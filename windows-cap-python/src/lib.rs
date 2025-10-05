#![warn(clippy::nursery)]
#![warn(clippy::cargo)]
#![allow(clippy::redundant_pub_crate)]
#![allow(clippy::multiple_crate_versions)] // Should update as soon as possible

use std::os::raw::{c_char, c_int};
use std::{ptr, slice};

use ::windows_capture::dxgi_duplication_api::{DxgiDuplicationApi, Error as DxgiDuplicationError};
use ::windows_capture::monitor::Monitor;
use ::windows_capture::settings::ColorFormat;
use pyo3::exceptions::PyException;
use pyo3::ffi;
use pyo3::prelude::*;
use pyo3::types::PyMemoryView;
use windows::Win32::Graphics::Direct3D11::{
    D3D11_BOX, D3D11_CPU_ACCESS_READ, D3D11_MAP_READ, D3D11_MAPPED_SUBRESOURCE, D3D11_TEXTURE2D_DESC,
    D3D11_USAGE_STAGING, ID3D11Texture2D,
};
use windows::Win32::Graphics::Dxgi::Common::{
    DXGI_FORMAT, DXGI_FORMAT_B8G8R8A8_UNORM, DXGI_FORMAT_R8G8B8A8_UNORM, DXGI_FORMAT_R16G16B16A16_FLOAT,
    DXGI_SAMPLE_DESC,
};

/// Fastest Windows Screen Capture Library For Python 🔥.
#[pymodule]
fn windows_capture(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<NativeDxgiDuplication>()?;
    m.add_class::<NativeDxgiDuplicationFrame>()?;
    Ok(())
}

#[pyclass(unsendable)]
pub struct NativeDxgiDuplication {
    duplication: DxgiDuplicationApi,
    monitor: Monitor,
}

impl NativeDxgiDuplication {
    fn new_duplication(monitor: Monitor) -> Result<(Monitor, DxgiDuplicationApi), DxgiDuplicationError> {
        let duplication = DxgiDuplicationApi::new(monitor)?;

        Ok((monitor, duplication))
    }

    fn recreate_duplication(&mut self) -> Result<(), DxgiDuplicationError> {
        let (_, duplication) = Self::new_duplication(self.monitor)?;
        self.duplication = duplication;
        Ok(())
    }

    const fn color_format_to_str(color_format: ColorFormat) -> &'static str {
        match color_format {
            ColorFormat::Bgra8 => "bgra8",
            ColorFormat::Rgba8 => "rgba8",
            ColorFormat::Rgba16F => "rgba16f",
        }
    }

    const fn bytes_per_pixel(color_format: ColorFormat) -> usize {
        match color_format {
            ColorFormat::Bgra8 | ColorFormat::Rgba8 => 4,
            ColorFormat::Rgba16F => 8,
        }
    }

    fn color_format_from_dxgi(format: DXGI_FORMAT) -> PyResult<ColorFormat> {
        match format {
            DXGI_FORMAT_B8G8R8A8_UNORM => Ok(ColorFormat::Bgra8),
            DXGI_FORMAT_R8G8B8A8_UNORM => Ok(ColorFormat::Rgba8),
            DXGI_FORMAT_R16G16B16A16_FLOAT => Ok(ColorFormat::Rgba16F),
            other => Err(PyException::new_err(format!("Unsupported DXGI color format: {other:?}"))),
        }
    }
}

#[pymethods]
impl NativeDxgiDuplication {
    #[new]
    #[pyo3(signature = (monitor_index=None))]
    pub fn new(monitor_index: Option<usize>) -> PyResult<Self> {
        let monitor = match monitor_index {
            Some(index) => Monitor::from_index(index)
                .map_err(|e| PyException::new_err(format!("Failed to resolve monitor from index {index}: {e}",)))?,
            None => Monitor::primary()
                .map_err(|e| PyException::new_err(format!("Failed to acquire primary monitor: {e}",)))?,
        };

        let (_, duplication) = Self::new_duplication(monitor)
            .map_err(|e| PyException::new_err(format!("Failed to create DXGI duplication session: {e}")))?;

        Ok(Self { duplication, monitor })
    }

    #[pyo3(signature = (timeout_ms=16, area=None))]
    pub fn acquire_next_frame(
        &mut self,
        timeout_ms: u32,
        area: Option<Vec<i32>>,
    ) -> PyResult<Option<NativeDxgiDuplicationFrame>> {
        match self.duplication.acquire_next_frame(timeout_ms) {
            Ok(frame) => {
                let texture_desc = *frame.texture_desc();
                let color_format = Self::color_format_from_dxgi(texture_desc.Format)?;
                let bytes_per_pixel = Self::bytes_per_pixel(color_format);
                let mut src_box = D3D11_BOX {
                    left: 0,
                    top: 0,
                    front: 0,
                    right: texture_desc.Width,
                    bottom: texture_desc.Height,
                    back: 1,
                };

                if let Some(xywh) = area {
                    if xywh.iter().all(|&x| x >= 0) {
                        src_box.left = xywh[0] as u32;
                        src_box.top = xywh[1] as u32;
                        src_box.right = xywh[2] as u32;
                        src_box.bottom = xywh[3] as u32;
                    }
                }

                let width = src_box.right - src_box.left;
                let height = src_box.bottom - src_box.top;

                let staging_desc = D3D11_TEXTURE2D_DESC {
                    Width: width,
                    Height: height,
                    MipLevels: 1,
                    ArraySize: 1,
                    Format: texture_desc.Format,
                    SampleDesc: DXGI_SAMPLE_DESC { Count: 1, Quality: 0 },
                    Usage: D3D11_USAGE_STAGING,
                    BindFlags: 0,
                    CPUAccessFlags: D3D11_CPU_ACCESS_READ.0 as u32,
                    MiscFlags: 0,
                };

                let device_context = frame.device_context().clone();
                let device = frame.device().clone();

                let mut staging = None;
                unsafe { device.CreateTexture2D(&staging_desc, None, Some(&mut staging)) }
                    .map_err(|e| PyException::new_err(format!("Failed to create staging texture: {e}")))?;
                let staging = staging.expect("CreateTexture2D returned Ok but no texture");

                unsafe {
                    device_context.CopySubresourceRegion(
                        &staging,
                        0,
                        0,
                        0,
                        0,
                        frame.texture(),
                        0,
                        Some(&src_box as *const _),
                    );
                }

                let mut mapped = D3D11_MAPPED_SUBRESOURCE::default();
                unsafe { device_context.Map(&staging, 0, D3D11_MAP_READ, 0, Some(&mut mapped)) }
                    .map_err(|e| PyException::new_err(format!("Failed to map duplication frame: {e}")))?;

                let row_pitch_u32 = mapped.RowPitch;
                let row_pitch = usize::try_from(row_pitch_u32)
                    .map_err(|_| PyException::new_err("Failed to convert row pitch to usize"))?;
                let height_usize =
                    usize::try_from(height).map_err(|_| PyException::new_err("Failed to convert height to usize"))?;
                let len = row_pitch
                    .checked_mul(height_usize)
                    .ok_or_else(|| PyException::new_err("Mapped frame size overflowed usize"))?;

                let frame_obj = NativeDxgiDuplicationFrame::new(
                    device_context,
                    staging,
                    mapped.pData.cast::<u8>(),
                    len,
                    width,
                    height,
                    bytes_per_pixel,
                    row_pitch,
                    Self::color_format_to_str(color_format),
                );

                Ok(Some(frame_obj))
            }
            Err(DxgiDuplicationError::Timeout) => Ok(None),
            Err(DxgiDuplicationError::AccessLost) => {
                Err(PyException::new_err("DXGI duplication access lost; call recreate() to re-establish the session"))
            }
            Err(other) => Err(PyException::new_err(format!("Failed to acquire duplication frame: {other}"))),
        }
    }

    #[pyo3(signature = (monitor_index))]
    pub fn switch_monitor(&mut self, monitor_index: usize) -> PyResult<()> {
        let monitor = Monitor::from_index(monitor_index)
            .map_err(|e| PyException::new_err(format!("Failed to resolve monitor from index {monitor_index}: {e}")))?;

        let (_, duplication) = Self::new_duplication(monitor)
            .map_err(|e| PyException::new_err(format!("Failed to create DXGI duplication session: {e}")))?;

        self.monitor = monitor;
        self.duplication = duplication;

        Ok(())
    }

    pub fn recreate(&mut self) -> PyResult<()> {
        self.recreate_duplication()
            .map_err(|e| PyException::new_err(format!("Failed to recreate DXGI duplication session: {e}")))?;
        Ok(())
    }
}

#[pyclass(unsendable)]
pub struct NativeDxgiDuplicationFrame {
    context: windows::Win32::Graphics::Direct3D11::ID3D11DeviceContext,
    staging: ID3D11Texture2D,
    ptr: *mut u8,
    len: usize,
    width: u32,
    height: u32,
    bytes_per_pixel: usize,
    row_pitch: usize,
    color_format: &'static str,
    mapped: bool,
}

#[allow(clippy::missing_const_for_fn)]
impl NativeDxgiDuplicationFrame {
    #[allow(clippy::too_many_arguments)]
    fn new(
        context: windows::Win32::Graphics::Direct3D11::ID3D11DeviceContext,
        staging: ID3D11Texture2D,
        ptr: *mut u8,
        len: usize,
        width: u32,
        height: u32,
        bytes_per_pixel: usize,
        row_pitch: usize,
        color_format: &'static str,
    ) -> Self {
        Self { context, staging, ptr, len, width, height, bytes_per_pixel, row_pitch, color_format, mapped: true }
    }
}

impl Drop for NativeDxgiDuplicationFrame {
    fn drop(&mut self) {
        if self.mapped {
            unsafe {
                self.context.Unmap(&self.staging, 0);
            }
            self.mapped = false;
            self.ptr = ptr::null_mut();
            self.len = 0;
        }
    }
}

#[pymethods]
#[allow(clippy::missing_const_for_fn)]
impl NativeDxgiDuplicationFrame {
    #[getter]
    pub fn width(&self) -> u32 {
        self.width
    }

    #[getter]
    pub fn height(&self) -> u32 {
        self.height
    }

    #[getter]
    pub fn bytes_per_pixel(&self) -> usize {
        self.bytes_per_pixel
    }

    #[getter]
    pub fn color_format(&self) -> &'static str {
        self.color_format
    }

    #[getter]
    pub fn bytes_per_row(&self) -> usize {
        self.row_pitch
    }

    pub fn buffer_ptr(&self) -> usize {
        self.ptr as usize
    }

    pub fn buffer_len(&self) -> usize {
        self.len
    }

    pub fn to_bytes(&self) -> Vec<u8> {
        unsafe { slice::from_raw_parts(self.ptr, self.len) }.to_vec()
    }

    pub fn buffer_view<'py>(&'py self, py: Python<'py>) -> PyResult<Bound<'py, PyMemoryView>> {
        let len = isize::try_from(self.len).map_err(|_| PyException::new_err("Frame too large for memoryview"))?;
        const PYBUF_READ: c_int = 0x100;
        let view = unsafe { ffi::PyMemoryView_FromMemory(self.ptr.cast::<c_char>(), len, PYBUF_READ) };
        if view.is_null() {
            Err(PyException::new_err("Failed to create memoryview for DXGI frame"))
        } else {
            let any = unsafe { Bound::from_owned_ptr(py, view) };
            any.downcast_into().map_err(|e| e.into())
        }
    }
}
