//! DXGI Desktop Duplication API wrapper.
//!
//! This module provides [`DxgiDuplicationApi`] to capture a monitor using the
//! Windows DXGI Desktop Duplication API. It integrates with [`crate::monitor::Monitor`]
//! to select the target output and exposes CPU-readable frames via [`crate::frame::FrameBuffer`].
//!
//! # Example
//! ```no_run
//! use windows_capture::dxgi_duplication_api::DxgiDuplicationApi;
//! use windows_capture::encoder::ImageFormat;
//! use windows_capture::monitor::Monitor;
//!
//! fn main() -> Result<(), Box<dyn std::error::Error>> {
//!     // Select the primary monitor
//!     let monitor = Monitor::primary()?;
//!
//!     // Create a duplication session for this monitor
//!     let mut dup = DxgiDuplicationApi::new(monitor)?;
//!
//!     // Try to grab one frame within ~33ms (about 30 FPS budget)
//!     let mut frame = dup.acquire_next_frame(33)?;
//!
//!     // Map the GPU image into CPU memory and save a PNG
//!     let mut buffer = frame.buffer()?;
//!     buffer.save_as_image("dup.png", ImageFormat::Png)?;
//!     Ok(())
//! }
//! ```
use std::path::Path;
use std::{fs, io, slice};

use rayon::iter::{IntoParallelIterator, ParallelIterator};
use windows::Win32::Graphics::Direct3D11::{
    D3D11_BOX, D3D11_CPU_ACCESS_READ, D3D11_CPU_ACCESS_WRITE, D3D11_MAP_READ_WRITE, D3D11_MAPPED_SUBRESOURCE,
    D3D11_TEXTURE2D_DESC, D3D11_USAGE_STAGING, ID3D11Device, ID3D11DeviceContext, ID3D11Texture2D,
};
use windows::Win32::Graphics::Dxgi::Common::{
    DXGI_FORMAT, DXGI_FORMAT_B8G8R8A8_UNORM, DXGI_FORMAT_R8G8B8A8_UNORM, DXGI_FORMAT_R16G16B16A16_FLOAT,
    DXGI_SAMPLE_DESC,
};
use windows::Win32::Graphics::Dxgi::{
    DXGI_ERROR_ACCESS_LOST, DXGI_ERROR_WAIT_TIMEOUT, DXGI_OUTDUPL_DESC, DXGI_OUTDUPL_FRAME_INFO, IDXGIDevice,
    IDXGIOutput1, IDXGIOutputDuplication,
};
use windows::core::Interface;

use crate::d3d11::{StagingTexture, create_d3d_device};
use crate::encoder::{ImageEncoder, ImageFormat};
use crate::monitor::Monitor;
use crate::settings::ColorFormat;

/// Errors that can occur while using the DXGI Desktop Duplication API wrapper.
#[derive(thiserror::Error, Debug)]
pub enum Error {
    /// The crop rectangle is invalid (start >= end on either axis).
    #[error("Invalid crop size")]
    InvalidSize,
    /// Failed to find a DXGI output that corresponds to the provided monitor.
    #[error("Failed to find DXGI output for the specified monitor")]
    OutputNotFound,
    /// AcquireNextFrame timed out without a new frame becoming available.
    #[error("AcquireNextFrame timed out")]
    Timeout,
    /// The duplication access was lost and must be recreated.
    #[error("Duplication access lost; the duplication must be recreated")]
    AccessLost,
    /// DirectX device creation or related error.
    #[error("DirectX error: {0}")]
    DirectXError(#[from] crate::d3d11::Error),
    /// Unsupported color format encountered.
    #[error("Unsupported color format: {0:?}")]
    UnsupportedColorFormat(DXGI_FORMAT),
    /// Invalid or mismatched staging texture supplied to [`DxgiDuplicationFrame::buffer_with`].
    #[error("Invalid staging texture: {0}")]
    InvalidStagingTexture(&'static str),
    /// Image encoding failed.
    ///
    /// Wraps [`crate::encoder::ImageEncoderError`].
    #[error("Failed to encode the image buffer to image bytes with the specified format: {0}")]
    ImageEncoderError(#[from] crate::encoder::ImageEncoderError),
    /// An I/O error occurred while writing the image to disk.
    ///
    /// Wraps [`std::io::Error`].
    #[error("I/O error: {0}")]
    IoError(#[from] io::Error),
    /// Windows API error.
    #[error("Windows API error: {0}")]
    WindowsError(#[from] windows::core::Error),
}

/// A minimal, ergonomic wrapper around the DXGI Desktop Duplication API for capturing a monitor.
///
/// This wrapper focuses on staying close to the native API while providing a simple Rust interface.
/// It integrates with [`crate::monitor::Monitor`] to select the target output.
pub struct DxgiDuplicationApi {
    /// Direct3D 11 device used for duplication operations.
    d3d_device: ID3D11Device,
    /// Direct3D 11 device context used for copy/map operations.
    d3d_device_context: ID3D11DeviceContext,
    /// The duplication interface used to acquire frames.
    duplication: IDXGIOutputDuplication,
    /// Description of the duplication, including format and dimensions.
    duplication_desc: DXGI_OUTDUPL_DESC,
    /// Whether the internal staging texture is currently holding a frame.
    is_holding_frame: bool,
}

impl DxgiDuplicationApi {
    /// Constructs a new duplication session for the specified monitor.
    ///
    /// Internally creates a Direct3D 11 device and immediate context using the crate's d3d11
    /// module.
    pub fn new(monitor: Monitor) -> Result<Self, Error> {
        // Create D3D11 device and context.
        let (d3d_device, d3d_device_context) = create_d3d_device()?;

        // Get the adapter used by the created device.
        let dxgi_device: IDXGIDevice = d3d_device.cast()?;
        let adapter = unsafe { dxgi_device.GetAdapter()? };

        // Find the DXGI output that corresponds to the provided HMONITOR.
        let found_output;
        let mut index = 0u32;
        loop {
            let output = unsafe { adapter.EnumOutputs(index) }?;
            let desc = unsafe { output.GetDesc()? };
            if desc.Monitor.0 == monitor.as_raw_hmonitor() {
                found_output = Some(output);
                break;
            }
            index += 1;
        }

        let Some(output) = found_output else {
            return Err(Error::OutputNotFound);
        };

        // IDXGIOutput1 is required for DuplicateOutput.
        let output1: IDXGIOutput1 = output.cast()?;

        // Create the duplication for this output using the supplied D3D11 device.
        let duplication = unsafe { output1.DuplicateOutput(&d3d_device)? };

        // Get the duplication description to determine the format for our internal texture.
        let duplication_desc = unsafe { duplication.GetDesc() };

        Ok(Self { d3d_device, d3d_device_context, duplication, duplication_desc, is_holding_frame: false })
    }

    /// Gets the underlying [`windows::Win32::Graphics::Direct3D11::ID3D11Device`] associated with
    /// this object.
    #[inline]
    #[must_use]
    pub const fn device(&self) -> &ID3D11Device {
        &self.d3d_device
    }

    /// Gets the underlying [`windows::Win32::Graphics::Direct3D11::ID3D11DeviceContext`] used for
    /// GPU operations.
    #[inline]
    #[must_use]
    pub const fn device_context(&self) -> &ID3D11DeviceContext {
        &self.d3d_device_context
    }

    /// Gets the underlying [`windows::Win32::Graphics::Dxgi::IDXGIOutputDuplication`] interface.
    #[inline]
    #[must_use]
    pub const fn duplication(&self) -> &IDXGIOutputDuplication {
        &self.duplication
    }

    /// Gets the [`windows::Win32::Graphics::Dxgi::DXGI_OUTDUPL_DESC`] of the duplication.
    #[inline]
    #[must_use]
    pub const fn duplication_desc(&self) -> &DXGI_OUTDUPL_DESC {
        &self.duplication_desc
    }

    /// Acquires the next frame and updates the internal texture.
    ///
    /// This call will block up to `timeout_ms` milliseconds. If no new frame arrives within
    /// the timeout, [`Error::Timeout`] is returned. If duplication access is lost,
    /// [`Error::AccessLost`] is returned and a new duplication should be created via
    /// [`DxgiDuplicationApi::new`].
    ///
    /// Main reasons for [`Error::AccessLost`] include:
    /// - The display mode of the output changed (e.g. resolution or color format change).
    /// - The user switched to a different desktop (e.g. via Ctrl+Alt+Del or Fast User Switching).
    /// - Switch from DWM on, DWM off, or other full-screen application
    ///
    /// The returned [`DxgiDuplicationFrame`] allows you to map the current full desktop image via
    /// [`DxgiDuplicationFrame::buffer`]. It contains the list of dirty rectangles reported for this
    /// frame.
    ///
    /// # Errors
    /// - [`Error::Timeout`] when no frame arrives within `timeout_ms`
    /// - [`Error::AccessLost`] when duplication access is lost and must be recreated
    /// - [`Error::WindowsError`] for other Windows API failures during frame acquisition
    #[inline]
    pub fn acquire_next_frame(&mut self, timeout_ms: u32) -> Result<DxgiDuplicationFrame<'_>, Error> {
        let mut frame_info = DXGI_OUTDUPL_FRAME_INFO::default();
        let mut resource = None;

        // Release the previous frame if we were holding one
        if self.is_holding_frame {
            unsafe { self.duplication.ReleaseFrame() }?;
        }

        // Acquire frame
        match unsafe { self.duplication.AcquireNextFrame(timeout_ms, &mut frame_info, &mut resource) } {
            Ok(()) => (),
            Err(e) => {
                if e.code() == DXGI_ERROR_WAIT_TIMEOUT {
                    return Err(Error::Timeout);
                } else if e.code() == DXGI_ERROR_ACCESS_LOST {
                    return Err(Error::AccessLost);
                } else {
                    return Err(Error::WindowsError(e));
                }
            }
        }
        self.is_holding_frame = true;

        let resource = resource.unwrap();

        // Convert the resource to an ID3D11Texture2D.
        let frame_texture: ID3D11Texture2D = resource.cast()?;

        // Obtain texture description to get size/format details.
        let mut frame_desc = D3D11_TEXTURE2D_DESC::default();
        unsafe { frame_texture.GetDesc(&mut frame_desc) };

        Ok(DxgiDuplicationFrame {
            d3d_device: &self.d3d_device,
            d3d_device_context: &self.d3d_device_context,
            duplication: &self.duplication,
            texture: frame_texture,
            texture_desc: frame_desc,
            frame_info,
        })
    }
}

/// Represents a pre-assembled full desktop image for the current frame,
/// backed by the internal GPU texture.
/// Call [`DxgiDuplicationFrame::buffer`] to obtain a CPU-readable [`crate::frame::FrameBuffer`].
pub struct DxgiDuplicationFrame<'a> {
    d3d_device: &'a ID3D11Device,
    d3d_device_context: &'a ID3D11DeviceContext,
    duplication: &'a IDXGIOutputDuplication,
    texture: ID3D11Texture2D,
    texture_desc: D3D11_TEXTURE2D_DESC,
    frame_info: DXGI_OUTDUPL_FRAME_INFO,
}

impl<'a> DxgiDuplicationFrame<'a> {
    /// Gets the width of the frame.
    #[inline]
    #[must_use]
    pub const fn width(&self) -> u32 {
        self.texture_desc.Width
    }

    /// Gets the height of the frame.
    #[inline]
    #[must_use]
    pub const fn height(&self) -> u32 {
        self.texture_desc.Height
    }

    /// Gets the underlying Direct3D device associated with this frame.
    #[inline]
    #[must_use]
    pub const fn device(&self) -> &ID3D11Device {
        self.d3d_device
    }

    /// Gets the underlying Direct3D device context used for GPU operations.
    #[inline]
    #[must_use]
    pub const fn device_context(&self) -> &ID3D11DeviceContext {
        self.d3d_device_context
    }

    /// Gets the underlying IDXGIOutputDuplication interface.
    #[inline]
    #[must_use]
    pub const fn duplication(&self) -> &IDXGIOutputDuplication {
        self.duplication
    }

    /// Gets the underlying [`windows::Win32::Graphics::Direct3D11::ID3D11Texture2D`] interface.
    #[inline]
    #[must_use]
    pub const fn texture(&self) -> &ID3D11Texture2D {
        &self.texture
    }

    /// Gets the [`windows::Win32::Graphics::Direct3D11::D3D11_TEXTURE2D_DESC`] of the underlying
    /// texture.
    #[inline]
    #[must_use]
    pub const fn texture_desc(&self) -> &D3D11_TEXTURE2D_DESC {
        &self.texture_desc
    }

    /// Gets the frame information for the current frame.
    #[inline]
    #[must_use]
    pub const fn frame_info(&self) -> &DXGI_OUTDUPL_FRAME_INFO {
        &self.frame_info
    }

    /// Maps the internal frame into CPU accessible memory and returns a
    /// [`crate::frame::FrameBuffer`].
    ///
    /// This creates a staging texture, copies the internal texture into it,
    /// and maps it for CPU read/write. The returned buffer may include row padding;
    /// you can use [`crate::frame::FrameBuffer::as_nopadding_buffer`] to obtain a packed
    /// representation.
    #[inline]
    pub fn buffer<'b>(&'b mut self) -> Result<DxgiDuplicationFrameBuffer<'b>, Error> {
        // Staging texture settings
        let texture_desc = D3D11_TEXTURE2D_DESC {
            Width: self.texture_desc.Width,
            Height: self.texture_desc.Height,
            MipLevels: 1,
            ArraySize: 1,
            Format: self.texture_desc.Format,
            SampleDesc: DXGI_SAMPLE_DESC { Count: 1, Quality: 0 },
            Usage: D3D11_USAGE_STAGING,
            BindFlags: 0,
            CPUAccessFlags: D3D11_CPU_ACCESS_READ.0 as u32 | D3D11_CPU_ACCESS_WRITE.0 as u32,
            MiscFlags: 0,
        };

        // Create a CPU-readable staging texture
        let mut staging = None;
        unsafe {
            self.d3d_device.CreateTexture2D(&texture_desc, None, Some(&mut staging))?;
        };
        let staging = staging.unwrap();

        // Copy from the internal GPU texture into the staging texture
        unsafe {
            self.d3d_device_context.CopyResource(&staging, &self.texture);
        };

        // Map the staging texture for CPU access
        let mut mapped = D3D11_MAPPED_SUBRESOURCE::default();
        unsafe {
            self.d3d_device_context.Map(&staging, 0, D3D11_MAP_READ_WRITE, 0, Some(&mut mapped))?;
        };

        // SAFETY: The staging texture remains alive for the scope of this function.
        let mapped_frame_data = unsafe {
            slice::from_raw_parts_mut(mapped.pData.cast(), (self.texture_desc.Height * mapped.RowPitch) as usize)
        };

        let color_format = match self.texture_desc.Format {
            DXGI_FORMAT_B8G8R8A8_UNORM => ColorFormat::Bgra8,
            DXGI_FORMAT_R8G8B8A8_UNORM => ColorFormat::Rgba8,
            DXGI_FORMAT_R16G16B16A16_FLOAT => ColorFormat::Rgba16F,
            _ => return Err(Error::UnsupportedColorFormat(self.texture_desc.Format)),
        };

        Ok(DxgiDuplicationFrameBuffer::new(
            mapped_frame_data,
            self.texture_desc.Width,
            self.texture_desc.Height,
            mapped.RowPitch,
            mapped.DepthPitch,
            color_format,
        ))
    }

    /// Gets a cropped frame buffer of the duplication frame.
    #[inline]
    pub fn buffer_crop<'b>(
        &'b mut self,
        start_x: u32,
        start_y: u32,
        end_x: u32,
        end_y: u32,
    ) -> Result<DxgiDuplicationFrameBuffer<'b>, Error> {
        if start_x >= end_x || start_y >= end_y {
            return Err(Error::InvalidSize);
        }

        let texture_width = end_x - start_x;
        let texture_height = end_y - start_y;

        // Staging texture settings for the cropped region
        let texture_desc = D3D11_TEXTURE2D_DESC {
            Width: texture_width,
            Height: texture_height,
            MipLevels: 1,
            ArraySize: 1,
            Format: self.texture_desc.Format,
            SampleDesc: DXGI_SAMPLE_DESC { Count: 1, Quality: 0 },
            Usage: D3D11_USAGE_STAGING,
            BindFlags: 0,
            CPUAccessFlags: D3D11_CPU_ACCESS_READ.0 as u32 | D3D11_CPU_ACCESS_WRITE.0 as u32,
            MiscFlags: 0,
        };

        // Create a CPU-readable staging texture of the crop size
        let mut staging = None;
        unsafe {
            self.d3d_device.CreateTexture2D(&texture_desc, None, Some(&mut staging))?;
        };
        let staging = staging.unwrap();

        // Define the source box to copy from the duplication texture
        let src_box = D3D11_BOX { left: start_x, top: start_y, front: 0, right: end_x, bottom: end_y, back: 1 };

        // Copy the selected region into the staging texture at (0,0)
        unsafe {
            self.d3d_device_context.CopySubresourceRegion(&staging, 0, 0, 0, 0, &self.texture, 0, Some(&src_box));
        }

        // Map the staging texture for CPU access
        let mut mapped = D3D11_MAPPED_SUBRESOURCE::default();
        unsafe {
            self.d3d_device_context.Map(&staging, 0, D3D11_MAP_READ_WRITE, 0, Some(&mut mapped))?;
        }

        // SAFETY: staging remains alive for the scope of this function.
        let mapped_frame_data =
            unsafe { slice::from_raw_parts_mut(mapped.pData.cast(), (texture_height * mapped.RowPitch) as usize) };

        let color_format = match self.texture_desc.Format {
            DXGI_FORMAT_B8G8R8A8_UNORM => ColorFormat::Bgra8,
            DXGI_FORMAT_R8G8B8A8_UNORM => ColorFormat::Rgba8,
            DXGI_FORMAT_R16G16B16A16_FLOAT => ColorFormat::Rgba16F,
            _ => return Err(Error::UnsupportedColorFormat(self.texture_desc.Format)),
        };

        Ok(DxgiDuplicationFrameBuffer::new(
            mapped_frame_data,
            texture_width,
            texture_height,
            mapped.RowPitch,
            mapped.DepthPitch,
            color_format,
        ))
    }

    /// Advanced: reuse your own CPU staging texture ([`crate::d3d11::StagingTexture`]).
    ///
    /// This avoids per-frame allocations and lets you manage the texture’s lifetime.
    /// The `staging` texture must be a `D3D11_USAGE_STAGING` 2D texture with CPU read/write access,
    /// matching the frame’s width/height/format.
    #[inline]
    pub fn buffer_with<'s>(
        &'s mut self,
        staging: &'s mut StagingTexture,
    ) -> Result<DxgiDuplicationFrameBuffer<'s>, Error> {
        // Validate geometry/format match.
        let desc = staging.desc();
        if desc.Width != self.texture_desc.Width || desc.Height != self.texture_desc.Height {
            return Err(Error::InvalidStagingTexture("geometry must match the frame"));
        }
        if desc.Format != self.texture_desc.Format {
            return Err(Error::InvalidStagingTexture("format must match the frame"));
        }

        // Unmap if was previously mapped
        if staging.is_mapped() {
            unsafe { self.d3d_device_context.Unmap(staging.texture(), 0) };
            staging.set_mapped(false);
        }

        // Copy the acquired duplication texture into the provided staging texture
        unsafe {
            self.d3d_device_context.CopyResource(staging.texture(), &self.texture);
        }

        // Map the staging texture for CPU access
        let mut mapped = D3D11_MAPPED_SUBRESOURCE::default();
        unsafe {
            self.d3d_device_context.Map(staging.texture(), 0, D3D11_MAP_READ_WRITE, 0, Some(&mut mapped))?;
        }
        staging.set_mapped(true);

        // SAFETY: staging lives for 's and remains alive while the FrameBuffer is borrowed.
        let mapped_frame_data = unsafe {
            slice::from_raw_parts_mut(mapped.pData.cast(), (self.texture_desc.Height * mapped.RowPitch) as usize)
        };

        let color_format = match self.texture_desc.Format {
            DXGI_FORMAT_B8G8R8A8_UNORM => ColorFormat::Bgra8,
            DXGI_FORMAT_R8G8B8A8_UNORM => ColorFormat::Rgba8,
            DXGI_FORMAT_R16G16B16A16_FLOAT => ColorFormat::Rgba16F,
            _ => return Err(Error::UnsupportedColorFormat(self.texture_desc.Format)),
        };

        Ok(DxgiDuplicationFrameBuffer::new(
            mapped_frame_data,
            self.texture_desc.Width,
            self.texture_desc.Height,
            mapped.RowPitch,
            mapped.DepthPitch,
            color_format,
        ))
    }

    /// Advanced: cropped buffer using a preallocated staging texture.
    /// The provided staging texture must be a D3D11_USAGE_STAGING 2D texture with CPU read/write
    /// access, of the same format as the duplication frame, and large enough to contain the
    /// crop region.
    #[inline]
    pub fn buffer_crop_with<'s>(
        &'s mut self,
        staging: &'s mut StagingTexture,
        start_x: u32,
        start_y: u32,
        end_x: u32,
        end_y: u32,
    ) -> Result<DxgiDuplicationFrameBuffer<'s>, Error> {
        // Validate crop rectangle
        if start_x >= end_x || start_y >= end_y {
            return Err(Error::InvalidSize);
        }

        let crop_width = end_x - start_x;
        let crop_height = end_y - start_y;

        // Validate format and capacity
        let desc = staging.desc();
        if desc.Format != self.texture_desc.Format {
            return Err(Error::InvalidStagingTexture("format must match the frame"));
        }
        if desc.Width < crop_width || desc.Height < crop_height {
            return Err(Error::InvalidStagingTexture("staging texture too small for crop region"));
        }

        // Unmap if was previously mapped
        if staging.is_mapped() {
            unsafe { self.d3d_device_context.Unmap(staging.texture(), 0) };
            staging.set_mapped(false);
        }

        // Define the source region to copy
        let src_box = D3D11_BOX { left: start_x, top: start_y, front: 0, right: end_x, bottom: end_y, back: 1 };

        // Copy the selected region to the top-left of the staging texture
        unsafe {
            self.d3d_device_context.CopySubresourceRegion(
                staging.texture(),
                0,
                0,
                0,
                0,
                &self.texture,
                0,
                Some(&src_box),
            );
        }

        // Map the staging texture
        let mut mapped = D3D11_MAPPED_SUBRESOURCE::default();
        unsafe {
            self.d3d_device_context.Map(staging.texture(), 0, D3D11_MAP_READ_WRITE, 0, Some(&mut mapped))?;
        }
        staging.set_mapped(true);

        // SAFETY: staging lives for 's and remains alive while the FrameBuffer is borrowed.
        let mapped_frame_data =
            unsafe { slice::from_raw_parts_mut(mapped.pData.cast(), (crop_height * mapped.RowPitch) as usize) };

        let color_format = match self.texture_desc.Format {
            DXGI_FORMAT_B8G8R8A8_UNORM => ColorFormat::Bgra8,
            DXGI_FORMAT_R8G8B8A8_UNORM => ColorFormat::Rgba8,
            DXGI_FORMAT_R16G16B16A16_FLOAT => ColorFormat::Rgba16F,
            _ => return Err(Error::UnsupportedColorFormat(self.texture_desc.Format)),
        };

        Ok(DxgiDuplicationFrameBuffer::new(
            mapped_frame_data,
            crop_width,
            crop_height,
            mapped.RowPitch,
            mapped.DepthPitch,
            color_format,
        ))
    }

    /// Saves the frame buffer as an image to the specified path.
    #[inline]
    pub fn save_as_image<T: AsRef<Path>>(&mut self, path: T, format: ImageFormat) -> Result<(), Error> {
        let mut frame_buffer = self.buffer()?;

        let width = frame_buffer.width();
        let height = frame_buffer.height();

        let bytes = ImageEncoder::new(format, frame_buffer.color_format())?.encode(
            frame_buffer.as_raw_buffer(),
            width,
            height,
        )?;

        fs::write(path, bytes)?;

        Ok(())
    }
}

/// Represents a frame buffer containing pixel data.
///
/// # Example
/// ```ignore
/// // Get a frame from the capture session
/// let mut buffer = frame.buffer()?;
/// buffer.save_as_image("screenshot.png", ImageFormat::Png)?;
/// ```
pub struct DxgiDuplicationFrameBuffer<'a> {
    raw_buffer: &'a mut [u8],
    width: u32,
    height: u32,
    row_pitch: u32,
    depth_pitch: u32,
    color_format: ColorFormat,
}

impl<'a> DxgiDuplicationFrameBuffer<'a> {
    /// Constructs a new `FrameBuffer`.
    #[inline]
    #[must_use]
    pub const fn new(
        raw_buffer: &'a mut [u8],
        width: u32,
        height: u32,
        row_pitch: u32,
        depth_pitch: u32,
        color_format: ColorFormat,
    ) -> Self {
        Self { raw_buffer, width, height, row_pitch, depth_pitch, color_format }
    }

    /// Gets the width of the frame buffer.
    #[inline]
    #[must_use]
    pub const fn width(&self) -> u32 {
        self.width
    }

    /// Gets the height of the frame buffer.
    #[inline]
    #[must_use]
    pub const fn height(&self) -> u32 {
        self.height
    }

    /// Gets the row pitch of the frame buffer.
    #[inline]
    #[must_use]
    pub const fn row_pitch(&self) -> u32 {
        self.row_pitch
    }

    /// Gets the depth pitch of the frame buffer.
    #[inline]
    #[must_use]
    pub const fn depth_pitch(&self) -> u32 {
        self.depth_pitch
    }

    /// Gets the color format of the frame buffer.
    #[inline]
    #[must_use]
    pub const fn color_format(&self) -> ColorFormat {
        self.color_format
    }

    /// Checks if the buffer has padding.
    #[inline]
    #[must_use]
    pub const fn has_padding(&self) -> bool {
        self.width * 4 != self.row_pitch
    }

    /// Gets the pixel data without padding.
    #[inline]
    #[must_use]
    pub fn as_nopadding_buffer<'b>(&'b self, buffer: &'b mut Vec<u8>) -> &'b [u8] {
        if !self.has_padding() {
            return self.raw_buffer;
        }

        let multiplier = match self.color_format {
            ColorFormat::Rgba16F => 8,
            ColorFormat::Rgba8 => 4,
            ColorFormat::Bgra8 => 4,
        };

        let frame_size = (self.width * self.height * multiplier) as usize;
        if buffer.capacity() < frame_size {
            buffer.resize(frame_size, 0);
        }

        let width_size = (self.width * multiplier) as usize;
        let buffer_address = buffer.as_mut_ptr() as isize;
        (0..self.height).into_par_iter().for_each(|y| {
            let index = (y * self.row_pitch) as usize;
            let ptr = buffer_address as *mut u8;

            unsafe {
                std::ptr::copy_nonoverlapping(
                    self.raw_buffer.as_ptr().add(index),
                    ptr.add(y as usize * width_size),
                    width_size,
                );
            }
        });

        &buffer[0..frame_size]
    }

    /// Gets the raw pixel data, which may include padding.
    #[inline]
    #[must_use]
    pub const fn as_raw_buffer(&mut self) -> &mut [u8] {
        self.raw_buffer
    }

    /// Saves the frame buffer as an image to the specified path.
    #[inline]
    pub fn save_as_image<T: AsRef<Path>>(&mut self, path: T, format: ImageFormat) -> Result<(), Error> {
        let width = self.width;
        let height = self.height;

        let mut buffer = Vec::new();
        let bytes = ImageEncoder::new(format, self.color_format)?.encode(
            self.as_nopadding_buffer(&mut buffer),
            width,
            height,
        )?;

        fs::write(path, bytes)?;

        Ok(())
    }
}
