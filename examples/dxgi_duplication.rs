use windows_capture::dxgi_duplication_api::DxgiDuplicationApi;
use windows_capture::frame::ImageFormat;
use windows_capture::monitor::Monitor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Select a monitor (primary in this example)
    let monitor = Monitor::primary()?;

    // Create a duplication session for this monitor
    let mut dup = DxgiDuplicationApi::new(monitor)?;

    // Try to grab one frame within ~33ms (about 30 FPS budget)
    let mut frame = dup.acquire_next_frame(33)?;

    // Map the GPU image into CPU memory and save a PNG
    // Note: The API could send an empty frame especially
    // in the first few calls, you can check this by seeing if
    // frame.frame_info().LastPresentTime is zero.
    frame.save_as_image("dxgi_screenshot.png", ImageFormat::Png)?;

    Ok(())
}
