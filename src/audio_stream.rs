use std::{
    thread::{self},
    time::Instant,
};

use crate::stream_instant::StreamInstant;
use crate::{
    audio_client::{AudioClientError, EventHandleWrapper, get_wait_error},
    sample_format::SampleFormat,
};
use windows::Win32::Media::Audio::{AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY, IAudioClock};
use windows::Win32::{
    Foundation::{HANDLE, WAIT_OBJECT_0},
    Media::Audio::{AUDCLNT_BUFFERFLAGS_SILENT, IAudioCaptureClient, IAudioClient, IAudioRenderClient},
    System::Threading::{
        CreateEventA, CreateEventW, GetCurrentThread, INFINITE, SetEvent, SetThreadPriority, THREAD_PRIORITY_TIME_CRITICAL,
        WaitForMultipleObjectsEx,
    },
};

pub(crate) struct StreamRunContext<T> {
    audio_client: IAudioClient,
    stream_client: T,
    stop_handle: HANDLE,
    format: SampleFormat,
}
unsafe impl<T> Send for StreamRunContext<T> {}

impl<T> StreamRunContext<T> {
    pub(crate) fn new(audio_client: IAudioClient, stream_client: T, stop_handle: HANDLE, format: SampleFormat) -> Self {
        Self {
            audio_client,
            stream_client,
            stop_handle,
            format,
        }
    }
}

pub struct AudioStreamConfig {
    stream_fn: Box<dyn FnOnce() + Send + 'static>,
    stop_handle: HANDLE,
    format: SampleFormat,
    thread_name: String,
}

unsafe impl Send for AudioStreamConfig {}

pub struct CapturePacket<'a> {
    data: &'a [u8],
    timestamp: StreamInstant,
}

impl<'a> CapturePacket<'a> {
    pub fn data(&self) -> &'a [u8] {
        self.data
    }

    pub fn timestamp(&self) -> &StreamInstant {
        &self.timestamp
    }
}

pub struct AudioStream {
    thread: Option<thread::JoinHandle<()>>,
    stop_handle: HANDLE,
}

unsafe impl Send for AudioStream {}

impl AudioStreamConfig {
    pub(crate) fn create_capture_stream<D, E>(
        data_callback: D,
        mut error_callback: E,
        audio_client: IAudioClient,
        format: Option<SampleFormat>,
    ) -> Result<AudioStreamConfig, AudioClientError>
    where
        D: FnMut(CapturePacket) + Send + 'static,
        E: FnMut(AudioClientError) + Send + 'static,
    {
        let capture_client =
            unsafe { audio_client.GetService::<IAudioCaptureClient>() }.map_err(AudioClientError::FailedToStartAudioClient)?;
        let stop_handle = unsafe { CreateEventW(None, false, false, None) }.map_err(AudioClientError::EventCreationError)?;

        let format = match format {
            Some(format) => format,
            None => {
                let mix_format = unsafe { audio_client.GetMixFormat() }.map_err(AudioClientError::FailedToGetMixFormat)?;
                SampleFormat::from_wave_format_ex(mix_format)
            }
        };

        let run_context = StreamRunContext {
            audio_client,
            stream_client: capture_client,
            stop_handle: stop_handle.clone(),
            format: format.clone(),
        };

        let capture_fn = move || {
            let res = Self::capture_audio(run_context, data_callback);
            if let Err(err) = res {
                error_callback(err);
            }
        };

        Ok(AudioStreamConfig {
            stream_fn: Box::new(capture_fn),
            stop_handle,
            format: format.clone(),
            thread_name: "capture".to_string(),
        })
    }

    pub(crate) fn create_playback_stream<D, E>(
        data_callback: D,
        mut error_callback: E,
        audio_client: IAudioClient,
        format: SampleFormat,
    ) -> Result<AudioStreamConfig, AudioClientError>
    where
        D: FnMut(&mut [u8]) -> bool + Send + 'static,
        E: FnMut(AudioClientError) + Send + 'static,
    {
        let render_client =
            unsafe { audio_client.GetService::<IAudioRenderClient>() }.map_err(AudioClientError::FailedToStartAudioClient)?;
        let stop_handle = unsafe { CreateEventW(None, false, false, None) }.map_err(AudioClientError::EventCreationError)?;

        let run_context = StreamRunContext {
            audio_client,
            stream_client: render_client,
            stop_handle: stop_handle.clone(),
            format: format.clone(),
        };

        let capture_fn = move || {
            let res = Self::playback_audio(run_context, data_callback);
            if let Err(err) = res {
                error_callback(err);
            }
        };

        Ok(AudioStreamConfig {
            stream_fn: Box::new(capture_fn),
            stop_handle,
            format,
            thread_name: "playback".to_string(),
        })
    }

    pub fn start(self) -> Result<AudioStream, AudioClientError> {
        let thr = thread::Builder::new()
            .name(self.thread_name)
            .spawn(self.stream_fn)
            .map_err(|_| AudioClientError::FailedToCreateThread)?;
        Ok(AudioStream {
            thread: Some(thr),
            stop_handle: self.stop_handle,
        })
    }

    pub fn format(&self) -> &SampleFormat {
        &self.format
    }

    fn capture_audio<D>(run_context: StreamRunContext<IAudioCaptureClient>, mut data_callback: D) -> Result<(), AudioClientError>
    where
        D: FnMut(CapturePacket),
    {
        Self::set_thread_priority();
        let (audio_client, capture_client) = (run_context.audio_client, run_context.stream_client);
        let audio_clock = unsafe { audio_client.GetService::<IAudioClock>() }.map_err(AudioClientError::FailedToGetAudioClock)?;

        let block_align = run_context.format.block_align() as usize;

        let mut buffer: *mut u8 = std::ptr::null_mut();
        let mut flags: u32 = 0;
        let mut pu64qpcposition: u64 = 0;

        let h_event = unsafe { CreateEventA(None, false, false, None) }.map_err(|h| AudioClientError::FailedToCreateStopEvent(h))?;
        let h_event = EventHandleWrapper(h_event);
        let handles = [*h_event, run_context.stop_handle];
        unsafe { audio_client.SetEventHandle(*h_event) }.map_err(|h| AudioClientError::FailedToSetupEventHandle(h))?;
        unsafe { audio_client.Start() }.map_err(|h| AudioClientError::FailedToStartAudioClient(h))?;

        while let Ok(mut frames_available) = unsafe { capture_client.GetNextPacketSize() } {
            let wait_res = unsafe { get_wait_error(WaitForMultipleObjectsEx(&handles, false, INFINITE, false))? };

            // Stop event was called
            if wait_res == WAIT_OBJECT_0.0 + 1 {
                break;
            }

            if frames_available == 0 {
                continue;
            }
            unsafe {
                capture_client.GetBuffer(
                    &mut buffer,
                    &mut frames_available as *mut _,
                    &mut flags as *mut _,
                    None,
                    Some(&mut pu64qpcposition as *mut _),
                )
            }
            .map_err(AudioClientError::FailedGettingBuffer)?;
            debug_assert!(!buffer.is_null());
            let now = convert_instant(pu64qpcposition);

            let buf_slice = unsafe { std::slice::from_raw_parts(buffer, frames_available as usize * block_align) };
            data_callback(CapturePacket {
                data: buf_slice,
                timestamp: now,
            });

            unsafe { capture_client.ReleaseBuffer(frames_available) }.map_err(AudioClientError::FailedReleasingBuffer)?;
        }
        unsafe {
            audio_client.Stop().map_err(AudioClientError::FailedStoppingAudioClient)?;
            audio_client.Reset().map_err(AudioClientError::FailedResettingAudioClient)?;
        }
        Ok(())
    }

    fn playback_audio<D>(run_context: StreamRunContext<IAudioRenderClient>, mut data_callback: D) -> Result<(), AudioClientError>
    where
        D: FnMut(&mut [u8]) -> bool,
    {
        Self::set_thread_priority();
        let (audio_client, render_client) = (run_context.audio_client, run_context.stream_client);

        let buffer_size = unsafe { audio_client.GetBufferSize() }.map_err(AudioClientError::FailedToStartAudioClient)?;
        let h_event = unsafe { CreateEventA(None, false, false, None) }.map_err(|h| AudioClientError::FailedToCreateStopEvent(h))?;
        let h_event = EventHandleWrapper(h_event);
        let handles = [*h_event, run_context.stop_handle];
        let block_align = run_context.format.block_align() as usize;

        unsafe { audio_client.SetEventHandle(*h_event) }.map_err(|h| AudioClientError::FailedToSetupEventHandle(h))?;
        unsafe { audio_client.Start() }.map_err(|h| AudioClientError::FailedToStartAudioClient(h))?;

        loop {
            let wait_res = unsafe { get_wait_error(WaitForMultipleObjectsEx(&handles, false, INFINITE, false))? };
            // Stop event was called
            if wait_res == WAIT_OBJECT_0.0 + 1 {
                break;
            }
            let padding = unsafe { audio_client.GetCurrentPadding() }.map_err(AudioClientError::FailedGettingBuffer)?;
            let available_frames = buffer_size - padding;
            if available_frames == 0 {
                continue;
            }

            let buffer = unsafe { render_client.GetBuffer(available_frames) }.map_err(AudioClientError::FailedGettingBuffer)?;
            let buffer = unsafe { std::slice::from_raw_parts_mut(buffer, available_frames as usize * block_align) };
            let is_active = data_callback(buffer);
            let flags = if is_active { 0u32 } else { AUDCLNT_BUFFERFLAGS_SILENT.0 as u32 };
            unsafe { render_client.ReleaseBuffer(available_frames, flags) }.map_err(AudioClientError::FailedReleasingBuffer)?;
        }

        Ok(())
    }

    fn set_thread_priority() {
        unsafe {
            let curr_thr = GetCurrentThread();
            let _ = SetThreadPriority(curr_thr, THREAD_PRIORITY_TIME_CRITICAL);
        }
    }
}

fn convert_instant(buffer_qpc_position: u64) -> StreamInstant {
    // The `qpc_position` is in 100 nanosecond units. Convert it to nanoseconds. source: `https://learn.microsoft.com/en-us/windows/win32/api/audioclient/nf-audioclient-iaudiocaptureclient-getbuffer`
    let qpc_nanos = buffer_qpc_position as i128 * 100;
    StreamInstant::from_nanos_i128(qpc_nanos).expect("performance counter out of range of `StreamInstant` representation")
}

impl AudioStream {
    // See drop implementation for cleanup
    pub fn stop_recording(self) {}
}

impl Drop for AudioStream {
    fn drop(&mut self) {
        unsafe {
            let _ = SetEvent(self.stop_handle);
        }
        let _ = self.thread.take().map(|thr| thr.join());
    }
}
