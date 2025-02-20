use std::thread;

use log::error;
use windows::Win32::{
    Foundation::{self, HANDLE, WAIT_FAILED, WAIT_OBJECT_0},
    Media::Audio::{IAudioCaptureClient, IAudioClient, IAudioRenderClient, AUDCLNT_BUFFERFLAGS_SILENT},
    System::Threading::{
        CreateEventA, CreateEventW, GetCurrentThread, SetEvent, SetThreadPriority, WaitForMultipleObjectsEx, INFINITE,
        THREAD_PRIORITY_TIME_CRITICAL,
    },
};

use crate::{
    audio_client::{get_wait_error, AudioError, EventHandleWrapper},
    sample_format::SampleFormat,
};

pub(crate) struct CaptureRunContext {
    audio_client: IAudioClient,
    capture_client: IAudioCaptureClient,
    stop_handle: HANDLE,
    format: SampleFormat,
}
unsafe impl Send for CaptureRunContext {}

impl CaptureRunContext {
    pub(crate) fn new(audio_client: IAudioClient, capture_client: IAudioCaptureClient, stop_handle: HANDLE, format: SampleFormat) -> Self {
        Self {
            audio_client,
            capture_client,
            stop_handle,
            format,
        }
    }
}

pub(crate) struct PlaybackRunContext {
    audio_client: IAudioClient,
    render_client: IAudioRenderClient,
    stop_handle: HANDLE,
    format: SampleFormat,
}

unsafe impl Send for PlaybackRunContext {}

impl PlaybackRunContext {
    pub(crate) fn new(audio_client: IAudioClient, render_client: IAudioRenderClient, stop_handle: HANDLE, format: SampleFormat) -> Self {
        Self {
            audio_client,
            render_client,
            stop_handle,
            format,
        }
    }
}

pub struct AudioStream {
    thread: Option<thread::JoinHandle<()>>,
    thread_stop_handle: HANDLE,
}

unsafe impl Send for AudioStream {}

impl AudioStream {
    pub(crate) fn start_capture_stream<D, E>(
        data_callback: D,
        mut error_callback: E,
        audio_client: IAudioClient,
        format: SampleFormat,
    ) -> Result<AudioStream, AudioError>
    where
        D: FnMut(&[u8]) + Send + 'static,
        E: FnMut(AudioError) + Send + 'static,
    {
        let capture_client = unsafe { audio_client.GetService::<IAudioCaptureClient>() }.map_err(AudioError::FailedToStartAudioClient)?;
        let stop_handle = unsafe { CreateEventW(None, false, false, None) }.map_err(AudioError::EventCreationError)?;

        let run_context = CaptureRunContext {
            audio_client,
            capture_client,
            stop_handle: stop_handle.clone(),
            format: format.clone(),
        };

        let thr = thread::spawn(move || {
            let res = Self::capture_audio(run_context, data_callback);
            if let Err(err) = res {
                error_callback(err);
            }
        });

        Ok(AudioStream {
            thread: Some(thr),
            thread_stop_handle: stop_handle,
        })
    }

    pub(crate) fn start_playback_stream<D, E>(
        data_callback: D,
        mut error_callback: E,
        audio_client: IAudioClient,
        format: SampleFormat,
    ) -> Result<AudioStream, AudioError>
    where
        D: FnMut(&mut [u8]) -> bool + Send + 'static,
        E: FnMut(AudioError) + Send + 'static,
    {
        let render_client = unsafe { audio_client.GetService::<IAudioRenderClient>() }.map_err(AudioError::FailedToStartAudioClient)?;
        let stop_handle = unsafe { CreateEventW(None, false, false, None) }.map_err(AudioError::EventCreationError)?;

        let run_context = PlaybackRunContext {
            audio_client,
            render_client,
            stop_handle: stop_handle.clone(),
            format: format.clone(),
        };

        let thr = thread::spawn(move || {
            let res = Self::playback_audio(run_context, data_callback);
            if let Err(err) = res {
                error_callback(err);
            }
        });

        Ok(AudioStream {
            thread: Some(thr),
            thread_stop_handle: stop_handle,
        })
    }

    // See drop implementation for cleanup
    pub fn stop_recording(self) {}

    fn capture_audio<D>(run_context: CaptureRunContext, mut data_callback: D) -> Result<(), AudioError>
    where
        D: FnMut(&[u8]),
    {
        Self::set_thread_priority();
        let (audio_client, capture_client) = (run_context.audio_client, run_context.capture_client);
        let mut buffer: *mut u8 = std::ptr::null_mut();
        let mut flags: u32 = 0;
        let mut pu64deviceposition: u64 = 0;
        let mut pu64qpcposition: u64 = 0;
        let block_align = run_context.format.block_align() as usize;

        let h_event = unsafe { CreateEventA(None, false, false, None) }.map_err(|h| AudioError::FailedToCreateStopEvent(h))?;
        let h_event = EventHandleWrapper(h_event);
        let handles = [*h_event, run_context.stop_handle];
        unsafe { audio_client.SetEventHandle(*h_event) }.map_err(|h| AudioError::FailedToSetupEventHandle(h))?;
        unsafe { audio_client.Start() }.map_err(|h| AudioError::FailedToStartAudioClient(h))?;

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
                    Some(&mut pu64deviceposition as *mut _),
                    Some(&mut pu64qpcposition as *mut _),
                )
            }
            .map_err(AudioError::FailedGettingBuffer)?;
            debug_assert!(!buffer.is_null());

            let buf_slice = unsafe { std::slice::from_raw_parts(buffer, frames_available as usize * block_align) };
            data_callback(buf_slice);

            unsafe { capture_client.ReleaseBuffer(frames_available) }.map_err(AudioError::FailedReleasingBuffer)?;
        }
        unsafe {
            audio_client.Stop().map_err(AudioError::FailedStoppingAudioClient)?;
            audio_client.Reset().map_err(AudioError::FailedResettingAudioClient)?;
        }
        Ok(())
    }

    fn playback_audio<D>(run_context: PlaybackRunContext, mut data_callback: D) -> Result<(), AudioError>
    where
        D: FnMut(&mut [u8]) -> bool,
    {
        Self::set_thread_priority();
        let (audio_client, render_client) = (run_context.audio_client, run_context.render_client);

        let buffer_size = unsafe { audio_client.GetBufferSize() }.map_err(AudioError::FailedToStartAudioClient)?;
        let h_event = unsafe { CreateEventA(None, false, false, None) }.map_err(|h| AudioError::FailedToCreateStopEvent(h))?;
        let h_event = EventHandleWrapper(h_event);
        let handles = [*h_event, run_context.stop_handle];
        let block_align = run_context.format.block_align() as usize;

        unsafe { audio_client.SetEventHandle(*h_event) }.map_err(|h| AudioError::FailedToSetupEventHandle(h))?;
        unsafe { audio_client.Start() }.map_err(|h| AudioError::FailedToStartAudioClient(h))?;

        loop {
            let wait_res = unsafe { get_wait_error(WaitForMultipleObjectsEx(&handles, false, 0, false))? };
            // Stop event was called
            if wait_res == WAIT_OBJECT_0.0 + 1 {
                break;
            }
            let padding = unsafe { audio_client.GetCurrentPadding() }.map_err(AudioError::FailedGettingBuffer)?;
            let available_frames = buffer_size - padding;
            if available_frames == 0 {
                continue;
            }

            let buffer = unsafe { render_client.GetBuffer(available_frames) }.map_err(AudioError::FailedGettingBuffer)?;
            let buffer = unsafe { std::slice::from_raw_parts_mut(buffer, available_frames as usize * block_align) };
            let is_active = data_callback(buffer);
            let flags = if is_active { 0u32 } else { AUDCLNT_BUFFERFLAGS_SILENT.0 as u32 };
            unsafe { render_client.ReleaseBuffer(available_frames, flags) }.map_err(AudioError::FailedReleasingBuffer)?;
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

impl Drop for AudioStream {
    fn drop(&mut self) {
        unsafe {
            let _ = SetEvent(self.thread_stop_handle);
        }
        let _ = self.thread.take().map(|thr| thr.join());
    }
}
