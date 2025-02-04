use std::thread;

use log::error;
use windows::Win32::{
    Foundation::{self, HANDLE, WAIT_FAILED, WAIT_OBJECT_0},
    Media::Audio::{IAudioCaptureClient, IAudioClient},
    System::Threading::{
        CreateEventA, CreateEventW, GetCurrentThread, SetEvent, SetThreadPriority, WaitForMultipleObjectsEx, INFINITE,
        THREAD_PRIORITY_TIME_CRITICAL,
    },
};

use crate::{
    audio_capture::{get_wait_error, EventHandleWrapper, RecordingError},
    sample_format::SampleFormat,
};

pub(crate) struct RunContext {
    audio_client: IAudioClient,
    capture_client: IAudioCaptureClient,
    stop_handle: HANDLE,
    format: SampleFormat,
}
unsafe impl Send for RunContext {}

impl RunContext {
    pub(crate) fn new(audio_client: IAudioClient, capture_client: IAudioCaptureClient, stop_handle: HANDLE, format: SampleFormat) -> Self {
        Self {
            audio_client,
            capture_client,
            stop_handle,
            format,
        }
    }
}

pub struct CaptureStream {
    thread: Option<thread::JoinHandle<()>>,
    thread_stop_handle: HANDLE,
}

impl CaptureStream {
    pub(crate) fn start_stream<D, E>(
        data_callback: D,
        mut error_callback: E,
        audio_client: IAudioClient,
        format: SampleFormat,
    ) -> Result<CaptureStream, RecordingError>
    where
        D: FnMut(&[u8]) + Send + 'static,
        E: FnMut(RecordingError) + Send + 'static,
    {
        let capture_client =
            unsafe { audio_client.GetService::<IAudioCaptureClient>() }.map_err(RecordingError::FailedToStartAudioClient)?;
        let stop_handle = unsafe { CreateEventW(None, false, false, None) }.map_err(RecordingError::EventCreationError)?;

        let run_context = RunContext {
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

        Ok(CaptureStream {
            thread: Some(thr),
            thread_stop_handle: stop_handle,
        })
    }

    // See drop implementation for cleanup
    pub fn stop_recording(self) {}

    fn capture_audio<D>(run_context: RunContext, mut data_callback: D) -> Result<(), RecordingError>
    where
        D: FnMut(&[u8]),
    {
        Self::set_thread_priority();
        let (audio_client, capture_client) = (run_context.audio_client, run_context.capture_client);
        let mut buffer: *mut u8 = std::ptr::null_mut();
        let mut flags: u32 = 0;
        let mut pu64deviceposition: u64 = 0;
        let mut pu64qpcposition: u64 = 0;

        let h_event = unsafe { CreateEventA(None, false, false, None) }.map_err(|h| RecordingError::FailedToCreateStopEvent(h))?;
        let h_event = EventHandleWrapper(h_event);
        let handles = [*h_event, run_context.stop_handle];
        unsafe { audio_client.SetEventHandle(*h_event) }.map_err(|h| RecordingError::FailedToSetupEventHandle(h))?;
        unsafe { audio_client.Start() }.map_err(|h| RecordingError::FailedToStartAudioClient(h))?;

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
            .map_err(|h| RecordingError::FailedGettingBuffer(h))?;
            debug_assert!(!buffer.is_null());

            let buf_slice =
                unsafe { std::slice::from_raw_parts(buffer, frames_available as usize * run_context.format.block_align() as usize) };
            data_callback(buf_slice);

            unsafe { capture_client.ReleaseBuffer(frames_available) }.map_err(|h| RecordingError::FailedReleasingBuffer(h))?;
        }
        unsafe {
            audio_client.Stop().map_err(|h| RecordingError::FailedStoppingAudioClient(h))?;
            audio_client.Reset().map_err(|h| RecordingError::FailedResettingAudioClient(h))?;
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

impl Drop for CaptureStream {
    fn drop(&mut self) {
        unsafe {
            let _ = SetEvent(self.thread_stop_handle);
        }
        let _ = self.thread.take().map(|thr| thr.join());
    }
}
