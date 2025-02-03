use std::{
    fmt::Display,
    sync::{atomic::AtomicBool, Arc},
    thread,
    time::Duration,
};

use crate::{activation_params::SafeActivationParams, sample_format::SampleFormat};
use crate::{com::com_initialized, manager::Device};
use log::error;
use windows::{
    core::{IUnknown, Interface, GUID, HRESULT},
    Win32::{
        Foundation::{self, HANDLE, WAIT_FAILED, WAIT_OBJECT_0, WIN32_ERROR},
        Media::Audio::*,
        System::{
            Com::{self, StructuredStorage::PROPVARIANT, CLSCTX},
            Threading::{
                CreateEventA, CreateEventW, GetCurrentThread, SetEvent, SetThreadPriority, WaitForMultipleObjectsEx, INFINITE,
                THREAD_PRIORITY_TIME_CRITICAL,
            },
        },
    },
};
use windows_core::implement;

pub struct AudioCapture {
    _activate_completed: Arc<AtomicBool>,
    _handler: IActivateAudioInterfaceCompletionHandler,
    _thread: Option<thread::JoinHandle<()>>,
    _thread_stop_handle: Option<HANDLE>,
    _format: SampleFormat,
}

struct RunContext {
    audio_client: IAudioClient,
    capture_client: IAudioCaptureClient,
    stop_handle: HANDLE,
}
unsafe impl Send for RunContext {}

#[derive(Debug, Clone)]
pub enum RecordingError {
    FailedToCreateStopEvent(windows_core::Error),
    FailedToSetupEventHandle(windows_core::Error),
    FailedToStartAudioClient(windows_core::Error),
    WaitFailed(WIN32_ERROR),
    FailedGettingBuffer(windows_core::Error),
    FailedReleasingBuffer(windows_core::Error),
    FailedStoppingAudioClient(windows_core::Error),
    FailedResettingAudioClient(windows_core::Error),
    AlreadyRecording,
}

impl Display for RecordingError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Recording error: {:?}", self)
    }
}

impl AudioCapture {
    pub fn new() -> Self {
        let _activate_completed = Arc::new(AtomicBool::new(false));

        Self {
            _activate_completed: _activate_completed.clone(),
            _handler: ActivateHandler::new(_activate_completed).into(),
            _thread: None,
            _thread_stop_handle: None,
            _format: SampleFormat::default(),
        }
    }

    pub fn set_format(&mut self, format: SampleFormat) -> Result<(), RecordingError> {
        if self._activate_completed.load(std::sync::atomic::Ordering::Relaxed) {
            return Err(RecordingError::AlreadyRecording);
        }
        self._format = format;
        Ok(())
    }

    pub fn get_format(&self) -> SampleFormat {
        self._format.clone()
    }

    /// Start recording audio from a process
    pub fn start_recording_process<D, E>(&mut self, pid: u32, data_callback: D, mut error_callback: E)
    where
        D: FnMut(&[u8]) + Send + 'static,
        E: FnMut(RecordingError) + Send + 'static,
    {
        if self._activate_completed.load(std::sync::atomic::Ordering::Relaxed) {
            panic!("Can only start one recording on each instance of AudioCapture");
        }
        com_initialized();
        let activate_params = SafeActivationParams::new(pid);

        let res = self.activate_audio_interface(activate_params.prop());
        let (audio_client, capture_client) = self.init_loopback_clients(&res);

        let stop_handle = unsafe { CreateEventW(None, false, false, None) }.expect("Failed creating stop event");
        self._thread_stop_handle = Some(stop_handle);

        let run_context = RunContext {
            audio_client,
            capture_client,
            stop_handle,
        };

        let thr = thread::spawn(move || {
            let res = Self::capture_audio(run_context, data_callback);
            if res.is_err() {
                error_callback(res.unwrap_err());
            }
        });
        self._thread = Some(thr);
    }

    /// Start recording audio from an input device
    pub fn start_recording_device<D, E>(&mut self, dev: &Device, data_callback: D, mut error_callback: E)
    where
        D: FnMut(&[u8]) + Send + 'static,
        E: FnMut(RecordingError) + Send + 'static,
    {
        if dev.is_playback {
            panic!("Device is not an input device");
        }
        if self._activate_completed.load(std::sync::atomic::Ordering::Relaxed) {
            panic!("Can only start one recording on each instance of AudioCapture");
        }
        com_initialized();

        let (audio_client, capture_client) = self.init_input_clients(dev);

        let stop_handle = unsafe { CreateEventW(None, false, false, None) }.expect("Failed creating stop event");
        self._thread_stop_handle = Some(stop_handle);

        let run_context = RunContext {
            audio_client,
            capture_client,
            stop_handle,
        };

        let thr = thread::spawn(move || {
            let res = Self::capture_audio(run_context, data_callback);
            if res.is_err() {
                error_callback(res.unwrap_err());
            }
        });
        self._thread = Some(thr);
    }

    pub fn stop_recording(&mut self) {
        unsafe {
            if let Some(stop_handle) = self._thread_stop_handle {
                let _ = SetEvent(stop_handle);
            }
        }
        self._thread.take().map(|thr| thr.join().unwrap());
    }

    fn init_loopback_clients(&mut self, res: &IActivateAudioInterfaceAsyncOperation) -> (IAudioClient, IAudioCaptureClient) {
        let mut activate_result = HRESULT::default();
        let mut activated_interface: Option<::windows::core::IUnknown> = Option::default();
        unsafe {
            res.GetActivateResult(
                &mut activate_result as *mut HRESULT,
                &mut activated_interface as *mut Option<IUnknown>,
            )
        }
        .expect("Failed getting activate result");

        let audio_client = activated_interface.unwrap().cast::<IAudioClient>().unwrap();
        let capture_format = self._format.clone().into();

        //let capture_format = self._format.clone().to_wave_format_ex();

        unsafe {
            audio_client.Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM | AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
                200000,
                0,
                &capture_format as *const WAVEFORMATEX,
                None,
            )
        }
        .expect("Failed initializing audio client");

        let capture_client = unsafe { audio_client.GetService::<IAudioCaptureClient>() }.unwrap();
        (audio_client, capture_client)
    }

    fn init_input_clients(&mut self, dev: &Device) -> (IAudioClient, IAudioCaptureClient) {
        let audio_client = unsafe { dev.inner.Activate::<IAudioClient>(Com::CLSCTX_ALL, None) }.expect("Failed activating device");
        let capture_format: WAVEFORMATEX = self._format.clone().into();

        unsafe {
            audio_client.Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM | AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
                200000,
                0,
                &capture_format,
                None,
            )
        }
        .expect("Failed initializing audio client");

        let capture_client = unsafe { audio_client.GetService::<IAudioCaptureClient>() }.unwrap();
        (audio_client, capture_client)
    }

    fn set_thread_priority() {
        unsafe {
            let curr_thr = GetCurrentThread();
            let _ = SetThreadPriority(curr_thr, THREAD_PRIORITY_TIME_CRITICAL);
        }
    }

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
        let handles = [h_event, run_context.stop_handle];
        unsafe { audio_client.SetEventHandle(h_event) }.map_err(|h| RecordingError::FailedToSetupEventHandle(h))?;
        unsafe { audio_client.Start() }.map_err(|h| RecordingError::FailedToStartAudioClient(h))?;

        while let Ok(mut frames_available) = unsafe { capture_client.GetNextPacketSize() } {
            let wait_res = unsafe { WaitForMultipleObjectsEx(&handles, false, INFINITE, false) };

            if wait_res == WAIT_FAILED {
                let err = unsafe { Foundation::GetLastError() };
                error!("Wait failed: {:?}", err);
                return Err(RecordingError::WaitFailed(err));
            }

            // Stop event was called
            if wait_res.0 == WAIT_OBJECT_0.0 + 1 {
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

            let buf_slice = unsafe { std::slice::from_raw_parts(buffer, frames_available as usize * 4) };
            data_callback(buf_slice);

            unsafe { capture_client.ReleaseBuffer(frames_available) }.map_err(|h| RecordingError::FailedReleasingBuffer(h))?;
        }
        unsafe {
            audio_client.Stop().map_err(|h| RecordingError::FailedStoppingAudioClient(h))?;
            audio_client.Reset().map_err(|h| RecordingError::FailedResettingAudioClient(h))?;
        }
        Ok(())
    }

    fn activate_audio_interface(&self, activate_params: *const PROPVARIANT) -> IActivateAudioInterfaceAsyncOperation {
        let res = unsafe {
            ActivateAudioInterfaceAsync(
                VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK,
                &IAudioClient::IID as *const GUID,
                Some(activate_params as *const PROPVARIANT),
                &self._handler,
            )
        }
        .expect("ActivateAudioInterfaceAsync failed");

        while !self._activate_completed.load(std::sync::atomic::Ordering::Relaxed) {
            std::thread::sleep(Duration::from_millis(10));
        }

        res
    }
}

#[implement(IActivateAudioInterfaceCompletionHandler)]
struct ActivateHandler {
    activate_completed: Arc<AtomicBool>,
}

impl ActivateHandler {
    fn new(activate_completed: Arc<AtomicBool>) -> Self {
        Self {
            activate_completed: activate_completed,
        }
    }
}

impl IActivateAudioInterfaceCompletionHandler_Impl for ActivateHandler_Impl {
    fn ActivateCompleted(&self, _: windows_core::Ref<'_, IActivateAudioInterfaceAsyncOperation>) -> windows::core::Result<()> {
        self.activate_completed.store(true, std::sync::atomic::Ordering::Relaxed);
        Ok(())
    }
}
