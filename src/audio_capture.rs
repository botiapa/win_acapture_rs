use std::{
    fmt::Display,
    ops::Deref,
    sync::{atomic::AtomicBool, Arc},
    time::Duration,
};

use crate::{activation_params::SafeActivationParams, capture_stream::CaptureStream, sample_format::SampleFormat};
use crate::{com::com_initialized, manager::Device};
use log::{error, trace};
use windows::{
    core::{IUnknown, Interface, GUID, HRESULT},
    Win32::{
        Foundation::{self, CloseHandle, HANDLE, WAIT_EVENT, WAIT_FAILED, WIN32_ERROR},
        Media::Audio::*,
        System::{
            Com::{self, StructuredStorage::PROPVARIANT},
            Threading::{CreateEventW, SetEvent, WaitForSingleObject, INFINITE},
        },
    },
};
use windows_core::implement;

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
    NotInputDevice,
    RecordingAlreadyStarted,
    FailedGettingActivationResult,
    EventCreationError(windows_core::Error),
}

impl Display for RecordingError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Recording error: {:?}", self)
    }
}

pub struct EventHandleWrapper(pub(crate) HANDLE);

impl Drop for EventHandleWrapper {
    fn drop(&mut self) {
        unsafe {
            let _ = CloseHandle(self.0);
        };
    }
}

impl Deref for EventHandleWrapper {
    type Target = HANDLE;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

pub struct AudioCapture {
    format: SampleFormat,
}

impl AudioCapture {
    pub fn new() -> Self {
        Self {
            format: SampleFormat::default(),
        }
    }

    pub fn set_format(&mut self, format: SampleFormat) -> Result<(), RecordingError> {
        self.format = format;
        Ok(())
    }

    pub fn get_format(&self) -> SampleFormat {
        self.format.clone()
    }

    /// Start recording audio from a process
    pub fn start_recording_process<D, E>(mut self, pid: u32, data_callback: D, error_callback: E) -> Result<CaptureStream, RecordingError>
    where
        D: FnMut(&[u8]) + Send + 'static,
        E: FnMut(RecordingError) + Send + 'static,
    {
        com_initialized();
        let activate_params = SafeActivationParams::new(pid);

        let res = self.activate_audio_interface(activate_params.prop())?;
        let audio_client = self.activate_loopback_client(&res)?;

        CaptureStream::start_stream(data_callback, error_callback, audio_client, self.format)
    }

    /// Start recording audio from an input device
    pub fn start_recording_device<D, E>(
        mut self,
        dev: &Device,
        data_callback: D,
        error_callback: E,
    ) -> Result<CaptureStream, RecordingError>
    where
        D: FnMut(&[u8]) + Send + 'static,
        E: FnMut(RecordingError) + Send + 'static,
    {
        if dev.is_playback {
            return Err(RecordingError::NotInputDevice);
        }
        com_initialized();

        let audio_client = self.activate_input_client(dev)?;
        CaptureStream::start_stream(data_callback, error_callback, audio_client, self.format)
    }

    fn activate_loopback_client(&mut self, res: &IActivateAudioInterfaceAsyncOperation) -> Result<IAudioClient, RecordingError> {
        let mut activate_result = HRESULT::default();
        let mut activated_interface: Option<::windows::core::IUnknown> = Option::default();
        unsafe {
            res.GetActivateResult(
                &mut activate_result as *mut HRESULT,
                &mut activated_interface as *mut Option<IUnknown>,
            )
        }
        .map_err(RecordingError::FailedToStartAudioClient)?;

        let audio_client = activated_interface
            .ok_or(RecordingError::FailedGettingActivationResult)?
            .cast::<IAudioClient>()
            .map_err(RecordingError::FailedToStartAudioClient)?;
        let capture_format = self.format.clone().into();
        self.initalize_client(
            audio_client,
            capture_format,
            AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
        )
    }

    fn activate_input_client(&mut self, dev: &Device) -> Result<IAudioClient, RecordingError> {
        let audio_client =
            unsafe { dev.inner.Activate::<IAudioClient>(Com::CLSCTX_ALL, None) }.map_err(RecordingError::FailedToStartAudioClient)?;
        let capture_format: WAVEFORMATEX = self.format.clone().into();
        self.initalize_client(audio_client, capture_format, AUDCLNT_STREAMFLAGS_EVENTCALLBACK)
    }

    fn initalize_client(
        &mut self,
        audio_client: IAudioClient,
        capture_format: WAVEFORMATEX,
        flags: u32,
    ) -> Result<IAudioClient, RecordingError> {
        unsafe { audio_client.Initialize(AUDCLNT_SHAREMODE_SHARED, flags, 200000, 0, &capture_format, None) }
            .map_err(RecordingError::FailedToStartAudioClient)?;

        Ok(audio_client)
    }

    fn activate_audio_interface(
        &self,
        activate_params: *const PROPVARIANT,
    ) -> Result<IActivateAudioInterfaceAsyncOperation, RecordingError> {
        let activate_event = unsafe { CreateEventW(None, false, false, None) }.expect("Failed to create event handle");
        let activate_event = Arc::new(EventHandleWrapper(activate_event));
        let handler: IActivateAudioInterfaceCompletionHandler = ActivateHandler::new(activate_event.clone()).into();
        let res = unsafe {
            ActivateAudioInterfaceAsync(
                VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK,
                &IAudioClient::IID as *const GUID,
                Some(activate_params as *const PROPVARIANT),
                &handler,
            )
        }
        .expect("ActivateAudioInterfaceAsync failed");

        unsafe { get_wait_error(WaitForSingleObject(**activate_event, INFINITE))? };
        Ok(res)
    }
}

pub(crate) fn get_wait_error(wait_event: WAIT_EVENT) -> Result<u32, RecordingError> {
    if wait_event == WAIT_FAILED {
        let err = unsafe { Foundation::GetLastError() };
        error!("Wait failed: {:?}", err);
        return Err(RecordingError::WaitFailed(err));
    }
    Ok(wait_event.0)
}

#[implement(IActivateAudioInterfaceCompletionHandler)]
struct ActivateHandler {
    activate_event: Arc<EventHandleWrapper>,
}

impl ActivateHandler {
    fn new(activate_completed: Arc<EventHandleWrapper>) -> Self {
        Self {
            activate_event: activate_completed,
        }
    }
}

impl IActivateAudioInterfaceCompletionHandler_Impl for ActivateHandler_Impl {
    fn ActivateCompleted(&self, _: windows_core::Ref<'_, IActivateAudioInterfaceAsyncOperation>) -> windows::core::Result<()> {
        unsafe { SetEvent(self.activate_event.0)? }
        Ok(())
    }
}
