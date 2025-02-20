use std::{
    fmt::Display,
    ops::Deref,
    sync::{atomic::AtomicBool, Arc},
    time::Duration,
};

use crate::{activation_params::SafeActivationParams, audio_stream::AudioStream, sample_format::SampleFormat};
use crate::{com::com_initialized, manager::Device};
use log::{error, trace};
use windows::{
    core::{IUnknown, Interface, GUID, HRESULT},
    Win32::{
        Foundation::{self, CloseHandle, HANDLE, WAIT_EVENT, WAIT_FAILED, WIN32_ERROR},
        Media::{Audio::*, Multimedia::WAVE_FORMAT_IEEE_FLOAT},
        System::{
            Com::{self, CoTaskMemFree, StructuredStorage::PROPVARIANT},
            Threading::{CreateEventW, SetEvent, WaitForSingleObject, INFINITE},
        },
    },
};
use windows_core::implement;

#[derive(Debug, Clone)]
pub enum AudioError {
    FailedToCreateStopEvent(windows_core::Error),
    FailedToSetupEventHandle(windows_core::Error),
    FailedToStartAudioClient(windows_core::Error),
    WaitFailed(WIN32_ERROR),
    FailedGettingBuffer(windows_core::Error),
    FailedReleasingBuffer(windows_core::Error),
    FailedStoppingAudioClient(windows_core::Error),
    FailedResettingAudioClient(windows_core::Error),
    NotInputDevice,
    NotPlaybackDevice,
    RecordingAlreadyStarted,
    FailedGettingActivationResult,
    EventCreationError(windows_core::Error),
}

impl Display for AudioError {
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

pub struct WaveFormatWrapper(*mut WAVEFORMATEX);

impl WaveFormatWrapper {
    fn from_ptr(ptr: *mut WAVEFORMATEX) -> Self {
        Self(ptr)
    }
}

impl Deref for WaveFormatWrapper {
    type Target = *mut WAVEFORMATEX;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

impl Drop for WaveFormatWrapper {
    fn drop(&mut self) {
        unsafe {
            let _ = CoTaskMemFree(Some(self.0 as _));
        }
    }
}

pub struct AudioClient {
    format: SampleFormat,
}

impl AudioClient {
    pub fn new() -> Self {
        Self {
            format: SampleFormat::default(),
        }
    }

    pub fn set_format(&mut self, format: SampleFormat) -> Result<(), AudioError> {
        self.format = format;
        Ok(())
    }

    pub fn get_format(&self) -> SampleFormat {
        self.format.clone()
    }

    /// Start recording audio from a process
    pub fn start_recording_process<D, E>(mut self, pid: u32, data_callback: D, error_callback: E) -> Result<AudioStream, AudioError>
    where
        D: FnMut(&[u8]) + Send + 'static,
        E: FnMut(AudioError) + Send + 'static,
    {
        com_initialized();
        let activate_params = SafeActivationParams::new(pid);

        let res = self.activate_audio_interface(activate_params.prop())?;
        let audio_client = self.activate_loopback_client(&res)?;

        AudioStream::start_capture_stream(data_callback, error_callback, audio_client, self.format)
    }

    /// Start recording audio from an input device
    pub fn start_recording_device<D, E>(mut self, dev: &Device, data_callback: D, error_callback: E) -> Result<AudioStream, AudioError>
    where
        D: FnMut(&[u8]) + Send + 'static,
        E: FnMut(AudioError) + Send + 'static,
    {
        if dev.is_playback {
            return Err(AudioError::NotInputDevice);
        }
        com_initialized();

        let audio_client = self.activate_input_client(dev)?;
        AudioStream::start_capture_stream(data_callback, error_callback, audio_client, self.format)
    }

    pub fn start_playback_device<D, E>(
        mut self,
        dev: &Device,
        data_callback: D,
        error_callback: E,
    ) -> Result<(AudioStream, SampleFormat), AudioError>
    where
        D: FnMut(&mut [u8]) -> bool + Send + 'static,
        E: FnMut(AudioError) + Send + 'static,
    {
        if !dev.is_playback {
            return Err(AudioError::NotPlaybackDevice);
        }
        com_initialized();

        let (format, audio_client) = self.activate_playback_client(dev)?;
        AudioStream::start_playback_stream(data_callback, error_callback, audio_client, self.format)
            .map(|stream| (stream, SampleFormat::from_wave_format_ex(format.0)))
    }

    fn activate_loopback_client(&mut self, res: &IActivateAudioInterfaceAsyncOperation) -> Result<IAudioClient, AudioError> {
        let mut activate_result = HRESULT::default();
        let mut activated_interface: Option<::windows::core::IUnknown> = Option::default();
        unsafe {
            res.GetActivateResult(
                &mut activate_result as *mut HRESULT,
                &mut activated_interface as *mut Option<IUnknown>,
            )
        }
        .map_err(AudioError::FailedToStartAudioClient)?;

        let audio_client = activated_interface
            .ok_or(AudioError::FailedGettingActivationResult)?
            .cast::<IAudioClient>()
            .map_err(AudioError::FailedToStartAudioClient)?;
        let capture_format = self.format.clone().into();
        self.initalize_client(
            audio_client,
            &capture_format,
            AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
            20,
        )
    }

    fn activate_playback_client(&mut self, dev: &Device) -> Result<(WaveFormatWrapper, IAudioClient), AudioError> {
        let audio_client =
            unsafe { dev.inner.Activate::<IAudioClient>(Com::CLSCTX_ALL, None) }.map_err(AudioError::FailedToStartAudioClient)?;
        let format = unsafe { audio_client.GetMixFormat() }.map_err(AudioError::FailedToStartAudioClient)?;
        let format = WaveFormatWrapper::from_ptr(format);
        self.initalize_client(audio_client, *format, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, 0)
            .map(|client| (format, client))
    }

    fn activate_input_client(&mut self, dev: &Device) -> Result<IAudioClient, AudioError> {
        let audio_client =
            unsafe { dev.inner.Activate::<IAudioClient>(Com::CLSCTX_ALL, None) }.map_err(AudioError::FailedToStartAudioClient)?;
        let capture_format: WAVEFORMATEX = self.format.clone().into();
        self.initalize_client(audio_client, &capture_format, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, 20)
    }

    fn initalize_client(
        &mut self,
        audio_client: IAudioClient,
        format: *const WAVEFORMATEX,
        flags: u32,
        buffer_duration_ms: u32,
    ) -> Result<IAudioClient, AudioError> {
        const REFTIME_MS: i64 = 10_000;
        unsafe {
            audio_client.Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                flags,
                REFTIME_MS * buffer_duration_ms as i64,
                0,
                format,
                None,
            )
        }
        .map_err(AudioError::FailedToStartAudioClient)?;

        Ok(audio_client)
    }

    fn activate_audio_interface(&self, activate_params: *const PROPVARIANT) -> Result<IActivateAudioInterfaceAsyncOperation, AudioError> {
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

pub(crate) fn get_wait_error(wait_event: WAIT_EVENT) -> Result<u32, AudioError> {
    if wait_event == WAIT_FAILED {
        let err = unsafe { Foundation::GetLastError() };
        error!("Wait failed: {:?}", err);
        return Err(AudioError::WaitFailed(err));
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

#[cfg(test)]
mod tests {
    use crate::manager::DeviceManager;

    use super::*;

    #[test]
    fn playback() {
        let client = AudioClient::new();
        let dev = DeviceManager::get_default_playback_device().unwrap();
        let (audio_stream, format) = client.start_playback_device(&dev, |data| false, |err| {}).unwrap();
    }
}
