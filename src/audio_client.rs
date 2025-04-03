use std::{fmt::Display, ops::Deref, sync::Arc};

use crate::manager::{DeviceEnumError, DeviceManager};
use crate::{activation_params::SafeActivationParams, audio_stream::AudioStreamConfig, sample_format::SampleFormat};
use crate::{com::com_initialized, manager::Device};
use log::error;
use thiserror::Error;
use windows::{
    core::{IUnknown, Interface, GUID, HRESULT},
    Win32::{
        Foundation::{self, CloseHandle, HANDLE, WAIT_EVENT, WAIT_FAILED, WIN32_ERROR},
        Media::Audio::*,
        System::{
            Com::{self, CoTaskMemFree, StructuredStorage::PROPVARIANT},
            Threading::{CreateEventW, SetEvent, WaitForSingleObject, INFINITE},
        },
    },
};
use windows_core::implement;

#[derive(Error, Debug, Clone)]
pub enum AudioClientError {
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
    DeviceEnumError(DeviceEnumError),
    FailedToGetMixFormat(windows_core::Error),
    StreamAlreadyStarted,
}

impl Display for AudioClientError {
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

const BUFFER_DURATION_MS: u32 = 20;

pub struct AudioClient {
    format: Option<SampleFormat>,
}

impl AudioClient {
    pub fn new() -> Self {
        Self { format: None }
    }

    pub fn set_format(&mut self, format: SampleFormat) -> Result<(), AudioClientError> {
        self.format = Some(format);
        Ok(())
    }

    pub fn get_format(&self) -> Option<SampleFormat> {
        self.format.clone()
    }

    /// Start recording audio from a process
    pub fn start_recording_process<D, E>(
        mut self,
        pid: u32,
        data_callback: D,
        error_callback: E,
    ) -> Result<AudioStreamConfig, AudioClientError>
    where
        D: FnMut(&[u8]) + Send + 'static,
        E: FnMut(AudioClientError) + Send + 'static,
    {
        com_initialized();
        let activate_params = SafeActivationParams::new(Some(pid));

        let res = self.activate_audio_interface(activate_params.prop())?;
        let audio_client = self.activate_process_capture_client(&res)?;

        AudioStreamConfig::create_capture_stream(data_callback, error_callback, audio_client, self.format.clone())
    }

    /// Start recording audio from the default loopback device
    pub fn start_recording_default_loopback<D, E>(self, data_callback: D, error_callback: E) -> Result<AudioStreamConfig, AudioClientError>
    where
        D: FnMut(&[u8]) + Send + 'static,
        E: FnMut(AudioClientError) + Send + 'static,
    {
        com_initialized();
        let dev = DeviceManager::get_default_playback_device().map_err(AudioClientError::DeviceEnumError)?;

        self.start_recording_loopback_device(&dev, data_callback, error_callback)
    }

    pub fn start_recording_default_input<D, E>(self, data_callback: D, error_callback: E) -> Result<AudioStreamConfig, AudioClientError>
    where
        D: FnMut(&[u8]) + Send + 'static,
        E: FnMut(AudioClientError) + Send + 'static,
    {
        com_initialized();
        let dev = DeviceManager::get_default_input_device().map_err(AudioClientError::DeviceEnumError)?;

        self.start_recording_device(&dev, data_callback, error_callback)
    }

    /// Start recording audio from an input device
    pub fn start_recording_device<D, E>(
        mut self,
        dev: &Device,
        data_callback: D,
        error_callback: E,
    ) -> Result<AudioStreamConfig, AudioClientError>
    where
        D: FnMut(&[u8]) + Send + 'static,
        E: FnMut(AudioClientError) + Send + 'static,
    {
        if dev.is_playback {
            return Err(AudioClientError::NotInputDevice);
        }
        com_initialized();

        let audio_client =
            unsafe { dev.inner.Activate::<IAudioClient>(Com::CLSCTX_ALL, None) }.map_err(AudioClientError::FailedToStartAudioClient)?;
        let format = match self.format.clone() {
            Some(format) => &mut format.into() as *mut WAVEFORMATEX,
            None => unsafe { audio_client.GetMixFormat() }.map_err(AudioClientError::FailedToGetMixFormat)?,
        };

        let audio_client = self.initalize_client(audio_client, format, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, BUFFER_DURATION_MS)?;

        AudioStreamConfig::create_capture_stream(data_callback, error_callback, audio_client, self.format.clone())
    }

    pub fn start_recording_loopback_device<D, E>(
        mut self,
        dev: &Device,
        data_callback: D,
        error_callback: E,
    ) -> Result<AudioStreamConfig, AudioClientError>
    where
        D: FnMut(&[u8]) + Send + 'static,
        E: FnMut(AudioClientError) + Send + 'static,
    {
        if !dev.is_playback {
            return Err(AudioClientError::NotPlaybackDevice);
        }
        com_initialized();

        let audio_client =
            unsafe { dev.inner.Activate::<IAudioClient>(Com::CLSCTX_ALL, None) }.map_err(AudioClientError::FailedToStartAudioClient)?;
        let capture_format = unsafe { audio_client.GetMixFormat() }.map_err(AudioClientError::FailedToGetMixFormat)?;
        let audio_client = self.initalize_client(
            audio_client,
            capture_format,
            AUDCLNT_STREAMFLAGS_EVENTCALLBACK | AUDCLNT_STREAMFLAGS_LOOPBACK,
            BUFFER_DURATION_MS,
        )?;

        AudioStreamConfig::create_capture_stream(data_callback, error_callback, audio_client, None)
    }

    pub fn start_playback_default_device<D, E>(
        self,
        data_callback: D,
        error_callback: E,
    ) -> Result<(AudioStreamConfig, SampleFormat), AudioClientError>
    where
        D: FnMut(&mut [u8]) -> bool + Send + 'static,
        E: FnMut(AudioClientError) + Send + 'static,
    {
        let dev = DeviceManager::get_default_playback_device().map_err(AudioClientError::DeviceEnumError)?;
        self.start_playback_device(&dev, data_callback, error_callback)
    }

    pub fn start_playback_device<D, E>(
        mut self,
        dev: &Device,
        data_callback: D,
        error_callback: E,
    ) -> Result<(AudioStreamConfig, SampleFormat), AudioClientError>
    where
        D: FnMut(&mut [u8]) -> bool + Send + 'static,
        E: FnMut(AudioClientError) + Send + 'static,
    {
        if !dev.is_playback {
            return Err(AudioClientError::NotPlaybackDevice);
        }
        com_initialized();

        let (format, audio_client) = self.activate_playback_client(dev)?;
        AudioStreamConfig::create_playback_stream(data_callback, error_callback, audio_client, self.format.unwrap_or_default())
            .map(|stream| (stream, SampleFormat::from_wave_format_ex(format.0)))
    }

    fn activate_process_capture_client(&mut self, res: &IActivateAudioInterfaceAsyncOperation) -> Result<IAudioClient, AudioClientError> {
        let mut activate_result = HRESULT::default();
        let mut activated_interface: Option<::windows::core::IUnknown> = Option::default();
        unsafe {
            res.GetActivateResult(
                &mut activate_result as *mut HRESULT,
                &mut activated_interface as *mut Option<IUnknown>,
            )
        }
        .map_err(AudioClientError::FailedToStartAudioClient)?;

        let audio_client = activated_interface
            .ok_or(AudioClientError::FailedGettingActivationResult)?
            .cast::<IAudioClient>()
            .map_err(AudioClientError::FailedToStartAudioClient)?;
        let capture_format = self.format.clone().unwrap_or_default().into();
        self.initalize_client(
            audio_client,
            &capture_format,
            AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
            BUFFER_DURATION_MS,
        )
    }

    fn activate_playback_client(&mut self, dev: &Device) -> Result<(WaveFormatWrapper, IAudioClient), AudioClientError> {
        let audio_client =
            unsafe { dev.inner.Activate::<IAudioClient>(Com::CLSCTX_ALL, None) }.map_err(AudioClientError::FailedToStartAudioClient)?;
        let format = unsafe { audio_client.GetMixFormat() }.map_err(AudioClientError::FailedToGetMixFormat)?;
        let format = WaveFormatWrapper::from_ptr(format);
        self.initalize_client(audio_client, *format, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, 0)
            .map(|client| (format, client))
    }

    fn initalize_client(
        &mut self,
        audio_client: IAudioClient,
        format: *const WAVEFORMATEX,
        flags: u32,
        buffer_duration_ms: u32,
    ) -> Result<IAudioClient, AudioClientError> {
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
        .map_err(AudioClientError::FailedToStartAudioClient)?;

        Ok(audio_client)
    }

    fn activate_audio_interface(
        &self,
        activate_params: *const PROPVARIANT,
    ) -> Result<IActivateAudioInterfaceAsyncOperation, AudioClientError> {
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

pub(crate) fn get_wait_error(wait_event: WAIT_EVENT) -> Result<u32, AudioClientError> {
    if wait_event == WAIT_FAILED {
        let err = unsafe { Foundation::GetLastError() };
        error!("Wait failed: {:?}", err);
        return Err(AudioClientError::WaitFailed(err));
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
        let (audio_stream, format) = client.start_playback_default_device(|data| false, |err| {}).unwrap();
    }
}
