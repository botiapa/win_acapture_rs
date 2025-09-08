use crate::audio_stream::CapturePacket;
use crate::manager::DeviceEnumError;
use crate::{activation_params::SafeActivationParams, audio_stream::AudioStreamConfig, sample_format::SampleFormat};
use crate::{com::com_initialized, manager::Device};
use log::error;
use std::{fmt::Display, ops::Deref, sync::Arc};
use thiserror::Error;
use windows::Win32::System::Com::StringFromIID;
use windows::{
    Win32::{
        Foundation::{self, CloseHandle, HANDLE, WAIT_EVENT, WAIT_FAILED, WIN32_ERROR},
        Media::Audio::*,
        System::{
            Com::{self, CoTaskMemFree, StructuredStorage::PROPVARIANT},
            Threading::{CreateEventW, INFINITE, SetEvent, WaitForSingleObject},
        },
    },
    core::{GUID, HRESULT, IUnknown, Interface},
};
use windows_core::{PWSTR, implement};

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
    FailedToCreateThread,
    StreamAlreadyStarted,
    FailedToGetAudioClock(windows_core::Error),
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

pub(crate) struct PWSTRWrapper(pub(crate) PWSTR);
impl Drop for PWSTRWrapper {
    fn drop(&mut self) {
        unsafe {
            CoTaskMemFree(Some(self.0.0 as _));
        }
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
        D: FnMut(CapturePacket) + Send + 'static,
        E: FnMut(AudioClientError) + Send + 'static,
    {
        com_initialized();
        let activate_params = SafeActivationParams::new(Some(pid));

        let audio_client = self.get_audio_client(VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK, Some(activate_params.prop()))?;
        let capture_format = self.format.clone().unwrap_or_default().into();

        let audio_client = self.initialize_client(
            audio_client,
            &capture_format,
            AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
            BUFFER_DURATION_MS,
        )?;

        let out_format = SampleFormat::from_wave_format_ex(&capture_format);
        AudioStreamConfig::create_capture_stream(data_callback, error_callback, audio_client, Some(out_format))
    }

    /// Start recording audio from an input device
    /// If `dev` is `None`, the default input device will be used
    pub fn start_recording_device<D, E>(
        mut self,
        dev: Option<&Device>,
        data_callback: D,
        error_callback: E,
    ) -> Result<AudioStreamConfig, AudioClientError>
    where
        D: FnMut(CapturePacket) + Send + 'static,
        E: FnMut(AudioClientError) + Send + 'static,
    {
        if let Some(dev) = dev
            && dev.is_playback
        {
            return Err(AudioClientError::NotInputDevice);
        }
        com_initialized();

        let audio_client = self.activate_device_or_default(dev, &DEVINTERFACE_AUDIO_CAPTURE)?;
        let format = match self.format.clone() {
            Some(format) => &mut format.into() as *mut WAVEFORMATEX,
            None => unsafe { audio_client.GetMixFormat() }.map_err(AudioClientError::FailedToGetMixFormat)?,
        };

        let audio_client = self.initialize_client(audio_client, format, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, BUFFER_DURATION_MS)?;

        AudioStreamConfig::create_capture_stream(data_callback, error_callback, audio_client, self.format.clone())
    }

    /// Start recording audio from a loopback device
    /// If `dev` is `None`, the default loopback device will be used
    pub fn start_recording_loopback_device<D, E>(
        mut self,
        dev: Option<&Device>,
        data_callback: D,
        error_callback: E,
    ) -> Result<AudioStreamConfig, AudioClientError>
    where
        D: FnMut(CapturePacket) + Send + 'static,
        E: FnMut(AudioClientError) + Send + 'static,
    {
        if let Some(dev) = dev
            && !dev.is_playback
        {
            return Err(AudioClientError::NotPlaybackDevice);
        }
        com_initialized();

        let audio_client = self.activate_device_or_default(dev, &DEVINTERFACE_AUDIO_RENDER)?;
        let capture_format = unsafe { audio_client.GetMixFormat() }.map_err(AudioClientError::FailedToGetMixFormat)?;
        let audio_client = self.initialize_client(
            audio_client,
            capture_format,
            AUDCLNT_STREAMFLAGS_EVENTCALLBACK | AUDCLNT_STREAMFLAGS_LOOPBACK,
            BUFFER_DURATION_MS,
        )?;

        AudioStreamConfig::create_capture_stream(data_callback, error_callback, audio_client, Some(self.format.unwrap_or_default()))
    }

    /// Start playback on the given device
    /// If `dev` is `None`, the default playback device will be used
    pub fn start_playback_device<D, E>(
        mut self,
        dev: Option<&Device>,
        data_callback: D,
        error_callback: E,
    ) -> Result<(AudioStreamConfig, SampleFormat), AudioClientError>
    where
        D: FnMut(&mut [u8]) -> bool + Send + 'static,
        E: FnMut(AudioClientError) + Send + 'static,
    {
        if let Some(dev) = dev
            && !dev.is_playback
        {
            return Err(AudioClientError::NotPlaybackDevice);
        }
        com_initialized();

        let audio_client = self.activate_device_or_default(dev, &DEVINTERFACE_AUDIO_RENDER)?;
        let format = unsafe { audio_client.GetMixFormat() }.map_err(AudioClientError::FailedToGetMixFormat)?;
        let format = WaveFormatWrapper::from_ptr(format);
        let audio_client = self.initialize_client(audio_client, *format, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, 0)?;

        AudioStreamConfig::create_playback_stream(data_callback, error_callback, audio_client, self.format.unwrap_or_default())
            .map(|stream| (stream, SampleFormat::from_wave_format_ex(format.0)))
    }

    fn activate_device_or_default(&self, dev: Option<&Device>, default_iid: &windows_core::GUID) -> Result<IAudioClient, AudioClientError> {
        match dev {
            Some(dev) => {
                unsafe { dev.inner.Activate::<IAudioClient>(Com::CLSCTX_ALL, None) }.map_err(AudioClientError::FailedToStartAudioClient)
            }
            None => {
                let audio_render_guid = unsafe { StringFromIID(default_iid).expect("can only fail on OOM") };
                let audio_render_guid = PWSTRWrapper(audio_render_guid);
                self.get_audio_client(audio_render_guid.0, None)
            }
        }
    }

    fn initialize_client(
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

    fn get_audio_client<P>(
        &self,
        device_interface_path: P,
        activate_params: Option<*const PROPVARIANT>,
    ) -> Result<IAudioClient, AudioClientError>
    where
        P: windows_core::Param<windows_core::PCWSTR>,
    {
        let activate_event = unsafe { CreateEventW(None, false, false, None) }.expect("Failed to create event handle");
        let activate_event = Arc::new(EventHandleWrapper(activate_event));
        let handler: IActivateAudioInterfaceCompletionHandler = ActivateHandler::new(activate_event.clone()).into();
        let res =
            unsafe { ActivateAudioInterfaceAsync(device_interface_path, &IAudioClient::IID as *const GUID, activate_params, &handler) }
                .expect("ActivateAudioInterfaceAsync failed");

        unsafe { get_wait_error(WaitForSingleObject(**activate_event, INFINITE))? };

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
        Ok(audio_client)
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
    use super::*;
    use std::sync::mpsc::channel;
    use std::time::Duration;

    #[test]
    fn playback() {
        let client = AudioClient::new();
        let (err_sender, err_recv) = channel();
        let (audio_stream, _format) = client
            .start_playback_device(None, |_data| false, move |err| err_sender.send(err).unwrap())
            .unwrap();
        audio_stream.start().unwrap();

        if let Some(err) = err_recv.recv_timeout(Duration::from_millis(10)).ok() {
            panic!("Error during playback: {:?}", err);
        }
    }

    #[test]
    fn process_capture() {
        let rendering_client = AudioClient::new();
        let (audio_stream_config, _format) = rendering_client.start_playback_device(None, |_data| false, |_err| {}).unwrap();
        audio_stream_config.start().unwrap();

        let client = AudioClient::new();
        let (err_sender, err_recv) = channel();
        let audio_stream_config_capture = client
            .start_recording_process(std::process::id(), |_data| {}, move |err| err_sender.send(err).unwrap())
            .unwrap();
        audio_stream_config_capture.start().unwrap();

        if let Some(err) = err_recv.recv_timeout(Duration::from_millis(10)).ok() {
            panic!("Error during process cap: {:?}", err);
        }
    }
}
