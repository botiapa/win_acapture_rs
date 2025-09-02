use std::{ffi::OsString, ops::Deref, os::windows::ffi::OsStrExt, string::FromUtf16Error};

use thiserror::Error;
use windows::Win32::{
    Devices::Properties,
    Foundation::{self, GetLastError, S_FALSE, S_OK},
    Media::Audio::{
        AUDCLNT_E_UNSUPPORTED_FORMAT, AUDCLNT_SHAREMODE_SHARED, AudioSessionStateActive, AudioSessionStateExpired,
        AudioSessionStateInactive, DEVICE_STATE_ACTIVE, EDataFlow, IAudioSessionControl, IAudioSessionControl2, IAudioSessionEnumerator,
        IAudioSessionManager2, IMMDevice, IMMDeviceCollection, IMMDeviceEnumerator, MMDeviceEnumerator, WAVEFORMATEX, eCapture, eConsole,
        eRender,
    },
    Storage::FileSystem::QueryDosDeviceW,
    System::{
        Com::{self, CLSCTX_ALL, CoCreateInstance, STGM_READ},
        Variant::VT_LPWSTR,
    },
};
use windows_core::{Interface, PCWSTR, PWSTR};

use crate::audio_client::PWSTRWrapper;
use crate::{com::com_initialized, event_args::DeviceState, sample_format::SampleFormat};

#[derive(Error, Debug)]
pub enum AudioError {
    #[error("Device enumeration error: {0}")]
    DeviceEnumError(DeviceEnumError),
    #[error("Failed getting device: {0}")]
    DeviceError(windows::core::Error),
    #[error("Failed activating device: {0}")]
    DeviceActivationError(windows::core::Error),
    #[error("Failed getting session enumerator: {0}")]
    SessionEnumeratorError(windows::core::Error),
    #[error("Failed getting session count: {0}")]
    SessionCountError(windows::core::Error),
    #[error("Failed getting session: {0}")]
    SessionError(windows::core::Error),
    #[error("Failed casting to IAudioSessionControl2: {0}")]
    SessionCastError(windows::core::Error),
    #[error("Failed getting process id: {0}")]
    ProcessIdError(windows::core::Error),
    #[error("Failed getting display name: {0}")]
    DisplayNameError(windows::core::Error),
    #[error("Failed getting state: {0}")]
    GetStateError(windows::core::Error),
    #[error("Failed getting icon path: {0}")]
    IconPathError(windows::core::Error),
    #[error("Failed parsing raw utf16 string: {0}")]
    RawStringParseError(FromUtf16Error),
    #[error("Session not found")]
    GetSessionError(windows::core::Error),
    #[error("Failed to find session with given id")]
    SessionNotFound,
    #[error("Failed reading from property store: {0}")]
    PropertyStoreError(windows::core::Error),
    #[error("Read invalid prop variant")]
    InvalidPropVariant,
    #[error("Failed getting mix format: {0}")]
    FailedGettingMixFormat(windows::core::Error),
    #[error("Failed reading closest format match")]
    FailedReadingClosestFormatMatch,
    #[error("Failed getting volume path name: {0}")]
    FailedGettingVolumePathName(windows::core::Error),
    #[error("Invalid path")]
    InvalidPath,
    #[error("Failed getting dos path: {0}")]
    FailedGettingDosPath(u32),
    #[error("Failed getting nt path: {0}")]
    FailedGettingNtPath(u32),
}

#[derive(Debug, Clone)]
pub struct Session {
    name: String,
    process_name: Option<String>,
    pid: u32,
    is_system: bool,
    session: IAudioSessionControl2,
    session1: IAudioSessionControl,
}

impl PartialEq for Session {
    fn eq(&self, other: &Self) -> bool {
        self.name == other.name
    }
}

impl Session {
    pub fn get_name(&self) -> &String {
        &self.name
    }

    pub fn get_process_name(&self) -> &Option<String> {
        &self.process_name
    }

    pub fn get_pid(&self) -> &u32 {
        &self.pid
    }

    pub fn is_system(&self) -> &bool {
        &self.is_system
    }

    pub fn get_session(&self) -> &IAudioSessionControl2 {
        &self.session
    }

    pub(crate) fn from_session(session: IAudioSessionControl2) -> Result<Self, AudioError> {
        let pid = unsafe { session.GetProcessId() }.map_err(AudioError::ProcessIdError)?;
        let name_pwstr = unsafe { session.GetSessionInstanceIdentifier().map_err(AudioError::DisplayNameError)? };
        let name_pwstr = PWSTRWrapper(name_pwstr);
        let name = unsafe { name_pwstr.0.to_string() }.map_err(AudioError::RawStringParseError)?;
        let process_name = Self::parse_process_name(&name);
        let is_system = unsafe { session.IsSystemSoundsSession() };
        let session1 = session.cast::<IAudioSessionControl>().map_err(AudioError::SessionCastError)?;
        Ok(Self {
            name,
            process_name,
            pid,
            is_system: is_system == S_OK,
            session,
            session1,
        })
    }

    /// Try to parse process name from the session identifier
    /// This is not a good idea, since the session identifier is not guaranteed to be in the same format
    fn parse_process_name(name_string: &String) -> Option<String> {
        Some(name_string.split_once('|')?.1.split_once('%')?.0.into())
    }

    pub fn get_display_name(&self) -> Result<String, AudioError> {
        let display_name = unsafe { self.session1.GetDisplayName() }.map_err(AudioError::DisplayNameError)?;
        let display_name = PWSTRWrapper(display_name);
        Ok(unsafe { display_name.0.to_string() }.unwrap())
    }

    pub fn get_state(&self) -> Result<AudioSessionState, AudioError> {
        let state = unsafe { self.session1.GetState() }.map_err(AudioError::GetStateError)?;
        Ok(state.into())
    }

    pub fn get_icon_path(&self) -> Result<String, AudioError> {
        let icon_path = unsafe { self.session1.GetIconPath() }.map_err(AudioError::IconPathError)?;
        let icon_path = PWSTRWrapper(icon_path);
        Ok(unsafe { icon_path.0.to_string() }.unwrap())
    }
}

struct WaveFormatExPtr(*mut WAVEFORMATEX);

impl Deref for WaveFormatExPtr {
    type Target = *mut WAVEFORMATEX;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

impl Drop for WaveFormatExPtr {
    fn drop(&mut self) {
        unsafe {
            Com::CoTaskMemFree(Some(self.0 as *mut _));
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum FormatSupport {
    Supported,
    Unsupported,
    ClosestMatch(SampleFormat),
}

#[derive(Debug, Clone)]
pub struct Device {
    pub(crate) inner: IMMDevice,
    pub(crate) is_playback: bool,
}

unsafe impl Send for Device {}

impl Device {
    pub fn get_id(&self) -> Result<String, AudioError> {
        let id = unsafe { self.inner.GetId() }.map_err(AudioError::DeviceError)?;
        let id = PWSTRWrapper(id);
        Ok(unsafe { id.0.to_string() }.map_err(AudioError::RawStringParseError)?)
    }

    pub fn get_state(&self) -> Result<DeviceState, AudioError> {
        let state = unsafe { self.inner.GetState() }.map_err(AudioError::GetStateError)?;
        Ok(state.into())
    }

    pub fn get_friendly_name(&self) -> Result<String, AudioError> {
        let prop_key: *const Foundation::PROPERTYKEY = &Properties::DEVPKEY_Device_FriendlyName as *const _ as *const _;
        self.read_string_property(prop_key)
    }

    pub fn get_mix_format(&self) -> Result<SampleFormat, AudioError> {
        com_initialized();
        let audio_client = unsafe { self.inner.Activate::<windows::Win32::Media::Audio::IAudioClient>(CLSCTX_ALL, None) }
            .map_err(AudioError::DeviceActivationError)?;
        let mix_format = unsafe {
            audio_client
                .GetMixFormat()
                .map(WaveFormatExPtr)
                .map_err(AudioError::FailedGettingMixFormat)?
        };
        let mix_format: SampleFormat = SampleFormat::from_wave_format_ex(mix_format.0 as *const WAVEFORMATEX);
        Ok(mix_format)
    }

    pub fn format_supported(&self, format: &SampleFormat) -> Result<FormatSupport, AudioError> {
        com_initialized();
        let audio_client = unsafe { self.inner.Activate::<windows::Win32::Media::Audio::IAudioClient>(CLSCTX_ALL, None) }
            .map_err(AudioError::DeviceActivationError)?;
        let mut closest_match_ptr: *mut WAVEFORMATEX = std::ptr::null_mut();
        let wave_format: WAVEFORMATEX = format.clone().into();
        let hr = unsafe {
            audio_client.IsFormatSupported(
                AUDCLNT_SHAREMODE_SHARED,
                &wave_format,
                Some(&mut closest_match_ptr as *mut *mut WAVEFORMATEX),
            )
        };
        let closest_match = WaveFormatExPtr(closest_match_ptr);

        if hr == S_OK {
            Ok(FormatSupport::Supported)
        } else if hr == AUDCLNT_E_UNSUPPORTED_FORMAT {
            Ok(FormatSupport::Unsupported)
        } else if hr == S_FALSE {
            if closest_match_ptr.is_null() {
                return Err(AudioError::FailedReadingClosestFormatMatch);
            }
            let closest_match: SampleFormat = SampleFormat::from_wave_format_ex(closest_match.0 as *const WAVEFORMATEX);
            Ok(FormatSupport::ClosestMatch(closest_match))
        } else {
            panic!("Unexpected error code: {}", hr);
        }
    }

    pub(crate) fn from(dev: IMMDevice, is_playback: bool) -> Self {
        Self { inner: dev, is_playback }
    }

    fn read_string_property(&self, prop_key: *const Foundation::PROPERTYKEY) -> Result<String, AudioError> {
        let store = unsafe { self.inner.OpenPropertyStore(STGM_READ) }.map_err(AudioError::PropertyStoreError)?;
        let propvar = unsafe { store.GetValue(prop_key).map_err(AudioError::PropertyStoreError)? };
        let propvar = unsafe { &propvar.Anonymous.Anonymous };
        if propvar.vt != VT_LPWSTR {
            return Err(AudioError::InvalidPropVariant);
        }

        let ptr = unsafe { *(&propvar.Anonymous as *const _ as *const *mut u16) };
        let str = PWSTR::from_raw(ptr);
        Ok(unsafe { str.to_string() }.map_err(AudioError::RawStringParseError)?)
    }
}

impl PartialEq for Device {
    fn eq(&self, other: &Self) -> bool {
        match (self.get_id(), other.get_id()) {
            (Ok(id1), Ok(id2)) => id1 == id2,
            _ => false,
        }
    }
}

pub struct SessionManager {}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AudioSessionState {
    AudioSessionStateInactive,
    AudioSessionStateActive,
    AudioSessionStateExpired,
}

impl From<windows::Win32::Media::Audio::AudioSessionState> for AudioSessionState {
    #[allow(non_upper_case_globals)]
    fn from(state: windows::Win32::Media::Audio::AudioSessionState) -> Self {
        match state {
            AudioSessionStateInactive => AudioSessionState::AudioSessionStateInactive,
            AudioSessionStateActive => AudioSessionState::AudioSessionStateActive,
            AudioSessionStateExpired => AudioSessionState::AudioSessionStateExpired,
            _ => panic!("Unknown audio session state"),
        }
    }
}

impl SessionManager {
    /// Queries all active audio sessions
    pub fn get_sessions() -> Result<Vec<Session>, AudioError> {
        com_initialized();
        let dev_collection = Devices::new(eRender).map_err(AudioError::DeviceEnumError)?;

        let mut processes = Vec::new();
        for dev in dev_collection {
            let sessions = AudioSessions::new(dev)?;
            for session in sessions {
                let s = Session::from_session(session)?;
                if !s.is_system() {
                    processes.push(s);
                }
            }
        }
        Ok(processes)
    }

    pub fn session_from_id(searched_id: &str) -> Result<Session, AudioError> {
        let dev_collection = Devices::new(eRender).map_err(AudioError::DeviceEnumError)?;
        // This is a bit inefficient, but it's the only way, I found, to get the session reliably IAudioSessionManager::GetAudioSessionControl wasn't reliable
        // It's still plenty fast, so it's not a big deal (on the order of tenths of microseconds)
        for dev in dev_collection {
            let dev: Device = Device::from(dev, true);
            let sessions = AudioSessions::new(dev.inner)?;
            for session in sessions {
                let id = unsafe {
                    session
                        .GetSessionInstanceIdentifier()
                        .map_err(AudioError::DisplayNameError)?
                        .to_string()
                        .map_err(AudioError::RawStringParseError)?
                };
                if id == searched_id {
                    return Ok(Session::from_session(session)?);
                }
            }
        }
        Err(AudioError::SessionNotFound)
    }
}

const MAX_PATH_LEN: usize = 1024;
/// Gets the NT path (\\Device\\HarddiskVolumeX\\...)
pub fn get_nt_path(path: &str) -> Result<String, AudioError> {
    let (drive_letter, path) = path.split_once(":\\").ok_or(AudioError::InvalidPath)?;
    let drive_letter = format!("{}:\0", drive_letter);
    let drive_letter_u16 = drive_letter.encode_utf16().collect::<Vec<u16>>();
    let drive_letter_wc = PCWSTR::from_raw(drive_letter_u16.as_ptr());
    let mut target_path_u16 = vec![0u16; MAX_PATH_LEN];
    let res = unsafe { QueryDosDeviceW(drive_letter_wc, Some(&mut target_path_u16)) };
    if res == 0 {
        let err = unsafe { GetLastError() };
        return Err(AudioError::FailedGettingNtPath(err.0));
    }
    let target_path_u16 = target_path_u16.into_iter().take_while(|&c| c != 0).collect::<Vec<u16>>(); // Take the first null terminated string
    let target_path = String::from_utf16(&target_path_u16).map_err(AudioError::RawStringParseError)?;
    Ok(format!("{}\\{}", target_path, path))
}

/// Convert the NT path to dos path by mapping drive letters
/// e.g. \\Device\\HarddiskVolume3\\... -> D:\...
pub fn get_dos_path(path: &str) -> Result<String, AudioError> {
    if !path.starts_with("\\Device\\") {
        return Err(AudioError::InvalidPath);
    }

    let parts: Vec<&str> = path.split('\\').collect();
    if parts.len() >= 3 {
        let device_name = parts[2]; // e.g., "HarddiskVolume3"

        for drive_letter in b'A'..=b'Z' {
            let dos_device = format!("{}:", drive_letter as char);
            let mut buffer = [0u16; 512];

            let wide_dos_device: Vec<u16> = OsString::from(dos_device)
                .as_os_str()
                .encode_wide()
                .chain(std::iter::once(0))
                .collect();

            let length = unsafe { QueryDosDeviceW(windows_core::PCWSTR(wide_dos_device.as_ptr()), Some(&mut buffer)) };

            if length > 0 {
                let device_path = String::from_utf16(&buffer[..length as usize - 1]).map_err(AudioError::RawStringParseError)?;

                // Check if this device contains our target device name
                if device_path.contains(device_name) {
                    // Reconstruct the DOS path
                    if parts.len() > 2 {
                        let remaining_path: Vec<&str> = parts[3..].to_vec();
                        return Ok(format!("{}:\\{}", drive_letter as char, remaining_path.join("\\")));
                    } else {
                        return Ok(format!("{}:\\", drive_letter as char));
                    }
                }
            }
        }
        return Err(AudioError::FailedGettingDosPath(2)); // ERROR_FILE_NOT_FOUND
    }
    return Err(AudioError::InvalidPath);
}

#[derive(Error, Debug, Clone)]
pub enum DeviceEnumError {
    #[error("Failed creating enumerator instance: {0}")]
    InstanceCreation(windows::core::Error),
    #[error("Failed enumerating endpoints: {0}")]
    EndpointEnumeration(windows::core::Error),
    #[error("Failed getting device count: {0}")]
    DeviceCountError(windows::core::Error),
    #[error("Failed getting default device: {0}")]
    DefaultDeviceError(windows::core::Error),
}

pub struct DeviceManager {}

impl DeviceManager {
    pub fn get_default_playback_device() -> Result<Device, DeviceEnumError> {
        com_initialized();
        let enumerator: IMMDeviceEnumerator =
            unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL) }.map_err(DeviceEnumError::InstanceCreation)?;
        let dev = unsafe { enumerator.GetDefaultAudioEndpoint(eRender, eConsole) }.map_err(DeviceEnumError::DefaultDeviceError)?;
        Ok(Device::from(dev, true))
    }

    pub fn get_default_input_device() -> Result<Device, DeviceEnumError> {
        com_initialized();
        let enumerator: IMMDeviceEnumerator =
            unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL) }.map_err(DeviceEnumError::InstanceCreation)?;
        let dev = unsafe { enumerator.GetDefaultAudioEndpoint(eCapture, eConsole) }.map_err(DeviceEnumError::DefaultDeviceError)?;
        Ok(Device::from(dev, false))
    }

    pub fn get_playback_devices() -> Result<Vec<Device>, DeviceEnumError> {
        com_initialized();
        let dev_collection = Devices::new(eRender)?;
        Ok(dev_collection.map(|d| Device::from(d, true)).collect())
    }

    pub fn get_capture_devices() -> Result<Vec<Device>, DeviceEnumError> {
        com_initialized();
        let dev_collection = Devices::new(eCapture)?;
        Ok(dev_collection.map(|d| Device::from(d, false)).collect())
    }
}

// Once again, taken from CPAL, thank you!
pub(crate) struct Devices {
    dev_collection: IMMDeviceCollection,
    dev_count: u32,
    next_index: u32,
}

impl Devices {
    pub(crate) fn new(dataflow: EDataFlow) -> Result<Self, DeviceEnumError> {
        let enumerator: IMMDeviceEnumerator =
            unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL) }.map_err(DeviceEnumError::InstanceCreation)?;
        let dev_collection =
            unsafe { enumerator.EnumAudioEndpoints(dataflow, DEVICE_STATE_ACTIVE) }.map_err(DeviceEnumError::EndpointEnumeration)?;
        let dev_count = unsafe { dev_collection.GetCount() }.map_err(DeviceEnumError::DeviceCountError)?;
        Ok(Self {
            dev_collection,
            dev_count,
            next_index: 0,
        })
    }
}

impl Iterator for Devices {
    type Item = IMMDevice;

    fn next(&mut self) -> Option<Self::Item> {
        if self.next_index < self.dev_count {
            let dev = unsafe { self.dev_collection.Item(self.next_index) }.expect("Failed iterating device");
            self.next_index += 1;
            Some(dev)
        } else {
            None
        }
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = (self.dev_count - self.next_index) as usize;
        (remaining, Some(remaining))
    }
}

pub(crate) struct AudioSessions {
    session_enum: IAudioSessionEnumerator,
    session_count: i32,
    next_index: i32,
}

impl AudioSessions {
    pub fn new(device: IMMDevice) -> Result<Self, AudioError> {
        let mgr = unsafe { device.Activate::<IAudioSessionManager2>(CLSCTX_ALL, None) }.map_err(AudioError::DeviceActivationError)?;
        let session_enum = unsafe { mgr.GetSessionEnumerator() }.map_err(AudioError::SessionEnumeratorError)?;
        let session_count = unsafe { session_enum.GetCount() }.map_err(AudioError::SessionCountError)?;
        Ok(Self {
            session_enum,
            session_count,
            next_index: 0,
        })
    }
}

impl Iterator for AudioSessions {
    type Item = IAudioSessionControl2;

    fn next(&mut self) -> Option<Self::Item> {
        if self.next_index < self.session_count {
            let session = unsafe { self.session_enum.GetSession(self.next_index) }.expect("Failed iterating session");
            self.next_index += 1;
            Some(
                session
                    .cast::<IAudioSessionControl2>()
                    .expect("Failed casting to IAudioSessionControl2"),
            )
        } else {
            None
        }
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = (self.session_count - self.next_index) as usize;
        (remaining, Some(remaining))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_sessions() {
        assert!(SessionManager::get_sessions().is_ok());
    }

    #[test]
    fn test_default_format() {
        let devs = DeviceManager::get_capture_devices().unwrap();
        let dev = devs.first().unwrap();
        let format = dev.get_mix_format().unwrap();
        dev.format_supported(&format.into()).unwrap();
    }

    #[test]
    fn test_device() {
        let devs = DeviceManager::get_capture_devices().unwrap();
        let dev = devs.first().unwrap();
        assert!(dev.get_id().is_ok());
        assert!(dev.get_friendly_name().is_ok());
    }
}
