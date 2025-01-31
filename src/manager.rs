use std::{string::FromUtf16Error, time::Instant};

use thiserror::Error;
use windows::Win32::{
    Foundation::S_OK,
    Media::Audio::{
        eRender, AudioSessionStateActive, AudioSessionStateExpired, AudioSessionStateInactive, IAudioSessionControl, IAudioSessionControl2,
        IAudioSessionEnumerator, IAudioSessionManager2, IMMDevice, IMMDeviceCollection, IMMDeviceEnumerator, In, MMDeviceEnumerator,
        DEVICE_STATE_ACTIVE,
    },
    System::Com::{CoCreateInstance, CLSCTX_ALL},
};
use windows_core::{Interface, GUID, PWSTR};

use crate::com::com_initialized;

#[derive(Error, Debug)]
pub enum AudioError {
    #[error("Failed creating instance: {0}")]
    InstanceCreationError(windows::core::Error),
    #[error("Failed getting device collection: {0}")]
    DeviceCollectionError(windows::core::Error),
    #[error("Failed getting device count: {0}")]
    DeviceCountError(windows::core::Error),
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
}

#[derive(Debug, Clone, PartialEq)]
pub struct SafeSessionId(pub(crate) PWSTR);
unsafe impl Send for SafeSessionId {}

#[derive(Debug, Clone)]
pub struct Session {
    name: PWSTR,
    process_name: Option<String>,
    pid: u32,
    is_system: bool,
    session: IAudioSessionControl2,
    session1: IAudioSessionControl,
}

impl PartialEq for Session {
    fn eq(&self, other: &Self) -> bool {
        let n1 = unsafe { self.name.to_string() };
        let n2 = unsafe { other.name.to_string() };
        if n1.is_err() || n2.is_err() {
            return false;
        }
        n1.unwrap() == n2.unwrap()
    }
}

impl Session {
    pub fn get_id(&self) -> SafeSessionId {
        SafeSessionId(self.name)
    }

    pub fn get_name(&self) -> &PWSTR {
        &self.name
    }

    pub fn get_name_string(&self) -> Result<String, AudioError> {
        Ok(unsafe { self.name.to_string() }.map_err(AudioError::RawStringParseError)?)
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
        let process_name = Self::parse_process_name(name_pwstr);
        let is_system = unsafe { session.IsSystemSoundsSession() };
        let session1 = session.cast::<IAudioSessionControl>().map_err(AudioError::SessionCastError)?;
        Ok(Self {
            name: name_pwstr,
            process_name,
            pid,
            is_system: is_system == S_OK,
            session,
            session1,
        })
    }

    /// Try to parse process name from the session identifier
    /// This is not a good idea, since the session identifier is not guaranteed to be in the same format
    fn parse_process_name(name_pwstr: PWSTR) -> Option<String> {
        Some(unsafe { name_pwstr.to_string() }.ok()?.split_once('|')?.1.split_once('%')?.0.into())
    }

    pub fn get_display_name(&self) -> Result<String, AudioError> {
        let display_name = unsafe { self.session1.GetDisplayName() }.map_err(AudioError::DisplayNameError)?;
        Ok(unsafe { display_name.to_string() }.unwrap())
    }

    pub fn get_state(&self) -> Result<AudioSessionState, AudioError> {
        let state = unsafe { self.session1.GetState() }.map_err(AudioError::GetStateError)?;
        Ok(state.into())
    }

    pub fn get_icon_path(&self) -> Result<String, AudioError> {
        let icon_path = unsafe { self.session1.GetIconPath() }.map_err(AudioError::IconPathError)?;
        Ok(unsafe { icon_path.to_string() }.unwrap())
    }
}

pub struct Device {
    pub(crate) inner: IMMDevice,
}

unsafe impl Send for Device {}

impl Device {
    pub fn get_id(&self) -> Result<String, AudioError> {
        let id = unsafe { self.inner.GetId() }.map_err(AudioError::DeviceError)?;
        Ok(unsafe { id.to_string() }.map_err(AudioError::RawStringParseError)?)
    }
}

impl From<IMMDevice> for Device {
    fn from(dev: IMMDevice) -> Self {
        Self { inner: dev }
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
        let dev_collection = Devices::new()?;

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

    pub fn session_from_id(searched_id: &SafeSessionId) -> Result<Session, AudioError> {
        let dev_collection = Devices::new()?;
        let searched_id = unsafe { searched_id.0.to_string() }.map_err(AudioError::RawStringParseError)?;
        // This is a bit inefficient, but it's the only way, I found, to get the session reliably IAudioSessionManager::GetAudioSessionControl wasn't reliable
        // It's still plenty fast, so it's not a big deal (on the order of tenths of microseconds)
        for dev in dev_collection {
            let dev: Device = dev.into();
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

pub struct DeviceManager {}

impl DeviceManager {
    pub fn get_devices() -> Result<Vec<Device>, AudioError> {
        com_initialized();
        let dev_collection = Devices::new()?;
        Ok(dev_collection.map(Device::from).collect())
    }
}

// Once again, taken from CPAL, thank you!
pub(crate) struct Devices {
    dev_collection: IMMDeviceCollection,
    dev_count: u32,
    next_index: u32,
}

impl Devices {
    pub fn new() -> Result<Self, AudioError> {
        let enumerator: IMMDeviceEnumerator =
            unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL) }.map_err(AudioError::InstanceCreationError)?;
        let dev_collection =
            unsafe { enumerator.EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE) }.map_err(AudioError::DeviceCollectionError)?;
        let dev_count = unsafe { dev_collection.GetCount() }.map_err(AudioError::DeviceCountError)?;
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
        com_initialized();

        assert!(SessionManager::get_sessions().is_ok());
    }
}
