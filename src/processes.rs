use thiserror::Error;
use windows::Win32::{
    Foundation::{PROPERTYKEY, S_OK},
    Media::Audio::{
        eRender, AudioSessionState, EDataFlow, ERole, IAudioSessionControl, IAudioSessionControl2, IAudioSessionEnumerator,
        IAudioSessionManager2, IControlInterface_Impl, IMMDevice, IMMDeviceCollection, IMMDeviceEnumerator, IMMNotificationClient,
        IMMNotificationClient_Impl, MMDeviceEnumerator, DEVICE_STATE, DEVICE_STATE_ACTIVE,
    },
    System::Com::{CoCreateInstance, CLSCTX_ALL},
};
use windows_core::{implement, Interface, GUID, PCWSTR, PWSTR};

use crate::com::com_initialized;

#[derive(Error, Debug)]
pub enum ProcessesError {
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
}

#[derive(Debug)]
pub struct Session {
    name: PWSTR,
    process_name: Option<String>,
    pid: u32,
    is_system: bool,
    session: IAudioSessionControl2,
    session1: IAudioSessionControl,
}

impl Session {
    pub fn get_name(&self) -> &PWSTR {
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

    pub fn from_session(session: IAudioSessionControl2) -> Result<Self, ProcessesError> {
        let pid = unsafe { session.GetProcessId() }.map_err(ProcessesError::ProcessIdError)?;
        let name_pwstr = unsafe { session.GetSessionInstanceIdentifier().map_err(ProcessesError::DisplayNameError)? };
        let process_name = Self::parse_process_name(name_pwstr);
        let is_system = unsafe { session.IsSystemSoundsSession() };
        let session1 = session.cast::<IAudioSessionControl>().map_err(ProcessesError::SessionCastError)?;
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

    pub fn get_display_name(&self) -> Result<String, ProcessesError> {
        let display_name = unsafe { self.session1.GetDisplayName() }.map_err(ProcessesError::DisplayNameError)?;
        Ok(unsafe { display_name.to_string() }.unwrap())
    }

    pub fn get_state(&self) -> Result<AudioSessionState, ProcessesError> {
        let state = unsafe { self.session1.GetState() }.map_err(ProcessesError::DisplayNameError)?;
        Ok(state)
    }
}
pub struct ProcessesManager {}

impl ProcessesManager {
    pub fn new() -> Result<Self, ProcessesError> {
        Ok(Self {})
    }

    /// Queries all active audio sessions
    pub fn query_sessions(&self) -> Result<Vec<Session>, ProcessesError> {
        com_initialized();
        let device_enumerator: IMMDeviceEnumerator =
            unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL) }.map_err(ProcessesError::InstanceCreationError)?;
        let dev_collection = Devices::new(&device_enumerator)?;

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
}

// Once again, taken from CPAL, thank you!
struct Devices {
    dev_collection: IMMDeviceCollection,
    dev_count: u32,
    next_index: u32,
}

impl Devices {
    pub fn new(enumerator: &IMMDeviceEnumerator) -> Result<Self, ProcessesError> {
        let dev_collection =
            unsafe { enumerator.EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE) }.map_err(ProcessesError::DeviceCollectionError)?;
        let dev_count = unsafe { dev_collection.GetCount() }.map_err(ProcessesError::DeviceCountError)?;
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

struct AudioSessions {
    session_enum: IAudioSessionEnumerator,
    session_count: i32,
    next_index: i32,
}

impl AudioSessions {
    pub fn new(device: IMMDevice) -> Result<Self, ProcessesError> {
        let mgr = unsafe { device.Activate::<IAudioSessionManager2>(CLSCTX_ALL, None) }.map_err(ProcessesError::DeviceActivationError)?;
        let session_enum = unsafe { mgr.GetSessionEnumerator() }.map_err(ProcessesError::SessionEnumeratorError)?;
        let session_count = unsafe { session_enum.GetCount() }.map_err(ProcessesError::SessionCountError)?;
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
    fn test_processes() {
        com_initialized();
        let p = ProcessesManager::new();
        assert!(p.is_ok());

        let mut p = p.unwrap();
        assert!(p.query_sessions().is_ok());
    }
}
