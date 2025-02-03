use windows::Win32::{
    Foundation::{self, PROPERTYKEY},
    Media::Audio::{AudioSessionDisconnectReason, AudioSessionState, EDataFlow, ERole, DEVICE_STATE},
};
use windows_core::PCWSTR;

use crate::notifications::NotificationError;

#[derive(Debug)]
pub enum AudioSessionEventArgs {
    DisplayNameChanged(DisplayNameChangedArgs),
    IconPathChanged(IconPathChangedArgs),
    SimpleVolumeChanged(SimpleVolumeChangedArgs),
    ChannelVolumeChanged(ChannelVolumeChangedArgs),
    GroupingParamChanged(GroupingParamChangedArgs),
    StateChanged(StateChangedArgs),
    SessionDisconnected(SessionDisconnectedArgs),
}

#[derive(Debug)]
pub struct DisplayNameChangedArgs {
    pub(crate) newdisplayname: PCWSTR,
    pub(crate) eventcontext: *const windows_core::GUID,
}

#[derive(Debug)]
pub struct SimpleVolumeChangedArgs {
    pub(crate) newvolume: f32,
    pub(crate) newmute: Foundation::BOOL,
    pub(crate) eventcontext: *const windows_core::GUID,
}

#[derive(Debug)]
pub struct ChannelVolumeChangedArgs {
    pub(crate) channelcount: u32,
    pub(crate) newchannelvolumearray: *const f32,
    pub(crate) changedchannel: u32,
    pub(crate) eventcontext: *const windows_core::GUID,
}

#[derive(Debug)]
pub struct GroupingParamChangedArgs {
    pub(crate) newgroupingparam: *const windows_core::GUID,
    pub(crate) eventcontext: *const windows_core::GUID,
}

#[derive(Debug, Clone)]
pub struct StateChangedArgs {
    pub(crate) newstate: AudioSessionState,
}

impl StateChangedArgs {
    pub fn get_state(&self) -> SessionState {
        match self.newstate.0 {
            0 => SessionState::AudioSessionStateInactive,
            1 => SessionState::AudioSessionStateActive,
            2 => SessionState::AudioSessionStateExpired,
            _ => panic!("Unknown session state"),
        }
    }
}

#[derive(Debug, Clone)]
pub enum SessionState {
    AudioSessionStateActive,
    AudioSessionStateExpired,
    AudioSessionStateInactive,
}

#[derive(Debug, Clone)]
pub struct SessionDisconnectedArgs {
    pub(crate) disconnectreason: AudioSessionDisconnectReason,
}

impl SessionDisconnectedArgs {
    pub fn get_reason(&self) -> SessionDisconnectReason {
        match self.disconnectreason.0 {
            0 => SessionDisconnectReason::DisconnectReasonDeviceRemoval,
            1 => SessionDisconnectReason::DisconnectReasonServerShutdown,
            2 => SessionDisconnectReason::DisconnectReasonFormatChanged,
            3 => SessionDisconnectReason::DisconnectReasonSessionLogoff,
            4 => SessionDisconnectReason::DisconnectReasonSessionDisconnected,
            5 => SessionDisconnectReason::DisconnectReasonExclusiveModeOverride,
            _ => panic!("Invalid session disconnect reason"),
        }
    }
}

#[derive(Debug, Clone)]
pub enum SessionDisconnectReason {
    DisconnectReasonDeviceRemoval,
    DisconnectReasonServerShutdown,
    DisconnectReasonFormatChanged,
    DisconnectReasonSessionLogoff,
    DisconnectReasonSessionDisconnected,
    DisconnectReasonExclusiveModeOverride,
}

#[derive(Debug)]
pub struct IconPathChangedArgs {
    pub(crate) newiconpath: PCWSTR,
    pub(crate) eventcontext: *const windows_core::GUID,
}

impl IconPathChangedArgs {
    pub fn get_icon_path(&self) -> Result<String, NotificationError> {
        unsafe { self.newiconpath.to_string() }.map_err(NotificationError::PCWSTRConversionError)
    }
}

//DeviceEventArgs
#[derive(Debug)]
pub enum DeviceNotificationEventArgs {
    DefaultDeviceChanged(DefaultDeviceChangedEventArgs),
    DeviceAdded(DeviceAddedEventArgs),
    DeviceRemoved(DeviceRemovedEventArgs),
    DeviceStateChanged(DeviceStateChangedEventArgs),
    DevicePropertyValueChanged(DevicePropertyValueChangedEventArgs),
}

#[derive(Debug)]
pub struct DefaultDeviceChangedEventArgs {
    pub(crate) flow: EDataFlow,
    pub(crate) role: ERole,
    pub(crate) defaultdevice: PCWSTR,
}

impl DefaultDeviceChangedEventArgs {
    pub fn get_default_device(&self) -> Result<String, NotificationError> {
        unsafe { self.defaultdevice.to_string() }.map_err(NotificationError::PCWSTRConversionError)
    }
}

#[derive(Debug)]
pub struct DeviceAddedEventArgs {
    pub(crate) pwstrDeviceId: PCWSTR,
}

impl DeviceAddedEventArgs {
    pub fn get_device_id(&self) -> Result<String, NotificationError> {
        unsafe { self.pwstrDeviceId.to_string() }.map_err(NotificationError::PCWSTRConversionError)
    }
}

#[derive(Debug)]
pub struct DeviceRemovedEventArgs {
    pub(crate) pwstrDeviceId: PCWSTR,
}

impl DeviceRemovedEventArgs {
    pub fn get_device_id(&self) -> Result<String, NotificationError> {
        unsafe { self.pwstrDeviceId.to_string() }.map_err(NotificationError::PCWSTRConversionError)
    }
}

#[derive(Debug)]
pub struct DeviceStateChangedEventArgs {
    pub(crate) pwstrDeviceId: PCWSTR,
    pub(crate) dwNewState: DEVICE_STATE,
}

impl DeviceStateChangedEventArgs {
    pub fn get_device_id(&self) -> Result<String, NotificationError> {
        unsafe { self.pwstrDeviceId.to_string() }.map_err(NotificationError::PCWSTRConversionError)
    }

    pub fn get_state(&self) -> DeviceState {
        self.dwNewState.into()
    }
}

#[derive(Debug)]
pub enum DeviceState {
    Active,
    Disabled,
    NotPresent,
    Unplugged,
    /// Mask: Includes audio endpoint devices in all states active, disabled, not present, and unplugged.
    All,
}

impl From<DEVICE_STATE> for DeviceState {
    fn from(state: DEVICE_STATE) -> Self {
        match state.0 {
            1u32 => DeviceState::Active,
            2u32 => DeviceState::Disabled,
            4u32 => DeviceState::NotPresent,
            8u32 => DeviceState::Unplugged,
            15u32 => DeviceState::All,
            _ => panic!("Invalid device state"),
        }
    }
}

pub const DEVICE_STATEMASK_ALL: u32 = 15u32;
pub const DEVICE_STATE_ACTIVE: DEVICE_STATE = DEVICE_STATE(1u32);
pub const DEVICE_STATE_DISABLED: DEVICE_STATE = DEVICE_STATE(2u32);
pub const DEVICE_STATE_NOTPRESENT: DEVICE_STATE = DEVICE_STATE(4u32);
pub const DEVICE_STATE_UNPLUGGED: DEVICE_STATE = DEVICE_STATE(8u32);

#[derive(Debug)]
pub struct DevicePropertyValueChangedEventArgs {
    pub(crate) pwstrDeviceId: PCWSTR,
    pub(crate) key: PROPERTYKEY,
}

impl DevicePropertyValueChangedEventArgs {
    pub fn get_device_id(&self) -> Result<String, NotificationError> {
        unsafe { self.pwstrDeviceId.to_string() }.map_err(NotificationError::PCWSTRConversionError)
    }

    pub fn get_property_key(&self) {
        unimplemented!()
    }
}
