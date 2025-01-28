use std::{collections::HashMap, string::FromUtf16Error};

use thiserror::Error;
use windows::Win32::Media::Audio::{AudioSessionDisconnectReason, AudioSessionState, IAudioSessionControl2};
use windows::Win32::{
    Foundation::{self, PROPERTYKEY},
    Media::Audio::{
        EDataFlow, ERole, IAudioSessionEvents, IAudioSessionEvents_Impl, IMMDeviceEnumerator, IMMNotificationClient,
        IMMNotificationClient_Impl, MMDeviceEnumerator, DEVICE_STATE,
    },
    System::Com::{CoCreateInstance, CLSCTX_ALL},
};
use windows_core::{implement, PCWSTR, PWSTR};

use crate::com::com_initialized;
use crate::processes::Session;

pub struct Notifications {
    _device_notification_client: Option<(IMMDeviceEnumerator, IMMNotificationClient)>,
    _session_notification_client: HashMap<String, (IAudioSessionControl2, IAudioSessionEvents)>,
}

#[derive(Error, Debug)]
pub enum NotificationError {
    #[error("Failed creating instance: {0}")]
    InstanceCreationError(windows::core::Error),
    #[error("Already registered for notifications")]
    NotificationAlreadyRegistered,
    #[error("Failed registering for notifications: {0}")]
    NotificationRegisterError(windows::core::Error),
    #[error("Failed converting raw PCWSTR string: {0}")]
    PCWSTRConversionError(FromUtf16Error),
    #[error("Failed activating device: {0}")]
    SessionManagerActivationError(windows::core::Error),
    #[error("Failed setting up notification through session manager: {0}")]
    FailedSettingUpNotification(windows::core::Error),
}

impl Notifications {
    pub fn new() -> Self {
        Self {
            _device_notification_client: None,
            _session_notification_client: HashMap::new(),
        }
    }

    pub fn setup_notifications(&mut self) -> Result<(), NotificationError> {
        if self._device_notification_client.is_some() {
            return Err(NotificationError::NotificationAlreadyRegistered);
        }
        com_initialized();
        let device_enumerator: IMMDeviceEnumerator =
            unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL) }.map_err(NotificationError::InstanceCreationError)?;
        let nclient: IMMNotificationClient = IDeviceNotificationClient::new().into();
        unsafe { device_enumerator.RegisterEndpointNotificationCallback(&nclient) }
            .map_err(NotificationError::NotificationRegisterError)?;
        self._device_notification_client = Some((device_enumerator, nclient));
        Ok(())
    }

    pub fn register_session_notification<CB>(&mut self, session: &Session, callback_fn: CB) -> Result<(), NotificationError>
    where
        CB: Fn(AudioSessionEventArgs) + Send + 'static,
    {
        let name = unsafe { session.get_name().to_string() }.map_err(NotificationError::PCWSTRConversionError)?;
        if self._session_notification_client.contains_key(&name) {
            return Err(NotificationError::NotificationAlreadyRegistered);
        }
        com_initialized();
        let session_notification_client = ISessionNotificationClient::new(session.get_name().clone(), callback_fn);
        let session_notification_client = session_notification_client.into();

        // Set up the notification
        unsafe { session.get_session().RegisterAudioSessionNotification(&session_notification_client) }
            .map_err(NotificationError::FailedSettingUpNotification)?;

        self._session_notification_client
            .insert(name, (session.get_session().clone(), session_notification_client));
        Ok(())
    }
}

impl Drop for Notifications {
    fn drop(&mut self) {
        if let Some((enumerator, nclient)) = self._device_notification_client.take() {
            unsafe {
                enumerator
                    .UnregisterEndpointNotificationCallback(&nclient)
                    .expect("Failed unregistering notification client");
            };
        }

        for (_, (sc, nc)) in self._session_notification_client.drain() {
            unsafe {
                sc.UnregisterAudioSessionNotification(&nc)
                    .expect("Failed unregistering session notification client");
            };
        }
    }
}

#[implement(IMMNotificationClient)]
struct IDeviceNotificationClient {}

impl IDeviceNotificationClient {
    pub fn new() -> Self {
        Self {}
    }
}

impl IMMNotificationClient_Impl for IDeviceNotificationClient_Impl {
    fn OnDefaultDeviceChanged(&self, flow: EDataFlow, role: ERole, pwstrDefaultDevice: &PCWSTR) -> windows::core::Result<()> {
        println!("Default device changed");
        todo!()
    }

    fn OnDeviceAdded(&self, pwstrDeviceId: &PCWSTR) -> windows::core::Result<()> {
        println!("Device added");
        todo!()
    }

    fn OnDeviceRemoved(&self, pwstrDeviceId: &PCWSTR) -> windows::core::Result<()> {
        println!("Device removed");
        todo!()
    }

    fn OnDeviceStateChanged(&self, pwstrDeviceId: &PCWSTR, dwNewState: DEVICE_STATE) -> windows::core::Result<()> {
        println!("Device state changed");
        todo!()
    }

    fn OnPropertyValueChanged(&self, pwstrDeviceId: &PCWSTR, key: &PROPERTYKEY) -> windows::core::Result<()> {
        println!("Property value changed");
        todo!()
    }
}

#[implement(IAudioSessionEvents)]
struct ISessionNotificationClient<CB>
where
    CB: Fn(AudioSessionEventArgs) + Send + 'static,
{
    _session_id: PWSTR,
    callback_fn: CB,
}

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
    newdisplayname: PCWSTR,
    eventcontext: *const windows_core::GUID,
}

#[derive(Debug)]
pub struct SimpleVolumeChangedArgs {
    newvolume: f32,
    newmute: Foundation::BOOL,
    eventcontext: *const windows_core::GUID,
}

#[derive(Debug)]
pub struct ChannelVolumeChangedArgs {
    channelcount: u32,
    newchannelvolumearray: *const f32,
    changedchannel: u32,
    eventcontext: *const windows_core::GUID,
}

#[derive(Debug)]
pub struct GroupingParamChangedArgs {
    newgroupingparam: *const windows_core::GUID,
    eventcontext: *const windows_core::GUID,
}

#[derive(Debug)]
pub struct StateChangedArgs {
    newstate: AudioSessionState,
}

#[derive(Debug)]
pub struct SessionDisconnectedArgs {
    disconnectreason: AudioSessionDisconnectReason,
}

#[derive(Debug)]
pub struct IconPathChangedArgs {
    newiconpath: PCWSTR,
    eventcontext: *const windows_core::GUID,
}

impl<CB> ISessionNotificationClient<CB>
where
    CB: Fn(AudioSessionEventArgs) + Send + 'static,
{
    pub fn new(session_id: PWSTR, callback_fn: CB) -> Self {
        Self {
            _session_id: session_id,
            callback_fn,
        }
    }
}

impl<CB> IAudioSessionEvents_Impl for ISessionNotificationClient_Impl<CB>
where
    CB: Fn(AudioSessionEventArgs) + Send + 'static,
{
    fn OnDisplayNameChanged(
        &self,
        newdisplayname: &windows_core::PCWSTR,
        eventcontext: *const windows_core::GUID,
    ) -> windows_core::Result<()> {
        (self.callback_fn)(AudioSessionEventArgs::DisplayNameChanged(DisplayNameChangedArgs {
            newdisplayname: newdisplayname.clone(),
            eventcontext,
        }));
        Ok(())
    }

    fn OnIconPathChanged(&self, newiconpath: &windows_core::PCWSTR, eventcontext: *const windows_core::GUID) -> windows_core::Result<()> {
        (self.callback_fn)(AudioSessionEventArgs::IconPathChanged(IconPathChangedArgs {
            newiconpath: newiconpath.clone(),
            eventcontext,
        }));
        Ok(())
    }

    fn OnSimpleVolumeChanged(
        &self,
        newvolume: f32,
        newmute: Foundation::BOOL,
        eventcontext: *const windows_core::GUID,
    ) -> windows_core::Result<()> {
        (self.callback_fn)(AudioSessionEventArgs::SimpleVolumeChanged(SimpleVolumeChangedArgs {
            newvolume,
            newmute,
            eventcontext,
        }));
        Ok(())
    }

    fn OnChannelVolumeChanged(
        &self,
        channelcount: u32,
        newchannelvolumearray: *const f32,
        changedchannel: u32,
        eventcontext: *const windows_core::GUID,
    ) -> windows_core::Result<()> {
        (self.callback_fn)(AudioSessionEventArgs::ChannelVolumeChanged(ChannelVolumeChangedArgs {
            channelcount,
            newchannelvolumearray,
            changedchannel,
            eventcontext,
        }));
        Ok(())
    }

    fn OnGroupingParamChanged(
        &self,
        newgroupingparam: *const windows_core::GUID,
        eventcontext: *const windows_core::GUID,
    ) -> windows_core::Result<()> {
        (self.callback_fn)(AudioSessionEventArgs::GroupingParamChanged(GroupingParamChangedArgs {
            newgroupingparam,
            eventcontext,
        }));
        Ok(())
    }

    fn OnStateChanged(&self, newstate: windows::Win32::Media::Audio::AudioSessionState) -> windows_core::Result<()> {
        (self.callback_fn)(AudioSessionEventArgs::StateChanged(StateChangedArgs { newstate }));
        Ok(())
    }

    fn OnSessionDisconnected(
        &self,
        disconnectreason: windows::Win32::Media::Audio::AudioSessionDisconnectReason,
    ) -> windows_core::Result<()> {
        (self.callback_fn)(AudioSessionEventArgs::SessionDisconnected(SessionDisconnectedArgs {
            disconnectreason,
        }));
        Ok(())
    }
}
