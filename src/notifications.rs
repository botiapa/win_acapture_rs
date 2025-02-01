use std::sync::mpsc::{self};
use std::thread::{self, JoinHandle};
use std::{collections::HashMap, string::FromUtf16Error};

use log::trace;
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
use crate::manager::{AudioError, Device, SafeSessionId, Session};
use crate::session_notification::{session_notification_thread, SessionCreated, SessionNotificationCommand, SessionNotificationMessage};

pub struct Notifications {
    _device_notification_client: Option<(IMMDeviceEnumerator, IMMNotificationClient)>,
    _session_event_client: HashMap<String, (IAudioSessionControl2, IAudioSessionEvents)>,
    _session_notification: Option<(
        mpsc::Sender<SessionNotificationCommand>,
        mpsc::Receiver<SessionNotificationMessage>,
        JoinHandle<()>,
    )>,
}

#[derive(Error, Debug)]
pub enum NotificationError {
    #[error("Failed creating instance: {0}")]
    InstanceCreationError(windows::core::Error),
    #[error("Already registered for notifications")]
    NotificationAlreadyRegistered,
    #[error("Failed registering for notifications: {0}")]
    NotificationRegisterError(windows::core::Error),
    #[error("Failed unregistering for notifications: {0}")]
    NotificationUnregisterError(windows::core::Error),
    #[error("Failed converting raw PCWSTR string: {0}")]
    PCWSTRConversionError(FromUtf16Error),
    #[error("Failed activating device: {0}")]
    SessionManagerActivationError(windows::core::Error),
    #[error("Failed setting up notification through session manager: {0}")]
    FailedSettingUpNotification(windows::core::Error),
    #[error("Failed enumerating devices: {0}")]
    FailedEnumeratingDevices(AudioError),
    #[error("Failed activating session manager: {0}")]
    FailedActivatingSessionManager(windows::core::Error),
    #[error("Failed getting device id: {0}")]
    FailedGettingDeviceId(windows::core::Error),
    #[error("Failed starting notification thread")]
    FailedStartingNotificationThread,
    #[error("Failed setting up notification")]
    FailedRegisteringSessionNotification,
    #[error("Failed unregistering notification")]
    FailedUnregisteringSessionNotification,
    #[error("Notification thread not running, can't unregister notification")]
    SessionNotificationThreadNotRunning,
}

impl Notifications {
    pub fn new() -> Self {
        Self {
            _device_notification_client: None,
            _session_event_client: HashMap::new(),
            _session_notification: None,
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

    pub fn register_session_event<CB>(&mut self, session: &Session, callback_fn: CB) -> Result<(), NotificationError>
    where
        CB: Fn(AudioSessionEventArgs) + Send + 'static,
    {
        let name = unsafe { session.get_name().to_string() }.map_err(NotificationError::PCWSTRConversionError)?;
        if self._session_event_client.contains_key(&name) {
            return Err(NotificationError::NotificationAlreadyRegistered);
        }
        com_initialized();
        let session_notification_client = ISessionEventClient::new(session.get_name().clone(), callback_fn);
        let session_notification_client = session_notification_client.into();

        // Set up the notification
        unsafe { session.get_session().RegisterAudioSessionNotification(&session_notification_client) }
            .map_err(NotificationError::FailedSettingUpNotification)?;

        self._session_event_client
            .insert(name.clone(), (session.get_session().clone(), session_notification_client));
        trace!("Session event registered: {}", name);
        Ok(())
    }

    pub fn unregister_session_event(&mut self, session_id: &SafeSessionId) -> Result<(), NotificationError> {
        let name = unsafe { session_id.0.to_string() }.map_err(NotificationError::PCWSTRConversionError)?;
        if let Some((sc, nc)) = self._session_event_client.remove(&name) {
            unsafe { sc.UnregisterAudioSessionNotification(&nc) }.map_err(NotificationError::NotificationUnregisterError)?;
        }
        trace!("Session event unregistered: {}", name);
        Ok(())
    }

    pub fn register_session_notification(
        &mut self,
        dev: Device,
        callback_fn: impl Fn(SessionCreated) + Send + 'static + Clone + Sync,
    ) -> Result<(), NotificationError> {
        self.notification_thread_running()
            .map_err(|_| NotificationError::FailedStartingNotificationThread)?;
        let (send, recv, _) = self._session_notification.as_ref().unwrap();
        send.send(SessionNotificationCommand::RegisterNotification(Box::new(callback_fn), dev))
            .unwrap();
        match recv.recv() {
            Ok(SessionNotificationMessage::NotificationRegistered) => Ok(()),
            _ => Err(NotificationError::FailedRegisteringSessionNotification),
        }
    }

    pub fn unregister_session_notification(&mut self, dev: Device) -> Result<(), NotificationError> {
        match &self._session_notification {
            Some((send, recv, _)) => {
                send.send(SessionNotificationCommand::UnregisterNotification(dev)).unwrap();
                match recv.recv() {
                    Ok(SessionNotificationMessage::NotificationUnregistered) => Ok(()),
                    _ => Err(NotificationError::FailedUnregisteringSessionNotification),
                }
            }
            None => Err(NotificationError::SessionNotificationThreadNotRunning),
        }
    }

    fn notification_thread_running(&mut self) -> Result<(), NotificationError> {
        if self._session_notification.is_some() {
            return Ok(());
        }

        let (response_send, response_recv) = std::sync::mpsc::channel();
        let (comm_send, comm_recv) = std::sync::mpsc::channel();

        let t = thread::spawn(move || session_notification_thread(response_send, comm_recv));
        match response_recv.recv() {
            Ok(SessionNotificationMessage::Ready) => {}
            _ => return Err(NotificationError::FailedStartingNotificationThread),
        }
        self._session_notification = Some((comm_send, response_recv, t));
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

        for (_, (sc, nc)) in self._session_event_client.drain() {
            unsafe {
                sc.UnregisterAudioSessionNotification(&nc)
                    .expect("Failed unregistering session notification client");
            };
        }

        if let Some((send, recv, t)) = self._session_notification.take() {
            send.send(SessionNotificationCommand::Stop).unwrap();
            t.join().unwrap();
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

#[derive(Debug, Clone)]
pub struct StateChangedArgs {
    newstate: AudioSessionState,
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
    disconnectreason: AudioSessionDisconnectReason,
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
    newiconpath: PCWSTR,
    eventcontext: *const windows_core::GUID,
}

impl IconPathChangedArgs {
    pub fn get_icon_path(&self) -> Result<String, NotificationError> {
        unsafe { self.newiconpath.to_string() }.map_err(NotificationError::PCWSTRConversionError)
    }
}

#[implement(IAudioSessionEvents)]
struct ISessionEventClient<CB>
where
    CB: Fn(AudioSessionEventArgs) + Send + 'static,
{
    _session_id: PWSTR,
    _callback_fn: CB,
}

impl<CB> ISessionEventClient<CB>
where
    CB: Fn(AudioSessionEventArgs) + Send + 'static,
{
    pub fn new(session_id: PWSTR, callback_fn: CB) -> Self {
        Self {
            _session_id: session_id,
            _callback_fn: callback_fn,
        }
    }
}

impl<CB> IAudioSessionEvents_Impl for ISessionEventClient_Impl<CB>
where
    CB: Fn(AudioSessionEventArgs) + Send + 'static,
{
    fn OnDisplayNameChanged(
        &self,
        newdisplayname: &windows_core::PCWSTR,
        eventcontext: *const windows_core::GUID,
    ) -> windows_core::Result<()> {
        (self._callback_fn)(AudioSessionEventArgs::DisplayNameChanged(DisplayNameChangedArgs {
            newdisplayname: newdisplayname.clone(),
            eventcontext,
        }));
        Ok(())
    }

    fn OnIconPathChanged(&self, newiconpath: &windows_core::PCWSTR, eventcontext: *const windows_core::GUID) -> windows_core::Result<()> {
        (self._callback_fn)(AudioSessionEventArgs::IconPathChanged(IconPathChangedArgs {
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
        (self._callback_fn)(AudioSessionEventArgs::SimpleVolumeChanged(SimpleVolumeChangedArgs {
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
        (self._callback_fn)(AudioSessionEventArgs::ChannelVolumeChanged(ChannelVolumeChangedArgs {
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
        (self._callback_fn)(AudioSessionEventArgs::GroupingParamChanged(GroupingParamChangedArgs {
            newgroupingparam,
            eventcontext,
        }));
        Ok(())
    }

    fn OnStateChanged(&self, newstate: windows::Win32::Media::Audio::AudioSessionState) -> windows_core::Result<()> {
        (self._callback_fn)(AudioSessionEventArgs::StateChanged(StateChangedArgs { newstate }));
        Ok(())
    }

    fn OnSessionDisconnected(
        &self,
        disconnectreason: windows::Win32::Media::Audio::AudioSessionDisconnectReason,
    ) -> windows_core::Result<()> {
        (self._callback_fn)(AudioSessionEventArgs::SessionDisconnected(SessionDisconnectedArgs {
            disconnectreason,
        }));
        Ok(())
    }
}
