use std::sync::mpsc::{self};
use std::thread::{self, JoinHandle};
use std::{collections::HashMap, string::FromUtf16Error};

use log::trace;
use thiserror::Error;
use windows::Win32::Media::Audio::{AudioSessionDisconnectReason, AudioSessionState, IAudioSessionControl2};
use windows::Win32::{
    Foundation::{self, PROPERTYKEY},
    Media::Audio::{
        DEVICE_STATE, EDataFlow, ERole, IAudioSessionEvents, IAudioSessionEvents_Impl, IMMDeviceEnumerator, IMMNotificationClient,
        IMMNotificationClient_Impl, MMDeviceEnumerator,
    },
    System::Com::{CLSCTX_ALL, CoCreateInstance},
};
use windows_core::{PCWSTR, PWSTR, implement};

use crate::com::com_initialized;
use crate::event_args::{
    AudioSessionEventArgs, ChannelVolumeChangedArgs, DefaultDeviceChangedEventArgs, DeviceAddedEventArgs, DeviceNotificationEventArgs,
    DevicePropertyValueChangedEventArgs, DeviceRemovedEventArgs, DeviceStateChangedEventArgs, DisplayNameChangedArgs,
    GroupingParamChangedArgs, IconPathChangedArgs, SessionDisconnectedArgs, SimpleVolumeChangedArgs, StateChangedArgs,
};
use crate::manager::{AudioError, Device, Session};
use crate::session_notification::{SessionCreated, SessionNotificationCommand, SessionNotificationMessage, session_notification_thread};

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

pub struct Notifications {
    _device_notification_client: Option<(IMMDeviceEnumerator, IMMNotificationClient)>,
    _session_event_client: HashMap<String, (IAudioSessionControl2, IAudioSessionEvents)>,
    _session_notification: Option<(
        mpsc::Sender<SessionNotificationCommand>,
        mpsc::Receiver<SessionNotificationMessage>,
        JoinHandle<()>,
    )>,
}

impl Notifications {
    pub fn new() -> Self {
        Self {
            _device_notification_client: None,
            _session_event_client: HashMap::new(),
            _session_notification: None,
        }
    }
    pub fn register_session_event<CB>(&mut self, session: &Session, callback_fn: CB) -> Result<(), NotificationError>
    where
        CB: Fn(AudioSessionEventArgs) + Send + 'static,
    {
        if self._session_event_client.contains_key(session.get_name()) {
            return Err(NotificationError::NotificationAlreadyRegistered);
        }
        com_initialized();
        let session_notification_client = ISessionEventClient::new(session.get_name().clone(), callback_fn);
        let session_notification_client = session_notification_client.into();

        // Set up the notification
        unsafe { session.get_session().RegisterAudioSessionNotification(&session_notification_client) }
            .map_err(NotificationError::FailedSettingUpNotification)?;

        self._session_event_client.insert(
            session.get_name().clone(),
            (session.get_session().clone(), session_notification_client),
        );
        trace!("Session event registered: {}", session.get_name());
        Ok(())
    }

    pub fn unregister_session_event(&mut self, name: &str) -> Result<(), NotificationError> {
        if let Some((sc, nc)) = self._session_event_client.remove(name) {
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

    pub fn register_device_notification<CB>(&mut self, callback_fn: CB) -> Result<(), NotificationError>
    where
        CB: Fn(DeviceNotificationEventArgs) + Send + 'static,
    {
        if self._device_notification_client.is_some() {
            return Err(NotificationError::NotificationAlreadyRegistered);
        }
        com_initialized();
        let device_enumerator: IMMDeviceEnumerator =
            unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL) }.map_err(NotificationError::InstanceCreationError)?;
        let nclient: IMMNotificationClient = IDeviceNotificationClient::new(callback_fn).into();

        unsafe { device_enumerator.RegisterEndpointNotificationCallback(&nclient) }
            .map_err(NotificationError::NotificationRegisterError)?;
        self._device_notification_client = Some((device_enumerator, nclient));
        Ok(())
    }

    pub fn unregister_device_notification(&mut self) -> Result<(), NotificationError> {
        if let Some((enumerator, nclient)) = self._device_notification_client.take() {
            unsafe { enumerator.UnregisterEndpointNotificationCallback(&nclient) }
                .map_err(NotificationError::NotificationUnregisterError)?;
        }
        Ok(())
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
            trace!("Device notification unregistered");
        }

        for (_, (sc, nc)) in self._session_event_client.drain() {
            unsafe {
                sc.UnregisterAudioSessionNotification(&nc)
                    .expect("Failed unregistering session notification client");
            };
            trace!("Session event unregistered");
        }

        if let Some((send, recv, t)) = self._session_notification.take() {
            send.send(SessionNotificationCommand::Stop).unwrap();
            t.join().unwrap();
            trace!("Session notification thread stopped");
        }
    }
}

#[implement(IMMNotificationClient)]
struct IDeviceNotificationClient<CB>
where
    CB: Fn(DeviceNotificationEventArgs) + Send + 'static,
{
    callback_fn: CB,
}

impl<CB> IDeviceNotificationClient<CB>
where
    CB: Fn(DeviceNotificationEventArgs) + Send + 'static,
{
    pub fn new(callback_fn: CB) -> Self {
        Self { callback_fn }
    }
}

impl<CB> IMMNotificationClient_Impl for IDeviceNotificationClient_Impl<CB>
where
    CB: Fn(DeviceNotificationEventArgs) + Send + 'static,
{
    fn OnDefaultDeviceChanged(&self, flow: EDataFlow, role: ERole, pwstrDefaultDevice: &PCWSTR) -> windows::core::Result<()> {
        (self.callback_fn)(DeviceNotificationEventArgs::DefaultDeviceChanged(DefaultDeviceChangedEventArgs {
            flow,
            role,
            defaultdevice: pwstrDefaultDevice.clone(),
        }));
        Ok(())
    }

    fn OnDeviceAdded(&self, pwstrDeviceId: &PCWSTR) -> windows::core::Result<()> {
        (self.callback_fn)(DeviceNotificationEventArgs::DeviceAdded(DeviceAddedEventArgs {
            pwstrDeviceId: pwstrDeviceId.clone(),
        }));
        Ok(())
    }

    fn OnDeviceRemoved(&self, pwstrDeviceId: &PCWSTR) -> windows::core::Result<()> {
        (self.callback_fn)(DeviceNotificationEventArgs::DeviceRemoved(DeviceRemovedEventArgs {
            pwstrDeviceId: pwstrDeviceId.clone(),
        }));
        Ok(())
    }

    fn OnDeviceStateChanged(&self, pwstrDeviceId: &PCWSTR, dwNewState: DEVICE_STATE) -> windows::core::Result<()> {
        (self.callback_fn)(DeviceNotificationEventArgs::DeviceStateChanged(DeviceStateChangedEventArgs {
            pwstrDeviceId: pwstrDeviceId.clone(),
            dwNewState,
        }));
        Ok(())
    }

    fn OnPropertyValueChanged(&self, pwstrDeviceId: &PCWSTR, key: &PROPERTYKEY) -> windows::core::Result<()> {
        (self.callback_fn)(DeviceNotificationEventArgs::DevicePropertyValueChanged(
            DevicePropertyValueChangedEventArgs {
                pwstrDeviceId: pwstrDeviceId.clone(),
                key: key.clone(),
            },
        ));
        Ok(())
    }
}

#[implement(IAudioSessionEvents)]
struct ISessionEventClient<CB>
where
    CB: Fn(AudioSessionEventArgs) + Send + 'static,
{
    _session_id: String,
    _callback_fn: CB,
}

impl<CB> ISessionEventClient<CB>
where
    CB: Fn(AudioSessionEventArgs) + Send + 'static,
{
    pub fn new(session_id: String, callback_fn: CB) -> Self {
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
