use std::{collections::HashMap, sync::mpsc};

use log::{debug, trace};
use windows::Win32::{
    Media::Audio::{
        IAudioSessionControl, IAudioSessionControl2, IAudioSessionManager2, IAudioSessionNotification, IAudioSessionNotification_Impl,
    },
    System::Com::{CLSCTX_ALL, COINIT_MULTITHREADED, CoInitializeEx},
};
use windows_core::{Interface, implement};

use crate::{
    manager::{Device, Session},
    notifications::NotificationError,
};

pub(crate) enum SessionNotificationMessage {
    Ready,
    #[allow(dead_code)]
    Error(NotificationError),
    NotificationRegistered,
    NotificationUnregistered,
    Stopped,
}

type SessionNotificationCallback = Box<dyn Fn(SessionCreated) + Send + 'static + Sync>;

pub(super) enum SessionNotificationCommand {
    RegisterNotification(SessionNotificationCallback, Device),
    UnregisterNotification(Device),
    Stop,
}

type NotificationsMap = HashMap<String, (IAudioSessionManager2, IAudioSessionNotification)>;

pub(crate) fn session_notification_thread(
    send: mpsc::Sender<SessionNotificationMessage>,
    recv: mpsc::Receiver<SessionNotificationCommand>,
) {
    unsafe { CoInitializeEx(None, COINIT_MULTITHREADED) }.unwrap();
    let mut notifications: NotificationsMap = HashMap::new();
    send.send(SessionNotificationMessage::Ready).expect("Failed sending ready message");
    loop {
        match thread_inner(&send, &recv, &mut notifications) {
            Ok(LoopResult::Continue) => {}
            Ok(LoopResult::Stop) => {
                send.send(SessionNotificationMessage::Stopped)
                    .expect("Failed sending stopped message");
                break;
            }
            Err(err) => {
                send.send(SessionNotificationMessage::Error(err))
                    .expect("Failed sending error message");
                break;
            }
        }
    }
}

enum LoopResult {
    Continue,
    Stop,
}

fn thread_inner(
    send: &mpsc::Sender<SessionNotificationMessage>,
    recv: &mpsc::Receiver<SessionNotificationCommand>,
    notifications: &mut NotificationsMap,
) -> Result<LoopResult, NotificationError> {
    match recv.recv() {
        Ok(SessionNotificationCommand::RegisterNotification(cb, dev)) => {
            let session_notification_client = IAudioSessionNotificationClient::new(cb);
            let session_notification_client: IAudioSessionNotification = session_notification_client.into();
            let dev = dev.inner;

            let session_manager = unsafe { dev.Activate::<IAudioSessionManager2>(CLSCTX_ALL, None) }
                .map_err(NotificationError::FailedActivatingSessionManager)?;
            let session_enumerator = unsafe {
                session_manager
                    .GetSessionEnumerator()
                    .map_err(NotificationError::FailedActivatingSessionManager)?
            };
            unsafe { session_manager.RegisterSessionNotification(&session_notification_client) }
                .map_err(NotificationError::FailedSettingUpNotification)?;
            let dev_id = unsafe {
                dev.GetId()
                    .map_err(NotificationError::FailedGettingDeviceId)?
                    .to_string()
                    .map_err(NotificationError::PCWSTRConversionError)?
            };
            notifications.insert(dev_id, (session_manager, session_notification_client));
            // Have to call GetCount() to start th enotifications (MS documentation)
            unsafe {
                session_enumerator
                    .GetCount()
                    .map_err(NotificationError::FailedActivatingSessionManager)?;
            }

            trace!("Notification registered, notifications: {}", notifications.len());
            send.send(SessionNotificationMessage::NotificationRegistered)
                .expect("Failed sending notification registered message");
        }
        Ok(SessionNotificationCommand::UnregisterNotification(dev)) => {
            let dev = dev.inner;
            let dev_id = unsafe {
                dev.GetId()
                    .map_err(NotificationError::FailedGettingDeviceId)?
                    .to_string()
                    .map_err(NotificationError::PCWSTRConversionError)?
            };
            if let Some((session_manager, notification_client)) = notifications.remove(&dev_id) {
                unsafe { session_manager.UnregisterSessionNotification(&notification_client) }
                    .map_err(|_| NotificationError::FailedUnregisteringSessionNotification)?;
                // TODO: Don't throw away inner error
                send.send(SessionNotificationMessage::NotificationUnregistered)
                    .expect("Failed sending notification unregistered message");
            }
            trace!("Notification unregistered, notifications: {}", notifications.len());
        }
        Ok(SessionNotificationCommand::Stop) => {
            // Unregister all notifications
            for (id, (session_manager, notification_client)) in notifications.drain() {
                unsafe { session_manager.UnregisterSessionNotification(&notification_client) }
                    .map_err(|_| NotificationError::FailedUnregisteringSessionNotification)?;
                debug!("Notification {} unregistered", id);
            }
            return Ok(LoopResult::Stop);
        }
        Err(err) => {
            panic!("Notification thread crashed, receiver error: {:?}", err);
        }
    }
    Ok(LoopResult::Continue)
}

#[derive(Debug)]
pub struct SessionCreated(String);

impl SessionCreated {
    pub fn get_name(&self) -> &String {
        &self.0
    }
}

#[implement(IAudioSessionNotification)]
struct IAudioSessionNotificationClient {
    callback_fn: SessionNotificationCallback,
}

impl IAudioSessionNotificationClient {
    pub fn new(callback_fn: SessionNotificationCallback) -> Self {
        Self { callback_fn }
    }
}

impl IAudioSessionNotification_Impl for IAudioSessionNotificationClient_Impl {
    fn OnSessionCreated(&self, newsession: windows_core::Ref<'_, IAudioSessionControl>) -> windows_core::Result<()> {
        let s = newsession.clone().expect("Failed cloning session");
        let new_session =
            Session::from_session(s.cast::<IAudioSessionControl2>().expect("Failed casting session")).expect("Failed creating session");
        (self.callback_fn)(SessionCreated(new_session.get_name().clone()));
        Ok(())
    }
}
