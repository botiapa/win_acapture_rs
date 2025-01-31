use std::io::stdin;

use win_acapture_rs::{
    manager::{DeviceManager, SessionManager},
    notifications::{AudioSessionEventArgs, Notifications},
    session_notification::SessionCreated,
};

/// Setup events for every session
fn main() {
    let mut notification_manager = Notifications::new();

    // Set up session events
    let sessions = SessionManager::get_sessions().unwrap();
    for session in sessions {
        notification_manager.register_session_event(&session, handle_event).unwrap();
    }

    // Set up session notification (NewSession) tied to devices
    let devices = DeviceManager::get_devices().unwrap();
    for dev in devices {
        notification_manager
            .register_session_notification(dev, handle_notification)
            .unwrap();
    }

    println!("Listening for events, press enter to exit");
    stdin().read_line(&mut String::new()).unwrap();
}

fn handle_event(event: AudioSessionEventArgs) {
    match event {
        AudioSessionEventArgs::DisplayNameChanged(display_name_changed_args) => {
            println!("Display name changed: {:?}", display_name_changed_args)
        }
        AudioSessionEventArgs::IconPathChanged(icon_path_changed_args) => println!("Icon path changed: {:?}", icon_path_changed_args),
        AudioSessionEventArgs::SimpleVolumeChanged(simple_volume_changed_args) => {
            println!("Simple volume changed: {:?}", simple_volume_changed_args)
        }
        AudioSessionEventArgs::ChannelVolumeChanged(channel_volume_changed_args) => {
            println!("Channel volume changed: {:?}", channel_volume_changed_args)
        }
        AudioSessionEventArgs::GroupingParamChanged(grouping_param_changed_args) => {
            println!("Grouping param changed: {:?}", grouping_param_changed_args)
        }
        AudioSessionEventArgs::StateChanged(state_changed_args) => println!("State changed: {:?}", state_changed_args.get_state()),
        AudioSessionEventArgs::SessionDisconnected(session_disconnected_args) => {
            println!("Session disconnected: {:?}", session_disconnected_args.get_reason())
        }
    }
}

fn handle_notification(event: SessionCreated) {
    println!("New session: {:?}", event);
}
