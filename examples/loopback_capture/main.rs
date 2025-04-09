use std::{
    error::Error,
    sync::{Arc, Mutex},
    thread,
    time::Duration,
};

use win_acapture_rs::audio_client::AudioClient;

fn main() -> Result<(), Box<dyn Error>> {
    let audio_capture = AudioClient::new();

    // Use a proper circular buffer for this, this is just a simple example
    let samples = Arc::new(Mutex::new(vec![0; 1024]));
    let samples_clone = samples.clone();

    // if AudioStream is dropped, the stream will be stopped, so we need to keep it alive
    let stream = audio_capture.start_recording_default_loopback(
        move |data| {
            let mut samples = samples_clone.lock().unwrap();
            samples.extend_from_slice(data.data());
        },
        move |error| {
            println!("error: {:?}", error);
        },
    )?;

    thread::sleep(Duration::from_secs(1));
    println!("read samples: {:?}", samples.lock().unwrap().len());

    // No need to stop the stream, it will be stopped when the stream is dropped

    Ok(())
}
