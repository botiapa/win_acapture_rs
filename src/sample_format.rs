use windows::Win32::Media::Audio::{WAVEFORMATEX, WAVE_FORMAT_PCM};

fn get_waveformatex(channel: u16, n_samples_per_sec: u32, w_bits_per_sample: u16) -> WAVEFORMATEX {
    let mut capture_format = WAVEFORMATEX::default();
    capture_format.wFormatTag = WAVE_FORMAT_PCM as u16;
    capture_format.nChannels = channel;
    capture_format.nSamplesPerSec = n_samples_per_sec;
    capture_format.wBitsPerSample = w_bits_per_sample;
    capture_format.nBlockAlign = capture_format.nChannels * capture_format.wBitsPerSample / 8;
    capture_format.nAvgBytesPerSec = capture_format.nSamplesPerSec * capture_format.nBlockAlign as u32;
    capture_format
}

impl From<SampleFormat> for WAVEFORMATEX {
    fn from(sample_format: SampleFormat) -> Self {
        get_waveformatex(
            sample_format.get_channel(),
            sample_format.get_n_samples_per_sec(),
            sample_format.get_w_bits_per_sample(),
        )
    }
}

#[derive(Debug, Clone)]
pub struct SampleFormat {
    channel: u16,
    n_samples_per_sec: u32,
    w_bits_per_sample: u16,
}

impl SampleFormat {
    pub fn new(channel: u16, n_samples_per_sec: u32, w_bits_per_sample: u16) -> Self {
        Self {
            channel,
            n_samples_per_sec,
            w_bits_per_sample,
        }
    }

    pub fn get_channel(&self) -> u16 {
        self.channel
    }

    pub fn get_n_samples_per_sec(&self) -> u32 {
        self.n_samples_per_sec
    }

    pub fn get_w_bits_per_sample(&self) -> u16 {
        self.w_bits_per_sample
    }
}

impl Default for SampleFormat {
    fn default() -> Self {
        Self {
            channel: 2,
            n_samples_per_sec: 44100,
            w_bits_per_sample: 16,
        }
    }
}
