use windows::Win32::Media::{
    Audio::{WAVEFORMATEX, WAVEFORMATEXTENSIBLE, WAVE_FORMAT_PCM},
    KernelStreaming::{KSDATAFORMAT_SUBTYPE_PCM, WAVE_FORMAT_EXTENSIBLE},
    Multimedia::{KSDATAFORMAT_SUBTYPE_IEEE_FLOAT, WAVE_FORMAT_IEEE_FLOAT},
};
use windows_core::GUID;

#[derive(Debug, Clone, PartialEq)]
pub struct SampleFormat {
    format_tag: FormatTag,
    channels: u16,
    sample_rate: u32,
    bits_per_sample: u16,
}

impl SampleFormat {
    pub fn new(format_tag: FormatTag, channel: u16, n_samples_per_sec: u32, w_bits_per_sample: u16) -> Self {
        Self {
            format_tag,
            channels: channel,
            sample_rate: n_samples_per_sec,
            bits_per_sample: w_bits_per_sample,
        }
    }

    pub fn get_format_tag(&self) -> &FormatTag {
        &self.format_tag
    }

    pub fn get_channel(&self) -> u16 {
        self.channels
    }

    pub fn get_n_samples_per_sec(&self) -> u32 {
        self.sample_rate
    }

    pub fn get_w_bits_per_sample(&self) -> u16 {
        self.bits_per_sample
    }

    pub fn block_align(&self) -> u16 {
        self.channels * self.bits_per_sample / 8
    }

    pub fn avg_bytes_per_sec(&self) -> u32 {
        self.sample_rate * self.block_align() as u32
    }

    pub const fn default() -> Self {
        Self {
            format_tag: FormatTag::WaveFormatIeeeFloat,
            channels: 2,
            sample_rate: 44100,
            bits_per_sample: 32,
        }
    }

    pub(crate) fn from_wave_format_ex(wave_format_ex: *const WAVEFORMATEX) -> Self {
        // thanks cpal
        fn cmp_guid(a: &GUID, b: &GUID) -> bool {
            (a.data1, a.data2, a.data3, a.data4) == (b.data1, b.data2, b.data3, b.data4)
        }
        let format_tag: FormatTag = unsafe { *wave_format_ex }.wFormatTag.into();
        let format_tag = match format_tag {
            FormatTag::WaveFormatExtensible => {
                if unsafe { *wave_format_ex }.cbSize < (size_of::<WAVEFORMATEXTENSIBLE>() - size_of::<WAVEFORMATEX>()) as u16 {
                    panic!("Invalid WAVEFORMATEXTENSIBLE size");
                }
                let wave_format_extensible_ptr = wave_format_ex as *const WAVEFORMATEXTENSIBLE;
                let subformat = unsafe { *wave_format_extensible_ptr }.SubFormat;
                if cmp_guid(&subformat, &KSDATAFORMAT_SUBTYPE_PCM) {
                    FormatTag::WaveFormatPcm
                } else if cmp_guid(&subformat, &KSDATAFORMAT_SUBTYPE_IEEE_FLOAT) {
                    FormatTag::WaveFormatIeeeFloat
                } else {
                    FormatTag::Unsupported
                }
            }
            _ => format_tag,
        };
        let wave_format_ex = unsafe { *wave_format_ex };
        Self {
            format_tag,
            channels: wave_format_ex.nChannels,
            sample_rate: wave_format_ex.nSamplesPerSec,
            bits_per_sample: wave_format_ex.wBitsPerSample,
        }
    }
}

impl From<SampleFormat> for WAVEFORMATEX {
    fn from(sample_format: SampleFormat) -> Self {
        let sample_size_bytes = sample_format.bits_per_sample / 8;
        let mut waveformatex = WAVEFORMATEX::default();
        waveformatex.wFormatTag = sample_format.format_tag.to_wave_format_tag();
        waveformatex.nChannels = sample_format.channels;
        waveformatex.nSamplesPerSec = sample_format.sample_rate;
        waveformatex.wBitsPerSample = sample_format.bits_per_sample;
        waveformatex.nBlockAlign = sample_format.channels * sample_size_bytes;
        waveformatex.nAvgBytesPerSec = sample_format.sample_rate * sample_format.channels as u32 * sample_size_bytes as u32;
        waveformatex
    }
}

impl Default for SampleFormat {
    fn default() -> Self {
        Self::default()
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum FormatTag {
    WaveFormatPcm,
    WaveFormatIeeeFloat,
    WaveFormatExtensible,
    Unsupported,
}

impl FormatTag {
    pub fn to_wave_format_tag(&self) -> u16 {
        match self {
            FormatTag::WaveFormatPcm => WAVE_FORMAT_PCM as u16,
            FormatTag::WaveFormatIeeeFloat => WAVE_FORMAT_IEEE_FLOAT as u16,
            FormatTag::WaveFormatExtensible => WAVE_FORMAT_EXTENSIBLE as u16,
            FormatTag::Unsupported => 0,
        }
    }
}

impl From<u16> for FormatTag {
    fn from(tag: u16) -> Self {
        match tag as u32 {
            WAVE_FORMAT_PCM => FormatTag::WaveFormatPcm,
            WAVE_FORMAT_IEEE_FLOAT => FormatTag::WaveFormatIeeeFloat,
            WAVE_FORMAT_EXTENSIBLE => FormatTag::WaveFormatExtensible,
            _ => FormatTag::Unsupported,
        }
    }
}
