use std::time::Duration;

/// Taken from the `cpal` library: `https://github.com/RustAudio/cpal`
/// Licensed under `Apache-2.0`
#[derive(Copy, Clone, Debug, Eq, Hash, PartialEq, PartialOrd, Ord)]
pub struct StreamInstant {
    secs: i64,
    nanos: u32,
}

impl StreamInstant {
    /// The amount of time elapsed from another instant to this one.
    ///
    /// Returns `None` if `earlier` is later than self.
    pub fn duration_since(&self, earlier: &Self) -> Option<Duration> {
        if self < earlier {
            None
        } else {
            (self.as_nanos() - earlier.as_nanos()).try_into().ok().map(Duration::from_nanos)
        }
    }

    /// Returns the instant in time after the given duration has passed.
    ///
    /// Returns `None` if the resulting instant would exceed the bounds of the underlying data
    /// structure.
    pub fn add(&self, duration: Duration) -> Option<Self> {
        self.as_nanos()
            .checked_add(duration.as_nanos() as i128)
            .and_then(Self::from_nanos_i128)
    }

    /// Returns the instant in time one `duration` ago.
    ///
    /// Returns `None` if the resulting instant would underflow. As a result, it is important to
    /// consider that on some platforms the [`StreamInstant`] may begin at `0` from the moment the
    /// source stream is created.
    pub fn sub(&self, duration: Duration) -> Option<Self> {
        self.as_nanos()
            .checked_sub(duration.as_nanos() as i128)
            .and_then(Self::from_nanos_i128)
    }

    fn as_nanos(&self) -> i128 {
        (self.secs as i128 * 1_000_000_000) + self.nanos as i128
    }

    #[allow(dead_code)]
    pub(crate) fn from_nanos(nanos: i64) -> Self {
        let secs = nanos / 1_000_000_000;
        let subsec_nanos = nanos - secs * 1_000_000_000;
        Self::new(secs, subsec_nanos as u32)
    }

    #[allow(dead_code)]
    pub(crate) fn from_nanos_i128(nanos: i128) -> Option<Self> {
        let secs = nanos / 1_000_000_000;
        if secs > i64::MAX as i128 || secs < i64::MIN as i128 {
            None
        } else {
            let subsec_nanos = nanos - secs * 1_000_000_000;
            debug_assert!(subsec_nanos < u32::MAX as i128);
            Some(Self::new(secs as i64, subsec_nanos as u32))
        }
    }

    #[allow(dead_code)]
    pub(crate) fn from_secs_f64(secs: f64) -> Self {
        let s = secs.floor() as i64;
        let ns = ((secs - s as f64) * 1_000_000_000.0) as u32;
        Self::new(s, ns)
    }

    pub fn new(secs: i64, nanos: u32) -> Self {
        Self { secs, nanos }
    }
}
