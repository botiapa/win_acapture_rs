use std::mem::ManuallyDrop;

use windows::Win32::{
    Media::Audio::{
        AUDIOCLIENT_ACTIVATION_PARAMS, AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK, PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE,
    },
    System::{
        Com::{
            CoTaskMemAlloc,
            StructuredStorage::{PropVariantClear, PROPVARIANT, PROPVARIANT_0, PROPVARIANT_0_0, PROPVARIANT_0_0_0},
            BLOB,
        },
        Variant::VT_BLOB,
    },
};

pub(crate) struct SafeActivationParams(PROPVARIANT);

impl SafeActivationParams {
    pub fn new(pid: u32) -> Self {
        let params_ptr = unsafe { CoTaskMemAlloc(size_of::<AUDIOCLIENT_ACTIVATION_PARAMS>()) } as *mut AUDIOCLIENT_ACTIVATION_PARAMS;
        debug_assert!(!params_ptr.is_null(), "Failed allocating memory for activation params");
        let audioclient_activate_params: &mut AUDIOCLIENT_ACTIVATION_PARAMS = unsafe { &mut *params_ptr };
        audioclient_activate_params.ActivationType = AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK;
        audioclient_activate_params.Anonymous.ProcessLoopbackParams.ProcessLoopbackMode = PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE;
        audioclient_activate_params.Anonymous.ProcessLoopbackParams.TargetProcessId = pid;

        let inner_prop = ManuallyDrop::new(PROPVARIANT_0_0 {
            vt: VT_BLOB,
            wReserved1: 0,
            wReserved2: 0,
            wReserved3: 0,
            Anonymous: PROPVARIANT_0_0_0 {
                blob: BLOB {
                    cbSize: size_of::<AUDIOCLIENT_ACTIVATION_PARAMS>() as u32,
                    pBlobData: audioclient_activate_params as *mut _ as *mut u8,
                },
            },
        });

        let activate_params = PROPVARIANT {
            Anonymous: PROPVARIANT_0 { Anonymous: inner_prop },
        };

        Self(activate_params)
    }

    pub fn prop(&self) -> &PROPVARIANT {
        &self.0
    }
}

impl Drop for SafeActivationParams {
    fn drop(&mut self) {
        unsafe {
            PropVariantClear(&mut self.0 as *mut _ as *mut PROPVARIANT).expect("Failed clearing activation params");
        }
    }
}
