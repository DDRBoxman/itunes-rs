#![allow(non_snake_case)]

use std::{ffi::c_void, ptr::null};

use windows::{
    core::*,
    Win32::{Foundation::BOOL, System::Com::*},
};

fn main() -> Result<()> {
    println!("Hello, world!");

    unsafe {
        CoInitializeEx(None, COINIT_MULTITHREADED)?;
    }

    // Create ISpellCheckerFactory
    let app: Iitunes = unsafe { CoCreateInstance(&CitunesApp, None, CLSCTX_ALL)? };

    unsafe {
        // let result = app.Play();
        //let result: HRESULT = app.put_SoundVolume(50);
        let version = BSTR::default();
        let result: HRESULT = app.get_Version(&version);
        println!("{}", result);
        println!("{}", version);

        let mut pads: *mut c_void = std::ptr::null_mut();
        let track = &mut pads as *mut *mut c_void;

        let result: HRESULT = app.get_CurrentTrack(track);
        println!("{}", result);

        if result == HRESULT(1) { 
            // iTunes doesn't have a track right now
            // Internally "Incorrect function."
        } else {
            let name = BSTR::default();

            println!("{}", track.is_null());

            let fixed = track as *const IITTrack;

            let result = (*fixed).get_Name(&name);
            println!("{}", result);
            println!("{}", name);
        }
    }

    Ok(())
}

#[windows::core::interface("4CB0915D-1E54-4727-BAF3-CE6CC9A225A1")]
unsafe trait IITTrack: IUnknown {
    unsafe fn GetTypeInfoCount(&self, pctinfo: u32) -> HRESULT;
    unsafe fn GetTypeInfo(&self, iTInfo: u32, lcid: u32, pptinfo: *mut ITypeInfo) -> HRESULT;
    unsafe fn GetIDsOfNames(
        &self,
        rgsznames: *const PWSTR,
        cnames: u32,
        lcid: u32,
        rgdispid: *mut i32,
    ) -> HRESULT;
    unsafe fn Invoke(
        &self,
        dispIdMember: i32,
        riid: *const ::windows::core::GUID,
        lcid: u32,
        wflags: u16,
        pDispParams: *mut DISPPARAMS,
        pVarResult: *mut VARIANT,
        pExcepInfo: *mut EXCEPINFO,
        puArgErr: *mut u32,
    ) -> HRESULT;
    unsafe fn GetITObjectIDs(
        &self,
        sourceID: *mut i64,
        playlistID: *mut i64,
        trackID: *mut i64,
        databaseID: *mut i64,
    ) -> HRESULT;
    unsafe fn get_Name(&self, name: &BSTR) -> HRESULT;
}

#[windows::core::interface("53AE1704-491C-4289-94A0-958815675A3D")]
unsafe trait IITPlaylist: IUnknown {
    unsafe fn GetTypeInfoCount(&self, pctinfo: u32) -> HRESULT;
    unsafe fn GetTypeInfo(&self, iTInfo: u32, lcid: u32, pptinfo: *mut ITypeInfo) -> HRESULT;
    unsafe fn GetIDsOfNames(
        &self,
        rgsznames: *const PWSTR,
        cnames: u32,
        lcid: u32,
        rgdispid: *mut i32,
    ) -> HRESULT;
    unsafe fn Invoke(
        &self,
        dispIdMember: i32,
        riid: *const ::windows::core::GUID,
        lcid: u32,
        wflags: u16,
        pDispParams: *mut DISPPARAMS,
        pVarResult: *mut VARIANT,
        pExcepInfo: *mut EXCEPINFO,
        puArgErr: *mut u32,
    ) -> HRESULT;
    unsafe fn GetITObjectIDs(
        &self,
        sourceID: *mut i64,
        playlistID: *mut i64,
        trackID: *mut i64,
        databaseID: *mut i64,
    ) -> HRESULT;
    unsafe fn get_Name(&self, name: &BSTR) -> HRESULT;
}

#[windows::core::interface("9FAB0E27-70D7-4e3a-9965-B0C8B8869BB6")]
unsafe trait IITObject: IUnknown {
    unsafe fn GetTypeInfoCount(&self, pctinfo: u32) -> HRESULT;
    unsafe fn GetTypeInfo(&self, iTInfo: u32, lcid: u32, pptinfo: *mut ITypeInfo) -> HRESULT;
    unsafe fn GetIDsOfNames(
        &self,
        rgsznames: *const PWSTR,
        cnames: u32,
        lcid: u32,
        rgdispid: *mut i32,
    ) -> HRESULT;
    unsafe fn Invoke(
        &self,
        dispIdMember: i32,
        riid: *const ::windows::core::GUID,
        lcid: u32,
        wflags: u16,
        pDispParams: *mut DISPPARAMS,
        pVarResult: *mut VARIANT,
        pExcepInfo: *mut EXCEPINFO,
        puArgErr: *mut u32,
    ) -> HRESULT;
    unsafe fn GetITObjectIDs(
        &self,
        sourceID: *mut i64,
        playlistID: *mut i64,
        trackID: *mut i64,
        databaseID: *mut i64,
    ) -> HRESULT;
}

#[windows::core::interface("206479C9-FE32-4f9b-A18A-475AC939B479")]
unsafe trait IITOperationStatus: IUnknown {
    unsafe fn GetTypeInfoCount(&self, pctinfo: u32) -> HRESULT;
    unsafe fn GetTypeInfo(&self, iTInfo: u32, lcid: u32, pptinfo: *mut ITypeInfo) -> HRESULT;
    unsafe fn GetIDsOfNames(
        &self,
        rgsznames: *const PWSTR,
        cnames: u32,
        lcid: u32,
        rgdispid: *mut i32,
    ) -> HRESULT;
    unsafe fn Invoke(
        &self,
        dispIdMember: i32,
        riid: *const ::windows::core::GUID,
        lcid: u32,
        wflags: u16,
        pDispParams: *mut DISPPARAMS,
        pVarResult: *mut VARIANT,
        pExcepInfo: *mut EXCEPINFO,
        puArgErr: *mut u32,
    ) -> HRESULT;
    unsafe fn get_InProgress(&self, isInProgress: BOOL) -> HRESULT;
    //unsafe fn get_Tracks (&self, iTrackCollection: *IITTrackCollection) -> HRESULT;
}

#[windows::core::interface("9DD6680B-3EDC-40db-A771-E6FE4832E34A")]
unsafe trait Iitunes: IUnknown {
    unsafe fn GetTypeInfoCount(&self, pctinfo: u32) -> HRESULT;
    unsafe fn GetTypeInfo(&self, iTInfo: u32, lcid: u32, pptinfo: *mut ITypeInfo) -> HRESULT;
    unsafe fn GetIDsOfNames(
        &self,
        rgsznames: *const PWSTR,
        cnames: u32,
        lcid: u32,
        rgdispid: *mut i32,
    ) -> HRESULT;
    unsafe fn Invoke(
        &self,
        dispIdMember: i32,
        riid: *const ::windows::core::GUID,
        lcid: u32,
        wflags: u16,
        pDispParams: *mut DISPPARAMS,
        pVarResult: *mut VARIANT,
        pExcepInfo: *mut EXCEPINFO,
        puArgErr: *mut u32,
    ) -> HRESULT;
    unsafe fn BackTrack(&self) -> windows::core::HRESULT;
    unsafe fn FastForward(&self) -> windows::core::HRESULT;
    unsafe fn NextTrack(&self) -> windows::core::HRESULT;
    unsafe fn Pause(&self) -> windows::core::HRESULT;
    unsafe fn Play(&self) -> windows::core::HRESULT;
    unsafe fn PlayFile(&self, filePath: BSTR) -> HRESULT;
    unsafe fn PlayPause(&self) -> HRESULT;
    unsafe fn PreviousTrack(&self) -> HRESULT;
    unsafe fn Resume(&self) -> HRESULT;
    unsafe fn Rewind(&self) -> HRESULT;
    unsafe fn Stop(&self) -> HRESULT;
    unsafe fn ConvertFile(&self, filePath: BSTR, iStatus: *mut IITOperationStatus) -> HRESULT;
    unsafe fn ConvertFiles(
        &self,
        filePaths: *const c_void,
        iStatus: *mut IITOperationStatus,
    ) -> HRESULT;
    unsafe fn ConvertTrack(
        &self,
        track: *const c_void,
        iStatus: *mut IITOperationStatus,
    ) -> HRESULT;
    unsafe fn ConvertTracks(
        &self,
        tracks: *const c_void,
        iStatus: *mut IITOperationStatus,
    ) -> HRESULT;
    unsafe fn CheckVersion(
        &self,
        majorVersion: i64,
        minorVersion: i64,
        isCompatible: *mut BOOL,
    ) -> HRESULT;
    unsafe fn GetITObjectByID(
        &self,
        sourceID: i64,
        playlistID: i64,
        trackID: i64,
        databaseID: i64,
        iObject: *mut IITObject,
    ) -> HRESULT;
    unsafe fn CreatePlaylist(&self, playlistName: BSTR, iPlaylist: *mut IITPlaylist) -> HRESULT;
    unsafe fn OpenURL(&self, url: BSTR) -> HRESULT;
    unsafe fn GotoMusicStoreHomePage(&self) -> HRESULT;
    unsafe fn UpdateIPod(&self) -> HRESULT;
    unsafe fn Authorize(&self) -> HRESULT;
    unsafe fn Quit(&self) -> HRESULT;

    unsafe fn get_Sources(&self); // todo: fix
    unsafe fn get_Encoders(&self); // todo: fix
    unsafe fn get_EQPresets(&self); // todo: fix
    unsafe fn get_Visuals(&self); // todo: fix
    unsafe fn get_Windows(&self); // todo: fix
    unsafe fn get_SoundVolume(&self); // todo: fix

    unsafe fn put_SoundVolume(&self, volume: i64) -> HRESULT;
    unsafe fn get_Mute(&self, isMuted: *mut bool) -> HRESULT;
    unsafe fn put_Mute(&self, shouldMute: bool) -> HRESULT;
    unsafe fn get_PlayerState(&self); // todo: fix
    unsafe fn get_PlayerPosition(&self, playerPos: &i64) -> HRESULT;
    unsafe fn put_PlayerPosition(&self, playerPos: i64) -> HRESULT;
    unsafe fn get_CurrentEncoder(&self); // todo: fix
    unsafe fn put_CurrentEncoder(&self); // todo: fix
    unsafe fn get_VisualsEnabled(&self, isEnabled: &bool) -> HRESULT;
    unsafe fn put_VisualsEnabled(&self, shouldEnable: bool) -> HRESULT;
    unsafe fn get_FullScreenVisuals(&self, isFullScreen: *mut bool) -> HRESULT;
    unsafe fn put_FullScreenVisuals(&self, shouldUseFullScreen: bool) -> HRESULT;
    unsafe fn get_VisualSize(&self); // todo: fix
    unsafe fn put_VisualSize(&self); // todo: fix
    unsafe fn get_CurrentVisual(&self); //todo: fix
    unsafe fn put_CurrentVisual(&self); // todo: fix
    unsafe fn get_EQEnabled(&self, isEnabled: *mut bool) -> HRESULT;
    unsafe fn put_EQEnabled(&self, shouldEnable: bool) -> HRESULT;
    unsafe fn get_CurrentEQPreset(&self); // todo: fix
    unsafe fn put_CurrentEQPreset(&self); // todo: fix
    unsafe fn get_CurrentStreamTitle(&self, streamTitle: &BSTR) -> HRESULT;
    unsafe fn get_CurrentStreamURL(&self, streamTitle: &BSTR) -> HRESULT;
    unsafe fn get_BrowserWindow(&self); // todo: fix
    unsafe fn get_EQWindow(&self); //todo: fix
    unsafe fn get_LibrarySource(&self); // todo: fix
    unsafe fn get_LibraryPlaylist(&self); // todo: fix
    unsafe fn get_CurrentTrack(&self, iTrack: *mut *mut ::core::ffi::c_void) -> HRESULT;
    unsafe fn get_CurrentPlaylist(&self); //todo: fix
    unsafe fn get_SelectedTracks(&self); // todo: fix
    unsafe fn get_Version(&self, version: &BSTR) -> HRESULT;
    unsafe fn SetOptions(&self, options: i64);
}

const CitunesApp: GUID = GUID::from_u128(0xDC0C2640_1415_4644_875C_6F4D769839BA);

#[windows::core::interface("DC0C2640-1415-4644-875C-6F4D769839BA")]
unsafe trait ItunesApp: IUnknown {}
