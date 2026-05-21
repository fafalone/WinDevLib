## Windows Development Library for twinBASIC - API Coverage

This is a partial list of the current API coverage of WinDevLib.

### Broad overview of feature coverage 

- Shell interfaces and APIs including shell automation
- COM and OLE foundations including type library interfaces
- General purpose APIs from common system DLLs like user32, kernel32, gdi32, etc.
- Common Controls APIs and Common Dialog APIs; UI Ribbon; UI Animation
- DirectX technologies, from 8-12, including most extended features only from the DirectX SDK
- OpenGL, including full coverage beyond the Windows SDK through v4.6 with ARB, EXT, and vendor-specific functions.
- Additional multimedia APIs including VFW, winmm, CoreAudio, Media Foundation, Windows Imaging Component, XAudio2, and XACT3.
- Image Color Management / Windows Color System API
- Windows cryptography and certificate APIs, including both older APIs like crypt32.dll and CNG APIs bcrypt and ncrypt.
- Security-related APIs like LSA and Authz
- Accessibility, UI Automation, and IME
- Configuration Manager and SetupAPI; WIM; Software Licensing 
- Radio Manager, Sensor and Location APIs, Bluetooth APIs, Portable Devices API
- Windows Filtering Platform 
- WebView2
- Internet APIs including Winsock, WinInet, WinHttp, WinDNS, HTTP server, IPHLPAPI, traffic management, WLAN, and various system COM interfaces.
- Human Interface Device APIs (HID)
- Debug APIs (dbphlp, imagehlp)
- Background Intelligent Transfer Service (BITS)
- Virtual Disk Service (VDS)
- Low-level SQL APIs
- Windows Terminal Services
- Low-level WMI
- ActiveScript; Uniscribe
- Windows Biometric Framework
- Event Tracing for Windows (ETW)
- Group Policy
- Filter Manager
- Cloud filter APIs
- UPNP Automation

This is not a complete list and numerous minor features are covered or part of what I just listed as 'general purposes'.

### SDK Header Coverage

  **This is a work in progress and not complete!** Additional headers may be covered but not listed.

- Verified 100% basic coverage (for SDK 10.0.22621.0 minimum,  most for 10.0.26100.0);  
Excluded from completed %:
    - Definitions unsupported by the tB language with no reasonable substitute.  
    - Definitions disabled by conditional compilation with version flags for XP and earlier,  non-Windows platforms,  or kernel mode only.

  Basic Coverage: excludes macros,  callbacks->delegates,  ANSI APIs (though most are covered),  and other headers from #include statements.\
  Anything else missing is a bug and a report should be filed\

    winuser.h, UtilApiSet.h, processtopologyapi.h, msdelta.h, handleapi.h, cfgmgr32.h, ole2.h, avrt.h, KnownFolders.h, keycredmgr.h, mcx.h, windef.h, winver.h, dlgs.h, 
    realtimeapiset.h, msime.h, msimeapi.h, ws2bth.h, VersionHelpers.h, minwinbase.h, wsman.h, wcmapi.h, nb30.h, GPEdit.h, InputPanelConfiguration.h, commoncontrols.h, commoncontrols.idl, 
    WinEFS.h, winstring.h, qos2.h, traffic.h, qosobjs.h, qos.h, qossp.h, bluetoothleapis.h, bluetoothapis.h, bthsdpdef.h, fhcfg.h, fhsvcctl.h, fhstatus.h, fherrors.h, 
    fwpmu.h, ipsectypes.h, iketypes.h, fdi_fcitypes.h, fdi.h, fci.h, namespaceapi.h, physicalmonitorenumerationapi.h, highlevelmonitorconfigurationapi.h, combaseapi.h, 
    lowlevelmonitorconfigurationapi.h, wmistr.h, evntcons.h, wincodec.h, wincodec.idl, wincodecsdk.h, wincodecsdk.idl, WinML.h, sysinfoapi.h, cderr.h, commctrl.h, 
    ShellScalingApi.h, imagehlp.h, dbghelp.h, interlockedapi.h, upnp.h, upnp.idl, upnphost.h, upnphost.idl, RTWorkQ.h, wlanapi.h, magnification.h, threadpoolapiset.h, 
    cfapi.h, amsi.h, tokenbinding.h, WcnApi.h, WcnTypes.h, WcnDevice.h, WcnFunctionDiscoveryKeys.h, lmserver.h, cimfs.h, icmpapi.h, LMJoin.h, LMMsg.h, LMShare.h, 
    ObjSel.h, DSClient.h, security.h, minschannel.h, sspi.h, issper16.h, credssp.h, vbinterf.h, vdserr.h, vdscmprv.idl, vsprvcm.idl, vdshwprv.idl, vdscmmn.idl, 
    vdslun.idl, vdssp.idl, vdshp.idl, vdsvd.idl, vds.idl, vdshpcm.idl, vds.h, vdshwprv.h, vdslun.h, vdssys.h, directml.h, restartmanager.h, dde.h, ddeml.h, zmouse.h, 
    winbio_err.h, winbio_ioctl.h, winbio_types.h, winbio.h, winsvc.h, mssign.h, shellapi.h, bits.idl, bits1_5.idl, bits2_0.idl, bits2_5.idl, bits3_0.idl, bits4_0.idl, 
    bits5_0.h, bits10_1.h, bits10_2.h, bits10_3.h, bitscfg.h, qmgr.h, bits2_0.h, bits2_5.h, bits3_0.h, bits4_0.h, bits5_0.idl, bits10_1.idl, bits10_2.idl, bits10_3.idl, 
    bitscfg.h, qmgr.h, bitsmsg.h, dwrite.h, dwrite_1.h, dwrite_2.h, dwrite_3.h, shdeprecated.h, UserEnv.h, mscat.h, processenv.h, netioapi.h, iwscapi.h, wscapi.h, http.h, 
    WebEvnts.idl, WebEvnts.h, propkey.h, propkeydef.h, winsafer.h, powerbase.h, powersetting.h, powrprof.h, synchapi.h, dpa_dsa.h, DocumentTarget.idl, DocumentTarget.h, 
    propsys.h, SrRestorePtApi.h, compressapi.h, wincrypt.h, dpapi.h, mssip.h, memoryapi.h, wintrust.h, bcrypt.h, ncrypt.h, ncryptprotect.h, mobsync.h, ProcessSnapshot.h, 
    wincred.h, winhttp.h, websocket.h, photoacquire.h, oleacc.h, sddl.h, securitybaseapi.h, dssec.h, oleauto.h, olectl.h, newdev.h, processthreadsapi.h, virtdisk.h, 
    cryptuiapi.h, limits.h, evntrace.h, evntprov.h, relogger.h, relogger.idl, WinNls.h, WinNls32.h, ktmw32.h, fileapi.h, AccCtrl.h, AclAPI.h, dbt.h, TlHelp32.h, winnetwk.h, 
    enclaveapi.h, wincon.h, wincontypes.h, consoleapi.h, consoleapi2.h, consoleapi3.h, winreg.h, lsalookup.h, adtgen.h, authz.h, cfg.h, sfc.h, secext.h, AudioAPOTypes.h, 
    audioclient.h, audioclient.idl, audioclientactivationparams.h, audioendpoints.h, audioendpoints.idl, audioenginebaseapo.h, audioenginebaseapo.idl, 
    audioengineendpoint.h, audioengineendpoint.idl, audiomediatype.h, audiomediatype.idl, audiostatemonitorapi.h, audiopolicy.h, audiopolicy.idl, audiosessiontypes.h, 
    devicetopology.h, devicetopology.idl, endpointvolume.h, mmdeviceapi.h, spatialaudioclient.h, spatialaudiohrtf.h, spatialaudiometadata.h, l2cnm.h, wlantypes.h, 
    ndis\ObjectHeader.h, eaptypes.h, gdiplusflat.h, imapi.h, manipulations.h, manipulations.idl, propsys.h, propsys.idl, spellcheck.h, spellcheck.idl, spellcheckprovider.h, 
    spellcheckprovider.idl, structuredquerycondition.h, structuredquerycondition.idl, threadpoollegacyapiset.h, thumbcache.h, thumbcache.idl, thumbnailstreamcache.h, 
    thumbnailstreamcache.idl, timezoneapi.h, TOM.h, tdh.h, winres.h, UIRibbon.h, UIRibbon.idl, UIRibbonKeydef.h, UIRibbonPropertyHelpers.h, UIAutomationClient.h, 
    UIAutomationClient.idl, UIAutomationCore.h, UIAutomationCore.idl, UIAnimation.h, UIAnimation.idl, shtypes.h, shtypes.idl, TextServ.h, servprov.h, ServProv.Idl, 
    shappmgr.h, shappmgr.idl, shobjidl.h, ShObjIdl.idl, shlobj.h*, richedit.h, richole.h, shimgdata.h, uxtheme.h, mmsystem.h, mmsyscom.h, mciapi.h, mmiscapi2.h, playsoundapi.h, 
    mmeapi.h, timeapi.h, joystickapi.h, commdlg.h, cderr.h, prsht.h, prsht.idl, comctrl.h, ShlObj_core.h, ShObjidl_core.h, ShObjidl_core.idl, credentialprovider.h, 
    credentialprovider.idl, aclui.h, RadioMgr.h, RadioMgr.idl, PortableDevice.h, PortableDeviceAPI.h, PortableDeviceAPI.idl, portabledeviceclassextension.h, 
    portabledeviceclassextension.idl, portabledevicetypes.h, portabledevicetypes.idl, dsound.h, WinDNS.h, dstorage.h, dstorageerr.h, wininet.h, propapi.h, psapi.h, 
    propidl.h, propidl.idl, propidlbase.h, propidlbase.idl, propsys.idl, propsys.h, propvarutil.h, Xinput.h, winperf.h, perlib.h, spapidef.h, devpropdef.h, devpkey.h, 
    devguid.h, setupapi.h, prnasnot.h, winspool.h, libloaderapi.h, libloaderapi2.h, ioapiset.h, wingdi.h, coml2api.h, evr9.h, dxgi_1.h, dxgi_1.idl, dxgi_2.h, dxgi_2.idl, 
    dxgi_3.h, dxgi_3.idl, dxgi_4.h, dxgi_4.idl, dxgi_5.h, dxgi_5.idl, dxgi_6.h, dxgi_6.idl, DXGI_Messages.h, dxgitype.h, dxgitype.idl, dxgicommon.h, dxgicommon.idl, dxgidebug.h, 
    ntlsa.h, vsstyle.h, vssym32.h, usp10.h, xapo.h, xapofx.h, xaudio2.h, xaudio2fx.h, x3daudio.h, hrtfapoapi.h, WpdShellExtension.h, WpdMtpExtensions.h, evr.h, evr.idl, d3d11.h, d3d11.idl, 
    d3d11_2.h, d3d11_2.idl, d3d11_3.h, d3d11_3.idl, d3d11_4.h, d3d11_4.idl, d3d11on12.h, d3d11on12.idl, d2d1.h, d2d1_1.h, d2d1_2.h, d2d1_3.h, d2d1effectauthor.h, d2d1effects.h, 
    d2d1effects_1.h, d2d1effects_2.h, d2d1EffectauIEVRTrustedVideoPluginthor.h, d3dcommon.h, d3dcommon.idl, d3d10.h, d3d10.idl, d3d10misc.h, d3d10shader.h, d3d10effects.h, 
    d3d10sdklayers.h, d3d10sdklayers.idl, d3d10_1shader.h, d3d10_1.h, d3d10_1.idl, d3dcsx.h, presentation.idl, presentationtypes.h, presentationtypes.idl, wldp.h, webauthn.h, 
    ActivityCoordinator.h, ActivityCoordinatorTypes.h, ActivScp.h, ActivScp.idl, atacct.h, lm.h, lmcons.h, lmaccess.h, lmalert.h, lmapibuf.h, lmat.h, lmaudit.h, lmconfig.h, 
    lmerrlog.h, lmjoin.h, lmmsg.h, lmremutl.h, lmrepl.h, lmserver.h, lmshare.h, lmsname.h, lmstats.h, lmsvc.h, lmuse.h, lmuseflg.h, lmwksta.h, lmerr.h, lmdfs.h, poclass.h, 
    datetimeapi.h, ElsCore.h, ElsSrvc.h, Gb18030.h, stringsetapi.h, ime.h, imm.h, tcpestats.h, tcpmib.h, mprapidef.h, ipifcons.h, ifdef.h, nldef.h, ipmib.h, iprtrmib.h, 
    ipexport.h, iptypes.h, iphlpapi.h, winsmcrd.h, SCardErr.h, winscard.h, schannel.h, axcore.idl, devenum.idl, axextendedenums.h, mediaerr.h, dxva.h, dxva9typ.h, dxva2api.h, 
    dxva2api.idl, dxvahd.h, dxvahd.idl, icodecapi.h, wmcontainer.h, medparam.h, medparam.idl, mediaobj.h, mediaobj.idl, dmoreg.h, ksopmapi.h, opmapi.h, opmapi.idl, 
    d3dx9core.h, d3d9x.h, d3d9xshader.h, d3d9xtex.h, d3dx9xof.h, d3dx9mesh.h, d3dx9shape.h, d3dx11core.h, d3dx11tex.h, d3dx11async.h, d3d12compatibility.h, d3d12compatibility.idl, 
    d3d12shader.h, d3d12video.h, d3d12video.idl, d3d12.h, d3d12.idl, D3D12MarkerApiEnums.idl, d3d12compiler.h, d3d12compiler.idl, 
    fltUserStructures.h, fltUser.h, SpOrder.h, Filter.h, Filterr.h, NTQuery.h, apiquery2.h, appnotify.h, DeleteBrowsingHistory.h, DeleteBrowsingHistory.idl, 
    cpl.h, cplext.h, ddraw.h, ddstream.h, ddstream.idl, vmr9.h, vmr9.idl, vmrender.idl, amvideo.h, Dvp.h, uuids.h, amaudio.h, evcode.h, dyngraph.idl, dvdmedia.h, edevdefs.h, qnetwork.h, 
    xprtdefs.h, axextend.idl, amparse.h, vidcap.h, vidcap.idl, videoacc.h, videoacc.idl, dmodshow.h, dmodshow.idl, CameraUIControl.h, CameraUIControl.idl, austream.h, austream.idl, il21dec.h, 
    iwstdec.h, dvdif.h, strmif.h, strmif.idl, control.h, control.idl, amstream.h, amstream.idl, amva.h, sherrors.h, bcp47mrm.h, regbag.h, regbag.idl, wimgapi.h, lsalookupi.h, 
    bdatypes.h, bdaiface_enums.h, bdaiface.h, bdaiface.idl, mpeg2structs.h, Mpeg2Structs.idl, Mpeg2Bits.h, Mpeg2Data.h, Mpeg2Data.idl, Mpeg2PsiParser.idl, AtscPsipParser.h, 
    callobj.h, callobj.idl, WbemCli.h, WbemCli.idl, WMIUtils.h, WMIUtils.idl, dinput.h, icm.h, wcsplugin.h, wcsplugin.idl, jobapi.h, jobapi2.h, mixerocx.h, mixerocx.idl, 
    SensAPI.h, Sens.h, SensEvts.idl, OleDlg.h, dplay8.h, dpaddr.h, dplobby8.h, dpnathlp.h, dvoice.h, d3d8.h, d3d8types.h, d3d8caps.h, dxcapi.h, dcompanimation.h, dcomptypes.h, dcomp.h, 
    tlogstg.h, tlogstg.idl, PathCch.h, appmgmt.h, Dimm.h, Dimm.idl, Reconcil.h, objbase.h, objidlbase.h, objidlbase.idl, MSAAText.h, MSAAText.idl, TextStor.h, TextStor.idl, 
    SubAuth.h, davclient.h, DsGetDC.h, errhandlingapi.h, msports.h, objsafe.h, objsafe.idl, winternl.h, AppxPackaging.h, AppxPackaging.idl, XmlDom.idl, wofapi.h, ntddvdeo.h, 
    ntddvol.h, hidclass.h, hidusage.h, hidpi.h, hidsdi.h, ntenclv.h, winenclave.h, winenclaveapi.h, ioringapi.h, ntioring_x.h, urlmon.h, urlmon.idl, 
    gl.h, glu.h, DocObj.h, DocObj.idl, slpublic.h, slerror.h, sliddefs.h, fwpstypes.h, htiface.h, htiface.idl, htiframe.h, htiframe.idl, dmort.h, 
    odbcinst.h, sqltypes.h, sql.h, sqlext.h, sqlucode.h, minidumpapiset.h, coguid.h,   
    dls1.h, dls2.h, dmerror.h, dmdls.h, dmusbuff.h, dmusicc.h, dmusicf.h, dmplugin.h, dmusici.h, dmksctrl.h, dmusics.h, d3d12compiler.h/d3d12compiler.idl
    xact3.h, xact3wb.h, xma2defs.h, audiodefs.h, xact3d3.h, dxfile.h, 
    d3dx10.h, d3dx10core.h, d3dx10tex.h, d3dx10async.h, d3dx10mesh.h, d3d9on12.h, 
    
- Coverage in the 90%+ range\
winbase.h,  oleidl.h,  oaidl.h,  ocidl.h,  ocidl.idl,  , objidl.h, objidl.idl, presentation.h, ScrnSave.h

- Substantial coverage\
mmsciapi.h,  winnt.h, immdev.h, winioctl.h, mmreg.h, WS2spi.h, winerror.h, WindowsSearchErrors.h, windowsx.h, 

- Minimal coverage\
windot11.h, peninputpanel.h, xapobase.h, wmcodecdsp.h, ksmedia.h, d3dkmdt.h, d3dukmdt.h, d3dkmthk.h, fwpsu.h

  
