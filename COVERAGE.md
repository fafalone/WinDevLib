## Windows Development Library for twinBASIC - API Coverage

This is a partial list of the current API coverage of WinDevLib.\
This is not a complete list and numerous minor features are covered or part of what is listed as 'general purpose'.

### Broad overview of feature coverage 

- Shell interfaces and APIs including shell automation and Property System
- COM and OLE foundations including type library interfaces
- General purpose APIs from common system DLLs like user32, kernel32, gdi32, ole32, oleaut32, etc.
- General purpose Native APIs from ntdll.
- Common Controls APIs and Common Dialog APIs; UI Ribbon; UI Animation
- DirectX technologies, from 8-12, including most extended features only from the DirectX SDK
- OpenGL, including full coverage beyond the Windows SDK through v4.6 with ARB, EXT, and vendor-specific functions.
- Windows Media Foundation
- CoreAudio/WASAPI, XAudio2, XACT3
- XInput
- Windows Imaging Component (WIC), GDIPlus
- WINMM.dll multimedia, timer, and joystick APIs
- Microsoft Speech API
- Image Color Management / Windows Color System API
- Windows cryptography and certificate APIs, including both older APIs like crypt32.dll and CNG APIs bcrypt and ncrypt.
- Security-related APIs like LSA and Authz
- Accessibility, UI Automation, and IME
- Configuration Manager and SetupAPI; WIM; Software Licensing; Package Manager (OPC)
- Radio Manager, Sensor and Location APIs, Bluetooth APIs, Portable Devices API
- Windows Media Device Manager 
- Windows Filtering Platform 
- WebView2
- Internet APIs including Winsock, WinInet, WinHttp, WinDNS, WNet, HTTP server, IPHLPAPI, websocket, urlmon, hlink, traffic management, WLAN, and various system COM interfaces.
- Local network APIs (netapi32)
- Active Directory 
- Human Interface Device APIs (HID)
- Winspool
- Debug APIs (dbghlp, imagehlp)
- Background Intelligent Transfer Service (BITS)
- Virtual Disk Service (VDS)
- Low-level SQL APIs
- Windows Terminal Services
- Low-level WMI
- ActiveScript; Uniscribe
- Windows Biometric Framework
- SmartCard APIs
- Event Tracing for Windows (ETW)
- Group Policy
- Filter Manager
- Cloud filter APIs
- UPNP Automation
- Windows Remote Management (WSMAN) low-level API


### SDK Header Coverage

  **This is a work in progress and not complete!** Additional headers may be covered but not listed.

  **Note:** Header file list is from the Windows 11 26000 SDK. Some header files are from other SDKs like the DirectX SDKs or other Microsoft SDKs or publications.

- Verified 100% basic coverage (for SDK 10.0.22621.0 minimum,  most for 10.0.26100.0);  
Excluded from completed %:
    - Definitions unsupported by the tB language with no reasonable substitute.  
    - Definitions disabled by conditional compilation with version flags for XP and earlier,  non-Windows platforms,  or kernel mode only.

  Basic Coverage: excludes macros,  callbacks->delegates,  ANSI APIs (though most are covered),  and other headers from #include statements.\
  Anything else missing is a bug and a report should be filed\

  AccCtrl.h, aclui.h, AclAPI.h, ActivScp.h, ActivScp.idl, ActivityCoordinator.h, ActivityCoordinatorTypes.h, adtgen.h, amaudio.h, amparse.h, amsi.h, amstream.h, amstream.idl, amva.h, amvideo.h, apiquery2.h, AppxPackaging.h, AppxPackaging.idl, appmgmt.h, appnotify.h, atacct.h, AtscPsipParser.h, AudioAPOTypes.h, audioclient.h, audioclient.idl, audioclientactivationparams.h, audiodefs.h, audioendpoints.h, audioendpoints.idl, audioenginebaseapo.h, audioenginebaseapo.idl, audioengineendpoint.h, audioengineendpoint.idl, audiomediatype.h, audiomediatype.idl, audiopolicy.h, audiopolicy.idl, audiosessiontypes.h, audiostatemonitorapi.h, austream.h, austream.idl, authz.h, avrt.h, axcore.idl, axextend.idl, axextendedenums.h, bcp47mrm.h, bcrypt.h, bdaiface.h, bdaiface.idl, bdaiface_enums.h, bdatypes.h, bits.idl, bits1_5.idl, bits10_1.h, bits10_1.idl, bits10_2.h, bits10_2.idl, bits10_3.h, bits10_3.idl, bits2_0.h, bits2_0.idl, bits2_5.h, bits2_5.idl, bits3_0.h, bits3_0.idl, bits4_0.h, bits4_0.idl, bits5_0.h, bits5_0.idl, bitscfg.h, bitsmsg.h, bluetoothapis.h, bluetoothleapis.h, bthdef.h, bthioctl.h, bthledef.h, bthsdpdef.h, CameraUIControl.h, CameraUIControl.idl, callobj.h, callobj.idl, cderr.h, cfg.h, cfgmgr32.h, cfapi.h, cimfs.h, combaseapi.h, codecapi.h, comctrl.h, commoncontrols.h, commoncontrols.idl, commctrl.h, commdlg.h, coml2api.h, compressapi.h, consoleapi.h, consoleapi2.h, consoleapi3.h, control.h, control.idl, coguid.h, cpl.h, cplext.h, credssp.h, credentialprovider.h, credentialprovider.idl, cryptuiapi.h, d2d1.h, d2d1_1.h, d2d1_2.h, d2d1_3.h, d2d1effectauthor.h, d2d1effectauthor_1.h, d2d1effects.h, d2d1effects_1.h, d2d1effects_2.h, d2dbasetypes.h, d2derr.h, d3d10.h, d3d10.idl, d3d10_1.h, d3d10_1.idl, d3d10_1shader.h, d3d10effects.h, d3d10misc.h, d3d10sdklayers.h, d3d10sdklayers.idl, d3d10shader.h, d3d11.h, d3d11.idl, d3d11_2.h, d3d11_2.idl, d3d11_3.h, d3d11_3.idl, d3d11_4.h, d3d11_4.idl, d3d11on12.h, d3d11on12.idl, d3d11sdklayers.h, d3d11sdklayers.idl, d3d11shader.h, d3d11shadertracing.h, d3d12.h, d3d12.idl, d3d12compatibility.h, d3d12compatibility.idl, d3d12compiler.h, d3d12compiler.idl, D3D12MarkerApiEnums.idl, d3d12shader.h, d3d12video.h, d3d12video.idl, d3d8.h, d3d8caps.h, d3d8types.h, d3d9on12.h, d3d9x.h, d3d9xshader.h, d3d9xtex.h, d3dcommon.h, d3dcommon.idl, d3dcsx.h, d3dx10.h, d3dx10async.h, d3dx10core.h, d3dx10mesh.h, d3dx10tex.h, d3dx11async.h, d3dx11core.h, d3dx11tex.h, d3dx9anim.h, d3dx9core.h, d3dx9mesh.h, d3dx9shape.h, d3dx9xof.h, davclient.h, dbt.h, dbghelp.h, dcomp.h, dcompanimation.h, dcomptypes.h, dde.h, ddeml.h, ddraw.h, ddstream.h, ddstream.idl, DeleteBrowsingHistory.h, DeleteBrowsingHistory.idl, devenum.idl, devguid.h, devicetopology.h, devicetopology.idl, devpkey.h, devpropdef.h, directml.h, Dimm.h, Dimm.idl, dinput.h, dlgs.h, dls1.h, dls2.h, dmdls.h, dmerror.h, dmksctrl.h, dmodshow.h, dmodshow.idl, dmort.h, dmoreg.h, dmusbuff.h, dmusicc.h, dmusicf.h, dmusici.h, dmplugin.h, dmusics.h, DocObj.h, DocObj.idl, DocumentTarget.h, DocumentTarget.idl, dpaddr.h, dpa_dsa.h, dplay8.h, dplobby8.h, dpapi.h, dpnathlp.h, dsclient.h, dsound.h, dssec.h, DsGetDC.h, dstorage.h, dstorageerr.h, dvcieetopology.h, dvdif.h, dvdmedia.h, dvoice.h, Dvp.h, dwrite.h, dwrite_1.h, dwrite_2.h, dwrite_3.h, dxcapi.h, dxdiag.h, dxfile.h, dxgicommon.h, dxgicommon.idl, dxgidebug.h, dxgi.h, dxgi.idl, dxgi1_2.h, dxgi1_2.idl, dxgi1_3.h, dxgi1_3.idl, dxgi1_4.h, dxgi1_4.idl, dxgi1_5.h, dxgi1_5.idl, dxgi1_6.h, dxgi1_6.idl, dxgi1_7.h, dxgi1_7.idl, DXGI_Messages.h, dxgitype.h, dxgitype.idl, dxmini.h, dxva.h, dxva2api.h, dxva2api.idl, dxva2SWDev.h, dxva2trace.h, dxva9typ.h, dxvahd.h, dxvahd.idl, dyngraph.idl, eaptypes.h, edevdefs.h, ElsCore.h, ElsSrvc.h, enclaveapi.h, enclaveium.h, endpointvolume.h, errhandlingapi.h, evcode.h, evntcons.h, evntprov.h, evntrace.h, evr.h, evr.idl, evr9.h, ExDisp.h, ExDisp.idl, ExDispid.h, fci.h, fdi.h, fdi_fcitypes.h, fhcfg.h, fherrors.h, fhstatus.h, fhsvcctl.h, fileapi.h, Filter.h, Filterr.h, fltUser.h, fltUserStructures.h, fwpstypes.h, fwpmu.h, Gb18030.h, gdiplusflat.h, GdiplusImaging.h, gl.h, glu.h, GPEdit.h, handleapi.h, hidsdi.h, hidclass.h, hidpi.h, hidusage.h, highlevelmonitorconfigurationapi.h, hrtfapoapi.h, htiface.h, htiface.idl, htiframe.h, htiframe.idl, http.h, icmpapi.h, icodecapi.h, icm.h, ifdef.h, il21dec.h, imagehlp.h, ime.h, imm.h, imapi.h, InputPanelConfiguration.h, interlockedapi.h, ioapiset.h, ioringapi.h, ipexport.h, iphlpapi.h, ipifcons.h, ipmib.h, iprtrmib.h, ipsectypes.h, iptypes.h, issper16.h, iwscapi.h, iwstdec.h, joystickapi.h, jobapi.h, jobapi2.h, keycredmgr.h, KnownFolders.h, ksopmapi.h, ktmtypes.h, ktmw32.h, l2cnm.h, libloaderapi.h, libloaderapi2.h, limits.h, lm.h, lmaccess.h, lmalert.h, lmapibuf.h, lmat.h, lmaudit.h, lmconfig.h, lmcons.h, lmdfs.h, lmerr.h, lmerrlog.h, lmjoin.h, LMJoin.h, lmmsg.h, LMMsg.h, lmremutl.h, lmrepl.h, lmserver.h, LMServer.h, lmshare.h, LMShare.h, lmsname.h, lmstats.h, lmsvc.h, lmuse.h, lmuseflg.h, lmwksta.h, lsalookup.h, lsalookupi.h, magnification.h, manipulations.h, manipulations.idl, mcx.h, mciapi.h, mediaerr.h, mediaobj.h, mediaobj.idl, medparam.h, medparam.idl, memoryapi.h, minidumpapiset.h, minschannel.h, minwinbase.h, mixerocx.h, mixerocx.idl, mmdeviceapi.h, mmeapi.h, mmiscapi2.h, mmsyscom.h, mmsystem.h, mobsync.h, Mpeg2Bits.h, Mpeg2Data.h, Mpeg2Data.idl, Mpeg2PsiParser.idl, mpeg2structs.h, Mpeg2Structs.idl, mprapidef.h, mscat.h, msdelta.h, MSAAText.h, MSAAText.idl, mssign.h, mssip.h, msime.h, msimeapi.h, msports.h, mstcipip.h, muiload.h, nb30.h, ncrypt.h, ncryptprotect.h, ndis\ObjectHeader.h, netioapi.h, newdev.h, nldef.h, ntenclv.h, ntddvdeo.h, ntddvol.h, ntioring_x.h, ntlsa.h, NTQuery.h, objbase.h, objidlbase.h, objidlbase.idl, ObjSel.h, objsafe.h, objsafe.idl, odbcinst.h, ole2.h, oleacc.h, OleDlg.h, oleauto.h, olectl.h, opmapi.h, opmapi.idl, PathCch.h, perlib.h, photoacquire.h, physicalmonitorenumerationapi.h, playsoundapi.h, poclass.h, PortableDevice.h, PortableDeviceAPI.h, PortableDeviceAPI.idl, portabledeviceclassextension.h, portabledeviceclassextension.idl, portabledevicetypes.h, portabledevicetypes.idl, powerbase.h, powersetting.h, powrprof.h, presentation.idl, presentationtypes.h, presentationtypes.idl, processenv.h, ProcessSnapshot.h, processthreadsapi.h, processtopologyapi.h, propapi.h, propidl.h, propidl.idl, propidlbase.h, propidlbase.idl, propkey.h, propkeydef.h, propsys.h, propsys.idl, propvarutil.h, prsht.h, prsht.idl, psapi.h, qmgr.h, qnetwork.h, qos.h, qos2.h, qosobjs.h, qossp.h, RadioMgr.h, RadioMgr.idl, realtimeapiset.h, Reconcil.h, regbag.h, regbag.idl, relogger.h, relogger.idl, restartmanager.h, richedit.h, richole.h, RTWorkQ.h, SCardErr.h, schannel.h, sddl.h, secext.h, security.h, securitybaseapi.h, Sens.h, SensAPI.h, SensEvts.idl, servprov.h, ServProv.Idl, setupapi.h, sfc.h, shappmgr.h, shappmgr.idl, shdeprecated.h, shellapi.h, ShellScalingApi.h, sherrors.h, shimgdata.h, shlobj.h*, ShlObj_core.h, shobjidl.h, ShObjIdl.idl, ShObjidl_core.h, ShObjidl_core.idl, shtypes.h, shtypes.idl, slerror.h, sliddefs.h, slpublic.h, spapidef.h, spatialaudioclient.h, spatialaudiohrtf.h, spatialaudiometadata.h, spellcheck.h, spellcheck.idl, spellcheckprovider.h, spellcheckprovider.idl, SpOrder.h, sql.h, sqlext.h, sqltypes.h, sqlucode.h, sspi.h, SrRestorePtApi.h, StorageProvider.h, StorageProvider.idl, strmif.h, strmif.idl, stringsetapi.h, structuredquerycondition.h, structuredquerycondition.idl, SubAuth.h, synchapi.h, sysinfoapi.h, tcpestats.h, tcpmib.h, tdh.h, TextServ.h, TextStor.h, TextStor.idl, threadpoolapiset.h, threadpoollegacyapiset.h, thumbcache.h, thumbcache.idl, thumbnailstreamcache.h, thumbnailstreamcache.idl, timeapi.h, timezoneapi.h, TlHelp32.h, tlogstg.h, tlogstg.idl, tokenbinding.h, TOM.h, traffic.h, UIRibbon.h, UIRibbon.idl, UIRibbonKeydef.h, UIRibbonPropertyHelpers.h, UIAnimation.h, UIAnimation.idl, UIAutomationClient.h, UIAutomationClient.idl, UIAutomationCore.h, UIAutomationCore.idl, upnp.h, upnp.idl, upnphost.h, upnphost.idl, urlmon.h, urlmon.idl, UserEnv.h, usp10.h, UtilApiSet.h, uuids.h, uxtheme.h, vbinterf.h, vds.h, vds.idl, vdscmmn.idl, vdscmprv.idl, vdserr.h, vdshpcm.idl, vdshp.idl, vdshwprv.h, vdshwprv.idl, vdslun.h, vdslun.idl, vdssp.idl, vdssys.h, vdsvd.idl, VersionHelpers.h, videoacc.h, videoacc.idl, vidcap.h, vidcap.idl, virtdisk.h, vmr9.h, vmr9.idl, vmrender.idl, vsprvcm.idl, vsstyle.h, vssym32.h, WbemCli.h, WbemCli.idl, WcmApi.h, WcnApi.h, WcnDevice.h, WcnFunctionDiscoveryKeys.h, WcnTypes.h, wcsplugin.h, wcsplugin.idl, webauthn.h, webauthnplugin.h, WebEvnts.h, WebEvnts.idl, websocket.h, wimgapi.h, wincodec.h, wincodec.idl, wincodecsdk.h, wincodecsdk.idl, WinBio.h, winbio_err.h, winbio_ioctl.h, winbio_types.h, wincon.h, wincontypes.h, wincred.h, wincrypt.h, WinDNS.h, windef.h, WinEFS.h, wingdi.h, winhttp.h, wininet.h, WinML.h, WinNls.h, WinNls32.h, winnetwk.h, winperf.h, winreg.h, winres.h, winsafer.h, winscard.h, winsmcrd.h, winspool.h, winstring.h, winternl.h, wintrust.h, winuser.h, winver.h, wlanapi.h, wlantypes.h, wldp.h, WMIUtils.h, WMIUtils.idl, wofapi.h, wmcontainer.h, wmistr.h, wscapi.h, ws2atm.h, ws2bth.h, ws2def.h, ws2ipdef.h, ws2tcpip.h, wsman.h, x3daudio.h, xact3.h, xact3d3.h, xact3wb.h, xapo.h, xapofx.h, xaudio2.h, xaudio2fx.h, XInput.h, XmlDom.idl, xma2defs.h, xprtdefs.h, zmouse.h


- Coverage in the 90%+ range\
winbase.h,  oleidl.h,  oaidl.h,  ocidl.h,  ocidl.idl,  , objidl.h, objidl.idl, presentation.h, ScrnSave.h, winsock.h

- Substantial coverage\
mmsciapi.h,  winnt.h, immdev.h, winioctl.h, mmreg.h, WS2spi.h, winerror.h, WindowsSearchErrors.h, windowsx.h, 

- Minimal coverage\
windot11.h, peninputpanel.h, xapobase.h, wmcodecdsp.h, ksmedia.h, d3dkmdt.h, d3dukmdt.h, d3dkmthk.h, fwpsu.h

  
