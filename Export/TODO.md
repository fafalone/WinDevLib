Known omissions where a priority update neccessitated not finishing 
before release, or waiting on feature support:
 
- Check SymGetSetHomeDirectory SymSrvGetFileIndexInfo
- d3d8 flag args

- Verify dinput static lib replacements

- copy  new NT sync APIs to tbKMode.

- Variant/PROPVARIANT overloads / As Any to overload, pending tB bug fix

- Finish DirectShow (now 98% done)
 
- Verify new xaudio2 and helper math functions

- Finish D2D1 helper class pending overload class bug fix
 
- ntlsa delegates

- BUG FIX PENDING: Overloaded interface vtable order vs source, reverse after patch:
IDCompositionVisual, IDCompositionVisual3, IDCompositionGaussianBlurEffect, IDCompositionBrightnessEffect, 
IDCompositionColorMatrixEffect, IDCompositionShadowEffect, IDCompositionHueRotationEffect, IDCompositionSaturationEffect,
IDCompositionLinearTransferEffect, IDCompositionTableTransferEffect, IDCompositionArithmeticCompositeEffect,
IDCompositionAffineTransform2DEffect, IDCompositionTranslateTransform, IDCompositionScaleTransform, IDCompositionRotateTransform,
IDCompositionSkewTransform, IDCompositionMatrixTransform, IDCompositionEffectGroup, IDCompositionTranslateTransform3D,
IDCompositionScaleTransform3D, IDCompositionRotateTransform3D, IDCompositionMatrixTransform3D, IDCompositionRectangleClip,
ID2D1SvgStrokeDashArray, IDWriteGdiInterop1, IDWriteFontFace4, IDWriteFactory4, IDWriteFontSet1
Note: ID2D1SvgElement overloads currently left tagged because tB cannot disambiguate 2 of them. 

- Implements/PreserveSig support IMFTimedTextNotify,IMFMediaSourceExtensionNotify,IMFBufferListNotify,IMFBufferListNotify
IMFMediaEngineNeedKeyNotify,IMFMediaEngineEMENotify,IMFMediaKeySessionNotify2

- OpenGL (needs Alias support)

- Dispinterfaces pending support:
    DWebBrowserEvents[2], DShellNameSpaceEvents 
    _ISpeechRecoContextEvents, _ISpeechVoiceEvents
dispinterface XMLDOMDocumentEvents 



- POLID_ guids only partially complete

- winbio_adapter.h
 
- vfw.h/vfsmsg/vfwext (AVIFil32 done; other sections under consideration)
  
- SQL.h/obdc32.dll (needs Alias support)

- test IsVariantString / IsPropVariantString

- Tablet interfaces (waiting on dispinterface support)

- ndfapi.h

- When enum descriptions supported, convert NErr_ netapi32 error values to enum for APIs

- NT reg VR differencing

- bluetooth btdef.h macros

- Convert additional static lib stuff to Emit()

- Put   delegates back
 

- Macros\
#define D3D_SET_OBJECT_NAME_N_A(pObject, Chars, pName) (pObject)->SetPrivateData(WKPDID_D3DDebugObjectName, Chars, pName)
#define D3D_SET_OBJECT_NAME_A(pObject, pName) D3D_SET_OBJECT_NAME_N_A(pObject, lstrlenA(pName), pName)
#define D3D_SET_OBJECT_NAME_N_W(pObject, Chars, pName) (pObject)->SetPrivateData(WKPDID_D3DDebugObjectNameW, Chars*2, pName)
#define D3D_SET_OBJECT_NAME_W(pObject, pName) D3D_SET_OBJECT_NAME_N_W(pObject, wcslen(pName), pName)
 
- advpub.h

- d3d9xmath.h; D3DX10 (base D3D10 covered; X is from the DX SDK, not Windows Platform SDK). D3DX11 already done. No 12.

- ID3D12GraphicsCommandList::OMSetBlendFactor, ID3D11DeviceContext::OMSetBlendState, ID3D11DeviceContext1::ClearView,
 ID3D12GraphicsCommandList::::ClearUnorderedAccessViewUint, ID3D12GraphicsCommandList::::ClearUnorderedAccessViewFloat
   et al adjust when tB has syntax support for [in] type t[x] syntax.

- (COMPLETED) CLSIDs for coclasses\
keep file list for other tasks\
wdAccessible,\
wdAMSI,\
wdAPI* 
wdBITS,\ 
wdCOM,\
wdCoreAudio,\
wdCredProv,\
wdD3D*\
wdDefs,\
wdDeprecated,\
wdDevices,\ 
wdDirectML,\
wdDirectShow,\
wdDSound,\ 
wdDStorage,\
wdDXVA,\
wdExplorer,\ 
wdFileHist,\
wdGDIP,\
wdGP,\
wdHelpers,\
wdIID,\
wdIMAPI,\
wdInternet,\ 
wdLegacy,\
wdManipulations,\ 
wdNetcon,\ 
wdOLE,\
wdOPC,\
wdPhotoAcq,\
wdPKEY,\
wdPrintNotify,\
wdRadio,\
wdRTWorkQ,\
wdScript,\
wdSearch,\
wdSecurity,\ 
wdSensors,\ 
wdShellCore,\ 
wdShellObj,\ 
wdSpeech,\ 
wdSpellCheck,\ 
wdSyncMgr,\ 
wdTablet,\ 
wdTaskScheduler,\ 
wdUIAnimation,\ 
wdUIRibbon,\ 
wdUPNP,\ 
wdWIC,\
wdWinML,\
wdWinRTBase,\
wdWMDM,\
wdWSC,\ 
wdWTS,\
wdXAudio

- Create coverage list by header. (WORK IN PROGRESS: So long as this list is in todo.md it's incomplete--numerous other headers are covered.)
Excluded from completed %:
    - Definitions unsupported by the tB language with no reasonable substitute.
    - Definitions disabled by conditional compilation with version flags for XP and earlier, non-Windows platforms, or kernel mode only.
Basic Coverage: excludes macros, callbacks->delegates, ANSI APIs (though most are covered), and other headers from #include statements.\
Anything else missing is a bug and a report should be filed
Verified 100% basic coverage (for SDK 10.0.22621.0 minimum, most for 10.0.26100.0);  winuser.h,UtilApiSet.h,processtopologyapi.h,msdelta.h,handleapi.h,cfgmgr32.h,ole2.h,avrt.h,KnownFolders.h,keycredmgr.h,mcx.h,windef.h,winver.h,dlgs.h,
    realtimeapiset.h,msime.h,msimeapi.h,ws2bth.h,VersionHelpers.h,minwinbase.h,wsman.h,wcmapi.h,nb30.h,GPEdit.h,InputPanelConfiguration.h,commoncontrols.h,
    WinEFS.h,winstring.h,qos2.h,traffic.h,qosobjs.h,qos.h,qossp.h,bluetoothleapis.h,bluetoothapis.h,bthsdpdef.h,fhcfg.h,fhsvcctl.h,fhstatus.h,fherrors.h,
    fwpmu.h,ipsectypes.h,iketypes.h,fdi_fcitypes.h,fdi.h,fci.h,namespaceapi.h,physicalmonitorenumerationapi.h,highlevelmonitorconfigurationapi.h,combaseapi.h,
    lowlevelmonitorconfigurationapi.h,wmistr.h,evntcons.h,wincodec.h,wincodec.idl,wincodecsdk.h,wincodecsdk.idl,WinML.h,sysinfoapi.h,cderr.h,commctrl.h,
    ShellScalingApi.h,imagehlp.h,dbghelp.h,interlockedapi.h,upnp.h,upnp.idl,upnphost.h,upnphost.idl,RTWorkQ.h,wlanapi.h,magnification.h,threadpoolapiset.h,
    cfapi.h,amsi.h,tokenbinding.h,WcnApi.h,WcnTypes.h,WcnDevice.h,WcnFunctionDiscoveryKeys.h,lmserver.h,cimfs.h,icmpapi.h,LMJoin.h,LMMsg.h,LMShare.h,
    ObjSel.h,DSClient.h,security.h,minschannel.h,sspi.h,issper16.h,credssp.h,vbinterf.h,vdserr.h,vdscmprv.idl,vsprvcm.idl,vdshwprv.idl,vdscmmn.idl,
    vdslun.idl,vdssp.idl,vdshp.idl,vdsvd.idl,vds.idl,vdshpcm.idl,vds.h,vdshwprv.h,vdslun.h,vdssys.h,directml.h,restartmanager.h,dde.h,ddeml.h,zmouse.h,
    winbio_err.h,winbio_ioctl.h,winbio_types.h,winbio.h,winsvc.h,mssign.h,shellapi.h,bits.idl,bits1_5.idl,bits2_0.idl,bits2_5.idl,bits3_0.idl,bits4_0.idl,
    bits5_0.h,bits10_1.h,bits10_2.h,bits10_3.h,bitscfg.h,qmgr.h,bits2_0.h,bits2_5.h,bits3_0.h,bits4_0.h,bits5_0.idl,bits10_1.idl,bits10_2.idl,bits10_3.idl,
    bitscfg.h,qmgr.h,bitsmsg.h,dwrite.h,dwrite_1.h,dwrite_2.h,dwrite_3.h,shdeprecated.h,UserEnv.h,mscat.h,processenv.h,netioapi.h,iwscapi.h,wscapi.h,http.h,
    WebEvnts.idl,WebEvnts.h,propkey.h,propkeydef.h,winsafer.h,powerbase.h,powersetting.h,powrprof.h,synchapi.h,dpa_dsa.h,DocumentTarget.idl,DocumentTarget.h,
    propsys.h,SrRestorePtApi.h,compressapi.h,wincrypt.h,dpapi.h,mssip.h,memoryapi.h,wintrust.h,bcrypt.h,ncrypt.h,ncryptprotect.h,mobsync.h,ProcessSnapshot.h,
    wincred.h,winhttp.h,websocket.h,photoacquire.h,oleacc.h,sddl.h,securitybaseapi.h,dssec.h,oleauto.h,olectl.h,newdev.h,processthreadsapi.h,virtdisk.h,
    cryptuiapi.h,limits.h,evntrace.h,evntprov.h,relogger.h,relogger.idl,WinNls.h,WinNls32.h,ktmw32.h,fileapi.h,AccCtrl.h,AclAPI.h,dbt.h,TlHelp32.h,winnetwk.h,
    enclaveapi.h,wincon.h,wincontypes.h,consoleapi.h,consoleapi2.h,consoleapi3.h,winreg.h,lsalookup.h,adtgen.h,authz.h,cfg.h,sfc.h,secext.h,AudioAPOTypes.h,
    audioclient.h,audioclient.idl,audioclientactivationparams.h,audioendpoints.h,audioendpoints.idl,audioenginebaseapo.h,audioenginebaseapo.idl,
    audioengineendpoint.h,audioengineendpoint.idl,audiomediatype.h,audiomediatype.idl,audiostatemonitorapi.h,audiopolicy.h,audiopolicy.idl,audiosessiontypes.h,
    devicetopology.h,devicetopology.idl,endpointvolume.h,mmdeviceapi.h,spatialaudioclient.h,spatialaudiohrtf.h,spatialaudiometadata.h,l2cnm.h,wlantypes.h,
    ndis\ObjectHeader.h,eaptypes.h,gdiplusflat.h,imapi.h,manipulations.h,manipulations.idl,propsys.h,propsys.idl,spellcheck.h,spellcheck.idl,spellcheckprovider.h,
    spellcheckprovider.idl,structuredquerycondition.h,structuredquerycondition.idl,threadpoollegacyapiset.h,thumbcache.h,thumbcache.idl,thumbnailstreamcache.h,
    thumbnailstreamcache.idl,timezoneapi.h,TOM.h,tdh.h,winres.h,UIRibbon.h,UIRibbon.idl,UIRibbonKeydef.h,UIRibbonPropertyHelpers.h,UIAutomationClient.h,
    UIAutomationClient.idl,UIAutomationCore.h,UIAutomationCore.idl,UIAnimation.h,UIAnimation.idl,shtypes.h,shtypes.idl,TextServ.h,servprov.h,ServProv.Idl,
    shappmgr.h,shappmgr.idl,shobjidl.h,ShObjIdl.idl,shlobj.h*,richedit.h,richole.h,shimgdata.h,uxtheme.h,mmsystem.h,mmsyscom.h,mciapi.h,mmiscapi2.h,playsoundapi.h,
    mmeapi.h,timeapi.h,joystickapi.h,commdlg.h,cderr.h,prsht.h,prsht.idl,comctrl.h,ShlObj_core.h,ShObjidl_core.h,ShObjidl_core.idl,credentialprovider.h,
    credentialprovider.idl,aclui.h,RadioMgr.h,RadioMgr.idl,PortableDevice.h,PortableDeviceAPI.h,PortableDeviceAPI.idl,portabledeviceclassextension.h,
    portabledeviceclassextension.idl,portabledevicetypes.h,portabledevicetypes.idl,dsound.h,WinDNS.h,dstorage.h,dstorageerr.h,wininet.h,propapi.h,
    propidl.h,propidl.idl,propidlbase.h,propidlbase.idl,propsys.idl,propsys.h,propvarutil.h,Xinput.h,winperf.h,perlib.h,spapidef.h,devpropdef.h,devpkey.h,
    devguid.h,setupapi.h,prnasnot.h,winspool.h,libloaderapi.h,libloaderapi2.h,ioapiset.h,wingdi.h,coml2api.h,evr9.h,dxgi_1.h,dxgi_1.idl,dxgi_2.h,dxgi_2.idl,
    dxgi_3.h,dxgi_3.idl,dxgi_4.h,dxgi_4.idl,dxgi_5.h,dxgi_5.idl,dxgi_6.h,dxgi_6.idl,DXGI_Messages.h,dxgitype.h,dxgitype.idl,dxgicommon.h,dxgicommon.idl,dxgidebug.h,
    ntlsa.h,vsstyle.h,vssym32.h,usp10.h,xapo.h,xaudio2.h,xaudio2fx.h,x3daudio.h,hrtfapoapi.h,WpdShellExtension.h,WpdMtpExtensions.h,evr.h,evr.idl,d3d11.h,d3d11.idl,
    d3d11_2.h,d3d11_2.idl,d3d11_3.h,d3d11_3.idl,d3d11_4.h,d3d11_4.idl,d3d11on12.h,d3d11on12.idl,d2d1.h,d2d1_1.h,d2d1_2.h,d2d1_3.h,d2d1effectauthor.h,d2d1effects.h,
    d2d1effects_1.h,d2d1effects_2.h,d2d1EffectauIEVRTrustedVideoPluginthor.h,d3dcommon.h,d3dcommon.idl,d3d10.h,d3d10.idl,d3d10misc.h,d3d10shader.h,d3d10effects.h,
    d3d10sdklayers.h,d3d10sdklayers.idl,d3d10_1shader.h,d3d10_1.h,d3d10_1.idl,d3dcsx.h,presentation.idl,presentationtypes.h,presentationtypes.idl,wldp.h,webauthn.h,
    ActivityCoordinator.h,ActivityCoordinatorTypes.h,ActivScp.h,ActivScp.idl,atacct.h,lm.h,lmcons.h,lmaccess.h,lmalert.h,lmapibuf.h,lmat.h,lmaudit.h,lmconfig.h,
    lmerrlog.h,lmjoin.h,lmmsg.h,lmremutl.h,lmrepl.h,lmserver.h,lmshare.h,lmsname.h,lmstats.h,lmsvc.h,lmuse.h,lmuseflg.h,lmwksta.h,lmerr.h,lmdfs.h,poclass.h,
    datetimeapi.h,ElsCore.h,ElsSrvc.h,Gb18030.h,stringsetapi.h,ime.h,imm.h,tcpestats.h,tcpmib.h,mprapidef.h,ipifcons.h,ifdef.h,nldef.h,ipmib.h,iprtrmib.h,
    ipexport.h,iptypes.h,iphlpapi.h,winsmcrd.h,SCardErr.h,winscard.h,schannel.h,axcore.idl,devenum.idl,axextendedenums.h,mediaerr.h,dxva.h,dxva9typ.h,dxva2api.h,
    dxva2api.idl,dxvahd.h,dxvahd.idl,icodecapi.h,wmcontainer.h,medparam.h,medparam.idl,mediaobj.h,mediaobj.idl,dmoreg.h,ksopmapi.h,opmapi.h,opmapi.idl,
    d3dx9core.h,d3d9x.h,d3d9xshader.h,d3d9xtex.h,d3dx9xof.h,d3dx9mesh.h,d3dx9shape.h,d3dx11core.h,d3dx11tex.h,d3dx11async.h,d3d12compatibility.h,d3d12compatibility.idl,
    d3d12shader.h,d3d12video.h,d3d12video.idl,d3d12.h,d3d12.idl,fltUserStructures.h,fltUser.h,SpOrder.h,Filter.h,Filterr.h,NTQuery.h,apiquery2.h,appnotify.h,
    cpl.h,cplext.h,ddraw.h,ddstream.h,ddstream.idl,vmr9.h,vmr9.idl,vmrender.idl,amvideo.h,Dvp.h,uuids.h,amaudio.h,evcode.h,dyngraph.idl,dvdmedia.h,edevdefs.h,
    xprtdefs.h,axextend.idl,amparse.h,vidcap.h,vidcap.idl,dmodshow.h,dmodshow.idl,CameraUIControl.h,CameraUIControl.idl,austream.h,austream.idl,qnetwork.h,il21dec.h,
    iwstdec.h,dvdif.h,strmif.h,strmif.idl,control.h,control.idl,amstream.h,amstream.idl,amva.h,sherrors.h,bcp47mrm.h,regbag.h,regbag.idl,wimgapi.h,lsalookupi.h,
    bdatypes.h,bdaiface_enums.h,bdaiface.h,bdaiface.idl,mpeg2structs.h,Mpeg2Structs.idl,Mpeg2Bits.h,Mpeg2Data.h,Mpeg2Data.idl,Mpeg2PsiParser.idl,AtscPsipParser.h,
    callobj.h,callobj.idl,WbemCli.h,WbemCli.idl,WMIUtils.h,WMIUtils.idl,dinput.h,icm.h,wcsplugin.h,wcsplugin.idl,jobapi.h,jobapi2.h,mixerocx.h,mixerocx.idl,
    SensAPI.h,Sens.h,SensEvts.idl,OleDlg.h,dplay8.h,dpaddr.h,dplobby8.h,dpnathlp.h,d3d8.h,d3d8types.h,d3d8caps.h,dxcapi.h,
    tlogstg.h,tlogstg.idl,PathCch.h,appmgmt.h,Dimm.h,Dimm.idl,Reconcil.h,objbase.h,objidlbase.h,objidlbase.idl,MSAAText.h,MSAAText.idl,TextStor.h,TextStor.idl,
    SubAuth.h,davclient.h,DsGetDC.h,errhandlingapi.h,msports.h,objsafe.h,objsafe.idl,winternl.h,AppxPackaging.h,AppxPackaging.idl,XmlDom.idl,wofapi.h,ntddvdeo.h,
    ntddvol.h,hidclass.h,hidusage.h,hidpi.h,hidsdi.h,ntenclv.h,winenclave.h,winenclaveapi.h,ioringapi.h,ntioring_x.h,urlmon.h,urlmon.idl,
    
      
Coverage in the 90%+ range  winbase.h, oleidl.h, oaidl.h, ocidl.h, ocidl.idl, ,objidl.h,objidl.idl,presentation.h,ScrnSave.h

Substantial coverage  mmsciapi.h, winnt.h,immdev.h,winioctl.h,mmreg.h,WS2spi.h,winerror.h,WindowsSearchErrors.h,windowsx.h,

Minimal coverage
    windot11.h,peninputpanel.h,xapobase.h,wmcodecdsp.h,ksmedia.h,d3dkmdt.h,d3dukmdt.h,d3dkmthk.h
    
Zero or near-zero coverage  (all other files)

 