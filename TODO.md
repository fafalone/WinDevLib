Known omissions where a priority update neccessitated not finishing 
before release, or waiting on feature support:
 
-BUG FIX PENDING: Overloaded interface vtable order vs source, reverse after patch:
IDCompositionVisual, IDCompositionVisual3, IDCompositionGaussianBlurEffect, IDCompositionBrightnessEffect, 
IDCompositionColorMatrixEffect, IDCompositionShadowEffect, IDCompositionHueRotationEffect, IDCompositionSaturationEffect,
IDCompositionLinearTransferEffect, IDCompositionTableTransferEffect, IDCompositionArithmeticCompositeEffect,
IDCompositionAffineTransform2DEffect, IDCompositionTranslateTransform, IDCompositionScaleTransform, IDCompositionRotateTransform,
IDCompositionSkewTransform, IDCompositionMatrixTransform, IDCompositionEffectGroup, IDCompositionTranslateTransform3D,
IDCompositionScaleTransform3D, IDCompositionRotateTransform3D, IDCompositionMatrixTransform3D, IDCompositionRectangleClip,
ID2D1SvgStrokeDashArray, IDWriteGdiInterop1, IDWriteFontFace4, IDWriteFactory4, IDWriteFontSet1
Note: ID2D1SvgElement overloads currently left tagged because tB cannot disambiguate 2 of them. 

-OpenGL (needs Alias support)

-Dispinterfaces pending support:
    DWebBrowserEvents[2], DShellNameSpaceEvents 
    _ISpeechRecoContextEvents, _ISpeechVoiceEvents

-POLID_ guids only partially complete

-winbio_adapter.h
 
-vfw.h/vfsmsg/vfwext (AVIFil32 done; other sections under consideration)
 

-SQL.h/obdc32.dll (needs Alias support)

-test IsVariantString / IsPropVariantString

-Tablet interfaces (waiting on dispinterface support)

-ndfapi.h

-NT reg VR differencing

-bluetooth btdef.h macros

-Convert additional static lib stuff to Emit()

-Put   delegates back

 
-advpub.h

-ID3D12GraphicsCommandList::OMSetBlendFactor, ID3D11DeviceContext::OMSetBlendState, ID3D11DeviceContext1::ClearView,
 ID3D12GraphicsCommandList::::ClearUnorderedAccessViewUint, ID3D12GraphicsCommandList::::ClearUnorderedAccessViewFloat
   et al adjust when tB has syntax support for [in] type t[x] syntax.

-CLSIDs for coclasses
done:  
wdAccessible, 
wdAMSI,
wdAPI* 
wdBITS, 
wdCOM,
wdCoreAudio,
wdCredProv,
wdD3D*
wdDefs,
wdDeprecated,
wdDevices, 
wdDirectML,

wdDSound, 
wdDStorage,
wdDXVA,
wdExplorer, 
wdFileHist,
wdGDIP,
wdGP,
wdHelpers,
wdIID,
wdIMAPI,
wdInternet, 
wdLegacy,
wdManipulations, 
wdNetcon, 








wdSearch,
wdSecurity, 
wdSensors, 
wdShellCore, 
wdShellObj, 
wdSpeech, 
wdSpellCheck, 
wdSyncMgr, 
wdTablet, 
wdTaskScheduler, 
wdUIAnimation, 
wdUIRibbon, 
wdUPNP, 
wdWIC,

wdWSC, 
wdWTS

-Create coverage list by header.
Verified 100% basic coverage (for SDK 10.0.22621.0 minimum, most for 10.0.26100.0); basic coverage excludes macros, callbacks->delegates,
ANSI APIs (though most are covered), and other headers from #include statements. Anything else missing is a bug and a report should be filed.
    UtilApiSet.h,processtopologyapi.h,msdelta.h,handleapi.h,cfgmgr32.h,ole2.h,lmuse.h,lmuseflg.h,lmrepl.h,lmat.h,avrt.h,
    realtimeapiset.h,msime.h,msimeapi.h,ws2bth.h,VersionHelpers.h,minwinbase.h,wsman.h,wcmapi.h,nb30.h,GPEdit.h,InputPanelConfiguration.h,
    WinEFS.h,winstring.h,qos2.h,traffic.h,qosobjs.h,qos.h,qossp.h,bluetoothleapis.h,bluetoothapis.h,bthsdpdef.h,fhcfg.h,fhsvcctl.h,fhstatus.h,fherrors.h,
    fwpmu.h,ipsectypes.h,iketypes.h,fdi_fcitypes.h,fdi.h,fci.h,namespaceapi.h,physicalmonitorenumerationapi.h,highlevelmonitorconfigurationapi.h,
    lowlevelmonitorconfigurationapi.h,wmistr.h,evntcons.h,wincodec.h,hidclass.h,hidusage.h,hidpi.h,hidsdi.h,WinML.h,sysinfoapi.h,cderr.h,
    ShellScalingApi.h,imagehlp.h,dbghelp.h,interlockedapi.h,upnp.h,upnp.idl,upnphost.h,upnphost.idl,RtwqSetLongRunning,RTWorkQ.h,wlanapi.h,magnification.h,
    cfapi.h,amsi.h,tokenbinding.h,WcnApi.h,WcnTypes.h,WcnDevice.h,WcnFunctionDiscoveryKeys.h,lmserver.h,cimfs.h,icmpapi.h,LMJoin.h,LMMsg.h,LMShare.h,
    ObjSel.h,DSClient.h,security.h,minschannel.h,sspi.h,issper16.h,credssp.h,vbinterf.h,vdserr.h,vdscmprv.idl,vsprvcm.idl,vdshwprv.idl,vdscmmn.idl,
    vdslun.idl,vdssp.idl,vdshp.idl,vdsvd.idl,vds.idl,vdshpcm.idl,vds.h,vdshwprv.h,vdslun.h,vdssys.h,directml.h,restartmanager.h,dde.h,ddeml.h,
    winbio_err.h,winbio_ioctl.h,winbio_types.h,winbio.h,winsvc.h,mssign.h,shellapi.h,bits.idl,bits1_5.idl,bits2_0.idl,bits2_5.idl,bits3_0.idl,bits4_0.idl,
    bits5_0.h,bits10_1.h,bits10_2.h,bits10_3.h,bitscfg.h,qmgr.h,bits2_0.h,bits2_5.h,bits3_0.h,bits4_0.h,bits5_0.h,bits10_1.h,bits10_2.h,bits10_3.h,
    bitscfg.h,qmgr.h,bitsmsg.h,dwrite.h,dwrite_1.h,dwrite_2.h,dwrite_3.h,shdeprecated.h,UserEnv.h,mscat.h,processenv.h,netioapi.h,iwscapi.h,wscapi.h,
    WebEvnts.idl,WebEvnts.h,propkey.h,propkeydef.h,winsafer.h,powerbase.h,powersetting.h,powrprof.h,synchapi.h,dpa_dsa.h,DocumentTarget.idl,DocumentTarget.h,
    propsys.h,SrRestorePtApi.h,compressapi.h,wincrypt.h,dpapi.h,mssip.h,memoryapi.h,wintrust.h,bcrypt.h,ncrypt.h,ncryptprotect.h,mobsync.h,ProcessSnapshot.h,
    wincred.h,winhttp.h,websocket.h,photoacquire.h,oleacc.h,sddl.h,securitybaseapi.h,dssec.h,oleauto.h,olectl.h,newdev.h,processthreadsapi.h,
    cryptuiapi.h,limits.h,winuser.h,evntrace.h,WinNls.h,ktmw32.h,fileapi.h,AccCtrl.h,AclAPI.h,dbt.h,TlHelp32.h,winnetwk.h,virtdisk.h,threadpoolapiset.h,
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
    devguid.h,setupapi.h,prnasnot.h,winspool.h,libloaderapi.h,libloaderapi2.h,ioapiset.h,wingdi.h,coml2api.h,dxgi_1.h,dxgi_1.idl,dxgi_2.h,dxgi_2.idl,
    dxgi_3.h,dxgi_3.idl,dxgi_4.h,dxgi_4.idl,dxgi_5.h,dxgi_5.idl,dxgi_6.h,dxgi_6.idl,DXGI_Messages.h,dxgitype.h,dxgitype.idl,dxgicommon.h,dxgicommon.idl,
    
Coverage in the 90%+ range
    winbase.h, oleidl.h, oaidl.h, 

Substantial coverage
    mmsciapi.h, ntlsa.h (100% verified through line 3130), winnt.h,winternl.h

Minimal coverage
    windot11.h
    peninputpanel.h

Zero or near-zero coverage
    


    * - Excludes definitions disabled by conditional compilation with version flags for XP and earlier or non-Windows platforms.
    