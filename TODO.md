Known omissions where a priority update neccessitated not finishing 
before release, or waiting on feature support:
 
-Check WV2 at end of March

-OpenGL (needs Alias support)

-Dispinterfaces pending support:
    DWebBrowserEvents[2], DShellNameSpaceEvents 
    _ISpeechRecoContextEvents, _ISpeechVoiceEvents

-POLID_ guids only partially complete

-winbio_adapter.h
 
-vfw.h/vfsmsg/vfwext (AVIFil32 done; other sections under consideration)

-missing winspool.h stuff 

-SQL.h/obdc32.dll (needs Alias support)

-test IsVariantString / IsPropVariantString

-Tablet interfaces (waiting on dispinterface support)

-ndfapi.h

-NT reg VR differencing

-bluetooth btdef.h macros

-Convert additional static lib stuff to Emit()

-Put   delegates back

 
-advpub.h

-CLSIDs for coclasses (done: wdShellCore, wdShellObj, wdInternet, wdAccessible, wdBITS, wdUIRibbon, wdWIC, wdCoreAudio,
                             wdSecurity, wdTaskScheduler, wdManipulations, wdSpellCheck, wdExplorer, wdSearch,
                             wdDevices, wdSyncMgr, wdSensors, wdWSC, wdUPNP, wdUIAnimation, wdSpeech, wdNetcon, wdCredProv,
                             wdDSound)
                             
-Create coverage list by header.
Verified 100% basic coverage (for SDK 10.0.22621.0 minimum, most for 10.0.26100.0); basic coverage excludes macros, callbacks->delegates,
ANSI APIs (though most are covered), and other headers from #include statements.
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
    ndis\ObjectHeader.h,eaptypes.h,gdiplusflat.h,imapi.h,manipulations,manipulations.idl,propsys.h,propsys.idl,spellcheck.h,spellcheck.idl,spellcheckprovider.h,
    spellcheckprovider.idl,structuredquerycondition.h,structuredquerycondition.idl,threadpoollegacyapiset.h,thumbcache.h,thumbcache.idl,thumbnailstreamcache.h,
    thumbnailstreamcache.idl,timezoneapi.h,TOM.h,tdh.h,winres.h,UIRibbon.h,UIRibbon.idl,UIRibbonKeydef.h,UIRibbonPropertyHelpers.h,UIAutomationClient.h,
    UIAutomationClient.idl,UIAutomationCore.h,UIAutomationCore.idl,UIAnimation.h,UIAnimation.idl,shtypes.h,shtypes.idl,TextServ.h,servprov.idl,ServProv.Idl,
    shappmgr.h,shappmgr.idl,shobjidl.h,ShObjIdl.idl,shlobj.h*,richedit.h,richole.h,shimgdata.h,uxtheme.h,mmsystem.h,mmsyscom.h,mciapi.h,mmiscapi2.h,playsoundapi.h,
    mmeapi.h,timeapi.h,joystickapi.h,commdlg.h,cderr.h,prsht.h,prsht.idl,comctrl.h,ShlObj_core.h,ShObjidl_core.h,ShObjidl_core.idl,credentialprovider.h,
    credentialprovider.idl,RadioMgr.h,RadioMgr.idl,PortableDevice.h,PortableDeviceAPI.h,PortableDeviceAPI.idl,portabledeviceclassextension.h,
    portabledeviceclassextension.idl,portabledevicetypes.h,portabledevicetypes.idl,dsound.h,WinDNS.h,dstorage.h,dstorageerr.h,wininet.h,propapi.h,
    propidl.h,propidl.idl,propidlbase.h,propidlbase.idl,propsys.idl,propsys.h,propvarutil.h,Xinput.h,winperf.h,perlib.h,spapidef.h,devpropdef.h,devpkey.h,
    devguid.h,setupapi.h,

Coverage in the 90%+ range
   wingdi.h, winbase.h

Substantial coverage
    mmsciapi.h, ntlsa.h

Minimal coverage
    windot11.h
    peninputpanel.h

Zero or near-zero coverage
    


    * - Excludes definitions disabled by conditional compilation with version flags for XP and earlier or non-Windows platforms.
    