Known omissions where a priority update neccessitated not finishing 
before release, or waiting on feature support:

-OpenGL (needs Alias support)

-DWebBrowserEvents[2], DShellNameSpaceEvents interfaces (needs dispinterface support)

-POLID_ guids only partially complete

-winbio_adapter.h
 
-vfw.h/vfsmsg/vfwext (AVIFil32 done; other sections under consideration)

-missing winspool.h stuff

-windns.h (in progress, 20%)

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
                             wdDevices, wdSyncMgr, wdSensors)