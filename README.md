# WinDevLib 
## Windows Development Library for twinBASIC

**Current Version: 8.0.405 (May 17th, 2024)**

(c) 2022-2023 Jon Johnson (fafalone)

This project is a comprehensive twinBASIC replacement for [oleexp.tlb](http://www.vbforums.com/showthread.php?786079-VB6-Modern-Shell-Interface-Type-Library-oleexp-tlb), my Modern Shell Interfaces Type Library project for VB6, that is x64 compatible, due to the many problems using midl to create a 64bit tlb.

This and oleexp are projects to supply Windows shell and component interfaces in a format consumable by VB6/VBA/tB. This involves not only defining interfaces, but using VB/tB compatible types-- so in some cases, even though there may be an existing way to import references to interfaces, they may be unusable due to e.g. the use of unsigned types, C-style arrays, double pointers, etc. All interfaces, types, consts, and APIs from oleexp are covered. For a full list of interfaces, see [INTERFACES.md](https://github.com/fafalone/WinDevLib/blob/main/INTERFACES.md).

In addition to the 2200+ common COM interfaces, WinDevLib now includes expansive coverage of Windows APIs from all the common modules. This makes it similar to working in C++ with `#include <Windows.h>` and a few others. Currently, approximately 6,000 of the most common APIs have been added- redone by hand from the original headers, in order to restore 64bit type info lost in VB6 versions, avoid the errors of automated conversion tools (e.g. Win32API_PtrSafe.txt is riddled with errors), and make them friendlier by converting groups of constants associated with a variable into an Enum so it comes up in Intellisense. This takes advantage of tB's ability to provide Intellisense for types besides Long in API defs (hopefully UDTs soon, this project has provisioning for that). 

This project is implemented purely in tB native code, as unlike VB6 there's language support for defining interfaces and coclasses. As a twinPACKAGE, regular code is supported in addition to the definitions, so the regular addin modules have been built in (mIID.bas, mPKEY.bas, etc). Does it still make sense to use a project like this when interfaces can be defined in-language? I'd say yes, because for a large number of interfaces, there's deep dependency chains with other interfaces and the types they rely on. It makes more sense to drop this in and be done with it than constantly have to define the interfaces you want and then stubs for their dependencies, especially when you might need those later on. This project is even more useful now with the API coverage; it should cover about 99% of your needs for core system DLLS. 

Please report any bugs via the Issues feature here on GitHub.

### Requirements

[twinBASIC Beta 424 or newer](https://github.com/twinbasic/twinbasic/releases) is required, 461 or newer is recommended.

### Adding WinDevLib to your project
You have 2 options for this:

#### Via the Package Server
twinBASIC has an online package server and WinDevLib is published on it. Open your project settings and scroll to the **COM Type Library / ActiveX References**, then click **TWINPACK PACKAGES**. Add "Windows Development Library for twinBASIC v7.0.272" (or whatever the newest version is). The other similar entry, "WinDevLib for Implements" contains `Implements` compatible versions of a small number of common interfaces not defined in a compatible way in the main project; you normally don't need this. For more details, including illustrations, [see this post](https://github.com/fafalone/WinDevLib/issues/9#issuecomment-1416767019).

#### From a local file
You can download the project from this repository and use the WinDevLib.twinpack file. Navigate to the same area as above, and click on the "Import from file" button. WinDevLib.twinproj is the source for the package, if you want to edit it.


### Optional Features

#### Compiler Flags
WinDevLib has some compiler constants you can enable:

`WINDEVLIB_LITE` - This flag disables most API declares and misc WinAPI definitions, including everything in wdAPIComCtl, wdAPI, and wdDefs. I used to like doing my APIs separate too, which is why oleexp never had the expansive coverage. But with that coverage now present, I think it's worth using, but this option will still be supported.

`WINDEVLIB_COMCTL_LIB_DEFINED` - You can use this flag if you already have an alternative common controls definition set, e.g. tbComCtlLib; it will disable wdAPIComCtl. (Note: WinDevLib has more complete comctl defs than tbComCtlLib, as that project was deprecated and not updated).

`WINDEVLIB_DLGH` - This enabled constants from dlg.h. These are extremely uncommon to use, and have very short, generic names likely to cause conflicts, so they're opt-in.

`WINDEVLIB_NOQUADLI` - Restores the old `LARGE_INTEGER` definition of lo/high Long values.

>[!WARNING]
>The `WINDEVLIB_NOQUADLI` constant will break alignment on numerous Types; most only on x64, but some on both. 

`WINDEVLIB_AVOID_INTRINSICS` - Uses the `Interlocked*` APIs that are exported from kernel32.dll (32bit mode only) instead of the static library containing compiler intrinsic versions of those in addition to all the ones not exported and all the 64bit ones.

`WINDEVLIB_NOLIBS` - Fully exclude static libraries (currently only Interlocked); mainly intended for comparing current tB versions to Beta 423 where the `Import Library` syntax is not yet supported.


#### Custom Helper Functions
In addition to coverage of common Windows SDK-defined macros and inlined functions, a small number of custom helper functions are provided to deal with Windows data types and similar not properly supported by the language. These are:

`Public Function LPWSTRtoStr(lPtr As LongPtr, Optional ByVal fFree As Boolean = True) As String`\
Converts a pointer to an LPWSTR/LPCWSTR/PWSTR/etc to an instrinsic `String` (BSTR)

`Public Function UtfToANSI(sIn As String) As String`\
Converts a Unicode string to ANSI. This function is `[ConstantFoldable]` -- it can be used to create strings resolved at compile time and stored as constants; this technique was developed to use ANSI strings in kernel mode, where the APIs that handle a normal `String` cannot  be used.

`Public Function VariantLPWSTRtoSTR(pVar As Variant, pOut As String) As Boolean`\
Retrieves a tB-style String from a VT_LPWSTR Variant. Returns False if pVar is a null pointer, or VT_LPWSTR or PropVariantToStringAlloc returns a nullptr.

`Public Function GetSystemErrorString(lErrNum As Long, Optional ByVal lpSource As LongPtr = 0) As String`\
`Public Function GetNtErrorString(lErrNum As Long) As String`\
Retrieve descriptions of `HRESULT` and `NTSTATUS` error codes, respectively.

`Public Function VariantSetType(pvar As Variant, [TypeHint(VARENUM)] ByVal vt As Integer, [TypeHint(VARENUM)] Optional ByVal vtOnlyIf As Integer = -1) As Boolean`\
Sets a Variant to the specified type without any alteration to the data. vtOnlyIf returns False if the original type is other than specified. This should only be used when `VariantChangeType` is not applicable, and only with full understanding of consequences like automation errors if you attempt to use intrinsic operations on unsupported types; e.g. if you set the type to `VT_UI4`, then `CLng()` will raise a 'type unsupported' runtime error.

`Public Function PointerAdd(ByVal Start As LongPtr, ByVal Incr As LongPtr) As LongPtr`\
`Public Function UnsignedAdd(ByVal Start As Long, ByVal Incr As Long) As Long`\
`Public Function UnsignedAdd(ByVal Start As Integer, ByVal Incr As Integer) As Integer`\
`Public Function UnsignedAdd(ByVal Start As LongLong, ByVal Incr As LongLong) As LongLong`\
Perform addition as if `Start` and `Incr` were unsigned, returning an unsigned result. Important for large address aware operations in 32bit.'

`Public Function SwapVtableEntry(pObj As LongPtr, EntryNumber As Integer, ByVal lpFN As LongPtr) As LongPtr`\
This is the common vtable redirection helper function rewritten to support both 32 and 64bit operations.

`Public Function PointToLongLong(pt As POINT) As LongLong`\
`Public Function PointToLongLong(ByVal ptx As Long, ByVal pty As Long) As LongLong`\
`Public Function PointToLongLong(ByVal ptx As Integer, ByVal pty As Integer) As LongLong`\
`Public Function PointFToLongLong(pt As POINTF) As LongLong`\
`Public Function PointFToLongLong(ptx As Single, pty As Single) As LongLong`\
`Public Function PointSToLong(pt As POINTS) As Long`\
`Public Function SizeToLongLong(z As SIZE) As LongLong`\
`Public Function SizeToLongLong(ByVal cx As Long, ByVal cy As Long) As LongLong`\
Functions for converting POINT and SIZE types or coords to Long or LongLong for methods requiring them to be passed ByVal, which is currently unsupported.

`Public Function CUIntToInt(ByVal Value As Long) As Integer` - Create unsigned Integer from a Long\
`Public Function CIntToUInt(ByVal Value As Integer) As Long` - Convert an Integer to Long as if it were unsigned (&HFFFF = 65536 instead of -1)\
`Public Function CULngToLng(ByVal Value As Double) As Long`\
`Public Function CULngToLng(ByVal Value As LongLong) As Long`\
`Public Function CLngToULng(ByVal Value As Long) As LongLong`\
`Public Sub CLngToULng(ByVal Value As Long, pULng As Double)`


### Guide to switching existing code to WinDevLib

WinDevLib presented the best opportunity there would be to ditch some olelib legacy baggage. It's fairly simple to move your VB6 projects to WinDevLib, just follow these steps:

#### oleexp type library issues
The follow steps  apply only if you're converting code that previously relied on my oleexp.tlb project:

1) Replace public aliases: It's important to do this first. Run a Replace All changing oleexp.LONG_PTR to LongPtr, oleexp.REFERENCE_TIME to LongLong, oleexp.HNSTIME to LongLong, oleexp.KNOWNFOLDERID to UUID, oleexp.EventRegistrationToken to LongLong, oleexp.BINDPTR to LongPtr, and oleexp.LPCRITICAL_SECTION to LongPtr. If you've used them without the oleexp. prefix, you'll also need to replace those, but if you've imported into tB they should be tagged.

2) Replace oleexp.IUnknown with IUnknownUnrestricted. WinDevLib keeps this separate due to the major issues with conflicts with the former approach. If your project has IUnknown *without* oleexp. in front of it, **do not** replace those, as it's not referring to oleexp. 

3) After you've done those two, you can now go ahead and simply delete all remaining instances of `oleexp.` (including the .). 

##### Issues specific to oleexpimp.tlb

There's 'WinDevLib for Implements' (WinDevLibImpl.twinpack/.twinproj) as well, but you'll note it has substantially fewer interfaces than oleexpimp.tlb. This is because there's two reasons for an interface to have an alternate version: It uses `[ Preservesig ]` on one or more methods, or it uses `As Any`. twinBASIC allows using `Implements` with `As Any` by replacing it with `As LongPtr` (which is what the alternate versions do). So many interfaces were in oleexpimp.tlb for this latter reason, and subsequently are *not* included in WinDevLibImpl as it's not neccessary.

If you find an oleexpimp.tlb interface is not in WinDevLibImpl, you will be able to use the one from WinDevLib, simply make sure `As Any` is changed to `As LongPtr`. 

tB has announced plans to support `[ PreserveSig ]` in implemented interfaces in the future; when that happens WinDevLibImpl will be deprecated.

>[!IMPORTANT]
>There currently seems to be an issue with using WinDevLib and WinDevLibImpl together if WinDevLibImpl does not use the current WinDevLib as a reference (it would usually use an old one as it's updated much less frequently). I've updated the reference on this repo and the package server, just note that you'll need to refresh both every time you update one if they're used together
>
>
#### API definition differences
This section applies both to API calls from oleexp.tlb and general `Declare` statements.

1) Convert `Currency` to `LongLong` for interfaces and APIs: It's no longer neccessary to worry about multiplying and dividing by 10,000 since tB supports a true 64bit type in both 32bit and 64bit mode. So this change is ultimately for the better, but existing codebases will have had to have used `Currency` for all interfaces and oleexp APIs expecting a 64bit integer.

2) (From type libraries only) If you see errors about wrong number of arguments or a mismatch on the return type, some oleexp.tlb and other type library defined APIs do a similar change as some interface arguents and rewrite the last argument as a return value. In these cases the return type will now be `Long`, and you can just move the receiving variable to the position of a final argument. If it uses `Set`, you can just drop that. 

3) Optional UDTs no longer use `As Any`. If you see errors like `Validation of call to 'CreateFile' failed.  Argument for 'lpSecurityAttributes': cannot coerce type 'Long' to 'SECURITY_ATTRIBUTES'`, this is an example of the issue. twinBASIC supports substituing `vbNullPtr` for a UDT (do not include `ByVal`), so WinDevLib can use the proper type while still permitting you to pass the equivalent of `ByVal 0`. 

Example:

VB6:
```vba
Public Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

hFile = CreateFileW(StrPtr("name"), 0, 0, ByVal 0, ...)
```
twinBASIC:
```vba
Public Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr

hFile = CreateFileW(StrPtr("name"), 0, 0, vbNullPtr, ...)
'---or---
Dim pSec As SECURITY_ATTRIBUTES
Dim lPtr As LongPtr = VarPtr(pSec)
hFile = CreateFileW(StrPtr("name"), 0, 0, lPtr, ...)
```

4) String vs Long(Ptr) in APIs with both ANSI and Unicode versions: Most VB programs are written with ANSI versions of APIs being the default. **This is not the case with WinDevLib**. With very few exceptions, APIs are Unicode by default-- i.e. they use the W, rather than A, version of APIs e.g. `DeleteFile` maps to `DeleteFileW` rather than `DeleteFileA`. The A and W variants use String/LongPtr, and in almost all cases, the mapped version uses `String` with twinBASIC's `DeclareWide` keyword-- this disables Unicode-ANSI conversion. Since this is automatic, you generally don't need to make any changes; you can still use `String` without `StrPtr` or any manual Unicode <-> ANSI conversion. Note this usually only applies to strings passed as input, you'll need to update any externally allocated strings returned as a pointer, where you previously used e.g. `lstrlenA`, to use `lstrlenW` and Unicode handling in general. 

All APIs are provided, as a minimum, as the explicit W variant, and an untagged version that maps to the W version. Most ANSI variants are also included, but code should use Unicode wherever possible. You generally don't need to change any code in the case of 

UDTs used by these calls are also supplied in the same manner, the W variant, an untagged variant that's the same as the W version, and in some cases, an A version. UDTs always use `LongPtr` for strings, even the untagged versions for `DeclareWide`. 

If you have any doubts about which API is being called, twinBASIC will show the full declaration when you hover your cursor over the API in your code.

Special thanks to GCUser99 for helping normalize API declaration in this project. 👍


> [!TIP]
> Reminder: `Nothing` can be used in place of an interface where WinDevLib has the interface as an argument but another signature used `Long`/`LongPtr`


> [!NOTE]
>  This is just for using WinDevLib-- you'll likely have a lot more changes to make if you want to make your project x64 compatible. 


#### Scope of coverage

The goal of the API coverage in WinDevLib is to provide the kind of programming experience you'd get in C/C++ by including windows.h and some of the more common feature sets like DirectX and GDIPlus. It currently includes about 5,500 APIs. But even that is just scratching the surface of the total Windows API set. Due to the low quality of automated conversion, even by Microsoft themselves (see: Win32API_PtrSafe.txt), I'm not interested in simply feeding headers through a conversion utility or using a database, so instead WinDevLib will be focused on the most commonly used features in the major system DLLs, though less common ones can be added by request or as time goes on and the existing DLLs are completed. I do not intend to include native APIs that have fully equivalent regular APIs; that's basically doubling the work for no benefit-- but if they offer additional features or substantially improved performance, they will be included. 

I've included the definitions, associated types, and associated constants, for extensive portions of the following modules: shell32.dll, shlwapi.dll, kernel32.dll, user32.dll, advapi32.dll, ole32.dll, oleaut32.dll, propsys.dll, gdi32.dll, gdiplus.dll, urlmon.dll, hlink.dll, winmm.dll, setupapi.dll, comctl32.dll, dwm.dll/uxtheme.dll, comdlg32.dll, winspool.drv, and netapi32.dll. Besides highly self-contained specialized sets in their own headers (unless small), please let me know any I've missed from these.\
Limited coverage (or full coverage of very small sets) is provided for ntdll.dll, version.dll, msimg32.dll, crypt32.dll, virtdisk.dll, sxs.dll, secur32.dll, imm32.dll, userenv.dll, wintrust.dll, msacm32.dll, url.dll, htmlhelp.dll, imagehlp.dll, and ws2_32.dll. If you feel any missing ones from these should be included, or would like to contribute more, let me know.\
Finally, there's small API sets for features, like DirectX DLLs, Webview2Loader, WIC, etc. Definitely let me know any missing from these.

**Future coverage:** In the future I'm planning to expand native APIs with no equivalents, add additional Winsock coverage, and add OpenGL-- though for this last one I may wait for tB to have `Alias` support since existing OpenGL codebases make heavy use of them by way of NeHe's TLB. I welcome contributions of any of these. If you've done the consts->enums conversions already, I'd even take 32bit-only versions.


### ***tbShellLib is now WinDevLib - Windows Development Library for twinBASIC***
This project has grown well beyond it's original mission of shell programming. While that's still the largest single part, it's no longer a majority of the code, and the name change now much better reflects the purpose of providing a general Windows API experience like windows.h. Compiler constants and module names/file names have been updated to reflect the name change. tbShellLibImpl is now WinDevLibImpl. There are also some major chanages associated with this update, please see the full changelog below.

### DLL Redirection Errors in Older Versions
 
twinBASIC now counts msvbvm60 redirects as legacy DLL redirects, which WinDevLib set to "Error". Please update to the latest version of WinDevLib to get rid of these errors and use it on twinBASIC Beta 456 and newer. Both this repo and the package server downloads have been updated.
 

### Updates

**Update (v8.0.405, 17 May 2024):**
-(Bug fix) IShellImageDataFactory method names incorrect
-(Bug fix) IShellImageData missing method, breaking 2nd half of interface. Some method names incorrect.

**Update (v8.0.404, 17 May 2024):**\
-Added missing constants from shimgdata.h (now 100% covered)\
-(Bug fix) URLDownloadToCacheFileW and URLDownloadToFileW still using String arguments.

**Update (v8.0.403, 17 May 2024):**\
-(API Standards) URLOpenStream, URLPullStream, URLDownloadToCacheFile, and URLDownloadToFile did not conform to standards, additionally W versions used String without DeclareWide. (Issue #29)

**Update (v8.0.402, 16 May 2024):**\
-Added Magnification API (magnification.h, 100% coverage)\
-Added Cloud Filter APIs (cfapi.h, 100% coverage). Note: These APIs use overloaded String/LongPtr declares, please report any problems.\
-Added Antimalware Scan Interfaces and APIs (amsi.h, 100% coverage). Note: These APIs use overloaded String/LongPtr declares, please report any problems.\
-Added tokenbinding.h/dll APIs (100% coverage)\
-Added Windows Connect Now interfaces/coclass (WcnApi.h, WcnTypes.h, WcnDevice.h, WcnFunctionDiscoveryKeys.h 100%)\
-Added all netapi32 APIs from lmserver.h (100% coverage)\
-Added Composite Image APIs (cimfs.h, 100% coverage)\
-Added AVI file interfaces and APIs from vfw.h\
-Added additional overloads for COM object APIs (e.g. CoMarshalInterThreadInterfaceInStream and CoGetInterfaceAndReleaseStream), to allow using LongPtr in addition to interfaces.\
-Added missing WIC proxy functions WICCreateColorContext_Proxy, WICCreateImagingFactory_Proxy, and WICSetEncoderFormat_Proxy.\
-DragQueryFile[A,W] now uses Optional for the last 2 arguments for compatibility with common usage.\
-DLLVERSIONINFO member names now match SDK\
-IOleInPlaceUIWindow.SetActiveObject now uses LongPtr in place of String for compatibility with OLEGuids\
-IOleInPlaceActiveObject now uses PreserveSig to return the HRESULT on all methods for compatibility with OLEGuids.\
    The original, Implements-compatible version, is now in WinDevLibImpl.\
-(API Standards) CreateFontIndirect now uses LOGFONT instead of LOGFONTW (identical besides name)\
-(API Standards) GetIconInfoEx was using ICONINFOEXW instead of (previously missing) ICONINFOEX.\
-(API Standards) CryptBinaryToString not marked DeclareWide. (Issue #26)\
-(Bug fix) StopTrace and QueryTrace missing aliases (Issue #28)\
-(Bug fix) DrawThemeParentBackgroundEx case incorrect\
-(Bug fix) GetCurrentThemeName missing ByVal on String argument\
-(Bug fix) GetFileVersionInfoA, GetFileVersionInfoSizeA, GetDiskFreeSpaceA incorrectly used W aliases. (Issue #27)\
-(Bug fix) RegCreateKey missing DeclareWide (Issue #27)\
-(Bug fix) Shell library helper functions incorrectly used Null instead of Nothing.\
-(Bug fix) SetFocus missing argument\
**WinDevLibImpl v1.3.18** `Implements`-compatible version of IOleInPlaceActiveObject added.


**Update (v7.10.396, 28 Apr 2024):**\
-**MAJOR CHANGE:** IShellIconOverlay will now no longer require using VarPtr() around the index output var.\
-Added WIC proxy functions (Issue #22)\
-Added iphlpapi ICMP APIs (icmpapi.h, 100%)\
-Added additional netapi32 APIs, LMJoin.h, LMMsg.h, 100%; some missing APIs from LMShare.h (100% now hopefully)\
-Added missing common API CreateBitmap (Issue #21)\
-LVTILEVIEWINFO.SizeTile no longer uses redundant SIZELVT UDT\
-First half of imagehlp.h/dbghelp.h added\
-(Bug fix) PathIsNetworkPathW/PathFindFileNameW were incorrectly misnamed PathIsNetworkPathA/PathFindFileNameW (creating overloads).\
-(Bug fix) BITMAPFILEHEADER definition incorrect (missing non-default packing alignment)\
-(Bug fix) ImageEnumerateCertificates definition incorrect (ByRef/ByVal mixup)\
-(Bug fix) STORAGE_BUS_TYPE values all off by one.


**Update (v7.9.392, 24 Apr 2024):**\
-Added additional security dialog stuff; the Directory Object Picker interfaces/coclass and DsBrowserForContainer API; ObjSel.h, DSClient.h 100%
  
**Update (v7.9.390, 24 Apr 2024):**\
-Large expansion of security APIs from security.h, minschannel.h, sspi.h, issper16.h, and credssp.h\
   All are 100% covered with the exception of kernel-mode only defs in sspi.h.\
-Added new helper function for APIs/COM interfaces expecting a ByVal GUID:\
   UUIDtoLong(UUID, pl1 As Long, pl2 As Long, pl3 As Long, pl4 As Long)\
   UUIDtoLong(UUID, pls() As Long)\
-Added VBA-related interfaces from vbinterf.h (100% coverage)\
-Adjusted custom buffers on DEV_BROADCAST_* types to not leave padding bytes.\
-Added non-aliased versions of RtlMoveMemory, RtlZeroMemory, and RtlFillMemory (Issue #20)\
-(Bug fix) LoadIconMetrics enum had incorrect values and is now also renamed to the proper LI_METRIC name.


**Update (v7.9.386, 19 April 2024):**\
-Added complete Virtual Disk Service interfaces and custom coclass VdsLoader\
-Added DirectML interfaces\
-Added Restart Manager APIs (restartmanager.h, 100% coverage)\
-Added DDE APIs (dde.h, ddeml.h 100%)\
-Added some misc missing extremely common APIs.


**Update (v7.8.382, 17 April 2024):**\
-Added coverage of all Windows Biometric Framework application APIs (winbio_err.h, winbio_ioctl.h, winbio_types.h, winbio.h 100%)\
-Added missing WMDM DRM interfaces/coclass (MS forgot to merge these into the SDK when it abandoned a separate WMDM sdk)\
-Some additional defs to bring winsvc.h coverage to 100%\
-Add some missing WIC GUIDs\
-(Bug fix) SERVICE_REQUIRED_PRIVILEGES_INFO[W] definitions incorrect for 64bit\
-(Bug fix) EnumServicesStatusEx, GetServiceDisplayName incorrect alias\
-(Bug fix) QueryServiceStatusEx, QueryServiceDynamicInformation, GetServiceRegistryStateKey, GetServiceDirectory, GetSharedServiceDirectory, RegisterServiceCtrlHandler[A,W,Ex,ExA,ExW] definitions incorrect for 64bit (Ex incorrect alias as well)\
-(Bug fix) QueryServiceStatusEx incorrect additional overload\
-(Bug fix) SECURITY_MAX_SID_SIZE value incorrect


**Update (v7.8.379, 12 April 2024):**\
-Large expansion of Direct3D 12 interfaces to cover latest SDK version of d3d12.idl\
-Added Direct3D 12 Video interfaces\
-Added some missing Direct2D and Direct3D 11 interfaces\
-Added Windows Media Device Manager application interfaces (mswmdm.h, 50%- provider interfaces todo)\
-Added cert signing APIs from Mssign32.dll (mssign.h, 100%)\
-(Bug fix) GdipGetLineColors definition incorrect [(Issue #18)](https://github.com/fafalone/WinDevLib/issues/18)\
-(Bug fix) GdipDrawImagePointsRect[I] definitions incorrect for 64bit [(Issue #19)](https://github.com/fafalone/WinDevLib/issues/19)\
-(Bug fix) GdipEnumerateMetafileDestPoint[I] definitions incorrect for 64bit


**Update (v7.7.372, 09 April 2024):**\
-Minor additions to bring coverage of shellapi.h to 100%\
-Added macros/helpers from mfapi.h and mfplay.idl\
-Add missing gdip function GdipDrawImageFX\
-(Bug fix) GdipFillClosedCurve2[I] definitions incorrect. [(Issue #17)](https://github.com/fafalone/WinDevLib/issues/17)

**Update (v7.7.370, 05 April 2024):**\
-Added all Background Intelligent Transfer Service interfaces; 100% coverage of:\
 bits.idl, bits1_5.idl, bits2_0.idl, bits2_5.idl, bits3_0.idl, bits4_0.idl, bits5_0.idl, bits10_1.idl, bits10_2.idl, bits10_3.idl, bitscfg.idl, qmgr.idl.
 
**Update (v7.7.360, 04 April 2024):**\
-Very large expansion of DirectWrite interfaces; only dwrite.h was covered; added 100%\
 coverage of dwrite_1.h, dwrite_2.h, and dwrite_3.h\
-Added shdeprecated.h (100% coverage). Many of these are still in undocumented use.\
-UserEnv.h expanded to 100% coverage\
-Added crypto catalog APIs from mscat.h (100% coverage)\
-(API Standards) GetClassInfo[A, ExA, Ex] did not conform to API standards. For compatibility, this has been resolved by adding overloads.\
-CreateProfile does not have A/W variants. I have *zero* idea where I found otherwise, and with differently named arguments... no search results anywhere. Weird.\
-Add DWRITE_RENDERING_MODE missing values


**Update (v7.7.350, 31 Mar 2024):**\
-Large expansion of mfapi.h coverage; all APIs and GUIDs are covered, only missing the macros\
-processenv.h coverage now 100%\
-avrt.h 100% coverage in prep. for mfapi.h (limited current coverage)\
-Added 100% cover of netioapi.h\
-GetEnvironmentStrings now redirects to GetEnvironmentStringsW, per SDK.\
-Added security center interfaces from iwscapi.h and APIs from wscapi.h (both 100% covered)\
-Added WINDEVLIB_NOLIBS compiler option, completely disabling static library use (intended mainly to be able to test with tB Beta 423 or earlier)\
-(Bug fix) SetCurrentDirectory[W] definitions incorrect.\
-(Bug fix) Certain obscure PE header types missing alternate alignment attribute\
-(Bug fix) GetNamedPipeClientComputerName[A.W] definitions incorrect\
-(Bug fix) GetNamedPipeHandleState[A,W] definitions incorrect


**Update (v7.7.345, 26 Mar 2024):**\
-Added tdh.dll event trace helper APIs (tdh.h; all APIs/types complete but macros not yet added)\
-Added some additional native APIs.\
-FlushViewOfFile was missing.\
-(Bug fix) IMAGE_OPTIONAL_HEADER64 had an extra member and pointer member incorrectly declared as LongPtr, making the UDT offsets incorrect when handling a 64bit PE from a 32bit build.
-(Bug fix) The extra member mentioned above *is* in the 32bit version; so the build-linked version (IMAGE_OPTIONAL_HEADER) had to have a conditional added.

**Update (v7.7.343, 22 Mar 2024):**\
-(Bug fix) Coclass ActCtx conflicted with type ACTCTX; the former has been renamed CActCtx.\
-(Bug fix) ReleaseActCtx had typo in name.


**Update (v7.7.342, 21 Mar 2024):**\
-**MAJOR CHANGE:** The common used enum SHGNO_Flags has been renamed SHGDNF, the proper name per SDK.\
-**MAJOR CHANGE:** The common used enum SVGIO_Flags has been renamed SVGIO, the proper name per SDK.\
-**MAJOR CHANGE:** The common used enum SVSI_Flags has been renamed SVSIF, the proper name per SDK.\
-Updated WebView2 to match current stable release 1.0.2365.46\
-Filled out KUSER_SHARED_DATA more.\
-(Bug fix) NET_ADDRESS_INFO union substitute sized incorrectly.

**WinDevLibImpl v1.3.16:** Updatrd to use enum name changes associated with WinDevLib update.

**Update (v7.7.341, 16 Mar 2024):**\
-**MAJOR CHANGE:** The commonly used enum SFGAO_Flags has been renamed SFGAOF, in accordance with a previously overlooked official name for the enum: `typedef ULONG SFGAOF;` It is safe (as far as this package knows) to do a find/replace all for this. Also added missing value SFGAO_PLACEHOLDER.\
-For code portability, over the coming weeks and months I'll be replacing `DeclareWide` with `Declare`.\
  This will only be done on functions where it doesn't matter; where no arguments or arg UDT members\
  are `String`. It will still be used where it matters (especially in A/W functions without the A/W)\
-Added missing winmm video/animation consts and structs\
-Added helper function InitVariantFromIDList (undocumented inline helper)\
-Added interfaces IWebBrowserEventsService, IWebBrowserEventsUrlService (WebEvnts.idl, 100%)\
-Added interfaces ILaunchUIContext, ILaunchUIContextProvider\
-Added numerous shell related GUIDs\
-Added some missing property key related enums from propkey.h (should be 100% now)\
-Some enums for shell automation have officially associated IIDs; added these with new EnumId attrib\
-Added some missing registry constants and enum associations\
-Added SDK helper macros ISLBUTTON, ISMBUTTON, ISRBUTTON, ISDBLCLICK\
-EnumWindows, EnumChildWindows, and EnumTaskWindows APIs were inexplicably missing.\
-(API Standards) GetAltTabInfo did not conform to WinDevLib API standards (LongPtr instead of String)\
-(API Standards) GetKeyboardLayoutName did not conform to WinDevLib API standards (LongPtr instead of String)\
-(API Standards) ShutdownBlockReasonQuery was inconsistent with ShutdownBlockReasonCreate for String vs LongPtr.\
-(API Standards) CreateDesktop[A,ExA,Ex] did not use appropriate `DEVMODE[A,W]` variants.\
-(API Standards) RegCreateKey[A,W,ExA,ExW] did not use SECURITY_ATTRIBUTES instead of ByVal LongPtr.\
-(API Standards) RegConnectRegistry[A, ExA] did not use String types\
-(Bug fix) OpenDesktopA incorrectly used `DeclareWide`\
-(Bug fix) FOLDERTYPEID_ GUIDs were not properly defined as Static\
-(Bug fix) RegCreateKey, RegConnectRegistryExA definitions incorrect\
-(Bug fix) RegCreateKeyTransacted definition incorrect (wrong alias)\
-(Bug fix) Some winmm UDTs lacked required PackingAlignment attribute\
-(Bug fix) WAVEFORMAT[EX,EXTENSIBLE] lacked required PackingAlignment attribute


**Update (v7.6.334, 08 Mar 2024):**\
-Added 100% coverage of winsafer.h\
-Expanded power API coverage; powerbase.h, powersetting.h, powrprof.h 100%.

**Update (v7.6.332, 06 Mar 2024):**\
-NamespaceTreeControl default changed to INamespaceTreeControl2\
-Added inline helper SDK macros FreeIDListArray[Full|Child], SetContractDelegateWindow\
-(Bug fix) INameSpaceTreeControlEvents::OnGetTooltip should be ByVal pszTip\
-(Bug fix) MSGBOXPARAMS[A,W], MSGBOXDATA defs incorrect for x64.


**Update (v7.6.330, 04 Mar 2024):**\
-Added some additional sync APIs; synchapi.h coverage now 100%.\
-IObjectCollection now uses proper types (IUnknown and IObjectArray)\
-(Bug fix) IsBadStringPtr missing alias\
-(Bug fix) GetTimeZoneInformationForYear definition incorrect (used Long instead of Integer; no change needed, would work either way)\
-(Bug fix) HIMC/HIMCC types for IME APIs were incorrectly Long instead of LongPtr; this was only true on early Windows versions

**Update (v7.6.325, 29 Feb 2024):**
-Suppress new tB warnings (configd as errors in WinDevLib) for msvbvm60 DLL redirects (this info is still noted in the descriptions for each API)
-(Bug fix) DF_ALLOWOTHERACCOUNTHOOK value incorrect

**Update (v7.6.324, 27 Feb 2024):**\
-Added additional Variant/PROPVARIANT helpers; propvarutil.h now 100% covered\
-Additional DirectX As Any->proper type\
-Substantial improvement to Task Scheduler 2.0 interfaces (intellisense, Boolean instead of Integer where appropriate, descriptions)\
-(Bug fix) InitVariantFromString was not a dll export (replaced by macro)\
-(Bug fix) VariantToFileTimeArray and VariantToFileTimeArrayAlloc don't exist\
-(Bug fix) IScheduledWorkItem missing 3 methods and GetRunTimes, SetCreator methods incorrect.\
-(Bug fix) ITaskSettings missing Compatibility Let/Get methods.\
-(Bug fix) ITaskSettings3 missing CreateMaintenanceSettings method\
-(Name change) ISchedulingAgent was apparently renamed ITaskScheduler by Windows 2000; coclass SchedulingAgent to CTaskScheduler.\
               Further, IEnumWorkItems was IEnumTasks before that; why olelib was inconsistent here, I don't know.\
               Since the SDK still defines these as aliases, WinDevLib now includes both names for all 3.\
-(Name change) TASK_RUNLEVEL corrected to more appropriate TASK_RUNLEVEL_TYPE


**Update (v7.6.322, 24 Feb 2024):**\
-Added DSA and DPA APIs (dpa_dsa.h, 100% coverage including macros)\
-Further compat updates for The trick's typelibs:\
   -IDWriteFontFileLoader.CreateStreamFromKey last arg now retval.\
   -ID2D1RenderTarget many arguments now optional, with correct default values where appropriate\
   -IWICBitmap.Lock last arg now retval\
-ID2D1Factory and ID2D1Geometry had many As Any arguments switched to their proper types\
-Added SizeToLongLong helper function\
-(Bug fix) PointFToLongLong helper function incorrect.\
-(Bug fix) ID2D1RenderTarget::CreateBitmap definition incompatible with 64bit

**Update: WinDevLibImpl v1.3.15** - `IPersistStream::GetMaxSize now LongLong instead of Currency, matching WinDevLib.


**Update (v7.6.320, 20 Feb 2024):**\
-Added IPrintDocumentPackage* interfaces and coclasses (DocumentTarget.idl, 100%)\
-Added un/under-documented MRU APIs from comctl32\
-For compatibility with The trick's D2D and WIC typelibs:\
   -D2D1_MATRIX_ types are now flat; the D2D alias versions remain the same, switch to these if you were using the previous defs.\
   -ID2D1Effect data arguments are now As Any (no change needed)\
   -Some arguments now optional (no change needed)\
      NOTE: Unlike VB6, twinBASIC supports ByVal Nothing to pass a null pointer to a ByRef interface/object method.\
   -ID2D1DeviceContext::CreateEffect last param now return value\
   -IWICBitmapDecoder::GetFrame last param now return value\
-Many Direct2D/DirectWrite types were changed from As Any to their real UDT, since tB supports vbNullPtr to pass the optional null.\
   While this reduces compatibility with The trick's TLBs (and oleexp), the extra info and intellisense benefits are worth it.\
-(Bug fix) PathRemoveBackslashW incorrectly used String.\
-(Bug fix) LookupPrivilegeValue[A] used LongPtr instead of String.\
-(Bug fix) PointToLongLong ambiguous overloads; new PointFToLongLong for POINTF.\
-(Bug fix) All Direct2D effects CLSID functions incorrect (returning UUID_NULL)\
-(Bug fix) IDWriteLocalizedStrings, IDWriteTextFormat, IDWriteTextLayout, IDWriteLocalFontFileLoader string arguments improperly ByRef\
-(Bug fix) IDWriteInlineObject, IDWriteTextRenderer, and IDWritePixelSnapping argument clientDrawingContext should be ByVal LongPtr.\
-(Bug fix) Several DirectWrite font UDTs had plocalename members incorrectly defined as Long, making them incompatible with 64bit


**Update (v7.6.312, 10 Feb 2024):**\
-Added IAccessControl/IAuditControl interfaces\
-Added numerous missing propsys APIs; propsys.h coverage now 100%\
-Added a few missing registry functions, also previously excluded deprecated ones-- winreg.h coverage is now 100%.\
-`GetProcessMemoryInfo` now uses As Any so `PROCESS_MEMORY_COUNTERS` and `PROCESS_MEMORY_COUNTERS_EX2` can also be used.\
-Added System Restore APIs from SrRestorePtApi.h (100%). IMPORTANT: Event types have been prefixed with SRPT_ due to common name conflicts (e.g. it has `BACKUP, RESTORE`, etc, that are now `SRPT_BACKUP, SRPT_RESTORE`, etc)\
-Added Compressor APIs from compressapi.h (100%). IMPORTANT: Compress and Decompress have been renamed CompressorCompress and CompressorDecompress, respectively, due to the short name conflict potential.\
-(Internal) Moved crypto APIs to their own file, wdAPICrypto.twin. Internet APIs moved to new module wdAPIInternet with wdInternet.twin. DEVPKEY and MiscGUID regions moved to wdDefs.twin. wdAPI.twin was becoming unmanageable and running into performance issues; it was up to 65k lines before this reorganization.\
-Implemented all basic Interlocked* APIs. These are implemented primarily as static libraries: Only a few of these are exported by the Windows API, and only on x86.\
 To handle this, I've included my Interlocked64 project as a static library. I've also produced a 32bit version to handle all the inline/instrinsic ones besides the basics.\
 If you wish to avoid static linking these obj files (while using the APIs), specify the compiler flag:\
 `#WINDEVLIB_AVOID_INTRINSICS`\
 This uses the kernel32 versions *where available*: You're limited to InterlockedIncrement, InterlockedDecrement, InterlockedExchange[Add], and InterlockedCompareExchange[64].
 Using any besides those 6 will trigger the static library to be included.
 NOTE: TEMPORARY: Due to editing instability, a default alternative of ONLY the kernel32s are set-- for use in Beta 423. See wdInterlocked.twin.
-Added addtional error codes
-Added cards.dll APIs for 32bit only (no 64bit build exists)

**Update (v7.5.310, 26 Jan 2024):**\
-Massive expansion of crypt APIs; coverage of wincrypt.h, dpapi.h (crypto data protection) and mssip.h now 100%\
-Coverage of wintrust.h is now 99%; all but a couple of difficult to decipher macros and a byte sequence the order needs to be verified for.\
-Coverage of memoryapi.h is now 100% (excluding APIs only available to Store Apps)\
-Added UserNotification2 coclass; oleexp had this with a default of IUserNotification2, and while WinDevLib had UserNotification as a coclass, it had IUserNotification as a default without listing 2. Added 2 and the additional coclass.\
-EVENT_FILTER_EVENT_ID is now buffered to the maximum number of IDs. This allows using it directly, at the expense of not being able to use LenB for size.\
-Virtual* memory functions now use ByVal addresses instead of ByRef As Any; 99% of code uses this definition.\
-(Bug fix) CertFreeCertificateContext definition incompatible with x64\
-(Bug fix) SwapVTableEntry helper not working with old defs

**Update (v7.4.308, 20 Jan 2024):**\
-Added interface IAttachmentExecute and coclass AttachmentServices.\
-Added interface IStorageProviderBanners, and coclass StorageProviderBanners.\
-Substantial expanson of crypto APIs; bcrypt.h, ncrypt.h, and ncryptprotect.h all now have 100% coverage, and wincrypt.h coverage has doubled (though still has quite a bit to go)\
-Crypto provider enum Crypt_Providers (dwProvType) renamed to CryptProviders to resolve conflict with SDK-defined CRYPT_PROVIDERS type.\
-Numerous missing IShellMenu related consts/types; fixed incorrect intellisense associations.\
-(Bug fix) MEMORYSTATUS definition incorrect (incompatible with 64bit). The associated API should not be used however, as it has problems with >4GB RAM. Use GlobalMemoryStatusEx.

**Update (v7.3.306, 17 Jan 2024):**\
-Some additional crypto APIs.\
-Added undocumented TaskDialogIndirect button flags (Abort, Ignore, Continue, Retry, Help) and renamed the enum to the proper SDK-defined name (replace TDBUTTONS with TASKDIALOG_COMMON_BUTTON_FLAGS)\
-Added x,y option to PointToLongLong helper.\
-Added some missing GDI defs and macros.\
-(Bug fix) Numerous duplicated enum values undetected last time.

**Update (v7.3.304, 15 Jan 2024):**\
-Added legacy Sync Manager interfaces/coclasses (mobsync.h, 100%)\
-Added process snapshot APIs (ProcessSnapshot.h, 100% coverage)\
-Added all consts (grouped as enums where possible) from propkey.h\
-Added new property keys from propkey.h\
-Added some missing STR_ binding strings.\
-Small additions to get shellapi.h coverage to 100%\
-Added undocumented interfaces IInfoBarMessage, IInfoBarHost, and IBrowserProgressSessionProvider (for the popup banner menus in NSEs)\
-Added undocumented interfaces IShellFolder3, IFilterItem, IItemFilter\
-Added undocumented interfaces IScope, IScopeItem (NSE filtering)\
-(Bug fix) LockWorkStation incorrect case.\
-(Bug fix) SHFILEOPSTRUCT[A,W] definition incorrect for x86

**Update (v7.2.301, 10 Jan 2024):** Bug fix: Numerous duplicated enum values.

**Update (v7.2.300, 09 Jan 2024):**\
-Added wincred advapi32.dll APIs; wincred.h, 100% coverage\
-Completed adding WinHttp APIs, winhttp.h coverage now 100% (note: The WinHttp interface/coclass is not included as it already has a VB/tB-compatible typelib to add)\
-Added remaining websocket.dll APIs, websocket.h coverage now 100%\
-Added pointer encode/decodes functions (and kernel32's Beep): utilapiset.h 100% coverage\
-A few missing WinInet APIs\
-Around 100 additional HRESULT error constants w/ descriptions.\
-Base WinRT IInspectable and some initialization APIs and HSTRING APIs added.\
-(Bug fix) All ERROR_DS_x constants were wrong. ICM ERROR_x constants were wrong.


**Update (v7.2.289, 06 Jan 2024):** Bug fix: InternetConnect definition incorrect.\
**Update (v7.2.288, 06 Jan 2024):**\
-Added Photo Acquisition interfaces and coclasses (photoacquire.h, 100%)\
-Added accessibility APIs from oleacc.dll (oleacc.h now 100% coverage). Really thought these were already added; there's a bug in oleexp where most are missing from that too despite presence in source.\
-Added inline Library helper functions from ShObjIdl_core.h; also some additional shell32.dll APIs.\
-Added SDDL language string constants; coverage of sddl.h now 100%.\
-Additional advapi32.dll security APIs, to bring coverage of securitybaseapi.h to 100%.\
-Added 100% coverage of dssec.h.\
-Cleaned up PROCESS_BASIC_INFORMATION\
-(Bug fix) LogonUserEx[A,W] definitions incorrect.\
-(Bug fix) CreateWellKnownSid definition incorrect.\
-(Bug fix) GetSidIdentifierAuthority definition likely incorrect.\
-(Bug fix) SHChangeUpdateImageIDList missing 1-byte packing attribute.\
-(Bug fix) A couple setup APIs missing 32bit 1-byte packing attribute.


**Update (v7.1.286, 02 Jan 2024):**\
-Added initial coverage of Lsa* APIs from advapi32.dll/NTSecAPI.h/LSALookup.h/ntlsa.h\
-WIC: Converted LongPtr buffer arguments to As Any, for more flexibility in what can be supplied.\
-WIC: Converted all ByVal VarPtr(WICRect) LongPtr's to ByRef WICRect.\
-(Bug fix) IWICBitmapSourceTransform::CopyPixels definition incorrect.\
-(WinDevLibImpl) Added Implements-compatible WIC interfaces for custom codec creation.

**Update (v7.0.283, 01 Jan 2024):**\
-Improved enum associations/formatting for WIC.\
-Added numerous missing GUIDs from wincodecsdk.h\
-(Bug fix) IWICPalette, IWICFormatConverter, IWICBitmapDecoderInfo, IWICPixelFormatInfo2, IWICMetadataReaderInfo, IWICMetadataHandlerInfo, IWICBitmapCodecInfo, IWICComponentInfo, WICMapGuidToShortName, WICMapSchemaToName had numerous ByVal/ByRef mixups.

**Update (v7.0.282, 01 Jan 2024):**\
-Added all variable conversion and arithmetic helpers from oleauto.h; coverage of that header now 100% (of supported by language).\
-Additional GUIDs and error consts from olectl.h to bring that header's coverage to 100%.\
-VARCMP enum renamed VARCMPRES to avoid conflict with VarCmp API.\
-Added missing flags for VariantChangeType[Ex]\
-SHFileOperation and SHFILEOPSTRUCT did not conform to API standards. Struct names were incorrect; the operations aborted member was incorrectly defined as Boolean, but the padding bytes prevented it from failing the entire function.\
-SysAllocStringByteLen now use ByVal As Any, since either a String or LongPtr would be ByVal.\
-(Bug fix) SysAllocString definition incorrect (Long instead of LongPtr, impacting 64bit)\
-(Bug fix) SysFreeString definition incorrect (ByRef instead of ByVal)\
-(Bug fix) SysReAllocStringLen should use DeclareWide\
-(Bug fix) LHashValOfName is a macro, not an export; now implemented properly.\
-(Bug fix) FORMATETC used a Long for CLIPFORMAT, which is incorrect.\
-(MAJOR BUG FIX) IStream was missing UnlockRegion. This impacted numerous derived interfaces, throwing off their vtables, completely breaking them. This bug was also present in WinDevLibImpl.


**Update (v7.0.280, 28 Dec 2023):**\
-INDEXTOOVERLAYMASK was inexplicably missing; also added inverse, OVERLAYMASKTOINDEX.\
-Additional setup APIs-- newdev.h, 100% coverage, and additional cfgmgr32 APIs.\
-Additional kernel32 APIs-- processthreadsapi.h now has 100% coverage\
-(Bug fix) SetupDiGetClassDevsW did not conform to WinDevLib API standards.\
-(Bug fix) Some SetupAPI defs did not have the required 1-byte packing on 32bit\
-(Bug fix) NMLVKEYDOWN and NMTVKEYDOWN did not have required packing alignment

**Update (v7.0.277, 21 Dec 2023):**\
-Added customer caller for AuthzReportSecurityEvent (experimental).\
-(Bug fix) SHEmptyRecycleBinW, PathRemoveBackslash, PathSkipRoot, CreateMailslot did not conform to API standards\
-(Bug fix) All SHReg* APIs missing W variants\
-(Bug fix) PathAddExtension, PathAddRoot, EnumSystemLanguageGroups, LoadCursorFromFile, waveInGetErrorText definitions incorrect (misplaced alias)\
-(Bug fix) PathIsDirectoryA/W, PdhAddEnglishCounterA definitions incorrect (invalid alias)\
-(Bug fix) GetLogicalDriveStringsA definition incorrect (DeclareWide on ANSI)\
-(Bug fix) Mising DeclareWide:\
    Get/SetComputerName[Ex]\
    All THelp32.h APIs\
    SHUpdateImage\
    ShellNotify_Icon\
    WaveIn/OutDevCaps\
    HttpQueryInfo


**Update (v7.0.276, 20 Dec 2023):**\
-Added cryptui.dll APIs (cryptuiapi.h, 100% coverage)\
-Some additional SetupAPI and Cfgmgr32 defs, as well as devmgr.dll APIs documented and not (show device manager, prop pages, problem wizard, etc)\
-More inexplicably missing shell32 APIs\
-Additional APIs from ShellScalingAPI.h (now 100% coverage)\
-(Bug fix) Duplicated DEVPROP_TYPE_* values.\
-(Bug fix) GetExplicitEntriesFromAcl definition incorrect (misplaced Alias)\
-(Bug fix) Wow64RevertWow64FsRedirection lacked explicit ByVal modifier.\
-(Bug fix) Get/SetProcessDpiAwareness definitions incorrect.

**Update (v7.0.272, 17 Dec 2023):**

***tbShellLib is now WinDevLib - Windows Development Library for twinBASIC***




***MAJOR CHANGES***
*`LARGE_INTEGER`*
I've been considering these for a long time, and decided to pull the trigger before tB goes 1.0. 

The LARGE_INTEGER type is defined  in C as:

```c
typedef union _LARGE_INTEGER {
    struct {
        DWORD LowPart;
        LONG HighPart;
    } DUMMYSTRUCTNAME;
    struct {
        DWORD LowPart;
        LONG HighPart;
    } u;
    LONGLONG QuadPart;
} LARGE_INTEGER;
```

The Windows API, from user to native to kernel, all recognize the QuadPart member and apply 8-byte packing rules.
VB6 and VBA (except 64bit) lack a LongLong type, so programmers have traditionally used the LowPart/HighPart option.
This *does not* trigger 8 byte packing rules, and while problems from this are rare in 32bit mode, they're quite common
in 64bit mode. As a result of this, WinDevLib has up until now kept the traditional definition for LARGE_INTEGER and 
instead substituted a QLARGE_INTEGER or ULARGE_INTEGER in it's own definitions.\
This will now change. The original plan was to wait for union support which would allow both while still triggering
the 8 byte alignment rules, but that has recently been confirmed as a post-1.0 feature. When that is added, the old
option will be added back in.\
LARGE_INTEGER now by default uses QuadPart, and all QLARGE_INTEGER have been changed to LARGE_INTEGER.

Reminder: This does greatly simplify things; you can remove all conversions to Currency and related multiply/divide 
          by 10,000. Also, note that if you use your own local definition, WinDevLib does not supercede it for your
          own code. It is strongly recommended to switch away from Currency when doing 64bit updates.

A compiler flag is available to restore the old definition (but not the use of QLARGE_INTEGER in WinDevLib defs):
`WINDEVLIB_NOQUADLI`

*SendMessage and PostMessage*
These will now conform to the same API standards as all other functions; the undenominated (without A or W suffix)
will now point to `SendMessageW` and `PostMessageW` and use `DeclareWide`. Note that these have never affected the target
itself, it's always just modified how String arguments are interpreted. 99% of usage of these will not be impacted
by this, since you'll still be able to use String and not nee to modify the result for ANSI/Unicode conversion.
PostMessage already used `DeclareWide`, which was perhaps causing unexpected issues in the edge cases.

**Addtional changes:**\
-Added interface IActCtx and coclass ActCtx.\
-Missing WH_ enum values and associated types for SetWindowsHookEx\
-Numerous missing VK_* virtual key codes\
-Missing WM_* wParam enums.\
-Several service APIs did not conform to tbShellLib API standards with respect to A/W/DeclareWide UDT naming.\
-Added a lot of additional user32 content.\
-Added variable min/max constants from limits.h (100% coverage)\
-Redid FILEDESCRIPTOR[A,W] to use proper FILETIME types and Integer for WCHAR instead of 2x Byte.\
-Added several types associated with clipboard formats.\
-Added unsigned variable helper functions (thanks to Krool for these): UnsignedAdd, CUIntToInt, CIntToUInt, CULngToLng, and CLngToULng. CULngToLng has an override between the original Double and LongLong, CLngToULng does too but rewrites the output into an argument since tB can't overload purely based on function return type.\
-Added gesture angle macros GID_ROTATE_ANGLE_TO_ARGUMENT/GID_ROTATE_ANGLE_FROM_ARGUMENT\
-Added hundreds of additional NTSTATUS values.\
-Added overloads to LOWORD and HIWORD to handle LongLong directly.\
-winuser.h now has 100% coverage of language-supported definitions (10.0.25309 SDK); the largest header to date with this distinction with over 16000 lines in the original.\
-(Bug fix) LBItemFromPt was marked Private.\
-(Bug fix) RealGetWindowClass definition incorrect (invalid alias).\
-(Bug fix) Duplicated constant: CCHILDREN_SCROLLBAR\
-(Bug fix) PostThreadMessage definition incorrect and did not meet API standards.\
-(Bug fix) InsertMenuItem[A,W] definitions technically incorrect although not causing an error. Also did not conform to API standards.\
-(Bug fix) PostThreadMessage definition incorrect.\
-(Bug fix) PostMessageA incorrectly had DeclareWide.\
-(Bug fix) ILCreateFromPathEx was removed as it's not exported from shell32 either by name or ordinal.\
-(Bug fix) ILCloneChild, ILCloneFull, ILIsAligned, ILIsChild, ILIsEmpty, ILNext, and ILSkip are only macros; they were declared as shell32.dll functions. Some of these were aliases and modified appropriate, the rest were implemented as functions.\
-(Bug fix) ILLoadFromStream is exported by ordinal only.

**WinDevLibImpl v1.2.11.272**
-(Bug fix, WinDevLibImpl) IPersistFile method definition incorrect.



---

For earlier version history, see CHANGELOG.md

For more information and a list of available interfaces, visit the [VB Forums thread](https://www.vbforums.com/showthread.php?897883-twinBASIC-tbShellLib-Shell-Interface-Library-(x64-compatible-successor-to-oleexp)) for this project.
