# WinDevLib 
## Windows Development Library for twinBASIC

### ***tbShellLib is now WinDevLib - Windows Development Library for twinBASIC***
This project has grown well beyond it's original mission of shell programming. While that's still the largest single part, it's no longer a majority of the code, and the name change now much better reflects the purpose of providing a general Windows API experience like windows.h. Compiler constants and module names/file names have been updated to reflect the name change. tbShellLibImpl is now WinDevLibImpl. There are also some major chanages associated with this update, please see the full changelog below.

### DLL Redirection Errors 
 
twinBASIC now counts msvbvm60 redirects as legacy DLL redirects, which WinDevLib set to "Error". Please update to the latest version of WinDevLib to get rid of these errors and use it on twinBASIC Beta 456 and newer. Both this repo and the package server downloads have been updated.
 

**Current Version: 7.7.342 (March 21st, 2024)**

(c) 2022-2023 Jon Johnson (fafalone)

This project is a comprehensive twinBASIC replacement for [oleexp.tlb](http://www.vbforums.com/showthread.php?786079-VB6-Modern-Shell-Interface-Type-Library-oleexp-tlb), my Modern Shell Interfaces Type Library project for VB6, that is x64 compatible, due to the many problems using midl to create a 64bit tlb.

This and oleexp are projects to supply Windows shell and component interfaces in a format consumable by VB6/VBA/tB. This involves not only defining interfaces, but using VB/tB compatible types-- so in some cases, even though there may be an existing way to import references to interfaces, they may be unusable due to e.g. the use of unsigned types, C-style arrays, double pointers, etc. All interfaces, types, consts, and APIs from oleexp are covered. For a full list of interfaces, see [INTERFACES.md](https://github.com/fafalone/WinDevLib/blob/main/INTERFACES.md).

In addition to the 2200+ common COM interfaces, WinDevLib now includes expansive coverage of Windows APIs from all the common modules. This makes it similar to working in C++ with `#include <Windows.h>` and a few others. Currently, approximately 5,500 of the most common APIs have been added- redone by hand from the original headers, in order to restore 64bit type info lost in VB6 versions, avoid the errors of automated conversion tools (e.g. Win32API_PtrSafe.txt is riddled with errors), and make them friendlier by converting groups of constants associated with a variable into an Enum so it comes up in Intellisense. This takes advantage of tB's ability to provide Intellisense for types besides Long in API defs (hopefully UDTs soon, this project has provisioning for that). 

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


### Guide to switching from oleexp.tlb

WinDevLib presented the best opportunity there would be to ditch some olelib legacy baggage. It's fairly simple to move your VB6 projects to WinDevLib, just follow these steps:

1) Replace public aliases: It's important to do this first. Run a Replace All changing oleexp.LONG_PTR to LongPtr, oleexp.REFERENCE_TIME to LongLong, oleexp.HNSTIME to LongLong, oleexp.KNOWNFOLDERID to UUID, oleexp.EventRegistrationToken to LongLong, oleexp.BINDPTR to LongPtr, and oleexp.LPCRITICAL_SECTION to LongPtr. If you've used them without the oleexp. prefix, you'll also need to replace those, but if you've imported into tB they should be tagged.

2) Replace oleexp.IUnknown with IUnknownUnrestricted. WinDevLib keeps this separate due to the major issues with conflicts with the former approach. If your project has IUnknown *without* oleexp. in front of it, **do not** replace those, as it's not referring to oleexp. 

3) After you've done those two, you can now go ahead and simply delete all remaining instances of `oleexp.` (including the .). 

4) Convert `Currency` to `LongLong` for interfaces and APIs: It's no longer neccessary to worry about multiplying and dividing by 10,000 since tB supports a true 64bit type in both 32bit and 64bit mode. So this change is ultimately for the better, but existing codebases will have had to have used `Currency` for all interfaces and oleexp APIs expecting a 64bit integer.

5) Manually address any errors remaining. Interfaces should be mostly fine at this point, but if you've made use of the APIs in oleexp, many of them have syntax differences, mainly not being able to rewrite an ending [out] argument as the return value, and changing String arguments to `LongPtr` you'll need `StrPtr` with. Another major difference is that the default for almost all APIs with ANSI/Unicode (A/W) versions, is now the Unicode version. In most cases, the W version is declared with `LongPtr` for strings, and the untagged alias version uses tB's new `DeclareWide` keyword to disable ANSI conversion while using `String`.\
Finally, a very small number of APIs and interfaces use ByVal UDTs. Since VB cannot do this, nor can tB yet, a typical workaround was to pass each member as an individual argument. This worked when arguments were 4 bytes each, but the x64 calling convention aligns arguments at 8 bytes. So the two options were to follow that convention, which also works for 32bit allowing a single call for both, or require two different calls for 32 and 64bit. Since one of the main points of twinBASIC is 64bit support, WinDevLib uses the former option. The downside of this is that VB-style calls will have to be rewritten. If you see, for example, `ByVal ptX As Long, ByVal ptY As Long` replaced with `ByVal pt As LongLong`, this was an unsupported `ByVal POINT`. You'd declare a LongLong, and use `PointToLongLong` (there's also PointFToLongLong for Single-based points), helper functions added to help here.

> [!NOTE]
>  This is just for using WinDevLib-- you'll likely have a lot more changes to make if you want to make your project x64 compatible.


#### Handling UDTs where normally they were As Any

>[!NOTE]
>If you see errors like `Validation of call to 'CreateFile' failed.  Argument for 'lpSecurityAttributes': cannot coerce type 'Long' to 'SECURITY_ATTRIBUTES'`, this section explains the cause and solution!

The best example of this is many APIs, like file APIs, where in traditional VB declarations, you see 'As Any' and in WinDevLib you see e.g. `SECURITY_ATTRIBUTES` or `OVERLAPPED`. These are the correct the definitions, but VB6 had no facility to specify 'NULL', which is what they usually would be set to as optional arguments. So the VB6 way was a workaround, where you could pass `ByVal 0`. 

twinBASIC has direct support for passing a null pointer instead of a UDT. You can pass `vbNullPtr` to these arguments where previously you would have used ByVal 0 on an `As Any` argument that you've found is now a UDT. You can also pass a non-null pointer; simply pass a `LongPtr` *without* `ByVal` (for now, twinBASIC will be changing this to require `ByVal` as that makes it far more clear you intend this kind of substitution and doesn't imply you're passing ByRef LongPtr). 

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


### Guide to switching from oleexpimp.tlb

There's 'WinDevLib for Implements' (WinDevLibImpl.twinpack/.twinproj) as well, but you'll note it has substantially fewer interfaces than oleexpimp.tlb. This is because there's two reasons for an interface to have an alternate version: It uses `[ Preservesig ]` on one or more methods, or it uses `As Any`. twinBASIC allows using `Implements` with `As Any` by replacing it with `As LongPtr` (which is what the alternate versions do). So many interfaces were in oleexpimp.tlb for this latter reason, and subsequently are *not* included in WinDevLibImpl as it's not neccessary.

If you find an oleexpimp.tlb interface is not in WinDevLibImpl, you will be able to use the one from WinDevLib, simply make sure `As Any` is changed to `As LongPtr`. 

tB has announced plans to support `[ PreserveSig ]` in implemented interfaces in the future; when that happens WinDevLibImpl will be deprecated.

>[!IMPORTANT]
>There currently seems to be an issue with using WinDevLib and WinDevLibImpl together if WinDevLibImpl does not use the current WinDevLib as a reference (it would usually use an old one as it's updated much less frequently). I've updated the reference on this repo and the package server, just note that you'll need to refresh both every time you update one if they're used together.


### WinDevLib API standards

This was mentioned above, but it's worth going into more detail. In addition to the COM interfaces, WinDevLib has a large selection of common Windows APIs; this is a much larger set than oleexp. WinDevLib and twinBASIC represented the best opportunity there would be to modernize standards... most VB programs are written with ANSI versions of APIs being the default. **This is not the case with WinDevLib**. With very few exceptions, APIs are Unicode by default-- i.e. they use the W, rather than A, version of APIs e.g. `DeleteFile` maps to `DeleteFileW` rather than `DeleteFileA`. The A and W variants use String/LongPtr, and in almost all cases, the mapped version uses `String` with twinBASIC's `DeclareWide` keyword-- this disables Unicode-ANSI conversion, so you can still use `String` without `StrPtr` or any Unicode <-> ANSI conversion. Note this usually only applies to strings passed as input, APIs passing a LPWSTR that's allocated externally will still be LongPtr, as they're not in the same BSTR format as VBx/TB strings.

All APIs are provided, as a minimum, as the explicit W variant, and an untagged version that maps to the W version. Some, but not all, APIs also have an explicit A variant defined that will perform the normal ANSI conversion for compatibility purposes. This is decided on a case by case basis depending on my impression of how much legacy code is around that needs the ANSI version. All new code should use the Unicode versions.

UDTs used by these calls are also supplied in the same manner, the W variant, an untagged variant that's the same as the W version, and in some cases, an A version. UDTs always use `LongPtr` for strings, even the untagged versions for `DeclareWide`. 

If you have any doubts about which API is being called, twinBASIC will show the full declaration when you hover your cursor over the API in your code.

#### Scope of coverage

The goal of the API coverage in WinDevLib is to provide the kind of programming experience you'd get in C/C++ by including windows.h and some of the more common feature sets like DirectX and GDIPlus. It currently includes about 5,500 APIs. But even that is just scratching the surface of the total Windows API set. Due to the low quality of automated conversion, even by Microsoft themselves (see: Win32API_PtrSafe.txt), I'm not interested in simply feeding headers through a conversion utility or using a database, so instead WinDevLib will be focused on the most commonly used features in the major system DLLs, though less common ones can be added by request or as time goes on and the existing DLLs are completed. I do not intend to include native APIs that have fully equivalent regular APIs; that's basically doubling the work for no benefit-- but if they offer additional features or substantially improved performance, they will be included. 

I've included the definitions, associated types, and associated constants, for extensive portions of the following modules: shell32.dll, shlwapi.dll, kernel32.dll, user32.dll, advapi32.dll, ole32.dll, oleaut32.dll, propsys.dll, gdi32.dll, gdiplus.dll, urlmon.dll, hlink.dll, winmm.dll, setupapi.dll, comctl32.dll, dwm.dll/uxtheme.dll, comdlg32.dll, winspool.drv, and netapi32.dll. Besides highly self-contained specialized sets in their own headers (unless small), please let me know any I've missed from these.\
Limited coverage (or full coverage of very small sets) is provided for ntdll.dll, version.dll, msimg32.dll, crypt32.dll, virtdisk.dll, sxs.dll, secur32.dll, imm32.dll, userenv.dll, wintrust.dll, msacm32.dll, url.dll, htmlhelp.dll, imagehlp.dll, and ws2_32.dll. If you feel any missing ones from these should be included, or would like to contribute more, let me know.\
Finally, there's small API sets for features, like DirectX DLLs, Webview2Loader, WIC, etc. Definitely let me know any missing from these.

**Future coverage:** In the future I'm planning to expand native APIs with no equivalents, add additional Winsock coverage, and add OpenGL-- though for this last one I may wait for tB to have `Alias` support since existing OpenGL codebases make heavy use of them by way of NeHe's TLB. I welcome contributions of any of these. If you've done the consts->enums conversions already, I'd even take 32bit-only versions.


### Updates

**Update (v7.7.342, 21 Mar 2024):**\
-**MAJOR CHANGE:** The common used enum SHGNO_Flags has been renamed SHGDNF, the proper name per SDK.\
-**MAJOR CHANGE:** The common used enum SVGIO_Flags has been renamed SVGIO, the proper name per SDK.\
-**MAJOR CHANGE:** The common used enum SVSI_Flags has been renamed SVSIF, the proper name per SDK.\
-Updated WebView2 to match current stable release 1.0.2365.46\
-Filled out KUSER_SHARED_DATA more.\
-(Bug fix) NET_ADDRESS_INFO union substitute sized incorrectly.

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


**Update (v6.6.269):**\
-Added helper function GetNtErrorString that gets strings for NTSTATUS values. GetSystemErrorString already exists for HRESULT.\
-SHLimitInputEdit didn't have the ByVal attribute included, making it easy to not realize it's then required when called.\
-CreateSymbolicLink API inexplicable missing.\
-LIMITINPUTSTRUCT has been renamed to the original, correct name LIMITINPUT. The original documentation and demos have made this change too with the recently released universal compatibility update.


**Update (v6.6.268, 11 Dec 2023):**\
-Added UI Animation interfaces and coclasses\
-Added Radio Manager interfaces and some undocumented coclasses to use them. Added undocumented interface IRadioManager with coclass RadioManagementAPI: This controls 'Airplane mode' on newer Windows.\
-Added IThumbnailStreamCache and coclass ThumbnailStreamCache. Note: Due to simple name potential conflicts, flags prefixed with TSC_. A ByVal SIZE is replaced with ByVal LongLong; copy into one.\
-Added additional event trace APIs; coverage of evntrace.h is now 100%.\
-Additional BCrypt APIs sufficient for basic public key crypto implementations.\
-Added additional language settings APIs from WinNls.h; coverage is near or at 100% now.\
-Added remaining transaction manager APIs; coverage of ktmw32.h is now 100%.\
-Added all remaining .ini/win.ini file APIs.\
-Added misc other APIs.\
-Added memcpy alias for RtlMoveMemory (in addition to CopyMemory and MoveMemory)\
-Several event trace APIs and transaction API improperly used 'As GUID', which is undefined in tbShellLib and will refer to the unsupported stdole GUID.\
-Reworked the way the REASON_CONTEXT union was set up; the old version would likely not work as implied.\
-(Bug fix) KSIDENTIFIER union size incorrect.

**Update (v6.5.263, 06 Dec 2023):**\
-Added numerous missing shell32 APIs.\
-Some additional kernel32 APIs, bringing coverage of fileapi.h to 100%.\
-Added numerous IOCTL_DISK_* constants and associated UDTs.\
-Converted some ListView-related consts to enums to use with their associated UDTs.\
-Added missing name mappings structs for SHFileOperation.\
-(Bug fix) BITMAPFILEHEADER, DISK_EXTENT, VOLUME_DISK_EXTENT, and STORAGE_PROPERTY_QUERY typed improperly marked Private.\
-(Bug fix) STORAGE_PROPERTY_QUERY definition incorrect\
-(Bug fix) SCSI_PASS_THROUGH_BUFFERED24 definition incorrect.\
-(Bug fix) GetVolumeInformationByHandle definition incorrect.\
-(Bug fix) ReadFile did not conform to tbShellLib API conventions (ByVal As Any instead of OVERLAPPED)

**Update (v6.5.260, 04 Dec 2023):**
-Added all authz APIs/consts/types from authz.h; note that AuthzReportSecurityEvent is currently unsupported by the language. However, it internally calls AuthzReportSecurityEventFromParams.\
-Added many missing shlwapi APIs; URL flags enum missing values\
-Updated shlwapi "Is" functions to use BOOL instead of Long where that way in sdk.\
-Completed all currently known PROCESSINFOCLASS structs for NtQueryInformationProcess.\
-Added custom enums for PROCESS_MITIGATION_* structs\
-(Bug fix) SHGetThreadRef/SHSetThreadRef definitions incorrect\
-(Bug fix) SHMessageBoxCheck definition incorrect\
-(Bug fix) Path[Un]QuoteSpaces definitions incorrect

**Update (v6.4.258), 28 Nov 2023):**\
-Large number of additional advapi security APIs (AccCtrl.h and AclAPI.h, 100% coverage)\
-Additional crypto APIs\
-(Bug fix) Missing FindFirstFileEx flag FIND_FIRST_EX_ON_DISK_ENTRIES_ONLY.

**Update (v6.4.257), 26 Nov 2023):** GdipGetImageEncoders/GdipGetImageDecoders definitions "incorrect" for unclear reasons... Documentation indicates it's an array of ImageCodecInfo, which does not contain any C-style arrays, but there's a mismatch between the byte size and number of structs * sizeof. Changed to As Any to allow byte buffers in addition to oversized ImageCodecInfo buffers.

**Update (v6.4.256, 25 Nov 2023):**\
-Added inexplicably missing basic versioning and sysinfo APIs from kernel32.\
-Added ListView subitem control undocumented CLSIDs.\
-Additional sys info classes (NtQuerySystemInformation).\
-Misc. API additions.\
-(Bug fix) GetAtomName[A,W] and GlobalGetAtomName[A,W] definitions incorrect.\
-(Bug fix) Multiple ole32 functions incorrectly passing ANSI strings.\
-(Bug fix) ListView_GetItemText was thoroughly broken.\
-(Bug fix) GetSystemDirectory definition incorrect.\
-(Bug fix) EnumPrintersA definition incorrect; GetPrinter, SetPrinter, and GetJob definitions technically incorrect but no impact unless you had redefined associated UDTs.\
-(Bug fix) UNICODE_STRING members renamed to their proper SDK names. I realize this is a substantial breaking change but it's a minor adjustment and I feel it's important to be faithful to the SDK.

**Update (v6.3.253, 17 Nov 2023):**\
-Additional crypto APIs (both classic and nextgen)\
-Added GetSystemErrorString helper function to look up system error messages.\
-(Bug fix) FormatMessage did not follow W/DeclareWideString convention; last param not ByVal.\
-(Bug fix) RtlDestroyHeap has but one p.\
-(Bug fix) CoCreateInstance overloads not playing nice. Only a single form available now.

**Update (v6.3.252, 11 Nov 2023):**\
-Expanded bcrypt coverage\
-Added RegisterDeviceChangeNotification and the numerous assorted consts/types (dbt.h, 100% coverage)\
-Added DISP_E_* and TYPE_E_* error messages w/ descriptions. Added additional errors and descriptions for several original oleexp error sets.\
-The WBIDM enum that was full of IDM_* values has had the values changed to WBIDM_*. IDM_ is the standard prefix for menu resources, so these would often conflict with projects not using the same resource id, and the ids here are for Win9x legacy content.\
-All the fairly useless system info UDTs and an actually useful one, SYSTEM_PROCESS_ID_INFORMATION was missing.\
-Additional shell32 APIs\
-(Bug fix) Helper function NT_SUCCESS was improperly Private\
-(Bug fix) SetupDiGetClassDevPropertySheets[W] definitions incorrect

**Update (v6.3.250, 5 Nov 2023):**\
-Added Credential Provider interfaces from credentialprovider.h\
-Added missing TlHelp32.h APIs/structs, now covered 100%.\
-Added several types/enums related to things already in project.\
-(Bug fix) Duplicate of NETRESOURCE type. Project was subsequently analyzed for further duplicated types, and 4 other bugs in this class were eliminated.\
-(Bug fix) No base PEB type defined.\
-(NOTICE) OpenGL is being deferred until twinBASIC has Alias support (planned).

---

For earlier version history, see CHANGELOG.md

For more information and a list of available interfaces, visit the [VB Forums thread](https://www.vbforums.com/showthread.php?897883-twinBASIC-tbShellLib-Shell-Interface-Library-(x64-compatible-successor-to-oleexp)) for this project.
