# WinDevLib 
## Windows Development Library for twinBASIC

**Current Version: 9.1.618 (November 4th, 2025)**

(c) 2022-2025 Jon Johnson (fafalone)

> [!IMPORTANT]
> **Version 9.1.564 and higher now requires twinBASIC Beta 814 or newer.** Even if you're not using anything new. The package has hit a size threshold that due to a WebView2 bug will crash earlier versions if they attempt to load it.

WinDevLib is a project to make all common Windows API COM interfaces, DLL declares, and related Types/Enums/Consts available while programming in twinBASIC.\
Included are definitions of 3300+ common COM interfaces and 10,000+ APIs from all the common system modules, a level of coverage which makes WDL an entirely different experience than any VBx library, the largest of which offer at most 1/10th as much with huge gaps.\
This makes working with WDL similar to working in C++ with `#include <Windows.h>` and a number of other headers for commonly used features. These have all been redone by hand from the original headers, in order to restore 64bit type info lost in VB6 versions, avoid the errors of automated conversion tools (e.g. Win32API_PtrSafe.txt is riddled with errors), and make them friendlier by converting groups of constants associated with a variable into an Enum so it comes up in Intellisense. This takes advantage of tB's ability to provide Intellisense for types besides Long in API defs (hopefully UDTs soon, this project has provisioning for that). 

Creating this involves not only writing the definitions, but using tB compatible types-- so in some cases, even though there may be an existing way to import references to interfaces, they may be unusable due to e.g. the use of unsigned types, C-style arrays, double pointers, etc. In most cases these definitions are also compatible with VBA7, and with minor adjustments VB6; where they're not it's usually minor syntax adjustments, so this is also a great resource for APIs for those, covering vastly more than other other similar project.

This project is implemented purely in tB native code, as unlike VB6 there's language support for defining interfaces and coclasses. As a twinPACKAGE, regular code is supported in addition to the definitions, so some content normally found in regular addin modules have been built in (like you'd find in oleexp's mIID.bas, mPKEY.bas, etc, and helper functions). Does it still make sense to use a project like this when interfaces can be defined in-language? I'd say yes, because for a large number of interfaces, there's deep dependency chains with other interfaces and the types they rely on. It makes more sense to drop this in and be done with it than constantly have to define the interfaces you want and then stubs for their dependencies, especially when you might need those later on. This project is even more useful now with the API coverage; it should cover about 99% of your needs for core system DLLs. 

This project also serves a comprehensive twinBASIC replacement for [oleexp.tlb](http://www.vbforums.com/showthread.php?786079-VB6-Modern-Shell-Interface-Type-Library-oleexp-tlb), my Modern Shell Interfaces Type Library project for VB6. 100% of the content is covered with little to no change (just String arguments in some places due to differences between how they're handled in typelibs). 

Please report any bugs via the Issues feature here on GitHub.

### Requirements

[twinBASIC Beta 814 or newer](https://github.com/twinbasic/twinbasic/releases) is required.

### Adding WinDevLib to your project
You have 2 options for this:

#### Via the Package Server
twinBASIC has an online package server and WinDevLib is published on it. Open your project settings and scroll to the **Library References**, then click **Available Packages**. Add "Windows Development Library for twinBASIC v7.0.272" (or whatever the newest version is). The other similar entry, "WinDevLib for Implements" contains `Implements` compatible versions of a small number of common interfaces not defined in a compatible way in the main project; you normally don't need this. For more details, including illustrations, [see this post](https://github.com/fafalone/WinDevLib/issues/9#issuecomment-1416767019).

#### From a local file
You can download the project from this repository and use the WinDevLib.twinpack file. Navigate to the same area as above, and click on the "Import from file" button. WinDevLib.twinproj is the source for the package, if you want to edit it.


### Optional Features

#### Compiler Flags
WinDevLib has some compiler constants you can enable:

`WINDEVLIB_LITE` - This flag disables most API declares and misc WinAPI definitions, including everything in wdAPIComCtl, wdAPI, and wdDefs. I used to like doing my APIs separate too, which is why oleexp never had the expansive coverage. But with that coverage now present, I think it's worth using, but this option will still be supported.

`WDL_NO_DIRECTX` - Excludes DirectX, Media Foundation, XAudio, and WinML content. This is useful to substantially cut down on Intellisense entries in non-multimedia apps. Basic 2D graphics remain (GDI, GDI+, WIC).

`WDL_NO_COMCTL` - You can use this flag if you already have an alternative common controls definition set, e.g. tbComCtlLib; it will disable wdAPIComCtl. (Note: WinDevLib has more complete comctl defs than tbComCtlLib, as that project was deprecated and not updated).

`WDL_DLGSH` - This enabled constants from dlg.h. These are extremely uncommon to use, and have very short, generic names likely to cause conflicts, so they're opt-in.

`WDL_NOQUADLI` - Restores the old `LARGE_INTEGER` definition of lo/high Long values.

>[!WARNING]
>The `WDL_NOQUADLI` constant will break alignment on numerous Types; most only on x64, but some on both. 

`WDL_AVOID_INTRINSICS` - Uses the `Interlocked*` APIs that are exported from kernel32.dll (32bit mode only) instead of the static library containing compiler intrinsic versions of those in addition to all the ones not exported and all the 64bit ones.

`WDL_NO_LIBS` - Fully exclude static libraries (currently only Interlocked); mainly intended for comparing current tB versions to Beta 423 where the `Import Library` syntax is not yet supported.

`WDL_NO_DELEGATES` - Do not use Delegate functions in place of function pointers.

`WDL_XAUDIO8` - Use XAudio8 DLLs for XAudio2 APIs (Windows 8)

`WDL_NOMATH` - Exclude built in math helper function (see below). Note: XAudio2 inlined helper functions unavailable when math disabled.

`WDL_ADS_DEFINED` - activeds.tlb is referenced, enable interfaces using its contents.

>[!IMPORTANT]
>Currently flags are not inherited from the main project, so the only way to use these is to set them in the compiler flags for WinDevLib.twinproj then build a custom twinpack.

#### Custom Helper Functions
In addition to coverage of common Windows SDK-defined macros and inlined functions, a small number of custom helper functions are provided to deal with Windows data types and similar not properly supported by the language. These are:

`Public Function GetMem(Of T)(ByVal ptr As LongPtr) As T` - A generic to dereference a pointer into any type. The native `CType(Of )` allows dereferencing to UDTs, but this helper allows instrinsic types in addition to UDTs, and is used the same way.

`Public Function DCast(Of T, T2)(v As T2) As T` - Direct Cast: Copies the data of v into any type, without modification, so no overflows, and possible to e.g. go from `LongLong` to `POINT`, with `Dim pt As POINT = DCast(Of POINT)(SomeLongLong)`

`Public Function LPWSTRtoStr(lPtr As LongPtr, Optional ByVal fFree As Boolean = True) As String`\
Converts a pointer to an LPWSTR/LPCWSTR/PWSTR/etc to an instrinsic `String` (BSTR)

`Public Function UtfToANSI(sIn As String) As String`\
Converts a Unicode string to ANSI. This function is `[ConstantFoldable]` -- it can be used to create strings resolved at compile time and stored as constants; this technique was developed to use ANSI strings in kernel mode, where the APIs that handle a normal `String` cannot be used.

`Public Function VariantLPWSTRtoSTR(pVar As Variant, pOut As String) As Boolean`\
Retrieves a tB-style String from a VT_LPWSTR Variant. Returns False if pVar is a null pointer, or the Variant is not a VT_LPWSTR, or PropVariantToStringAlloc returns a nullptr.

`Public Function GetSystemErrorString(lErrNum As Long, Optional ByVal lpSource As LongPtr = 0) As String`\
`Public Function GetNtErrorString(lErrNum As Long) As String`\
Retrieve descriptions of `HRESULT` and `NTSTATUS` error codes, respectively.

`Public Function VariantSetType(pvar As Variant, [TypeHint(VARENUM)] ByVal vt As Integer, [TypeHint(VARENUM)] Optional ByVal vtOnlyIf As Integer = -1) As Boolean`\
Sets a Variant to the specified type without any alteration to the data. vtOnlyIf will abort the change and return False if the original type is other than specified. This should only be used when `VariantChangeType` is not applicable, and only with full understanding of consequences like automation errors if you attempt to use intrinsic operations on unsupported types; e.g. if you set the type to `VT_UI4`, then `CLng()` will raise a 'type unsupported' runtime error.

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

Math helpers:
```
   Functions: The first column take Double arguments, the second (with f) take Single (float).
   Log10, Log10f   - Base 10 logarithm; native Log is actually Ln 
   Pow, powf       - Power function for easier porting of code from langs w/o x^y.
   Asin, Asinf     - Arcsine
   Acos, Acosf     - Arccosine
   Atan, Atanf     - Arctangent (alias for Atn)
   Sec, Secf       - Secant
   Asec, Asecf     - Arcsecant
   Cosec, Cosecf   - Cosecant
   Acosec, Acosecf - Arccosecant
   Acotan, Acotan  - Arccotangent
   Sinh, Sinhf     - Hyperbolic sine
   Cosh, Coshf     - Hyperbolic cosine
   Tanh, Tanhf     - Hyperbolic tangent
   Sech, Sech      - Hyperbolic secant
   Cosech, Cosechf - Hyperbolic cosecant
   Cotanh, Cotanhf - Hyperbolic cotangent
   Asinh, Asinhf   - Hyperbolic arcsine
   Acosh, Acoshf   - Hyperbolic arccosine
   Atanh, Atanhf   - Hyperbolic arccotangent
   Asech, Asechf   - Hyperbolic arcsecant
   Acosech, Acosechf - Hyperbolic arccosecant
   Acotanh, Acotanh - Hyperbolic arccotangent
```

### Guide to switching existing code to WinDevLib

WinDevLib presented the best opportunity there would be to ditch some olelib legacy baggage. It's fairly simple to move your VB6 projects to WinDevLib, just follow these steps:

#### oleexp type library issues
The follow steps apply only if you're converting code that previously relied on my oleexp.tlb project:

1) Replace public aliases: It's important to do this first. Run a Replace All changing oleexp.LONG_PTR to LongPtr, oleexp.REFERENCE_TIME to LongLong, oleexp.HNSTIME to LongLong, oleexp.KNOWNFOLDERID to UUID, oleexp.EventRegistrationToken to LongLong, oleexp.BINDPTR to LongPtr, and oleexp.LPCRITICAL_SECTION to LongPtr. If you've used them without the oleexp. prefix, you'll also need to replace those, but if you've imported into tB they should be tagged. (These aliases will be restored as soon as tB supports it).

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

1) Convert `Currency` to `LongLong` for interfaces and APIs: It's no longer neccessary to worry about multiplying and dividing by 10,000 since tB supports a true 64bit integer type in both 32bit and 64bit mode. So this change is ultimately for the better, but existing codebases will have had to have used `Currency` for all interfaces and oleexp APIs expecting a 64bit integer.

2) Optional UDTs no longer use `As Any`. If you see errors like `Validation of call to 'CreateFile' failed.  Argument for 'lpSecurityAttributes': cannot coerce type 'Long' to 'SECURITY_ATTRIBUTES'`, this is an example of the issue. twinBASIC supports substituing `vbNullPtr` for a UDT (do not include `ByVal`), so WinDevLib can use the proper type while still permitting you to pass the equivalent of `ByVal 0`. 

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

3) String vs Long(Ptr) in APIs with both ANSI and Unicode versions: Most VB programs are written with ANSI versions of APIs being the default. **This is not the case with WinDevLib**. APIs are Unicode by default-- i.e. they use the W, rather than A, version of APIs e.g. `DeleteFile` maps to `DeleteFileW` rather than `DeleteFileA`. The A and W variants use String/LongPtr, and in almost all cases, the mapped version uses `String` with twinBASIC's `DeclareWide` keyword-- this disables Unicode-ANSI conversion. Since this is automatic, you generally don't need to make any changes; you can still use `String` without `StrPtr` or any manual Unicode <-> ANSI conversion. Note this usually only applies to strings passed as input, you'll need to update any externally allocated strings returned as a pointer, where you previously used e.g. `lstrlenA`, to use `lstrlenW` and Unicode handling in general. 

All APIs are provided, as a minimum, as the explicit W variant, and an untagged version that maps to the W version. Most ANSI variants are also included, but code should use Unicode wherever possible.

UDTs used by these calls are also supplied in the same manner, the W variant, an untagged variant that's the same as the W version, and in some cases, an A version. UDTs always use `LongPtr` for strings, even the untagged versions for `DeclareWide`. 

If you have any doubts about which API is being called, twinBASIC will show the full declaration when you hover your cursor over the API in your code.

Special thanks to GCUser99 for helping normalize API declaration in this project. 👍

4) Member names in UDTs have in almost all cases use their official SDK names, even where VB6 programmers traditionally used others. If you encounter errors where UDT members are missing, check the definition to see if the name has changed. This may also happen where unions are worked around in different ways.
   
6) **CURRENTLY N/A DUE TO BUGS** Callbacks that were previously LongPtr now expect a delegate function in many cases. A delegate is a typed function pointer defined with the correct prototype for the function it references. For compatibility these function the same way as previously and you need not make any changes, it's merely a better, easier way of defining callbacks that's also much closer to the C/C++ source. tB may show a warning, but it can be turned off.\
To use these, you can view the definition then make a regular function with that name and those arguments. See the [twinBASIC documentation overview of delegates](https://github.com/twinbasic/documentation/wiki/twinBASIC-Features#delegate-types-for-call-by-pointer) for more details and examples of using this new feature.

> [!TIP]
> Reminder: `Nothing` can be used in place of an interface where WinDevLib has the interface as an argument but another signature used `Long`/`LongPtr`


> [!NOTE]
>  This is just for using WinDevLib-- you'll likely have a lot more changes to make if you want to make your project x64 compatible. 


#### Scope of coverage

The goal of the API coverage in WinDevLib is to provide the kind of programming experience you'd get in C/C++ by including windows.h and many of the more common feature sets like DirectX and GDIPlus. Even the 10,000+ APIs are just scratching the surface of the total Windows API set, and due to the low quality of automated conversion, even by Microsoft themselves (see: Win32API_PtrSafe.txt), I'm not interested in simply feeding headers through a conversion utility or using a database, so instead WinDevLib will be focused on the most commonly used features in the major system DLLs-- everything 99%+ of apps need; though less common ones can be added by request or as time goes on and the existing DLLs are completed.

I've included the definitions, associated types, and associated constants, for extensive portions of the following modules: shell32.dll, shlwapi.dll, kernel32.dll, ktmw32.dll, user32.dll, advapi32.dll, tdh.dll, authz.dll, crypt32.dll, wintrust.dll, bcrypt.dll, ncrypt.dll, cryptui.dll, ole32.dll, oleaut32.dll, propsys.dll, gdi32.dll, gdiplus.dll, virtdisk.dll, userenv.dll, dbghelp.dll, mpr.dll, iphlpapi.dll, urlmon.dll, hlink.dll, winmm.dll, cfgmgr32.dll, setupapi.dll, comctl32.dll, dwm.dll/uxtheme.dll, comdlg32.dll, winbio.dll, winspool.drv, imm32.dll, hid.dll, cldapi.dll, pdh.dll, powrprof.dll, wtsapi32.dll, and netapi32.dll. Please let me know any I've missed from these.\
Limited coverage (or full coverage of very small sets) is provided for ntdll.dll, version.dll, msimg32.dll, fwpuclnt.dll, sxs.dll, secur32.dll, msacm32.dll, url.dll, htmlhelp.dll, avifil32.dll, and ws2_32.dll. If you feel any missing ones from these should be included, or would like to contribute more, let me know.\
Finally, there's numerous additional API sets from small to large for independent Windows features. These include small sets like restartmgr.dll through very large sets like the various Media Foundation and DirectX DLLs. In the future I'll better organize coverage lists, but the bottom line is let me know if any common APIs or built in API sets for components should be added. TODO.md in the WDL project files contains ones planned but not yet done.

**Future coverage:** In the future I'm planning to expand native APIs, complete legacy DirectX coverage, add SQL APIs, and add OpenGL-- though for these last two I may wait for tB to have `Alias` support since the SQL API has all custom SQL types, as does OpenGL which additionally has existing VB6 codebases which make heavy use of them by way of NeHe's TLB. I welcome contributions of any of these. If you've done the consts->enums conversions already, I'd even take 32bit-only versions. See CONTRIBUTING.MD for more information on that;


### ***tbShellLib is now WinDevLib - Windows Development Library for twinBASIC***
This project has grown well beyond it's original mission of shell programming. While that's still the largest single part, it's no longer a majority of the code, and the name change now much better reflects the purpose of providing a general Windows API experience like windows.h. Compiler constants and module names/file names have been updated to reflect the name change. tbShellLibImpl is now WinDevLibImpl. There are also some major chanages associated with this update, please see the full changelog below.

### Updates

**Update (v9.1.618, 04 Nov 2025):**
- Additional native APIs, inc dozens of missing types for NtQuerySystemInformation- I believe they're now all present
 through the latest phnt header covering through Win11 25H2
- Numerous arguments made optional for signature compatibility with oleexp.tlb.
- (Bug fix) Direct3DCreate9 definition incorrect
- (Bug fix) D3DPERF_ APIs used String without DeclareWide when LPCWSTR was expected


**Update (v9.1.614, 01 Nov 2025):**
- Added D3D Compiler interfaces (dxcapi.h, 100%)
- Added DsGetDC.h (100%)
- Additional undocumented shell interfaces and APIs
- Additional native APIs
- Additional types/consts for DeviceIoControl commands

**Update (v9.1.612, 30 Oct 2025):**
- Added WebDAV APIs (davclient.h, 100% inc. delegates)
- Additional native APIs
- (Bug fix) ITaskbarList3 missing method

**Update (v9.1.610, 27 Oct 2025):**
- (BREAKING CHANGE) LdrGetDllHandle now uses phnt signature of ByRef DllCharacteristics As Long for 2nd argument
- Additional loader, native, and low level actctx APIs
- (Bug fix) Some Websocket APIs used String for PCSTR* 

**Update (v9.1.608, 26 Oct 2025):**
- Experimental: WDL_QUALIFY compiler const will remove everything except interfaces and coclasses from the global
 namespace and require it to be prefixed with "WinDevLib."
- Added coverage of SubAuth.h (100%, inc delegates)
- (Bug fix) ISecurityInformation::GetSecurity missing argument
- (Bug fix) IEffectivePermission2 incorrect argument and API Standards noncompliance
- (Bug fix) MEMORY_BASIC_INFORMATION extra member in 32bit

**Update (v9.1.607, 16 Oct 2025):**
- (Bug fix) GetEnvironmentStrings[A], GetCommandLine[A], StrCat[A], StrCpyN[A], CharUpper[A], CharLower[A], 
      D3D10GetPixelShaderProfile, D3D10GetVertexShaderProfile, D3D10GetGeometryShaderProfile had String returns
      for non-BSTR strings, causing access violations or incorrect values.
- (WinDevLibImpl) Added Media Foundation PreserveSig notify interfaces. You *should* be able to use the
   versions in WinDevLib main, and indeed IMFTimedTextNotify has PreserveSig commented out and appears to
   work, but just in case I added them.
- (WinDevLibImpl) Removed empty modules since all they did was cause name conflicts.

**Update (v9.1.606, 10 Oct 2025):**
- Added Microsoft Active Accessibility Text Services interfaces/coclasses (MSAAText.h/.idl, TextStor.h/.idl 100%)
- Added many Comctl/dwn/uxtheme overloads for using either String or LongPtr/StrPtr
- Restored version gating around va_list APIs per https://github.com/fafalone/WinDevLib/issues/41
- Remaining _CONTEXT usages changed to CONTEXT.
- Misc Native API additions
- (Bug fix) wvnsprintfW definition incorrect
- (Bug fix) DrawShadowText expects LPCWSTR but used String without DeclareWide (now overloaded to accept either properly)

**Update (v9.1.603, 24 Sep 2025):**
- Some additional process heap APIs
- Misc bug fixes and API standards corrections
- (Bug fix) ICallFrameWalker, D3D10CreateBlob ByVal/ByRef
- (Bug fix) ISearchCatalogManager::GetParameter definition incorrect
- (Bug fix) IViewObject::Draw definition incorrect for x64

**Update (v9.1.602, 24 Sep 2025):**
- Added numerous missing COM APIs/interfaces from objidl.idl, objidlbase.idl and objbase.h
- (Bug fix) Numerous instances of LongPtr that should be As Any and ByVal/ByRef mixups in additions from last release.

**Update (v9.1.600, 23 Sep 2025):**
- Added numerous missing COM APIs/interfaces from objidl.idl, objidlbase.idl and objbase.h
- (Bug fix) CoFileTimeNow PreserveSig(False) overload definition incorrect.
- (Bug fix) Some 'As GUID' arguments escaped replacement with UUID.

**Update (v9.1.596, 21 Sep 2025):**
- Added missing standard shell header tlogstg.h/.idl (100%)
- Added missing standard shell header PathCch.h (100%)
- Added missing standard shell header ScrnSave.h (95%; some constants were highly generic names and skipped)
  Note: ids* constants prefixed with scrnsv_ and placed in enum ScreenSaverIDs.
- Added missing standard shell header appmgmt.h. Note: Some constants with simple, common names had prefixes added. See header region in wdAPI.twin.
- Added missing standard shell header Reconcil.h (100%)
- Added ActiveIMM interfaces (Dimm.h/.idl, 100%)
  
**Update (v9.1.595, 06 Sep 2025):**
- (Bug fix) ID2D1DeviceContext5::CreateSvgDocument missing argument
- (Bug fix/API Standards) Many uses of Currency not replaced with LongLong (where not explicitly Currency in the SDK). In some cases this would have caused improper alignment.

**Update (v9.1.594, 03 Sep 2025):**
- Added Direct3D 8 and DirectPlay interfaces for additional dxvb conversions. I now intend to cover all DXVB equivalent C defs.
- (Bug fix) Duplicated consts MDITILE_* and MDIS_ALLCHILDSTYLES, LOGONID_CURRENT, SERVERNAME_CURRENT, D3D11_DEFAULT_SLOPE_SCALED_DEPTH_BIAS,
      D3D11_DEFAULT_VIEWPORT_MAX_DEPTH, D3D11_DEFAULT_VIEWPORT_MIN_DEPTH, OLEIVERB_PROPERTIES, DISPID_IADCCTL_*, DVB_ST_PID_16/17/18/19/20,
      WINSTATIONNAME_LENGTH, and DOMAIN_LENGTHDOMAIN_LENGTH.
  
**Update (v9.1.592, 29 Aug 2025):**
- (Bug fix) hostent and netent definitions incorrect (https://github.com/fafalone/WinDevLib/issues/40)

**Update (v9.1.591, 27 Aug 2025):**
- (Bug fix) Some DirectDraw interfaces missing correct inheritance
- 
**Update (v9.1.590, 27 Aug 2025):**
- Added SENS APIs/interfaces (SensAPI.h, Sens.h, SensEvts.idl 100%)
- Added GetProcAddress overload (https://github.com/fafalone/WinDevLib/issues/38)
- (Bug fix) SCardUIDlgSelectCard[A,W], GetOpenCardName[A,W] are in scarddlg.dll, not winscard.dll
- (Bug fix) PostMessage[A,W] wParam arg incorrect for x64 (https://github.com/fafalone/WinDevLib/issues/39)

**Update (v9.1.588, 09 Aug 2025):**
- Misc API additions for upcoming project

**Update (v9.1.586, 09 Aug 2025):**
- Added HTTP Server API (http.h, 100% inc. macros)
- Added Image Color Management / Windows Color System APIs (icm.h, wcsplugin.h/.idl, 100%)
- Misc API additions

**Update (v9.1.585, 03 Aug 2025):**
- (Bug fix) New DirectInput interfaces used stdole.GUID

**Update (v9.1.584, 02 Aug 2025):**
- Added DirectInput (dinput.h, 100% inc. macros, delegates and statically exported data)
- Added WMI utility interfaces (WMIUtils.h/.idl, 100%; the system typelib for this is full of unsupported types)
- (Bug fix) Many helper functions used ByRef instead of ByVal for in only args, which causes issues with other functions that call them.

**Update (v9.1.581, 30 Jul 2025):**
- Some new WBEM interfaces used inconvenient LongPtr instead of ByRef Interface
**Update (v9.1.580, 29 Jul 2025):**
- Added WBEM Client COM interfaces (WbemCli.h/.idl, 100%). Note: WDL will not duplicate the WMI Scripting Library, and in most cases you should use that.
- Misc. API additions

**Update (v9.1.578, 29 Jul 2025):**
- Add coverage of COM interceptors (callobj.h, 100%)
- Add coverage of WinNls32.h, ime.h (100%)
- Add coverage of poclass.h (100%)
- Misc. API additions
- (Bug fix) WNetDisconnectDialog name incorrect; WNetRestoreSingleConnectionA does not exist

**Update (v9.1.572, 25 Jul 2025):**
- Added deleted file restore APIs (fmapi.dll)
- Breaking Change: NtCreateToken[Ex] now uses proper LARGE_INTEGER type instead of LongLong.
- Added numerous missing keys from propkey.h added from when modPKEY was initially done with the
   Windows 7 SDK to the latest Windows 11 SDK. 
- Breaking Change: IEnumExplorerCommand::Next now returns HRESULT; Implements version added to WinDevLibImpl. 
- Misc. API additions

**Update (v9.1.570, 15 Jul 2025):**
- **BREAKING CHANGES** 
   - LUID_AND_ATTRIBUTES LUID member is not a pointer so "pLuid" was not only 
       wrong but misleading. Now just Luid to match SDK.
   - TOKEN_OWNER and TOKEN_PRIMARY_GROUP members now use their name rather than type.
- Updated WebView2 definitions to match stable release 1.0.3351.48 
- Added windowsx.h macros for ListBox, ComboBox, and ScrollBar.
- Misc API additions (inc. native api sync and richedit undoc'd)
- (Bug fix) Duplicated constant ST_PLACEHOLDERTEXT 
- (Bug fix) NtCreateToken / NtCreateTokenEx missing ObjectAttributes argument.
  
**Update (v9.1.567, 08 Jul 2025):**
- IDWriteColorGlyphRunEnumerator had its GetCurrentRun method named GetCurrentRun1, which would be confusing when IDWriteColorGlyphRunEnumerator1 was just GetCurrentRun. They're both GetCurrentRun now as they are in SDK. 
- (Bug fix) IFileOperationProgressSink::PostNewItem missing argument.

**Update (v9.1.566, 02 Jul 2025):**
- Added coclasses for ListView subitem controls (using their common CLSID-derived names,
   CBooleanControl for CLSID_CBooleanControl, etc).
- Added DirectShow BDA interfaces not covered by VBx/tB compatible tuner typelib. 
- Some DirectWrite enum values from dwrite.h were missing.
- Added numerous additional PE header types/consts from winnt.h.
- Added undocumented IGlobalOptions/ISecurityOptions and GlobalOptions coclass.
- Misc API additions
- (Bug fix) DWM_TIMING_INFO and DWM_THUMBNAIL_PROPERTIES missing req'd PackingAlignment attrib.

**Update (v9.1.564, 22 Jun 2025):**
- **IMPORTANT:** WinDevLib now requires twinBASIC Beta 814 or newer, *regardless of whether you're
  using anything new.* This is due a longstanding bug concerning the size of packages, and WDL is 
  now large enough that it triggers this bug. 
- (API Standards) **BREAKING CHANGE** :: Shell functions taking pidl arrays were inconsistently
  defined. Some took ByVal and some took ByRef (VarPtr(pidls(0)) vs just pidls(0)). For the sake
  of consistency, correctness, and WDL API standards, SHCreateShellItemArrayFromIDLists, SHCreateDataObject, 
  SHCreateFileDataObject, and IDefaultFolderMenuInitialize::Initialize have now been changed to 
  use the more correct ByRef semantics. Where you passed `VarPtr(pidls(0))` you'll need to change
  that to just `pidls(0)`. oleexp will also change in its next release.
- Added some urlmon.h content that was strongly related to that already included.
- Added all error consts from sherrors.h
- Added META_ metafile function codes missing from current SDK headers (but present in older ones)
- AVISave[A,W] functions no longer [Unimplemented] 
- PROPVARIANT APIs now all take As Any to accomonodate use of `PROPVARIANT` UDT as well as Variant. Most
  inlined APIs do not yet, pending a bug fix in overload resolution. 
- New helpers InitPropVariantFromStringPtr/VariantSetTypePtr for versions of the original that take 
  a LongPtr to a Variant/PROPVARIANT instead. LongPtr for String overloads for InitPropVariantFromString[Ptr].
- For compatibility, IPropertyValue will now use `PROPVARIANT` UDT instead of tB Variant.
- (Bug fix) IPropertyValue::InitValue definition incorrect.

  
**Update (v9.0.562, 13 Jun 2025):**
- Added some remaining DirectShow content (dvdif.h 100%, strmif.h now 100%)
- (Bug fix) STRRET did not account for x64 union padding. 

**Update (v9.0.560, 12 Jun 2025):**
- Added complete coverage of DirectDraw (ddraw.h, ddstream.h 100%)
  - (While highly similar, this is not equivalent to the DX7VB implementation. That uses
    a C++ intermediate that rewrites and translates a lot of stuff; it's not practical
    to reimplement. If anyone finds themselves struggling with a missing helper from that,
    I can help with a reimplementation. Note that other DirectX 7 and 8 technologies won't
    be added in the near term; DirectDraw was added for a DirectShow expansion)
- Added legacy DDraw Video Mixer interfaces (vmr9.h, vmr9.idl, vmrender.idl 100%), Video 
   (amvideo.h 100%), and Video Port interfaces (Dvp.h 100%)
- Major expansion DirectShow coverage (axextend.idl, amvideo.h, amaudio.h, MpegTypes.h, VpConfig.h,
    VpTypes.h, dvdmedia.h, edevdefs.h/xprtdefs.h, amparse.h, vidcap.h/.idl, dmodshow.h/.idl, 
    CameraUIControl.h/.idl, il21dec.h, iwstdec.h 100%)
- Added all DXVA types and interfaces (dxva.h, 100%; DXVA2 and DXVAHD already covered)
- More D3DX coverage (d3dx9xof.h, d3dx9mesh.h, d3dx9shape.h, d3d9xmath.h 100%)
- Added missing evr.h APIs.
- New project flag: WDL_NO_DIRECTX. Excludes directly all DirectX technologies:  
   Media Foundation is considered part of and is tightly linked with it, and is also excluded.  
   WinML is dependent on it and also excluded.  
   Some parts of WIC, TextServices (RichEdit), and WMDM have DirectX interfaces replaced with IUnknown.  
   This is part of a planned series of flags to disable major components you don't need to limit how
   much is in the symbol table for Intellisense. Everything will remain enabled by default* until tB
   supports namespaces properly (far in the future).  
    - * Constants requiring a flag to be enabled now will remain that way.
- WINDEVLIB_LITE flag now also disables GDIP, ETW, WIM, and, XAudio inlines.
- Other flags shortened or modified for consistency:    
   WINDEVLIB_NO_WS_ALIASES -> WDL_NO_WS_ALIASES   
   WINDEVLIB_COMCTL_LIB_DEFINED -> WDL_NO_COMCTL   
   WINDEVLIB_DLGSH -> WDL_DLGSH  
   WINDEVLIB_NOQUADLI -> WDL_NOQUADLI  
   WINDEVLIB_AVOID_INTRINSICS -> WDL_AVOID_INTRINSICS   
   WINDEVLIB_NOLIBS -> WDL_NO_LIBS  
   ADS_DEFINED -> WDL_ADS_DEFINED   
   WINDEVLIB_NO_DELEGATES -> WDL_NO_DELEGATES (still disabled)  
   WINDEVLIB_XAUDIO8 -> WDL_XAUDIO8   
   WINDEVLIB_NOMATH -> WDL_NO_MATH   
- wvnsprintf and wvsprintf reinstated. This requires tB Beta 797 or newer. For the next few months,
   these will be gated off in `#If TWINBASIC_BUILD >= 797` version checks so the minimum for the whole
   project isn't raised, but that will change eventually.
- (Bug fix) A handful of GUID function (IID_, GUID_, etc) were wrongly defined and would return GUID_NULL.
- (Bug fix) Some duplicated enum values.
- (Internal) Changelog/readme markdown files formatted for markdown preview mode now that it's default.  
   Reminder: If you're reading this on the web, you can also view this changelog by navigating to the WDL
   package in the Project Explorer under Packages. This is updated with every new release.


**Update (v8.12.552, 06 Jun 2025):**\
-More D3DX coverage (d3dx11.h, d3dx11async.h).\
--Note: D3DX11 APIs use d3dx11_43.dll, and D3D9X APIs use d3dx9_43.dll. This are the most recent versions, but may not be included with Windows 10 and 11 installations. It's recommended you obtain the June 2010 DirectX SDK for redistributable files you can install. You can also downgrade to installed versions.\
        https://www.microsoft.com/en-us/download/details.aspx?id=6812 and https://www.microsoft.com/en-us/download/details.aspx?id=8109
        
-Added missing constants for IFilter HRESULTs.\
-Added NTQuery.h coverage\
-Added appnotify.h coverage (100%)\
-Added missing WIN32_FILE_ATTRIBUTE_DATA\
-Misc API additions\
-(Bug fix) IFilter::GetValue definition incorrect.

**Update (v8.12.550, 05 Jun 2025):**\
-Added some D3DX coverage (d3dx9core.h, d3d9x.h, d3d9xshader.h, d3d9xtex.h, d3dx11core.h, d3dx11tex.h 100%)\
-Added additional Winsock APIs (ws_closesocket from ws2api; then WS2spi.h partial, SpOrder.h 100%)\
-Added numerous missing shlwapi aliases and some missing functions.\
-Added Filter Manager usermode APIs (fltUserStructures.h, fltUser.h 100%)\
-wvsprintf[A,W] didn't make use of ByRef ParamArray args As Any()\
-(API Standards) Changed numerous byte array inputs typed as Byte to As Any to conform with standard.\
-(API Standards) StrCpy didn't use String.\
-(Bug fix) Duplicated constant: FACILITY_HID_ERROR_CODE\
-(Bug fix) IAMMediaTypeSample::GetPointer incorrectly returned Byte instead of LongPtr for a double pointer.\
-(Bug fix) String overload for StrCmpLogicalW didn't use DeclareWide\
-(Bug fix) WSASocket invalid duplicate (Thanks to forliny)\
-(Bug fix) COINIT_MULTITHREADED value incorrect (Thanks to forliny)\
-(Disabled) wvnsprintf and wvsprintf functions commented out pending tB bugfix

**Update (v8.12.544, 27 May 2025):**\
-Added additional Media Foundation interfaces/APIs from wmcontainer.h, ksopmapi.h, opmapi.h (100%)\
-Added additional DirectShow interfaces (axcore.idl, now 100%; medparam.h, dmoreg.h 100%)\
-(Bug fix) GetMem(Of T) helper generic broken

**Update (v8.12.542, 24 May 2025):**\
-Added DXGI debug interfaces/APIs (dxgidebug.h, 100%)\
-Added effect processor CLSIDs, MEDIASUBTYPE_ GUIDs, and MFPKEY_ PROPERTYKEYs from wmcodecdsp.h\
-Added structs/guids from dxva9typ.h (100%)\
-Added some additional undocumented RichEdit constants/enums/types.
-(Bug fix) DCast helper wasn't working with a UDT as the source.\
-(Bug fix) MAX_DEINTERLACE_SURFACES value incorrect.

**Update (v8.12.539, 21 May 2025):**\
-(Bug fix) Several ByRef As Byte that should be ByRef As LongPtr in Media Foundation interfaces.\
-(Bug fix) Numerous ByVal/ByRef mixups in Media Foundation interfaces. 

**Update (v8.12.538, 20 May 2025):**\
-Added SmartCard API (winscard.h, winsmcrd.h, SCardErr.h 100% inc delegates etc)\
-Added SSL-related APIs from schannel.h (100% including delegates etc)\
-Added numerous missing WIC error consts\
-Helper generic DCast now includes a safety check that the source type isn't smaller than the destination type, and if it is, only copies the number of bytes in the source.
-The following interfaces are clearly meant to be used with Implements but used [PreserveSig]\
      IMFTimedTextNotify, IMFMediaSourceExtensionNotify, IMFBufferListNotify, IMFBufferListNotify, IMFMediaEngineNeedKeyNotify, IMFMediaEngineEMENotify, IMFMediaKeySessionNotify2\
   [PreserveSig] was removed but that means they'll likely require v-table swaps or redirects to not crash.\
   Tip: You can copy these interfaces to your project and use [RedirectToStaticImplementation] to simplify.\
-(Bug fix) MFInitAMMediaTypeFromMFMediaType definition incorrect.\
-(Bug fix) New GetMem generic helper used Len instead of LenB.


**Update (v8.12.534, 16 May 2025):**\
-Added common control macros for Edit, Button, Tab, DateTime, MonthCal, Static, IPAddress, Animate controls.\ 
   In all cases, these include the macros from both commctrl.h and windowsx.h.\
-Added helper function `GetMem(Of T)` generic to dereference and cast a LongPtr to any type, even intrinsic types.\
-Added helper function `DCast(Of T, T2)` (direct cast) to copy `LenB(Of T)` bytes from any type, with no conversion like CInt would do where 65535 would overflow instead of giving -1. Also allows converting to UDTs, e.g. If you have ptll As LongLong containing a `POINT`, `Dim pt As POINT = DCast(Of POINT)(ptll)`.\
-Some Tooltip types were only defined by their tag names instead of proper names. Tag names remain for compatibility.\
-(Bug fix) Some GET_*_WPARAM helpers would overflow due to use of CLng().

**Update (v8.12.532, 13 May 2025):**\
-Added lcid/LANGID helpers and some additional internationalization APIs\
-Added WINDEVLIB_NO_WS_ALIASES compile const to remove ws_ prefix from Winsock functions with short, generic names (bind, socket, recv, etc)\
-Added keycredmgr.h, 100% all\
-Added lzexpand.h, 100% all\
-(Bug fix) MappingRecognizeText used MAPPING_ENUM_OPTIONS instead of MAPPING_OPTIONS

**Update (v8.12.530, 10 May 2025):**\
-Basic date/time format APIs from datetimeapi.h were inexplicably not done yet.\
-Added Extended Linguistic Services (ELS) APIs from ELSCore.h and ElsSrvc.h, 100% coverage.\
-Added D3DX11 General Purpose GPU computing algorithms (d3dcsx.h, 100%)\
-Added remaining ETW interfaces/APIs from evntprov.h, relogger.h (100% inc. delegates, macros, and inlines)\
-Added DirectManipulation interfaces/etc (directmanipulation.h, 100%)\
   Note: This was done assuming "LIFTED_SDK" was not defined. There's some deleted vtable entries, additional interfaces, additional coclasses, and entirely different GUIDs for everything if that is defined; the meaning is entirely undocumented. Will look into it in the future.\
-D2D1 PredeclaredId class from The trick's bas for e.g. D2D1::RectF. Disabled by default, to enable, set WINDEVLIB_DXHELPERS\
   Note: __F functions will be converted to overloads pending a tB bug fix concerning them.\
-Added some missing content from lmaccess.h and lmwksta.h to bring coverage to 100%; added LMalert.h, LMaudit.h, LMErrlog.h, LMRemUtl.h, LMSvc.h, LMDFS.h 100%
-Some netapi32 structs changed from String to LongPtr for consistency with vast majority of others.

**Update (v8.11.528, 08 May 2025):**\
-Added WebAuthN APIs (Windows Hello and other new security tokens; webauthn.h 100%)\
-IWICImageEncoder methods now use proper ID2D1Image type. (This is a breaking change against typelibs, but the next version of oleexp will use it too)\
-PROPVARIANT now uses more convenient 2x/4x Long, renamed pVar/pVar2/etc to harmonize with oleexp (unnamed in SDK)\
-(Bug fix) WICImageParameters improperly substituted Long for D2D1_PIXEL_FORMAT (now used).

**Update (v8.11.526, 05 May 2025):**\
-Added Direct3D 10. Was weird having 9, 11, and 12 but not 10.\
   100% coverage of d3d10.h, d3d10misc.h, d3d10shader.h, d3d10effects.h, d3d10sdklayers.h, d3d10_1shader.h, d3d10_1.h\
-Added Windows Lockdown Policy APIs (wldp.h, 100% inc. all). Note: VALUENAME enum renamed WLDP_VALUENAME.\
-Added Activity Coordinator API ActivityCoordinator.h, ActivityCoordinatorTypes.h - 100% (Win11+)\
-(Bug fix) ID2D1Bitmap inherits from ID2D1Image. No consequences besides a warning in some circumstances, since ID2D1Image has no methods.\
-(Bug fix) Some D3D_PRIMITIVE_TOPOLOGY values incorrect.\
-(Bug fix) A number of uxtheme APIs were missing ByVal on LPWSTR arguments.

**Update (v8.10.524, 02 May 2025):**\
-Added XAudio2 interfaces and APIs - xaudio2.h, xaudio2fx.h, x3daudio.h, xapo.h, xapobase.h, hrtfapoapi.h 100%\
   IMPORTANT: For Windows 8, define compiler constant WINDEVLIB_XAUDIO8.\
   NOTE: Inlined functions included, but the math conversion from C to tB has not yet been verified accurate.\
-Misc Native API additions, including NtCurrentTeb implemented by `Emit()`.\
-Completed adding known documented CLSID_xxx constants in usuable UUID form for all coclasses.\
-Added numerous overloads for compatibility with oleexp.tlb API signatures using `[PreserveSig(False)]` (where the last argument becomes the return)\
-CoInitialize/OleInitialize/vbCoInitialize now use Optional ByVal LongPtr for useless reserved argument.\
-Added math helpers. Constants from corecrt_math_defines.h;\
   Functions: The first column take Double arguments, the second (with f) take Single (float).\
   Log10, Log10f   - Base 10 logarithm; native Log is actually Ln\
   Pow, powf       - Power function for easier porting of code from langs w/o x^y.\
   Asin, Asinf     - Arcsine\
   Acos, Acosf     - Arccosine\
   Atan, Atanf     - Arctangent (alias for Atn)\
   Sec, Secf       - Secant\
   Asec, Asecf     - Arcsecant\
   Cosec, Cosecf   - Cosecant\
   Acosec, Acosecf - Arccosecant\
   Acotan, Acotan  - Arccotangent\
   Sinh, Sinhf     - Hyperbolic sine\
   Cosh, Coshf     - Hyperbolic cosine\
   Tanh, Tanhf     - Hyperbolic tangent\
   Sech, Sech      - Hyperbolic secant\
   Cosech, Cosechf - Hyperbolic cosecant\
   Cotanh, Cotanhf - Hyperbolic cotangent\
   Asinh, Asinhf   - Hyperbolic arcsine\
   Acosh, Acoshf   - Hyperbolic arccosine\
   Atanh, Atanhf   - Hyperbolic arccotangent\
   Asech, Asechf   - Hyperbolic arcsecant\
   Acosech, Acosechf - Hyperbolic arccosecant\
   Acotanh, Acotanh - Hyperbolic arccotangent\
   As with the native trig functions, these are in radians.\
   To disable, define #WINDEVLIB_NOMATH. Note: XAudio2 inlined helper functions unavailable when math disabled.\
   Note: Currently not verified for accuracy; I believe I tested most of these when I wrote them decades ago, but can't remember for sure and will need time to re-check them.\
-(Bug fix) Certain oleaut32 Var*, and some hlink, functions improperly used String without DeclareWide\
-(Bug fix) StrRetToStr[A] incorrect signature, inconsistent use of ByRef/ByVal


**Update (v8.9.520, 27 Apr 2025):**\
-Added Uniscribe API (usp10.h, 100%)/ UDTs harmonized with work by Michael Kaplan and Tanner Helland\
      However the APIs they used have signatures that just stray way too far from the documentation; many ByVal LongPtr arguments are now ByRef. Reminder: vbNullPtr replaces ByVal 0 for skipping an optional UDT.\
-Added coverage of DSAdmin.h. Note: The interfaces for this rely on activeds.tlb. After you add a reference to that, add the compiler option ADS_DEFINED=1.\
-Added numerous missing Visual Styles theme constants, vssym32.h 100%\
-Added basic Winstation APIs from phnt winsta.h.\
-Because they may contain pointers to data stored in a contiguous byte array, MEM_EXTENDED_PARAMETERS arguments have been changed to As Any. No change is needed to existing code.\
-Misc Native API additions\
-(Bug fix) HD_TEXTFILTERW name typo.


**Update (v8.9.518, 23 Apr 2025):**\
-**BREAKING CHANGE** SHCreateShellItemArray will now use the proper definition of ByRef ppidl As LongPtr. Workarounds using ByVal VarPtr() should remove that.\
-**BREAKING CHANGE** Since tB supports overloads, DirectComposition overloaded methods have had their tag (usally _A) removed. Affected interfaces:\
   IDCompositionVisual, IDCompositionVisual3, IDCompositionGaussianBlurEffect, IDCompositionBrightnessEffect,  IDCompositionColorMatrixEffect, IDCompositionShadowEffect, IDCompositionHueRotationEffect, IDCompositionSaturationEffect, IDCompositionLinearTransferEffect, IDCompositionTableTransferEffect, IDCompositionArithmeticCompositeEffect, IDCompositionAffineTransform2DEffect, IDCompositionTranslateTransform, IDCompositionScaleTransform, IDCompositionRotateTransform, IDCompositionSkewTransform, IDCompositionMatrixTransform, IDCompositionEffectGroup, IDCompositionTranslateTransform3D, IDCompositionScaleTransform3D, IDCompositionRotateTransform3D, IDCompositionMatrixTransform3D, IDCompositionRectangleClip, ID2D1SvgStrokeDashArray, IDWriteGdiInterop1, IDWriteFontFace4, IDWriteFactory4, IDWriteFontSet1\
   Note: ID2D1SvgElement overloads currently left tagged because tB cannot disambiguate 2 of them.\
   Note: This is experimental. Please report any problems. May be reverted if any arise.\
-Added missing IDXGIFactory6/7 interfaces from dxgi_6.h\
-Added custom UUIDs for system default GDIP encoders: ImageCodecBMP, ImageCodecJPG, ImageCodecGIF, ImageCodecTIF, ImageCodecPNG,  and ImageCodecICO. It's still advisable to use the documented way of finding these.\
-Added some missing interfaces, enums, and consts from oleidl.h.\
-Some imagehlp (dbghelp) APIs with only ANSI versions now use String for input instead of LongPtr\
-Misc API additions\
-(API Standards) WTSSetUserConfig[A,W] did not follow String/LongPtr convention for buffer arg\
-(Bug fix) DXGI_FORMAT missing and incorrect values\
-(Bug fix) SELFREG_E_CLASS value incorrect\
-(Bug fix) WTSSetUserConfig incorrect alias\
-(Bug fix) ByRef/ByVal mixups:\
         UiaNavigate, UiaFind, UiaNodeFromPoint, UiaNodeFromFocus\
         ISyncMgrHandler::Synchronize\
         IDXGIDevice2::ReclaimResources/::OfferResources, IDXGISwapChain::GetFullscreenState, IDXGIDevice::QueryResourceResidency, IDXGIDevice4::OfferResources1/::ReclaimResources1, ID3DXInclude::Open 
         
**Update (v8.8.516, 15 Apr 2025):**\
-Added all missing MetaFile/ENHMF APIs and structs\
-Added numerous other missing gdi32 APIs\
-Added missing APIs from coml2api.h, now 100% covered\
-Changed As BITMAPINFO args to As Any since this sometimes uses a variable C-style array.\
-(Bug fix) EnumEnhMetaFile, DeleteEnhMetaFile returned Boolean (2 bytes) instead of BOOL (4 bytes)\
-(Bug fix) ENHMETA_SIGNATURE conditional compilation value wrong\
-(Bug fix) CFSEPCHAR type and value incorrect

**Update (v8.8.513, 31 Mar 2025):**\
-winspool.h now covered 100%; added async printer notification ifaces/apis from prnasnot.h (100% coverage)\
-(Bug fix) PRINTER_NOTIFY_INFO_DATA, INPUT incorrect union substitution sizes; sorry don't know how I missed them in the 8.8.504 fix.\
-(Bug fix) PRINTER_OPTION_FLAGS incorrect and missing values.

**Update (v8.8.512, 26 Mar 2025):**\
-Updated WebView2 to 1.0.3124.44 Release SDK\
-Added missing 32bit aliases for GetWindowLongPtr[A,W]/SetWindowLongPtr[A,W]/GetClassLongPtr[A,W]/SetClassLongPtr[A,W]\
-UNREFERENCED_PARAMETER is now available as a generic; this lets you opt individual variables/arguments out of compiler messages about unused variables instead of opting out whole functions.
-Misc minor additions/fixes   

**Update (v8.8.511, 20 Mar 2025):**\
-(Bug fix) DirectX 2D arrays updated to match the layout you see with oleexp and other VB6 typelibs. The dimensions are inverted, e.g. m(y,x) instead of m(x,y) in VB/tB arrays, in order to get the same memory layout C/C++ expects from a caller of these interfaces/APIs. While VB6's object browser shows it as x,y, when you actually try to use the oleexp.tlb matricies, being compiled with C tooling, you'll see the compiler treats it as y,x.\
            So where the TLB has `FLOAT m[3][2]` in the source, the VB6 Object Browser says `m(0 To 2, 0 To 1) As Single`, but then `m(2, 1) = 1` will raise a 'Subscript out of range' error, while `m(1, 2) = 1` will work. tB matches this behavior (but shows the definition consistently), so this change is to match VB6/oleexp/other typelibs and is easier than remapping to the different coordinates.\
            This was previously applied to some but not all matricies.\
-(Bug fix) ID3D12GraphicsCommandList::OMSetBlendFactor, ID3D11DeviceContext::OMSetBlendState, ID3D11DeviceContext1::ClearView, ID3D12GraphicsCommandList::::ClearUnorderedAccessViewUint, ID3D12GraphicsCommandList::::ClearUnorderedAccessViewFloat improperly had a SAFEARRAY.\
            Note: Due to unsupported syntax, the array notation isn't used, but you would pass ArrayOfValues(0).


**Update (v8.8.509, 19 Mar 2025):**\
-LOWORD and HIWORD now use assembly functions made from the C macros on x64.\
-(Bug fix) IDragSourceHelper IDataObject params missing ByVal, causing crashing.

**Update (v8.8.507, 17 Mar 2025):**\
-While tB language features make using them as-is possible, for compatibility with VB6 code, QueryServiceConfig[A,W], EnumDependentServices[A,W], EnumServicesStatus[A,W,Ex,ExA,ExW], QueryServiceLockStatus[A,W] and GetUserObjectSecurity require a buffer for all the strings pointed to by the return type, so must use As Any instead of As the UDT mentioned.\
   **NOTE:** This is a breaking change if you were already using the tB-language way; you'd have to add ByVal. No error will be generated, it will just crashing without being changed.\
-(Bug fix) ChangeServiceConfig2[A,W], RegisterServiceCtrlHandlerEx[A,W], ReportEvent[A,w], and GetModuleHandle had As Any params marked Optional (unsupported by language) 

**Update (v8.8.506, 15 Mar 2025):**\
-Added Performance Counter APIs from perflib.h and winperf.h (100% coverage inc delegates and UseGetLastError)\
-Added Xinput APIs. Note: DLL name for Win8+ used. Separate defs for Win7/Vista are provided with the suffic -7, e.g. XInputGetState7\
-Large expansion of Setup APIs; SetupAPI.h should now be 100% inc. Delegates.\
-Completed updating Direct3D 12 to SDK 10.0.26100.0\
-Misc API additions, including additional native APIs.\
-(Bug fix) ID3D12GraphicsCommandList10 method definitions incorrect.\
-(Bug fix) ChangeWindowMessageFilterEx 'action' was set to the wrong enum, and the right one was missing.\
-(Bug fix) InitializeSid missing ByVal\
-(Internal) ntdll and kernel32 APIs moved to wdAPINTKernel.twin to reduce size of wdAPI

**Update (v8.8.504, 10 Mar 2025):**\
-Added DirectStorage - dstorage.h, dstorageerr.h - 100% coverage (Note: Some versions of Windows may not have DLL preinstalled)\
-Added Windows Imaging Interface APIs (wimgapi.h, 100% coverage)\
-Completed WinDNS.h coverage (now 100% including macros, delegates, and UseGetLastError)\
-Added some missing DirectSound interfaces and constants from dsound.h.\
-Added some missing Portable Devices interfaces and coclasses from portabledeviceclassextension.h and portabledevicetypes.h.\
-Min/max/etc for Single and Double from float.h\
-(Bug fix) DS3D_DEFERRED name typo\
-(Bug fix) Numerous errors when WINDEVLIB_LITE flag set; had been ignoring that since it could only be used by compiling a custom version of the package, but that should change soon.\
-(Bug fix) Dozens of union substitutions incorrect due to not always accounting for padding needed before or after the bytes making up the union data, or in some cases the size of the union data itself (particularly for x64).
           
**Update (v8.7.502, 06 Mar 2025):**\
-Added 100% coverage of processtopologyapi.h and systemtopologyapi.h\
-Added 100% coverage of audiostatemonitorapi.h\
-Added improperly excluded vararg functions of oledlg.h, now 100% coverage\
-Added some missing items to bring shobjidl.h/.idl, ShlObj_core.h, thumbcache.h/.idl and timezoneapi.h to 100%\
-For consistency, GETTEXTEX now uses LongPtr instead of String.\
-Renamed MENUPOPUPPOPUPFLAGS to SDK-defined MP_POPUPFLAGS\
-Added IShellIconOverlayIdentifier::GetOverlayInfo missing flags\
-Continued work to supply usable UUID types for documented CLSID_ constants for coclasses.\
-Continued implementation of [UseGetLastError(False)]; applied to all NTSTATUS-returning APIs\
-Numerous other misc additions and small fixes \
-(Bug fix) PUNCTUATION name typo; also now uses LongPtr instead of String.\
-(Bug fix) SPC_LINK had extra trailing _\
-(Bug fix) SpatialAudioObjectRenderStreamActivationParams2 missing packing alignment attrib\
-(Bug fix) `boolean` values on IDiscMasterProgressEvents::QueryCancel, IDiscMaster::RecordDisc, and IDiscRecorder::Erase should be Byte\
-(Bug fix) PSGetPropertyDescriptionByName definition incorrect\
-(Bug fix) IShellLibrary::ResolveFolder name typo\
-(Bug fix) PROP_CONTRACT_DELEGATE definition incorrect\
-(Bug fix) ICredentialProviderEvents::CredentialsChanged argument type incompatible with x64


**Update (v8.7.500, 28 Feb 2025):**\
-Added 100% coverage of msdelta.h\
-Added CompressedFolder coclass that creates an instance of the Zip Folder extension; replaces CoCreateInstance of {E88DCCE0-B7B3-11d1-A9F0-00AA0060FA31}.\
-Added STDIO_BUFFER and related flags per https://github.com/fafalone/WinDevLib/issues/37 request\
-Updated IStorage to make reserved and some other arguments optional for oleexp/olelib compatibility\
-Made last argument optional in a number of IEnum*::Next methods where already using [PreserveSig]\
-Misc API additions\
-NTDLL APIs now use `[UseGetLastError(False)]` since it's always n/a there; going forward I'll be adding this attribute as appropriate, but it will be a very long term project as documentation will need to be checked; can't just apply it to anything not returning BOOL.\
-(Bug fix) SysAllocString now uses DeclareWide

**Update (v8.7.498, 21 Feb 2025):**\
-For Property Get/Lets in TOM (RichEdit) interfaces, the actual typelib uses the more nature Property Get/Let Prop vs the SDK which uses GetProp/SetProp; the latter is more natural for BASIC so the names are being changed to that for usability and oleexp compatibility.\
-Added numerous missing tom* constants, including many undocumented ones for Office richedit.\
-Like previous tom* constants, I did my best to sort them into enums according to their usage, and all the TOM interfaces have been updated to make use of these.\
-(Bug fix) ITextRange2 missing GetProperty and SetText2 methods\
-(Bug fix) ITextFont2 missing SpaceExtension and UnderlinePositionMode prop get/lets.\
-(Bug fix) ITextDocument2::GetClientRect missing Type argument.\
-(Bug fix) ITextDocument2::GetEffectColor 2nd param is not retval\
-(Bug fix) ITextServices::TxDraw argument pfnContinue incorrect for x64

**Update (v8.7.496, 20 Feb 2025):**\
-Added missing functions from ole2.h; now 100% coverage\
-Added ATL helpers AtlPixelToHiMetric and AtlHiMetricToPixel (also as PixelToHiMetric and HiMetricToPixel)

**Update (v8.7.494, 19 Feb 2025):**\
-Added 100% coverage of lmuse.h/lmuseflg.h, lmrepl.h and lmat.h 

**Update (v8.7.493, 17 Feb 2025):**\
-Misc minor fixes/adjustments for upcoming project.\
**Update (v8.7.492, 17 Feb 2025):**\
-Misc minor fixes/adjustments for upcoming project.\
-(BREAKING CHANGE) CHARRANGE members renamed to their actual SDK names.\
-UIRibbonPropertyHelpers.h helper functions now manually validate PROPERTYKEY inputs.\
-Continued work to supply usable UUID types for documented CLSID_ constants for coclasses. 

**Update (v8.7.490, 11 Feb 2025):**\
-Added FolderShortcut coclass\
-Added mountmgr.h IOCTLs and structs; macros not covered.\
-Added OleTranslateColorPtr to allow the last parameter as ByVal VarPtr in x64\
-Additions to Direct3D 12 covering new stuff from SDK 10.0.22621 to 10.0.26000. Incomplete until next release.\
-Misc additions and fixes for upcoming project.\
-Continued work to supply usable UUID types for documented CLSID_ constants for coclasses. 

**Update (v8.7.487, 06 Feb 2025):**\
-Added min/max helper functions as they're in minwindef.h\
-Some missing constants for upcoming projects.\
-Added Ribbon helper functions from UIRibbonPropertyHelpers.h. Note that while functions are implemented, they don't have the PKEY type checking done by all the generic template stuff because of no language support, so it's up to the user to ensure the PKEY uses the correct type for the call.\
-(Bug fix) UI_HSB macro incorrect\
-(Bug fix) SI_NO_TREE_APPLY name typo


**Update (v8.7.486, 02 Feb 2025):**\
-Added realtimeapiset.h - 100% coverage\
-(Bug fix) FWPM_FILTER0 definition incorrect

**Update (v8.7.485, 16 Jan 2025):**\
-All delegate-using UDTs, APIs, and macros disabled again pending fix of tB issues:\
   https://github.com/twinbasic/twinbasic/issues/1999 Can't declare Delegates outside of module they're declared in for packages\
   https://github.com/twinbasic/twinbasic/issues/1890 Project compiler constants not applied to packages\
   https://discord.com/channels/927638153546829845/1293249305355747409 Delegates in interfaces in packages thoroughly broken\
-Additional bug fixes and improvements to RichEdit interfaces


**Update (v8.7.483, 11 Jan 2025):**\
-Began restoring delegates in API functions. By default these will generate a warning if you use LongPtr (or Long/LongLong). You can ignore these warnings through project settings or `[IgnoreWarnings(TB0026)]`.\
 You may also opt out of the use of delegates entirely by specifying the new compiler option WINDEVLIB_NO_DELEGATES = 1  (when fixed). Incomplete until next release.\
-Additions to Direct3D 12 covering new stuff from SDK 10.0.22621 to 10.0.26000. Incomplete until next release.\
-Added 100% coverage of msime.h, msimeapi.h\
-There's disagreement between sources for the names and arguments for several ITextHost2 members. They've been changed to match the Win10 SDK (10.0.22621) and Win11 SDK (10.0.26000).\
 This also applies to WinDevLibImpl.\
-(Bug fix) D3D12_VERSIONED_ROOT_SIGNATURE_DESC union member sizes incorrect; since all members had equivalent tB types they're now used in place of byte arrays.\
-(Bug fix) ITransferAdviseSink ByRef/ByVal and Long/LongPtr bugs\
-(Bug fix) IShellItemResources Long/LongPtr bug\
-(Bug fix) ITextHost::TxSetScrollPos, TxGetCharFormat, TxGetParaFormat definitions incorrect.\
-(Bug fix) RichEdit's SELCHANGE definition incorrect.\
-(Bug fix) Because they're mixed up in the SDK defs, some CHARFORMAT[2] dwMask values were in the dwEffects enum, and vice versa.

           
**Update (v8.7.480, 18 Dec 2024):**\
-Substantial additional winsock stuff; about 95% of winsock2.h/ws2def.h now covered; 33% of ws2tcpip.h\
   **REMINDER:** Due to their short genericnames, all Winsock APIs (ws2_32,dll) starting with a lower case letter are prefixed by ws_, e.g. ws_bind for bind.\
-Misc additions\
-(Bug fix) TOKEN_ALL_ACCESS, PROCESS_ALL_ACCESS values incorrect\
-(Bug fix) Many JOBOBJECTINFOCLASS values incorrect\
-(Bug fix) SHShowManageLibraryUI takes Unicode but used String without DeclareWide.\
-(Bug fix) Some constants for min/max values of types declared improperly or missing.\
-(Bug fix) WSAAsyncSelect definition incorrect for x64.


**Update (v8.6.476, 12 Dec 2024):**\
-Some additional winsock stuff\
-Some additional bluetooth stuff (including ws2bth.h, 100%)\
-Added all inlined functions from VersionHelpers.h (100% coverage)\
    IsWindowsVersionOrGreater has a optional custom argument, NoVersionLie, which returns the current Windows version regardless of manifest.\
-Some misc defs to bring coverage of minwinbase.h to 100%\
-(Bug fix) WideCharToMultiByte definition incorrect.

**Update (v8.6.474, 08 Dec 2024):**\
-Added Windows Remote Management APIs (wsman.h, 100% coverage)\
-Added Windows Connection Manager APIs (wcmapi.h, 100% coverage)\
-Added coverage of Netbios function (nb30.h, 100% coverage)\
-Added additional Windows Resource Protection APIs, including undocumented ones to list all protected files on Vista+.\
-VS_VERSIONINFO_FIXED_PORTION used 1-based arrays inconsistent with rest of project. Padding1 should not be an array.

**Update (v8.6.472, 26 Nov 2024):**\
-Added Group Policy APIs/interfaces from GPEdit.h (100% coverage)\
-Added InputPanelConfiguration.h 100% coverage\
-Added missing WIC interfaces, enums, and GUIDs.\
-Added missing event tracing related APIs and defs from wmistr.h and evntcons.h (now both 100% coverage)\
-(Bug fix) DeriveCapabilitySidsFromName typo in name; in kernelbase, not kernel32 or advapi32\
-(Bug fix) LsaConnectUntrusted, LsaInsertProtectedProcessAddress, LsaRemoveProtectedProcessAddress are in secur32, not advapi32.\
-(Bug fix) GetServiceRegistryStateKey, GetServiceDirectory, GetSharedServiceRegistryStateKey, GetSharedServiceDirectory are in sechost, not advapi32

**Update (v8.6.470, 18 Nov 2024):**\
-Large expansion of cfgmgr32.h APIs, now 100% coverage\
-Added 100% coverage of WinEFS.h\
-SHOpenFolderAndSelectItems will now use ByRef apidl As LongPtr in line with the official definition; if you previously used VarPtr you must either remove it or change to ByVal VarPtr.\
-Added 100% coverage of winstring.h\
-(Bug fix) MFP_GET_* functions improperly modified reference counts, leading to use-after-free crashes\
-(Bug fix) MSDN lists dialog macros as Sub (void); but the actual SDK macros would retain the return so they should be functions returning the result of the API they wrap.

           
**Update (v8.6.468, 11 Nov 2024):**\
-Added QoS APIs from qos2.h (100% coverage)\
-Added QoS Traffic APIs from traffic.h (100% coverage, also for qosobjs.h, qos.h, and qossp.h)\
-ServiceType custom enum renamed SystemServiceType to avoid conflict with official-named SERVICETYPE in QoS APIs\
-Added some additional Setup APIs

**Update (v8.6.466, 10 Nov 2024):**\
-Added Bluetooth LE APIs (bluetoothleapis.h, 100% coverage; bthledef.h 90% -- still need to do macros)\
-(Bug fix) FDI and FCI APIs and Delegates are _cdecl.

**Update (v8.6.464, 10 Nov 2024):**\
-Added Bluetooth APIs (bluetoothapis.h, 100% coverage; bthsdpdef.h 100%, bthdef.h 90% -- still need to do macros)\
-Added File History interfaces and APIs (fhcfg.h, fhsvcctl.h, fhstatus.h, fherrors.h 100%)\
-Added some undocumented APIs for immersive colors, dark mode, and SDR/HDR mode and brightness\
-Started medium term effort to supply usable UUID types for documented CLSID_ constants for coclasses. Covered wdShellCore and wdExplorer so far, the largest set, and also wdAccessible and wdBITS. Previously these weren't  provided because the objects could be created with the New keyword, but it's worthwhile to provide these for manual use with CoCreateInstance so other create options can be specified.
 
**Update (v8.5.462, 09 Nov 2024):**\
-All uses of delegates temporarily replaced with LongPtr pending backwards compatibility fix.

**Update (v8.5.461, 09 Nov 2024):**\
-Finished coverage of Windows Filtering Platform fwpmu.h (ipsectypes.h and iketypes.h now also 100%); also added IPSec errors.\
-Added Cabinet APIs (fdi_fcitypes.h, fdi.h, fci.h 100% coverage)\
-Additional callbacks declared as delegates\
-Additional work DNS API coverage\
-(Bug fix) StrFromTimeIntervalW missing ByVal, aliased version (StrFromTimeInterval) missing

**Update (v8.5.458, 26 Oct 2024):**\
-Added missing functions from handleapi.h (now 100% coverage)\
-Added private namespace api functions (namespaceapi.h, 100% coverage)\
-Misc winbase.h apis not added yet\
-(API Standards) [Global]AddAtom, FindAtom, GlobalFindAtom, [Global]GetAtomName used LongPtr instead of String\
-(Bug fix) CreatePipe ByVal/ByRef mixup. **IF YOU USED VARPTR AS A WORKAROUND MAKE SURE TO CHANGE IT!**\
-(WinDevLibImpl) Added IPerPropertyBrowsing, IOleControl

**Update (v8.5.456, 20 Oct 2024):**\
-Changed C-style buffered name args in file info UDTs to use MAX_PATH - 1 instead of MAX_PATH to eliminate excess padding to simplify operations on buffers full of them.\
-Fixed MagSetWindowSource misleading argument names.\
-Added undocumented antialiasing APIs for magnification.dll\
-(Bug fix) FILE_RENAME_INFO definition incorrect

**Update (v8.5.454, 15 Oct 2024):**\
**twinBASIC Beta 617 or newer is now required!**\
-I've begun replacing specifically defined callbacks with Delegate function pointers. These will allow you to,
 like C/C++, see the prototype for the function you implement for it.\
 This will not break existing code, however it may generate a warning about implicit conversion to a Delegate
 if you use a Long(Ptr) variable. You can change the type to the Delegate, or disable the warning.\
 This will be an ongoing process and only a small percentage are completed in this initial update.\
-Turns out several of us forgot variadic functions are actually supported (in user mode at least)... so now
 AuthzReportSecurityEvent, ShellMessageBox[A,W], and DbgPrint use the proper ByVal ParamArray vargs As Any()
 syntax to support it. These are all ByVal so pass ByRefs as ByVal VarPtr() etc.\
-Added DXVA2 monitor APIs (physicalmonitorenumerationapi.h, highlevelmonitorconfigurationapi.h, and lowlevelmonitorconfigurationapi.h; 100% coverage)\
-Added missing inlined APIs from evntcons.h\
-Added missing winuser.h functions wsprintf/wsvprintf and related.\
-(Bug fix) DXVA2CreateDirect3DDeviceManager9 typo in name.\
-(Bug fix) GdipEnumerateMetafile* API definition issues\
-(Bug fix) GDIP APIs with invalid Optional ByRef As Any arguments\
-(Bug fix) RtlCrc64 definition incorrect.

**Update (v8.5.451, 04 Oct 2024):**\
-CryptProtectMemory and CryptUnProtectMemory in crypt32 are just forwarders; these now point directly at their targets in dpapi.\
-(Bug fix) WindowsCreateString[Reference] definitions incorrect.

**Update (v8.5.450, 03 Oct 2024):**\
***NOTE:*** These bug fixes were identified through scanning for the actual entry points in DLLs.\
            About 75% of these bugs are errors in MSDN documentation or the SDK itself.\
-Removed some -A variants of functions that do not exist (many erroneously documented by MSDN or the SDK)\
-(Bug fix) MapViewOfFile2, LookupAccountSidLocal[A,W] is an inline macro, not dll export.\
-(Bug fix) DisconnectWindowsDialog name typo, also exported by ordinal only\
-(Bug fix) PssCaptureSnapshot, CreateCursor, CreateDIBPatternBrushPt, PropVariantToUInt16Vector, SetupDiSetDeviceRegistryPropertyW, CM_Query_And_Remove_SubTree[A,W], AddPrinterDriverExA, ShowHideMenuCtl, GetThemeFilename, BCryptProcessMultiOperations, MFDeserializeAttributesFromStream, GdipPathIterNextSubpathPath, GdipSetImageAttributesNoOp, GetComputerObjectName[A,W] name typos\
-(Bug fix) TabbedTextOut[A,W] is in user32, not gdi32\
-(Bug fix) FreePrintPropertyValue is in spoolss.dll, not winspool.drv\
-(Bug fix) GetListBoxInfo is in user32, not comctl32\
-(Bug fix) ImageList_CoCreateInstance dll name typo\
-(Bug fix) CryptProtectDataNoUI, CryptUnprotectDataNoUI are in dpapi, not crypt32\
-(Bug fix) MFCreateAVIMediaSink, MFCreateWAVEMediaSink are in mfsrcsnk.dll, not mf.dll\
-(Bug fix) CryptRetrieveObjectByUrl[A,W], CryptInstallCancelRetrieval, CryptUninstallCancelRetrieval, CryptCancelAsyncRetrieval, CryptGetObjectUrl, CryptGetTimeValidObject, CryptFlushTimeValidObject are in cryptnet, not crypt32\
-(Bug fix) CredPackAuthenticationBuffer[A,W], CredUnPackAuthenticationBuffer[A,W] are in credui, not advapi32\
-(Bug fix) [Un]SubscribeServiceChangeNotifications, LsaLookupOpenLocalPolicy, LsaLookupClose, LsaLookupTranslateSids, LsaLookupTranslateNames, LsaLookupGetDomainInfo, OpenTraceFrom*, ProcessTraceBufferIncrementReference, ProcessTraceBufferDecrementReference, ProcessTraceAddBufferToBufferStream, CredProtectEx, CredUnprotectEx, are in sechost, not advapi32\
-(Bug fix) CheckTokenCapability, DeriveCapabilitySidsFromNamed, GetAppContainerAce are kernel32, not advapi32\
-(Bug fix) QueryContextAttributesEx[A,W], QueryCredentialsAttributes[A,W] are in sspicli, not secur32\
-(Bug fix) LsaRegisterLogonProcess, LsaDeregisterLogonProcess, LsaLogonUser, LsaLookupAuthenticationPackage, LsaCallAuthenticationPackage, LsaFreeReturnBuffer, LsaEnumerateLogonSessions, LsaGetLogonSessionData, LsaRegisterPolicyChangeNotification, LsaUnregisterPolicyChangeNotification, LsaConnectUntrusted, are in secur32, not advapi32\
-(Bug fix) CreateRestrictedToken is in advapi32, not kernel32\
-(Bug fix) RegisterWindowMessage, SHCreateDirectoryEx, GetCPInfoEx, GetStartupInfo, FindText, ReplaceText, GetIconInfoEx, DrawText, EnumICMProfiles, HttpSendRequest, ChangeAccountPassword missing aliases\
-(Bug fix) TerminateProcessOnMemoryExhaustion, GetIntegratedDisplaySize, GetOsManufacturingMode, LoadStringByReference, VirtualAlloc2, SetProcessValidCallTargets[ForMappedView], QueryVirtualMemoryInformation, LoadEnclaveImage[A,W], CallEnclave, TerminateEnclave, DeleteEnclave, EncodeRemotePointer, DecodeRemotePointer, MapViewOfFileNuma2, MapViewOfFile3, UnmapViewOfFile2, SetSystemTimeAdjustmentPrecise, GetSystemTimeAdjustmentPrecise, ImpersonateNamedPipeClient, OpenCommPort, GetCommPorts are in kernelbase, not kernel32.\
-(Bug fix) GetPerformanceInfo, GetProcessMemoryInfo, InitializeProcessForWsWatch, GetWsChanges[Ex], QueryWorkingSet[Ex], GetModuleFileNameEx[A,W], GetProcessImageFileName[A,W], EnumProcesses, EnumProcessModules[Ex], EnumPageFiles[A,W], EnumDeviceDrivers, GetDeviceDriverBaseName[A,W], GetMappedFileName[A,W], GetModuleBaseName[A,W], GetModuleInformation, are in psapi, not kernel32\
-(Bug fix) SHRunControlPanel, SHOpenPropSheetA, SHStartNetConnectionDialogA, RunFileDlg, SHCreateFilter, CheckDiskSpace, CopyStreamUI, CreateInfoTipFromItem[2], GetAppPathFromLink, IsElevationRequired, IsSearchEnabled, PathGetPathDisplayName, SHGetUserPicturePath[Ex], SHSetUserPicturePath, PathUnExpandEnvStringsForUser[A,W], AssocGetUrlAction, SHCreateStreamOnDllResource[W], SHCreateStreamOnModuleResource[W], SHAreIconsEqual, SHGlobalCounterGetValue, SHGlobalCounterIncrement, SHGlobalCounterDecrement, ImageList_SetColorTable exported by ordinal only

           
**Update (v8.4.446, 26 Sep 2024):**\
-WebView2 definitions updated to match current stable release 1.0.2792.45\
-(Experimental) InterlockedIncrement, InterlockedDecrement and InterlockedExchange are now inline assembly via Emit() 
                instead of in a static library.\
                twinBASIC Beta 606 or newer is required for this; using the new TWINBASIC_BUILD compiler constant,
                this feature is only enabled if supported and older versions use the static library version.
-MIXERLINECONTROLS[A,W].dwControlType name changed to dwControlTypeOrID to properly indicate it's a union that can take either.\
-Cleared new compiler warnings to maintain strict mode compliance.\
-(Bug fix) Several String constants still had escaped slashes (\\), which in VBx and tB incorrectly produced both.\
-(Bug fix) PropSheet_ShowWizButtons macro incorrect.\
-(Bug fix) MIXERCONTROL[A,W] missing terminating reserved Long, so LenB would be incorrect.\
-(Bug fix) ICoreWebView2Profile7 missing method (also breaking ICoreWebView2Profile8)\
-TreeView_GetItemRect did not appear to be correct; it may or may not fixed now... it's one of those ridiculous 
   pointer messes like *(*(HTREEITEM))prc where the lParam is used for both item handle and RECT.\
   I won't call the bug fixed until some thorough testing.
   
**Update (v8.3.444, 12 Sep 2024):**\
-Added some missing netapi32 APIs from lmaccess.h\
-LPWSTRToStr now sets the pointer to zero when fFree = True to prevent use-after-free crashes.\
-Added missing ENDPOINT_HARDWARE_SUPPORT_* values for IAudioEndpointVolume::QueryHardwareSupport\
-Buffered AUDIO_VOLUME_NOTIFICATION_DATA for 128 channels instead of 2\
-Added some missing oaidl.idl types.\
-There was no reason ITypeFactory should extend IUnknownUnrestricted instead of IUnknown\
-(API Standards) FindFirstVolume[A], FindNextVolume[A] used LongPtr instead of String.\
-(Bug fix) FindNextVolume[A,W] incorrect return type (only impacted x64).\
-(Bug fix) IOwnerDataCallback.SetItemPosition takes a ByVal POINT, not ByRef (temp substitution of LongLong used pending proper support for ByVal UDTs)\
-(Bug fix) LookupAccountName[A,W] ByVal/ByRef mixup.\
-(Bug fix) LookupAccountSidLocal, ConvertStringSidToSid, FindFirstStreamTransacted, GetModuleHandleEx, GlobalGetAtomName, GetDiskFreeSpaceEx, GetSystemDirectory, 
           GetStringTypeEx, GetTempPath2, EnumPropsEx, RegOpenKeyTransacted, RegConnectRegistryEx, SetupDiGetDeviceInterfacePropertyKeys, SetDefaultPrinter, 
           AddPrinterDriverEx, DeletePrinterDriverEx, GetPrinterDriver2, DlgDirSelectEx, InternetGetPerSiteCookieDecision incorrect aliases.\
           I created a routine to scan for this class of error, so hopefully this kind of mistake should be eliminated now.\
-(Bug fix) FindFirstFileExTransacted should be FindFirstFileTransacted.\
-(Bug fix) ExpandEnvironmentStringForUser should be ExpandEnvironmentStringsForUser.\
-(Bug fix) RegRenameKey does not have A/W variants, only Unicode; these were removed, but this function is now overloaded to allow either String or LongPtr.\
IMPORTANT: THIS MAY REQUIRE CODE CHANGES. If you use any of the following and used the workaround of VarPtr(), the VarPtr must now be removed:\
-(Bug fix) ITypeComp::Bind last param should be ByRef BINDPTR.


**Update (v8.3.442, 2 Sep 2024):**\
-Added missing explicit A/W versions of [Get,Set]WindowLongPtr[A,W] and [Get,Set]ClassLongPtr[A,W].\
  Put those and also moved the aliased versions to the Win64 block as they're not exported from the 32bit user32.dll\
-Added interface IFileOperation2 (Win10RS4+).\
-(API Standards) GetCommandStringFlags (GCS_* values) used ANSI as the unmarked (not -A or -W) version.\
-(Bug fix) IPropertyBag2::Read/Write last args should be ByRef.\
-(Bug fix) GetCharacterPlacement alias typo\
-(Bug fix) COPYFILE2_MESSAGE union placeholder size incorrect for x64. Also renamed 'union' to 'Info', the name of the union.

**Update (v8.3.440, 27 Aug 2024):**\
-Misc shell32 and kernel32 API additions.\
-SHELLFLAGSTATE was only for use to hold settings; not for use with API.\
  This version has been renamed SHELLFLAGSTATEFlags and SHELLFLAGSTATE is now just a single Long
  representing the bitfield suitable for use with SHGetSettings.\
-(Bug fix) ID3D11DeviceContext::ClearRenderTargetView/ClearUnorderedAccessViewUint/ClearUnorderedAccessViewFloat, ID3D12GraphicsCommandList::ClearRenderTargetView definitions incorrect.\
-(Bug fix) ReadDirectoryChangesA does not exist\
-(Bug fix) SHGetSettings definition incorrect.\
-(Bug fix) SHChangeNotifyEntry missing packing alignment, leading to wrong size


**Update (v8.3.439, 21 Aug 2024):**\
-(Bug fix) While checking BOOL was used where appropriate in MediaFoundation, numerous ByVal args that should be ByRef were uncovered...\
           IMPORTANT: THIS MAY REQUIRE CODE CHANGES. If you use any of the following and used the workaround of VarPtr(), the VarPtr must now be removed:\
           IMF2DBuffer::IsContiguousFormat, IMFContentEnabler::IsAutomaticSupported, IMFByteStreamCacheControl2::IsBackgroundTransferActive, IMFByteStreamTimeSeek::IsTimeSeekSupported, 
           IMFNetCredential::LoggedOnUser, IMFSSLCertificateManager::GetCertificatePolicy, IMFTrustedOutput::IsFinal, IMFVideoDisplayControl::GetFullscreen, IMFPMediaPlayer::GetMute, 
           IMFRateControl::GetRate, IMFPMediaItem::Has(Audio,Video),GetStreamSelection, IMFMediaEngineEx::GetRealTimeMode,IsProtected, IMFHDCPStatus::Query, IMFMediaEngineOPMInfo::GetOPMInfo, 
           IMFMediaEngineClassFactoryEx::IsTypeSupported, IMFMediaEngineSupportsSourceTransfer::ShouldTransferSource, IMFMediaKeySession2::Load, IMFNetCrossOriginSupport::IsSameOrigin,GetSourceOrigin,
           IMFHttpDownloadRequest::HasNullSourceOrigin,QueryHeader,GetUrl,GetAtEndOfPayload, IMFSensorProfile::IsMediaTypeSupported, IMFSensorProcessActivity::GetStreamingState, 
           IMediaBuffer::GetBufferAndLength, IMFContentEnabler::GetEnableURL,GetEnableData, IMFMetadata::GetLanguage, IMFByteStreamCacheControl2::GetByteRanges, IMFOutputTrustAuthority::SetPolicy, 
           IMFSecureChannel::GetCertificate, IMFSampleProtection::GetProtectionCertificate, IMFSAMIStyle::GetSelectedStyle, IMFSystemId::Setup, IMFMediaEngineAudioEndpointId::GetAudioEndpointId, 
           IMFMediaEngineClassFactory3::CreateMediaKeySystemAccess, IMFExtendedCameraControl::LockPayload, MFEnumDeviceSources, MFSerializePresentationDescriptor, IMFSimpleAudioVolume::GetMute, 
           MFIsContentProtectionDeviceSupported, IAMAsyncReaderTimestampScaling::GetTimestampMode, IAMAudioInputMixer::get_Enable,Mono,Loudness, IUri::IsEqual,HasProperty, 
           IAppVisibility::IsLauncherVisible, IDataObjectAsyncCapability::GetAsyncMode,InOperation, IApplicationAssociationRegistration::QueryAppIsDefault[All], IDCompositionDevice::CheckDeviceState, 
           IOpLockStatus::IsOplockValid,IsOplockBroken, ISearchCrawlScopeManager::IncludedInCrawlScopeEx, ISearchViewChangedSink::OnChange, IInternetSecurityManagerEx2::QueryCustomPolicyEx2,
           WinHttpOpenRequest, ID3D11DeviceContext[1]::(numerous), ID3D11On12Device::ReleaseWrappedResources,AcquireWrappedResources, ID3D12VersionedRootSignatureDeserializer::GetRootSignatureDescAtVersion,
           ID3D12GraphicsCommandList::SetDescriptorHeaps, ID3D12CommandQueue::ExecuteCommandLists, ID3D12Device::MakeResident, ID3D12Device::Evict, ID3D12Device1::SetEventOnMultipleFenceCompletion,SetResidencyPriority
           UiaNodeFromHandle, UiaNodeFromProvider, UiaGetRootNode, UiaHUiaNodeFromVariant, UiaHPatternObjectFromVariant, UiaHTextRangeFromVariant, UiaGetPatternProvider, UiaAddEvent 
           
-(Bug fix) IMFSampleProtection::InitOutputProtection ppbSeed should be ByRef LongPtr.\
-(Bug fix) IMFSourceReader::SetCurrentMediaType dwReserved should be ByVal LongPtr.\
-(Bug fix) ID3D11DeviceContext::ClearRenderTargetView/ClearUnorderedAccessViewUint/ClearUnorderedAccessViewFloat use [in] type var[4]; which shouldn't be a safearray.\
            Used best guess for workaround; see https://github.com/twinbasic/twinbasic/issues/1892.\
-(Bug fix) ID3D12VersionedRootSignatureDeserializer::GetUnconvertedRootSignatureDesc and ID3D12RootSignatureDeserializer::GetRootSignatureDesc should return LongPtr.


**Update (v8.3.437, 20 Aug 2024):**\
-Added Native Registry APIs (ntregapi.h, 90%)\
-(Bug fix) WOW64_LDT_ENTRY duplicate type (Issue #32)\
-(API Standards) SHUpdateImage[A] used LongPtr instead of String; added overloads for standards due to common use of pidls instead. (Issue #33)


**Update (v8.3.430, 01 Jul 2024):**\
-Added HID APIs (hidclass.h, hidusage.h, hidpi.h, hidsdi.h 100%; HidD_ and HidP_ APIs in hid.dll)\
-Added WinML interfaces (WinML.h, 100%)\
-Added some additional APIs from sysinfoapi.h to bring coverage to 100%\
-Added Common Dialog extended error codes from cderr.h (100% coverage)\
-New helper function VarTypeEx returns the VarType without filtering flags like VT_BYREF.\
-WinDevLib is now strict mode compliant\
-(Bug fix) V_ISBYREF, V_ISARRAY, and V_ISVECTOR helper functions relied on VarType which filtered those flags.\
-(Bug fix) DispatchMessage[A,W], SendNotifyMessage[A,W] return types incorrect for x64.\
-(Bug fix) IMFVideoDisplayControl.GetCurrentImage second argument ByVal/ByRef mixup.\
-(Bug fix) ListView_SetItemText macro incorrect.\
-(Bug fix) SHSaveLibraryInFolderPath type mismatch.\
-Note: ShellScalingApi.h was verified to be 100% covered.

**Update (v8.3.428, 13 Jun 2024):**\
-Some additional system info structs to support upcoming project\
-PRIVILEGE_SET and TOKEN_PRIVILEGES were intended to be buffered to the max number of privileges, but that was set too low; it's now 45.\
-SE_DELEGATE_SESSION_USER_IMPERSONATE_NAME was missing.\
-(Bug fix) MAXSIZE_T only defined for 64bit

**Update (v8.3.426, 10 Jun 2024):**\
-Completed imagehlp.h/dbghelp.h API coverage, now 100%\
   Note: I've tried to implement the unusual alias struct in the header files as faithfully as possible, and a great many of these APIs do have aliases, so always consult the SDK source and wdAPIDbgHlp.twin in addition to MSDN-- MSDN covers only actual entry points.\
-Added a large number of overloads for compatibility with oleexp.tlb APIs that use `[PreserveSig(False)]` to rewrite a last [out] parameter as the return value. This is for compatibility only and will not be expanded beyond oleexp APIs using it.\
  **IMPORTANT:** Due to this change, WinDevLib now requires twinBASIC Beta 553 or newer.\
-Updated DirectML for recent additions (feature set >= 0x6000)\
-Added ITipAutoCompleteProvider, ITipAutoCompleteClient, and coclass TipAutoCompleteClient\
-Added IObjectWithPackageFullName\
-Added coverage of interlockedapi.h (100%)\
-Some additional system info structs\
-(Bug fix) MkParseDisplayName should not use ANSI conversion.\
-(Bug fix) MFCreateADTMediaSink should be MFCreateADTSMediaSink\
-(Bug fix) IMFMediaType.GetMajorType, IQueueCommand methods used stdole.GUID instead of UUID, leading to automation type incompatible errors.\
-(Bug fix) IMFMediaEngineEx.GetVideoSwapchainHandle Long instead of LongPtr.\
-(Bug fix) SLIST_HEADER definition incorrect.


**Update (v8.2.424, 06 Jun 2024):**\
-Added INATExternalIPAddressCallback for use with NATUPnP Type Library v1.0 (NATUPNPLib, included with Windows)\
-Removed LOWORD(LongLong) and HIWORD(LongLong) overloads due to too many circumstances with ambiguity errors.\
-(Bug fix) GetAdaptersAddresses returns variable length data, not a single UDT.

**Update (v8.2.423, 04 Jun 2024):**\
-Added UPnP interfaces (upnp.h, upnphost.h, 100%)\
-Added Real-time Work Queue (RTWorkQ.h) APIs and interfaces\
-(Bug fix) WSAStartup used Integer instead of Long for first arg\
-(Bug fix) RtlIpv4StringToAddressEx[A,W] arg 'Strict' should be ByVal\
-(Bug fix) IP_ADDRESS_STRING/IP_MASK_STRING and several downstream types definitions incorrect (+/* typo)\
-(Bug fix) GetAdaptersInfo returns variable length data, not a single UDT.


**Update (v8.2.413, 02 Jun 2024):**\
-Missing common winmm time APIs timeSetEvent/timeKillEvent and related consts\
**Update (v8.2.412, 02 Jun 2024):**\
-Added Direct3D 9 interfaces (base interfaces courtesy of The trick's Dx9vb type library); d3d9.h, d3d9types.h, d3d9caps.h, d3dx9shader.h\
-Added DXVA2, DXVA-HD, and EVR9 interfaces (evr9.h, dxva2api.h, dxvahd.h)\
-Added Native WiFi APIs (wlanapi.h, 100%, see wdAPIWLAN.twin for dependent header coverage details)\
-Coverage of oleexp's oledlg.inc was entirely missing; added and expanded to include 98% of oledlg.h (currently unsupported vararg APIs pending)\
-(Bug fix) Numerous incorrect constants due to << overflowing to zero after exceeding Integer.\
-(Bug fix) IBackgroundCopyJob2.GetReplyFileName, IBackgroundCopyJobHttpOptions.GetCertificateName used String for LPWSTR*

**WinDevLibImpl v1.3.20:** Add Implements-compatible IOleUILinkContainer


**Update (v8.1.409, 25 May 2024):**\
-(Bug fix) GDI+ enum values incorrect in PixelOffsetMode, EncoderParameterValueType, SmoothingMode, InterpolationMode, MetafileFrameUnit, and CompositingQuality.\
**Update (v8.1.408, 25 May 2024):**\
-Began coverage of the Windows Filtering Platform. Initially, enough is declared to set up basic filters, like blocking all traffic from a given process.\
-IShellItem2.GetCLSID now uses standardized UUID type instead of UUID.\
-Add missing GDI+ startup output and inputex structs and enums\
-(Bug fix) GdiplusStartupInput definition incorrect (did not cause runtime errors because size was > minimum, but optional args wouldn't work on x64)\
-(Bug fix) IAudioEndpointOffloadStreamMute method arg types incorrect (but likely was harmless)\
-(Bug fix) Switch imagehlp to dbghelp in identical parts of editor; DLL exports are not identical. Note: Dbghelp APIs are a work in progress; 40% done.\
-(Bug fix) MFMEDIASOURCE_CHARACTERISTICS, MF_SOURCE_READER_FLAG, and MF_SOURCE_READER_CONTROL_FLAG enums all values incorrect.


**Update (v8.0.406, 17 May 2024):**\
-(Bug fix) Numerous String/LongPtr bugs and standards issues; see Issue #30.

**Update (v8.0.405, 17 May 2024):**\
-(Bug fix) IShellImageDataFactory method names incorrect\
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

---

For earlier version history, see CHANGELOG.md

For more information and a list of available interfaces, visit the [VB Forums thread](https://www.vbforums.com/showthread.php?897883-twinBASIC-tbShellLib-Shell-Interface-Library-(x64-compatible-successor-to-oleexp)) for this project.
