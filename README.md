# WinDevLib 
## Windows Development Library for twinBASIC

**Current Version: 8.12.552 (June 6th, 2025)**

(c) 2022-2025 Jon Johnson (fafalone)

WinDevLib is a project to supply Windows API COM interfaces and DLL declares in a format consumable by twinBASIC. This involves not only writing the definitions, but using tB compatible types-- so in some cases, even though there may be an existing way to import references to interfaces, they may be unusable due to e.g. the use of unsigned types, C-style arrays, double pointers, etc. In most cases these definitions are also compatible with VBA7, and with minor adjustments VB6; where they're not it's usually minor syntax adjustments, so this is also a great resource for APIs for those, covering vastly more than other other similar project.

Included are definitions of 2800+ common COM interfaces and 8500+ APIs from all the common system modules. This makes it similar to working in C++ with `#include <Windows.h>` and a number of other headers for commonly used features. These have all been redone by hand from the original headers, in order to restore 64bit type info lost in VB6 versions, avoid the errors of automated conversion tools (e.g. Win32API_PtrSafe.txt is riddled with errors), and make them friendlier by converting groups of constants associated with a variable into an Enum so it comes up in Intellisense. This takes advantage of tB's ability to provide Intellisense for types besides Long in API defs (hopefully UDTs soon, this project has provisioning for that). 

This project is implemented purely in tB native code, as unlike VB6 there's language support for defining interfaces and coclasses. As a twinPACKAGE, regular code is supported in addition to the definitions, so some content normally found in regular addin modules have been built in (like you'd find in oleexp's mIID.bas, mPKEY.bas, etc, and helper functions). Does it still make sense to use a project like this when interfaces can be defined in-language? I'd say yes, because for a large number of interfaces, there's deep dependency chains with other interfaces and the types they rely on. It makes more sense to drop this in and be done with it than constantly have to define the interfaces you want and then stubs for their dependencies, especially when you might need those later on. This project is even more useful now with the API coverage; it should cover about 99% of your needs for core system DLLs. 

This project also serves a comprehensive twinBASIC replacement for [oleexp.tlb](http://www.vbforums.com/showthread.php?786079-VB6-Modern-Shell-Interface-Type-Library-oleexp-tlb), my Modern Shell Interfaces Type Library project for VB6. 100% of the content is covered with little to no change (just String arguments in some places due to differences between how they're handled in typelibs). 

Please report any bugs via the Issues feature here on GitHub.

### Requirements

[twinBASIC Beta 617 or newer](https://github.com/twinbasic/twinbasic/releases) is required.

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

`WINDEVLIB_COMCTL_LIB_DEFINED` - You can use this flag if you already have an alternative common controls definition set, e.g. tbComCtlLib; it will disable wdAPIComCtl. (Note: WinDevLib has more complete comctl defs than tbComCtlLib, as that project was deprecated and not updated).

`WINDEVLIB_DLGSH` - This enabled constants from dlg.h. These are extremely uncommon to use, and have very short, generic names likely to cause conflicts, so they're opt-in.

`WINDEVLIB_NOQUADLI` - Restores the old `LARGE_INTEGER` definition of lo/high Long values.

>[!WARNING]
>The `WINDEVLIB_NOQUADLI` constant will break alignment on numerous Types; most only on x64, but some on both. 

`WINDEVLIB_AVOID_INTRINSICS` - Uses the `Interlocked*` APIs that are exported from kernel32.dll (32bit mode only) instead of the static library containing compiler intrinsic versions of those in addition to all the ones not exported and all the 64bit ones.

`WINDEVLIB_NOLIBS` - Fully exclude static libraries (currently only Interlocked); mainly intended for comparing current tB versions to Beta 423 where the `Import Library` syntax is not yet supported.

`WINDEVLIB_NO_DELEGATES` - Do not use Delegate functions in place of function pointers.

`WINDEVLIB_XAUDIO8` - Use XAudio8 DLLs for XAudio2 APIs (Windows 8)

`WINDEVLIB_NOMATH` - Exclude built in math helper function (see below). Note: XAudio2 inlined helper functions unavailable when math disabled.

>[!IMPORTANT]
>Currently flags are not inherited from the main project, so the only way to use these is to set them in the compiler flags for WinDevLib.twinproj then build a custom twinpack.

#### Custom Helper Functions
In addition to coverage of common Windows SDK-defined macros and inlined functions, a small number of custom helper functions are provided to deal with Windows data types and similar not properly supported by the language. These are:

`Public Function GetMem(Of T)(ByVal ptr As LongPtr) As T` - A generic to dereference a pointer into any type. The native `CType(Of )` allows dereferencing to UDTs, but this helper allows instrinsic types in addition to UDTs, and is used the same way.

`Public Function DCast(Of T, T2)(ByVal v As T2) As T` - Direct Cast: Copies the data of v into any type, without modification, so no overflows, and possible to e.g. go from `LongLong` to `POINT`, with `Dim pt As POINT = DCast(Of POINT)(SomeLongLong)`

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

The goal of the API coverage in WinDevLib is to provide the kind of programming experience you'd get in C/C++ by including windows.h and many of the more common feature sets like DirectX and GDIPlus. It currently includes about 5,500 APIs. But even that is just scratching the surface of the total Windows API set. Due to the low quality of automated conversion, even by Microsoft themselves (see: Win32API_PtrSafe.txt), I'm not interested in simply feeding headers through a conversion utility or using a database, so instead WinDevLib will be focused on the most commonly used features in the major system DLLs, though less common ones can be added by request or as time goes on and the existing DLLs are completed. I do not intend to include most native APIs that have fully equivalent regular APIs; that's basically doubling the work for no benefit-- but if they offer additional features or substantially improved performance, they will be included. 

I've included the definitions, associated types, and associated constants, for extensive portions of the following modules: shell32.dll, shlwapi.dll, kernel32.dll, ktmw32.dll, user32.dll, advapi32.dll, tdh.dll, authz.dll, crypt32.dll, wintrust.dll, bcrypt.dll, ncrypt.dll, cryptui.dll, ole32.dll, oleaut32.dll, propsys.dll, gdi32.dll, gdiplus.dll, virtdisk.dll, userenv.dll, dbghelp.dll, mpr.dll, iphlpapi.dll, urlmon.dll, hlink.dll, winmm.dll, cfgmgr32.dll, setupapi.dll, comctl32.dll, dwm.dll/uxtheme.dll, comdlg32.dll, winbio.dll, winspool.drv, imm32.dll, hid.dll, cldapi.dll, pdh.dll, powrprof.dll, wtsapi32.dll, and netapi32.dll. Please let me know any I've missed from these.\
Limited coverage (or full coverage of very small sets) is provided for ntdll.dll, version.dll, msimg32.dll, fwpuclnt.dll, sxs.dll, secur32.dll, msacm32.dll, url.dll, htmlhelp.dll, avifil32.dll, and ws2_32.dll. If you feel any missing ones from these should be included, or would like to contribute more, let me know.\
Finally, there's numerous additional API sets from small to large for independent Windows features. These include small sets like restartmgr.dll through very large sets like the various Media Foundation and DirectX DLLs. In the future I'll better organize coverage lists, but the bottom line is let me know if any common APIs or built in API sets for components should be added. TODO.md in the WDL project files contains ones planned but not yet done.

**Future coverage:** In the future I'm planning to expand native APIs with no equivalents, add additional Winsock coverage, and add OpenGL-- though for this last one I may wait for tB to have `Alias` support since existing OpenGL codebases make heavy use of them by way of NeHe's TLB. I welcome contributions of any of these. If you've done the consts->enums conversions already, I'd even take 32bit-only versions.


### ***tbShellLib is now WinDevLib - Windows Development Library for twinBASIC***
This project has grown well beyond it's original mission of shell programming. While that's still the largest single part, it's no longer a majority of the code, and the name change now much better reflects the purpose of providing a general Windows API experience like windows.h. Compiler constants and module names/file names have been updated to reflect the name change. tbShellLibImpl is now WinDevLibImpl. There are also some major chanages associated with this update, please see the full changelog below.

### Updates

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
