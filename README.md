# WinDevLib 
## Windows Development Library for twinBASIC

**Current Version: 9.3.692 (June 12th, 2026)**

(c) 2022-2026 Jon Johnson (fafalone)

> [!IMPORTANT]
> **Version 9.2.634 and higher now requires twinBASIC Beta 923 or newer.** The project is now using twinBASIC's new `Alias` syntax support, which is impractical to version-gate. 954+ is required for OpenGL.

WinDevLib is a project to make all common Windows API COM interfaces, DLL declares, and related Types/Enums/Consts available while programming in twinBASIC.\
Included are definitions of 3800+ common COM interfaces and 15,000+ APIs from all the common system modules, a level of coverage which makes WDL an entirely different experience than any VBx library, the largest of which offer at most 1/10th as much with huge gaps.\
This makes working with WDL similar to working in C++ with `#include <Windows.h>` and a number of other headers for commonly used features. These have all been redone by hand from the original headers, in order to restore 64bit type info lost in VB6 versions, avoid the errors of automated conversion tools (e.g. Win32API_PtrSafe.txt is riddled with errors), and make them friendlier by converting groups of constants associated with a variable into an Enum so it comes up in Intellisense. This takes advantage of tB's ability to provide Intellisense for types besides Long in API defs (hopefully UDTs soon, this project has provisioning for that). 

Creating this involves not only writing the definitions, but using tB compatible types-- so in some cases, even though there may be an existing way to import references to interfaces, they may be unusable due to e.g. the use of unsigned types, C-style arrays, double pointers, etc. In most cases these definitions are also compatible with VBA7, and with minor adjustments VB6; where they're not it's usually minor syntax adjustments, so this is also a great resource for APIs for those, covering vastly more than other other similar project.

This project is implemented purely in tB native code, as unlike VB6 there's language support for defining interfaces and coclasses. As a twinPACKAGE, regular code is supported in addition to the definitions, so some content normally found in regular modules have been built in (like you'd find in oleexp.tlb's mIID.bas, mPKEY.bas, etc, and helper functions). Does it still make sense to use a project like this when interfaces can be defined in-language? I'd say yes, because for a large number of interfaces, there's deep dependency chains with other interfaces and the types they rely on. It makes more sense to drop this in and be done with it than constantly have to define the interfaces you want and then stubs for their dependencies, especially when you might need those later on. This project is even more useful now with the API coverage; it should cover about 99% of your needs for core system DLLs. 

This project also serves a comprehensive twinBASIC replacement for [oleexp.tlb](http://www.vbforums.com/showthread.php?786079-VB6-Modern-Shell-Interface-Type-Library-oleexp-tlb), my Modern Shell Interfaces Type Library project for VB6. 100% of the content is covered with little to no change (just String arguments in some places due to differences between how they're handled in typelibs). 

Please report any bugs via the Issues feature here on GitHub.

### Requirements

[twinBASIC Beta 923 or newer](https://github.com/twinbasic/twinbasic/releases) is required.

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

`WDL_NO_DIRECTX` - Excludes DirectX, Media Foundation, XAudio, and WinML content. This is useful to substantially cut down on Intellisense entries in non-multimedia apps. Basic 2D graphics remain (GDI, GDI+, WIC).\
`WDL_NO_GL` - Excludes OpenGL.

`WDL_NO_BYVAL_UDT` - Do not use ByVal UDT arguments in interfaces/APIs. Reverts to the previous workarounds from VBx like `LongLong` for `POINT` and separate 32/64 bit defs for GUIDs. Useful for VBx code compatibility.

`WDL_NO_COMCTL` - You can use this flag if you already have an alternative common controls definition set, e.g. tbComCtlLib; it will disable wdAPIComCtl. (Note: WinDevLib has more complete comctl defs than tbComCtlLib, as that project was deprecated and not updated).

`WDL_DLGSH` - This enabled constants from dlg.h. These are extremely uncommon to use, and have very short, generic names likely to cause conflicts, so they're opt-in.

`WDL_NOQUADLI` - Restores the old `LARGE_INTEGER` definition of lo/high Long values.

>[!WARNING]
>The `WDL_NOQUADLI` constant will break alignment on numerous Types; most only on x64, but some on both. 

`WDL_AVOID_INTRINSICS` - Uses the `Interlocked*` APIs that are exported from kernel32.dll (32bit mode only) instead of the static library containing compiler intrinsic versions of those in addition to all the ones not exported and all the 64bit ones.

`WDL_NO_LIBS` - Fully exclude static libraries (currently only Interlocked); mainly intended for comparing current tB versions to Beta 423 where the `Import Library` syntax is not yet supported.

`WDL_NO_DELEGATES` - Do not use Delegate functions in place of function pointers.

`WDL_NO_WS_ALIASES` - Do not use `ws_`prefixes for the short name common word Winsock APIs (`send`, `connect`, `bind`, etc)

`WDL_XAUDIO8` - Use XAudio8 DLLs for XAudio2 APIs (Windows 8)

`WDL_NOMATH` - Exclude built in math helper function (see below). Note: XAudio2 inlined helper functions unavailable when math disabled.

`WDL_ADS_DEFINED` - activeds.tlb is referenced, enable interfaces using its contents.

>[!IMPORTANT]
>Currently flags are not inherited from the main project, so the only way to use these is to set them in the compiler flags for WinDevLib.twinproj then build a custom twinpack.

#### Custom Helpers
In addition to coverage of common Windows SDK-defined macros and inlined functions, a small number of custom helpers are provided to deal with Windows data types and similar not properly supported by the language. These are:

`Public Function GetMem(Of T)(ByVal ptr As LongPtr) As T` - A generic to dereference a pointer into any type. The native `CType(Of )` allows dereferencing to UDTs, but this helper allows instrinsic types in addition to UDTs, and is used the same way.

`Public Function DCast(Of T, T2)(v As T2) As T` - Direct Cast: Copies the data of v into any type, without modification, so no overflows, and possible to e.g. go from `LongLong` to `POINT`, with `Dim pt As POINT = DCast(Of POINT)(SomeLongLong)`

`Public Type CTypeHelper(Of T)` / `Public Type TType(Of T)` - These are helpers for the `CType(Of )` operator, intended to allow a pointer to refer to any type, not just a UDT. For example, if you have only a pointer to an array of Single, you could pass it to a ByRef As Single argument with `CType(Of TType(Of Single))(ptr).x` and the API would still be able to access all members just like f(0).

`Public Function LPWSTRtoStr(lPtr As LongPtr, Optional ByVal fFree As Boolean = True) As String`\
Converts a pointer to an LPWSTR/LPCWSTR/PWSTR/etc to an instrinsic `String` (BSTR)

`Public Function WCHARtoSTR(aCh() As Integer) As String`\
Converts an Integer arrat of WCHARs to a tB String (BSTR).

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
`Public Function toPOINT(ByVal x As Long, ByVal y As Long) As POINT`
`Public Function toPOINTF(ByVal x As Long, ByVal y As Long) As POINTF`
`Public Function toSIZE(ByVal cx As Long, ByVal cy As Long) As SIZE`\
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
   Sech, Sechf      - Hyperbolic secant
   Cosech, Cosechf - Hyperbolic cosecant
   Cotanh, Cotanhf - Hyperbolic cotangent
   Asinh, Asinhf   - Hyperbolic arcsine
   Acosh, Acoshf   - Hyperbolic arccosine
   Atanh, Atanhf   - Hyperbolic arccotangent
   Asech, Asechf   - Hyperbolic arcsecant
   Acosech, Acosechf - Hyperbolic arccosecant
   Acotanh, Acotanhf - Hyperbolic arccotangent
```

### Guide to switching existing code to WinDevLib

#### API definition differences
This section applies both to API calls and type library interface methods.

1) Convert `Currency` to `LongLong` for interfaces and APIs: It's no longer neccessary to worry about multiplying and dividing by 10,000 since tB supports a true 64bit integer type in both 32bit and 64bit mode. So this change is ultimately for the better, but existing codebases will have had to have used `Currency` for all interfaces and oleexp APIs expecting a 64bit integer.

2) Optional UDTs no longer use `As Any`. If you see errors like `Validation of call to 'CreateFile' failed.  Argument for 'lpSecurityAttributes': cannot coerce type 'Long' to 'SECURITY_ATTRIBUTES'`, this is an example of the issue. twinBASIC supports substituing `ByVal vbNullPtr` or any `LongPtr` for a UDT (the `ByVal` is now required), so WinDevLib can use the proper type while still permitting you to pass the equivalent of `ByVal 0`. 

Example:

VB6:
```vba
Public Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

hFile = CreateFileW(StrPtr("name"), 0, 0, ByVal 0, ...)
```
twinBASIC:
```vba
Public Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr

hFile = CreateFileW(StrPtr("name"), 0, 0, ByVal vbNullPtr, ...)
'---or---
Dim pSec As SECURITY_ATTRIBUTES
Dim lPtr As LongPtr = VarPtr(pSec)
hFile = CreateFileW(StrPtr("name"), 0, 0, ByVal lPtr, ...)
```

3) ByVal UDTs are supported in newer twinBASIC versions and are used in all cases in this package if used with tB Beta 896 or newer. So in some cases you may see a 'wrong number of arguments' error where for example a ByVal `POINT` or `SIZE` was split into two ByVal `Long`s. You'd now use the proper type.

4) String vs Long(Ptr) in APIs with both ANSI and Unicode versions: Most VB programs are written with ANSI versions of APIs being the default. **This is not the case with WinDevLib**. APIs are Unicode by default-- i.e. they use the W, rather than A, version of APIs e.g. `DeleteFile` maps to `DeleteFileW` rather than `DeleteFileA`. The A and W variants use String/LongPtr, and in almost all cases, the mapped version uses `String` with twinBASIC's `DeclareWide` keyword-- this disables Unicode-ANSI conversion. Since this is automatic, you generally don't need to make any changes; you can still use `String` without `StrPtr` or any manual Unicode <-> ANSI conversion. Note this usually only applies to strings passed as input, you'll need to update any externally allocated strings returned as a pointer, where you previously used e.g. `lstrlenA`, to use `lstrlenW` and Unicode handling in general. 

All APIs are provided, as a minimum, as the explicit W variant, and an untagged version that maps to the W version. Most ANSI variants are also included, but code should use Unicode wherever possible.

UDTs used by these calls are also supplied in the same manner, the W variant, an untagged variant that's the same as the W version, and in some cases, an A version. UDTs always use `LongPtr` for strings, even the untagged versions for `DeclareWide`. 

If you have any doubts about which API is being called, twinBASIC will show the full declaration when you hover your cursor over the API in your code.

Special thanks to GCUser99 for helping normalize API declaration in this project. 👍

5) Member names in UDTs have in almost all cases use their official SDK names, even where VB6 programmers traditionally used others. If you encounter errors where UDT members are missing, check the definition to see if the name has changed. This may also happen where unions are worked around in different ways.
   
6) **CURRENTLY N/A DUE TO BUGS** Callbacks that were previously LongPtr now expect a delegate function in many cases. A delegate is a typed function pointer defined with the correct prototype for the function it references. For compatibility these function the same way as previously and you need not make any changes, it's merely a better, easier way of defining callbacks that's also much closer to the C/C++ source. tB may show a warning, but it can be turned off.\
To use these, you can view the definition then make a regular function with that name and those arguments. See the [twinBASIC documentation overview of delegates](https://github.com/twinbasic/documentation/wiki/twinBASIC-Features#delegate-types-for-call-by-pointer) for more details and examples of using this new feature.

> [!TIP]
> Reminder: `Nothing` can be used in place of an interface where WinDevLib has the interface as an argument but another signature used `Long`/`LongPtr`


> [!NOTE]
>  This is just for using WinDevLib-- you'll likely have a lot more changes to make if you want to make your project x64 compatible. 


#### oleexp type library issues

WinDevLib started as a project to replace my VB6 COM interface type library, oleexp.tlb. This is still probably the most common use case. WinDevLib presented the best opportunity there would be to ditch some olelib legacy baggage. It's fairly simple to move your VB6 projects to WinDevLib, just follow these steps:

The follow steps apply only if you're converting code that previously relied on my oleexp.tlb project:
 
1) Replace oleexp.IUnknown with IUnknownUnrestricted. WinDevLib keeps this separate due to the major issues with conflicts with the former approach. If your project has IUnknown *without* oleexp. in front of it, **do not** replace those, as it's not referring to oleexp. 

2) After you've done that, you can now go ahead and simply delete all remaining instances of `oleexp.` (including the .). 

##### Issues specific to oleexpimp.tlb

There's 'WinDevLib for Implements' (WinDevLibImpl.twinpack/.twinproj) as well, but you'll note it has substantially fewer interfaces than oleexpimp.tlb. This is because there's two reasons for an interface to have an alternate version: It uses `[ Preservesig ]` on one or more methods, or it uses `As Any`. twinBASIC allows using `Implements` with `As Any` by replacing it with `As LongPtr` (which is what the alternate versions do). So many interfaces were in oleexpimp.tlb for this latter reason, and subsequently are *not* included in WinDevLibImpl as it's not neccessary.

If you find an oleexpimp.tlb interface is not in WinDevLibImpl, you will be able to use the one from WinDevLib, simply make sure `As Any` is changed to `As LongPtr`. 

tB has announced plans to support `[ PreserveSig ]` in implemented interfaces in the future; when that happens WinDevLibImpl will be deprecated.

>[!IMPORTANT]
>There currently seems to be an issue with using WinDevLib and WinDevLibImpl together if WinDevLibImpl does not use the current WinDevLib as a reference (it would usually use an old one as it's updated much less frequently). I've updated the reference on this repo and the package server, just note that you'll need to refresh both every time you update one if they're used together
>
>

#### Scope of coverage

The goal of the API coverage in WinDevLib is to provide the kind of programming experience you'd get in C/C++ by including windows.h and many of the more common feature sets like DirectX and GDIPlus. Even the 15,000+ APIs are just scratching the surface of the total Windows API set, and due to the low quality of automated conversion, even by Microsoft themselves (see: Win32API_PtrSafe.txt), I'm not interested in simply feeding headers or metadata through a conversion utility or using a database, so instead WinDevLib is focused on the most commonly used features in the major system DLLs-- everything 99%+ of apps need; though less common ones can be added by request or as time goes on and coverage is expanded.

Current coverage is already quite extensive, covering hundreds of Windows SDK header files. For details, see [COVERAGE.md](COVERAGE.md).
 
### Updates

**Update (v9.3.692, 12 Jun 2026):** 
- Misc additions from Windows SDKs from after 26000.100 through 28000.1839
- Minor custom d3d9 definitions added for compatibility with The trick's typelib.
- (Breaking change) For X64 CONTEXT struct I switched the active union arm as IMO the floating point registers are more common to need. CONTEXT_XMMSAVE is available with the old def.
- (Bug fix) IAttachmentExecute::SetClientGuid name typo
- (Bug fix) IID_ID3DXConstantTable duplicate definition

**Update (v9.3.688, 23 May 2026):** 
- Add d2d1effectauthor_1.h
- Add d3dx9anim.h
- Add DirectPlay Voice (dvoice.h, 100%; same differences vs dxvb dll as others) 
- (Bug fix) IID_IMarshalOptions function name typo

**Update (v9.3.686, 16 May 2026):** 
- Initial coverage of D3DX10 (d3dx10.h, d3dx10core.h, d3dx10tex.h, d3dx10async.h, d3dx10mesh.h 100%)
- Add d3d9on12.h, 100%
- Add DeleteBrowsingHistory.h, 100%
- Continued work to add [UseGetLastError(False)] for performance where appropriate.
- IDispError was missed in implementing ByVal UDTs
- Add common alias SendMessageLong (lParam=ByVal LongPtr)
- (API Standards, breaking change) IAdviseSink::OnRename now uses proper IMoniker type instead of LongPtr
- (API Standards, breaking change) STARTUPINFO size member is named cb, not cbSize
- (API Standards, breaking change) IDataObjectAsyncCapability, IThumbnailHandlerFactory now uses proper IBindCtx types.
- (API Standards, breaking change) NotifyServiceStatusChange[A] used -W UDT in argument.
- (API Standards, breaking change) PICTDESC hPalette should be named hpal.
- (API Standards) Some OLE functions now use As Any to be more correct than e.g. stdole.IPictureDisp. Not a breaking change.
- (Bug fix) D3D_SHADER_FEATURE_EXTENDED_COMMAND_INFO definition incorrect.
- (Bug fix) IXACT3Engine::PrepareStreamingWave no-byval-udt 32bit argument split missing end padding bytes.
- (Internal) Moved comdlg defs to comctl module.

**Update (v9.3.684, 29 Mar 2026):** 
- Added generic type CTypeHelper(Of T). This is designed to allow `CType(Of T)(pointer)` to work 
with basic types as it would with UDTs. The scenario that led to this was if you have e.g. a 
ByRef f As Single argument and you want to pass a pointer. ByVal LongPtr wouldn't work, but with 
this new helper, you can use `CType(Of CTypeHelper(Of Single))(ptr).x`. This works with arrays-- 
where the API is expecting a pointer to the first member of an array, this will still allow the 
API to read all of the members, not just the first one.\
For brevity this helper is also available as `TType(Of T)`
- Add X3DAudio for XACT3 (xact3d3.h, 100%)
- Add XAudio FX APIs from xapofx.h (100%)
- XACT3 now uses v3.7 GUIDs instead of 3.6.
- Add DX9 file handling interfaces (dxfile.h, 100%)
- (Bug fix) X3DAudioInitialize/X3DAudioCalculate should be CDecl. The former should also be a Sub. 
The X3DAUDIO_HANDLE arguments in both should take the type ByRef, not a Byte.
- (Bug fix) XAudio2 direct DLL exports are in xaudio2_9.dll, not _9d
- (Bug fix) ILRemoveLastID pidl should be ByVal

**Update (v9.3.682, 29 Mar 2026):** 
- Added XACT3 audio definitions (xact3.h, xact3wb.h, xma2defs.h 100%)  
Notably, these include not only the large inlined functions, but the full C++ type implementations 
with Subs within the UDTs. Additionally, these too have an extensive custom type set that has 
been preserved through Aliases.

**Update (v9.3.680, 28 Mar 2026):** 
- Added IMsoComponent/IMsoComponentManager for .NET interop use
- Added new MSVC compiler instrinsics implementations: _byteswap_ulong and _byteswap_ushort
- Misc Native API additions
- Continued work to add [UseGetLastError(False)] for performance where appropriate.
- (API Standards, breaking change) CreateThread, CreateRemoteThread[Ex] now specify proper lpThreadAttributes type. Change to ByVal vbNullPtr if 0 was used.
- (API Standards, breaking change) IEnumSpellingError should use PreserveSig  [#47](https://github.com/fafalone/WinDevLib/issues/47)
- (Bug fix) LdrOpenImageFileOptionsKey duplicate entry
- (Bug fix) RtlIsAnyDebuggerPresent used wrong offset for KUSER_SHARED_DATA
- (Bug fix) NoVersionLie custom option for IsWindowsVersionOrGreater was still subject to version lie.

**Update (v9.3.678, 27 Mar 2026):** 
- Updated D3D12 to match latest d3d12.h, d3d12compatibility.h, d3d12sdklayers.h, and d3d12video.h. Add d3d12compiler.h/.idl. (D3D12_SDK_VERSION = 619)
- netioapi.h now complete
- Continued work to add [UseGetLastError(False)] for performance where appropriate.
- Misc Native API additions
- (API Standards) IHttpSecurity now extends its base instead of duplicating the methods.
- (API Standards, breaking change) DS3DALG_* values are now proper UUIDs instead of Strings.
- (Bug fix) XInputGetBatteryInformation used wrong enum in TypeHint.
- (Bug fix) IDirectSound::CreateSoundBuffer last arg incorrect
- (Bug fix) IDirectSoundBuffer::Unlock definition incorrect

**Update (v9.3.676, 17 Mar 2026):** 
- Continued work to add [UseGetLastError(False)] for performance where appropriate.
- Add some missing SAFEARRAY APIs and now all are marked as [UseGetLastError(False)] for performance (they do not use it anyway)
- Direct3D 12 now implements ByVal UDTs
- IBindCtx, CoGetObject should take As Any to accommodate BINDOPTS2/3
- (API Standards, breaking change) IMoniker::IsRunning, GetDisplayName not consistent in taking IMoniker/IBindCtx
- (Breaking change) FormatMessage[A,W] now uses ByRef ParamArray for va_list instead of ByRef LongPtr.  
This will not break most uses that simply pass 0 or ByVal 0, but would impact uses that passed a pointer to a valid va_list memory structure.
- (Breaking change) TraceMessageVa now uses ByRef ParamArray for va_list instead of Any
- (Bug fix) TraceMessage no longer marked Unimplemented with missing vararg param 
- (Bug fix) ICatInformation::GetCategoryDesc incorrect for x64
- (Bug fix) Improper use of SAFEARRAY for C-type `FLOAT f[4]` etc defs  in ID3D12GraphicsCommandList
- (Bug fix) ID3D12Device8::CreateSamplerFeedbackUnorderedAccessView definition incorrect for 32bit

**Update (v9.3.674, 11 Mar 2026):** 
- Add some missing comctl constants and types
- (Bug fix) ImageList_CoCreateInstance definition incorrect for x64
- (Bug fix) NMTOOLBAR missing member. Also missing explicit A/W types.
- (Bug fix) TTOOLINFO[A,W] should be named TTTOOLINFO[A,W]. The old incorrect names have been left in as aliases to not break existing code.
- (Bug fix) TRACKBAR_CLASS definition incorrect
- (Bug fix) TVGETITEMPARTRECTINFO definition incorrect
- (Bug fix) NMTCKEYDOWN definition incorrect

**Update (v9.3.672, 11 Mar 2026):** 
- (Bug fix) IDataObject::DAdvise missing argument
- (WinDevLibImpl, bug fix) IDataObject::DAdvise missing argument

**Update (v9.3.670, 11 Mar 2026):** 
- Since WDL nows requires a higher minimum tB build for Alias support, the version gate
around ByVal UDTs has now been changed to an optional new compiler constant:  
WDL_NO_BYVAL_UDT is now available to use VBx-compatible definitions without ByVal UDTs.
- Updated WebView2 defs to latest stable release 1.0.3800.47
- GetTextExtentExPointI, IDvdInfo2::GetDVDDirectory now use more convenient LongPtr/String instead of ByRef Integer for a string input
- (API Standards, breaking change) IOleInPlaceSite::GetWindowContext, IOleInPlaceObject::SetObjectRects
now use proper ByRef types 
- (Bug fix) IDropSource::QueryContinueDrag missing argument
- (Bug fix) IEnumOLEVERB::Next, ICategorizer::GetDescription, ICategoryProvider::GetCategoryName, IBandSite::QueryBand, 
IOpenControlPanel::GetPath, IPropertySystem::FormatForDisplay, ICDBurn::GetRecordedDriveLetter, UrlGetPartW, 
UrlApplySchemeW, CryptSetAsyncParam, CryptGetAsyncParam, PathCompactPathExW, PathRelativePathToW, IImageList2::GetOriginalSize, 
IAdviseSink::OnViewChange, NotifyServiceStatusChange[A,W] definitions incorrect 
- (WinDevLibImpl, bug fix) IEnumOLEVERB::Next definition incorrect 

**Update (v9.2.668, 10 Mar 2026):** 
- Add DirectMusic for legacy compatibility
- Add missing DirectShow header dmort.h (100%)
- Misc Native API additions
- Continued Alias implementation
- (Bug fix) Some D3D_BLOB_TEST_* values incorrect.
- (Bug fix) RtlCompareUnicodeString return type incorrect

**Update (v9.2.664, 08 Mar 2026):** 
- Add additional DirectComposition interfaces/APIs from newer SDK/for newer Win10/11
- (Breaking change) IDCompositionAnimation::SetAbsoluteBeginTime now uses proper LARGE_INTEGER type.
- (Breaking change) NMLVLINK now uses proper member names from SDK to conform to WDL API Standards.
- (Bug fix) L_MAX_URL_LENGTH value incorrect, subsequently breaking NMLVLINK, NMLVEMPTYMARKUP etc

**Update (v9.2.662, 08 Mar 2026):** 
- Add minidumpapiset.h (100% inc delegates, inlines, and aliases)
- (Breaking change) WinHttpCrackUrl/WinHttpCreateUrl now use an alternate UDT, WINHTTP_URL_COMPONENTS, since
 it uses alternate values for the nScheme member. The SDK has these two definitions in direct conflict.
- (Breaking change) The WCHARtoSTR helper function now has two Optional arguments:  
bStopOnNull - Ends the string if a null character is encountered. This is True by default, where previously the
string would be continued and null characters simply ignored.  
bFilterNull - Sets whether to include null characters in the destination string. This is False by default, where previously
null characters would be ignored. This optional only applies if bStopOnNull is False.
- Misc API additions
- (Bug fix) INSTALLDATA missing union substitution padding on x64.
- (Bug fix) SetCurrentProcessExplicitAppUserModelID String overload missing DeclareWide.

**Update (v9.2.660, 01 Mar 2026):** 
- Misc Native API additions
- Additional error constants
- (Bug fix) IMediaEvent::GetEvent lParams should be ByRef

**Update (v9.2.658, 28 Feb 2026):** 
- Add additional SQL APIs from odbcinst.h (100%)  
   Includes custom ANSI versions; the entry points for them are the unmarked versions remapped to Unicode,
   so custom versions with -A suffixes are aliased to them instead.
- (Bug fix) SECURITY_TRUSTED_INSTALLER_RID2/5 were specified as decimal literals but were above the signed
 long limit, so would overflow when assigned to a Long as done in typical usage. Changed to hex literals.

**Update (v9.2.656, 25 Feb 2026):** 
- Add initial coverage of odbc32.dll SQL APIs; sqltypes.h, sql.h, sqlext.h, sqlucode.h 100%  
Note: Since these have their own complete, unique type set, original types are preserved via aliases
 rather than mapped to intrinsic types.  
Note: SQLCHAR/SQLWCHAR are aliased to String as they're exclusively used in this way.  
Note: OBDC_STD is supported if defined (is not by default).  
- Add capture and MCI window APIs from vfw.h
- Add ShellHandwriting.h/.idl (Note: APIs do not appear in any DLL; they may only exist in C .lib files)
- (Bug fix) BIND_OPTS3 definition incorrect
- (Bug fix) IRecordInfo::RecordCreate incorrect for x64
- (Bug fix) IMFProtectedEnvironmentAccess::ReadGRL, IMFSignedLibrary::GetProcedureAddress, IMF2DBuffer2::Lock2DSize,
IMF2DBuffer::Lock2D,GetScanline0AndPitch ByVal/ByRef


**Update (v9.2.653, 29 Jan 2026):** 
- (API Standards, BREAKING CHANGES) Non-ANSI APIs with As Any should be using DeclareWide. This was inconsistently applied. 
Fixing this will be ongoing. 
- (API Standards, BREAKING CHANGE) GetFileInformationByHandleEx should use As Any for multiple UDT PVOID.
- (Bug fix) ID3D10Blob::GetBufferPointer definition incorrect
- (Bug fix) EnumCalendarInfoExEx, lstrlen used String without DeclareWide when expecting LPWSTR
- (Bug fix) IMediaObjectInPlace::GetLatency arg should be ByRef
- (Bug fix) IMFAsyncCallbackLogging definition issues; Implements-compat version added to WinDevLibImpl
- (WinDevLibImpl, bug fix) Some MF interfaces did not have [PreserveSig] commented out

**Update (v9.2.651, 28 Jan 2026):** 
- Add htiface.h/.idl, htiframe.h/.idl 100%
- Misc native API additions
- Added constants for backwards compatibility with NeHe's OpenGL typelib
- (Bug fix) Corrected a number of constants in the form of Const x = &H8000-FFFF as these would 
improperly become negative Integer types.

**Update (v9.2.648, 25 Jan 2026):** 
- (BREAKING CHANGE) For compatibility with existing OpenGL work, void* arguments have been changed from ByVal As LongPtr to ByRef As Any, except in cases where it's explicitly asking for a pointer.
- Misc native API additions
- (Bug fix) D2D1_VECTOR_3F Alias/Type duplicate 

**Update (v9.2.646, 24 Jan 2026):** 
- Added Intel vendor-specific and WGL extensions OpenGL functions. 
- Added WFP ALE Endpoint APIs (fwpstypes.h 100%, fwpsu.h minimal)
- Misc native API additions
- Continued implementation of Alias types
- (Bug fix) PROCESS_MITIGATION_POLICY_INFORMATION union size incorrect
- (Bug fix) CLSID_SchedulingAgent should be CLSID_CSchedulingAgent

**Update (v9.2.644, 23 Jan 2026):** 
- Added AMD and NVIDIA vendor-specific OpenGL functions. OpenGL coverage is now complete.
- Add Software Licensing APIs (slpublic.h, slerror.h, sliddefs.h 100%)
- Add some missing DispIds for shell interfaces

**Update (v9.2.643, 22 Jan 2026):** 
- (Bug fix) bind/ws_bind namelen should be ByVal

**Update (v9.2.642, 18 Jan 2026):** 
- Add missing interfaces and consts from ShObjIdl_core.h
- Add missing interfaces and consts from DocObj.h/.idl
- GLchar/GLcharARB is now LongPtr since String would pass Unicode
- Misc native API additions
- (Bug fix) IPrint::Print definition incorrect
- (Bug fix) RtlSetProcessIsCritical, RtlSetThreadIsCritical missing CDecl 
- (Bug fix) PFNGLSHADERSOURCEPROC missing argument
- (Bug fix) Many OpenGL constants were not properly marked as Long
- NOTE: twinBASIC Beta 954 or newer is required for the OpenGL delegates

**Update (v9.2.640, 15 Jan 2026):** 
- Added OpenGL EXT and MS vendor-specific functions
- opengl32, glu32, and GDI+ APIs now use `[UseGetLastError(False)]` for performance, since they don't use SetLastError.
- Added compiler const WDL_NO_GL to disable OpenGL.
- (Bug fix) BitmapData last member should be LongPtr [#44](https://github.com/fafalone/WinDevLib/issues/44)

**Update (v9.2.639, 13 Jan 2026):** 
- (Bug fix) glu.h functions are DLL exports from glu32.dll, not loaded by wglGetProcAddress

**Update (v9.2.638, 13 Jan 2026):** 
- Added initial OpenGL coverage.  
-- Included: Windows SDK gl.h, glu.h; OpenGL 1.2-4.6; ARB approved extension; FreeGLUT. Planned but not yet included: EXT functions, vendor-specific functions.  
-- Note: Most functions are loaded dynamically, and a context must be created first.\
 WDL will automatically initialize all dynamic functions on the first use of any dynamic function. Thanks to Wayne Phillips for the technique.
- GdipAddPathStringI is missing [#42](https://github.com/fafalone/WinDevLib/issues/42)
- (Bug fix) GdipPathIterNextMarkerPath, GdipBitmapApplyEffect, GdipBitmapCreateApplyEffect definitions incorrect [#43](https://github.com/fafalone/WinDevLib/issues/43)
 

**Update (v9.2.634, 27 Dec 2025):** 
- Began process of implementing Alias syntax:  
-- **twinBASIC Beta 923 or newer is now required.** This is impractical to version-gate, therefore support for old tB betas is ending.
-- `CBoolean` added as alias for `Byte`, all `BOOLEAN` C types will be changed  
-- oleexp.tlb public aliases added and usages being restored (e.g. REFERENCE_TIME)  
-- ANSI/Unicode UDTs will use an alias for the non-explicit version  
-- UDT aliases will be used, e.g. D2D1_POINT_2F for D2D_POINT_2F 
-- Not all C types will be used; only where they're far removed from their underlying  
type. Like DWORD wouldn't be used, but D2D1_TAG would be as that's not clearly just an 
alternate for LongLong. Short wouldn't be used but ATOM would.  
-- This will be an ongoing process as to not hold up bug fixes and new features.
- Additional urlmon.h content
- Add missing DateTime notify type aliases (e.g. NMDATETIMEFORMATQUERY for NMDATETIMEFORMATQUERYW)
- (Breaking change) ZONEATTRIBUTES now uses proper Integer type for arrays
- (Bug fix) IMediaSample::GetPointer incorrect for x64
- (Bug fix) IMediaSample::GetMediaType, IEnumMediaTypes::Next definitions incorrect
- (Bug fix) COMBOBOXEXITEM improperly used String; renamed to COMBOBOXEXITEMA and alias for W version added.  
NMCOMBOBOXEX corrected to match. Same bug fixed for NMCBEDRAGBEGIN and NMCBEENDEDIT.
- (Bug fix) Huge number of ByVals that should be ByRef in DX11/12 interfaces.
- (Bug fix) GetComponentIDFromCLSSPEC missing argument
- (Bug fix) SetupComm definition incorrect for x64

**Update (v9.2.632, 09 Dec 2025):** 
- Added a number of missing combaseapi.h APIs to bring coverage to 100%
- Added Enclave APIs (ntenclv.h, winenclave.h, winenclaveapi.h 100% inc. delegates)
- Added IORing APIs (ioringapi.h, ntioring_x.h 100%)
- Added 100% coverage of wofapi.h
- Added monitor/video IOCTL defs (ntddvdeo.h, 100%)
- Added additional ntddvol.h IOCTL defs to bring coverage to 100%

**Update (v9.2.630, 06 Dec 2025):** 
- Added AppxPackaging.h/.idl, 100%
- Added XmlDom.idl, 100% -- This covers IXMLDOMDocument used in other headers and related interfaces/coclasses,
 and also covers the XMLHttpRequest object, but does not cover the full MSXML library (msxml[6].h)
- Some additional XML interfacesfrom msxml.
- Misc Native API additions.
- (Bug fix) For ListView macros, changed numerous improper ByRef 0 SendMessage lParam arguments to ByVal 0. Misc other corrections.
- (Bug fix) DbgPrint paramarray should be ByVal 

**Update (v9.2.627, 04 Dec 2025):**
- (Bug fix) IDsObjectPicker interface id incorrect, IID_IDsObjectPicker definition incorrect
- (WinDevLibImpl) Add IMoniker

**Update (v9.2.626, 03 Dec 2025):**
- **BREAKING CHANGES** Work has begun to standardize variable C-style array substitutions and make them able to
 work with tB's ability to turn off array bounds checking per-procedure. Types that used a buffer for a reasonable
 guess at the maximum are unchanged. If a 1-member array was omitted, it's now added. If it used a SAFEARRAY, it 
 will be changed to a single-member array and the SAFEARRAY version will be offered seperately with a `_sa` suffix.\
 This will be ongoing work due to the volume and lack of consistent labeling.
- **BREAKING CHANGE** Due to a recurring and current bug with inability to resolve `ShowWindow` (API)
vs `SHOWWINDOW` (Enum), the latter has been renamed to `eSHOWWINDOW`
- Large expansion of UDTs etc for queries to D3DKMTQueryAdapterInfo and D3DKMTQueryStatistics
- Added custom helpers toPOINT[F] and toSIZE to easily convert an x,y to the UDTs
- Added missing mciapi.h APIs, and macros + delegates
- Misc additions
- (Bug fix) mciGetYieldProc return type incorrect
- (Bug fix) OpenDedicatedMemoryPartition, QueryPartitionInformation are in kernelbase, not kernel32.

**Update (v9.2.624, 19 Nov 2025):**
- Added common ETW MOF structs
- Added helpers DEFINE_GUID, toPOINT[F], toSIZE
- (Bug fix) Two ByVal UDT APIs not version gated.

**Update (v9.2.622, 18 Nov 2025):**
- **MAJOR BREAKING CHANGES** twinBASIC now supports ByVal UDTs and WinDevLib will now also use these in all
 applicable situations. The initial conversion has been completed for all files. All types >8 bytes were 
 easy to identify as these had separate 32/64bit defs; 8 byte types should be complete as all uses of
 LongLong were examined. However for less than 8 bytes, it's possible some were missed if they weren't
 tagged, so reviewing for these will be ongoing and lengthy. Please notify me of any UDTs that should
 now be ByVal that were missed.\
 These changes are version gated, so are only active in tB Beta 896 and newer and the old definitions are
 still active in Beta 895 and earlier.\
 Please report any crashes or related bugs.
- Added custom helper WCHARtoSTR. Converts an Integer arrat of WCHARs to a tB String (BSTR). 
- Misc additions
- (Bug fix) IApplicationDesignModeSettings::IsApplicationViewStateSupported last argument should be ByRef.
- (Bug fix) ID3D12VideoEncoder::GetCodecProfile,GetCodecConfiguration defs incorrect
- (Bug fix) ID2D1Transform::MapInvalidRect definition incorrect for 32bit
- (Bug fix) GdipWarpPath definition incorrect for 32bit

**Update (v9.1.620, 14 Nov 2025):**
- Added ComDB APIs (msports.h, 100% inc. delegates)
- Added some missing functions from errhandlingapi.h to bring coverage to 100%

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

---

For earlier version history, see CHANGELOG.md

