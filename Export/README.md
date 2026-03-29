# WinDevLib 
## Windows Development Library for twinBASIC 

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

`Public Function DCast(Of T, T2)(ByVal v As T2) As T` - Direct Cast: Copies the data of v into any type, without modification, so no overflows, and possible to e.g. go from `LongLong` to `POINT`, with `Dim pt As POINT = DCast(Of POINT)(SomeLongLong)`

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
Functions for converting POINT and SIZE types or coords to Long or LongLong.

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

3) ByVal UDTs are supported in twinBASIC and are used in all cases in this package.

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


#### Scope of coverage

The goal of the API coverage in WinDevLib is to provide the kind of programming experience you'd get in C/C++ by including windows.h and many of the more common feature sets like DirectX and GDIPlus. It currently includes about 5,500 APIs. But even that is just scratching the surface of the total Windows API set. Due to the low quality of automated conversion, even by Microsoft themselves (see: Win32API_PtrSafe.txt), I'm not interested in simply feeding headers through a conversion utility or using a database, so instead WinDevLib will be focused on the most commonly used features in the major system DLLs, though less common ones can be added by request or as time goes on and the existing DLLs are completed. I do not intend to include most native APIs that have fully equivalent regular APIs; that's basically doubling the work for no benefit-- but if they offer additional features or substantially improved performance, they will be included. 

I've included the definitions, associated types, and associated constants, for extensive portions of the following modules: shell32.dll, shlwapi.dll, kernel32.dll, ktmw32.dll, user32.dll, advapi32.dll, tdh.dll, authz.dll, crypt32.dll, wintrust.dll, bcrypt.dll, ncrypt.dll, cryptui.dll, ole32.dll, oleaut32.dll, propsys.dll, gdi32.dll, gdiplus.dll, virtdisk.dll, userenv.dll, dbghelp.dll, mpr.dll, iphlpapi.dll, urlmon.dll, hlink.dll, winmm.dll, cfgmgr32.dll, setupapi.dll, comctl32.dll, dwm.dll/uxtheme.dll, comdlg32.dll, winbio.dll, winspool.drv, imm32.dll, hid.dll, cldapi.dll, pdh.dll, powrprof.dll, wtsapi32.dll, and netapi32.dll. Please let me know any I've missed from these.\
Limited coverage (or full coverage of very small sets) is provided for ntdll.dll, version.dll, msimg32.dll, fwpuclnt.dll, sxs.dll, secur32.dll, msacm32.dll, url.dll, htmlhelp.dll, avifil32.dll, and ws2_32.dll. If you feel any missing ones from these should be included, or would like to contribute more, let me know.\
Finally, there's numerous additional API sets from small to large for independent Windows features. These include small sets like restartmgr.dll through very large sets like the various Media Foundation and DirectX DLLs. In the future I'll better organize coverage lists, but the bottom line is let me know if any common APIs or built in API sets for components should be added. TODO.md in the WDL project files contains ones planned but not yet done.

**Future coverage:** In the future I'm planning to expand native APIs with no equivalents, add additional Winsock coverage, and add OpenGL-- though for this last one I may wait for tB to have `Alias` support since existing OpenGL codebases make heavy use of them by way of NeHe's TLB. I welcome contributions of any of these. If you've done the consts->enums conversions already, I'd even take 32bit-only versions.


### ***tbShellLib is now WinDevLib - Windows Development Library for twinBASIC***
This project has grown well beyond it's original mission of shell programming. While that's still the largest single part, it's no longer a majority of the code, and the name change now much better reflects the purpose of providing a general Windows API experience like windows.h. Compiler constants and module names/file names have been updated to reflect the name change. tbShellLibImpl is now WinDevLibImpl. There are also some major chanages associated with this update, please see the full changelog below.
