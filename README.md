# WinDevLib 
## Windows Development Library for twinBASIC
### Formerly tbShellLib

### ***tbShellLib is now WinDevLib - Windows Development Library for twinBASIC***
This project has grown well beyond it's original mission of shell programming. While that's still the largest single part, it's no longer a majority of the code, and the name change now much better reflects the purpose of providing a general Windows API experience like windows.h. Compiler constants and module names/file names have been updated to reflect the name change. tbShellLibImpl is now WinDevLibImpl. There are also some major chanages associated with this update, please see the full changelog below.

### Big News!
**As of the [twinBASIC version Beta 368 and newer](https://github.com/twinbasic/twinbasic/releases), massive improvements to Intellisense mean that WinDevLib is vastly more usable, with no long delays. Intellisense is now cached and lag-free.** Thanks to Wayne for tackling this issue and continuing to make twinBASIC the programming tool of the future 👍

---

**Current Version: 7.0.282 (January 1st, 2024)**

(c) 2022-2023 Jon Johnson (fafalone)

This project is a comprehensive twinBASIC replacement for [oleexp.tlb](http://www.vbforums.com/showthread.php?786079-VB6-Modern-Shell-Interface-Type-Library-oleexp-tlb), my Modern Shell Interfaces Type Library project for VB6, that is x64 compatible, due to the many problems using midl to create a 64bit tlb.

This and oleexp are projects to supply Windows shell and component interfaces in a format consumable by VB6/VBA/tB. This involves not only defining interfaces, but using VB/tB compatible types-- so in some cases, even though there may be an existing way to import references to interfaces, they may be unusable due to e.g. the use of unsigned types, C-style arrays, double pointers, etc. All interfaces, types, consts, and APIs from oleexp are covered. For a full list of interfaces, see [INTERFACES.md](https://github.com/fafalone/WinDevLib/blob/main/INTERFACES.md).

In addition to the 2200+ common COM interfaces, WinDevLib now includes expansive coverage of Windows APIs from all the common modules. This makes it similar to working in C++ with `#include <Windows.h>` and a few others. Currently, approximately 5,500 of the most common APIs have been added- redone by hand from the original headers, in order to restore 64bit type info lost in VB6 versions, avoid the errors of automated conversion tools (e.g. Win32API_PtrSafe.txt is riddled with errors), and make them friendlier by converting groups of constants associated with a variable into an Enum so it comes up in Intellisense. This takes advantage of tB's ability to provide Intellisense for types besides Long in API defs (hopefully UDTs soon, this project has provisioning for that). 

This project is implemented purely in tB native code, as unlike VB6 there's language support for defining interfaces and coclasses. As a twinPACKAGE, regular code is supported in addition to the definitions, so the regular addin modules have been built in (mIID.bas, mPKEY.bas, etc). Does it still make sense to use a project like this when interfaces can be defined in-language? I'd say yes, because for a large number of interfaces, there's deep dependency chains with other interfaces and the types they rely on. It makes more sense to drop this in and be done with it than constantly have to define the interfaces you want and then stubs for their dependencies, especially when you might need those later on. This project is even more useful now with the API coverage; it should cover about 99% of your needs for core system DLLS. 

Please report any bugs via the Issues feature here on GitHub.

### Requirements

[twinBASIC Beta 269 or newer](https://github.com/twinbasic/twinbasic/releases) is required.

### Adding WinDevLib to your project
You have 2 options for this:

#### Via the Package Server
twinBASIC has an online package server and WinDevLib is published on it. Open your project settings and scroll to the **COM Type Library / ActiveX References**, then click **TWINPACK PACKAGES**. Add "Windows Development Library for twinBASIC v7.0.272" (or whatever the newest version is). The other similar entry, "WinDevLib for Implements" contains `Implements` compatible versions of a small number of common interfaces not defined in a compatible way in the main project; you normally don't need this. For more details, including illustrations, [see this post](https://github.com/fafalone/WinDevLib/issues/9#issuecomment-1416767019).

#### From a local file
You can download the project from this repository and use the WinDevLib.twinpack file. Navigate to the same area as above, and click on the "Import from file" button. WinDevLib.twinproj is the source for the package, if you want to edit it.


### Optional Features
WinDevLib has some compiler constants:

`WINDEVLIB_LITE` - This flag disables most API declares and misc WinAPI definitions, including everything in wdAPIComCtl, wdAPI, and wdDefs. I used to like doing my APIs separate too, which is why oleexp never had the expansive coverage. But with that coverage now present, I think it's worth using, but this option will still be supported.

`WINDEVLIB_COMCTL_LIB_DEFINED` - You can use this flag if you already have an alternative common controls definition set, e.g. tbComCtlLib; it will disable wdAPIComCtl. (Note: WinDevLib has more complete comctl defs than tbComCtlLib, as that project was deprecated and not updated).

`WINDEVLIB_DLGH` - This enabled constants from dlg.h. These are extremely uncommon to use, and have very short, generic names likely to cause conflicts, so they're opt-in.

`WINDEVLIB_NOQUADLI` - Restores the old `LARGE_INTEGER` definition of lo/high Long values.

>[!WARNING]
>The `WINDEVLIB_NOQUADLI` constant will break alignment on numerous Types; most only on x64, but some on both. 

`

### Guide to switching from oleexp.tlb

WinDevLib presented the best opportunity there would be to ditch some olelib legacy baggage. It's fairly simple to move your VB6 projects to WinDevLib, just follow these steps:

1) Replace public aliases: It's important to do this first. Run a Replace All changing oleexp.LONG_PTR to LongPtr, oleexp.REFERENCE_TIME to LongLong, oleexp.HNSTIME to LongLong, oleexp.KNOWNFOLDERID to UUID, oleexp.EventRegistrationToken to LongLong, oleexp.BINDPTR to LongPtr, and oleexp.LPCRITICAL_SECTION to LongPtr. If you've used them without the oleexp. prefix, you'll also need to replace those, but if you've imported into tB they should be tagged.

2) Replace oleexp.IUnknown with IUnknownUnrestricted. WinDevLib keeps this separate due to the major issues with conflicts with the former approach. If your project has IUnknown *without* oleexp. in front of it, **do not** replace those, as it's not referring to oleexp. 

3) After you've done those two, you can now go ahead and simply delete all remaining instances of `oleexp.` (including the .). 

4) Convert `Currency` to `LongLong` for interfaces and APIs: It's no longer neccessary to worry about multiplying and dividing by 10,000 since tB supports a true 64bit type in both 32bit and 64bit mode. So this change is ultimately for the better, but existing codebases will have had to have used `Currency` for all interfaces and oleexp APIs expecting a 64bit integer.

5) Manually address any errors remaining. Interfaces should be mostly fine at this point, but if you've made use of the APIs in oleexp, many of them have syntax differences, mainly not being able to rewrite an ending [out] argument as the return value, and changing String arguments to `LongPtr` you'll need `StrPtr` with. Another major difference is that the default for almost all APIs with ANSI/Unicode (A/W) versions, is now the Unicode version. In most cases, the W version is declared with `LongPtr` for strings, and the untagged alias version uses tB's new `DeclareWide` keyword to disable ANSI conversion while using `String`.\
Finally, a very small number of APIs and interfaces use ByVal UDTs. Since VB cannot do this, nor can tB yet, a typical workaround was to pass each member as an individual argument. This worked when arguments were 4 bytes each, but the x64 calling convention aligns arguments at 8 bytes. So the two options were to follow that convention, which also works for 32bit allowing a single call for both, or require two different calls for 32 and 64bit. Since one of the main points of twinBASIC is 64bit support, WinDevLib uses the former option. The downside of this is that VB-style calls will have to be rewritten. If you see, for example, `ByVal ptX As Long, ByVal ptY As Long` replaced with `ByVal pt As LongLong`, this was an unsupported `ByVal POINT`. You'd declare a LongLong, and use `CopyMemory` to fill it: `Dim pt As POINT: Dim ptt As Long: ...: CopyMemory ptt, pt, 8`.

> [!NOTE]
>  This is just for using WinDevLib-- you'll likely have a lot more changes to make if you want to make your project x64 compatible.

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

As noted before, an exception to the rule is `SendMessage`, due to the enourmous volume of existing code expecting SendMessage to map to SendMessageA.

If you have any doubts about which API is being called, twinBASIC will show the full declaration when you hover your cursor over the API in your code.

#### Scope of coverage

The goal of the API coverage in WinDevLib is to provide the kind of programming experience you'd get in C/C++ by including windows.h and some of the more common feature sets like DirectX and GDIPlus. It currently includes about 5,500 APIs. But even that is just scratching the surface of the total Windows API set. Due to the low quality of automated conversion, even by Microsoft themselves (see: Win32API_PtrSafe.txt), I'm not interested in simply feeding headers through a conversion utility or using a database, so instead WinDevLib will be focused on the most commonly used features in the major system DLLs, though less common ones can be added by request or as time goes on and the existing DLLs are completed. I do not intend to include native APIs that have fully equivalent regular APIs; that's basically doubling the work for no benefit-- but if they offer additional features or substantially improved performance, they will be included. 

I've included the definitions, associated types, and associated constants, for extensive portions of the following modules: shell32.dll, shlwapi.dll, kernel32.dll, user32.dll, advapi32.dll, ole32.dll, oleaut32.dll, propsys.dll, gdi32.dll, gdiplus.dll, urlmon.dll, hlink.dll, winmm.dll, setupapi.dll, comctl32.dll, dwm.dll/uxtheme.dll, comdlg32.dll, winspool.drv, and netapi32.dll. Besides highly self-contained specialized sets in their own headers (unless small), please let me know any I've missed from these.\
Limited coverage (or full coverage of very small sets) is provided for ntdll.dll, version.dll, msimg32.dll, crypt32.dll, virtdisk.dll, sxs.dll, secur32.dll, imm32.dll, userenv.dll, wintrust.dll, msacm32.dll, url.dll, htmlhelp.dll, imagehlp.dll, and ws2_32.dll. If you feel any missing ones from these should be included, or would like to contribute more, let me know.\
Finally, there's small API sets for features, like DirectX DLLs, Webview2Loader, WIC, etc. Definitely let me know any missing from these.

**Future coverage:** In the future I'm planning to expand crypto coverage from advapi32 and wintrust, expand native APIs with no equivalents, add additional Winsock coverage, and add OpenGL-- though for this last one I may wait for tB to have `Alias` support since existing OpenGL codebases make heavy use of them by way of NeHe's TLB. I welcome contributions of any of these. If you've done the consts->enums conversions already, I'd even take 32bit-only versions.


#### A note on seeing UDTs where before they were As Any

>[!NOTE]
>If you see errors like `Validation of call to 'CreateFile' failed.  Argument for 'lpSecurityAttributes': cannot coerce type 'Long' to 'SECURITY_ATTRIBUTES'`, this section explains the cause and solution!

The best example of this is many APIs, like file APIs, where in traditional VB declarations, you see 'As Any' and in WinDevLib you see e.g. `SECURITY_ATTRIBUTES` or `OVERLAPPED`. These are the correct the definitions, but VB6 had no facility to specify 'NULL', which is what they usually would be set to as optional arguments. So the VB6 way was a workaround, where you could pass ByVal 0. 

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

### Updates

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
-(MAJOR BUG FIX) IStream was missing UnlockRegion. This impacted numerous derived interfaces, throwing off their vtables, completely breaking them.


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

**Update (v6.3.240):**\
-Added interfaces IComputerAccounts, IEnumAccounts, IComputerAccountNotify, and IProfileNotify with coclasses LocalUserAccounts, LocalGroups, LoggedOnAccounts, ProfileAccounts, UserAccounts, and ProfileNotificationHandler. Also added numerous PROPERTYKEYs associated with this functionality.\
-Added a limited set of Winsock APIs. Note that with the exception of WSA* APIs, the short, generic names have been prefixed with ws_.\
-Misc API additions including undocumented shell32 APIs, and additional ntdll APIs.\
-Additional PE file structs\
-(Bug fix) Several WebView2 interface had incompatible Property Get defs for ByVal UDT workarounds.

**Update (v6.2.238):**\
-Added a limited set of winhttp APIs\
-Added misc APIs for recent projects\
-(Bug fix) RegQueryValueEx/RegQueryValueExW/RegQueryValueExA definitions incorrect.

**Update (v6.2.237):** Missing consts for upcoming project.

**Update (v6.2.234):**
-Added additional file info structs, exe header structs, and ntdll API\
-(Bug fix) Some Disk Quota interface enums had incorrect names and in some cases values.

**Update (v6.2.232):** 
-Added gdi32 Color Management (ICM) APIs.\
-Additional sysinfo UDTs.\
-TypeHints for NT functions missing them.

**Update (v6.2.230):** 
-Added Windows Networking (WNet) APIs (winnetwk.h, 100% coverage (mpr.dll))\
-Major expansion of internationalization API coverage from winnls.h.\
-Added numerous missing common User32 functions.\
-Misc bug fixes, inc. InsertMenuItem entry-point not found, missing menu alternates (W or A variations)\
-Added overloads for a number of functions, if you have any trouble with the following, please file a bug report:\
CoUnMarshalIface, IsValidLocaleName, EnumDateFormatsExEx, EnumCalendarInfoExEx, GetSystemDefaultLocaleName, GetCurrencyFormatEx, GetNumberFormatEx, GetCalendarInfoEx, SetUserGeoName, GetThreadPreferredUILanguages, SetThreadPreferredUILanguages, SetProcessPreferredUILanguages, GetProcessPreferredUILanguages, LocaleNameToLCID, GetDurationFormat, GetDurationFormatEx, GetLocaleInfoEx, ResolveLocaleName, GetNLSVersion, GetNLSVersionEx, ToUnicode, LoadBitmap[A,W], ModifyMenu, InsertMenu. 

**Update (v6.1.229):** Bug fix: A number of APIs had missing 'As <type>` statements, which were upgraded to errors. tB had previosly not caught these.

**Update (v6.1.228):**
-Completed imm32 APIs\
-Added Job Object APIs\
-Completed Virtual Disk APIs (virtdisk.h, 100% coverage)\
-Many missing gdi32.dll APIs\
-Misc APIs, inc. some power APIs\
-All UDTs for NtQueryInformationFile (through current Win11)\
-Bug fix: GDI object enum duplicate\
-Bug fix: Some incorrect UDTs

**Update (v6.0.220):**\
-Added Network List Manager interfaces and coclass NetworkListManager.\
-Added WININET APis (wininet.h, 99% coverage-- autoproxy defs unsupported by language)\
-Added all APIs from iphlpapi.h (IP Helper; network stats); netioapi.h not included. Will be in future release.\
-Added all Console APIs (wincon.h/wincontypes.h/consoleapi[, 2,3].h) and Comm APIs. WinEvent APIs and consts.\
-FileDeviceTypes has been renamed DEVICE_TYPE, per usage in km\
-Added most UDTs for GetFileInformationByHandle and native equivalents\
-Added Vista+ Thread Pool APIs, including inlined ones (threadpoolapiset.h, 100% coverage)\
-Added Windows 10+ Secure Enclave APIs (enclaveapi.h, 100% coverage)\
-dlgs.h, part of windows.h, has been added *AS AN OPTIONAL EXTENSION* due to anticipated naming conflicts with common names like 'lst1'. Add the compiler constant `TB_SHELLLIB_DLGH = 1` to include these.\
-Bug fix: Numerous UDTs with LARGE_INTEGER changed to QLARGE_INTEGER where the lack of 8-byte QuadPart was throwing alignment off. Note that in the future, tB will have union support, at which point LARGE_INTEGER will be changed to one, and all QLARGE_INTEGER replaced.

**Update (v5.3.214):** Added all DWM APIs from dwmapi.h. Added undoc'd shell app manager interfaces/coclasses. Added CPL applet defs. Misc API additions and bugfixes.

**Update (v5.2.210/212):** Additional APIs for upcoming project release.

**Update (v5.2.208):** Substantial API additions; inc. SystemParametersInfo structs/enums, display config, raw input, missing dialog stuff. Additional standard helper macros found in Windows headers.

**Update (5.1.206/207):**\
-Added PropSheet macros.\
-Set PROPSHEETPAGE to V4 by default.\
-Add missing PropSheet consts.\
-Bug fix: PROPSHEETHEADER definitions incorrect.\
-Bug fix: PostMessage API not 64bit compatible.\
-Bug fix: Several ListView macros not 64bit compatible.\
-Updated WebView2 to match 1.0.1901.177.\
-Completed all advapi32 registry functions.\
-Expanded Media Foundation APIs.\
-Bug fix: Property Sheet callback enums were missing values and improperly organized.\
-Misc bug fixes and additions to APIs.\

**Update (v5.0.203):** Bug fix: D3DMATRIX layout with 2d array was incorrect.

**Update (v5.0.201):**\
-Added some missing DirectShow media stream interfaces.\
-Complete coverage of winmm API sets for wave, midi, time, sound, mmio, joystick, mci, aux, and mixer.\
-Complete coverage of printer and print spooler APIs from winspool.\
-Major expansion of security-related APIs\
-Added D3D compiler APIs and effects interfaces\ 
-Added basic DirectSound interfaces/apis.\
-Bug fix: ShowWindow relocated to slShellCore.twin to avoid amibiguity with SHOWWINDOW enum.\
-Bug fix: Misc. bug fixes to APIs.


---

For earlier version history, see CHANGELOG.md

For more information and a list of available interfaces, visit the [VB Forums thread](https://www.vbforums.com/showthread.php?897883-twinBASIC-tbShellLib-Shell-Interface-Library-(x64-compatible-successor-to-oleexp)) for this project.
