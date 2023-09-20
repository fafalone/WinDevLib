# tbShellLib
## twinBASIC Shell Library

### Big News!
**As of the [twinBASIC version Beta 368 and newer](https://github.com/twinbasic/twinbasic/releases),, massive improvements to Intellisense mean that tbShellLib is vastly more usable, with no long delays. Intellisense is now cached and lag-free.** Thanks to Wayne for tackling this issue and continuing to make twinBASIC the programming tool of the future 👍

---

**Current Version: 6.1.228 (September 20th, 2023)**

(c) 2022-2023 Jon Johnson (fafalone)

This project is a comprehensive twinBASIC replacement for [oleexp.tlb](http://www.vbforums.com/showthread.php?786079-VB6-Modern-Shell-Interface-Type-Library-oleexp-tlb), my Modern Shell Interfaces Type Library project for VB6, that is x64 compatible, due to the many problems using midl to create a 64bit tlb.

This and oleexp are projects to supply Windows shell and component interfaces in a format consumable by VB6/VBA/tB. This involves not only defining interfaces, but using VB/tB compatible types-- so in some cases, even though there may be an existing way to import references to interfaces, they may be unusable due to e.g. the use of unsigned types, C-style arrays, double pointers, etc.

All interfaces, types, consts, and APIs from oleexp are covered, and there's additional API coverage not included in oleexp. For a full list of interfaces, see [INTERFACES.md](https://github.com/fafalone/tbShellLib/blob/main/INTERFACES.md).

This project is implemented purely in tB native code, as unlike VB6 there's language support for defining interfaces and coclasses. As a twinPACKAGE, regular code is supported in addition to the definitions, so the regular addin modules have been built in (mIID.bas, mPKEY.bas, etc). Does it still make sense to use a project like this when interfaces can be defined in-language? I'd say yes, because for a large number of interfaces, there's deep dependency chains with other interfaces and the types they rely on. It makes more sense to drop this in and be done with it than constantly have to define the interfaces you want and then stubs for their dependencies, especially when you might need those later on. This project is even more useful now with the API coverage; it should cover about 99% of your needs for system DLLS. 

Please report any bugs via the Issues feature here on GitHub.

### Requirements

[twinBASIC Beta 269 or newer](https://github.com/twinbasic/twinbasic/releases) is required.

### Adding tbShellLib to your project
You have 2 options for this:

#### Via the Package Server
twinBASIC has an online package server and tbShellLib is published on it. Open your project settings and scroll to the **COM Type Library / ActiveX References**, then click **TWINPACK PACKAGES**. Add "twinBASIC Shell Library v3.4.46" (or whatever the newest version is). "twinBASIC Shell Library for Implements" contains `Implements` compatible versions of a small number of common interfaces not defined in a compatible way in the main project; you normally don't need this. For more details, including illustrations, [see this post](https://github.com/fafalone/tbShellLib/issues/9#issuecomment-1416767019).

#### From a local file
You can download the project from this repository and use the .twinpack file. Navigate to the same area as above, and click on the "Import from file" button. 

### Guide to switching from oleexp.tlb

tbShellLib presented the best opportunity there would be to ditch some olelib legacy baggage. It's fairly simple to move your VB6 projects to tbShellLib, just follow these steps:

1) Replace public aliases: It's important to do this first. Run a Replace All changing oleexp.LONG_PTR to LongPtr, oleexp.REFERENCE_TIME to LongLong, oleexp.HNSTIME to LongLong, oleexp.KNOWNFOLDERID to UUID, oleexp.EventRegistrationToken to LongLong, oleexp.BINDPTR to LongPtr, and oleexp.LPCRITICAL_SECTION to LongPtr. If you've used them without the oleexp. prefix, you'll also need to replace those, but if you've imported into tB they should be tagged.

2) Replace oleexp.IUnknown with IUnknownUnrestricted. tbShellLib keeps this separate due to the major issues with conflicts with the former approach. If your project has IUnknown *without* oleexp. in front of it, **do not** replace those, as it's not referring to oleexp. 

3) After you've done those two, you can now go ahead and simply delete all remaining instances of `oleexp.` (including the .). 

4) Convert `Currency` to `LongLong` for interfaces and APIs: It's no longer neccessary to worry about multiplying and dividing by 10,000 since tB supports a true 64bit type in both 32bit and 64bit mode. So this change is ultimately for the better, but existing codebases will have had to have used `Currency` for all interfaces and oleexp APIs expecting a 64bit integer.

5) Manually address any errors remaining. Interfaces should be mostly fine at this point, but if you've made use of the APIs in oleexp, many of them have syntax differences, mainly not being able to rewrite an ending [out] argument as the return value, and changing String arguments to LongPtr you'll need StrPtr with. Another major difference is that the default for almost all APIs with ANSI/Unicode (A/W) versions, is now the Unicode version. A notable exception is `SendMessage` due to the overwhelming amount of VBx code expecting it to mean `SendMessageA`. In most cases, the W version is declared with `LongPtr` for strings, and the untagged alias version uses tB's new `DeclareWide` keyword to disable ANSI conversion while using `String`.\
Finally, a very small number of APIs and interfaces use ByVal UDTs. Since VB cannot do this, nor can tB yet, a typical workaround was to pass each member as an individual argument. This worked when arguments were 4 bytes each, but the x64 calling convention aligns arguments at 8 bytes. So the two options were to follow that convention, which also works for 32bit allowing a single call for both, or require two different calls for 32 and 64bit. Since one of the main points of twinBASIC is 64bit support, tbShellLib uses the former option. The downside of this is that VB-style calls will have to be rewritten. If you see, for example, `ByVal ptX As Long, ByVal ptY As Long` replaced with `ByVal pt As LongLong`, this was an unsupported `ByVal POINT`. You'd declare a LongLong, and use `CopyMemory` to fill it: `Dim pt As POINT: Dim ptt As Long: ...: CopyMemory ptt, pt, 8`.

> [!NOTE]
>  This is just for using tbShellLib-- you'll likely have a lot more changes to make if you want to make your project x64 compatible.

### Guide to switching from oleexpimp.tlb

There's 'twinBASIC Shell Library for Implements' (tbShellLibImpl.twinpack/.twinproj) as well, but you'll note it has substantially fewer interfaces than oleexpimp.tlb. This is because there's two reasons for an interface to have an alternate version: It uses `[ Preservesig ]` on one or more methods, or it uses `As Any`. twinBASIC allows using `Implements` with `As Any` by replacing it with `As LongPtr` (which is what the alternate versions do). So many interfaces were in oleexpimp.tlb for this latter reason, and subsequently are *not* included in tbShellLibImpl as it's not neccessary.

If you find an oleexpimp.tlb interface is not in tbShellLibImpl, you will be able to use the one from tbShellLib, simply make sure `As Any` is changed to `As LongPtr`. 

tB has announced plans to support `[ PreserveSig ]` in implemented interfaces in the future; when that happens tbShellLibImpl will be deprecated.

### tbShellLib API standards

This was mentioned above, but it's worth going into more detail. In addition to the COM interfaces, tbShellLib has a large selection of common Windows APIs; this is a much larger set than oleexp. tbShellLib and twinBASIC represented the best opportunity there would be to modernize standards... most VB programs are written with ANSI versions of APIs being the default. **This is not the case with tbShellLib**. With very few exceptions, APIs are Unicode by default-- i.e. they use the W, rather than A, version of APIs e.g. `DeleteFile` maps to `DeleteFileW` rather than `DeleteFileA`. The A and W variants use String/LongPtr, and in almost all cases, the mapped version uses `String` with twinBASIC's `DeclareWide` keyword-- this disables Unicode-ANSI conversion, so you can still use `String` without `StrPtr` or any Unicode <-> ANSI conversion. Note this usually only applies to strings passed as input, APIs passing a LPWSTR that's allocated externally will still be LongPtr, as they're not in the same BSTR format as VBx/TB strings.

All APIs are provided, as a minimum, as the explicit W variant, and an untagged version that maps to the W version. Some, but not all, APIs also have an explicit A variant defined that will perform the normal ANSI conversion for compatibility purposes. This is decided on a case by case basis depending on my impression of how much legacy code is around that needs the ANSI version. All new code should use the Unicode versions.

UDTs used by these calls are also supplied in the same manner, the W variant, an untagged variant that's the same as the W version, and in some cases, an A version. UDTs always use `LongPtr` for strings, even the untagged versions for `DeclareWide`. 

As noted before, an exception to the rule is `SendMessage`, due to the enourmous volume of existing code expecting SendMessage to map to SendMessageA.

If you have any doubts about which API is being called, twinBASIC will show the full declaration when you hover your cursor over the API in your code.

#### Scope of coverage

The goal of the API coverage in tbShellLib is to provide the kind of programming experience you'd get in C/C++ by including windows.h and some of the more common feature sets like DirectX and GDIPlus. It currently includes about 3,500 APIs. But even that is just scratching the surface of the total Windows API set. Due to the low quality of automated conversion, even by Microsoft themselves (see: Win32API_PtrSafe.txt), I'm not interested in simply feeding headers through a conversion utility or using a database, so instead tbShellLib will be focused on the most commonly used features in the major system DLLs. For example I've not included the Event Tracing APIs, as even though they're in advapi32.dll, they're a self-contained highly specialized set not found in the standard headers. I also do not intend to include native APIs that have fully equivalent regular APIs; that's basically doubling the work for no benefit. 

I've included the definitions, associated types, and associated constants, for extensive portions of the following modules: shell32.dll, shlwapi.dll, kernel32.dll, user32.dll, advapi32.dll, ole32.dll, oleaut32.dll, propsys.dll, gdi32.dll, gdiplus.dll, urlmon.dll, hlink.dll, winmm.dll, setupapi.dll, comctl32.dll, dwm.dll/uxtheme.dll, comdlg32.dll, winspool.drv, and netapi32.dll. Besides highly self-contained specialized sets in their own headers (unless small), please let me know any I've missed from these.\
Limited coverage (or full coverage of very small sets) is provided for ntdll.dll, version.dll, msimg32.dll, crypt32.dll, virtdisk.dll, sxs.dll, secur32.dll, imm32.dll, userenv.dll, wintrust.dll, msacm32.dll, url.dll, htmlhelp.dll, imagehlp.dll, and ws2_32.dll. If you feel any missing ones from these should be included, or would like to contribute more, let me know.\
Finally, there's small API sets for features, like DirectX DLLs, Webview2Loader, WIC, etc. Definitely let me know any missing from these.

**Future coverage:** In the future I'm planning to expand crypto coverage from advapi32 and wintrust, expand native APIs with no equivalents, and add coverage of Iphlpapi.dll. Winsock coverage is also planned, but this will likely be as an opt-in requiring a compiler constant, as there's just so many common two or three letter functions and types that would cause endless conflicts. I welcome contributions of any of these. If you've done the consts->enums conversions already, I'd even take 32bit-only versions.


#### A note on seeing UDTs where before they were As Any

The best example of this is many APIs, like file APIs, where in traditional VB declarations, you see 'As Any' and in tbShellLib you see e.g. `SECURITY_ATTRIBUTES` or `OVERLAPPED`. These are the correct the definitions, but VB6 had no facility to specify 'NULL', which is what they usually would be set to. So the VB6 way was a workaround, where you could pass ByVal 0. 

twinBASIC has direct support for passing a null pointer instead of a UDT. You can pass `vbNullPtr` to these arguments where previously you would have used ByVal 0 on an `As Any` argument that you've found is now a UDT. 

Example:

VB6:
```
Public Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

hFile = CreateFileW(StrPtr("name"), 0, 0, ByVal 0, ...)
```
twinBASIC:
```
Public Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr

hFile = CreateFileW(StrPtr("name"), 0, 0, vbNullPtr, ...)
```

### Updates

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


**Update (v4.16.191):** Critical bug fix: Multiple instances of errors for auto-declaring Variants. Bug fix: `GetClipboardData` incorrectly returned a Long (should be LongPtr).

**Update (v4.16.190):**\
-Critical bug fix: `TB_SHELLLIB_LITE` mode was broken.\
-Added additional DirectX errors w/ desciprtions.\
-Added initial D3D compiler apis, note that by default, these direct to d3dcompiler_47.dll, however you can specify compiler flag D3D_COMPILER = 44, 45, and 46 to use those.

**Update (v4.15.188):** Added `SAFEARRAY` APIs for manual operations on them and some more TypeLib-related APIs.

**Update (v4.15.185):** Bug fix: lstrcmp, lstrcmpi, and lstrcat declarations were incorrect. Some additional `[ TypeHint ]` attributes add.

**Update (v4.14.184):** Added SxS Assembly interfaces and APIs. Added MAKEINTRESOURCE macro. Added additional error messages. Made TaskDialogIndirect returns Optional per MSDN.

**Update (v4.14.182):** Added missing kernel32 string functions. Added SUCCEEDED helper function.

**Update (v4.14.181):** Bug fix: CHARFORMAT2[A|W] was incorrectly declared.

**Update (v4.14.180):** Much more extensive coverage of PROPVARIANT and Variant helpers for supported VB types (use changetype first to use them with unsigned et al).

**Update (v4.14.178):** Added partial Virtual Disk APIs and unsigned PROPVARIANT helpers.

**Update (v4.13.177):** Bug fix: Helper function UI_HSB had a syntax error.

**Update (v4.13.175):** Bug fix: UI Ribbon IIDs were missing.

**Update (v4.13.174):** Added caret APIs. Bug fix: Certain DirectWrite interfaces had members incompatible with x64. *IMPORTANT:* Having a single format for both 32 and 64bit breaks compatibility with the 32bit-only version. Previously `DWRITE_TEXT_RANGE` arguments were passed as two separate arguments, you'll now need to copy them to a single LongLong to pass.

**Update (v4.12.172):** User info APIs added.

 **Update (v4.12.170):** Bug fix: IOleInPlaceSite::Scroll scrollExtant should be ByVal. Added common error consts w/ descriptions. (171 is a version number change only for testing the package manager).
 
**Update (v4.12.166):**
-Added HTMLHelp APIs and misc ones that should be grouped with existing sets.

-New option: tbShellLib now has a 'Lite mode' designed to increase performance for users who typically define APIs themselves. In this mode, all API definitions in slAPI and slAPIComCtl are excluded, as are all misc API enums/types/consts in slDefs, and mPKEY.

-To use Lite mode, go to your project settings, go to 'Project: Conditional compilation constants', ensure it's checked to enable, and add `TB_SHELLLIB_LITE = 1`. This is applied during design time, so provides large benefits for Intellisense, syntax coloring/completion, etc.

**Update (v4.11.164):** Added Sensor APIs and Location APIs, including all related GUIDs/PKEYs from sensors.h. Added some APIs that belong with the previously added ones; major additions are likely over for now. Misc bugfixes to APIs.

**Update (v4.10.160):** Added IStorageProviderHandler and IStorageProviderPropertyHandler. Substantial updates to API sets.

**Update (v4.9.154):** Updated WebView2 interface set to latest stable release, v1.0.1774.30. Added additional APIs, focusing on Setup APIs, NTDLL, and data protection APIs. 

**Update (v4.8.147):** The OPENFILENAME[A,W] definitions were, inexplicably, still incorrect even though I thought I modified them when I made the issue for the pending fix.

**Update (v4.8.146):** The Common Controls API set did not conform to the project API standards at all; sometimes even within a single control's definitions. Be mindful if you've been using untagged aliases of A/W here. Numerous other small bug fixes. Many additional APIs.

**Update (v4.7.144):** Numerous bug fixes, including changing all olepro32.dll APIs to oleaut32, as the former doesn't exist in 64bit Windows and the functions have been exported by the latter since Win2k. Also added another large batch of APIs, with a focus on GDI drawing.

**Update (v4.6.142):** Some improvements/fixes to certain argument types in DirectX ifaces. Added a large number of font and text APIs in preparation for an upcoming project.

**Update (v4.6.139):** Bug fix: DirectComposition uses numerous overloaded methods; it's apparently an undocumented compiler behavior that these appear in reverse order from their declarations in the v-table, so the order had to be swapped for all overloads. These are currently uniquely named rather than taking advantage of tB's overloading supporting until I hear back from Wayne about the internals of support/implementation for it.

**Update (v4.6.138):** Several bug fixes, added misc commonly used APIs so far overlooked, and a number of additional APIs, focusing on registery, setup apis, and display settings apis.

**Update (v4.6.134): Critical bug fix:** A second `WM_USER` was accidentally made Public, which would cause numerous ambiguity and constant expression errors in any project using it or a constant derived from it. Also added keyboard APIs and some misc common ones that had been overlooked.

**Update (v4.6.132):** Numerous bug fixes related to string handling (ByRef LongPtrs that should have been ByVal), added another large batch of APIs.

**Update (v4.5.130):** Some minor bug fixes, added IInputPaneAnimationCoordinator, added another batch of APIs (focused on GDI, thread synchronization, and activation contexts). 

**Update (v4.5.128):** A number of DirectX interfaces were incompatible with x64 due to ByVal UDTs; these were imported from VB6 declares as e.g. 2 ByVal Longs for a point, but that won't work on x64 because of an 8 byte stack alignment. To keep codebases simple, points now use a single LongLong for *both* 32 and 64 bit. You declare a LongLong to pass, then use CopyMemory to copy your D2D1_POINT_F or other type into it. Also added some more APIs.

**Update (v4.5.126):** Added DirectComposition Presentation Manager interfaces, added additional APIs (focused on window management and file i/o), some minor bugfixes.

**Update (v4.4.124):** Important bug fixes and additional APIs (GDI printing and window transparency).

**Update (v4.4.122):**

-Critical bug fix for new tB builds (correctly) flagging Optional UDTs as errors. 

-Added UI Ribbon interfaces, coclasses, and PKEYs. (UIRibbon.h).

-Added interface IContextCallback with coclass ContextSwitcher (and related APIs).


**Update (v4.3.120):** 

-Added Disk Quota interfaces IDiskQuotaControl (with coclass DiskQuotaControl), IDiskQuotaUser, IDiskQuotaUserBatch, IEnumDiskQuotaUsers, and IDiskQuotaEvents.

-Bug fixes for certain `Optional` issues

-Added missing Direct2D flag to enable color fonts

-Expanded APIs focusing on subclassing, file mapping, memory management, and NT objects.

**Update (v4.3.114):** Important bug fixes for CreateThread ([#14](https://github.com/fafalone/tbShellLib/issues/14)), other bug fixes including IDataObject::DAdvise sink arg, and additional APIs.

**Update (v4.3.112):** Added some base OLE/COM interfaces I feel were substantial oversights from both olelib and oleexp; IDataAdviseHolder, IOleAdviseHolder, IDropSourceNotify, IEnterpriseDropTarget, and IContinue.

**Update (v4.3.102):** 

-Added some missing base OLE/COM interfaces: IQuickActivate, IAdviseSinkEx, IPointerInactive, IOleUndoManager, IEnumOleUndoUnits, IOleParentUndoUnit, IOleUndoUnit, IViewObjectEx, IOleInPlaceSiteWindowless, IOleInPlaceSiteEx, IOleInPlaceObjectWindowless.

-Additional APIs, focused on desktop/winstation APIs and DPI awareness APIs.

**Update (v4.2.98):** Numerous new APIs; some minor bugfixes.

**Update (v4.2.96):** Added missing Core Audio interfaces/GUIDs. Significant API coverage expansion.

**Update (v4.1.94):** Added Packaging API interfaces (msopc.idl). Added Netaddress control defs (newer version of old IP address control, msctls_netaddress; the old one, SysIPAddress32, is still there).

**Update (v4.0.93):** `Currency` in new interfaces changed to `LongLong`. 

**Update (v4.0.92):** 

-Completed Media Foundation interfaces up through the most recent Windows 11 SDK. This includes the capture engine and other entirely new feature sets.

-Added CoreAudio Spatial Audio interfaces (newer Win10 versions/Win11 only)

-Added IPropertyPage[2] and IPropertyPageSite interfaces.

-Added ISimpleFrameSite interface

-Bug fix: AUDCLNT_RETURNCODES were all incorrect.

---

For earlier version history, see CHANGELOG.md

For more information and a list of available interfaces, visit the [VB Forums thread](https://www.vbforums.com/showthread.php?897883-twinBASIC-tbShellLib-Shell-Interface-Library-(x64-compatible-successor-to-oleexp)) for this project.
