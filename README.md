# tbShellLib
**twinBASIC Shell Library**

Current Version: 4.12.166 (July 3rd, 2023)

(c) 2022-2023 Jon Johnson (fafalone)

This project is a comprehensive twinBASIC replacement for [oleexp.tlb](http://www.vbforums.com/showthread.php?786079-VB6-Modern-Shell-Interface-Type-Library-oleexp-tlb), my Modern Shell Interfaces Type Library project for VB6, that is x64 compatible, due to the many problems using midl to create a 64bit tlb.

This and oleexp are projects to supply Windows shell and component interfaces in a format consumable by VB6/VBA/tB. This involves not only defining interfaces, but using VB/tB compatible types-- so in some cases, even though there may be an existing way to import references to interfaces, they may be unusable due to e.g. the use of unsigned types, C-style arrays, double pointers, etc.

All interfaces, types, consts, and APIs from oleexp are covered, and there's additional API coverage not included in oleexp. For a full list of interfaces, see [INTERFACES.md](https://github.com/fafalone/tbShellLib/blob/main/INTERFACES.md).

This project is implemented purely in tB native code, as unlike VB6 there's language support for defining interfaces and coclasses. As a twinPACKAGE, regular code is supported in addition to the definitions, so the regular addin modules have been built in (mIID.bas, mPKEY.bas, etc). Does it still make sense to use a project like this when interfaces can be defined in-language? I'd say yes, because for a large number of interfaces, there's deep dependency chains with other interfaces and the types they rely on. It makes more sense to drop this in and be done with it than constantly have to define the interfaces you want and then stubs for their dependencies, especially when you might need those later on. This project is even more useful now with the API coverage; it should cover about 99% of your needs for system DLLS. 

Please report any bugs via the Issues feature here on GitHub.

## Requirements

[twinBASIC Beta 269 or newer](https://github.com/twinbasic/twinbasic/releases) is required.

## Adding tbShellLib to your project
You have 2 options for this:

#### Via the Package Server
twinBASIC has an online package server and tbShellLib is published on it. Open your project settings and scroll to the **COM Type Library / ActiveX References**, then click **TWINPACK PACKAGES**. Add "twinBASIC Shell Library v3.4.46" (or whatever the newest version is). "twinBASIC Shell Library for Implements" contains `Implements` compatible versions of a small number of common interfaces not defined in a compatible way in the main project; you normally don't need this. For more details, including illustrations, [see this post](https://github.com/fafalone/tbShellLib/issues/9#issuecomment-1416767019).

#### From a local file
You can download the project from this repository and use the .twinpack file. Navigate to the same area as above, and click on the "Import from file" button. 

## Guide to switching from oleexp.tlb

tbShellLib presented the best opportunity there would be to ditch some olelib legacy baggage. It's fairly simple to move your VB6 projects to tbShellLib, just follow these steps:

1) Replace public aliases: It's important to do this first. Run a Replace All changing oleexp.LONG_PTR to LongPtr, oleexp.REFERENCE_TIME to LongLong, oleexp.HNSTIME to LongLong, oleexp.KNOWNFOLDERID to UUID, oleexp.EventRegistrationToken to LongLong, oleexp.BINDPTR to LongPtr, and oleexp.LPCRITICAL_SECTION to LongPtr. If you've used them without the oleexp. prefix, you'll also need to replace those, but if you've imported into tB they should be tagged.

2) Replace oleexp.IUnknown with IUnknownUnrestricted. tbShellLib keeps this separate due to the major issues with conflicts with the former approach. If your project has IUnknown *without* oleexp. in front of it, **do not** replace those, as it's not referring to oleexp. 

3) After you've done those two, you can now go ahead and simply delete all remaining instances of `oleexp.` (including the .). 

4) Convert `Currency` to `LongLong` for interfaces and APIs: It's no longer neccessary to worry about multiplying and dividing by 10,000 since tB supports a true 64bit type in both 32bit and 64bit mode. So this change is ultimately for the better, but existing codebases will have had to have used `Currency` for all interfaces and oleexp APIs expecting a 64bit integer.

5) Manually address any errors remaining. Interfaces should be mostly fine at this point, but if you've made use of the APIs in oleexp, many of them have syntax differences, mainly not being able to rewrite an ending [out] argument as the return value, and changing String arguments to LongPtr you'll need StrPtr with. Another major difference is that the default for almost all APIs with ANSI/Unicode (A/W) versions, is now the Unicode version. A notable exception is `SendMessage` due to the overwhelming amount of VBx code expecting it to mean `SendMessageA`. In most cases, the W version is declared with `LongPtr` for strings, and the untagged alias version uses tB's new `DeclareWide` keyword to disable ANSI conversion while using `String`.

Note that this is just for using tbShellLib-- you'll likely have a lot more changes to make if you want to make your project x64 compatible.

## Guide to switching from oleexpimp.tlb

There's 'twinBASIC Shell Library for Implements' (tbShellLibImpl.twinpack/.twinproj) as well, but you'll note it has substantially fewer interfaces than oleexpimp.tlb. This is because there's two reasons for an interface to have an alternate version: It uses `[ Preservesig ]` on one or more methods, or it uses `As Any`. twinBASIC allows using `Implements` with `As Any` by replacing it with `As LongPtr` (which is what the alternate versions do). So many interfaces were in oleexpimp.tlb for this latter reason, and subsequently are *not* included in tbShellLibImpl as it's not neccessary.

If you find an oleexpimp.tlb interface is not in tbShellLibImpl, you will be able to use the one from tbShellLib, simply make sure `As Any` is changed to `As LongPtr`. 

tB has announced plans to support `[ PreserveSig ]` in implemented interfaces in the future; when that happens tbShellLibImpl will be deprecated.

## tbShellLib API standards

This was mentioned above, but it's worth going into more detail. In addition to the COM interfaces, tbShellLib has a large selection of common Windows APIs; this is a much larger set than oleexp. tbShellLib and twinBASIC represented the best opportunity there would be to modernize standards... most VB programs are written with ANSI versions of APIs being the default. **This is not the case with tbShellLib**. With very few exceptions, APIs are Unicode by default-- i.e. they use the W, rather than A, version of APIs e.g. `DeleteFile` maps to `DeleteFileW` rather than `DeleteFileA`. The A and W variants use String/LongPtr, and in almost all cases, the mapped version uses `String` with twinBASIC's `DeclareWide` keyword-- this disables Unicode-ANSI conversion, so you can still use `String` without `StrPtr` or any Unicode <-> ANSI conversion. Note this usually only applies to strings passed as input, APIs passing a LPWSTR that's allocated externally will still be LongPtr, as they're not in the same BSTR format as VBx/TB strings.

All APIs are provided, as a minimum, as the explicit W variant, and an untagged version that maps to the W version. Some, but not all, APIs also have an explicit A variant defined that will perform the normal ANSI conversion for compatibility purposes. This is decided on a case by case basis depending on my impression of how much legacy code is around that needs the ANSI version. All new code should use the Unicode versions.

UDTs used by these calls are also supplied in the same manner, the W variant, an untagged variant that's the same as the W version, and in some cases, an A version. UDTs always use `LongPtr` for strings, even the untagged versions for `DeclareWide`. 

As noted before, an exception to the rule is `SendMessage`, due to the enourmous volume of existing code expecting SendMessage to map to SendMessageA.

If you have any doubts about which API is being called, twinBASIC will show the full declaration when you hover your cursor over the API in your code.

### A note on seeing UDTs where before they were As Any

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

## Updates
**Update (v4.12.166):**
-Added HTMLHelp APIs and misc ones that should be grouped with existing sets.

-New option: tbShellLib now has a 'Lite mode' designed to increase performance for users who typically define APIs themselves. In this mode, all API definitions in slAPI and slAPIComCtl are excluded, as are all misc API enums/types/consts in slDefs, and mPKEY.

-To use Lite mode, go to your project settings, go to 'Project: Conditional compilation constants', ensure it's checked to enable, and add `TB_SHELLLIB_LITE = 1`.

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
