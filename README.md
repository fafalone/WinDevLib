# tbShellLib
**twinBASIC Shell Library**

Current Version: 4.3.102 (April 14th, 2023)

(c) 2022-2023 Jon Johnson (fafalone)

This project is a comprehensive twinBASIC replacement for [oleexp.tlb](http://www.vbforums.com/showthread.php?786079-VB6-Modern-Shell-Interface-Type-Library-oleexp-tlb), my Modern Shell Interfaces Type Library project for VB6, that is x64 compatible, due to the many problems using midl to create a 64bit tlb.

All interfaces, types, consts, and APIs from oleexp are covered, and there's additional API coverage not included in oleexp. For a full list of interfaces, see [INTERFACES.md](https://github.com/fafalone/tbShellLib/blob/main/INTERFACES.md).

This project is implemented purely in tB native code, as unlike VB6 there's language support for defining interfaces and coclasses. As a twinPACKAGE, regular code is supported in addition to the definitions, so the regular addin modules have been built in (mIID.bas, mPKEY.bas, etc).

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

It's fairly simple to move your VB6 projects to tbShellLib, just follow these steps:

1) Replace public aliases: It's important to do this first. Run a Replace All changing oleexp.LongPtr to LongPtr, oleexp.LONG_PTR to LongPtr, oleexp.REFERENCE_TIME to LongLong, oleexp.HNSTIME to LongLong, and oleexp.KNOWNFOLDERID to UUID. For all except the first, if you've used them without the oleexp. prefix, you'll also need to replace those.

2) Replace oleexp.IUnknown with IUnknownUnrestricted. tbShellLib keeps this separate due to the major issues with conflicts with the former approach. If your project has IUnknown *without* oleexp. in front of it, **do not** replace those, as it's not referring to oleexp. 

3) After you've done those two, you can now go ahead and simply delete all remaining instances of `oleexp.` (including the .). 

4) Manually address any errors remaining. Interfaces should be mostly fine at this point, but if you've made use of the APIs in oleexp, many of them have syntax differences, mainly not being able to rewrite an ending [out] argument as the return value, and changing String arguments to LongPtr you'll need StrPtr with. Note that while most APIs have been converted to use Unicode as the default, this is done with tB's `DeclareWide` keyword, so the types are still `String`, so you don't need to change anything from legacy oleexp/olelib code.

Note that this is just for using tbShellLib-- you'll likely have a lot more changes to make if you want to make your project x64 compatible.

## Guide to switching from oleexpimp.tlb

There's 'twinBASIC Shell Library for Implements' (tbShellLibImpl.twinpack/.twinproj) as well, but you'll note it has substantially fewer interfaces than oleexpimp.tlb. This is because there's two reasons for an interface to have an alternate version: It uses `[ Preservesig ]` on one or more methods, or it uses `As Any`. twinBASIC allows using `Implements` with `As Any` by replacing it with `As LongPtr` (which is what the alternate versions do). So many interfaces were in oleexpimp.tlb for this latter reason, and subsequently are *not* included in tbShellLibImpl as it's not neccessary.

If you find an oleexpimp.tlb interface is not in tbShellLibImpl, you will be able to use the one from tbShellLib, simply make sure `As Any` is changed to `As LongPtr`. 

tB has announced plans to support `[ PreserveSig ]` in implemented interfaces in the future; when that happens tbShellLibImpl will be deprecated.

## Updates
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

**Update (v3.12.88):** Added misc. interfaces IDelayedPropertyStoreFactory, IStorageProviderCopyHook, IDesktopGadget/Coclass DesktopGadget, IQueryCodePage, IStreamUnbufferedInfo, IUserAccountChangeCallback, IOpenSearchSource, IDestinationStreamFactory, ICreateProcessInputs, and ICreatingProcess. Continued adding APIs and Media Foundation interfaces.

**Update (v3.11.84, v3.11.86):** Additional APIs and Media Foundation stuff.

**Update (V3.11.82):** Additional API expansion for upcoming projects. Added Media Foundation / D3D12 sync interfaces/GUIDs. Added Media Foundation Capture Engine interfaces/GUIDs. Realized I actually have a ton more Media Foundation stuff not yet included. 

**Update (v3.10.80):** Additional API expansion for upcoming projects.

**Update (v3.10.72):** Added a number of important APIs for upcoming projects. Added EP_* GUIDs for IExplorerPaneVisibility, added some missing SID_ guids.
**tbShellLib (v1.2.7):** Added IMessageFilter. **NOTE: tbShellLibImpl IS NOW WORKING!** I hadn't realized the old VSCode plugin was continually refusing to save settings, thus ignoring the setting to disable the autoprettifying that didn't understand interfaces and thus ran together the declares, making them invalid. This has been fixed in 1.2.7.

**Update (v3.9.70):** Reworked APIs to be more consistent when there's A/W versions. For most of these APIs, tbShellLib offers 3 versions: An explicit A version, an explicit W version, and an undecorated version that uses `DeclareWide` and `String` that's an alias for the W version. Some of the more advanced/newer APIs don't have the ANSI version declared. For APIs from oleexp/olelib without A/W but accepting strings, they've been left as `LongPtr`, but new ones added will use String. Also continued to add new APIs.

**Update (v3.8.66):**

-Added IActiveScript and all related ActiveX Script Host / Engine interfaces

-Added IDispatchEx interface and related interfaces IDispError, IVariantChangeType, IProvideRuntimeContext, IObjectIdentity, and ICanHandleException

-Added IFileSearchBand, coclass FileSearchBand

-Corrected some Direct3D type names that got caught up in an autoreplace unintentionally.

-Misc bug fixes (Long->LongPtr, APIs pointing to wrong A/W version, missing A/W versions)

-Continued expanding API coverage.

**Update (v3.7.62):** Added all remaining missing oleexp interfaces simply for completeness and not needing to qualify 'contains everything in oleexp'. IHostDialog/coclass HostDialog seemed like a major omission from those legacy interfaces so added it. Continued to substantially expand API coverage.

**Update (v3.6.56):** Numerous bug fixes to IPinnedList[2,3], including their names: IPinnedListVista->IPinnedList, IPinnedList->IPinnedList2, IPinnedList10->IPinnedList3, to be more consistent with all other sources. Please do not abuse these interfaces: Never pin without permission. Added IWinEventHandler, IFolderBandPriv, and IAccessibleObject; added coclass TaskBand, and added numerous missing IIDs.

**Update (v3.6.54):** Some items were Private that should have been Public; put SW_Flags back to SHOWWINDOW now that bug is resolved for compatibility purposes (SHOWWINDOW is in oleexp). To use this, twinBASIC Beta 269 or newer is needed. Misc bug fixes.

**Update (v3.6.52):**

-By popular request to expand the API coverage, tbShellLib now has had tbComCtlLib merged into it. You can exclude these definitions with the TB_COMCTL_LIB_DEFINED compiler constant. 

-Substantially expanded general API coverage.

-Misc bugfixes including renaming SHOWWINDOW enum to SW_Flags to work around a tB bug. 

For earlier version history, see CHANGELOG.md

For more information and a list of available interfaces, visit the [VB Forums thread](https://www.vbforums.com/showthread.php?897883-twinBASIC-tbShellLib-Shell-Interface-Library-(x64-compatible-successor-to-oleexp)) for this project.
