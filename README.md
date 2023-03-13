# tbShellLib
**twinBASIC Shell Library**

Current Version: 3.5.48 (March 9th, 2023)

(c) 2022-2023 Jon Johnson (fafalone)

This project is a comprehensive twinBASIC replacement for [oleexp.tlb](http://www.vbforums.com/showthread.php?786079-VB6-Modern-Shell-Interface-Type-Library-oleexp-tlb), my Modern Shell Interfaces Type Library project for VB6, that is x64 compatible, due to the many problems using midl to create a 64bit tlb.

All interfaces are covered except ones useless on anything newer than Win9x, and all APIs are covered.

This project is implemented purely in tB native code, as unlike VB6 there's language support for defining interfaces and coclasses. As a twinPACKAGE, regular code is supported in addition to the definitions, so the regular addin modules have been built in (mIID.bas, mPKEY.bas, etc).

Please report any bugs via the Issues feature here on GitHub.

## Requirements

[twinBASIC Beta 239 or newer](https://github.com/twinbasic/twinbasic/releases) is required.

## Adding tbShellLib to your project
You have 2 options for this:

#### Via the Package Server
twinBASIC has an online package server and tbShellLib is published on it. Open your project settings and scroll to the **COM Type Library / ActiveX References**, then click **TWINPACK PACKAGES**. Add "twinBASIC Shell Library v3.4.46" (or whatever the newest version is). "twinBASIC Shell Library for Implements" contains `Implements` compatible versions of a small number of common interfaces not defined in a compatible way in the main project. For more details, including illustrations, [see this post](https://github.com/fafalone/tbShellLib/issues/9#issuecomment-1416767019).

#### From a local file
You can download the project from this repository and use the .twinpack file. Navigate to the same area as above, and click on the "Import from file" button. 

## Guide to switching from oleexp.tlb

It's fairly simple to move your VB6 projects to tbShellLib, just follow these steps:

1) Replace public aliases: It's important to do this first. Run a Replace All changing oleexp.LongPtr to LongPtr, oleexp.LONG_PTR to LongPtr, oleexp.REFERENCE_TIME to LongLong, oleexp.HNSTIME to LongLong, and oleexp.KNOWNFOLDERID to UUID. For all except the first, if you've used them without the oleexp. prefix, you'll also need to replace those.

2) Replace oleexp.IUnknown with IUnknownUnrestricted. tbShellLib keeps this separate due to the major issues with conflicts with the former approach. If your project has IUnknown *without* oleexp. in front of it, **do not** replace those, as it's not referring to oleexp. 

3) After you've done those two, you can now go ahead and simply delete all remaining instances of `oleexp.` (including the .). 

4) Manually address any errors remaining. Interfaces should be mostly fine at this point, but if you've made use of the APIs in oleexp, many of them have syntax differences, mainly not being able to rewrite an ending [out] argument as the return value, and changing String arguments to LongPtr you'll need StrPtr with.

Note that this is just for using tbShellLib-- you'll likely have a lot more changes to make if you want to make your project x64 compatible.

## Updates
**Update (v3.5.48):**

-Added accessibility UI Automation interfaces and APIs. 

---NOTE: This API had a number of *very* generically named enums, like FillMode and ToggleState; these have been prefixed with Uia_ to avoid conflicts. In most cases, the actual members were left   alone, with the exception of LiveSetting (renamed Uia_LiveSetting), which had Off, Polite, and Assertive; these have been prefixed with Uia_ as well.
  
---NOTE: IUIAutomation, IUIAutomationProxyFactoryMapping, IUIAutomationAndCondition and IUIAutomationOrCondition have members that use a SAFEARRAY of IUIAutomationCondition... MKTYPLIB does not support this so these return a pointer you'll need to dereference.

-The package now includes a common helper function for interfaces: SwapVTableEntry, updated for use in both 32bit and 64bit mode.

-Added misc interfaces ICurrentWorkingDirectory, IPropertyKeyStore, ISortColumnArray, and IBannerNotificationHandler

-Added IHandlerInfo2, IDeskBar, IDeskBarClient amd IShellFolderBand

-Began expanding general API coverage

**Update (v3.4.46):** Added all GDIPlus APIs and all Common Dialog APIs.

**Update (v3.3.41):** Bug fix: IExplorerBrowserEvents::NavigationFailed was misspelled.

**Update (v3.3.40):**
-Inexplicably, the IDeskBand, IDockingWindow, IDockingWindowFrame, and IDockingWindowSite interfaces were missing.

-Added ITrayBand, IDeskBand2, IDeskBandInfo, IBandHost, and IBandSite interfaces, IMenuBand, and coclasses TrayDeskBand, TrayBandSiteServices, and AddressBand.

-Added IRegTreeItem interface

-Added IPrintDialogCallback/IPrintDialogServices interfaces.

-Bug fix: Certain DirectWrite interface members had ByRef Long for strings where they should have had ByVal.

-Bug fix: SHELLSTATE had an extra member on the end (shouldn't have impacted use, but if MS changed the API to look for an exact size it would be an issue).

-Bug fix: Attempted to correct INameSpaceTreeControlEvents context menu crashing.

-Bug fix: INameSpaceTreeCustomDraw::ItemPrePaint was missing members.


**Update (v3.2.30):** Several Speech API interfaces were missing. Also, began using BOOL type as as enum with CFALSE (0) and CTRUE (1) members. I'll be slowly working on changing all the Long items that are actually BOOL to this over the coming months.

**Update (v3.2.24):** Numerous bugfixes to Speech API interfaces.

**Update (v3.2.22):** Missed some WebView2 interfaces that should have LongPtr instead of String; changed all String args to LongPtr just to be safe.

**Update (v3.2.20):** Bug fix: [out] LPWSTR* and [in] LPWSTR for Implements interfaces in WebView2 args changed to LongPtr. IEnumVARIANT was missed; added. 

**tbShellLibImpl (v1.2.6):** Also was missing `Implements` version of IEnumVARIANT.


**Update (v3.2.16):** Added WebView2 (EXPERIMENTAL). Added IObjectProvider, IEnumObjects, and IIOCancelInformation interfaces.

**Update (v3.1.14):** Added Microsoft Speech APIs v5.4. Added IHttpNegotiate3.

**Update (v3.0.13):**
-Added missing PROPSHEETHEADER and PROPSHEETHEADER_V2 types and PropSheet/PropSheetW APIs. Also corrected wrong values for PSN_TRANSLATEACCELERATOR/PSN_QUERYINITIALFOCUS.

-Began adding back in some Optionals in DirectX interfaces which weren't supported by MKTYPLIB so weren't in oleexp, where the tB code was generated from.

-(Bug fix) StringFromGUID2 now uses a Long instead of LPWSTR since the latter was not working.

-(Bug fix) D3D11CreateDevice and D3D11CreateDeviceAndSwapChain were declared incorrectly for 64bit compatibility (Softare param should be LongPtr).

**Update (v3.0.10):** Added all missing Direct2D interfaces/types/enums and corrected bugs in slDirectX.

**Update (v2.9.90):** EXPERIEMENTAL: Added Direct3D 11 and 12.

**Update (v2.9.85):** 

-Bug fix: ITypeInfo::AddressOfMember returned Long instead of LongPtr ([#11](https://github.com/fafalone/tbShellLib/issues/11)); ICreateTypeLib2 incorrectly extended IUnknown instead of ICreateTypeLib, and other misc bugfixes.

-Added objidl.idl interfaces IAdviseSink2, IClientSecurity, IServerSecurity, IMallocSpy, IClassActivator, IProgressNotify, IStdMarshalInfo, IExternalConnection and IThumbnailExtractor (w/coclass ThumbnailFCNHandler).

-Added undocmented hardware enum interfaces/coclasses.

**Update (v2.9.81):** Bug fix: ITypeLib::GetTypeInfoCount and several others never had `[ PreserveSig ]` restored after support was added.

**Update (v2.9.80):** Substantially expanded Media Foundation set, also now includes all related GUIDs defined in mfidl.idl. Fixed incorrect IID for ITypeInfo/IID_ITypeInfo. 

**Update (v2.8.76):** Added basic Media Foundation interfaces from oleexp that were missing up until now. **(v2.8.78):** Fixed missing PtrSafe attributes and SwapVTable type errors.

**Update (v2.8.75):** Added DXGI and DirectComposition interfaces (experimental).

**Update (v2.7.70):** Shell automation intefaces using VARIANT_BOOL have been changed to Boolean to be more correct than Integer (the underlying typedef is `short`, which is why is was Integer at first).

**Update (v2.7.66):**

-HIGHLY EXPERIMENTAL: Added DirectWrite and Direct2D interfaces via merging d2dvb.tlb/dwvb.tlb by [@Thetrik](https://github.com/thetrik) with oleexp, then copying out of tB's typelib browser. I've done an initial review to find pointer types, but may have missed some, and the defs are extremely VB-hostile and even more MKTYPLIB hostile so there may be function prototype issues. These are all in slDirectX.twin if you wanted to remove it from a build.

-Added Windowless RichEdit interfaces (ITextServices[2], ITextHost[2], IRicheditUiaOverrides)

-Bug fix: CreateTypeLib definition incorrect ([#8](https://github.com/fafalone/tbShellLib/issues/8))

-(tbShellLibImpl) Updated to v1.1.4- added `Implements` compatible versions of ITextHost[2]. WARNING: The signatures are incorrect as `[ PreserveSig ]` is not optional. You will need to swap the vtable entries to a correct prototype. These are provided merely to get around the invalid signature compiler error.

**Update (v2.6.64):** IWebBrowserApp managed to escape all the replacements I ran to add `[ OleAutomation(False) ]`

**Update (tbShellLibImpl v1.0.3):** Updated internal tbShellLib reference.

**Update (v2.6.62):** Corrected remaining hex literals that would be interpreted incorrectly.


**Update (v2.6.60):**

-**IMPORTANT:** tbShellLib now requires [twinBASIC Beta 239 or newer](https://github.com/twinbasic/twinbasic/releases). This is due to the requirement to use the new `[ OleAutomation(False) ]` attribute in order for interfaces to be used in compiled Active-X controls. All tbShellLib interfaces have been marked this way. This should not impact regular usage or UserControl usage; if it does, please create an issue.

-Completed Text Object Model interfaces as of Win10 (TOM.h).

-Added interfaces IObjectWithAppUserModelID, IObjectWithProgID, IObjectWithCancelEvent, IObjectWithSelection, and IObjectWithBackReferences. 

-Added interface IRemoteComputer

-Added interface IUpdateIDList

-Added interfaces IAccessibilityDockingService and IAccessibilityDockingServiceCallback with coclass AccessibilityDockingService

-Bug fix: [#7](https://github.com/fafalone/tbShellLib/issues/7) FILE_ATTRIBUTE_PINNED incorrectly defined


**Update (v2.5.56):**

-Added Sync Manager interfaces and coclasses (SyncMgr.h), including undocumented ITransferConfirmation/coclass TransferConfirmationUI.

-Added interfaces IPersistSerializedPropStorage, IPersistSerializedPropStorage2, and IPropertySystemChangeNotify

-Added missing propsys coclasses CLSID_InMemoryPropertyStore, CLSID_InMemoryPropertyStoreMarshalByValue, CLSID_PropertySystem

-Added IListViewVista interface (Vista-only version of IListView)

-Added IPinnedList with variants IPinnedListVista (Windows Vista) and IPinnedList10 (Windows 10 build 1809 and newer). Also added IStartMenuPin, ITrayNotify and INotificationCB. These are undocumented taskbar interfaces for programmatically pinning items to the start menu and taskbar. Added TaskbandPin, TrayNotify, and StartMenuPin coclasses (the last one is officially documented for the IStartMenuPinnedList interface with remove pin only, but it implements the undocumented pinning interfaces too and those have been added to the supported list). 

-Bug fix: Numerous enum values defined incorrectly as &H8000, causing sign issues in bitwise operations and downstream issues from that,

---

Update (tbShellLibImpl v1.0.2): CRITICAL BUG FIX: IFolderView was missing GetDefaultSpacing, breaking any use of it and IFolderView2.

Update (v2.4.49): IShellView::TranslateAccelerator was incorrectly named IShellView::TranslateAcceleratorSB.

**Update (v2.4.48):** 

-CRITICAL BUG FIX: IFolderView was missing GetDefaultSpacing, breaking any use of it and IFolderView2.

-Bug fix: IsEqualIID API declare was not marked PtrSafe.

-IServiceProvider did not use PreserveSig in the original oleexp, so that has been changed to match here, for use with Implements.

-Added IShellUIHelper[2,3,4,5,6,7,8,9], IShellFavoritesNameSpace, IShellNameSpace, IScriptErrorList and related coclasses.

-Added IDesktopWallpaper with coclass DesktopWallpaper

-Added IAppVisibility and IAppVisibilityEvents

-Added coclass AppStartupLink

-Added IApplicationActivationManager with coclass ApplicationActivationManager

-Added IContactManagerInterop, IAppActivationUIInfo, IHandlerActivationHost, IHandlerInfo, ILaunchSourceAppUserModelId, ILaunchTargetViewSizePreference, ILaunchSourceViewSizePreference, ILaunchTargetMonitor, IApplicationDesignModeSettings, IApplicationDesignModeSettings2, IExecuteCommandApplicationHostEnvironment, IPackageDebugSettings, IPackageDebugSettings2, IPackageExecutionStateChangeNotification, IDataObjectProvider, IDataTransferManagerInterop.
-Added coclasses for above: PackageDebugSettings, SuspensionDependencyManager, ApplicationDesignModeSettings


**tbShellLibImpl (v1.0)**: tbShellLib for Implements initial release. This does not cover all of oleexpimp.tlb because there's no need for an out only vs in, out distinction which many had as the only difference.

**Update (v2.3.44):** ICategoryProvider and ICategorizer had BSTR instead of LPWSTR (LongPtr) arguments.

**Update (v2.3.40):** Fixes for SHGetPathFromIDList[W] and IVirtualDesktopManager::IsWindowOnCurrentVirtualDesktop.

**Update (v2.3.38):** ICategorizer::GetCategory had apidl argument incorrectly defined as ByVal.

**Update (v2.3.35):** IShellIconOverlay had incorrect pIndex params in both methods. This didn't effect 32bit projects as pointers were the same size as the index. 

**Update (v2.3.32):** Fixed GWL_* duplicate error and LARGE_INTEGER restored to hipart/lowpart for compatibility; ULARGE_INTEGER still uses quadpart if desired.

**Update (v2.3.30):** Fixed CM_COLUMNINFO bug since it was causing SetColumnInfo to trigger an automation error.

**Update (v2.3.26):** Minor bug fixes.

**Update (v2.2.24):** Added IWebBrowser2 interface I thought was already there.

**Update (v2.1.24):** twinBASIC now supports in-project `CoClass` syntax! All coclasses from oleexp have been added (I think, if you find one missing please create an issue), and can once again be used with the New keyword. The prior sCLSID constants have been left in. Also greatly expanded the API declare coverage to match what was in oleexp, though a few DLLs are still pending. Finally, tbShellLib now declares an compiler constant, `TB_SHELLLIB_DEFINED`, to help avoid conflicts with other projects (chiefly, my upcoming Common Controls 64-bit compatible library). *tbShellLib now requires twinBASIC Beta 167 or newer*.

**Major Update (v2.0.20):**  The project has reached it's initial goal of implementing all but the most obsolete oleexp.tlb interfaces. In addition, with similar exception of a small set of highly obsolete items, the API coverage is now available. Note that this was subject to extensive cleanup; native-language declares can't use the last param as retval on APIs, so all of those were converted, and TLB APIs pass Strings as BSTR, while native language passes ANSI strings, so there's currently a mix of either using LongPtr or tB's DeclareWide for BSTR/LPWSTR support (if it says String you can use a String without StrPtr). 

Update (v1.9.17): Extensive new interface additions; all remaining oleexp additions have been added including WIC and NetCon, and the majority of remaining original olelib interfaces have been added as well. 

Update (v1.2.10): twinBASIC now supports [ PreserveSig ] as an attribute to have HRESULT values as a function return instead of only available via Err.LastHResult; tbShellLib now has this implemented wherever it is in oleexp. Like in VB, this means they're not Implements compatible, and at some point in the next few weeks, there will also be a tbShellLibImpl as a counterpart to oleexpimp.tlb. This update also adds all DirectShow interfaces (and mDirectShow.bas), most remaining oleexp interfaces, and several additional olelib interfaces.

Update (v1.1.8): Added class factory/typelib interfaces from olelib plus oleexp extensions; added manipulation.idl stuff (internial scrolling), and a few misc others.

Update (v1.1.6): Added a small number of interfaces I shouldn't have left out of the first major release... IShellExtInit, IShellExtPropPage (and related structs/apis), IQueryAssociations, IItemNameLimits, IObjectWithSite, a few others. 


For more information and a list of available interfaces, visit the [VB Forums thread](https://www.vbforums.com/showthread.php?897883-twinBASIC-tbShellLib-Shell-Interface-Library-(x64-compatible-successor-to-oleexp)) for this project.
