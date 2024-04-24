
**Update (v7.9.390, 24 Apr 2024):**
-Large expansion of security APIs from security.h, minschannel.h, sspi.h, issper16.h, and credssp.h
   All are 100% covered with the exception of kernel-mode only defs in sspi.h.
-Added new helper function for APIs/COM interfaces expecting a ByVal GUID:
   UUIDtoLong(UUID, pl1 As Long, pl2 As Long, pl3 As Long, pl4 As Long)
   UUIDtoLong(UUID, pls() As Long)
-Added VBA-related interfaces from vbinterf.h (100% coverage)
-Adjusted custom buffers on DEV_BROADCAST_* types to not leave padding bytes.
-Added non-aliased versions of RtlMoveMemory, RtlZeroMemory, and RtlZeroMemory (Issue #20)
-(Bug fix) LoadIconMetrics enum had incorrect values and is now also renamed 
'          to the proper LI_METRIC name.

**Update (v7.9.386, 19 April 2024):**
-Added complete Virtual Disk Service interfaces and custom coclass VdsLoader
   (vdserr.h, vdscmprv.idl, vsprvcm.idl, vdshwprv.idl, vdscmmn.idl, vdslun.idl, 
    vdssp.idl, vdshp.idl, vdsvd.idl, vds.idl, vdshpcm.idl; (from derivation, also
    vds.h, vdshwprv.h, vdslun.h, vdssys.idl/vdssys.h)-- 100%)
-Added DirectML interfaces (directml.h, 100%)
-Added Restart Manager APIs (restartmanager.h, 100% coverage)
-Added DDE APIs (dde.h, ddeml.h 100%)
-Added some misc missing extremely common APIs.

**Update (v7.8.382, 17 April 2024):**
-Added coverage of all Windows Biometric Framework application APIs (winbio_err.h, winbio_ioctl.h, winbio_types.h, winbio.h 100%)
-Added missing WMDM DRM interfaces/coclass (MS forgot to merge these into the SDK when it abandoned a separate WMDM sdk)
-Some additional defs to bring winsvc.h coverage to 100%
-Add some missing WIC GUIDs
-(Bug fix) SERVICE_REQUIRED_PRIVILEGES_INFO[W] definitions incorrect for 64bit
-(Bug fix) EnumServicesStatusEx, GetServiceDisplayName incorrect alias
-(Bug fix) QueryServiceStatusEx, QueryServiceDynamicInformation, GetServiceRegistryStateKey, GetServiceDirectory, GetSharedServiceDirectory, RegisterServiceCtrlHandler[A,W,Ex,ExA,ExW] definitions incorrect for 64bit (Ex incorrect alias as well)
-(Bug fix) QueryServiceStatusEx incorrect additional overload
-(Bug fix) SECURITY_MAX_SID_SIZE value incorrect

**Update (v7.8.379, 12 April 2024):**
-Large expansion of Direct3D 12 interfaces to cover latest SDK version of d3d12.idl
-Added Direct3D 12 Video interfaces
-Added some missing Direct2D and Direct3D 11 interfaces
-Added Windows Media Device Manager application interfaces (mswmdm.h, 50%- provider interfaces todo)
-Added cert signing APIs from Mssign32.dll (mssign.h, 100%)
-(Bug fix) GdipGetLineColors definition incorrect (Issue #18)
-(Bug fix) GdipDrawImagePointsRect[I] definitions incorrect for 64bit (Issue #19)
-(Bug fix) GdipEnumerateMetafileDestPoint[I] definitions incorrect for 64bit

**Update (v7.7.372, 09 April 2024):**
-Minor additions to bring coverage of shellapi.h to 100%
-Added macros/helpers from mfapi.h and mfplay.idl
-Add missing gdip function GdipDrawImageFX
-(Bug fix) GdipFillClosedCurve2[I] definitions incorrect. (Issue #17)

**Update (v7.7.370, 05 April 2024):**
-Added all Background Intelligent Transfer Service interfaces; 100% coverage of:
 bits.idl, bits1_5.idl, bits2_0.idl, bits2_5.idl, bits3_0.idl, bits4_0.idl, bits5_0.idl,
 bits10_1.idl, bits10_2.idl, bits10_3.idl, bitscfg.idl, qmgr.idl.

**Update (v7.7.360, 04 April 2024):**
-Very large expansion of DirectWrite interfaces; only dwrite.h was covered; added 100%
 coverage of dwrite_1.h, dwrite_2.h, and dwrite_3.h
-Added shdeprecated.h (100% coverage). Many of these are still in undocumented use.
-UserEnv.h expanded to 100% coverage
-Added crypto catalog APIs from mscat.h (100% coverage)
-(API Standards) GetClassInfo[A, ExA, Ex] did not conform to API standards. For compatibility,
                 this has been resolved by adding overloads.
-CreateProfile does not have A/W variants. I have *zero* idea where I found otherwise, and with
 differently named arguments... no search results anywhere. Weird.
-Add DWRITE_RENDERING_MODE missing values
                 
**Update (v7.7.350, 31 Mar 2024):**
-Large expansion of mfapi.h coverage; all APIs and GUIDs are covered, only missing the macros
-processenv.h coverage now 100%
-avrt.h 100% coverage in prep. for mfapi.h (limited current coverage)
-Added 100% cover of netioapi.h
-GetEnvironmentStrings now redirects to GetEnvironmentStringsW, per SDK.
-Added security center interfaces from iwscapi.h and APIs from wscapi.h (both 100% covered)
-Added WINDEVLIB_NOLIBS compiler option, completely disabling static library use (intended
 mainly to be able to test with tB Beta 423 or earlier)
-(Bug fix) SetCurrentDirectory[W] definitions incorrect.
-(Bug fix) Certain obscure PE header types missing alternate alignment attribute
-(Bug fix) GetNamedPipeClientComputerName[A.W] definitions incorrect
-(Bug fix) GetNamedPipeHandleState[A,W] definitions incorrect


**Update (v7.7.345, 26 Mar 2024):**
-Added tdh.dll event trace helper APIs (tdh.h; all APIs/types complete but macros not yet added)
-Added some additional native APIs.
-FlushViewOfFile was missing.
-(Bug fix) IMAGE_OPTIONAL_HEADER64 had an extra member and pointer member incorrectly declared as
           LongPtr, making the UDT offsets incorrect when handling a 64bit PE from a 32bit build.
-(Bug fix) The extra member mentioned above *is* in the 32bit version; so the build-linked verson
           (IMAGE_OPTIONAL_HEADER) had to have a conditional added.

**Update (v7.7.343, 22 Mar 2024):**
-(Bug fix) Coclass ActCtx conflicted with type ACTCTX; the former has been renamed CActCtx.
-(Bug fix) ReleaseActCtx had typo in name.

**Update (v7.7.342, 21 Mar 2024):**
-**MAJOR CHANGE:** The common used enum SHGDN_Flags has been renamed SHGDNF, the proper name per SDK.
-**MAJOR CHANGE:** The common used enum SVGIO_Flags has been renamed SVGIO, the proper name per SDK.
-**MAJOR CHANGE:** The common used enum SVSI_Flags has been renamed SVSIF, the proper name per SDK.
-Updated WebView2 to match current stable release 1.0.2365.46
-Filled out KUSER_SHARED_DATA more.
-(Bug fix) NET_ADDRESS_INFO union substitute sized incorrectly.


**Update (v7.7.341, 16 Mar 2024):**
-**MAJOR CHANGE:** The commonly used enum SFGAO_Flags has been renamed SFGAOF, in accordance with a
                   previously overlooked official name for the enum: `typedef ULONG SFGAOF;`
                   It is safe (as far as this package knows) to do a find/replace all for this.
                   Also added missing value SFGAO_PLACEHOLDER.
-For code portability, over the coming weeks and months I'll be replacing `DeclareWide` with `Declare`.
  This will only be done on functions where it doesn't matter; where no arguments or arg UDT members
  are `String`. It will still be used where it matters (especially in A/W functions without the A/W)
-Added missing winmm video/animation consts and structs
-Added helper function InitVariantFromIDList (undocumented inline helper)
-Added interfaces IWebBrowserEventsService, IWebBrowserEventsUrlService (WebEvnts.idl, 100%)
-Added interfaces ILaunchUIContext, ILaunchUIContextProvider
-Added numerous shell related GUIDs
-Added some missing property key related enums from propkey.h (should be 100% now)
-Some enums for shell automation have officially associated IIDs; added these with new EnumId attrib
-Added some missing registry constants and enum associations
-Added SDK helper macros ISLBUTTON, ISMBUTTON, ISRBUTTON, ISDBLCLICK
-EnumWindows, EnumChildWindows, and EnumTaskWindows APIs were inexplicably missing.
-(API Standards) GetAltTabInfo did not conform to WinDevLib API standards (LongPtr instead of String) 
-(API Standards) GetKeyboardLayoutName did not conform to WinDevLib API standards (LongPtr instead of String) 
-(API Standards) ShutdownBlockReasonQuery was inconsistent with ShutdownBlockReasonCreate for String vs LongPtr.
-(API Standards) CreateDesktop[A,ExA,Ex] did not use appropriate `DEVMODE[A,W]` variants.
-(API Standards) RegCreateKey[A,W,ExA,ExW] did not use SECURITY_ATTRIBUTES instead of ByVal LongPtr.
-(API Standards) RegConnectRegistry[A, ExA] did not use String types
-(Bug fix) OpenDesktopA incorrectly used `DeclareWide`
-(Bug fix) FOLDERTYPEID_ GUIDs were not properly defined as Static
-(Bug fix) RegCreateKey, RegConnectRegistryExA definitions incorrect 
-(Bug fix) RegCreateKeyTransacted definition incorrect (wrong alias)
-(Bug fix) Some winmm UDTs lacked required PackingAlignment attribute
-(Bug fix) WAVEFORMAT[EX,EXTENSIBLE] lacked required PackingAlignment attribute

**Update (v7.6.334, 08 Mar 2024):**
-Added 100% coverage of winsafer.h
-Expanded power API coverage; powerbase.h, powersetting.h, powrprof.h 100%.

**Update (v7.6.332, 06 Mar 2024):**
-NamespaceTreeControl default changed to INamespaceTreeControl2
-Added inline helper SDK macros FreeIDListArray[Full|Child], SetContractDelegateWindow
-(Bug fix) INameSpaceTreeControlEvents::OnGetTooltip should be ByVal pszTip
-(Bug fix) MSGBOXPARAMS[A,W], MSGBOXDATA defs incorrect for x64.

**Update (v7.6.330, 04 Mar 2024):**
-Added some additional sync APIs; synchapi.h coverage now 100%.
-IObjectCollection now uses proper types (IUnknown and IObjectArray)
-(Bug fix) IsBadStringPtr missing alias 
-(Bug fix) GetTimeZoneInformationForYear definition incorrect (used Long instead of Integer; no change needed, would work either way)
-(Bug fix) HIMC/HIMCC types for IME APIs were incorrectly Long instead of LongPtr; this was only true on early Windows versions

**Update (v7.6.325, 29 Feb 2024):**
-Suppress new tB warnings for msvbvm60 DLL redirects (this info is still noted in the descriptions for each API)
-(Bug fix) DF_ALLOWOTHERACCOUNTHOOK value incorrect

**Update (v7.6.324, 27 Feb 2024):**
-Added additional Variant/PROPVARIANT helpers; propvarutil.h now 100% covered
-Additional DirectX As Any->proper type
-Substantial improvement to Task Scheduler 2.0 interfaces (intellisense, Boolean instead of Integer where appropriate, descriptions)
-(Bug fix) InitVariantFromString was not a dll export (replaced by macro)
-(Bug fix) VariantToFileTimeArray and VariantToFileTimeArrayAlloc don't exist
-(Bug fix) IScheduledWorkItem missing 3 methods and GetRunTimes, SetCreator methods incorrect.
-(Bug fix) ITaskSettings missing Compatibility Let/Get methods.
-(Bug fix) ITaskSettings3 missing CreateMaintenanceSettings method
-(Name change) ISchedulingAgent was apparently renamed ITaskScheduler by Windows 2000; coclass SchedulingAgent to CTaskScheduler.
               Further, IEnumWorkItems was IEnumTasks before that; why olelib was inconsistent here, I don't know.
               Since the SDK still defines these as aliases, WinDevLib now includes both names for all 3.
-(Name change) TASK_RUNLEVEL corrected to more appropriate TASK_RUNLEVEL_TYPE


**Update (v7.6.322, 24 Feb 2024):**
-Added DSA and DPA APIs (dpa_dsa.h, 100% coverage including macros)
-Further compat updates for The trick's typelibs:
   -IDWriteFontFileLoader.CreateStreamFromKey last arg now retval.
   -ID2D1RenderTarget many arguments now optional, with correct default values where appropriate
   -IWICBitmap.Lock last arg now retval
-ID2D1Factory and ID2D1Geometry had many As Any arguments switched to their proper types
-Added SizeToLongLong helper function
-(Bug fix) PointFToLongLong helper function incorrect.   
-(Bug fix) ID2D1RenderTarget::CreateBitmap definition incompatible with 64bit

**Update (v7.6.320, 20 Feb 2024):**
-Added IPrintDocumentPackage* interfaces and coclasses (DocumentTarget.idl, 100%)
-Added un/under-documented MRU APIs from comctl32
-For compatibility with The trick's D2D and WIC typelibs:
   -D2D1_MATRIX_ types are now flat; the D2D alias versions remain the same, switch to these if you were using the previous defs.
   -ID2D1Effect data arguments are now As Any (no change needed)
   -Some arguments now optional (no change needed)
      NOTE: Unlike VB6, twinBASIC supports ByVal Nothing to pass a null pointer to a ByRef interface/object method.
   -ID2D1DeviceContext::CreateEffect last param now return value
   -IWICBitmapDecoder::GetFrame last param now return value
-Many Direct2D/DirectWrite types were changed from As Any to their real UDT, since tB supports vbNullPtr to pass the optional null.
   While this reduces compatibility with The trick's TLBs (and oleexp), the extra info and intellisense benefits are worth it.
-(Bug fix) PathRemoveBackslashW incorrectly used String.
-(Bug fix) LookupPrivilegeValue[A] used LongPtr instead of String.
-(Bug fix) PointToLongLong ambiguous overloads; new PointFToLongLong for POINTF.
-(Bug fix) All Direct2D effects CLSID functions incorrect (returning UUID_NULL)
-(Bug fix) IDWriteLocalizedStrings, IDWriteTextFormat, IDWriteTextLayout, IDWriteLocalFontFileLoader string arguments improperly ByRef
-(Bug fix) IDWriteInlineObject, IDWriteTextRenderer, and IDWritePixelSnapping argument clientDrawingContext should be ByVal LongPtr.
-(Bug fix) Several DirectWrite font UDTs had plocalename members incorrectly defined as Long, making them incompatible with 64bit


**Update (v7.6.312, 10 Feb 2024):**
-Added IAccessControl/IAuditControl interfaces
-Added numerous missing propsys APIs; propsys.h coverage now 100%
-Added a few missing registry functions, also previously excluded deprecated ones-- winreg.h coverage is now 100.
-GetProcessMemoryInfo now uses As Any so PROCESS_MEMORY_COUNTERS and PROCESS_MEMORY_COUNTERS_EX2 can also be used.
-Added System Restore APIs from SrRestorePtApi.h (100%). IMPORTANT: Event types have been prefixed with SRPT_ due to common name conflicts (e.g. it has BACKUP, RESTORE, etc, that are now SRPT_BACKUP, SRPT_RESTORE, etc)
-Added Compressor APIs from compressapi.h (100%). IMPORTANT: Compress and Decompress have been renamed CompressorCompress and CompressorDecompress, respectively, due to the short name conflict potential.
-(Internal) Moved crypto APIs to their own file, wdAPICrypto.twin. Internet APIs moved to new module wdAPIInternet with wdInternet.twin. DEVPKEY and MiscGUID regions moved to wdDefs.twin. wdAPI.twin was becoming unmanageable and running into performance issues; it was up to 65k lines before this reorganization.
-Implemented all basic Interlocked* APIs. These are implemented primarily as static libraries: Only a few of these are exported by the Windows API, and only on x86.
 To handle this, I've included my Interlocked64 project as a static library. I've also produced a 32bit version to handle all the inline/instrinsic ones besides the basics.
 If you wish to avoid static linking these obj files (while using the APIs), specify the compiler flag:
 #WINDEVLIB_AVOID_INTRINSICS
 This uses the kernel32 versions *where available*: You're limited to InterlockedIncrement, InterlockedDecrement, InterlockedExchange[Add], and InterlockedCompareExchange[64].
 Using any besides those 6 will trigger the static library to be included.
 NOTE: TEMPORARY: Due to editing instability, a default alternative of ONLY the kernel32s are set-- for use in Beta 423. See wdInterlocked.twin.
-Added addtional error codes
-Added cards.dll APIs for 32bit only (no 64bit build exists)
 
**Update (v7.5.310, 26 Jan 2024):**
-Massive expansion of crypt APIs; coverage of wincrypt.h, dpapi.h (crypto data protection) and mssip.h now 100%
-Coverage of wintrust.h is now 99%; all but a couple of difficult to decipher macros and a byte sequence the order needs to be verified for.
-Coverage of memoryapi.h is now 100% (excluding APIs only available to Store Apps)
-Added UserNotification2 coclass; oleexp had this with a default of IUserNotification2, and while WinDevLib had UserNotification as a coclass, it had IUserNotification as a default without listing 2. Added 2 and the additional coclass.
-EVENT_FILTER_EVENT_ID is now buffered to the maximum number of IDs. This allows using it directly, at the expense of not being able to use LenB for size.
-Virtual* memory functions now use ByVal addresses instead of ByRef As Any; 99% of code uses this definition.
-(Bug fix) CertFreeCertificateContext definition incompatible with x64
-(Bug fix) SwapVTableEntry helper not working with old defs


**Update (v7.4.308, 20 Jan 2024):**
-Added interface IAttachmentExecute and coclass AttachmentServices.
-Added interface IStorageProviderBanners, and coclass StorageProviderBanners.
-Substantial expanson of crypto APIs; bcrypt.h, ncrypt.h, and ncryptprotect.h all now have 100% coverage, and wincrypt.h coverage has doubled (though still has quite a bit to go)
-Crypto provider enum Crypt_Providers (dwProvType) renamed to CryptProviders to resolve conflict with SDK-defined CRYPT_PROVIDERS type.
-Numerous missing IShellMenu related consts/types; fixed incorrect intellisense associations.
-(Bug fix) MEMORYSTATUS definition incorrect (incompatible with 64bit). The associated API should not be used however, as it has problems with >4GB RAM. Use GlobalMemoryStatusEx.


**Update (v7.3.306, 17 Jan 2024):**
-Some additional crypto APIs.
-Added undocumented TaskDialogIndirect button flags (Abort, Ignore, Continue, Retry, Help) and renamed the enum to the proper SDK-defined name (replace TDBUTTONS with TASKDIALOG_COMMON_BUTTON_FLAGS)
-Added x,y option to PointToLongLong helper.
-Added some missing GDI defs and macros.
-(Bug fix) Numerous duplicated enum values undetected last time.

**Update (v7.3.304, 15 Jan 2024):**
-Added legacy Sync Manager interfaces/coclasses (mobsync.h, 100%)
-Added process snapshot APIs (ProcessSnapshot.h, 100% coverage)
-Added all consts (grouped as enums where possible) from propkey.h
-Added new property keys from propkey.h
-Added some missing STR_ binding strings.
-Small additions to get shellapi.h coverage to 100%
-Added undocumented interfaces IInfoBarMessage, IInfoBarHost, and IBrowserProgressSessionProvider (for the popup banner menus in NSEs)
-Added undocumented interfaces IShellFolder3, IFilterItem, IItemFilter
-Added undocumented interfaces IScope, IScopeItem (NSE filtering)
-(Bug fix) LockWorkStation incorrect case.
-(Bug fix) SHFILEOPSTRUCT[A,W] definition incorrect for x86


**Update (v7.2.301, 10 Jan 2024):** Bug fix: Numerous duplicated enum values.

**Update (v7.2.300, 09 Jan 2024):**:
-Added wincred advapi32.dll APIs; wincred.h, 100% coverage
-Completed adding WinHttp APIs, winhttp.h coverage now 100% (note: The WinHttp interface/coclass is )
-Added remaining websocket.dll APIs, websocket.h coverage now 100%
-Added pointer encode/decodes functions (and kernel32's Beep): utilapiset.h 100% coverage
-A few missing WinInet APIs
-Around 100 additional HRESULT error constants w/ descriptions.
-Base WinRT IInspectable and some initialization APIs and HSTRING APIs added.
-(Bug fix) All ERROR_DS_x constants were wrong. ICM ERROR_x constants were wrong.

**Update (v7.2.289, 06 Jan 2024):** Bug fix: InternetConnect definition incorrect.

**Update (v7.2.288, 06 Jan 2024):**
-Added Photo Acquisition interfaces and coclasses (photoacquire.h, 100%)
-Added accessibility APIs from oleacc.dll (oleacc.h now 100% coverage). Really thought these were already added; there's a bug in oleexp where most are missing from that too despite presence in source.
-Added inline Library helper functions from ShObjIdl_core.h; also some additional shell32.dll APIs.
-Added SDDL language string constants; coverage of sddl.h now 100%.
-Additional advapi32.dll security APIs, to bring coverage of securitybaseapi.h to 100%.
-Added 100% coverage of dssec.h.
-Cleaned up PROCESS_BASIC_INFORMATION
-(Bug fix) LogonUserEx[A,W] definitions incorrect.
-(Bug fix) CreateWellKnownSid definition incorrect.
-(Bug fix) GetSidIdentifierAuthority definition likely incorrect.
-(Bug fix) SHChangeUpdateImageIDList missing 1-byte packing attribute.
-(Bug fix) A couple setup APIs missing 32bit 1-byte packing attribute.


**Update (v7.1.286, 02 Jan 2024):**
-Added initial coverage of Lsa* APIs from advapi32.dll/NTSecAPI.h/LSALookup.h/ntlsa.h
-WIC: Converted LongPtr buffer arguments to As Any, for more flexibility in what can be supplied.
-WIC: Converted all ByVal VarPtr(WICRect) LongPtr's to ByRef WICRect.
-(Bug fix) IWICBitmapSourceTransform::CopyPixels definition incorrect.
-(WinDevLibImpl) Added Implements-compatible WIC interfaces for custom codec creation.

**Update (v7.0.283, 01 Jan 2024):**
-Improved enum associations/formatting for WIC.
-Added numerous missing GUIDs from wincodecsdk.h
-(Bug fix) IWICPalette, IWICFormatConverter, IWICBitmapDecoderInfo, IWICPixelFormatInfo2, IWICMetadataReaderInfo, IWICMetadataHandlerInfo, IWICBitmapCodecInfo, IWICComponentInfo, WICMapGuidToShortName, WICMapSchemaToName had numerous ByVal/ByRef mixups.


**Update (v7.0.282, 01 Jan 2024):**
-Added all variable conversion and arithmetic helpers from oleauto.h; coverage of that header now 100% (of supported by language). 
-Additional GUIDs and error consts from olectl.h to bring that header's coverage to 100%.
-VARCMP enum renamed VARCMPRES to avoid conflict with VarCmp API.
-Added missing flags for VariantChangeType[Ex]
-SHFileOperation and SHFILEOPSTRUCT did not conform to API standards. Struct names were incorrect; the operations aborted member was incorrectly defined as Boolean, but the padding bytes prevented it from failing the entire function.
-SysAllocStringByteLen now use ByVal As Any, since either a String or LongPtr would be ByVal.
-(Bug fix) SysAllocString definition incorrect (Long instead of LongPtr, impacting 64bit)
-(Bug fix) SysFreeString definition incorrect (ByRef instead of ByVal)
-(Bug fix) SysReAllocStringLen should use DeclareWide
-(Bug fix) LHashValOfName is a macro, not an export; now implemented properly.
-(Bug fix) FORMATETC used a Long for CLIPFORMAT, which is incorrect.
-(MAJOR BUG FIX) IStream was missing UnlockRegion. This impacted numerous derived interfaces, throwing off their vtables, completely breaking them.

**Update (v7.0.280, 28 Dec 2023):**
-INDEXTOOVERLAYMASK was inexplicably missing; also added inverse, OVERLAYMASKTOINDEX.
-Additional setup APIs-- newdev.h, 100% coverage, and additional cfgmgr32 APIs.
-Additional kernel32 APIs-- processthreadsapi.h now has 100% coverage
-(Bug fix) SetupDiGetClassDevsW did not conform to WinDevLib API standards.
-(Bug fix) Some SetupAPI defs did not have the required 1-byte packing on 32bit
-(Bug fix) NMLVKEYDOWN and NMTVKEYDOWN did not have required packing alignment

**Update (v7.0.277, 21 Dec 2023):**
-Added customer caller for AuthzReportSecurityEvent (experimental).
-(Bug fix) SHEmptyRecycleBinW, PathRemoveBackslash, PathSkipRoot, CreateMailslot did not conform to API standards
-(Bug fix) All SHReg* APIs missing W variants
-(Bug fix) PathAddExtension, PathAddRoot, EnumSystemLanguageGroups, LoadCursorFromFile, waveInGetErrorText definitions incorrect (misplaced alias)
-(Bug fix) PathIsDirectoryA/W, PdhAddEnglishCounterA definitions incorrect (invalid alias)
-(Bug fix) GetLogicalDriveStringsA definition incorrect (DeclareWide on ANSI)
-(Bug fix) Missing DeclareWide:
    Get/SetComputerName[Ex]
    All THelp32.h APIs
    SHUpdateImage
    ShellNotify_Icon
    WaveIn/OutDevCaps
    HttpQueryInfo


**Update (v7.0.276, 20 Dec 2023):**
-Added cryptui.dll APIs (cryptuiapi.h, 100% coverage)
-Some additional SetupAPI and Cfgmgr32 defs, as well as devmgr.dll APIs documented and not (show device manager, prop pages, problem wizard, etc)
-More inexplicably missing shell32 APIs
-Additional APIs from ShellScalingAPI.h (now 100% coverage)
-(Bug fix) Duplicated DEVPROP_TYPE_* values.
-(Bug fix) GetExplicitEntriesFromAcl definition incorrect (misplaced Alias)
-(Bug fix) Wow64RevertWow64FsRedirection lacked explicit ByVal modifier.
-(Bug fix) Get/SetProcessDpiAwareness definitions incorrect.

**Update (v7.0.272, 17 Dec 2023):**

***tbShellLib is now WinDevLib - Windows Development Library for twinBASIC***




***MAJOR CHANGES***
*LARGE_INTEGER*
I've been considering these for a long time, and decided to pull the trigger before tB goes 1.0. 

The LARGE_INTEGER type is defined  in C as:

```
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
instead substituted a QLARGE_INTEGER or ULARGE_INTEGER in it's own definitions.
This will now change. The original plan was to wait for union support which would allow both while still triggering
the 8 byte alignment rules, but that has recently been confirmed as a post-1.0 feature. When that is added, the old
option will be added back in.
LARGE_INTEGER now by default uses QuadPart, and all QLARGE_INTEGER have been changed to LARGE_INTEGER.

Reminder: This does greatly simplify things; you can remove all conversions to Currency and related multiply/divide 
          by 10,000. Also, note that if you use your own local definition, WinDevLib does not supercede it for your
          own code. It is strongly recommended to switch away from Currency when doing 64bit updates.

A compiler flag is available to restore the old definition (but not the use of QLARGE_INTEGER in WinDevLib defs):
WINDEVLIB_NOQUADLI

*SendMessage and PostMessage*
These will now conform to the same API standards as all other functions; the undenominated (without A or W suffix)
will now point to SendMessageW and PostMessageW and use DeclareWide. Note that these have never affected the target
itself, it's always just modified how String arguments are interpreted. 99% of usage of these will not be impacted
by this, since you'll still be able to use String and not nee to modify the result for ANSI/Unicode conversion.
PostMessage already used DeclareWide, which was perhaps causing unexpected issues in the edge cases.

**Addtional changes:**
-Added interface IActCtx and coclass ActCtx.
-Missing WH_ enum values and associated types for SetWindowsHookEx
-Numerous missing VK_* virtual key codes
-Missing WM_* wParam enums.
-Several service APIs did not conform to WinDevLib API standards with respect to A/W/DeclareWide UDT naming.
-Added a lot of additional user32 content.
-Added variable min/max constants from limits.h (100% coverage)
-Redid FILEDESCRIPTOR[A,W] to use proper FILETIME types and Integer for WCHAR instead of 2x Byte.
-Added several types associated with clipboard formats.
-Added unsigned variable helper functions (thanks to Krool for these): UnsignedAdd, CUIntToInt, CIntToUInt, CULngToLng, and CLngToULng. CULngToLng has an override between the original Double and LongLong, CLngToULng does too but rewrites the output into an argument since tB can't overload purely based on function return type.
-Added gesture angle macros GID_ROTATE_ANGLE_TO_ARGUMENT/GID_ROTATE_ANGLE_FROM_ARGUMENT
-Added hundreds of additional NTSTATUS values.
-Added overloads to LOWORD and HIWORD to handle LongLong directly.
-winuser.h now has 100% coverage of language-supported definitions (10.0.25309 SDK); the largest header to date with this distinction with over 16000 lines in the original.
-(Bug fix) LBItemFromPt was marked Private.
-(Bug fix) RealGetWindowClass definition incorrect (invalid alias).
-(Bug fix) Duplicated constant: CCHILDREN_SCROLLBAR
-(Bug fix) PostThreadMessage definition incorrect and did not meet API standards.
-(Bug fix) InsertMenuItem[A,W] definitions technically incorrect although not causing an error. Also did not conform to API standards.
-(Bug fix) PostThreadMessage definition incorrect.
-(Bug fix) PostMessageA incorrectly had DeclareWide.
-(Bug fix) ILCreateFromPathEx was removed as it's not exported from shell32 either by name or ordinal.
-(Bug fix) ILCloneChild, ILCloneFull, ILIsAligned, ILIsChild, ILIsEmpty, ILNext, and ILSkip are only macros; they were declared as shell32.dll functions. Some of these were aliases and modified as appropriate, the rest were implemented as functions.
-(Bug fix) ILLoadFromStream is exported by ordinal only.
-(Bug fix, WinDevLibImpl) IPersistFile method definition incorrect.


**Update (v6.6.269):**
-Added helper function GetNtErrorString that gets strings for NTSTATUS values. GetSystemErrorString already exists for HRESULT.
-SHLimitInputEdit didn't have the ByVal attribute included, making it easy to not realize it's then required when called.
-CreateSymbolicLink API inexplicable missing.
-LIMITINPUTSTRUCT has been renamed to the original, correct name LIMITINPUT. The original documentation and demos have made this change too with the recently released universal compatibility update.

**Update (v6.6.268, 11 Dec 2023):**
-Added UI Animation interfaces and coclasses
-Added Radio Manager interfaces and some undocumented coclasses to use them. Added undocumented interface IRadioManager with coclass RadioManagementAPI: This controls 'Airplane mode' on newer Windows.
-Added IThumbnailStreamCache and coclass ThumbnailStreamCache. Note: Due to simple name potential conflicts, flags prefixed with TSC_. A ByVal SIZE is replaced with ByVal LongLong; copy into one.
-Added additional event trace APIs; coverage of evntrace.h is now 100%.
-Additional BCrypt APIs sufficient for basic public key crypto implementations.
-Added additional language settings APIs from WinNls.h; coverage is near or at 100% now.
-Added remaining transaction manager APIs; coverage of ktmw32.h is now 100%.
-Added all remaining .ini/win.ini file APIs.
-Added misc other APIs.
-Added memcpy alias for RtlMoveMemory (in addition to CopyMemory and MoveMemory)
-Several event trace APIs and transaction API improperly used 'As GUID', which is undefined in tbShellLib and will refer to the unsupported stdole GUID.
-Reworked the way the REASON_CONTEXT union was set up; the old version would likely not work as implied.
-(Bug fix) KSIDENTIFIER union size incorrect.

**Update (v6.5.263, 06 Dec 2023):**
-Added numerous missing shell32 APIs.
-Some additional kernel32 APIs, bringing coverage of fileapi.h to 100%.
-Added numerous IOCTL_DISK_* constants and associated UDTs.
-Converted some ListView-related consts to enums to use with their associated UDTs.
-Added missing name mappings structs for SHFileOperation.
-(Bug fix) BITMAPFILEHEADER, DISK_EXTENT, VOLUME_DISK_EXTENT, and STORAGE_PROPERTY_QUERY typed improperly marked Private.
-(Bug fix) STORAGE_PROPERTY_QUERY definition incorrect
-(Bug fix) SCSI_PASS_THROUGH_BUFFERED24 definition incorrect.
-(Bug fix) GetVolumeInformationByHandle definition incorrect.
-(Bug fix) ReadFile did not conform to tbShellLib API conventions (ByVal As Any instead of OVERLAPPED)


**Update (v6.5.260, 04 Dec 2023):**
-Added all authz APIs/consts/types from authz.h; note that AuthzReportSecurityEvent is currently unsupported by the language. However, it internally calls AuthzReportSecurityEventFromParams.
-Added many missing shlwapi APIs; URL flags enum missing values
-Updated shlwapi "Is" functions to use BOOL instead of Long where that way in sdk.
-Completed all currently known PROCESSINFOCLASS structs for NtQueryInformationProcess.
-Added custom enums for PROCESS_MITIGATION_* structs
-(Bug fix) SHGetThreadRef/SHSetThreadRef definitions incorrect
-(Bug fix) SHMessageBoxCheck definition incorrect
-(Bug fix) Path[Un]QuoteSpaces definitions incorrect

**Update (v6.4.258), 28 Nov 2023):**
-Large number of additional advapi security APIs (AccCtrl.h and AclAPI.h, 100% coverage)
-Additional crypto APIs
-(Bug fix) Missing FindFirstFileEx flag FIND_FIRST_EX_ON_DISK_ENTRIES_ONLY.

**Update (v6.4.257), 26 Nov 2023):** GdipGetImageEncoders/GdipGetImageDecoders definitions "incorrect" for unclear reasons... Documentation indicates it's an array of ImageCodecInfo, which does not contain any C-style arrays, but there's a mismatch between the byte size and number of structs * sizeof. Changed to As Any to allow byte buffers in addition to oversized ImageCodecInfo buffers.
**Update (v6.4.256, 25 Nov 2023):**
-Added inexplicably missing basic versioning and sysinfo APIs from kernel32.
-Added ListView subitem control undocumented CLSIDs.
-Additional sys info classes (NtQuerySystemInformation).
-Misc. API additions.
-(Bug fix) GetAtomName[A,W] and GlobalGetAtomName[A,W] definitions incorrect.
-(Bug fix) Multiple ole32 functions incorrectly passing ANSI strings.
-(Bug fix) ListView_GetItemText was thoroughly broken.
-(Bug fix) GetSystemDirectory definition incorrect.
-(Bug fix) EnumPrintersA definition incorrect; GetPrinter, SetPrinter, and GetJob definitions technically incorrect but no impact unless you had redefined associated UDTs.
-(Bug fix) UNICODE_STRING members renamed to their proper SDK names. I realize this is a substantial breaking change but it's a minor adjustment and I feel it's important to be faithful to the SDK.

**Update (v6.3.253, 17 Nov 2023):**
-Additional crypto APIs (both classic and nextgen)
-Added GetSystemErrorString helper function to look up system error messages.
-(Bug fix) FormatMessage did not follow W/DeclareWideString convention; last param not ByVal.
-(Bug fix) RtlDestroyHeap has but one p.
-(Bug fix) CoCreateInstance overloads not playing nice. Only a single form available now.

**Update (v6.3.252, 11 Nov 2023):**
-Expanded bcrypt coverage
-Added RegisterDeviceChangeNotification and the numerous assorted consts/types (dbt.h, 100% coverage)
-Added DISP_E_* and TYPE_E_* error messages w/ descriptions. Added additional errors and descriptions for several original oleexp error sets.
-The WBIDM enum that was full of IDM_* values has had the values changed to WBIDM_*. IDM_ is the standard prefix for menu resources, so these would often conflict with projects not using the same resource id, and the ids here are for Win9x legacy content.
-All the fairly useless system info UDTs and an actually useful one, SYSTEM_PROCESS_ID_INFORMATION was missing.
-Additional shell32 APIs
-(Bug fix) Helper function NT_SUCCESS was improperly Private
-(Bug fix) SetupDiGetClassDevPropertySheets[W] definitions incorrect


**Update (v6.3.250, 5 Nov 2023):**
-Added Credential Provider interfaces from credentialprovider.h
-Added missing TlHelp32.h APIs/structs, now covered 100%. 
-Added several types/enums related to things already in project.
-(Bug fix) Duplicate of NETRESOURCE type. Project was subsequently analyzed for further duplicated types, and 4 other bugs in this class were eliminated.
-(Bug fix) No base PEB type defined.
-(NOTICE) OpenGL is being deferred until twinBASIC has Alias support (planned).

**Update (v6.3.240):**
-Added interfaces IComputerAccounts, IEnumAccounts, IComputerAccountNotify, and IProfileNotify with coclasses LocalUserAccounts, LocalGroups, LoggedOnAccounts, ProfileAccounts, UserAccounts, and ProfileNotificationHandler. Also added numerous PROPERTYKEYs associated with this functionality.
-Added a limited set of Winsock APIs. Note that with the exception of WSA* APIs, the short, generic names have been prefixed with ws_.
-Misc API additions including undocumented shell32 APIs, and additional ntdll APIs.
-Additional PE file structs
-(Bug fix) Several WebView2 interface had incompatible Property Get defs for ByVal UDT workarounds.


**Update (v6.2.238):**
-Added a limited set of winhttp APIs
-Added misc APIs for recent projects
-(Bug fix) RegQueryValueEx/RegQueryValueExW/RegQueryValueExA definitions incorrect.

**Update (v6.2.237):** Missing consts for upcoming project.

**Update (v6.2.234):**
-Added additional file info structs, exe header structs, and ntdll APIs
-(Bug fix) Some Disk Quota interface enums had incorrect names and in some cases values.

**Update (v6.2.232):** 
-Added gdi32 Color Management (ICM) APIs. 
-Additional sysinfo UDTs. 
-TypeHints for NT functions missing them.

**Update (v6.2.230):** 
-Added Windows Networking (WNet) APIs (winnetwk.h, 100% coverage (mpr.dll))
-Major expansion of internationalization API coverage from winnls.h.
-Added numerous missing common User32 functions.
-Misc bug fixes, inc. InsertMenuItem entry-point not found, missing menu alternates (W or A variations)
-Added overloads for a number of functions, if you have any trouble with the following, please file a bug report:
CoUnMarshalIface

IsValidLocaleName
EnumDateFormatsExEx
EnumCalendarInfoExEx

GetSystemDefaultLocaleName
GetCurrencyFormatEx
GetNumberFormatEx
GetCalendarInfoEx

SetUserGeoName

GetThreadPreferredUILanguages
SetThreadPreferredUILanguages
SetProcessPreferredUILanguages
GetProcessPreferredUILanguages

LocaleNameToLCID

GetDurationFormat
GetDurationFormatEx

GetLocaleInfoEx
ResolveLocaleName

GetNLSVersion
GetNLSVersionEx

ToUnicode

LoadBitmap[A,W]

ModifyMenu
InsertMenu

StgMakeUniqueName

SHEvaluateSystemCommandTemplate
SHIsFileAvailableOffline
SHSetLocalizedName
SHGetLocalizedName
SHRemoveLocalizedName

GetClassInfo[A, Ex, ExA]

RmAddFilter
RmRemoveFilter


**Update (v6.1.229):** Bug fix: A number of APIs had missing 'As <type>` statements, which were upgraded to errors. tB had previosly not caught these.

**Update (v6.1.228):**
-Completed imm32 APIs
-Added Job Object APIs
-Completed Virtual Disk APIs (virtdisk.h, 100% coverage)
-Many missing gdi32.dll APIs
-Misc APIs, inc. some power APIs
-All UDTs for NtQueryInformationFile (through current Win11)
-Bug fix: GDI object enum duplicate
-Bug fix: Some incorrect UDTs


**Update (v6.0.220):** 
-Added Network List Manager interfaces and coclass NetworkListManager.
-Added WININET APis (wininet.h, 99% coverage-- autoproxy defs unsupported by language)
-Added all APIs from iphlpapi.h (IP Helper; network stats); netioapi.h not included. Will be in future release.
-Added all Console APIs (wincon.h/wincontypes.h/consoleapi[, 2,3].h) and Comm APIs. WinEvent APIs and consts.
-FileDeviceTypes has been renamed DEVICE_TYPE, per usage in km
-Added most UDTs for GetFileInformationByHandle and native equivalents.
-Added Vista+ Thread Pool APIs, including inlined ones (threadpoolapiset.h, 100% coverage)
-Added Windows 10+ Secure Enclave APIs (enclaveapi.h, 100% coverage)
-dlgs.h, part of windows.h, has been added *AS AN OPTIONAL EXTENSION* due to anticipated naming conflicts with common names like 'lst1'. Add the compiler constant `TB_SHELLLIB_DLGH = 1` to include these. 
-Bug fix: Numerous UDTs with LARGE_INTEGER changed to QLARGE_INTEGER where the lack of 8-byte QuadPart was throwing alignment off. Note that in the future, tB will have union support, at which point LARGE_INTEGER will be changed to one, and all QLARGE_INTEGER replaced.

**Update (v5.3.214):** Added all DWM APIs from dwmapi.h. Added undoc'd shell app manager interfaces/coclasses. Added CPL applet defs. Misc API additions and bugfixes.

**Update (v5.2.210-212):** Additional APIs for upcoming project release.

**Update (v5.2.208):** Substantial API additions; inc. SystemParametersInfo structs/enums, display config, raw input, missing dialog stuff. Additional standard helper macros found in Windows headers.

**Update (5.1.207):** 
-Added PropSheet macros
-Set PROPSHEETPAGE to V4 by default
-Add missing PropSheet consts
-Bug fix: PROPSHEETHEADER definitions incorrect
-Bug fix: PostMessage API not 64bit compatible
-Bug fix: Several ListView macros not 64bit compatible

**Update (5.1.206):** 
-Updated WebView2 to match 1.0.1901.177. 
-Completed all advapi32 registry functions.
-Expanded Media Foundation APIs.
-Bug fix: Property Sheet callback enums were missing values and improperly organized.
-Misc bug fixes and additions to APIs.

**Update (v5.0.203):** Bug fix: D3DMATRIX layout with 2d array was incorrect.

**Update (v5.0.201):** 
-Added some missing DirectShow media stream interfaces. 
-Complete coverage of winmm API sets for wave, midi, time, sound, mmio, joystick, mci, aux, and mixer.
-Complete coverage of printer and print spooler APIs from winspool.
-Major expansion of security-related APIs
-Added D3D compiler APIs and effects interfaces; 
-Added basic DirectSound interfaces/apis.
-Bug fix: ShowWindow relocated to slShellCore.twin to avoid amibiguity with SHOWWINDOW enum.
-Bug fix: Misc. bug fixes to APIs.

**Update (v4.16.193):** Small API update for upcoming project; some resource loading APIs were missing.

**Update (v4.16.191):** Bug fix: Multiple instances of errors for auto-declaring Variants, Bug fix: `GetClipboardData` incorrectly returned a Long (should be LongPtr).

**Update (v4.16.190):** Critical bug fix: TB_SHELLLIB_LITE mode was broken. Added additional DirectX errors w/ desciprtions. Added initial D3D compiler apis, note that by default, these direct to d3dcompiler_47.dll, however you can specify compiler flag D3D_COMPILER = 44, 45, and 46 to use those.

**Update (v4.15.188):** Added SAFEARRAY APIs for manual operations on them and some more TypeLib-related APIs.

**Update (v4.14.185):** Bug fix: lstrcmp, lstrcmpi, and lstrcat declarations were incorrect. Some additional [ TypeHint ] attributes add.

**Update (v4.14.184):** Added SxS Assembly interfaces and APIs. Added MAKEINTRESOURCE macro. Added additional error messages. Made TaskDialogIndirect returns Optional per MSDN.

**Update (v4.14.182):** Added missing kernel32 string functions. Added SUCCEEDED helper function.

**Update (v4.14.181):** Bug fix: CHARFORMAT2[A|W] was incorrectly declared.

**Update (v4.14.180):** Much more extensive coverage of PROPVARIANT and Variant helpers for supported VB types (use changetype first to use them with unsigned et al).

**Update (v4.14.178):** Added partial Virtual Disk APIs and unsigned PROPVARIANT helpers.

**Update (v4.13.177):** Bug fix: Helper function UI_HSB had a syntax error.

**Update (v4.13.175):** Ribbon UI IIDs were missing.

**Update (v4.13.174):** Added caret APIs. Bug fix: Certain DirectWrite interfaces had members incompatible with x64. *IMPORTANT:* Having a single format for both 32 and 64bit breaks compatibility with the 32bit-only version. Previously `DWRITE_TEXT_RANGE` arguments were passed as two separate arguments, you'll now need to copy them to a single LongLong to pass.

**Update (v4.12.172):** User info APIs added.

**Update (v4.12.171):** No change; version number incremented to test package manager.

**Update (v4.12.170):** Bug fix: IOleInPlaceSite::Scroll scrollExtant should be ByVal. Added common error consts w/ descriptions.

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

Initial release: v1.0.1: 26 Sept 2022