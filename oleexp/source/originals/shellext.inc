typedef enum QueryContextMenuFlags {
    CMF_NORMAL              = 0x00000000,
    CMF_DEFAULTONLY         = 0x00000001,
    CMF_VERBSONLY           = 0x00000002,
    CMF_EXPLORE             = 0x00000004,
    CMF_NOVERBS             = 0x00000008,
    CMF_CANRENAME           = 0x00000010,
    CMF_NODEFAULT           = 0x00000020,
    CMF_INCLUDESTATIC       = 0x00000040,
    CMF_RESERVED            = 0xffff0000,
} QueryContextMenuFlags;

typedef enum GetCommandStringFlags {
    [helpstring("Canonical verb")]
    GCS_VERB         = 0x00000000,
    [helpstring("Help text for status bar")]
    GCS_HELPTEXT     = 0x00000001,
    [helpstring("Validate command exists")]
    GCS_VALIDATE     = 0x00000002,
    GCS_VERBW        = 0x00000004,
    GCS_HELPTEXTW    = 0x00000005,
    GCS_VALIDATEW    = 0x00000006,
    GCS_VERBICONW    = 0x00000014,
    [helpstring("Return in Unicode")]
    GCS_UNICODE      = 0x00000004,
} GetCommandStringFlags;

typedef enum InvokeCommandMask {
    CMIC_MASK_HOTKEY        = 0x00000020,
    CMIC_MASK_ICON          = 0x00000010,
    CMIC_MASK_FLAG_NO_UI    = 0x00000400,
    CMIC_MASK_UNICODE       = 0x00004000,
    CMIC_MASK_NO_CONSOLE    = 0x00008000,
//        CMIC_MASK_HASLINKNAME   = SEE_MASK_HASLINKNAME,
//        CMIC_MASK_FLAG_SEP_VDM  = SEE_MASK_FLAG_SEPVDM,
//        CMIC_MASK_HASTITLE      = SEE_MASK_HASTITLE,
    CMIC_MASK_ASYNCOK       = 0x00100000,
    CMIC_MASK_PTINVOKE      = 0x20000000,
    CMIC_MASK_SHIFT_DOWN    = 0x10000000,
    CMIC_MASK_CONTROL_DOWN  = 0x20000000,
} InvokeCommandMask;

typedef struct CMINVOKECOMMANDINFO {
    [helpstring("sizeof(CMINVOKECOMMANDINFO)")]
    LONG cbSize;
    [helpstring("Invoke mask")]
    InvokeCommandMask fMask;
    [helpstring("Might be NULL (indicating no owner window)")]
    LONG hwnd;
    [helpstring("Either a string or MAKEINTRESOURCE(idOffset)")]
    LONG lpVerb;
    [helpstring("Might be NULL (indicating no parameter)")]
    LONG lpParameters;
    [helpstring("Might be NULL (indicating no specific directory)")]
    LONG lpDirectory;
    [helpstring("One of SW_* values for ShowWindow() API")]
    LONG nShow;
    [helpstring("Hot key")]
    LONG dwHotKey;
    [helpstring("Icon handle")]
    LONG hIcon;
} CMINVOKECOMMANDINFO;

typedef struct CMINVOKECOMMANDINFOEX {
    LONG cbSize;
    InvokeCommandMask fMask;
    LONG hwnd;
    LONG lpVerb;
    LONG lpParameters;
    LONG lpDirectory;
    LONG nShow;
    LONG dwHotKey;
    LONG hIcon;
    [helpstring("For CreateProcess-StartupInfo.lpTitle")]
    LONG lpTitle;
    [helpstring("Unicode verb (for those who can use it)")]
    LONG lpVerbW;
    [helpstring("Unicode parameters (for those who can use it)")]
    LONG lpParametersW;
    [helpstring("Unicode directory (for those who can use it)")]
    LONG lpDirectoryW;
    [helpstring("Unicode title (for those who can use it)")]
    LONG lpTitleW;
    [helpstring("Point where it's invoked")]
    POINT ptInvoke;
} CMINVOKECOMMANDINFOEX;

[
    odl,
    uuid(000214E4-0000-0000-C000-000000000046),
    helpstring("IContextMenu interface")
]    
interface IContextMenu : stdole.IUnknown {

    [helpstring("Adds menu items to the specified menu.")]
    HRESULT QueryContextMenu(
        [in] long hMenu,
        [in] long indexMenu,
        [in] long idCmdFirst,
        [in] long idCmdLast,
        [in] QueryContextMenuFlags uFlags);

    [helpstring("Carries out the command associated with a context menu item.")]
    HRESULT InvokeCommand(
        [in, out] CMINVOKECOMMANDINFO* lpici);
    
    [helpstring("Retrieves the language-independent command string or the help text for a context menu item.")]
    HRESULT GetCommandString(
        [in] long idCmd,
        [in] GetCommandStringFlags uType,
        [in, out] int *pwReserved,
        [in] long pszName,
        [in] int cchMax);

};

[
    uuid(000214F4-0000-0000-C000-000000000046),
    helpstring("IContextMenu2 interface"),
    odl
]
interface IContextMenu2 : stdole.IUnknown {

    [helpstring("Adds menu items to the specified menu.")]
    HRESULT QueryContextMenu(
        [in] long hMenu,
        [in] long indexMenu,
        [in] long idCmdFirst,
        [in] long idCmdLast,
        [in] QueryContextMenuFlags uFlags);

    [helpstring("Carries out the command associated with a context menu item.")]
    HRESULT InvokeCommand(
        [in, out] CMINVOKECOMMANDINFO* lpici);
    
    [helpstring("Retrieves the language-independent command string or the help text for a context menu item.")]
    HRESULT GetCommandString(
        [in] long idCmd,
        [in] GetCommandStringFlags uType,
        [in, out] int *pwReserved,
        [in] long pszName,
        [in] int cchMax);

    [helpstring("Allows client objects of IContextMenu to handle messages associated with owner-drawn menu items.")]
    HRESULT HandleMenuMsg(
        [in] LONG uMsg,
        [in] LONG wParam,
        [in] LONG lParam);
        
};

[
    odl,
    helpstring("IShellExtInit interface"),
    uuid(000214E8-0000-0000-C000-000000000046)
]
interface IShellExtInit : stdole.IUnknown {

    [helpstring("Initializes a property sheet extension, context menu extension, or drag and drop handler.")]
    HRESULT Initialize (
        [in] LONG pidlFolder,
        [in] IDataObject *lpIDataObject,
        [in] LONG hkeyProgID);
};

typedef enum SHELLEXECUTEMASK {
    SEE_MASK_CLASSNAME        = 0x00000001,
    SEE_MASK_CLASSKEY         = 0x00000003,
    SEE_MASK_IDLIST           = 0x00000004,
    SEE_MASK_INVOKEIDLIST     = 0x0000000c,
    SEE_MASK_ICON             = 0x00000010,
    SEE_MASK_HOTKEY           = 0x00000020,
    SEE_MASK_NOCLOSEPROCESS   = 0x00000040,
    SEE_MASK_CONNECTNETDRV    = 0x00000080,
    SEE_MASK_FLAG_DDEWAIT     = 0x00000100,
    SEE_MASK_DOENVSUBST       = 0x00000200,
    SEE_MASK_FLAG_NO_UI       = 0x00000400,
    SEE_MASK_UNICODE          = 0x00004000,
    SEE_MASK_NO_CONSOLE       = 0x00008000,
    SEE_MASK_ASYNCOK          = 0x00100000,
    SEE_MASK_HMONITOR         = 0x00200000,
} SHELLEXECUTEMASK;

typedef struct SHELLEXECUTEINFO {
    LONG cbSize;
    SHELLEXECUTEMASK fMask;
    LONG hwnd;
    LONG lpVerb;
    LONG lpFile;
    LONG lpParameters;
    LONG lpDirectory;
    LONG nShow;
    LONG hInstApp;
} SHELLEXECUTEINFO;

typedef struct SHELLEXECUTEINFO2 {
    LONG cbSize;
    SHELLEXECUTEMASK fMask;
    LONG hwnd;
    LONG lpVerb;
    LONG lpFile;
    LONG lpParameters;
    LONG lpDirectory;
    LONG nShow;
    LONG hInstApp;

    LONG lpIDList;
    LONG lpClass;
    LONG hkeyClass;
    LONG dwHotKey;
    LONG hIcon_hMonitor;
    LONG hProcess;
} SHELLEXECUTEINFO2;

[
    uuid(000214F5-0000-0000-C000-000000000046),
    helpstring("IShellExecuteHook Ansi version"),
    odl
]
interface IShellExecuteHookA : stdole.IUnknown {

    HRESULT Execute(
        [in, out] SHELLEXECUTEINFO *pei);
};

[
    uuid(000214FB-0000-0000-C000-000000000046),
    helpstring("IShellExecuteHook Unicode version"),
    odl
]
interface IShellExecuteHookW : stdole.IUnknown {

    HRESULT Execute(
        [in, out] SHELLEXECUTEINFO *pei);
        
};

typedef enum GILFlags {
    [helpstring("allows containers to specify an \"open\" look")]
    GIL_OPENICON     = 0x0001, 
    [helpstring("icon is to be displayed in a ShellFolder")]
    GIL_FORSHELL     = 0x0002, 
    [helpstring("this is an async extract, return E_ASYNC")]
    GIL_ASYNC        = 0x0020, 
} GILFlags;

typedef enum GILRes {
    [helpstring("simulate this document icon for this")]
    GIL_SIMULATEDOC  = 0x0001, 
    [helpstring("icons from this class are per instance (each file has its own)")]
    GIL_PERINSTANCE  = 0x0002, 
    [helpstring("icons from this class per class (shared for all files of this type)")]
    GIL_PERCLASS     = 0x0004, 
    [helpstring("location is not a filename, must call ::ExtractIcon")]
    GIL_NOTFILENAME  = 0x0008,
    [helpstring("this icon should not be cached")]
    GIL_DONTCACHE    = 0x0010, 
} GILRes;

[
    odl,
    uuid(000214EB-0000-0000-C000-000000000046),
    helpstring("IExtractIcon ANSI version")
]
interface IExtractIconA : stdole.IUnknown {

    HRESULT GetIconLocation(
        [in] GILFlags uFlags,
        [in] LONG szIconFile,
        [in] LONG cchMax,
        [in, out] int * piIndex,
        [out, retval] GILRes *pwFlags);

    HRESULT Extract(
        [in] LONG pszFile,
        [in] LONG nIconIndex,
        [in, out] LONG *phiconLarge,
        [in, out] LONG *phiconSmall,
        [in] LONG nIconSize);
};

[
    odl,
    uuid(000214FA-0000-0000-C000-000000000046),
    helpstring("IExtractIcon Unicode version")
]
interface IExtractIconW : stdole.IUnknown {

    HRESULT GetIconLocation(
        [in] GILFlags uFlags,
        [in] LONG pwszIconFile,
        [in] LONG cchMax,
        [in, out] int * piIndex,
        [out, retval] GILRes * pwFlags);

    HRESULT Extract(
        [in] LONG pwszFile,
        [in] LONG nIconIndex,
        [in, out] LONG *phiconLarge,
        [in, out] LONG *phiconSmall,
        [in] LONG nIconSize);
};

[
    odl,
    uuid(000214E9-0000-0000-C000-000000000046),
    helpstring("IShellPropSheetExt interface")
]
interface IShellPropSheetExt: stdole.IUnknown {

    HRESULT AddPages(
        [in] LONG lpfnAddPage,
        [in] LONG lParam);

    HRESULT ReplacePage(
        [in] LONG uPageID,
        [in] LONG lpfnReplaceWith,
        [in] LONG lParam);

};

typedef enum PSP_FLAGS {
    PSP_DEFAULT                = 0x00000000,
    PSP_DLGINDIRECT            = 0x00000001,
    PSP_USEHICON               = 0x00000002,
    PSP_USEICONID              = 0x00000004,
    PSP_USETITLE               = 0x00000008,
    PSP_RTLREADING             = 0x00000010,
    PSP_HASHELP                = 0x00000020,
    PSP_USEREFPARENT           = 0x00000040,
    PSP_USECALLBACK            = 0x00000080,
    PSP_PREMATURE              = 0x00000400,
    PSP_HIDEHEADER             = 0x00000800,
    PSP_USEHEADERTITLE         = 0x00001000,
    PSP_USEHEADERSUBTITLE      = 0x00002000,
} PSP_FLAGS;

typedef enum PSPCALLBACKMESSAGES {
    PSPCB_RELEASE           = 1,
    PSPCB_CREATE            = 2,
} PSPCALLBACKMESSAGES;

typedef struct PROPSHEETPAGE {
    LONG dwSize;
    PSP_FLAGS dwFlags;
    LONG hInstance;
    LONG pResource;
    LONG hIcon;
    LONG pszTitle;
    LONG pfnDlgProc;
    LONG lParam;
    LONG pfnCallback;
    LONG pcRefParent;
} PROPSHEETPAGE;

[
    dllname("COMCTL32.DLL")
]
module comctl32 {

    const LONG MK_ALT = 0x0020;
    const LONG DD_DEFSCROLLINSET = 11;
    const LONG DD_DEFSCROLLDELAY = 50;
    const LONG DD_DEFSCROLLINTERVAL = 50;
    const LONG DD_DEFDRAGDELAY = 200;
    const LONG DD_DEFDRAGMINDIST = 2;

    [entry("CreatePropertySheetPageA")]
    LONG CreatePropertySheetPage(
        [in] PROPSHEETPAGE *PSP);

    [entry("CreatePropertySheetPageW")]
    LONG CreatePropertySheetPageW(
        [in] PROPSHEETPAGE *PSP);

    [entry("DestroyPropertySheetPage")]
    LONG DestroyPropertySheetPage(
        [in] LONG hPage);
}

typedef struct PSHNOTIFY {
    NMHDR hdr;
    LONG lParam;
} PSHNOTIFY;

typedef enum PSNOTIFYCODES {    
    PSN_FIRST               = (-200),
    PSN_LAST                = (-299),
    PSN_SETACTIVE           = (-200),
    PSN_KILLACTIVE          = (-201),
    PSN_APPLY               = (-202),
    PSN_RESET               = (-203),
    PSN_HELP                = (-205),
    PSN_WIZBACK             = (-206),
    PSN_WIZNEXT             = (-207),
    PSN_WIZFINISH           = (-208),
    PSN_QUERYCANCEL         = (-209),
    PSN_GETOBJECT           = (-210),
} PSNOTIFYCODES;

typedef enum PSNOTIFYRESULTS {
    PSNRET_NOERROR              = 0,
    PSNRET_INVALID              = 1,
    PSNRET_INVALID_NOCHANGEPAGE = 2
} PSNOTIFYRESULTS;

typedef enum PSMESSAGES {
    [helpstring("SNDMSG(hDlg, PSM_SETCURSEL, index, hpage)")]
    PSM_SETCURSEL            = (0x400 + 101),

    [helpstring("SNDMSG(hDlg, PSM_REMOVEPAGE, index, hpage)")]
    PSM_REMOVEPAGE           = (0x400 + 102),

    [helpstring("SNDMSG(hDlg, PSM_ADDPAGE, 0, hpage)")]
    PSM_ADDPAGE             = (0x400 + 103),

    [helpstring("SNDMSG(hDlg, PSM_CHANGED, hwnd, 0L)")]
    PSM_CHANGED             = (0x400 + 104),

    [helpstring("SNDMSG(hDlg, PSM_RESTARTWINDOWS, 0, 0L)")]
    PSM_RESTARTWINDOWS      = (0x400 + 105),

    [helpstring("SNDMSG(hDlg, PSM_REBOOTSYSTEM, 0, 0L)")]
    PSM_REBOOTSYSTEM        = (0x400 + 106),

    [helpstring("PostMessage(hDlg, PSM_CANCELTOCLOSE, 0, 0L)")]
    PSM_CANCELTOCLOSE       = (0x400 + 107),

    [helpstring("SNDMSG(hDlg, PSM_QUERYSIBLINGS, wParam, lParam)")]
    PSM_QUERYSIBLINGS       = (0x400 + 108),

    [helpstring("SNDMSG(hDlg, PSM_UNCHANGED, hwnd, 0L)")]
    PSM_UNCHANGED           = (0x400 + 109),

    [helpstring("SNDMSG(hDlg, PSM_APPLY, 0, 0L)")]
    PSM_APPLY               = (0x400 + 110),

    [helpstring("SNDMSG(hDlg, PSM_SETTITLE, wStyle, lpszText)")]
    PSM_SETTITLEA           = (0x400 + 111),
    PSM_SETTITLEW           = (0x400 + 120),

    [helpstring("PostMessage(hDlg, PSM_SETWIZBUTTONS, 0, dwFlags)")]
    PSM_SETWIZBUTTONS       = (0x400 + 112),

    [helpstring("PostMessage(hDlg, PSM_PRESSBUTTON, iButton, 0)")]
    PSM_PRESSBUTTON         = (0x400 + 113),

    [helpstring("SNDMSG(hDlg, PSM_SETCURSELID, 0, id)")]
    PSM_SETCURSELID         = (0x400 + 114),

    [helpstring("SNDMSG(hDlg, PSM_SETFINISHTEXT, 0, lpszText)")]
    PSM_SETFINISHTEXTA      = (0x400 + 115),
    PSM_SETFINISHTEXTW      = (0x400 + 121),

    [helpstring("HWND = SNDMSG(hDlg, PSM_GETTABCONTROL, 0, 0)")]
    PSM_GETTABCONTROL       = (0x400 + 116),

    [helpstring("BOOL = SNDMSG(hDlg, PSM_ISDIALOGMESSAGE, 0, pMsg)")]
    PSM_ISDIALOGMESSAGE     = (0x400 + 117),

    [helpstring("HWND = SNDMSG(hDlg, PSM_GETCURRENTPAGEHWND, 0, 0L)")]
    PSM_GETCURRENTPAGEHWND  = (0x400 + 118),

} PSMESSAGES;

    //ID_PSRESTARTWINDOWS     0x2
    //ID_PSREBOOTSYSTEM       (ID_PSRESTARTWINDOWS | 0x1)

typedef enum PSWIZBUTTONS {
    PSWIZB_BACK = 0x00000001,
    PSWIZB_NEXT = 0x00000002,
    PSWIZB_FINISH = 0x00000004,
    PSWIZB_DISABLEDFINISH = 0x00000008,
} PSWIZBUTTONS;

typedef enum PSBUTTONS {
    PSBTN_BACK              = 0,
    PSBTN_NEXT              = 1,
    PSBTN_FINISH            = 2,
    PSBTN_OK                = 3,
    PSBTN_APPLYNOW          = 4,
    PSBTN_CANCEL            = 5,
    PSBTN_HELP              = 6,
    PSBTN_MAX               = 6,
} PSBUTTONS;

//
// IQueryInfo
//
[
    odl,
    uuid(00021500-0000-0000-C000-000000000046),
]
interface IQueryInfo : stdole.IUnknown {

    typedef enum QITipFlags {
        QITIPF_DEFAULT          = 0x00000000,
        QITIPF_USENAME          = 0x00000001,
        QITIPF_LINKNOTARGET     = 0x00000002,
        QITIPF_LINKUSETARGET    = 0x00000004,
    } QITipFlags;

    typedef enum QIFlags {
        QIF_CACHED           = 0x00000001,
        QIF_DONTEXPANDFOLDER = 0x00000002,
    } QIFlags;

    HRESULT GetInfoTip(
        [in] QITipFlags dwFlags, 
        [in, out] LONG *ppwszTip);

    HRESULT GetInfoFlags(
        [in, out] QITipFlags *pdwFlags);

};
 
typedef enum IEIFlags {
    IEIFLAG_ASYNC      = 0x0001,      // ask the extractor if it supports ASYNC extract (free threaded)
    IEIFLAG_CACHE      = 0x0002,      // returned from the extractor if it does NOT cache the thumbnail
    IEIFLAG_ASPECT     = 0x0004,      // passed to the extractor to beg it to render to the aspect ratio of the supplied rect
    IEIFLAG_OFFLINE    = 0x0008,      // if the extractor shouldn't hit the net to get any content neede for the rendering
    IEIFLAG_GLEAM      = 0x0010,      // does the image have a gleam ? this will be returned if it does
    IEIFLAG_SCREEN     = 0x0020,      // render as if for the screen  (this is exlusive with IEIFLAG_ASPECT )
    IEIFLAG_ORIGSIZE   = 0x0040,      // render to the approx size passed, but crop if neccessary
} IEIFlags;

[
    odl,
    uuid(BB2E617C-0920-11d1-9A0B-00C04FC2D6C1)
]
interface IExtractImage : stdole.IUnknown {

    HRESULT GetLocation(
        [in] LONG pszPathBuffer,
        [in] LONG cch,
        [in, out] LONG *pdwPriority,
        [in, out] SIZE *prgSize,
        [in] LONG dwRecClrDepth,
        [in, out] IEIFlags *pdwFlags);

    HRESULT Extract(
        [out, retval] LONG *phBmpThumbnail);
};

[
    odl,
    uuid(953BB1EE-93B4-11d1-98A3-00C04FB687DA)
]
interface IExtractImage2 : IExtractImage {

    HRESULT GetLocation(
        [in] LONG pszPathBuffer,
        [in] LONG cch,
        [in, out] LONG *pdwPriority,
        [in, out] SIZE *prgSize,
        [in] LONG dwRecClrDepth,
        [in, out] IEIFlags *pdwFlags);

    HRESULT Extract(
        [out, retval] LONG *phBmpThumbnail);

    HRESULT GetDateStamp(
        [out, retval] CURRENCY *pDateStamp);
};

typedef enum FILEOP {
    FO_MOVE           = 0x0001,
    FO_COPY           = 0x0002,
    FO_DELETE         = 0x0003,
    FO_RENAME         = 0x0004,
} FILEOP;

typedef enum FILEOP_FLAGS {
    FOF_MULTIDESTFILES         =0x0001,
    FOF_CONFIRMMOUSE           =0x0002,
    FOF_SILENT                 =0x0004,  // don't create progress/report
    FOF_RENAMEONCOLLISION      =0x0008,
    FOF_NOCONFIRMATION         =0x0010,  // Don't prompt the user.
    FOF_WANTMAPPINGHANDLE      =0x0020,  // Fill in SHFILEOPSTRUCT.hNameMappings
                                  // Must be freed using SHFreeNameMappings
    FOF_ALLOWUNDO              =0x0040,
    FOF_FILESONLY              =0x0080,  // on *.*, do only files
    FOF_SIMPLEPROGRESS         =0x0100,  // means don't show names of files
    FOF_NOCONFIRMMKDIR         =0x0200,  // don't confirm making any needed dirs
    FOF_NOERRORUI              =0x0400,  // don't put up error UI
    FOF_NOCOPYSECURITYATTRIBS  =0x0800,  // dont copy NT file Security Attributes
    FOF_NORECURSION            =0x1000,  // don't recurse into directories.
    FOF_NO_CONNECTED_ELEMENTS  =0x2000,  // don't operate on connected file elements.
    FOF_WANTNUKEWARNING        =0x4000,  // during delete operation, warn if nuking instead of recycling (partially overrides FOF_NOCONFIRMATION)
} FILEOP_FLAGS;

typedef enum PRINTEROP_FLAGS {
    PO_DELETE       =0x0013,  // printer is being deleted
    PO_RENAME       =0x0014,  // printer is being renamed
    PO_PORTCHANGE   =0x0020,  // port this printer connected to is being changed
                            // if this id is set, the strings received by
                            // the copyhook are a doubly-null terminated
                            // list of strings.  The first is the printer
                            // name and the second is the printer port.
    PO_REN_PORT     =0x0034,  // PO_RENAME and PO_PORTCHANGE at same time.
} PRINTEROP_FLAGS;

[
    odl,
    uuid(000214EF-0000-0000-C000-000000000046)
]
interface ICopyHookA : stdole.IUnknown {

    HRESULT CopyCallback(
        [in] LONG hwnd,
        [in] FILEOP wFunc,
        [in] FILEOP_FLAGS wFlags,
        [in] LONG pszSrcFile,
        [in] LONG dwSrcAttribs,
        [in] LONG pszDestFile, 
        [in] LONG dwDestAttribs);
};

[
    odl,
    uuid(000214FC-0000-0000-C000-000000000046)
]
interface ICopyHookW : stdole.IUnknown {

    HRESULT CopyCallback(
        [in] LONG hwnd,
        [in] FILEOP wFunc,
        [in] FILEOP_FLAGS wFlags,
        [in] LONG pszSrcFile,
        [in] LONG dwSrcAttribs,
        [in] LONG pszDestFile, 
        [in] LONG dwDestAttribs);
};

typedef struct SHCOLUMNID {
    UUID fmtid;
    LONG pid;
} SHCOLUMNID;

typedef struct SHCOLUMNINIT {
    LONG dwFlags ;        // initialization flags
    LONG dwReserved ;     // reserved for future use.
    BYTE wszFolder[520];  // fully qualified folder path (or empty if multiple folders)
} SHCOLUMNINIT;

typedef enum SHCOLSTATE {
    SHCOLSTATE_TYPE_STR     = 0x00000001,
    SHCOLSTATE_TYPE_INT     = 0x00000002,
    SHCOLSTATE_TYPE_DATE    = 0x00000003,
    SHCOLSTATE_TYPEMASK     = 0x0000000F,
    SHCOLSTATE_ONBYDEFAULT  = 0x00000010,   // should on by default in details view
    SHCOLSTATE_SLOW         = 0x00000020,   // will be slow to compute, do on a background thread
    SHCOLSTATE_EXTENDED     = 0x00000040,   // provided by a handler, not the folder
    SHCOLSTATE_SECONDARYUI  = 0x00000080,   // not displayed in context menu, but listed in the "More..." dialog
    SHCOLSTATE_HIDDEN       = 0x00000100,   // not displayed in the UI
} SHCOLSTATE;

typedef struct SHCOLUMNINFO {
    SHCOLUMNID scid;           // OUT the unique identifier of this column
    short vt;                  // OUT the native type of the data returned
    LONG fmt;                  // OUT this listview format (LVCFMT_LEFT, usually)
    LONG cChars;               // OUT the default width of the column, in characters
    SHCOLSTATE csFlags;        // OUT SHCOLSTATE flags
    BYTE wszTitle[160];        // OUT the title of the column
    BYTE wszDescription[256];  // OUT full description of this column
} SHCOLUMNINFO;

// #define SHCDF_UPDATEITEM        0x00000001      // this flag is a hint that the file has changed since the last call to GetItemData

typedef struct SHCOLUMNDATA {
    LONG dwFlags ;            // combination of SHCDF_ flags.
    LONG dwFileAttributes ;   // file attributes.
    LONG dwReserved ;         // reserved for future use.
    LONG pwszExt ;            // address of file name extension
    BYTE wszFile[520] ;       // Absolute path of file.
} SHCOLUMNDATA;

[
    odl,
    uuid(E8025004-1C42-11d2-BE2C-00A0C9A83DA1)
]
interface IColumnProvider : stdole.IUnknown {

    HRESULT Initialize(
        [in, out] SHCOLUMNINIT *psci);

    HRESULT GetColumnInfo(
        [in] LONG dwIndex, 
        [in, out] SHCOLUMNINFO *psci);

    HRESULT GetItemData(
        [in, out] SHCOLUMNID *pscid, 
        [in, out] SHCOLUMNDATA *pscd, 
        [in, out] VARIANT *pvarData);

};

/*

 URLSearchHooks are installed here_

 HKEY_LOCAL_MACHINE
   Software
     Microsoft
       Internet Explorer
         UrlSearchHooks
*/

[
    odl,
    uuid(ac60f6a0-0fd9-11d0-99cb-00c04fd64497)
]
interface IURLSearchHook : stdole.IUnknown {

    HRESULT Translate(
        [in] LONG lpwszSearchURL,
        [in] LONG cchBufferSize);
        
};

[
    odl,
    uuid(09F656A2-41AF-480C-88F7-16CC0D164615)
]
interface ISearchContext : stdole.IUnknown {

    HRESULT GetSearchUrl(
        [in, out] BSTR *pbstrSearchUrl);
        
    HRESULT GetSearchText(
        [in, out] BSTR * pbstrSearchText);
    
    HRESULT GetSearchStyle(
        [in, out] LONG * pdwSearchStyle);

};

[
    odl,
    uuid(5ee44da4-6d32-46e3-86bc-07540dedd0e0)
]
interface IURLSearchHook2 : stdole.IUnknown {

    // *** IURLSearchHook methods ***
    HRESULT Translate(
        [in] LONG lpwszSearchURL,
        [in] LONG cchBufferSize);

    // *** IURLSearchHook2 methods ***
    HRESULT TranslateWithSearchContext(
        [in] LONG lpwszSearchURL,
        [in] LONG cchBufferSize,
        [in] ISearchContext *pSearchContext);

};

[
    odl,
    uuid(000214e1-0000-0000-c000-000000000046)
]
interface INewShortcutHookA : stdole.IUnknown {

    HRESULT SetReferent(
        [in] LONG pcszReferent,
        [in] LONG hwnd);

    HRESULT GetReferent(
        [in] LONG pszReferent,
        [in] int cchReferent);
    
    HRESULT SetFolder(
        [in] LONG pcszFolder);
    
    HRESULT GetFolder(
        [in] LONG pszFolder,
        [in] int cchFolder);
    
    HRESULT GetName(
        [in] LONG pszName,
        [in] int cchName);
    
    HRESULT GetExtension(
        [in] LONG pszExtension,
        [in] int cchExtension);

};

[
    odl,
    uuid(000214f7-0000-0000-c000-000000000046)
]
interface INewShortcutHookW : stdole.IUnknown {

    HRESULT SetReferent(
        [in] LONG pcszReferent,
        [in] LONG hwnd);

    HRESULT GetReferent(
        [in] LONG pszReferent,
        [in] int cchReferent);
    
    HRESULT SetFolder(
        [in] LONG pcszFolder);
    
    HRESULT GetFolder(
        [in] LONG pszFolder,
        [in] int cchFolder);
    
    HRESULT GetName(
        [in] LONG pszName,
        [in] int cchName);
    
    HRESULT GetExtension(
        [in] LONG pszExtension,
        [in] int cchExtension);

};
