 typedef enum QueryContextMenuFlags {
    CMF_NORMAL              = 0x00000000,
    CMF_DEFAULTONLY         = 0x00000001,
    CMF_VERBSONLY           = 0x00000002,
    CMF_EXPLORE             = 0x00000004,
    CMF_NOVERBS             = 0x00000008,
    CMF_CANRENAME           = 0x00000010,
    CMF_NODEFAULT           = 0x00000020,
    CMF_INCLUDESTATIC       = 0x00000040,
    //Vista+
    CMF_ITEMMENU            = 0x00000080,
    CMF_EXTENDEDVERBS       = 0x00000100,
    CMF_DISABLEDVERBS       = 0x00000200,
    //Win7+
    CMF_ASYNCVERBSTATE      = 0x00000400,
    CMF_OPTIMIZEFORINVOKE   = 0x00000800,
    CMF_SYNCCASCADEMENU     = 0x00001000,
    CMF_DONOTPICKDEFAULT    = 0x00002000,
	//Now how'd these get here?
    CMF_DVFILE              = 0x00010000,      // "File" pulldown
    CMF_ICM3                = 0x00020000,      // QueryContextMenu can assume IContextMenu3 semantics (i.e.,
                                                // will receive WM_INITMENUPOPUP, WM_MEASUREITEM, WM_DRAWITEM,
                                                // and WM_MENUCHAR, via HandleMenuMsg2)
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
    CMIC_MASK_HASLINKNAME   = 0x00010000, //SEE_MASK_HASLINKNAME,
	CMIC_MASK_FLAG_SEP_VDM  = 0x00040000, //SEE_MASK_FLAG_SEPVDM,
	CMIC_MASK_HASTITLE      = 0x00020000, //SEE_MASK_HASTITLE,
    CMIC_MASK_ASYNCOK       = 0x00100000,
    CMIC_MASK_SHIFT_DOWN    = 0x10000000,
    CMIC_MASK_CONTROL_DOWN  = 0x40000000,
    CMIC_MASK_FLAG_LOG_USAGE= 0x04000000,
    CMIC_MASK_NOZONECHECKS  = 0x00800000,  // SEE_MASK_NOZONECHECKS
    CMIC_MASK_PTINVOKE      = 0x20000000,
    CMIC_MASK_DATAOBJECT    = 0x40000000,   // lpPara is IDataObject* 
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
    SHOWWINDOW nShow;
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
    SHOWWINDOW nShow;
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
	HRESULT InvokeCommand([in] LONG lpici);
        //[in, out] CMINVOKECOMMANDINFO* lpici);
    
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
	HRESULT InvokeCommand([in] LONG lpici);
        //[in, out] CMINVOKECOMMANDINFO* lpici);
    
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
    SEE_MASK_NOZONECHECKS     = 0x00800000,
	SEE_MASK_NOQUERYCLASSSTORE  = 0x01000000,
	SEE_MASK_WAITFORINPUTIDLE   = 0x02000000,
	SEE_MASK_FLAG_LOG_USAGE     = 0x04000000,
	SEE_MASK_FLAG_HINST_IS_SITE = 0x08000000,
} SHELLEXECUTEMASK;

typedef enum SE_ERR {
	SE_ERR_FNF             = 2,       // file not found
	SE_ERR_PNF             = 3,       // path not found
	SE_ERR_ACCESSDENIED    = 5,       // access denied
	SE_ERR_OOM             = 8,       // out of memory
	SE_ERR_DLLNOTFOUND             = 32,
/* error values for ShellExecute() beyond the regular WinExec() codes */
	SE_ERR_SHARE                   = 26,
	SE_ERR_ASSOCINCOMPLETE         = 27,
	SE_ERR_DDETIMEOUT              = 28,
	SE_ERR_DDEFAIL                 = 29,
	SE_ERR_DDEBUSY                 = 30,
	SE_ERR_NOASSOC                 = 31
} SE_ERR;

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
    LONG lpIDList;
    LONG lpClass;
    LONG hkeyClass;
    LONG dwHotKey;
    LONG hIcon_hMonitor;
    LONG hProcess;
} SHELLEXECUTEINFO;
typedef SHELLEXECUTEINFO SHELLEXECUTEINFOW;


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
    [helpstring("get the default icon location if the final one takes too long to get")]
    GIL_DEFAULTICON  = 0x0040,
    [helpstring("the icon is for a shortcut to the object")]
    GIL_FORSHORTCUT  = 0x0080,
    [helpstring("return GIL_SHIELD or GIL_FORCENOSHIELD, don't block if GIL_ASYNC is set")]
    GIL_CHECKSHIELD  = 0x0200,
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
    [helpstring("icon should be 'stamped' with the LUA shield")]
    GIL_SHIELD       = 0x0200,
    [helpstring("icon must *not* be 'stamped' with the LUA shield")]
    GIL_FORCENOSHIELD = 0x0400,
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
    PSP_USEFUSIONCONTEXT       = 0x00004000,
} PSP_FLAGS;

typedef enum PSPCALLBACKMESSAGES {
	PSPCB_ADDREF            = 0,
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
typedef struct PROPSHEETPAGE_V2
{
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
    LONG pszHeaderTitle;
    LONG pszHeaderSubtitle;
} PROPSHEETPAGE_V2;
typedef struct PROPSHEETPAGE_V3
{
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
    LONG pszHeaderTitle;
    LONG pszHeaderSubtitle;
    LONG hActCtx;
} PROPSHEETPAGE_V3;
typedef struct PROPSHEETPAGE_V4
{
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
    LONG pszHeaderTitle;
    LONG pszHeaderSubtitle;
    LONG hActCtx;
    LONG bmHeader;
} PROPSHEETPAGE_V4;

typedef enum PropSheetHeaderFlags
{
	PSH_DEFAULT = 0x00000000,
	PSH_PROPTITLE = 0x00000001,
	PSH_USEHICON = 0x00000002,
	PSH_USEICONID = 0x00000004,
	PSH_PROPSHEETPAGE = 0x00000008,
	PSH_WIZARDHASFINISH = 0x00000010,
	PSH_WIZARD = 0x00000020,
	PSH_USEPSTARTPAGE = 0x00000040,
	PSH_NOAPPLYNOW = 0x00000080,
	PSH_USECALLBACK = 0x00000100,
	PSH_HASHELP = 0x00000200,
	PSH_MODELESS = 0x00000400,
	PSH_RTLREADING = 0x00000800,
	PSH_WIZARDCONTEXTHELP = 0x00001000,
//----- New flags for wizard97 -----------
	PSH_WIZARD97 = 0x01000000,
	PSH_WATERMARK = 0x00008000,
	PSH_USEHBMWATERMARK = 0x00010000, //  ' user pass in a hbmWatermark instead of pszbmWatermark
	PSH_USEHPLWATERMARK = 0x00020000, //  '
	PSH_STRETCHWATERMARK = 0x00040000, //  ' stretchwatermark also applies for the header
	PSH_HEADER = 0x00080000,
	PSH_USEHBMHEADER = 0x00100000,
	PSH_USEPAGELANG = 0x00200000,  //' use frame dialog template matched to page
	PSH_WIZARD_LITE = 0x00400000,
	PSH_NOCONTEXTHELP = 0x02000000,
	PSH_AEROWIZARD = 0x00004000,
	PSH_RESIZABLE = 0x04000000,
	PSH_HEADERBITMAP = 0x08000000,
	PSH_NOMARGIN = 0x10000000
} PropSheetHeaderFlags;

typedef struct PROPSHEETHEADER
{
    DWORD dwSize;
    PropSheetHeaderFlags dwFlags;
    LONG hwndParent;
    LONG hInstance;
    LONG hIcon;
    LONG pszCaption;
    LONG nPages;
    LONG nStartPage;
    LONG ppsp;
} PROPSHEETHEADER;
typedef struct PROPSHEETHEADER_V2
{
    DWORD dwSize;
    PropSheetHeaderFlags dwFlags;
    LONG hwndParent;
    LONG hInstance;
    LONG hIcon;
    LONG pszCaption;
    LONG nPages;
    LONG nStartPage;
    LONG ppsp;
    LONG hbmWatermark;
    LONG hplWatermark;
    LONG hbmHeader;
} PROPSHEETHEADER_V2;

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
        [in] void *PSP);

    [entry("CreatePropertySheetPageW")]
    LONG CreatePropertySheetPageW(
        [in] void *PSP);

    [entry("DestroyPropertySheetPage")]
    LONG DestroyPropertySheetPage(
        [in] LONG hPage);

    [entry("PropertySheetA")]
    LONG PropertySheet([in] void *psh);

    [entry("PropertySheetW")]
    LONG PropertySheetW([in] void *psh);
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
    PSN_TRANSLATEACCELERATOR= (-212),
    PSN_QUERYINITIALFOCUS   = (-213),
} PSNOTIFYCODES;

typedef enum PSNOTIFYRESULTS {
    PSNRET_NOERROR              = 0,
    PSNRET_INVALID              = 1,
    PSNRET_INVALID_NOCHANGEPAGE = 2,
    PSNRET_MESSAGEHANDLED       = 3,
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

	PSM_INSERTPAGE             = (0x400 + 119),
	PSM_SETHEADERTITLEA        = (0x400 + 125),
	PSM_SETHEADERTITLEW        = (0x400 + 126),
	PSM_SETHEADERSUBTITLEA     = (0x400 + 127),
	PSM_SETHEADERSUBTITLEW     = (0x400 + 128),
	PSM_HWNDTOINDEX            = (0x400 + 129),
	PSM_INDEXTOHWND            = (0x400 + 130),
	PSM_PAGETOINDEX            = (0x400 + 131),
	PSM_INDEXTOPAGE            = (0x400 + 132),
	PSM_IDTOINDEX              = (0x400 + 133),
	PSM_INDEXTOID              = (0x400 + 134),
	PSM_GETRESULT              = (0x400 + 135),
	PSM_RECALCPAGESIZES        = (0x400 + 136),
	PSM_SETNEXTTEXT            = (0x400 + 137),
	PSM_SHOWWIZBUTTONS         = (0x400 + 138),
	PSM_ENABLEWIZBUTTONS       = (0x400 + 139),
	PSM_SETBUTTONTEXT          = (0x400 + 140),
} PSMESSAGES;

    //ID_PSRESTARTWINDOWS     0x2
    //ID_PSREBOOTSYSTEM       (ID_PSRESTARTWINDOWS | 0x1)

typedef enum PSWIZBUTTONS {
    PSWIZB_BACK = 0x00000001,
    PSWIZB_NEXT = 0x00000002,
    PSWIZB_FINISH = 0x00000004,
    PSWIZB_DISABLEDFINISH = 0x00000008,
    PSWIZB_CANCEL         = 0x00000010,
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
        [helpstring("Flag says it's OK to take a long time generating tip")]
        QITIPF_USESLOWTIP      = 0x00000008,
        QITIPF_SINGLELINE      = 0x00000010,
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
     [helpstring("(deprecated) ask the extractor if it supports ASYNC extract (free threaded)")]
     IEIFLAG_ASYNC      = 0x0001,
     [helpstring("returned from the extractor if it does NOT cache the thumbnail")]
     IEIFLAG_CACHE      = 0x0002,
     [helpstring("passed to the extractor to beg it to render to the aspect ratio of the supplied rect")]
     IEIFLAG_ASPECT     = 0x0004,
     [helpstring("if the extractor shouldn't hit the net to get any content neede for the rendering")]
     IEIFLAG_OFFLINE    = 0x0008,
     [helpstring("does the image have a gleam ? this will be returned if it does")]
     IEIFLAG_GLEAM      = 0x0010,
     [helpstring("render as if for the screen  (this is exlusive with IEIFLAG_ASPECT )")]
     IEIFLAG_SCREEN     = 0x0020,
     [helpstring("render to the approx size passed, but crop if neccessary")]
     IEIFLAG_ORIGSIZE   = 0x0040,
     [helpstring("returned from the extractor if it does NOT want an icon stamp on the thumbnail")]
     IEIFLAG_NOSTAMP    = 0x0080,
     [helpstring("returned from the extractor if it does NOT want an a border around the thumbnail")]
     IEIFLAG_NOBORDER   = 0x0100,
     [helpstring("passed to the Extract method to indicate that a slower, higher quality image is desired, re-compute the thumbnail")]
     IEIFLAG_QUALITY    = 0x0200,
     [helpstring("returned from the extractor if it would like to have Refresh Thumbnail available")]
     IEIFLAG_REFRESH    = 0x0400,
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

    //HRESULT GetLocation(
    //    [in] LONG pszPathBuffer,
    //    [in] LONG cch,
    //    [in, out] LONG *pdwPriority,
    //    [in, out] SIZE *prgSize,
    //    [in] LONG dwRecClrDepth,
    //    [in, out] IEIFlags *pdwFlags);

    //HRESULT Extract(
    //    [out, retval] LONG *phBmpThumbnail);

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
    FOF_MULTIDESTFILES        = 0x0001,
    FOF_CONFIRMMOUSE          = 0x0002,
    [helpstring("don't display progress UI (confirm prompts may be displayed still)")]
    FOF_SILENT                = 0x0004,
    [helpstring("automatically rename the source files to avoid the collisions")]
    FOF_RENAMEONCOLLISION     = 0x0008,
    [helpstring("don't display confirmation UI, assume 'yes' for cases that can be bypassed, 'no' for those that can not")]
    FOF_NOCONFIRMATION        = 0x0010,
    [helpstring("Fill in SHFILEOPSTRUCT.hNameMappings. Must be freed using SHFreeNameMappings")]
    FOF_WANTMAPPINGHANDLE     = 0x0020,
    [helpstring("enable undo including Recycle behavior for IFileOperation::Delete()")]
    FOF_ALLOWUNDO              =0x0040,
    [helpstring("on *.*, do only files")]
    FOF_FILESONLY              =0x0080,
    [helpstring("means don't show names of files")]
    FOF_SIMPLEPROGRESS         =0x0100,
    [helpstring("don't confirm making any needed dirs")]
    FOF_NOCONFIRMMKDIR         =0x0200,
    [helpstring("don't put up error UI")]
    FOF_NOERRORUI              =0x0400,
    [helpstring("dont copy NT file Security Attributes")]
    FOF_NOCOPYSECURITYATTRIBS  =0x0800,
    [helpstring("don't recurse into directories.")]
    FOF_NORECURSION            =0x1000,
    [helpstring("don't operate on connected file elements.")]
    FOF_NO_CONNECTED_ELEMENTS  =0x2000,
    [helpstring("during delete operation, warn if nuking instead of recycling (partially overrides FOF_NOCONFIRMATION)")]
    FOF_WANTNUKEWARNING        =0x4000,
    [helpstring("deprecated; the operations engine always does the right thing on FolderLink objects (symlinks, reparse points, folder shortcuts)")]
    FOF_NORECURSEREPARSE       =0x8000,
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

typedef enum SHCOLSTATEF {
    SHCOLSTATE_TYPE_STR     = 0x00000001,
    SHCOLSTATE_TYPE_INT     = 0x00000002,
    SHCOLSTATE_TYPE_DATE    = 0x00000003,
    SHCOLSTATE_TYPEMASK     = 0x0000000F,
    [helpstring("should on by default in details view")]
    SHCOLSTATE_ONBYDEFAULT  = 0x00000010,
    [helpstring("will be slow to compute, do on a background thread")]
    SHCOLSTATE_SLOW         = 0x00000020,
    [helpstring("provided by a handler, not the folder")]
    SHCOLSTATE_EXTENDED     = 0x00000040,
    [helpstring("not displayed in context menu, but listed in the 'More...' dialog")]
    SHCOLSTATE_SECONDARYUI  = 0x00000080,
    [helpstring("not displayed in the UI")]
    SHCOLSTATE_HIDDEN       = 0x00000100,
    [helpstring("VarCmp() (with folderness test) produces same result as CompareIDs()")]
    SHCOLSTATE_PREFER_VARCMP = 0x00000200,
    [helpstring("comparison of PSFormatForDisplay result produces same result as CompareIDs()")]
    SHCOLSTATE_PREFER_FMTCMP = 0x00000400,
    [helpstring("do not sort folders separately")]
    SHCOLSTATE_NOSORTBYFOLDERNESS = 0x00000800,
    [helpstring("only displayed in the UI")]
    SHCOLSTATE_VIEWONLY      = 0x00010000,
    [helpstring("marks columns with values that should be read in a batch")]
    SHCOLSTATE_BATCHREAD     = 0x00020000,
    [helpstring("grouping is disabled for this column")]
    SHCOLSTATE_NO_GROUPBY    = 0x00040000,
    [helpstring("Can't resize the column")]
    SHCOLSTATE_FIXED_WIDTH   = 0x00001000,
    [helpstring("Width is the same in all DPIs")]
    SHCOLSTATE_NODPISCALE    = 0x00002000,
    [helpstring("Fixed width augments with height")]
    SHCOLSTATE_FIXED_RATIO   = 0x00004000,
    [helpstring("To filter out new display flags")]
    SHCOLSTATE_DISPLAYMASK   = 0x0000F000,
} SHCOLSTATEF;

typedef struct SHCOLUMNINFO {
    SHCOLUMNID scid;           // OUT the unique identifier of this column
    short vt;                  // OUT the native type of the data returned
    LONG fmt;                  // OUT this listview format (LVCFMT_LEFT, usually)
    LONG cChars;               // OUT the default width of the column, in characters
    SHCOLSTATEF csFlags;        // OUT SHCOLSTATE flags
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
