typedef enum IDs {
    FCIDM_SHVIEWFIRST = 0x0000,
    FCIDM_SHVIEWLAST = 0x7fff,
    FCIDM_BROWSERFIRST = 0xa000,
    FCIDM_BROWSERLAST = 0xbf00,
    FCIDM_GLOBALFIRST = 0x8000,
    FCIDM_GLOBALLAST = 0x9fff,
//
// Global submenu IDs and separator IDs
//
    FCIDM_MENU_FILE = (0x8000+0x0000),
    FCIDM_MENU_EDIT = (0x8000+0x0040),
    FCIDM_MENU_VIEW = (0x8000+0x0080),
    FCIDM_MENU_VIEW_SEP_OPTIONS = (0x8000+0x0081),
    FCIDM_MENU_TOOLS = (0x8000+0x00c0),
    FCIDM_MENU_TOOLS_SEP_GOTO = (0x8000+0x00c1),
    FCIDM_MENU_HELP = (0x8000+0x0100),
    FCIDM_MENU_FIND = (0x8000+0x0140),
    FCIDM_MENU_EXPLORE = (0x8000+0x0150),
    FCIDM_MENU_FAVORITES = (0x8000+0x0170),

//--------------------------------------------------------------------------
// control IDs known to the view
//--------------------------------------------------------------------------

    FCIDM_TOOLBAR = (0xa000 + 0),
    FCIDM_STATUS = (0xa000 + 1),
} IDs;

typedef enum FOLDERFLAGS {
    FWF_AUTOARRANGE =       0x0001,
    FWF_ABBREVIATEDNAMES =  0x0002,
    FWF_SNAPTOGRID =        0x0004,
    FWF_OWNERDATA =         0x0008,
    FWF_BESTFITWINDOW =     0x0010,
    FWF_DESKTOP =           0x0020,
    FWF_SINGLESEL =         0x0040,
    FWF_NOSUBFOLDERS =      0x0080,
    FWF_TRANSPARENT  =      0x0100,
    FWF_NOCLIENTEDGE =      0x0200,
    FWF_NOSCROLL     =      0x0400,
    FWF_ALIGNLEFT    =      0x0800,
    FWF_NOICONS      =      0x1000,
    FWF_SHOWSELALWAYS =     0x2000,
    FWF_NOVISIBLE    =      0x4000,
    FWF_SINGLECLICKACTIVATE=0x8000  // TEMPORARY -- NO UI FOR THIS
} FOLDERFLAGS;

typedef enum FOLDERVIEWMODE {
    FVM_ICON =              1,
    FVM_SMALLICON =         2,
    FVM_LIST =              3,
    FVM_DETAILS =           4,
} FOLDERVIEWMODE;

typedef struct TBBUTTON {
    LONG iBitmap;
    LONG idCommand;
    BYTE fsState;
    BYTE fsStyle;
    BYTE bReserved[2];
    LONG dwData;
    LONG iString;
} TBBUTTON;

typedef enum ESTRRET {
    STRRET_WSTR = 0,
    STRRET_OFFSET = 1,
    STRRET_CSTR = 2
} ESTRRET;

typedef struct STRRET {
    ESTRRET uType;
    BYTE cStr[260];
} STRRET;

typedef enum CSIDLs {
	CSIDL_DESKTOP                 = 0x0000, // <desktop>
	CSIDL_INTERNET                = 0x0001, 
	CSIDL_PROGRAMS                = 0x0002,
	CSIDL_CONTROLS                = 0x0003,
	CSIDL_PRINTERS                = 0x0004,
	CSIDL_PERSONAL                = 0x0005, // My Documents
	CSIDL_FAVORITES               = 0x0006,
	CSIDL_STARTUP                 = 0x0007,
	CSIDL_RECENT                  = 0x0008,
	CSIDL_SENDTO                  = 0x0009,
	CSIDL_BITBUCKET               = 0x000a,
	CSIDL_STARTMENU               = 0x000b,
	CSIDL_MYDOCUMENTS             = 0x000c,        // logical "My Documents" desktop icon
	CSIDL_MYMUSIC                 = 0x000d,        // "My Music" folder
	CSIDL_MYVIDEO                 = 0x000e,        // "My Videos" folder
	CSIDL_DESKTOPDIRECTORY        = 0x0010,
	CSIDL_DRIVES                  = 0x0011,
	CSIDL_NETWORK                 = 0x0012,
	CSIDL_NETHOOD                 = 0x0013,
	CSIDL_FONTS                   = 0x0014,
	CSIDL_TEMPLATES               = 0x0015,
	CSIDL_COMMON_STARTMENU        = 0x0016,
	CSIDL_COMMON_PROGRAMS         = 0x0017,
	CSIDL_COMMON_STARTUP          = 0x0018,
	CSIDL_COMMON_DESKTOPDIRECTORY = 0x0019,
	CSIDL_APPDATA                 = 0x001a,
	CSIDL_PRINTHOOD               = 0x001b,
	CSIDL_LOCAL_APPDATA           = 0x001c,        // <user name>\Local Settings\Applicaiton Data (non roaming)
	CSIDL_ALTSTARTUP              = 0x001d,        // DBCS
	CSIDL_COMMON_ALTSTARTUP       = 0x001e,        // DBCS
	CSIDL_COMMON_FAVORITES        = 0x001f,
	CSIDL_INTERNET_CACHE          = 0x0020,
	CSIDL_COOKIES                 = 0x0021,
	CSIDL_HISTORY                 = 0x0022,
	CSIDL_COMMON_APPDATA          = 0x0023,        // All Users\Application Data
	CSIDL_WINDOWS                 = 0x0024,        // GetWindowsDirectory()
	CSIDL_SYSTEM                  = 0x0025,        // GetSystemDirectory()
	CSIDL_PROGRAM_FILES           = 0x0026,        // C:\Program Files
	CSIDL_MYPICTURES              = 0x0027,        // C:\Program Files\My Pictures
	CSIDL_PROFILE                 = 0x0028,        // USERPROFILE
	CSIDL_SYSTEMX86               = 0x0029,        // x86 system directory on RISC
	CSIDL_PROGRAM_FILESX86        = 0x002a,        // x86 C:\Program Files on RISC
	CSIDL_PROGRAM_FILES_COMMON    = 0x002b,        // C:\Program Files\Common
	CSIDL_PROGRAM_FILES_COMMONX86 = 0x002c,        // x86 Program Files\Common on RISC
	CSIDL_COMMON_TEMPLATES        = 0x002d,        // All Users\Templates
	CSIDL_COMMON_DOCUMENTS        = 0x002e,        // All Users\Documents
	CSIDL_COMMON_ADMINTOOLS       = 0x002f,        // All Users\Start Menu\Programs\Administrative Tools
	CSIDL_ADMINTOOLS              = 0x0030,        // <user name>\Start Menu\Programs\Administrative Tools
	CSIDL_CONNECTIONS             = 0x0031,        // Network and Dial-up Connections
	CSIDL_COMMON_MUSIC            = 0x0035,        // All Users\My Music
	CSIDL_COMMON_PICTURES         = 0x0036,        // All Users\My Pictures
	CSIDL_COMMON_VIDEO            = 0x0037,        // All Users\My Video
	CSIDL_RESOURCES               = 0x0038,        // %windir%\Resources\, For theme and other windows resources.
	CSIDL_RESOURCES_LOCALIZED     = 0x0039,        // %windir%\Resources\<LangID>, for theme and other windows specific resources.
	CSIDL_COMMON_OEM_LINKS        = 0x003a,        // Links to All Users OEM specific apps
	CSIDL_CDBURN_AREA             = 0x003b,        // USERPROFILE\Local Settings\Application Data\Microsoft\CD Burning
	CSIDL_COMPUTERSNEARME         = 0x003d,        // Computers Near Me (computered from Workgroup membership)

	CSIDL_FLAG_PER_USER_INIT      = 0x0800,        // combine with CSIDL_ value to indicate per-user init (eg. upgrade)
	CSIDL_FLAG_NO_ALIAS           = 0x1000,        // combine with CSIDL_ value to insure non-alias versions of the pidl
	CSIDL_FLAG_CREATE             = 0x8000,        // combine with CSIDL_ value to force folder creation in SHGetFolderPath()
	CSIDL_FLAG_DONT_VERIFY        = 0x4000,        // combine with CSIDL_ value to return an unverified folder path
	CSIDL_FLAG_MASK               = 0xFF00,        // mask for all possible flag values
} CSIDLs;


typedef enum SHCIDS {
	SHCIDS_ALLFIELDS       = 0x80000000,
	SHCIDS_CANONICALONLY   = 0x10000000,
	SHCIDS_BITMASK         = 0xFFFF0000,
	SHCIDS_COLUMNMASK      = 0x0000FFFF,
} SHCIDS;

[
	odl,
	uuid(000214EA-0000-0000-C000-000000000046),
	helpstring("IPersistFolder interface")
]
interface IPersistFolder : stdole.IUnknown /* IPersist */ {

    HRESULT GetClassID(
        [in, out] UUID *lpClassID);

    HRESULT Initialize(
        [in] LONG pidl);
}

[
    odl,
    uuid(1AC3D9F0-175C-11d1-95BE-00609797EA4F)
]
interface IPersistFolder2 : stdole.IUnknown {
	
	HRESULT GetClassID(
		[in, out] UUID *lpClassID);

	HRESULT Initialize(
		[in] LONG pidl);

    HRESULT GetCurFolder(
		[out] LONG *ppidl);
};
    
[
    odl,
    uuid(1079acfc-29bd-11d3-8e0d-00c04f6837d5),
]
interface IPersistIDList : stdole.IUnknown {

	HRESULT GetClassID(
		[in, out] UUID *lpClassID);

    // sets or gets a fully qualifed idlist for an object
    HRESULT SetIDList(
		[in] LONG pidl);
		
    HRESULT GetIDList(
		[out] LONG *ppidl);

};

typedef struct EXTRASEARCH {
    UUID guidSearch;
    BYTE wszFriendlyName[160];
    BYTE wszUrl[4096];
} EXTRASEARCH;

[
	odl,
    helpstring("IEnumExtraSearch"),
    uuid(0E700BE1-9DB6-11d1-A1CE-00C04FD75D13),
]
interface IEnumExtraSearch : stdole.IUnknown {
    
    LONG Next(
        [in] LONG celt,
        [in, out] EXTRASEARCH *rgelt,
        [out, defaultvalue(0)] LONG *pceltFetched);

    HRESULT Skip(
		[in] LONG celt);
    
    HRESULT Reset();
    
    HRESULT Clone(
		[out, retval] IEnumExtraSearch **ppenum);
};


//
// Values for wFlags parameter of ISB::BrowseObject() member.
//
typedef enum BO_Flags {
    SBSP_DEFBROWSER = 0x0000,
    SBSP_SAMEBROWSER = 0x0001,
    SBSP_NEWBROWSER = 0x0002,

    SBSP_DEFMODE = 0x0000,
    SBSP_OPENMODE = 0x0010,
    SBSP_EXPLOREMODE = 0x0020,

    SBSP_ABSOLUTE = 0x0000,
    SBSP_RELATIVE = 0x1000,
    SBSP_PARENT = 0x2000,
    SBSP_NAVIGATEBACK = 0x4000,
    SBSP_NAVIGATEFORWARD = 0x8000,

    SBSP_ALLOW_AUTONAVIGATE = 0x10000,

    SBSP_INITIATEDBYHLINKFRAME = 0x80000000,
    SBSP_REDIRECT = 0x40000000,

    SBSP_WRITENOHISTORY = 0x08000000,
    SBSP_NOAUTOSELECT = 0x04000000
} BO_Flags;

typedef enum GCW_IDs {
    FCW_STATUS      = 0x0001,
    FCW_TOOLBAR     = 0x0002,
    FCW_TREE        = 0x0003,
    FCW_VIEW        = 0x0004,
    FCW_BROWSER     = 0x0005,
    FCW_INTERNETBAR = 0x0006,
    FCW_MENUBAR     = 0x0007,
    FCW_PROGRESS    = 0x0008,
    FCW_ADDRESSBAR  = 0x0009,
    FCW_TOOLBAND    = 0x000a,
    FCW_LINKSBAR    = 0x000b,
} GCW_IDs;

//
// Values for uFlags paremeter of ISB::SetToolbarItems member.
//

typedef enum STI_Flags {
    FCT_MERGE = 0x0001,
    FCT_CONFIGABLE = 0x0002,
    FCT_ADDTOEND = 0x0004,
} STI_Flags;

[
    odl,
    uuid(000214E2-0000-0000-C000-000000000046),
    helpstring("IShellBrowser ")
]
interface IShellBrowser : stdole.IUnknown /* IOleWindow */{

    [helpstring("IOleWindow.GetWindow")]
    HRESULT GetWindow(
        [out, retval] LONG * lpLONG);

    [helpstring("IOleWindow.ContextSensitiveHelp")]
    HRESULT ContextSensitiveHelp(
        [in] LONG fEnterMode);

    HRESULT InsertMenusSB(
        [in] LONG hmenuShared,
        [in, out] OLEMENUGROUPWIDTHS *lpMenuWidths);

    HRESULT SetMenuSB(
        [in] LONG hmenuShared,
        [in] LONG holemenuReserved,
        [in] LONG LONGActiveObject);

    HRESULT RemoveMenusSB(
        [in] LONG hmenuShared);

    HRESULT SetStatusTextSB(
        [in] LPWSTR lpszStatusText);

    HRESULT EnableModelessSB(
        [in] BOOL fEnable);

    HRESULT TranslateAcceleratorSB(
        [in, out] MSG *lpmsg,
        [in] short wID);

    HRESULT BrowseObject(
        [in] LONG pidl,
        [in] BO_Flags wFlags);

    HRESULT GetViewStateStream(
        [in] LONG grfMode,
        [out, retval] IStream **ppStrm);

    HRESULT GetControlWindow(
        [in] GCW_IDs id,
        [out, retval] LONG *lpLONG);

    HRESULT SendControlMsg(
        [in] GCW_IDs id,
        [in] LONG uMsg,
        [in] LONG wParam,
        [in] void *lParam,
        [out, retval] LONG * pret);

    HRESULT QueryActiveShellView(
        [out, retval] IShellView ** ppshv);

    HRESULT OnViewWindowActive(
        [in] IShellView * ppshv);

    HRESULT SetToolbarItems(
        [in] void *lpButtons,
        [in] LONG nButtons,
        [in] STI_Flags uFlags);
};

//
// shellview select item flags
//
typedef enum SVSI_Flags {
    SVSI_DESELECT       = 0x0000,
    SVSI_SELECT         = 0x0001,
    SVSI_EDIT           = 0x0003, // includes select
    SVSI_DESELECTOTHERS = 0x0004,
    SVSI_ENSUREVISIBLE  = 0x0008,
    SVSI_FOCUSED        = 0x0010,
    SVSI_TRANSLATEPT    = 0x0020,
} SVSI_Flags;

//
// shellview get item object flags
//
typedef enum SVGIO_Flags {
    SVGIO_BACKGROUND = 0x00000000,
    SVGIO_SELECTION  = 0x00000001,
    SVGIO_ALLVIEW    = 0x00000002,
} SVGIO_Flags;

//
// uState values for IShellView::UIActivate
//
typedef enum SVUIA_STATUS {
    SVUIA_DEACTIVATE       = 0,
    SVUIA_ACTIVATE_NOFOCUS = 1,
    SVUIA_ACTIVATE_FOCUS   = 2,
    SVUIA_INPLACEACTIVATE  = 3         // new flag for IShellView2
} SVUIA_STATUS;

typedef struct FOLDERSETTINGS {
    FOLDERVIEWMODE ViewMode;
    FOLDERFLAGS fFlags;
} FOLDERSETTINGS;

[
    odl,
    uuid(000214E3-0000-0000-C000-000000000046),
    helpstring("IShellView interface")
]
interface IShellView : stdole.IUnknown {

    // *** IOleWindow methods ***
    HRESULT GetWindow(
        [out, retval] LONG * lpLONG);

    HRESULT ContextSensitiveHelp(
        [in] LONG fEnterMode);

    // *** IShellView methods ***
    HRESULT TranslateAccelerator(
        [in, out] MSG *lpmsg);

    HRESULT EnableModelessSV(
        [in] LONG fEnable);

    HRESULT UIActivate(
        [in] SVUIA_STATUS uState);

    HRESULT Refresh();

    HRESULT CreateViewWindow(
        [in] IShellView *lpPrevView,
        [in] FOLDERSETTINGS *lpfs,
        [in] IShellBrowser *psb,
        [in] RECT *prcView,
        [out, retval] LONG *pLONG);

    HRESULT DestroyViewWindow();

    HRESULT GetCurrentInfo(
        [in, out] FOLDERSETTINGS *lpfs);

    HRESULT AddPropertySheetPages(
        [in] LONG dwReserved,
        [in] LONG lpfn,
        [in] LONG lParam);

    HRESULT SaveViewState();

    HRESULT SelectItem(
        [in] LONG pidlItem,
        [in] SVSI_Flags uFlags);

    HRESULT GetItemObject(
        [in] LONG uItem,
        [in, out] UUID *riid,
        [out, retval] LONG *ppv);

};

typedef enum SV2GV_Flags {
    SV2GV_CURRENTVIEW = -1,
    SV2GV_DEFAULTVIEW = -2
} SV2GV_Flags;

typedef struct SV2CVW2_PARAMS {
    LONG cbSize;
    IShellView *psvPrev;
    LONG pFOLDERSETTING;
    IShellBrowser *psbOwner;
    LONG prcView;
    LONG pvUUID;
    LONG hwndView;
} SV2CVW2_PARAMS;

[
    uuid(88E39E80-3578-11CF-AE69-08002B2E1262),
    odl
]
interface IShellView2 : stdole.IUnknown {

    // *** IOleWindow methods ***
    HRESULT GetWindow(
        [out, retval] LONG * lpLONG);

    HRESULT ContextSensitiveHelp(
        [in] LONG fEnterMode);

    // *** IShellView methods ***
    HRESULT TranslateAccelerator(
        [in, out] MSG *lpmsg);

    HRESULT EnableModelessSV(
        [in] LONG fEnable);

    HRESULT UIActivate(
        [in] SVUIA_STATUS uState);

    HRESULT Refresh();

    HRESULT CreateViewWindow(
        [in] IShellView *lpPrevView,
        [in] FOLDERSETTINGS *lpfs,
        [in] IShellBrowser *psb,
        [in] RECT *prcView,
        [out, retval] LONG *pLONG);

    HRESULT DestroyViewWindow();

    HRESULT GetCurrentInfo(
        [in, out] FOLDERSETTINGS *lpfs);

    HRESULT AddPropertySheetPages(
        [in] LONG dwReserved,
        [in] LONG lpfn,
        [in] LONG lParam);

    HRESULT SaveViewState();

    HRESULT SelectItem(
        [in] LONG pidlItem,
        [in] SVSI_Flags uFlags);

    HRESULT GetItemObject(
        [in] LONG uItem,
        [in, out] UUID *riid,
        [out, retval] LONG *ppv);

    // *** IShellView2 methods ***
    HRESULT GetView(
        [in, out] UUID* pvid, 
        [in] SV2GV_Flags uView);

    HRESULT CreateViewWindow2(
        [in, out] SV2CVW2_PARAMS *lpParams);

    HRESULT HandleRename(
        [in] LONG pidlNew);

    HRESULT SelectAndPositionItem(
        [in] LONG pidlItem,
        [in] SVSI_Flags uFlags,
        [in, out] POINT *point);
};

[
    odl,
    uuid(000214F2-0000-0000-C000-000000000046),
    helpstring("IEnumIDList interface")
]
interface IEnumIDList : stdole.IUnknown {

    HRESULT Next(
        [in] LONG celt,
        [in, out] LONG *rgelt,
        [out, retval] LONG *pceltFetched);

    HRESULT Skip(
        [in] LONG celt);

    HRESULT Reset();

    HRESULT Clone(
        [out, retval] IEnumIDList **ppenum);
}

typedef [helpstring("IShellFolder.GetDisplayNameOf/SetNameOf uFlags")]
enum SHGNO_Flags {
    [helpstring("default (display purpose)")]
    SHGDN_NORMAL = 0,
    [helpstring("displayed under a folder (relative)")]
    SHGDN_INFOLDER = 1,
    [helpstring("for displaying in the address (drives dropdown) bar")]
    SHGDN_FORADDRESSBAR = 0x4000,
    [helpstring("ParseDisplayName or path")]
    SHGDN_FORPARSING = 0x8000,
} SHGNO_Flags;

typedef [helpstring("IShellFolder.EnumObjects Flags")]
enum SHCONTF {
	[helpstring("only want folders enumerated (SFGAO_FOLDER)")]
    SHCONTF_FOLDERS             = 0x0020,
	[helpstring("include non folders")]
    SHCONTF_NONFOLDERS          = 0x0040,
	[helpstring("show items normally hidden")]
    SHCONTF_INCLUDEHIDDEN       = 0x0080,
	[helpstring("allow EnumObject() to return before validating enum")]
    SHCONTF_INIT_ON_FIRST_NEXT  = 0x0100,
	[helpstring("hint that client is looking for printers")]
    SHCONTF_NETPRINTERSRCH      = 0x0200,
	[helpstring("hint that client is looking sharable resources (remote shares)")]
    SHCONTF_SHAREABLE           = 0x0400,
	[helpstring("include all items with accessible storage and their ancestors")]
    SHCONTF_STORAGE             = 0x0800,
} SHCONTF;

typedef [helpstring("IShellFolder.GetAttributesOf flags")]
enum SFGAO_Flags {
    [helpstring("Objects can be copied")]
    SFGAO_CANCOPY                    = 0x00000001,
    [helpstring("Objects can be moved")]
    SFGAO_CANMOVE                    = 0x00000002,
    [helpstring("Objects can be linked")]
    SFGAO_CANLINK                    = 0x00000004,
    [helpstring("Objects can be renamed")]
    SFGAO_CANRENAME                  = 0x00000010,
    [helpstring("Objects can be deleted")]
    SFGAO_CANDELETE                  = 0x00000020,
    [helpstring("Objects have property sheets")]
    SFGAO_HASPROPSHEET               = 0x00000040,
    [helpstring("Objects are drop target")]
    SFGAO_DROPTARGET                 = 0x00000100,
    SFGAO_CAPABILITYMASK             = 0x00000177,

    [helpstring("object is encrypted (use alt color)")]
SFGAO_ENCRYPTED                  = 0x00002000,
    [helpstring("'slow' object")]
SFGAO_ISSLOW                     = 0x00004000,
    [helpstring("Ghosted icon")]
    SFGAO_GHOSTED                    = 0x00008000,
    [helpstring("Objects are shortcuts")]
    SFGAO_LINK                       = 0x00010000,
    [helpstring("Objects are shared")]
    SFGAO_SHARE                      = 0x00020000,
    [helpstring("Objects are read-only")]
    SFGAO_READONLY                   = 0x00040000,
    [helpstring("Objects are hidden")]
    SFGAO_HIDDEN                     = 0x00080000,
    SFGAO_DISPLAYATTRMASK            = 0x000FC000,

    [helpstring("It contains file system folder")]
    SFGAO_FILESYSANCESTOR            = 0x10000000,
    [helpstring("Is a folder")]
    SFGAO_FOLDER                     = 0x20000000,
    [helpstring("Is a file system thing (file/folder/root)")]
    SFGAO_FILESYSTEM                = 0x40000000,
    [helpstring("Expandable in the map pane")]
    SFGAO_HASSUBFOLDER              = 0x80000000,
    SFGAO_CONTENTSMASK              = 0x80000000,

    [helpstring("Invalidate cached information")]
    SFGAO_VALIDATE                  = 0x01000000,
    [helpstring("Is this removeable media?")]
    SFGAO_REMOVABLE                 = 0x02000000,
    [helpstring("Object is compressed (use alt color)")]
    SFGAO_COMPRESSED                = 0x04000000,
    [helpstring("Is in-place browsable")]
    SFGAO_BROWSABLE                 = 0x08000000,
    [helpstring("Is a non-enumerated object")]
    SFGAO_NONENUMERATED             = 0x00100000,
    [helpstring("Should show bold in explorer tree")]
    SFGAO_NEWCONTENT                = 0x00200000,

    SFGAO_CANMONIKER                = 0x00400000,
    SFGAO_HASSTORAGE                = 0x00400000,
    [helpstring("supports BindToObject(IID_IStream)")]
    SFGAO_STREAM                    = 0x00400000,     
    [helpstring("may contain children with SFGAO_STORAGE or SFGAO_STREAM")]
    SFGAO_STORAGEANCESTOR           = 0x00800000,
    [helpstring("for determining storage capabilities, ie for open/save semantics")]
    SFGAO_STORAGECAPMASK            = 0x70C50008,
} SFGAO_Flags;

[
  odl,
  uuid(000214E6-0000-0000-C000-000000000046),
  helpstring("IShellFolder interface")
]
interface IShellFolder : stdole.IUnknown {

    HRESULT ParseDisplayName(
        [in] long hwndOwner,
        [in] long pbcReserved,
        [in] long lpszDisplayName,
        [in, out] long* pchEaten,
        [in, out] long* ppidl,
        [in, out] long* pdwAttributes);

    HRESULT EnumObjects(
        [in] long hwndOwner,
        [in] SHCONTF grfFlags,
        [out, retval] IEnumIDList** ppenumIDList);

    HRESULT BindToObject(
        [in] long pidl,
        [in] long pbcReserved,
        [in, out] UUID *riid,
        [in, out] long *ppvOut);

    HRESULT BindToStorage(
        [in] long pidl,
        [in] long pbcReserved,
        [in, out] UUID *riid,
        [out, retval] IStorage **ppvObj);

    HRESULT CompareIDs(
        [in] long lparam,
        [in] long pidl1,
        [in] long pidl2);

    HRESULT CreateViewObject(
        [in] long hwndOwner,
        [in, out] UUID* riid,
        [out, retval] long *ppvOut);

    HRESULT GetAttributesOf(
        [in] long cidl,
        [in, out] long* apidl,
        [in, out] long* rgfInOut);

    HRESULT GetUIObjectOf(
        [in] long hwndOwner,
        [in] long cidl,
        [in, out] long *apidl,
        [in, out] UUID  *riid,
        [in, out] long *prgfInOut,
        [out, retval] long *ppvOut);

    HRESULT GetDisplayNameOf(
        [in] long pidl,
        [in] SHGNO_Flags uFlags,
        [in, out] STRRET* lpName);

    HRESULT SetNameOf(
        [in] long hwndOwner,
        [in] long pidl,
        [in] long lpszName,
        [in] SHGNO_Flags uFlags,
        [out, retval] long* ppidlOut);
};

typedef struct SHELLDETAILS {
    int     fmt;            // LVCFMT_* value (header only)
    int     cxChar;         // Number of "average" characters (header only)
    STRRET  str;            // String information
} SHELLDETAILS;

[
	odl,
    helpstring("IShellFolder2"),
    uuid(93F2F68C-1D1B-11d3-A30E-00C04F79ABD1),
]
interface IShellFolder2 : stdole.IUnknown {

    HRESULT ParseDisplayName(
        [in] long hwndOwner,
        [in] long pbcReserved,
        [in] long lpszDisplayName,
        [in, out] long* pchEaten,
        [in, out] long* ppidl,
        [in, out] long* pdwAttributes);

    HRESULT EnumObjects(
        [in] long hwndOwner,
        [in] SHCONTF grfFlags,
        [out, retval] IEnumIDList** ppenumIDList);

    HRESULT BindToObject(
        [in] long pidl,
        [in] long pbcReserved,
        [in, out] UUID *riid,
        [in, out] long *ppvOut);

    HRESULT BindToStorage(
        [in] long pidl,
        [in] long pbcReserved,
        [in, out] UUID *riid,
        [out, retval] IStorage **ppvObj);

    HRESULT CompareIDs(
        [in] long lparam,
        [in] long pidl1,
        [in] long pidl2);

    HRESULT CreateViewObject(
        [in] long hwndOwner,
        [in, out] UUID* riid,
        [out, retval] long *ppvOut);

    HRESULT GetAttributesOf(
        [in] long cidl,
        [in, out] long* apidl,
        [in, out] long* rgfInOut);

    HRESULT GetUIObjectOf(
        [in] long hwndOwner,
        [in] long cidl,
        [in, out] long *apidl,
        [in, out] UUID  *riid,
        [in, out] long *prgfInOut,
        [out, retval] long *ppvOut);

    HRESULT GetDisplayNameOf(
        [in] long pidl,
        [in] SHGNO_Flags uFlags,
        [in, out] STRRET* lpName);

    HRESULT SetNameOf(
        [in] long hwndOwner,
        [in] long pidl,
        [in] long lpszName,
        [in] SHGNO_Flags uFlags,
        [out, retval] long* ppidlOut);

    // Returns the guid of the search that is to be invoked when user clicks 
    // on the search toolbar button
    HRESULT GetDefaultSearchGUID(
		[out] GUID *pguid);

    // gives an enumerator of the searches to be added to the search menu
    HRESULT EnumSearches(
		[out, retval] IEnumExtraSearch **ppenum);
    
    HRESULT GetDefaultColumn(
		[in] LONG dwRes,
		[out] LONG *pSort,
		[out] LONG *pDisplay);

    // return SHCOLSTATE_ values
    HRESULT GetDefaultColumnState(
		[in] INT iColumn, 
		[out, retval] SHCOLSTATE *pcsFlags);
    
    HRESULT GetDetailsEx(
		[in] LONG pidl,
		[in] SHCOLUMNID *pscid,
		[out, retval] VARIANT *pv);
    
    HRESULT GetDetailsOf(
		[in] LONG pidl,
		[in] INT iColumn,
		[out] SHELLDETAILS *psd);
    
    HRESULT MapColumnToSCID(
		[in] INT iColumn,
		[in] SHCOLUMNID *pscid);
};

[
    odl,
    uuid(000214E5-0000-0000-C000-000000000046)
]
interface IShellIcon : stdole.Iunknown {

    typedef enum IconIndexes {
        Unknown = 0,
        Document = 1,
        Application = 2,
        ClosedFolder = 3,
        OpenFolder = 4,
    } IconIndexes;

    HRESULT GetIconOf(
        [in] LONG pidl,
        [in] GILFlags flags,
        [in] IconIndexes *lpIconIndex);

};


// IShellLink::Resolve fFlags
typedef enum SLR_FLAGS {
    SLR_NO_UI               = 0x0001,   // don't post any UI durring the resolve operation, not msgs are pumped
    SLR_ANY_MATCH           = 0x0002,   // no longer used
    SLR_UPDATE              = 0x0004,   // save the link back to it's file if the track made it dirty
    SLR_NOUPDATE            = 0x0008,
    SLR_NOSEARCH            = 0x0010,   // don't execute the search heuristics
    SLR_NOTRACK             = 0x0020,   // don't use NT5 object ID to track the link
    SLR_NOLINKINFO          = 0x0040,   // don't use the net and volume relative info
    SLR_INVOKE_MSI          = 0x0080,   // if we have a darwin link, then call msi to fault in the applicaion
    SLR_NO_UI_WITH_MSG_PUMP = 0x0101,   // SLR_NO_UI + requires an enable modeless site or HWND
} SLR_FLAGS;

// IShellLink::GetPath fFlags
typedef enum SLGP_FLAGS {
    SLGP_SHORTPATH      = 0x0001,
    SLGP_UNCPRIORITY    = 0x0002,
    SLGP_RAWPATH        = 0x0004,
} SLGP_FLAGS;

typedef struct WIN32_FIND_DATAA {
    LONG dwFileAttributes;
    CURRENCY ftCreationTime;
    CURRENCY ftLastAccessTime;
    CURRENCY ftLastWriteTime;
    LONG nFileSizeHigh;
    LONG nFileSizeLow;
    LONG dwReserved0;
    LONG dwReserved1;
    BYTE cFileName[260];
    BYTE cAlternateFileName[14];
} WIN32_FIND_DATAA;

typedef struct WIN32_FIND_DATAW {
    LONG dwFileAttributes;
    CURRENCY ftCreationTime;
    CURRENCY ftLastAccessTime;
    CURRENCY ftLastWriteTime;
    LONG nFileSizeHigh;
    LONG nFileSizeLow;
    LONG dwReserved0;
    LONG dwReserved1;
    BYTE cFileName[520];
    BYTE cAlternateFileName[28];
} WIN32_FIND_DATAW;

[
    uuid(000214EE-0000-0000-C000-000000000046),
    odl,
]
interface IShellLinkA : stdole.IUnknown {
    HRESULT GetPath(
		[in, out] LPSTR pszFile,
		[in] int cch,
		[in, out] WIN32_FIND_DATAA *pfd,
		[in] SLGP_FLAGS fFlags);
		
    HRESULT GetIDList(
		[out, retval] LONG *ppidl);
		
    HRESULT SetIDList(
		[in] LONG pidl);
		
    HRESULT GetDescription(
		[in, out] LPSTR pszName,
		[in] int cch);
		
    HRESULT SetDescription(
		[in] LPSTR pszName);
    
    HRESULT GetWorkingDirectory(
		[in, out] LPSTR pszDir,
		[in] int cch);
    
    HRESULT SetWorkingDirectory(
		[in] LPSTR pszDir);
    
    HRESULT GetArguments(
		[in, out] LPSTR pszArgs,
		[in] int cch);
    
    HRESULT SetArguments(
		[in] LPSTR pszArgs);
		
    HRESULT GetHotkey(
		[out, retval] short *pwHotkey);
		
    HRESULT SetHotkey(
		[in] short wHotkey);
    
    HRESULT GetShowCmd(
		[out, retval] int *piShowCmd);
		
    HRESULT SetShowCmd(
		[in] int iShowCmd);
		
    HRESULT GetIconLocation(
		[in, out] LPSTR pszIconPath,
		[in] int cch,
		[in, out] int *piIcon);
		
    HRESULT SetIconLocation(
		[in] LPSTR pszIconPath,
		[in] int iIcon);
    
    HRESULT SetRelativePath(
		[in] LPSTR pszPathRel,
		[in] LONG dwReserved);
    
    HRESULT Resolve(
		[in] LONG hwnd,
		[in] SLR_FLAGS fFlags);
    
    HRESULT SetPath(
		[in] LPSTR pszFile);
};

[
    odl,
    uuid(000214F9-0000-0000-C000-000000000046)
]
interface IShellLinkW : IUnknown
{
    HRESULT GetPath(
		[in, out] LPWSTR pszFile,
		[in] int cch,
		[in, out] WIN32_FIND_DATAW *pfd,
		[in] SLGP_FLAGS fFlags);
		
    HRESULT GetIDList(
		[out, retval] LONG *ppidl);
		
    HRESULT SetIDList(
		[in] LONG pidl);
		
    HRESULT GetDescription(
		[in, out] LPWSTR pszName,
		[in] int cch);
		
    HRESULT SetDescription(
		[in] LPWSTR pszName);
    
    HRESULT GetWorkingDirectory(
		[in, out] LPWSTR pszDir,
		[in] int cch);
    
    HRESULT SetWorkingDirectory(
		[in] LPWSTR pszDir);
    
    HRESULT GetArguments(
		[in, out] LPWSTR pszArgs,
		[in] int cch);
    
    HRESULT SetArguments(
		[in] LPWSTR pszArgs);
		
    HRESULT GetHotkey(
		[out, retval] short *pwHotkey);
		
    HRESULT SetHotkey(
		[in] short wHotkey);
    
    HRESULT GetShowCmd(
		[out, retval] int *piShowCmd);
		
    HRESULT SetShowCmd(
		[in] int iShowCmd);
		
    HRESULT GetIconLocation(
		[in, out] LPWSTR pszIconPath,
		[in] int cch,
		[in, out] int *piIcon);
		
    HRESULT SetIconLocation(
		[in] LPWSTR pszIconPath,
		[in] int iIcon);
    
    HRESULT SetRelativePath(
		[in] LPWSTR pszPathRel,
		[in] LONG dwReserved);
    
    HRESULT Resolve(
		[in] LONG hwnd,
		[in] SLR_FLAGS fFlags);
    
    HRESULT SetPath(
		[in] LPWSTR pszFile);
};

[
	odl,
    helpstring("IFolderView"),
    uuid(cde725b0-ccc9-4519-917e-325d72fab4ce),
]
interface IFolderView : stdole.IUnknown {
    
    HRESULT GetCurrentViewMode(
		[in, out] INT *pViewMode);
    
    HRESULT SetCurrentViewMode(
		[in] INT ViewMode);

    HRESULT GetFolder(
		[in] UUID *riid,
		[out] LONG *ppv);

    HRESULT Item(
		[in] int iItemIndex,
		[out] LONG *ppidl);

    HRESULT ItemCount(
		[in] INT uFlags,
		[out] int *pcItems);
    
    HRESULT Items(
		[in] INT uFlags,
		[in] UUID *riid,
		[out] LONG *ppv);

    HRESULT GetSelectionMarkedItem(
		[out] int *piItem);
		
    HRESULT GetFocusedItem(
		[out] int *piItem);

    HRESULT GetItemPosition(
		[in] LONG pidl,
		[out] POINT *ppt);
    
    HRESULT GetSpacing(
		[in, out] POINT *ppt);
		
    HRESULT GetDefaultSpacing(
		[out] POINT *ppt);
    
    HRESULT GetAutoArrange();

    // like IShellView::SelectItem() by index
    HRESULT SelectItem(
		[in] int iItem,
		[in] LONG dwFlags);
		
    HRESULT SelectAndPositionItems(
		[in] INT cidl, 
		[in] LONG *apidl,
		[in] POINT *apt,
		[in] LONG dwFlags);
};

[
	odl,
    helpstring("IFolderFilterSite"),
    uuid(C0A651F5-B48B-11d2-B5ED-006097C686F6)
]
interface IFolderFilterSite : stdole.IUnknown {
    HRESULT SetFilter(
		[in] stdole.IUnknown *punk);
};

[
	odl,
    helpstring("IFolderFilter"),
    uuid(9CC22886-DC8E-11d2-B1D0-00C04F8EEB3E),
]
interface IFolderFilter : stdole.IUnknown {
    
    HRESULT ShouldShow(
		[in] IShellFolder *psf,
		[in] LONG pidlFolder,
		[in] LONG pidlItem);
    
    HRESULT GetEnumFlags(
		[in] IShellFolder *psf,
		[in] LONG pidlFolder,
		[in] LONG *phwnd,
		[out] LONG *pgrfFlags);

};

[
	odl,
    helpstring("IProfferService"),
    uuid(cb728b20-f786-11ce-92ad-00aa00a74cd0), // IID_IProfferService
]
interface IProfferService : stdole.IUnknown {
    
    HRESULT ProfferService(
		[in] UUID *rguidService,
		[in] IServiceProvider *psp,
		[out, retval] LONG *pdwCookie);
    
    HRESULT RevokeService(
		[in] LONG dwCookie);
}

typedef enum PROPERTYUI_NAME_FLAGS {
    PUIFNF_DEFAULT          = 0x00000000,
    PUIFNF_MNEMONIC         = 0x00000001,   // include mnemonic in display name
} PROPERTYUI_NAME_FLAGS;

typedef enum PROPERTYUI_FLAGS {
    PUIF_DEFAULT            = 0x00000000,
    PUIF_RIGHTALIGN         = 0x00000001,   // this property should be right alligned
    PUIF_NOLABELININFOTIP   = 0x00000002,   // this property should not display a label in the infotip
} PROPERTYUI_FLAGS;

typedef enum PROPERTYUI_FORMAT_FLAGS {
    PUIFFDF_DEFAULT         = 0x00000000,
    PUIFFDF_RIGHTTOLEFT     = 0x00000001,   // BIDI support, right to left caller
    PUIFFDF_SHORTFORMAT     = 0x00000002,   // short format version of string
    PUIFFDF_NOTIME          = 0x00000004,   // truncate time to days, not hours/mins/sec
    PUIFFDF_FRIENDLYDATE    = 0x00000008,   // "Today", "Yesterday", etc
    PUIFFDF_NOUNITS         = 0x00000010,   // don't do "KB", "MB", "KHz"
} PROPERTYUI_FORMAT_FLAGS;

[
	odl,
    helpstring("IPropertyUI - helper for dealing with properties ala IPropertySetStorage"),
    uuid(757a7d9f-919a-4118-99d7-dbb208c8cc66),
]
interface IPropertyUI : stdole.IUnknown {

    HRESULT ParsePropertyName(
		[in] LPWSTR pszName,
		[out] UUID *pUUID,
		[out] LONG *ppid,
		[in, out] LONG *pchEaten);
		 
    HRESULT GetCannonicalName(
		[in] UUID *UUID,
		[in] LONG pid,
		[out] LPWSTR pwszText,
		[in] LONG cchText);
		
    HRESULT GetDisplayName(
		[in] UUID *UUID,
		[in] LONG pid,
		[in] PROPERTYUI_NAME_FLAGS flags,
		[out] LPWSTR pwszText,
		[in] LONG cchText);
    
    HRESULT GetPropertyDescription(
		[in] UUID *UUID,
		[in] LONG pid,
		[out] LPWSTR pwszText,
		[in] LONG cchText);
		
    HRESULT GetDefaultWidth(
		[in] UUID *UUID,
		[in] LONG pid,
		[out, retval] LONG *pcxChars);
    
    HRESULT GetFlags(
		[in] UUID *UUID,
		[in] LONG pid,
		[out, retval] PROPERTYUI_FLAGS *pFlags);
		
    HRESULT FormatForDisplay(
		[in] UUID *UUID,
		[in] LONG pid,
		[in] VARIANT *pvar,
		[in] PROPERTYUI_FORMAT_FLAGS flags,
		[out] LPWSTR pwszText,
		[in] LONG cchText);
		
    HRESULT GetHelpInfo(
		[in] UUID *UUID,
		[in] LONG pid,
		[out] LPWSTR pwszHelpFile,
		[in] LONG cch,
		[out] LONG *puHelpID);
};


[
	odl,
    helpstring("ICategoryProvider - Provides a list of catagorizers registered on an IShellFolder"),
    uuid(9af64809-5864-4c26-a720-c1f78c086ee3)
]
interface ICategoryProvider : stdole.IUnknown {
    
    [helpstring("Returns S_OK if the view should display this column in category selection UI, or S_FALSE to remove it.")]
    HRESULT CanCategorizeOnSCID(
		[in] SHCOLUMNID *pscid);
    
    [helpstring("Returns either a GUID to create in CreateCategory, or a SHCOLUNNID that is used by the default categorizer. Return S_FALSE if you do not support a default group. GUID_NULL returned in pguid indicates to the client to use pscid as the default category.")]
    HRESULT GetDefaultCategory(
		[in, out] UUID *pguid,
		[in, out] SHCOLUMNID *pscid);
    
    [helpstring("Returns either a GUID that represents the categoizer to use for the specified SHCOLUMNID.")]
    HRESULT GetCategoryForSCID(
		[in, out] SHCOLUMNID* pscid,
		[in, out] UUID* pguid);
    
    [helpstring("Returns an IEnumGUID that has a list of GUIDs that represent categories.")]
    HRESULT EnumCategories(
		[out, retval] IEnumGUID ** penum);
		
    [helpstring("Returns the name of the given category.")]
    HRESULT GetCategoryName(
		[in, out] UUID* pguid,
		[in, out] LPWSTR pszName,
		[in] INT cch);
    
    [helpstring("Creates the category.")]
    HRESULT CreateCategory(
		[in, out] UUID * pguid,
		[in, out] UUID *riid,
		[in, out] LONG * ppv);
}

typedef enum CATEGORYINFO_FLAGS {
    CATINFO_NORMAL      = 0x00000000,   // Apply default properties to this category
    CATINFO_COLLAPSED   = 0x00000001,   // This category should appear collapsed. useful for the "None" category. 
    CATINFO_HIDDEN      = 0x00000002,   // This category should follow the "Hidden" files setting for being displayed
} CATEGORYINFO_FLAGS;

typedef enum CATSORT_FLAGS {
    CATSORT_DEFAULT     = 0x00000000,   // Default Sort order
    CATSORT_NAME        = 0x00000001,   // Sort by name
} CATSORT_FLAGS;

typedef struct CATEGORY_INFO {
    CATEGORYINFO_FLAGS cif;
    BYTE wszName[520];
} CATEGORY_INFO;

[
	odl,
    helpstring("ICategorizer - Categorizes pidls"),
    uuid(a3b14589-9174-49a8-89a3-06a1ae2b9ba7),
]
interface ICategorizer : stdole.IUnknown {
    
    [helpstring("Returns the description of this category that will be displayed in the UI")]
    HRESULT GetDescription(
		[in, out] LPWSTR pszDesc,
		[in] INT cch);
		
    [helpstring("Returns a list of categories associated with a list of ID Lists. NOTE: -1 is an invalid Category ID, and they cannot be persisted")]
    HRESULT GetCategory(
		[in] INT cidl, 
		[in] LONG *apidl, 
		[in, out] LONG *rgCategoryIds);

    [helpstring("Returns information about the category, such as default display and the text to display in the UI")]
    HRESULT GetCategoryInfo(
		[in] LONG dwCategoryId,
		[in] CATEGORY_INFO *pci);
    
    [helpstring("Returns HRESULTFromShort. -1, 0, 1 indicate the comparison of the IDs. Used for sorting categories in the UI")]
    HRESULT CompareCategory(
		[in] CATSORT_FLAGS csfFlags, 
		[in] LONG dwCategoryId1,
		[in] LONG dwCategoryId2);
};

[
    odl,
    uuid(49ff1172-eadc-446d-9285-156453a6431c),
    helpstring("IActionProgressDialog Interface"),
]
interface IActionProgressDialog : stdole.IUnknown {
    
    typedef enum SPINITF{
        SPINITF_NORMAL         = 0x00000000,      // default normal progress behavior
        SPINITF_MODAL          = 0x00000001,      // call punkSite->EnableModeless() or EnableWindow()
        SPINITF_NOMINIMIZE     = 0x00000008,      // Do not have a minimize button in the caption bar.
    } SPINITF;
    
    HRESULT Initialize(
        [in] SPINITF flags, 
        [in] LPWSTR pszTitle, 
        [in] LPWSTR pszCancel);

    HRESULT Stop();
    
};

[
    odl,
    uuid(C1FB73D0-EC3A-4ba2-B512-8CDB9187B6D1),
    helpstring("HW Event Handler Interface"),
]
interface IHWEventHandler : stdole. IUnknown {
    
    HRESULT Initialize(
		[in] LPWSTR pszParams);
		
    HRESULT HandleEvent(
		[in] LPWSTR pszDeviceID,
        [in] LPWSTR pszAltDeviceID,
        [in] LPWSTR pszEventType);
        
    HRESULT HandleEventWithContent(
		[in] LPWSTR pszDeviceID,
        [in] LPWSTR pszAltDeviceID,
        [in] LPWSTR pszEventType,
        [in] LPWSTR pszContentTypeHandler,
        [in] IDataObject* pdataobject);
};

typedef enum ARCONTENT {
	ARCONTENT_AUTORUNINF      = 0x00000002, // That's the one we have today, and always had 
	ARCONTENT_AUDIOCD         = 0x00000004, // Audio CD (not MP3 and the like, the stuff you buy at the store) 
	ARCONTENT_DVDMOVIE        = 0x00000008, // DVD Movie (not MPEGs, the stuff you buy at the store) 
	ARCONTENT_BLANKCD         = 0x00000010, // Blank CD-R/CD-RW 
	ARCONTENT_BLANKDVD        = 0x00000020, // Blank DVD-R/DVD-RW 
	ARCONTENT_UNKNOWNCONTENT  = 0x00000040, // Whatever files.  Mean that it's formatted.
	ARCONTENT_AUTOPLAYPIX     = 0x00000080, // Whatever files.  Mean that it's formatted.
	ARCONTENT_AUTOPLAYMUSIC   = 0x00000100, // Whatever files.  Mean that it's formatted.
	ARCONTENT_AUTOPLAYVIDEO   = 0x00000200, // Whatever files.  Mean that it's formatted.
} ARCONTENT;

[
    odl,
    uuid(DDEFE873-6997-4e68-BE26-39B633ADBE12),
    helpstring("must register in ROT")
]
interface IQueryCancelAutoPlay : stdole.IUnknown {
    
    HRESULT AllowAutoPlay(
		[in] LONG pszPath,
		[in] ARCONTENT dwContentType,
		[in] LONG pszLabel,
		[in] LONG dwSerialNumber);
};

[
    odl,
    uuid(49ff1173-eadc-446d-9285-156453a6431c),
    helpstring("IActionProgress Interface")
]
interface IActionProgress : stdole.IUnknown {
    
    typedef enum SPBEGINF {
        SPBEGINF_NORMAL             = 0x00000000,      // default normal progress behavior
        SPBEGINF_AUTOTIME           = 0x00000002,      // automatically updates the "time remaining" text 
        SPBEGINF_NOPROGRESSBAR      = 0x00000010,      // Don't display the progress bar (SetProgress() wont be called)
        SPBEGINF_MARQUEEPROGRESS    = 0x00000020,      // use marquee progress (comctl32 v6 required)
    } SPBEGINF;

    typedef enum SPACTION {
        SPACTION_NONE            = 0,      
        SPACTION_MOVING,
        SPACTION_COPYING,
        SPACTION_RECYCLING,
        SPACTION_APPLYINGATTRIBS,
        SPACTION_DOWNLOADING,
        SPACTION_SEARCHING_INTERNET,
        SPACTION_CALCULATING,
        SPACTION_UPLOADING,
        SPACTION_SEARCHING_FILES,
    } SPACTION;
    
    HRESULT Begin(
        [in] SPACTION action,
        [in] SPBEGINF flags);

    HRESULT UpdateProgress(
        [in] CURRENCY ulCompleted, 
        [in] CURRENCY ulTotal);

    typedef enum SPTEXT {
        SPTEXT_ACTIONDESCRIPTION        = 1,
        SPTEXT_ACTIONDETAIL,
    } SPTEXT;
    
    HRESULT UpdateText(
        [in] SPTEXT sptext, 
        [in] LPWSTR pszText, 
        [in] BOOL fMayCompact);

    HRESULT QueryCancel(
		[out, retval] BOOL *pfCancelled);
		
    HRESULT ResetCancel();

    HRESULT End();

};

[
	odl,
    uuid(7307055c-b24a-486b-9f25-163e597a28a9)
]
interface IQueryContinue : stdole.IUnknown {
    HRESULT QueryContinue();    // S_OK -> Continue, other 
};

[
	odl,
    uuid(ba9711ba-5893-4787-a7e1-41277151550b)
]
interface IUserNotification : stdole.IUnknown {
    
    HRESULT SetBalloonInfo(
		[in] LPWSTR pszTitle,
		[in] LPWSTR pszText,
		[in] LONG dwInfoFlags);

	// times in msec
    HRESULT SetBalloonRetry(
		[in] LONG dwShowTime,
		[in] LONG dwInterval,
		[in] INT cRetryCount);
		
    HRESULT SetIconInfo(
		[in] LONG hIcon,
		[in] LPWSTR pszToolTip);
		
    HRESULT Show(
		[in] IQueryContinue *pqc,
		[in] LONG dwContinuePollInterval);
		
    HRESULT PlaySound(
		[in] LPWSTR pszSoundName);
};

[
	odl,
    helpstring("IUserEventTimerCallback "),
    uuid(e9ead8e6-2a25-410e-9b58-a9fbef1dd1a2),     // IID_IUserEventTimerCallback
]
interface IUserEventTimerCallback : stdole.IUnknown {
    
    HRESULT UserEventTimerProc(
        [in] LONG uUserEventTimerID,
        [in] INT uTimerElapse); 
};

[
	odl,
    helpstring("IUserEventTimer "),
    uuid(0F504B94-6E42-42E6-99E0-E20FAFE52AB4),     // IID_IUserEventTimer
]
interface IUserEventTimer : stdole.IUnknown {
    
    HRESULT SetUserEventTimer (
        [in] LONG hWnd,
        [in] INT uCallbackMessage,
        [in] INT uTimerElapse,
        [in] IUserEventTimerCallback *pUserEventTimerCallback,
        [in, out] LONG * puUserEventTimerID);

    HRESULT KillUserEventTimer(    
        [in] LONG hWnd,
        [in] LONG uUserEventTimerID);

    HRESULT GetUserEventTimerElapsed(
        [in] LONG hWnd,
        [in] LONG uUserEventTimerID,
        [out, retval] LONG *puTimerElapsed);

    HRESULT InitTimerTickInterval(
        [in] INT uTimerTickIntervalMs);                            
};

[
	odl,
    uuid(56FDF342-FD6D-11d0-958A-006097C9A090)
]
interface ITaskbarList : stdole.IUnknown {
    
    HRESULT HrInit();
    
    HRESULT AddTab(
		[in] LONG hwnd);
		
    HRESULT DeleteTab(
		[in] LONG hwnd);
		
    HRESULT ActivateTab(
		[in] LONG hwnd);
		
    HRESULT SetActiveAlt(
		[in] LONG hwnd);
};

[
	odl,
    uuid(602D4995-B13A-429b-A66E-1935E44F4317)
]
interface ITaskbarList2 : ITaskbarList {
    
    HRESULT MarkFullscreenWindow(
		[in] LONG hwnd,
		[in] BOOL fFullscreen);
		
};

typedef [helpstring("Flags for GetWallpaperOptions and SetWallpaperOptions")]
enum WPSTYLE {
	WPSTYLE_CENTER = 0,
	WPSTYLE_TILE = 1,
	WPSTYLE_STRETCH = 2,
	WPSTYLE_MAX = 3,
} WPSTYLE;

typedef struct WALLPAPEROPT {
	[helpstring("size of this Structure")]
	LONG dwSize;
	[helpstring("Style")]
	WPSTYLE dwStyle;
} WALLPAPEROPT;

typedef struct COMPONENTSOPT {
	[helpstring("Size of this structure")]
	LONG dwSize;
	[helpstring("Enable components?")]
	BOOL fEnableComponents;
	[helpstring("Active desktop enabled ?")]
	BOOL fActiveDesktop;
} COMPONENTSOPT;

typedef struct COMPPOS {
	[helpstring("Size of this structure")]
	LONG dwSize;
	[helpstring("Left of top-left corner in screen co-ordinates.")]
	LONG iLeft;
	[helpstring("Top of top-left corner in screen co-ordinates.")]
	LONG iTop;
	[helpstring("Width in pixels.")]
	LONG dwWidth;
	[helpstring("Height in pixels.")]
	LONG dwHeight;
	[helpstring("Indicates the Z-order of the component.")]
	LONG izIndex;
	[helpstring("Is the component resizeable?")]
	BOOL fCanResize;
	[helpstring("Resizeable in X-direction?")]
	BOOL fCanResizeX;
	[helpstring("Resizeable in Y-direction?")]
	BOOL fCanResizeY;
	[helpstring("Left of top-left corner as percent of screen width")]
	LONG iPreferredLeftPercent;
	[helpstring("Top of top-left corner as percent of screen height")]
	LONG iPreferredTopPercent;
} COMPPOS;

typedef struct COMPSTATEINFO {
	[helpstring("Size of this structure.")]
	LONG dwSize;
	[helpstring("Left of the top-left corner in screen co-ordinates.")]
	LONG iLeft;
	[helpstring("Top of top-left corner in screen co-ordinates.")]
	LONG iTop;
	[helpstring("Width in pixels.")]
	LONG dwWidth;
	[helpstring("Height in pixels.")]
	LONG dwHeight;
	[helpstring("State of the component (full-screen mode or split-screen or normal state.")]
	LONG dwItemState;
} COMPSTATEINFO;

typedef enum COMZORDER {
	[helpstring("izOrder value meaning component is at the top")]
	COMPONENT_TOP = (0x3fffffff)
} COMZORDER;

// iCompType values
typedef enum COMP_TYPE {
	COMP_TYPE_HTMLDOC       = 0,
	COMP_TYPE_PICTURE       = 1,
	COMP_TYPE_WEBSITE       = 2,
	COMP_TYPE_CONTROL       = 3,
	COMP_TYPE_CFHTML        = 4,
	COMP_TYPE_MAX           = 4,
} COMP_TYPE;

// The following is the COMPONENT structure used 
// in IE4.01, IE4.0 and Memphis. It is kept here 
// for compatibility reasons.

typedef struct IE4COMPONENT {
	LONG dwSize;                  //Size of this structure
	LONG dwID;                    //Reserved: Set it always to zero.
	COMP_TYPE iComponentType;     //One of COMP_TYPE_*
	BOOL fChecked;                // Is this component enabled?
	BOOL fDirty;                  // Had the component been modified and not yet saved to disk?
	BOOL fNoScroll;               // Is the component scrollable?
	COMPPOS cpPos;                // Width, height etc.,
	BYTE wszFriendlyName[520];    // Friendly name of component.
	BYTE wszSource[4168];         //URL of the component.
	BYTE wszSubscribedURL[4168];  //Subscrined URL
} IE4COMPONENT;

//
// The following is the new NT5 component structure. 
// Note that the initial portion of this component exactly
// matches the IE4COMPONENT structure. All new fields 
// are added at the bottom and the dwSize field is used to
// distinguish between IE4COMPONENT and the new COMPONENT 
// structures.
//

// Defines for dwCurItemState
typedef enum ITEMSTATE {
	IS_NORMAL             = 0x00000001,
	IS_FULLSCREEN         = 0x00000002,
	IS_SPLIT              = 0x00000004,
	IS_VALIDSIZESTATEBITS = 0x00000007,  // The set of IS_* state bits which define the "size" of the component - these bits are mutually exclusive.
	IS_VALIDSTATEBITS     = 0xC0000007,  // All of the currently defined IS_* bits.
} ITEMSTATE;

typedef struct COMPONENT {
	LONG dwSize;             //Size of this structure
	LONG dwID;               //Reserved: Set it always to zero.
	COMP_TYPE iComponentType;     //One of COMP_TYPE_*
	BOOL fChecked;           // Is this component enabled?
	BOOL fDirty;             // Had the component been modified and not yet saved to disk?
	BOOL fNoScroll;          // Is the component scrollable?
	COMPPOS cpPos;              // Width, height etc.,
	BYTE wszFriendlyName[520];    // Friendly name of component.
	BYTE wszSource[4168];         //URL of the component.
	BYTE wszSubscribedURL[4168];  //Subscrined URL

	// New fields are added below. 
	// Everything above here must exactly match 
	// the IE4COMPONENT Structure.

	ITEMSTATE       dwCurItemState; // Current state of the Component.
	COMPSTATEINFO   csiOriginal;    // Original state of the component when it was first added.
	COMPSTATEINFO   csiRestored;    // Restored state of the component.
} COMPONENT;


// Flags for IActiveDesktop::ApplyChanges()
typedef enum AD_APPLY {
	AD_APPLY_SAVE             = 0x00000001,
	AD_APPLY_HTMLGEN          = 0x00000002,
	AD_APPLY_REFRESH          = 0x00000004,
	AD_APPLY_ALL              = 0x00000007,
	AD_APPLY_FORCE            = 0x00000008,
	AD_APPLY_BUFFERED_REFRESH = 0x00000010,
	AD_APPLY_DYNAMICREFRESH   = 0x00000020,
} ADAPPLY;


// Flags for IActiveDesktop::ModifyComponent()
typedef enum COMP_ELEM {
	COMP_ELEM_TYPE          = 0x00000001,
	COMP_ELEM_CHECKED       = 0x00000002,
	COMP_ELEM_DIRTY         = 0x00000004,
	COMP_ELEM_NOSCROLL      = 0x00000008,
	COMP_ELEM_POS_LEFT      = 0x00000010,
	COMP_ELEM_POS_TOP       = 0x00000020,
	COMP_ELEM_SIZE_WIDTH    = 0x00000040,
	COMP_ELEM_SIZE_HEIGHT   = 0x00000080,
	COMP_ELEM_POS_ZINDEX    = 0x00000100,
	COMP_ELEM_SOURCE        = 0x00000200,
	COMP_ELEM_FRIENDLYNAME  = 0x00000400,
	COMP_ELEM_SUBSCRIBEDURL = 0x00000800,
	COMP_ELEM_ORIGINAL_CSI  = 0x00001000,
	COMP_ELEM_RESTORED_CSI  = 0x00002000,
	COMP_ELEM_CURITEMSTATE  = 0x00004000,
	COMP_ELEM_ALL           = 0x00007FFF,
} COMP_ELEM;

// Flags for IActiveDesktop::AddDesktopItemWithUI()
typedef enum DTI_ADTIWUI {
	DTI_ADDUI_DEFAULT       = 0x00000000,
	DTI_ADDUI_DISPSUBWIZARD = 0x00000001,
	DTI_ADDUI_POSITIONITEM  = 0x00000002,
} DTI_ADTIWUI;


// Flags for IActiveDesktop::AddUrl()
typedef enum ADDURL {
	ADDURL_SILENT = 0x0001,
} ADDURL;

// Default positions for ADI
typedef enum DEFAULTPOS {
	COMPONENT_DEFAULT_LEFT    = 0xFFFF,
	COMPONENT_DEFAULT_TOP     = 0xFFFF,
} DEFAULTPOS;

//
//  Interface for manipulating the Active Desktop.
//

[
	odl,
	uuid(F490EB00-1240-11D1-9888-006097DEACF9),
	helpstring("IActiveDesktop interface")
]
interface IActiveDesktop : IUnknown {

	// IActiveDesktop methods

	HRESULT ApplyChanges(
		[in] ADAPPLY dwFlags);

	LONG GetWallpaper(
		[in, out] LPWSTR pwszWallpaper,
		[in] LONG cchWallpaper,
		[in, defaultvalue(0)] LONG dwReserved);

	LONG SetWallpaper(
		[in] LPWSTR pwszWallpaper,
		[in, defaultvalue(0)] LONG dwReserved);

	LONG GetWallpaperOptions(
		[in, out] WALLPAPEROPT *pwpo,
		[in, defaultvalue(0)] LONG dwReserved);

	LONG SetWallpaperOptions(
		[in] WALLPAPEROPT *pwpo,
		[in, defaultvalue(0)] LONG dwReserved);

	LONG GetPattern(
		[in, out] LPWSTR pwszPattern,
		[in] LONG cchPattern, 
		[in, defaultvalue(0)] LONG dwReserved);

	LONG SetPattern(
		[in] LPWSTR pwszPattern,
		[in, defaultvalue(0)] LONG dwReserved);

	LONG GetDesktopItemOptions(
		[in] void *pco,
		[in, defaultvalue(0)] LONG dwReserved);

	LONG SetDesktopItemOptions(
		[in] void *pco,
		[in, defaultvalue(0)] LONG dwReserved);

	LONG AddDesktopItem(
		[in] void *pcomp,
		[in, defaultvalue(0)] LONG dwReserved);

	LONG AddDesktopItemWithUI(
		[in] LONG hwnd, 
		[in] void *pcomp, 
		[in, defaultvalue(0)] LONG dwReserved);

	LONG ModifyDesktopItem(
		[in] void *pcomp,
		[in] COMP_ELEM dwFlags);

	LONG RemoveDesktopItem(
		[in] void *pcomp, 
		[in, defaultvalue(0)] LONG dwReserved);

	LONG GetDesktopItemCount(
		[in, out] LONG *lpiCount,
		[in, defaultvalue(0)] LONG dwReserved);

	LONG GetDesktopItem(
		[in] LONG nComponent,
		[in, out] void *pcomp,
		[in, defaultvalue(0)] LONG dwReserved);

	LONG GetDesktopItemByID(
		[in] LONG dwID,
		[in] void *pcomp,
		[in, defaultvalue(0)] LONG dwReserved);

	LONG GenerateDesktopItemHtml(
		[in] LPWSTR pwszFileName,
		[in] void *pcomp,
		[in, defaultvalue(0)] LONG dwReserved);

	LONG AddUrl(
		[in] LONG hwnd,
		[in] LPWSTR pszSource,
		[in] void *pcomp,
		[in, defaultvalue(0)] LONG dwFlags);

	LONG GetDesktopItemBySource(
		[in] LPWSTR pwszSource, 
		[in] void *pcomp, 
		[in, defaultvalue(0)] LONG dwReserved);
};

[
	uuid(75048700-EF1F-11D0-9888-006097DEACF9),
	helpstring("ActiveDesktop class")
]
coclass ActiveDesktop {
	[default] interface IActiveDesktop;
}

[
    odl,
    uuid(3d73a659-e5d0-4d42-afc0-5121ba425c8d),     // IID_ICDBurn
]
interface ICDBurn : stdole.IUnknown {

    HRESULT GetRecorderDriveLetter(
		[out] LPWSTR pszDrive,
		[in] INT cch);
		
    HRESULT Burn(
		[in] LONG hwnd);
    
    HRESULT HasRecordableDrive(
		[out, retval] BOOL *pfHasRecorder);
};

[
	odl,
    uuid(C9D81948-443A-40C7-945C-5E171B8C66B4)
]
interface IAddressBarParser : stdole.IUnknown {

    HRESULT ParseFromOutsideSource(
		[in] LPWSTR pcszUrlIn,
		[in] LONG dwParseFlags,
		[out, retval] BOOL *pfWasCorrected);
		
    HRESULT GetUrl(
		[out] LPWSTR pszUrlOut,
		[in] LONG cchUrlOutSize);
		
    HRESULT SetUrl(
		[in] LPWSTR pcszUrlIn,
		[in] LONG dwGenType);
		
    HRESULT GetDisplayName(
		[out] LPWSTR pszUrlOut,
		[in] LONG cchUrlOutSize);
		
    HRESULT GetPidl(
		[out, retval] LONG *ppidl);
		
    HRESULT SetPidl(
		[in] LONG  pidl);
		
    HRESULT GetArgs(
		[out] LPWSTR pszArgsOut,
		[in] LONG cchArgsOutSize);
    
    HRESULT AddPath(
		[in] LONG pidl);
};

// Wizard Extension objects.  These interfaces defined methods for extending 
// Win32 wizard in a progromatic way.

// Range of ID's that extensions can used, these mustn't clash with
// the existing wizards dialog IDS.  (That enables them to still
// do PropSheet_SetCurSelByID).

typedef enum IDD_WIZEXTN {
	IDD_WIZEXTN_FIRST    = 0x5000,
	IDD_WIZEXTN_LAST     = 0x5100,
} IDD_WIZEXTN;

// This site object is requested via a QueryService of the objects site,
// it allows the extension to navigate in/out out itself, eg. when the
// extension has shown all of its pages and wants to navigate to the
// next page it would call GetNextPage and select the specified HPAGE.

[
	odl,
    uuid(88960f5b-422f-4e7b-8013-73415381c3c3),
    helpstring("Wizard Extension Site"),
]
interface IWizardSite : stdole.IUnknown {

    HRESULT GetPreviousPage(
		[out, retval] LONG *phpage);
    
    HRESULT GetNextPage(
		[out, retval] LONG *phpage);
    
    HRESULT GetCancelledPage(
		[out, retval] LONG *phpage);
};

// A wizard extension is implemented using this object, the extension will declare the
// pages that it supports using the AddPages method, and then when its host needs to navigate
// into the extenion it will do so via GetFirstPage and selecting that.

[
	odl,
    uuid(c02ea696-86cc-491e-9b23-74394a0444a8),
    helpstring("Wizard Extension"),
]
interface IWizardExtension : stdole.IUnknown {
    
    HRESULT AddPages(
		[in, out] LONG *aPages,
		[in] INT cPages,
		[out, retval] INT *pnPagesAdded);
		
    HRESULT GetFirstPage(
		[out, retval] LONG *phpage);
		
    HRESULT GetLastPage(
		[out, retval] LONG *phpage);
};

// The Web Wizard is a HTML host for wizard pages, it allows you
// create a HTML wizard starting at the URL defined via SetInitialURL.

[
	odl,
    helpstring("Web Wizard Page Extension"),
    uuid(0e6b3f66-98d1-48c0-a222-fbde74e2fbc5),
]
interface IWebWizardExtension : IWizardExtension {
    
    HRESULT SetInitialURL(
		[in] LPWSTR pszURL);
		
    HRESULT SetErrorURL(
		[in] LPWSTR pszErrorURL);
}

typedef enum SHPWHF {
	SHPWHF_NORECOMPRESS             = 0x00000001,  // don't allow/prompt for recompress of streams")
	SHPWHF_NONETPLACECREATE         = 0x00000002,  // don't create a network place when transfer is complete")
	SHPWHF_NOFILESELECTOR           = 0x00000004,  // don't show the file selector")
	SHPWHF_VALIDATEVIAWEBFOLDERS    = 0x00010000,  // enable web folders to validate network places (ANP support)")
} SHPWHF;

[ 
	odl,
    helpstring("Web Publishing Wizard"),
    uuid(aa9198bb-ccec-472d-beed-19a4f6733f7a),
]
interface IPublishingWizard : IWizardExtension {
    
    HRESULT Initialize(
		[in] IDataObject *pdo,
		[in] LONG dwOptions,
		[in] LPWSTR pszServiceProvider);
		
    HRESULT GetTransferManifest(
		[out] LONG *phrFromTransfer,
		[out] stdole.IUnknown **pdocManifest);
}

[
	odl,
    uuid(1ea58f02-d55a-411d-b09e-9e65ac21605b),
    helpstring("Shell Folder Host"),
]
interface IFolderViewHost : stdole.IUnknown {
    HRESULT Initialize(
		[in] LONG hwndParent,
		[in] IDataObject *pdo,
		[in] RECT *prc);
};

typedef enum PROGDLG {
    PROGDLG_NORMAL          = 0x00000000,      // default normal progress dlg behavior
    PROGDLG_MODAL           = 0x00000001,      // the dialog is modal to its hwndParent (default is modeless)
    PROGDLG_AUTOTIME        = 0x00000002,      // automatically updates the "Line3" text with the "time remaining" (you cant call SetLine3 if you passs this!)
    PROGDLG_NOTIME          = 0x00000004,      // we dont show the "time remaining" if this is set. We need this if dwTotal < dwCompleted for sparse files
    PROGDLG_NOMINIMIZE      = 0x00000008,      // Do not have a minimize button in the caption bar.
    PROGDLG_NOPROGRESSBAR   = 0x00000010,      // Don't display the progress bar
} PROGDLG;

// Time Actions (dwTimerAction)
typedef enum PDTIMER {
    PDTIMER_RESET       = 0x00000001,       // Reset the timer so the progress will be calculated from now until the first ::SetProgress() is called so
                                            // those this time will correspond to the values passed to ::SetProgress().  Only do this before ::SetProgress() is called.
} PDTIMER;

[
    odl,
    uuid(EBBC7C04-315E-11d2-B62F-006097DF5BD4)
]
interface IProgressDialog : stdole.IUnknown {

    HRESULT StartProgressDialog(
        [in] LONG hwndParent,
        [in] IUnknown *punkEnableModless,
        [in] PROGDLG dwFlags,
        [in, defaultvalue(0)] void *pvResevered);

    HRESULT StopProgressDialog();

    HRESULT SetTitle(
        [in] LPWSTR pwzTitle);

    HRESULT SetAnimation(
        [in] LONG hInstAnimation,
        [in] LONG idAnimation);

    BOOL HasUserCancelled();

    HRESULT SetProgress(
        [in] LONG dwCompleted,
        [in] LONG dwTotal);

    HRESULT SetProgress64(
        [in] CURRENCY ullCompleted,
        [in] CURRENCY ullTotal);

    HRESULT SetLine(
        [in] LONG dwLineNum,
        [in] LPWSTR pwzString,
        [in] BOOL fCompactPath,
        [in, defaultvalue(0)] void *pvResevered);

    HRESULT SetCancelMsg(
        [in] LPWSTR pwzCancelMsg,
        [in, defaultvalue(0)] void *pvResevered);

    HRESULT Timer(
        [in] LONG dwTimerAction,
        [in, defaultvalue(0)] void *pvResevered);
};

typedef enum SNCF {
	SNCF_REFRESHLIST = 0x00000001  // refresh the list (eg. from F5 or opening a folder)")
} SNCF;

[
    odl,
    helpstring("INetCrawler"),
    uuid(49c929ee-a1b7-4c58-b539-e63be392b6f3),
]
interface INetCrawler : stdole.IUnknown {
    HRESULT Update(
		[in] SNCF dwFlags);
}

[
    odl,
    helpstring("Interface to allow the IDataObject operation to occur asynchronously."),
    uuid(3D8B0590-F691-11d2-8EA9-006097DF5BD4)
]
interface IAsyncOperation : stdole.IUnknown {

    HRESULT SetAsyncMode(
        [in] BOOL fDoOpAsync);

    HRESULT GetAsyncMode(
        [out, retval] BOOL * pfIsOpAsync);

    HRESULT StartOperation(
        [in] IBindCtx * pbcReserved);

    HRESULT InOperation(
        [out, retval] BOOL * pfInAsyncOp);

    HRESULT EndOperation(
        [in] HRESULT hResult, 
        [in] IBindCtx *pbcReserved,
        [in] LONG dwEffects);
}

// CLSID_QueryCancelAutoPlay
[ uuid(331F1768-05A9-4ddd-B86E-DAE34DDC998A) ]
coclass QueryCancelAutoPlay {
	interface IQueryCancelAutoPlay;
}

// CLSID_DriveSizeCategorizer
[ uuid(94357B53-CA29-4b78-83AE-E8FE7409134F) ]
coclass DriveSizeCategorizer {
	interface ICategorizer;
}

// CLSID_DriveTypeCategorizer
[ uuid(B0A8F3CF-4333-4bab-8873-1CCB1CADA48B) ]
coclass DriveTypeCategorizer {
	interface ICategorizer;
}

// CLSID_FreeSpaceCategorizer
[ uuid(B5607793-24AC-44c7-82E2-831726AA6CB7) ]
coclass FreeSpaceCategorizer {
	interface ICategorizer;
}

// CLSID_TimeCategorizer
[ uuid(3bb4118f-ddfd-4d30-a348-9fb5d6bf1afe) ]
coclass TimeCategorizer {
	interface ICategorizer;
}

// CLSID_SizeCategorizer
[ uuid(55d7b852-f6d1-42f2-aa75-8728a1b2d264) ]
coclass SizeCategorizer {
	interface ICategorizer;
}

// CLSID_AlphabeticalCategorizer
[ uuid(3c2654c6-7372-4f6b-b310-55d6128f49d2) ]
coclass AlphabeticalCategorizer {
	interface ICategorizer;
}

// CLSID_MergedCategorizer
[ uuid(8e827c11-33e7-4bc1-b242-8cd9a1c2b304) ]
coclass MergedCategorizer {
	interface ICategorizer;
}

// CLSID_ImageProperties
[ uuid(7ab770c7-0e23-4d7a-8aa2-19bfad479829) ]
coclass ImageProperties {
	interface stdole.IUnknown;
//	interface IPersistFile;
}

// CLSID_ImageProperties
[ uuid(d912f8cf-0396-4915-884e-fb425d32943b) ]
coclass PropertiesUI {
	interface IPropertyUI;
}

// CLSID_UserNotification
[ uuid(0010890e-8789-413c-adbc-48f5b511b3af) ]
coclass UserNotification {
	interface IUserNotification;
}

// CLSID_UserEventTimerCallback
[ uuid(15fffd13-5140-41b8-b89a-c8d5759cd2b2) ]
coclass UserEventTimerCallback {
	interface IUserEventTimerCallback;
}

// CLSID_UserEventTimer
[ uuid(864A1288-354C-4D19-9D68-C2742BB14997) ]
coclass UserEventTimer {
	interface IUserEventTimer;
}

// CLSID_NetCrawler
[ uuid(601ac3dc-786a-4eb0-bf40-ee3521e70bfb) ] 
coclass NetCrawler {
	interface INetCrawler;
}

// CLSID_CDBurn
[ uuid(fbeb8a05-beee-4442-804e-409d6c4515e9) ]
coclass CDBurn {
	interface ICDBurn;
}

// CLSID_AddressBarParser
[ uuid(e0e11a09-5cb8-4b6c-8332-e00720a168f2) ]
coclass AddressBarParser {
	interface IAddressBarParser;
}

// CLSID_TaskbarList
[ uuid(56FDF344-FD6D-11d0-958A-006097C9A090) ]
coclass TaskbarList {
	interface ITaskbarList;
}

// For supporting HTML wizard extensions we provide this object, it implements the IWizardExtenion
// and allows the site to specific via an IPropertyBag an URL which should be displayed.  The property
// bag is requested from the SID_WebWizardHost, and it used inturn to return parameter information
// back to the site (eg. any information that the displayed HTML would like to communicate).
    
// CLSID_WebWizardHost
[ uuid(c827f149-55c1-4d28-935e-57e47caed973) ]
coclass WebWizardHost {
	interface IWebWizardExtension;
}

// CLSID_PublishDropTarget
[ uuid(CC6EEFFB-43F6-46c5-9619-51D571967F7D) ]
coclass PublishDropTarget {
	interface stdole.IUnknown;
}

// CLSID_PublishingWizard
[ uuid(6b33163c-76a5-4b6c-bf21-45de9cd503a1) ] 
coclass PublishingWizard {
	interface IPublishingWizard;
}

// CLSID_InternetPrintOrdering
[ uuid(add36aa8-751a-4579-a266-d66f5202ccbb) ]
coclass InternetPrintOrdering {
	interface stdole.IUnknown;
}

// CLSID_FolderViewHost
[ uuid(20b1cb23-6968-4eb9-b7d4-a66d00d07cee) ]
coclass FolderViewHost {
	interface IFolderViewHost;
}

// CLSID_ProgressDialog
[ uuid(F8383852-FCD3-11d1-A6B9-006097DF5BD4) ]
coclass ProgressDialog {
	interface IProgressDialog;
}
