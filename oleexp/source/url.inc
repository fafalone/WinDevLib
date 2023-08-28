//tbShellLib - Interfaces done; APIs not

typedef [helpstring("Flags for IUniformResourceLocator::SetURL()")]
enum IURL_SETURL_FLAGS {
    [helpstring("Guess protocol if missing")]
    IURL_SETURL_FL_GUESS_PROTOCOL = 0x0001,
    [helpstring("Use default protocol if missing")]
    IURL_SETURL_FL_USE_DEFAULT_PROTOCOL = 0x0002,
} IURL_SETURL_FLAGS;

typedef [helpstring("Flags for IUniformResourceLocator()::InvokeCommand()")]
enum IURL_INVOKECOMMAND_FLAGS {
    IURL_INVOKECOMMAND_FL_ALLOW_UI = 0x0001,
    [helpstring("Ignore pcszVerb")]
    IURL_INVOKECOMMAND_FL_USE_DEFAULT_VERB = 0x0002,
    IURL_INVOKECOMMAND_FL_DDEWAIT           = 0x0004,
} IURL_INVOKECOMMAND_FLAGS;

typedef [helpstring("Command info for IUniformResourceLocator::InvokeCommand()")]
struct URLINVOKECOMMANDINFO {
    [helpstring("Size of structure")]
    LONG  dwcbSize;
    [helpstring("Bit field of IURL_INVOKECOMMAND_FLAGS")]
    IURL_INVOKECOMMAND_FLAGS dwFlags;
    [helpstring("Parent window.  Valid only if IURL_INVOKECOMMAND_FL_ALLOW_UI is set.")]
    LONG   hwndParent;
    [helpstring("Verb to invoke.  Ignored if IURL_INVOKECOMMAND_FL_USE_DEFAULT_VERB is set.")]
    LONG pcszVerb;
} URLINVOKECOMMANDINFO;

[
    odl,
    uuid(FBF23B80-E3F0-101B-8488-00AA003E56F8),
    helpstring("IuniformResourceLocator - ANSI version")
]
interface IUniformResourceLocatorA: IUnknown {

    [helpstring("Sets an object's URL")]
    HRESULT SetURL(
        [in] LPSTR pcszURL,
        [in] IURL_SETURL_FLAGS dwInFlags);

    [helpstring("Gets an object's URL")]
    HRESULT GetURL(
        [out, retval] LONG *ppszURL);

    [helpstring("Invokes a command on an object's URL.")]
    HRESULT InvokeCommand(
        [in] URLINVOKECOMMANDINFO *purlici);
};

[
    odl,
    uuid(CABB0DA0-DA57-11CF-9974-0020AFD79762),
    helpstring("IUniformResourceLocator - Unicode version")
]
interface IUniformResourceLocatorW: IUnknown {

    [helpstring("Sets an object's URL")]
    HRESULT SetURL(
        [in] LPWSTR pcszURL,
        [in] IURL_SETURL_FLAGS dwInFlags);

    [helpstring("Gets an object's URL")]
    HRESULT GetURL(
        [out, retval] LONG *ppszURL);

    [helpstring("Invokes a command on an object's URL.")]
    HRESULT InvokeCommand(
        [in] URLINVOKECOMMANDINFO *purlici);
};

typedef [helpstring("PROPIDs for Internet Shortcuts (FMTID_Intshcut) to be used with IPropertySetStorage/IPropertyStorage")]
enum PID_IS {
    [helpstring("URL")]
    PID_IS_URL = 2,
    [helpstring("Name of the internet shortcut")]
    PID_IS_NAME = 4,
    [helpstring("Working directory for the shortcut")]
    PID_IS_WORKINGDIR = 5,
    [helpstring("Hotkey for the shortcut")]
    PID_IS_HOTKEY = 6,
    [helpstring("Show command for shortcut")]
    PID_IS_SHOWCMD = 7,
    [helpstring("Index into file that has icon")]
    PID_IS_ICONINDEX = 8,
    [helpstring("File that has the icon")]
    PID_IS_ICONFILE = 9,
    [helpstring("What's New text")]
    PID_IS_WHATSNEW = 10,
    [helpstring("Author")]
    PID_IS_AUTHOR = 11,
    [helpstring("Description text of site")]
    PID_IS_DESCRIPTION = 12,
    [helpstring("User annotated comment")]
    PID_IS_COMMENT = 13,
    [helpstring("True when shortcut is roamed for first time")]
    PID_IS_ROAMED = 14,
} PID_IS;

typedef [helpstring("PROPIDs for Internet Sites (FMTID_InternetSite) to be used with IPropertySetStorage/IPropertyStorage")]
enum PID_INTSITE {
    [helpstring("[VT_LPWSTR] What's New text")]
    PID_INTSITE_WHATSNEW = 2,
    [helpstring("[VT_LPWSTR] Author")]
    PID_INTSITE_AUTHOR = 3,
    [helpstring("[VT_CURRENCY] Time site was last visited")]
    PID_INTSITE_LASTVISIT = 4,
    [helpstring("[VT_CURRENCY] Time site was last modified")]
    PID_INTSITE_LASTMOD = 5,
    [helpstring("[VT_UI4] Number of times user has visited")]
    PID_INTSITE_VISITCOUNT = 6,
    [helpstring("[VT_LPWSTR] Description text of site")]
    PID_INTSITE_DESCRIPTION = 7,
    [helpstring("[VT_LPWSTR] User annotated comment")]
    PID_INTSITE_COMMENT = 8,
    [helpstring("[VT_UI4] PIDISM_ flags")]
    PID_INTSITE_FLAGS = 9,
    PID_INTSITE_CONTENTLEN = 10,
    PID_INTSITE_CONTENTCODE = 11,
    [helpstring("[VT_UI4] Levels to recurse (0-3)")]
    PID_INTSITE_RECURSE = 12,
    PID_INTSITE_WATCH = 13,
    [helpstring("[VT_UI8] Subscription cookie")]
    PID_INTSITE_SUBSCRIPTION = 14,
    [helpstring("[VT_LPWSTR] URL")]
    PID_INTSITE_URL = 15,
    [helpstring("[VT_LPWSTR] Title")]
    PID_INTSITE_TITLE = 16,
    [helpstring("[VT_UI4] Codepage of the document")]
    PID_INTSITE_CODEPAGE = 18,
    [helpstring("[VT_UI4] Tracking")]
    PID_INTSITE_TRACKING = 19,
     [helpstring("Retrieve the index to the icon")]
     PID_INTSITE_ICONINDEX    = 20,
     [helpstring("Retrieve the file containing the icon index.")]
     PID_INTSITE_ICONFILE     = 21,
     [helpstring("Indicates that this entry was roamed from a different machine")]
     PID_INTSITE_ROAMED       = 34,
   

} PID_INTSITE;

typedef [helpstring("Values for PID_INTSITE_WATCH")]
enum PIDISM {
    [helpstring("Monitor based on global setting")]
    PIDISM_GLOBAL = 0,
    [helpstring("User says watch")]
    PIDISM_WATCH = 1,
    [helpstring("User says don't watch")]
    PIDISM_DONTWATCH = 2
} PIDISM;

[
    uuid(FBF23B40-E3F0-101B-8488-00AA003E56F8)
]
coclass InternetShortcut {
    [default] interface IUniformResourceLocatorW;
    interface IUniformResourceLocatorA;
    interface IPersistFile;
    interface IPropertySetStorage;
    interface IObjectWithSite;
}

typedef enum TRANSLATEURL_IN_FLAGS {
    [helpstring("Guess protocol if missing")]
    TRANSLATEURL_FL_GUESS_PROTOCOL = 0x0001,
    [helpstring("Use default protocol if missing")]
    TRANSLATEURL_FL_USE_DEFAULT_PROTOCOL = 0x0002,
} TRANSLATEURL_IN_FLAGS;

typedef enum URLASSOCIATIONDIALOG_IN_FLAGS {
    URLASSOCDLG_FL_USE_DEFAULT_NAME = 0x0001,
    URLASSOCDLG_FL_REGISTER_ASSOC = 0x0002,
} URLASSOCIATIONDIALOG_IN_FLAGS;

typedef enum MIMEASSOCIATIONDIALOG_IN_FLAGS {
    MIMEASSOCDLG_FL_REGISTER_ASSOC = 0x0001,
} MIMEASSOCIATIONDIALOG_IN_FLAGS;

[
    dllname("URL.DLL")
]
module URLDll {

    //const LONG MAX_PATH = 260;
    const LONG INTERNET_MAX_PATH_LENGTH = 2048;
    const LONG INTERNET_MAX_SCHEME_LENGTH = 32;
    const LONG INTERNET_MAX_URL_LENGTH = 2083;
    const LONG FILE_ALL_ACCESS = 0x1F01FF;

    [entry("TranslateURLA")]
    HRESULT TranslateURL(
        [in] LPSTR pcszURL,
        [in] TRANSLATEURL_IN_FLAGS dwInFlags,
        [out, retval] LONG *ppszTranslatedURL);

    [entry("TranslateURLW")]
    HRESULT TranslateURLW(
        [in] LPWSTR pcszURL,
        [in] TRANSLATEURL_IN_FLAGS dwInFlags,
        [out, retval] LONG *ppszTranslatedURL);

    [entry("URLAssociationDialogA")]
    HRESULT URLAssociationDialog(
        [in] LONG hwndParent,
        [in] URLASSOCIATIONDIALOG_IN_FLAGS dwInFlags,
        [in] LPSTR pcszFile,
        [in] LPSTR pcszURL,
        [in] LPSTR pszAppBuf,
        [in] LONG ucAppBufLen);

    [entry("URLAssociationDialogW")]
    HRESULT URLAssociationDialogW(
        [in] LONG hwndParent,
        [in] URLASSOCIATIONDIALOG_IN_FLAGS dwInFlags,
        [in] LPWSTR pcszFile,
        [in] LPWSTR pcszURL,
        [in] LPWSTR pszAppBuf,
        [in] LONG ucAppBufLen);

    [entry("MIMEAssociationDialogA")]
    HRESULT MIMEAssociationDialog(
        [in] LONG hwndParent,
        [in] URLASSOCIATIONDIALOG_IN_FLAGS dwInFlags,
        [in] LPSTR pcszFile,
        [in] LPSTR pcszMIMEContentType,
        [in] LPSTR pszAppBuf,
        [in] LONG ucAppBufLen);

    [entry("MIMEAssociationDialogW")]
    HRESULT MIMEAssociationDialogW(
        [in] LONG hwndParent,
        [in] URLASSOCIATIONDIALOG_IN_FLAGS dwInFlags,
        [in] LPWSTR pcszFile,
        [in] LPWSTR pcszMIMEContentType,
        [in] LPWSTR pszAppBuf,
        [in] LONG ucAppBufLen);

    [entry("InetIsOffline")] 
    LONG InetIsOffline([in] LONG dwFlags);
    
}