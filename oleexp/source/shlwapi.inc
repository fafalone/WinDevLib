typedef enum StrToIntExFlags {
    STIF_DEFAULT = 0x00000000,
    STIF_SUPPORT_HEX = 0x00000001
} StrToIntExFlags;

// Return flags for PathGetCharType
typedef enum GCTFlags {
    GCT_INVALID = 0x0000,
    GCT_LFNCHAR = 0x0001,
    GCT_SHORTCHAR = 0x0002,
    GCT_WILD = 0x0004,
    GCT_SEPARATOR = 0x0008
} GCTFlags;

typedef enum {
    URL_SCHEME_INVALID = -1,
    URL_SCHEME_UNKNOWN = 0,
    URL_SCHEME_FTP,
    URL_SCHEME_HTTP,
    URL_SCHEME_GOPHER,
    URL_SCHEME_MAILTO,
    URL_SCHEME_NEWS,
    URL_SCHEME_NNTP,
    URL_SCHEME_TELNET,
    URL_SCHEME_WAIS,
    URL_SCHEME_FILE,
    URL_SCHEME_MK,
    URL_SCHEME_HTTPS,
    URL_SCHEME_SHELL,
    URL_SCHEME_SNEWS,
    URL_SCHEME_LOCAL,
    URL_SCHEME_JAVASCRIPT,
    URL_SCHEME_VBSCRIPT,
    URL_SCHEME_ABOUT,
    URL_SCHEME_RES,
    URL_SCHEME_MAXVALUE
} URL_SCHEME;

typedef enum {
    URL_PART_NONE = 0,
    URL_PART_SCHEME = 1,
    URL_PART_HOSTNAME,
    URL_PART_USERNAME,
    URL_PART_PASSshort,
    URL_PART_PORT,
    URL_PART_QUERY,
} URL_PART;

typedef enum {
    URLIS_URL,
    URLIS_OPAQUE,
    URLIS_NOHISTORY,
    URLIS_FILEURL,
    URLIS_APPLIABLE,
    URLIS_DIRECTORY,
    URLIS_HASQUERY,
} URLIS;

typedef enum {
    URL_UNESCAPE = 0x10000000,
    URL_ESCAPE_UNSAFE = 0x20000000,
    URL_PLUGGABLE_PROTOCOL = 0x40000000,
    URL_WININET_COMPATIBILITY = 0x80000000,
    URL_DONT_ESCAPE_EXTRA_INFO = 0x02000000,
    URL_DONT_UNESCAPE_EXTRA_INFO = 0x02000000,
    URL_BROWSER_MODE = 0x02000000,
    URL_ESCAPE_SPACES_ONLY = 0x04000000,
    URL_DONT_SIMPLIFY = 0x08000000,
    URL_NO_META = 0x08000000,
    URL_UNESCAPE_INPLACE = 0x00100000,
    URL_CONVERT_IF_DOSPATH = 0x00200000,
    URL_UNESCAPE_HIGH_ANSI_ONLY = 0x00400000,
    URL_INTERNAL_PATH = 0x00800000, // Will escape #'s in paths
    URL_FILE_USE_PATHURL = 0x00010000,
    URL_ESCAPE_PERCENT = 0x00001000,
    URL_ESCAPE_SEGMENT_ONLY = 0x00002000, // Treat the entire URL param as one URL segment
} URL;

typedef enum {
    URL_PARTFLAG_KEEPSCHEME = 0x00000001
} URL_PARTFLAG;

typedef enum {
    URL_APPLY_DEFAULT = 0x00000001,
    URL_APPLY_GUESSSCHEME = 0x00000002,
    URL_APPLY_GUESSFILE = 0x00000004,
    URL_APPLY_FORCEAPPLY = 0x00000008,
} URL_APPLY;

typedef enum {
    SHREGDEL_DEFAULT = 0x00000000, // Delete's HKCU, or HKLM if HKCU is not found.
    SHREGDEL_HKCU = 0x00000001, // Delete HKCU only
    SHREGDEL_HKLM = 0x00000010, // Delete HKLM only.
    SHREGDEL_BOTH = 0x00000011, // Delete both HKCU and HKLM.
} SHREGDEL_FLAGS;

typedef enum {
   SHREGENUM_DEFAULT = 0x00000000, // Enumerates HKCU or HKLM if not found.
   SHREGENUM_HKCU = 0x00000001, // Enumerates HKCU only
   SHREGENUM_HKLM = 0x00000010, // Enumerates HKLM only.
   SHREGENUM_BOTH = 0x00000011, // Enumerates both HKCU and HKLM without duplicates. This option is NYI.
} SHREGENUM_FLAGS;

typedef enum {
    SHREGSET_HKCU = 0x00000001, // Write to HKCU if empty.
    SHREGSET_FORCE_HKCU = 0x00000002, // Write to HKCU.
    SHREGSET_HKLM = 0x00000004, // Write to HKLM if empty.
    SHREGSET_FORCE_HKLM = 0x00000008, // Write to HKLM.
    SHREGSET_DEFAULT = 0x00000006 // Default is SHREGSET_FORCE_HKCU | SHREGSET_HKLM.
} SHREGSET;


typedef enum {
	KEY_QUERY_VALUE        = 0x0001,
	KEY_SET_VALUE          = 0x0002,
	KEY_CREATE_SUB_KEY     = 0x0004,
	KEY_ENUMERATE_SUB_KEYS = 0x0008,
	KEY_NOTIFY             = 0x0010,
	KEY_CREATE_LINK        = 0x0020,
	KEY_READ               = 0x00020019,
	KEY_WRITE              = 0x00020006,
	KEY_EXECUTE            = 0x00020019,
	KEY_ALL_ACCESS         = 0x000F003F
} REGSAM;

typedef enum {
    ASSOCF_INIT_NOREMAPCLSID = 0x00000001, // do not remap clsids to progids
    ASSOCF_INIT_BYEXENAME = 0x00000002, // executable is being passed in
    ASSOCF_OPEN_BYEXENAME = 0x00000002, // executable is being passed in
    ASSOCF_INIT_DEFAULTTOSTAR = 0x00000004, // treat "*" as the BaseClass
    ASSOCF_INIT_DEFAULTTOFOLDER = 0x00000008, // treat "Folder" as the BaseClass
    ASSOCF_NOUSERSETTINGS = 0x00000010, // dont use HKCU
    ASSOCF_NOTRUNCATE = 0x00000020, // dont truncate the return string
    ASSOCF_VERIFY = 0x00000040, // verify data is accurate (DISK HITS)
    ASSOCF_REMAPRUNDLL = 0x00000080, // actually gets info about rundlls target if applicable
    ASSOCF_NOFIXUPS = 0x00000100, // attempt to fix errors if found
    ASSOCF_IGNOREBASECLASS = 0x00000200,  // dont recurse into the baseclass
    ASSOCF_INIT_IGNOREUNKNOWN          = 0x00000400,  //  dont use the "Unknown" progid, instead fail
//#if (NTDDI_WIN8)
    ASSOCF_INIT_FIXED_PROGID           = 0x00000800,  //  the Init() pszAssoc value is a ProgId that should not be mapped using the current user defaults
    ASSOCF_IS_PROTOCOL                 = 0x00001000,  //  the Init() pszAssoc value is a protocol that should be mapped using the current user defaults
    ASSOCF_INIT_FOR_FILE               = 0x00002000  //  use this flag when specifying ASSOCF_INIT_FIXED_PROGID if the ProgId corresponds with a file extension based association
//#endif

} ASSOCF;

typedef enum {
    ASSOCSTR_COMMAND = 1, // shell\verb\command string
    ASSOCSTR_EXECUTABLE, // the executable part of command string
    ASSOCSTR_FRIENDLYDOCNAME, // friendly name of the document type
    ASSOCSTR_FRIENDLYAPPNAME, // friendly name of executable
    ASSOCSTR_NOOPEN, // noopen value
    ASSOCSTR_SHELLNEWVALUE, // query values under the shellnew key
    ASSOCSTR_DDECOMMAND, // template for DDE commands
    ASSOCSTR_DDEIFEXEC, // DDECOMMAND to use if just create a process
    ASSOCSTR_DDEAPPLICATION, // Application name in DDE broadcast
    ASSOCSTR_DDETOPIC, // Topic Name in DDE broadcast
    ASSOCSTR_INFOTIP, // info tip for an item, or list of properties to create info tip from
    ASSOCSTR_QUICKTIP, // same as ASSOCSTR_INFOTIP, except, this list contains only quickly retrievable properties
    ASSOCSTR_TILEINFO,          //  similar to ASSOCSTR_INFOTIP - lists important properties for tileview
    ASSOCSTR_CONTENTTYPE,       //  MIME Content type
    ASSOCSTR_DEFAULTICON,       //  Default icon source
    ASSOCSTR_SHELLEXTENSION,    //  Guid string pointing to the Shellex\Shellextensionhandler value.
    ASSOCSTR_DROPTARGET,        //  The CLSID of DropTarget
    ASSOCSTR_DELEGATEEXECUTE,   //  The CLSID of DelegateExecute
    ASSOCSTR_SUPPORTED_URI_PROTOCOLS, 
//#if (NTDDI_VERSION >= NTDDI_WIN10)
    ASSOCSTR_PROGID,            // The ProgId provided by the app associated with the file type or uri scheme based on user default settings.
    ASSOCSTR_APPID,             // The AppUserModelID of the app associated with the file type or uri scheme based on user default settings.
    ASSOCSTR_APPPUBLISHER,      // THe publisher of the app associated with the file type or uri scheme based on user default settings.
    ASSOCSTR_APPICONREFERENCE,  // The icon reference of the app associated with the file type or uri scheme based on user default settings.
//#end if
    ASSOCSTR_MAX // last item in enum...
} ASSOCSTR;

typedef enum {
    ASSOCKEY_SHELLEXECCLASS = 1, // the key that should be passed to ShellExec(hkeyClass)
    ASSOCKEY_APP, // the "Application" key for the association
    ASSOCKEY_CLASS, // the progid or class key
    ASSOCKEY_BASECLASS, // the BaseClass key
    ASSOCKEY_MAX // last item in enum...
} ASSOCKEY;

typedef enum {
    ASSOCDATA_MSIDESCRIPTOR = 1, // Component Descriptor to pass to MSI APIs
    ASSOCDATA_NOACTIVATEHANDLER, // restrict attempts to activate window
    ASSOCDATA_QUERYCLASSSTORE, // should check with the NT Class Store
    ASSOCDATA_HASPERUSERASSOC, // defaults to user specified association
    ASSOCDATA_EDITFLAGS,          //  Edit flags.
    ASSOCDATA_VALUE,              //  use pszExtra as the Value name
    ASSOCDATA_MAX
} ASSOCDATA;

typedef enum {
    ASSOCENUM_NONE
} ASSOCENUM;

typedef enum ASSOCCLASS
{                               //  which other members are used
    ASSOCCLASS_SHELL_KEY = 0,   //  hkeyClass
    ASSOCCLASS_PROGID_KEY,      //  hkeyClass
    ASSOCCLASS_PROGID_STR,      //  pszClass (HKCR\pszClass)
    ASSOCCLASS_CLSID_KEY,       //  hkeyClass
    ASSOCCLASS_CLSID_STR,       //  pszClass (HKCR\CLSID\pszClass)
    ASSOCCLASS_APP_KEY,         //  hkeyClass
    ASSOCCLASS_APP_STR,         //  pszClass (HKCR\Applications\PathFindFileName(pszClass))
    ASSOCCLASS_SYSTEM_STR,      //  pszClass
    ASSOCCLASS_FOLDER,          //  none
    ASSOCCLASS_STAR,            //  none
//#if (NTDDI_WIN8)
    ASSOCCLASS_FIXED_PROGID_STR,//  pszClass (HKCR\pszClass), do not apply mapping of pszClass based on user defaults
    ASSOCCLASS_PROTOCOL_STR,    //  pszClass is a protocol, apply mapping of pszClass based on user defaults
//#endif
} ASSOCCLASS;

typedef struct ASSOCIATIONELEMENT
{
    ASSOCCLASS ac;              // required
    HKEY hkClass;               // may be NULL
    LONG pszClass;            // may be NULL
} ASSOCIATIONELEMENT;

typedef enum {
	SHACF_DEFAULT =       0x00000000, // Currently (SHACF_FILESYSTEM | SHACF_URLALL)
	SHACF_FILESYSTEM =    0x00000001, // This includes the File System as well as the rest of the shell (Desktop\My Computer\Control Panel\)
	SHACF_URLHISTORY =    0x00000002, // URLs in the User's History
	SHACF_URLMRU =        0x00000004, // URLs in the User's Recently Used list.
	SHACF_URLALL =        0x00000006, // (SHACF_URLHISTORY | SHACF_URLMRU)
	SHACF_USETAB =        0x00000008, // Use the tab to move thru the autocomplete possibilities instead of to the next dialog/window control.
	SHACF_FILESYS_ONLY =  0x00000010, // Don't AutoComplete non-File System items.
    SHACF_FILESYS_DIRS  = 0x00000020, // Same as SHACF_FILESYS_ONLY except it only includes directories, UNC servers, and UNC server shares.
    SHACF_VIRTUAL_NAMESPACE = 0x00000040, // Also include the virtual namespace
    SHACF_AUTOSUGGEST_FORCE_ON = 0x10000000, // Ignore the registry default and force the feature on.
	SHACF_AUTOSUGGEST_FORCE_OFF = 0x20000000, // Ignore the registry default and force the feature off.
	SHACF_AUTOAPPEND_FORCE_ON =   0x40000000, // Ignore the registry default and force the feature on. (Also know as AutoComplete)
	SHACF_AUTOAPPEND_FORCE_OFF =  0x80000000, // Ignore the registry default and force the feature off. (Also know as AutoComplete)
} SHACF;

typedef enum {
	CTF_INSIST             = 0x00000001, // SHCreateThread() dwFlags - call pfnThreadProc synchronously if CreateThread() fails
	CTF_THREAD_REF         = 0x00000002, // hold a reference to the creating thread
	CTF_PROCESS_REF        = 0x00000004, // hold a reference to the creating process
	CTF_COINIT             = 0x00000008, // init COM for the created thread
	CTF_FREELIBANDEXIT     = 0x00000010, // hold a ref to the DLL and call FreeLibraryAndExitThread() when done
	CTF_REF_COUNTED        = 0x00000020, //Set thread reference
	CTF_WAIT_ALLOWCOM      = 0x00000040, //The calling thread waits and pumps COM and SendMessage messages.
	CTF_UNUSED             = 0x00000080, //Used. Jk.
	CTF_INHERITWOW64       = 0x00000100, //Inherits WoW64 disabled state (if set)
	CTF_WAIT_NO_REENTRANCY = 0x00000200, //SendMessage to a window hosted on the calling thread causes the synchronous proc to deadlock.
	CTF_KEYBOARD_LOCALE    = 0x00000400, //Inherit keyboard locale from calling thread
	CTF_OLEINITIALIZE      = 0x00000800, //Initialize COM with STA model for created thread.
	CTF_COINIT_MTA         = 0x00001000, //Initialize COM with MTA model for created thread.
	CTF_NOADDREFLIB        = 0x00002000, //Opposite of CTF_FREELIBANDEXIT.
} CTF;

typedef struct DLLVERSIONINFO {
    LONG cbSize;
    LONG dwMajorVersion; // Major version
    LONG dwMinorVersion; // Minor version
    LONG dwBuildNumber; // Build number
    LONG dwPlatformID; // DLLVER_PLATFORM_*
} DLLVERSIONINFO;

// Platform IDs for DLLVERSIONINFO
typedef enum {
	DLLVER_PLATFORM_WINDOWS = 0x00000001, // Windows 95
	DLLVER_PLATFORM_NT = 0x00000002, // Windows NT
} DLLVER;

typedef struct {
    DLLVERSIONINFO info1;
    LONG dwFlags; // No flags currently defined
    CURRENCY ullVersion; // Encoded as:
    // Major 0xFFFF 0000 0000 0000
    // Minor 0x0000 FFFF 0000 0000
    // Build 0x0000 0000 FFFF 0000
    // QFE 0x0000 0000 0000 FFFF
} DLLVERSIONINFO2;

[
    dllname("SHLWAPI.DLL")
]
module shlwapi {

//oleexp additions

	[entry("SHCreateStreamOnFileW")]
	int SHCreateStreamOnFileW(
		[in] LPWSTR pszFile,
		[in] STGM grfMode,
		[out] IStream **ppstm);

    [entry("SHCreateStreamOnFileEx")]
	int SHCreateStreamOnFileEx(
		[in] LONG pszFile,
		[in] STGM grfMode,
		[in] FILE_ATTRIBUTES dwAttributes,
		[in] BOOL fCreate,
		[in] IStream *pstmTemplate,
		[out] IStream **ppstm);


    [entry("SHCreateMemStream")]
    IStream* SHCreateMemStream([in] LPVOID pInit, [in] ULONG cbInit);

	[entry("IUnknown_QueryService")]
	int IUnknown_QueryService(
		[in] IUnknown *punk,
		[in] UUID *guidService,
		[in] UUID *riid,
		[out] LPVOID ppv);
		
	
//
//=============== String Routines ===================================
//
    [entry("StrChrA")]
    LONG StrChrA(
        [in] LPSTR lpStart,
        [in] short wMatch);

    [entry("StrChr")]
    LONG StrChr(
        [in] LPWSTR lpStart,
        [in] short wMatch);

    [entry("StrChrI")]
    LONG StrChrIA(
        [in] LPSTR lpStart,
        [in] short wMatch);

    [entry("StrChrI")]
    LONG StrChrIW(
        [in] LPWSTR lpStart,
        [in] short wMatch);

    [entry("StrCmpNA")]
    LONG StrCmpNA(
        [in] LPSTR lpStr1,
        [in] LPSTR lpStr2,
        [in] int nChar);

    [entry("StrCmpN")]
    LONG StrCmpN(
        [in] LPWSTR lpStr1,
        [in] LPWSTR lpStr2,
        [in] int nChar);

    [entry("StrCmpNIA")]
    LONG StrCmpNIA(
        [in] LPSTR lpStr1,
        [in] LPSTR lpStr2,
        [in] int nChar);

    [entry("StrCmpNIW")]
    LONG StrCmpNI(
        [in] LPWSTR lpStr1,
        [in] LPWSTR lpStr2,
        [in] int nChar);

    [entry("StrCSpnA")]
    LONG StrCSpnA(
        [in] LPSTR lpStr,
        [in] LPSTR lpSet);

    [entry("StrCSpnW")]
    LONG StrCSpn(
        [in] LPWSTR lpStr,
        [in] LPWSTR lpSet);

    [entry("StrCSpnIA")]
    LONG StrCSpnIA(
        [in] LPSTR lpStr,
        [in] LPSTR lpSet);

    [entry("StrCSpnIW")]
    LONG StrCSpnI(
        [in] LPWSTR lpStr,
        [in] LPWSTR lpSet);

    [entry("StrDupA")]
    LONG StrDupA(
        [in] LPSTR lpSrch);

    [entry("StrDupW")]
    LONG StrDup(
        [in] LPWSTR lpSrch);

    [entry("StrFormatByteSizeA")]
    LONG StrFormatByteSizeA(
        [in] LONG dw,
        [in] LPSTR szBuf,
        [in] LONG uiBufSize);

    [entry("StrFormatByteSize64A")]
    LONG StrFormatByteSize64A(
        [in] CURRENCY qdw,
        [in] LPSTR szBuf,
        [in] LONG uiBufSize);

    [entry("StrFormatByteSizeW")]
    LONG StrFormatByteSize(
        [in] CURRENCY qdw,
        [in] LPWSTR szBuf,
        [in] LONG uiBufSize);

    [entry("StrformatKBSizeW")]
    LONG StrFormatKBSize(
        [in] CURRENCY qdw,
        [in] LPWSTR szBuf,
        [in] LONG uiBufSize);

    [entry("StrFormatKBSizeA")]
    LONG StrFormatKBSizeA(
        [in] CURRENCY qdw,
        [in] LPSTR szBuf,
        [in] LONG uiBufSize);

    [entry("StrFromTimeIntervalA")]
    LONG StrFromTimeIntervalA(
        [in] LPSTR pszOut,
        [in] LONG cchMax,
        [in] LONG dwTimeMS,
        [in] int digits);

    [entry("StrFromTimeIntervalW")]
    LONG StrFromTimeInterval(
        [in] LPWSTR pszOut,
        [in] LONG cchMax,
        [in] LONG dwTimeMS,
        [in] int digits);

    [entry("StrIsIntlEqualA")]
    BOOL StrIsIntlEqualA(
        [in] BOOL fCaseSens,
        [in] LPSTR lpString1,
        [in] LPSTR lpString2,
        [in] int nChar);

    [entry("StrIsIntlEqualW")]
    BOOL StrIsIntlEqual(
        [in] BOOL fCaseSens,
        [in] LPWSTR lpString1,
        [in] LPWSTR lpString2,
        [in] int nChar);

    [entry("StrNCatA")]
    LONG StrNCatA(
        [in] LPSTR psz1,
        [in] LPSTR psz2,
        [in] int cchMax);

    [entry("StrNCatW")]
    LONG StrNCat(
        [in] LPWSTR psz1,
        [in] LPWSTR psz2,
        [in] int cchMax);

    [entry("StrPBrkA")]
    LONG StrPBrkA(
        [in] LPSTR psz,
        [in] LPSTR pszSet);

    [entry("StrPBrkW")]
    LONG StrPBrk(
        [in] LPWSTR psz,
        [in] LPWSTR pszSet);

    [entry("StrRChrA")]
    LONG StrRChrA(
        [in] LPSTR lpStart,
        [in] LPSTR lpEnd,
        [in] short wMatch);

    [entry("StrRChrW")]
    LONG StrRChr(
        [in] LPWSTR lpStart,
        [in] LPWSTR lpEnd,
        [in] short wMatch);

    [entry("StrRChrIA")]
    LONG StrRChrIA(
        [in] LPSTR lpStart,
        [in] LPSTR lpEnd,
        [in] short wMatch);

    [entry("StrRChrIW")]
    LONG StrRChrI(
        [in] LPWSTR lpStart,
        [in] LPWSTR lpEnd,
        [in] short wMatch);

    [entry("StrRStrIA")]
    LONG StrRStrIA(
        [in] LPSTR lpSource,
        [in] LPSTR lpLast,
        [in] LPSTR lpSrch);

    [entry("StrRStrIW")]
    LONG StrRStrI(
        [in] LPWSTR lpSource,
        [in] LPWSTR lpLast,
        [in] LPWSTR lpSrch);

    [entry("StrSpnA")]
    LONG StrSpnA(
        [in] LPSTR psz,
        [in] LPSTR pszSet);

    [entry("StrSpnW")]
    LONG StrSpn(
        [in] LPWSTR psz,
        [in] LPWSTR pszSet);

    [entry("StrStrA")]
    LONG StrStrA(
        [in] LPSTR lpFirst,
        [in] LPSTR lpSrch);

    [entry("StrStr")]
    LONG StrStr(
        [in] LPWSTR lpFirst,
        [in] LPWSTR lpSrch);

    [entry("StrStrIA")]
    LONG StrStrIA(
        [in] LPSTR lpFirst,
        [in] LPSTR lpSrch);

    [entry("StrStrIW")]
    LONG StrStrI(
        [in] LPWSTR lpFirst,
        [in] LPWSTR lpSrch);

    [entry("StrToIntA")]
    LONG StrToIntA(
        [in] LPSTR lpSrc);

    [entry("StrToIntW")]
    LONG StrToInt(
        [in] LPWSTR lpSrc);

    [entry("StrToIntExA")]
    BOOL StrToIntExA(
        [in] LPSTR pszString,
        [in] StrToIntExFlags dwFlags,
        [in, out] int *piRet);

    [entry("StrToIntExW")]
    BOOL StrToIntEx(
        [in] LPWSTR pszString,
        [in] StrToIntExFlags dwFlags,
        [in, out] int *piRet);

    [entry("StrTrimA")]
    BOOL StrTrimA(
        [in] LPSTR psz,
        [in] LPSTR pszTrimChars);

    [entry("StrTrimW")]
    BOOL StrTrim(
        [in] LPWSTR psz,
        [in] LPWSTR pszTrimChars);

    [entry("StrCatW")]
    LONG StrCat(
        [in] LPWSTR psz1,
        [in] LPWSTR psz2);

    [entry("StrCmpW")]
    LONG StrCmp(
        [in] LPWSTR psz1,
        [in] LPWSTR psz2);

    [entry("StrCmpIW")]
    LONG StrCmpI(
        [in] LPWSTR psz1,
        [in] LPWSTR psz2);

    [entry("StrCpyW")]
    LONG StrCpy(
        [in] LPWSTR psz1,
        [in] LPWSTR psz2);

    [entry("StrCpyNW")]
    LONG StrCpyN(
        [in] LPWSTR psz1,
        [in] LPWSTR psz2,
        [in] int cchMax);

    [entry("StrCatBuffW")]
    LONG StrCatBuff(
        [in] LPWSTR pszDest,
        [in] LPWSTR pszSrc,
        [in] int cchDestBuffSize);

    [entry("StrCatBuffA")]
    LONG StrCatBuffA(
        [in] LPSTR pszDest,
        [in] LPSTR pszSrc,
        [in] int cchDestBuffSize);

    [entry("ChrCmpIA")]
    BOOL ChrCmpIA(
        [in] short w1,
        [in] short w2);

    [entry("ChrCmpIW")]
    BOOL ChrCmpI(
        [in] short w1,
        [in] short w2);

    [entry("wvnsprintfA")]
    LONG wvnsprintfA(
        [in] LPSTR lpOut,
        [in] int cchLimitIn,
        [in] LPSTR lpFmt,
        [in] void *arglist);

    [entry("wvnsprintfW")]
    LONG wvnsprintf(
        [in] LPWSTR lpOut,
        [in] int cchLimitIn,
        [in] LPWSTR lpFmt,
        [in] void *arglist);

    [entry("StrRetToStrA")]
    HRESULT StrRetToStrA(
        [in] LONG pstrret,
        [in] LONG pidl,
        [out, retval] LONG *ppsz);

    [entry("StrRetToStrW")]
    HRESULT StrRetToStr(
        [in] LONG pstrret,
        [in] LONG pidl,
        [out, retval] LONG *ppsz);

    [entry("StrRetToBufA")]
    HRESULT StrRetToBufA(
        [in] LONG pstrret,
        [in] LONG pidl,
        [in] LPSTR pszBuf,
        [in] LONG cchBuf);

    [entry("StrRetToBufW")]
    HRESULT StrRetToBuf(
        [in] LONG pstr,
        [in] LONG pidl,
        [in] LPWSTR pszBuf,
        [in] LONG cchBuf);

// helper to duplicate a string using the task allocator

    [entry("SHStrDupA")]
    HRESULT SHStrDupA(
        [in] LPSTR psz,
        [out, retval] LONG *ppwsz);

    [entry("SHStrDupW")]
    HRESULT SHStrDup(
        [in] LPWSTR psz,
        [out, retval] LONG *ppwsz);

// Backward compatible to NT's non-standard naming ( strictly
// for comctl32)
//

    [entry("IntlStrEqWorkerA")]
    BOOL IntlStrEqWorkerA(
        [in] BOOL fCaseSens,
        [in] LPSTR lpString1,
        [in] LPSTR lpString2,
        [in] int nChar);

    [entry("IntlStrEqWorkerW")]
    BOOL IntlStrEqWorker(
        [in] BOOL fCaseSens,
        [in] LPWSTR lpString1,
        [in] LPWSTR lpString2,
        [in] int nChar);

    const LPSTR SZ_CONTENTTYPE_HTML = "text/html";
    const LPSTR SZ_CONTENTTYPE_CDFA = "application/x-cdf";

//
//=============== Path Routines ===================================
//

    [entry("PathAddBackslashA")]
    LONG PathAddBackslashA(
        [in] LPSTR pszPath);

    [entry("PathAddBackslashW")]
    LONG PathAddBackslash(
        [in] LPWSTR pszPath);

    [entry("PathAddExtensionA")]
    BOOL PathAddExtensionA(
        [in] LPSTR pszPath,
        [in] LPSTR pszExt);

    [entry("PathAddExtensionW")]
    BOOL PathAddExtension(
        [in] LPWSTR pszPath,
        [in] LPWSTR pszExt);

    [entry("PathAppendA")]
    BOOL PathAppendA(
        [in] LPSTR pszPath,
        [in] LPSTR pMore);

    [entry("PathAppendW")]
    BOOL PathAppend(
        [in] LPWSTR pszPath,
        [in] LPWSTR pMore);

    [entry("PathBuildRootA")]
    LONG PathBuildRootA(
        [in] LPSTR pszRoot,
        [in] int iDrive);

    [entry("PathBuildRootW")]
    LONG PathBuildRoot(
        [in] LPWSTR pszRoot,
        [in] int iDrive);

    [entry("PathCanonicalizeA")]
    BOOL PathCanonicalizeA(
        [in] LPSTR pszBuf,
        [in] LPSTR pszPath);

    [entry("PathCanonicalizeW")]
    BOOL PathCanonicalize(
        [in] LPWSTR pszBuf,
        [in] LPWSTR pszPath);

    [entry("PathCombineA")]
    LONG PathCombineA(
        [in] LPSTR pszDest,
        [in] LPSTR pszDir,
        [in] LPSTR pszFile);

    [entry("PathCombineW")]
    LONG PathCombine(
        [in] LPWSTR pszDest,
        [in] LPWSTR pszDir,
        [in] LPWSTR pszFile);

    [entry("PathCompactPathA")]
    BOOL PathCompactPathA(
        [in] LONG hDC,
        [in] LPSTR pszPath,
        [in] LONG dx);

    [entry("PathCompactPathW")]
    BOOL PathCompactPath(
        [in] LONG hDC,
        [in] LPWSTR pszPath,
        [in] LONG dx);

    [entry("PathCompactPathExA")]
    BOOL PathCompactPathExA(
        [in] LPSTR pszOut,
        [in] LPSTR pszSrc,
        [in] LONG cchMax,
        [in] LONG dwFlags);

    [entry("PathCompactPathExW")]
    BOOL PathCompactPathEx(
        [in] LPWSTR pszOut,
        [in] LPWSTR pszSrc,
        [in] LONG cchMax,
        [in] LONG dwFlags);

    [entry("PathCommonPrefixA")]
    LONG PathCommonPrefixA(
        [in] LPSTR pszFile1,
        [in] LPSTR pszFile2,
        [in] LPSTR achPath);

    [entry("PathCommonPrefixW")]
    LONG PathCommonPrefix(
        [in] LPWSTR pszFile1,
        [in] LPWSTR pszFile2,
        [in] LPWSTR achPath);

    [entry("PathFileExistsA")]
    BOOL PathFileExistsA(
        [in] LPSTR pszPath);

    [entry("PathFileExistsW")]
    BOOL PathFileExists(
        [in] LPWSTR pszPath);

    [entry("PathFindExtensionA")]
    LONG PathFindExtensionA(
        [in] LPSTR pszPath);

    [entry("PathFindExtensionW")]
    LONG PathFindExtension(
        [in] LPWSTR pszPath);

    [entry("PathFindFileNameA")]
    LONG PathFindFileNameA(
        [in] LPSTR pszPath);

    [entry("PathFindFileNameW")]
    LONG PathFindFileName(
        [in] LPWSTR pszPath);

    [entry("PathFindNextComponentA")]
    LONG PathFindNextComponentA(
        [in] LPSTR pszPath);

    [entry("PathFindNextComponentW")]
    LONG PathFindNextComponent(
        [in] LPWSTR pszPath);

    [entry("PathFindOnPathA")]
    BOOL PathFindOnPathA(
        [in] LPSTR pszPath,
        [in] LPSTR *ppszOtherDirs);

    [entry("PathFindOnPathW")]
    BOOL PathFindOnPath(
        [in] LPWSTR pszPath,
        [in] LPWSTR *ppszOtherDirs);

    [entry("PathGetArgsA")]
    LONG PathGetArgsA(
        [in] LPSTR pszPath);

    [entry("PathGetArgsW")]
    LONG PathGetArgs(
        [in] LPWSTR pszPath);

    [entry("PathFindSuffixArrayA")]
    LONG PathFindSuffixArrayA(
        [in] LPSTR pszPath,
        [in] LPSTR *apszSuffix,
        [in] int iArraySize);

    [entry("PathFindSuffixArrayW")]
    LONG PathFindSuffixArray(
        [in] LPWSTR pszPath,
        [in] LPWSTR *apszSuffix,
        [in] int iArraySize);

    [entry("PathIsLFNFileSpecA")]
    BOOL PathIsLFNFileSpecA(
        [in] LPSTR lpName);

    [entry("PathIsLFNFileSpecW")]
    BOOL PathIsLFNFileSpec(
        [in] LPWSTR lpName);

    [entry("PathGetCharTypeA")]
    GCTFlags PathGetCharTypeA(
        [in] CHAR ch);

    [entry("PathGetCharTypeW")]
    GCTFlags PathGetCharType(
        [in] short ch);

    [entry("PathGetDriveNumberA")]
    LONG PathGetDriveNumberA(
        [in] LPSTR pszPath);

    [entry("PathGetDriveNumberW")]
    LONG PathGetDriveNumber(
        [in] LPWSTR pszPath);
    [entry("(")]
    BOOL PathIsDirectoryA(
        [in] LPSTR pszPath);

    [entry("PathIsDirectoryW")]
    BOOL PathIsDirectory(
        [in] LPWSTR pszPath);

    [entry("PathIsDirectoryEmptyA")]
    BOOL PathIsDirectoryEmptyA(
        [in] LPSTR pszPath);

    [entry("PathIsDirectoryEmptyW")]
    BOOL PathIsDirectoryEmpty(
        [in] LPWSTR pszPath);

    [entry("PathIsFileSpecA")]
    BOOL PathIsFileSpecA(
        [in] LPSTR pszPath);

    [entry("PathIsFileSpecW")]
    BOOL PathIsFileSpec(
        [in] LPWSTR pszPath);

    [entry("PathIsPrefixA")]
    BOOL PathIsPrefixA(
        [in] LPSTR pszPrefix,
        [in] LPSTR pszPath);

    [entry("PathIsPrefixW")]
    BOOL PathIsPrefix(
        [in] LPWSTR pszPrefix,
        [in] LPWSTR pszPath);

    [entry("PathIsRelativeA")]
    BOOL PathIsRelativeA(
        [in] LPSTR pszPath);

    [entry("PathIsRelativeW")]
    BOOL PathIsRelative(
        [in] LPWSTR pszPath);

    [entry("PathIsRootA")]
    BOOL PathIsRootA(
        [in] LPSTR pszPath);

    [entry("PathIsRootW")]
    BOOL PathIsRoot(
        [in] LPWSTR pszPath);

    [entry("PathIsSameRootA")]
    BOOL PathIsSameRootA(
        [in] LPSTR pszPath1,
        [in] LPSTR pszPath2);

    [entry("PathIsSameRootW")]
    BOOL PathIsSameRoot(
        [in] LPWSTR pszPath1,
        [in] LPWSTR pszPath2);

    [entry("PathIsUNCA")]
    BOOL PathIsUNCA(
        [in] LPSTR pszPath);

    [entry("PathIsUNCW")]
    BOOL PathIsUNC(
        [in] LPWSTR pszPath);

    [entry("PathIsNetworkPathA")]
    BOOL PathIsNetworkPathA(
        [in] LPSTR pszPath);

    [entry("PathIsNetworkPathW")]
    BOOL PathIsNetworkPath(
        [in] LPWSTR pszPath);

    [entry("PathIsUNCServerA")]
    BOOL PathIsUNCServerA(
        [in] LPSTR pszPath);

    [entry("PathIsUNCServerW")]
    BOOL PathIsUNCServer(
        [in] LPWSTR pszPath);

    [entry("PathIsUNCServerShareA")]
    BOOL PathIsUNCServerShareA(
        [in] LPSTR pszPath);

    [entry("PathIsUNCServerShareW")]
    BOOL PathIsUNCServerShare(
        [in] LPWSTR pszPath);

    [entry("PathIsContentTypeA")]
    BOOL PathIsContentTypeA(
        [in] LPSTR pszPath,
        [in] LPSTR pszContentType);

    [entry("PathIsContentTypeW")]
    BOOL PathIsContentType(
        [in] LPWSTR pszPath,
        [in] LPWSTR pszContentType);

    [entry("PathIsURLA")]
    BOOL PathIsURLA(
        [in] LPSTR pszPath);

    [entry("PathIsURLW")]
    BOOL PathIsURL(
        [in] LPWSTR pszPath);

    [entry("PathMakePrettyA")]
    BOOL PathMakePrettyA(
        [in] LPSTR pszPath);

    [entry("PathMakePrettyW")]
    BOOL PathMakePretty(
        [in] LPWSTR pszPath);

    [entry("PathMatchSpecA")]
    BOOL PathMatchSpecA(
        [in] LPSTR pszFile,
        [in] LPSTR pszSpec);

    [entry("PathMatchSpecW")]
    BOOL PathMatchSpec(
        [in] LPWSTR pszFile,
        [in] LPWSTR pszSpec);

    [entry("PathParseIconLocationA")]
    LONG PathParseIconLocationA(
        [in] LPSTR pszIconFile);

    [entry("PathParseIconLocationW")]
    LONG PathParseIconLocation(
        [in] LPWSTR pszIconFile);

    [entry("PathQuoteSpacesA")]
    void PathQuoteSpacesA(
        [in] LPSTR lpsz);

    [entry("PathQuoteSpacesW")]
    void PathQuoteSpaces(
        [in] LPWSTR lpsz);

    [entry("PathRelativePathToA")]
    BOOL PathRelativePathToA(
        [in] LPSTR pszPath,
        [in] LPSTR pszFrom,
        [in] LONG dwAttrFrom,
        [in] LPSTR pszTo,
        [in] LONG dwAttrTo);

    [entry("PathRelativePathToW")]
    BOOL PathRelativePathTo(
        [in] LPWSTR pszPath,
        [in] LPWSTR pszFrom,
        [in] LONG dwAttrFrom,
        [in] LPWSTR pszTo,
        [in] LONG dwAttrTo);

    [entry("PathRemoveArgsA")]
    void PathRemoveArgsA(
        [in] LPSTR pszPath);

    [entry("PathRemoveArgsW")]
    void PathRemoveArgs(
        [in] LPWSTR pszPath);

    [entry("PathRemoveBackslashA")]
    LONG PathRemoveBackslashA(
        [in] LPSTR pszPath);

    [entry("PathRemoveBackslashW")]
    LONG PathRemoveBackslash(
        [in] LPWSTR pszPath);

    [entry("PathRemoveBlanksA")]
    void PathRemoveBlanksA(
        [in] LPSTR pszPath);

    [entry("PathRemoveBlanksW")]
    void PathRemoveBlanks(
        [in] LPWSTR pszPath);

    [entry("PathRemoveExtensionA")]
    void PathRemoveExtensionA(
        [in] LPSTR pszPath);

    [entry("PathRemoveExtensionW")]
    void PathRemoveExtension(
        [in] LPWSTR pszPath);

    [entry("PathRemoveFileSpecA")]
    BOOL PathRemoveFileSpecA(
        [in] LPSTR pszPath);

    [entry("PathRemoveFileSpecW")]
    BOOL PathRemoveFileSpec(
        [in] LPWSTR pszPath);

    [entry("PathRenameExtensionA")]
    BOOL PathRenameExtensionA(
        [in] LPSTR pszPath,
        [in] LPSTR pszExt);

    [entry("PathRenameExtensionW")]
    BOOL PathRenameExtension(
        [in] LPWSTR pszPath,
        [in] LPWSTR pszExt);

    [entry("PathSearchAndQualifyA")]
    BOOL PathSearchAndQualifyA(
        [in] LPSTR pszPath,
        [in] LPSTR pszBuf,
        [in] LONG cchBuf);

    [entry("PathSearchAndQualifyW")]
    BOOL PathSearchAndQualify(
        [in] LPWSTR pszPath,
        [in] LPWSTR pszBuf,
        [in] LONG cchBuf);

    [entry("PathSetDlgItemPathA")]
    void PathSetDlgItemPathA(
        [in] LONG hDlg,
        [in] int id,
        [in] LPSTR pszPath);

    [entry("PathSetDlgItemPathW")]
    void PathSetDlgItemPath(
        [in] LONG hDlg,
        [in] int id,
        [in] LPWSTR pszPath);

    [entry("PathSkipRootA")]
    LONG PathSkipRootA(
        [in] LPSTR pszPath);

    [entry("PathSkipRootW")]
    LONG PathSkipRoot(
        [in] LPWSTR pszPath);

    [entry("PathStripPathA")]
    void PathStripPathA(
        [in] LPSTR pszPath);

    [entry("PathStripPathW")]
    void PathStripPath(
        [in] LPWSTR pszPath);

    [entry("PathStripToRootA")]
    BOOL PathStripToRootA(
        [in] LPSTR pszPath);

    [entry("PathStripToRootW")]
    BOOL PathStripToRoot(
        [in] LPWSTR pszPath);

    [entry("PathUnquoteSpacesA")]
    void PathUnquoteSpacesA(
        [in] LPSTR lpsz);

    [entry("PathUnquoteSpacesW")]
    void PathUnquoteSpaces(
        [in] LPWSTR lpsz);

    [entry("PathMakeSystemFolderA")]
    BOOL PathMakeSystemFolderA(
        [in] LPSTR pszPath);

    [entry("PathMakeSystemFolderW")]
    BOOL PathMakeSystemFolder(
        [in] LPWSTR pszPath);

    [entry("PathUnmakeSystemFolderA")]
    BOOL PathUnmakeSystemFolderA(
        [in] LPSTR pszPath);

    [entry("PathUnmakeSystemFolderW")]
    BOOL PathUnmakeSystemFolder(
        [in] LPWSTR pszPath);

    [entry("PathIsSystemFolderA")]
    BOOL PathIsSystemFolderA(
        [in] LPSTR pszPath,
        [in] LONG dwAttrb);

    [entry("PathIsSystemFolderW")]
    BOOL PathIsSystemFolder(
        [in] LPWSTR pszPath,
        [in] LONG dwAttrb);

    [entry("PathUndecorateA")]
    void PathUndecorateA(
        [in] LPSTR pszPath);

    [entry("PathUndecorateW")]
    void PathUndecorate(
        [in] LPWSTR pszPath);

    [entry("PathUnExpandEnvStringsA")]
    BOOL PathUnExpandEnvStringsA(
        [in] LPSTR pszPath,
        [in] LPSTR pszBuf,
        [in] LONG cchBuf);

    [entry("PathUnExpandEnvStringsW")]
    BOOL PathUnExpandEnvStrings(
        [in] LPWSTR pszPath,
        [in] LPWSTR pszBuf,
        [in] LONG cchBuf);

//
// URL functions
//

    [entry("UrlCompareA")]
    LONG UrlCompareA(
        [in] LPSTR psz1,
        [in] LPSTR psz2,
        [in] BOOL fIgnoreSlash);

    [entry("UrlCompareW")]
    LONG UrlCompare(
        [in] LPWSTR psz1,
        [in] LPWSTR psz2,
        [in] BOOL fIgnoreSlash);

    [entry("UrlCombineA")]
    HRESULT UrlCombineA(
        [in] LPSTR pszBase,
        [in] LPSTR pszRelative,
        [in] LPSTR pszCombined,
        [in] LONG *pcchCombined,
        [in] LONG dwFlags);

    [entry("UrlCombineW")]
    HRESULT UrlCombine(
        [in] LPWSTR pszBase,
        [in] LPWSTR pszRelative,
        [in] LPWSTR pszCombined,
        [in] LONG *pcchCombined,
        [in] LONG dwFlags);

    [entry("UrlCanonicalizeA")]
    HRESULT UrlCanonicalizeA(
        [in] LPSTR pszUrl,
        [in] LPSTR pszCanonicalized,
        [in] LONG *pcchCanonicalized,
        [in] LONG dwFlags);

    [entry("UrlCanonicalizeW")]
    HRESULT UrlCanonicalize(
        [in] LPWSTR pszUrl,
        [in] LPWSTR pszCanonicalized,
        [in] LONG *pcchCanonicalized,
        [in] LONG dwFlags);

    [entry("UrlIsOpaqueA")]
    BOOL UrlIsOpaqueA(
        [in] LPSTR pszURL);

    [entry("UrlIsOpaqueW")]
    BOOL UrlIsOpaque(
        [in] LPWSTR pszURL);

    [entry("UrlIsNoHistoryA")]
    BOOL UrlIsNoHistoryA(
        [in] LPSTR pszURL);

    [entry("UrlIsNoHistoryW")]
    BOOL UrlIsNoHistory(
        [in] LPWSTR pszURL);

    [entry("UrlIsA")]
    BOOL UrlIsA(
        [in] LPSTR pszUrl,
        [in] URLIS UrlIs);

    [entry("UrlIsW")]
    BOOL UrlIs(
        [in] LPWSTR pszUrl,
        [in] URLIS UrlIs);

    [entry("UrlGetLocationA")]
    LONG UrlGetLocationA(
        [in] LPSTR psz1);

    [entry("UrlGetLocationW")]
    LONG UrlGetLocation(
        [in] LPWSTR psz1);

    [entry("UrlUnescapeA")]
    HRESULT UrlUnescapeA(
        [in] LPSTR pszUrl,
        [in] LPSTR pszUnescaped,
        [in] LONG *pcchUnescaped,
        [in] LONG dwFlags);

    [entry("UrlUnescapeW")]
    HRESULT UrlUnescape(
        [in] LPWSTR pszUrl,
        [in] LPWSTR pszUnescaped,
        [in] LONG *pcchUnescaped,
        [in] LONG dwFlags);

    [entry("UrlEscapeA")]
    HRESULT UrlEscapeA(
        [in] LPSTR pszUrl,
        [in] LPSTR pszEscaped,
        [in] LONG *pcchEscaped,
        [in] LONG dwFlags);

    [entry("UrlEscapeW")]
    HRESULT UrlEscape(
        [in] LPWSTR pszUrl,
        [in] LPWSTR pszEscaped,
        [in] LONG *pcchEscaped,
        [in] LONG dwFlags);

    [entry("UrlCreateFromPathA")]
    HRESULT UrlCreateFromPathA(
        [in] LPSTR pszPath,
        [in] LPSTR pszUrl,
        [in] LONG *pcchUrl,
        [in] LONG dwFlags);

    [entry("UrlCreateFromPathW")]
    HRESULT UrlCreateFromPath(
        [in] LPWSTR pszPath,
        [in] LPWSTR pszUrl,
        [in] LONG *pcchUrl,
        [in] LONG dwFlags);

    [entry("PathCreateFromUrlA")]
    HRESULT PathCreateFromUrlA(
        [in] LPSTR pszUrl,
        [in] LPSTR pszPath,
        [in] LONG *pcchPath,
        [in] LONG dwFlags);

    [entry("PathCreateFromUrlW")]
    HRESULT PathCreateFromUrl(
        [in] LPWSTR pszUrl,
        [in] LPWSTR pszPath,
        [in] LONG *pcchPath,
        [in] LONG dwFlags);

    [entry("UrlHashA")]
    HRESULT UrlHashA(
        [in] LPSTR pszUrl,
        [in] void *pbHash,
        [in] LONG cbHash);

    [entry("UrlHashW")]
    HRESULT UrlHash(
        [in] LPWSTR pszUrl,
        [in] void *pbHash,
        [in] LONG cbHash);

    [entry("UrlGetPartW")]
    HRESULT UrlGetPart(
        [in] LPWSTR pszIn,
        [in] LPWSTR pszOut,
        [in] LONG *pcchOut,
        [in] LONG dwPart,
        [in] LONG dwFlags);

    [entry("UrlGetPartA")]
    HRESULT UrlGetPartA(
        [in] LPSTR pszIn,
        [in] LPSTR pszOut,
        [in] LONG *pcchOut,
        [in] LONG dwPart,
        [in] LONG dwFlags);

    [entry("UrlApplySchemeA")]
    HRESULT UrlApplySchemeA(
        [in] LPSTR pszIn,
        [in] LPSTR pszOut,
        [in] LONG *pcchOut,
        [in] LONG dwFlags);

    [entry("UrlApplySchemeW")]
    HRESULT UrlApplyScheme(
        [in] LPWSTR pszIn,
        [in] LPWSTR pszOut,
        [in] LONG *pcchOut,
        [in] LONG dwFlags);

    [entry("HashData")]
    HRESULT HashData(
        [in] void *pbData,
        [in] LONG cbData,
        [in] void *pbHash,
        [in] LONG cbHash);

//
//=============== Registry Routines ===================================
//

// SHDeleteEmptyKey mimics RegDeleteKey as it behaves on NT.
// SHDeleteKey mimics RegDeleteKey as it behaves on Win95.

    [entry("SHDeleteEmptyKeyA")]
    LONG SHDeleteEmptyKeyA(
        [in] LONG hkey,
        [in] LPSTR pszSubKey);

    [entry("SHDeleteEmptyKeyW")]
    LONG SHDeleteEmptyKey(
        [in] LONG hkey,
        [in] LPWSTR pszSubKey);

    [entry("SHDeleteKeyA")]
    LONG SHDeleteKeyA(
        [in] LONG hkey,
        [in] LPSTR pszSubKey);

    [entry("SHDeleteKeyW")]
    LONG SHDeleteKey(
        [in] LONG hkey,
        [in] LPWSTR pszSubKey);

    [entry("SHRegDuplicateHKey")]
    LONG SHRegDuplicateHKey(
        [in] LONG hkey);

// These functions open the key, get/set/delete the value,
// then close the key.

    [entry("SHDeleteValueA")]
    LONG SHDeleteValueA(
        [in] LONG hkey,
        [in] LPSTR pszSubKey,
        [in] LPSTR pszValue);

    [entry("SHDeleteValueW")]
    LONG SHDeleteValue(
        [in] LONG hkey,
        [in] LPWSTR pszSubKey,
        [in] LPWSTR pszValue);

    [entry("SHGetValueA")]
    LONG SHGetValueA(
        [in] LONG hkey,
        [in] LPSTR pszSubKey,
        [in] LPSTR pszValue,
        [in] LONG *pdwType,
        [in] void *pvData,
        [in] LONG *pcbData);

    [entry("SHGetValueW")]
    LONG SHGetValue(
        [in] LONG hkey,
        [in] LPWSTR pszSubKey,
        [in] LPWSTR pszValue,
        [in] LONG *pdwType,
        [in] void *pvData,
        [in] LONG *pcbData);

    [entry("SHSetValueA")]
    LONG SHSetValueA(
        [in] LONG hkey,
        [in] LPSTR pszSubKey,
        [in] LPSTR pszValue,
        [in] LONG dwType,
        [in] void *pvData,
        [in] LONG cbData);

    [entry("SHSetValueW")]
    LONG SHSetValue(
        [in] LONG hkey,
        [in] LPWSTR pszSubKey,
        [in] LPWSTR pszValue,
        [in] LONG dwType,
        [in] void *pvData,
        [in] LONG cbData);

// These functions work just like RegQueryValueEx, except if the
// data type is REG_EXPAND_SZ, then these will go ahead and expand
// out the string. *pdwType will always be massaged to REG_SZ
// if this happens. REG_SZ values are also guaranteed to be null
// terminated.

    [entry("SHQueryValueExA")]
    LONG SHQueryValueExA(
        [in] LONG hkey,
        [in] LPSTR pszValue,
        [in] LONG *pdwReserved,
        [in] LONG *pdwType,
        [in] void *pvData,
        [in] LONG *pcbData);

    [entry("SHQueryValueExW")]
    LONG SHQueryValueEx(
        [in] LONG hkey,
        [in] LPWSTR pszValue,
        [in] LONG *pdwReserved,
        [in] LONG *pdwType,
        [in] void *pvData,
        [in] LONG *pcbData);

// Enumeration functions support.

    [entry("SHEnumKeyExA")]
    LONG SHEnumKeyExA(
        [in] LONG hkey,
        [in] LONG dwIndex,
        [in] LPSTR pszName,
        [in] LONG *pcchName);

    [entry("SHEnumKeyExW")]
    LONG SHEnumKeyEx(
        [in] LONG hkey,
        [in] LONG dwIndex,
        [in] LPWSTR pszName,
        [in] LONG *pcchName);

    [entry("SHEnumValueA")]
    LONG SHEnumValueA(
        [in] LONG hkey,
        [in] LONG dwIndex,
        [in] LPSTR pszValueName,
        [in] LONG *pcchValueName,
        [in] LONG *pdwType,
        [in] void *pvData,
        [in] LONG *pcbData);

    [entry("SHEnumValueW")]
    LONG SHEnumValue(
        [in] LONG hkey,
        [in] LONG dwIndex,
        [in] LPWSTR pszValueName,
        [in] LONG *pcchValueName,
        [in] LONG *pdwType,
        [in] void *pvData,
        [in] LONG *pcbData);

    [entry("SHQueryInfoKeyA")]
    LONG SHQueryInfoKeyA(
        [in] LONG hkey,
        [in] LONG *pcSubKeys,
        [in] LONG *pcchMaxSubKeyLen,
        [in] LONG *pcValues,
        [in] LONG *pcchMaxValueNameLen);

    [entry("SHQueryInfoKeyW")]
    LONG SHQueryInfoKey(
        [in] LONG hkey,
        [in] LONG *pcSubKeys,
        [in] LONG *pcchMaxSubKeyLen,
        [in] LONG *pcValues,
        [in] LONG *pcchMaxValueNameLen);

// recursive key copy
    [entry("SHCopyKeyA")]
    LONG SHCopyKeyA(
        [in] LONG hkeySrc,
        [in] LPSTR szSrcSubKey,
        [in] LONG hkeyDest,
        [in] LONG fReserved);

    [entry("SHCopyKeyW")]
    LONG SHCopyKey(
        [in] LONG hkeySrc,
        [in] LPWSTR wszSrcSubKey,
        [in] LONG hkeyDest,
        [in] LONG fReserved);

// Getting and setting file system paths with environment variables

    [entry("SHRegGetPathA")]
    LONG SHRegGetPathA(
        [in] LONG hKey,
        [in] LPSTR pcszSubKey,
        [in] LPSTR pcszValue,
        [in] LPSTR pszPath,
        [in] LONG dwFlags);

    [entry("SHRegGetPathW")]
    LONG SHRegGetPath(
        [in] LONG hKey,
        [in] LPWSTR pcszSubKey,
        [in] LPWSTR pcszValue,
        [in] LPWSTR pszPath,
        [in] LONG dwFlags);

    [entry("SHRegSetPathA")]
    LONG SHRegSetPathA(
        [in] LONG hKey,
        [in] LPSTR pcszSubKey,
        [in] LPSTR pcszValue,
        [in] LPSTR pcszPath,
        [in] LONG dwFlags);

    [entry("SHRegSetPathW")]
    LONG SHRegSetPath(
        [in] LONG hKey,
        [in] LPWSTR pcszSubKey,
        [in] LPWSTR pcszValue,
        [in] LPWSTR pcszPath,
        [in] LONG dwFlags);

//////////////////////////////////////////////
// User Specific Registry Access Functions
//////////////////////////////////////////////

    [entry("SHRegCreateUSKeyA")]
    LONG SHRegCreateUSKeyA(
        [in] LPSTR pszPath,
        [in] REGSAM samDesired,
        [in] LONG hRelativeUSKey,
        [in] LONG *phNewUSKey,
        [in] LONG dwFlags);

    [entry("SHRegCreateUSKeyW")]
    LONG SHRegCreateUSKey(
        [in] LPWSTR pwzPath,
        [in] REGSAM samDesired,
        [in] LONG hRelativeUSKey,
        [in] LONG *phNewUSKey,
        [in] LONG dwFlags);

    [entry("SHRegOpenUSKeyA")]
    LONG SHRegOpenUSKeyA(
        [in] LPSTR pszPath,
        [in] REGSAM samDesired,
        [in] LONG hRelativeUSKey,
        [in] LONG *phNewUSKey,
        [in] BOOL fIgnoreHKCU);

    [entry("SHRegOpenUSKeyW")]
    LONG SHRegOpenUSKey(
        [in] LPWSTR pwzPath,
        [in] REGSAM samDesired,
        [in] LONG hRelativeUSKey,
        [in] LONG *phNewUSKey,
        [in] BOOL fIgnoreHKCU);

    [entry("SHRegQueryUSValueA")]
    LONG SHRegQueryUSValueA(
        [in] LONG hUSKey,
        [in] LPSTR pszValue,
        [in] LONG *pdwType,
        [in] void *pvData,
        [in] LONG *pcbData,
        [in] BOOL fIgnoreHKCU,
        [in] void *pvDefaultData,
        [in] LONG dwDefaultDataSize);

    [entry("SHRegQueryUSValueW")]
    LONG SHRegQueryUSValue(
        [in] LONG hUSKey,
        [in] LPWSTR pwzValue,
        [in] LONG *pdwType,
        [in] void *pvData,
        [in] LONG *pcbData,
        [in] BOOL fIgnoreHKCU,
        [in] void *pvDefaultData,
        [in] LONG dwDefaultDataSize);

    [entry("SHRegWriteUSValueA")]
    LONG SHRegWriteUSValueA(
        [in] LONG hUSKey,
        [in] LPSTR pszValue,
        [in] LONG dwType,
        [in] void *pvData,
        [in] LONG cbData,
        [in] SHREGSET dwFlags);

    [entry("SHRegWriteUSValueW")]
    LONG SHRegWriteUSValue(
        [in] LONG hUSKey,
        [in] LPWSTR pwzValue,
        [in] LONG dwType,
        [in] void *pvData,
        [in] LONG cbData,
        [in] SHREGSET dwFlags);

    [entry("SHRegDeleteUSValueA")]
    LONG SHRegDeleteUSValueA(
        [in] LONG hUSKey,
        [in] LPSTR pszValue,
        [in] SHREGDEL_FLAGS delRegFlags);

    [entry("SHRegDeleteEmptyUSKeyW")]
    LONG SHRegDeleteEmptyUSKey(
        [in] LONG hUSKey,
        [in] LPWSTR pwzSubKey,
        [in] SHREGDEL_FLAGS delRegFlags);

    [entry("SHRegDeleteEmptyUSKeyA")]
    LONG SHRegDeleteEmptyUSKeyA(
        [in] LONG hUSKey,
        [in] LPSTR pszSubKey,
        [in] SHREGDEL_FLAGS delRegFlags);

    [entry("SHRegDeleteUSValueW")]
    LONG SHRegDeleteUSValue(
        [in] LONG hUSKey,
        [in] LPWSTR pwzValue,
        [in] SHREGDEL_FLAGS delRegFlags);

    [entry("SHRegEnumUSKeyA")]
    LONG SHRegEnumUSKeyA(
        [in] LONG hUSKey,
        [in] LONG dwIndex,
        [in] LPSTR pszName,
        [in] LONG *pcchName,
        [in] SHREGENUM_FLAGS enumRegFlags);

    [entry("SHRegEnumUSKeyW")]
    LONG SHRegEnumUSKey(
        [in] LONG hUSKey,
        [in] LONG dwIndex,
        [in] LPWSTR pwzName,
        [in] LONG *pcchName,
        [in] SHREGENUM_FLAGS enumRegFlags);

    [entry("SHRegEnumUSValueA")]
    LONG SHRegEnumUSValueA(
        [in] LONG hUSKey,
        [in] LONG dwIndex,
        [in] LPSTR pszValueName,
        [in] LONG *pcchValueName,
        [in] LONG *pdwType,
        [in] void *pvData,
        [in] LONG *pcbData,
        [in] SHREGENUM_FLAGS enumRegFlags);

    [entry("SHRegEnumUSValueW")]
    LONG SHRegEnumUSValue(
        [in] LONG hUSKey,
        [in] LONG dwIndex,
        [in] LPWSTR pszValueName,
        [in] LONG *pcchValueName,
        [in] LONG *pdwType,
        [in] void *pvData,
        [in] LONG *pcbData,
        [in] SHREGENUM_FLAGS enumRegFlags);

    [entry("SHRegQueryInfoUSKeyA")]
    LONG SHRegQueryInfoUSKeyA(
        [in] LONG hUSKey,
        [in] LONG *pcSubKeys,
        [in] LONG *pcchMaxSubKeyLen,
        [in] LONG *pcValues,
        [in] LONG *pcchMaxValueNameLen,
        [in] SHREGENUM_FLAGS enumRegFlags);

    [entry("SHRegQueryInfoUSKeyW")]
    LONG SHRegQueryInfoUSKey(
        [in] LONG hUSKey,
        [in] LONG *pcSubKeys,
        [in] LONG *pcchMaxSubKeyLen,
        [in] LONG *pcValues,
        [in] LONG *pcchMaxValueNameLen,
        [in] SHREGENUM_FLAGS enumRegFlags);

    [entry("SHRegCloseUSKey")]
    LONG SHRegCloseUSKey(
        [in] LONG hUSKey);

// These calls are equal to an SHRegOpenUSKey, SHRegQueryUSValue, and then a SHRegCloseUSKey.

    [entry("SHRegGetUSValueA")]
    LONG SHRegGetUSValueA(
        [in] LPSTR pszSubKey,
        [in] LPSTR pszValue,
        [in] LONG *pdwType,
        [in] void * pvData,
        [in] LONG *pcbData,
        [in] BOOL fIgnoreHKCU,
        [in] void *pvDefaultData,
        [in] LONG dwDefaultDataSize);

    [entry("SHRegGetUSValueW")]
    LONG SHRegGetUSValue(
        [in] LPWSTR pwzSubKey,
        [in] LPWSTR pwzValue,
        [in] LONG *pdwType,
        [in] void * pvData,
        [in] LONG *pcbData,
        [in] BOOL fIgnoreHKCU,
        [in] void *pvDefaultData,
        [in] LONG dwDefaultDataSize);

    [entry("SHRegSetUSValueA")]
    LONG SHRegSetUSValueA(
        [in] LPSTR pszSubKey,
        [in] LPSTR pszValue,
        [in] LONG dwType,
        [in] void *pvData,
        [in] LONG cbData,
        [in] SHREGSET dwFlags);

    [entry("SHRegSetUSValueW")]
    LONG SHRegSetUSValue(
        [in] LPWSTR pwzSubKey,
        [in] LPWSTR pwzValue,
        [in] LONG dwType,
        [in] void *pvData,
        [in] LONG cbData,
        [in] SHREGSET dwFlags);

    [entry("SHRegGetIntW")]
    LONG SHRegGetInt(
        [in] LONG hk,
        [in] LPWSTR pwzKey,
        [in] int iDefault);

    [entry("SHRegGetBoolUSValueA")]
    BOOL SHRegGetBoolUSValueA(
        [in] LPSTR pszSubKey,
        [in] LPSTR pszValue,
        [in] BOOL fIgnoreHKCU, 
        [in] BOOL fDefault);

    [entry("SHRegGetBoolUSValueW")]
    BOOL SHRegGetBoolUSValue(
        [in] LPWSTR pszSubKey,
        [in] LPWSTR pszValue,
        [in] BOOL fIgnoreHKCU,
        [in] BOOL fDefault);

//
// Association APIs
//
// these APIs are to assist in accessing the data in HKCR
// getting the Command strings and exe paths
// for different verbs and extensions are simplified this way
//

	//Original def was missing 4 bytes. A CLSID is 16 bytes, 2 longs and 2 shorts is 12.
    [entry("AssocCreate")]
    HRESULT AssocCreate(
        [in] LONG clsid_data1,
        [in] short clsid_data2,
        [in] short clsid_data3,
        [in] BYTE clsid_data4[8],
        [in] UUID *riid,
        [out] void *ppv);

// wrappers for the interface

    [entry("AssocQueryStringA")]
    HRESULT AssocQueryStringA(
        [in] ASSOCF flags,
        [in] ASSOCSTR str,
        [in] LPSTR pszAssoc,
        [in] LPSTR pszExtra,
        [in] LPSTR pszOut,
        [in, out] LONG *pcchOut);

// wrappers for the interface

    [entry("AssocQueryStringW")]
    HRESULT AssocQueryString(
        [in] ASSOCF flags,
        [in] ASSOCSTR str,
        [in] LPWSTR pszAssoc,
        [in] LPWSTR pszExtra,
        [in] LPWSTR pszOut,
        [in, out] LONG *pcchOut);

    [entry("AssocQueryStringByKeyA")]
    HRESULT AssocQueryStringByKeyA(
        [in] ASSOCF flags,
        [in] ASSOCSTR str,
        [in] LONG hkAssoc,
        [in] LPSTR pszExtra,
        [in] LPSTR pszOut,
        [in, out] LONG *pcchOut);

    [entry("AssocQueryStringByKeyW")]
    HRESULT AssocQueryStringByKey(
        [in] ASSOCF flags,
        [in] ASSOCSTR str,
        [in] LONG hkAssoc,
        [in] LPWSTR pszExtra,
        [in] LPWSTR pszOut,
        [in, out] LONG *pcchOut);

    [entry("AssocQueryKeyA")]
    HRESULT AssocQueryKeyA(
        [in] ASSOCF flags,
        [in] ASSOCKEY key,
        [in] LPSTR pszAssoc,
        [in] LPSTR pszExtra,
        [in, out] LONG *phkeyOut);

    [entry("AssocQueryKeyW")]
    HRESULT AssocQueryKey(
        [in] ASSOCF flags,
        [in] ASSOCKEY key,
        [in] LPWSTR pszAssoc,
        [in] LPWSTR pszExtra,
        [in, out] LONG *phkeyOut);

//
//=============== Stream Routines ===================================
//

    [entry("SHOpenRegStreamA")]
    IStream *SHOpenRegStreamA(
        [in] LONG hkey,
        [in] LPSTR pszSubkey,
        [in] LPSTR pszValue,
        [in] STGM grfMode);

    [entry("SHOpenRegStreamW")]
    IStream *SHOpenRegStream(
        [in] LONG hkey,
        [in] LPWSTR pszSubkey,
        [in] LPWSTR pszValue,
        [in] STGM grfMode);

    [entry("SHOpenRegStream2A")]
    IStream *SHOpenRegStream2A(
        [in] LONG hkey,
        [in] LPSTR pszSubkey,
        [in] LPSTR pszValue,
        [in] STGM grfMode);

    [entry("SHOpenRegStream2W")]
    IStream *SHOpenRegStream2(
        [in] LONG hkey,
        [in] LPWSTR pszSubkey,
        [in] LPWSTR pszValue,
        [in] STGM grfMode);

    [entry("SHCreateStreamOnFileA")]
    HRESULT SHCreateStreamOnFileA(
        [in] LPSTR pszFile,
        [in] STGM grfMode,
        [out, retval] IStream **ppstm);

    [entry("SHCreateStreamOnFileW")]
    HRESULT SHCreateStreamOnFile(
        [in] LPWSTR pszFile,
        [in] STGM grfMode,
        [out, retval] IStream **ppstm);

// SHAutoComplete
// hwndEdit - HWND of editbox, ComboBox or ComboBoxEx.
// dwFlags - Flags to indicate what to AutoAppend or AutoSuggest for the editbox.
//
// WARNING:
// Caller needs to have called CoInitialize() or OleInitialize()
// and cannot call CoUninit/OleUninit until after
// WM_DESTROY on hwndEdit.
//
// dwFlags values:

    [entry("SHAutoComplete")]
    HRESULT SHAutoComplete(
        [in] LONG hwndEdit,
        [in] SHACF dwFlags);

    [entry("SHSetThreadRef")]
    HRESULT SHSetThreadRef(
        [in] stdole.IUnknown *punk);

    [entry("SHGetThreadRef")]
    HRESULT SHGetThreadRef(
        [out, retval] IUnknown **ppunk);

    [entry("SHSkipJunction")]
    BOOL SHSkipJunction(
        [in] IBindCtx *pbc,
        [in] UUID *pclsid);

    [entry("SHCreateThread")]
    BOOL SHCreateThread(
        [in] LONG pfnThreadProc,
        [in] void *pData,
        [in] CTF dwFlags,
        [in] LONG pfnCallback);

//
//====== GDI helper functions ================================================
//

    [entry("SHCreateShellPalette")]
    LONG SHCreateShellPalette(
        [in] LONG hdc);

    [entry("ColorRGBToHLS")]
    void ColorRGBToHLS(
        [in] LONG clrRGB,
        [out] short* pwHue,
        [out] short* pwLuminance,
        [out] short* pwSaturation);

    [entry("ColorHLSToRGB")]
    LONG ColorHLSToRGB(
        [in] short wHue,
        [in] short wLuminance,
        [in] short wSaturation);

    [entry("ColorAdjustLuma")]
    LONG ColorAdjustLuma(
        [in] LONG clrRGB,
        [in] int n,
        [in] BOOL fScale);

}

[
    odl,
    uuid(c46ca590-3c3f-11d2-bee6-0000f805ca57)
]
interface IQueryAssociations : stdole.IUnknown {

    HRESULT Init(
        [in] ASSOCF flags,
        [in] LPWSTR pszAssoc,
        [in] LONG hkProgid,
        [in] LONG hwnd);

    HRESULT GetString(
        [in] ASSOCF flags,
        [in] ASSOCSTR str,
        [in] LPWSTR pszExtra,
        [in] LPWSTR pszOut,
        [in, out] LONG *pcchOut);

    HRESULT GetKey(
        [in] ASSOCF flags,
        [in] ASSOCKEY key,
        [in] LPWSTR pszExtra,
        [in, out] LONG *phkeyOut);

    HRESULT GetData(
        [in] ASSOCF flags,
        [in] ASSOCDATA data,
        [in] LPWSTR pszExtra,
        [in] void *pvOut,
        [in, out] LONG *pcbOut);

    HRESULT GetEnum(
        [in] ASSOCF flags,
        [in] ASSOCENUM assocenum,
        [in] LPWSTR pszExtra,
        [in] UUID *riid,
        [in] void *ppvOut);
};

