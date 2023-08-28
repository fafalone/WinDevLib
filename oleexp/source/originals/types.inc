
typedef unsigned char BYTE;
typedef [public] LONG BOOL;

typedef enum HRESULTS {
    S_OK = 0,
    S_FALSE = 1,

    E_NOTIMPL = 0x80004001,
    E_OUTOFMEMORY = 0x8007000E,
    E_INVALIDARG = 0x80070057,
    E_NOINTERFACE = 0x80004002,
    E_POINTER = 0x80004003,
    E_HANDLE = 0x80070006,
    E_ABORT = 0x80004004,
    E_FAIL = 0x80004005,
    E_ACCESSDENIED = 0x80070005,

    // IConnectionPoint errors
        
    CONNECT_E_FIRST = 0x80040200,
    CONNECT_E_NOCONNECTION,  // there is no connection for this connection id
    CONNECT_E_ADVISELIMIT,   // this implementation's limit for advisory connections has been reached
    CONNECT_E_CANNOTCONNECT, // connection attempt failed
    CONNECT_E_OVERRIDDEN,    // must use a derived interface to connect
    
    // DllRegisterServer/DllUnregisterServer errors
    SELFREG_E_TYPELIB = 0x80040200, // failed to register/unregister type library
    SELFREG_E_CLASS,                // failed to register/unregister class

    // IPerPropertyBrowsing status codes

    PERPROP_E_NOPAGEAVAILABLE = 0x80040200, // no page available for requested property

    // INET errors
    
    INET_E_INVALID_URL = 0x800C0002,
    INET_E_NO_SESSION = 0x800C0003,
    INET_E_CANNOT_CONNECT = 0x800C0004,
    INET_E_RESOURCE_NOT_FOUND = 0x800C0005,
    INET_E_OBJECT_NOT_FOUND = 0x800C0006,
    INET_E_DATA_NOT_AVAILABLE = 0x800C0007,
    INET_E_DOWNLOAD_FAILURE = 0x800C0008,
    INET_E_AUTHENTICATION_REQUIRED = 0x800C0009,
    INET_E_NO_VALID_MEDIA = 0x800C000A,
    INET_E_CONNECTION_TIMEOUT = 0x800C000B,
    INET_E_INVALID_REQUEST = 0x800C000C,
    INET_E_UNKNOWN_PROTOCOL = 0x800C000D,
    INET_E_SECURITY_PROBLEM = 0x800C000E,
    INET_E_CANNOT_LOAD_DATA = 0x800C000F,
    INET_E_CANNOT_INSTANTIATE_OBJECT = 0x800C0010,
    INET_E_USE_DEFAULT_PROTOCOLHANDLER = 0x800C0011,
    INET_E_DEFAULT_ACTION = 0x800C0011,
    INET_E_USE_DEFAULT_SETTING = 0x800C0012,
    INET_E_QUERYOPTION_UNKNOWN = 0x800C0013,
    INET_E_REDIRECTING = 0x800C0014,
    INET_E_REDIRECT_FAILED = 0x800C0014,
    INET_E_REDIRECT_TO_DIR = 0x800C0015,
    INET_E_CANNOT_LOCK_REQUEST = 0x800C0016,
    INET_E_USE_EXTEND_BINDING = 0x800C0017,
    INET_E_ERROR_FIRST = 0x800C0002,
    INET_E_ERROR_LAST = 0x800C0017,
    INET_E_CODE_DOWNLOAD_DECLINED = 0x800C0100,
    INET_E_RESULT_DISPATCHED = 0x800C0200,
    INET_E_CANNOT_REPLACE_SFP_FILE = 0x800C0300,

} HRESULTS;

typedef struct UUID {
    LONG Data1;
    SHORT Data2;
    SHORT Data3;
    BYTE Data4[8];
} UUID;

typedef enum VARENUM {
    VT_EMPTY = 0,
    VT_NULL = 1,
    VT_I2 = 2,
    VT_I4 = 3,
    VT_R4 = 4,
    VT_R8 = 5,
    VT_CY = 6,
    VT_DATE = 7,
    VT_BSTR = 8,
    VT_DISPATCH = 9,
    VT_ERROR = 10,
    VT_BOOL = 11,
    VT_VARIANT = 12,
    VT_UNKNOWN = 13,
    VT_DECIMAL = 14,
    VT_I1 = 16,
    VT_UI1 = 17,
    VT_UI2 = 18,
    VT_UI4 = 19,
    VT_I8 = 20,
    VT_UI8 = 21,
    VT_INT = 22,
    VT_UINT = 23,
    VT_VOID = 24,
    VT_HRESULT = 25,
    VT_PTR = 26,
    VT_SAFEARRAY = 27,
    VT_CARRAY = 28,
    VT_USERDEFINED = 29,
    VT_LPSTR = 30,
    VT_LPWSTR = 31,
    VT_FILETIME = 64,
    VT_BLOB = 65,
    VT_STREAM = 66,
    VT_STORAGE = 67,
    VT_STREAMED_OBJECT = 68,
    VT_STORED_OBJECT = 69,
    VT_BLOB_OBJECT = 70,
    VT_CF = 71,
    VT_CLSID = 72,
    VT_STREAMED_PROPSET = 73,
    VT_STORED_PROPSET = 74,
    VT_BLOB_PROPSET = 75,
    VT_VERBOSE_ENUM = 76,
    VT_BSTR_BLOB = 0xFFF,
    VT_VECTOR = 0x1000,
    VT_ARRAY = 0x2000,
    VT_BYREF = 0x4000,
    VT_RESERVED = 0x8000,
    VT_ILLEGAL = 0xffff,
    VT_ILLEGALMASKED = 0x0fff,
    VT_TYPEMASK = 0x0fff
} VARENUM;

typedef enum SummaryInfoProperties {
    [helpstring("Document Title - VT_LPSTR")]
    PIDSI_TITLE = 0x00000002,
    [helpstring("Document Subject - VT_LPSTR")]
    PIDSI_SUBJECT = 0x00000003,
    [helpstring("Document Author - VT_LPSTR")]
    PIDSI_AUTHOR = 0x00000004,
    [helpstring("Keywords - VT_LPSTR")]
    PIDSI_KEYWORDS = 0x00000005,
    [helpstring("Comments - VT_LPSTR")]
    PIDSI_COMMENTS = 0x00000006,
    [helpstring("Template name - VT_LPSTR")]
    PIDSI_TEMPLATE = 0x00000007,
    [helpstring("Last Saved By - VT_LPSTR")]
    PIDSI_LASTAUTHOR = 0x00000008,
    [helpstring("Revision Number (save count) - VT_LPSTR")]
    PIDSI_REVNUMBER = 0x00000009,
    [helpstring("Total editing time - VT_CURRENCY (UTC)")]
    PIDSI_EDITTIME = 0x0000000a,
    [helpstring("Last printed date - VT_CURRENCY (UTC)")]
    PIDSI_LASTPRINTED = 0x0000000b,
    [helpstring("Create Date - VT_CURRENCY (UTC)")]
    PIDSI_CREATE_DTM = 0x0000000c,
    [helpstring("Last save date - VT_CURRENCY (UTC)")]
    PIDSI_LASTSAVE_DTM = 0x0000000d,
    [helpstring("Page count - VT_I4")]
    PIDSI_PAGECOUNT = 0x0000000e,
    [helpstring("Word count - VT_I4")]
    PIDSI_WORDCOUNT = 0x0000000f,
    [helpstring("Char count - VT_I4")]
    PIDSI_CHARCOUNT = 0x00000010,
    [helpstring("Thumbnail - VT_CF")]
    PIDSI_THUMBNAIL = 0x00000011,
    [helpstring("Application Name - VT_LPSTR")]
    PIDSI_APPNAME = 0x00000012,
    [helpstring("Document security - VT_I4")]
    PIDSI_DOC_SECURITY = 0x00000013,
} SummaryInfoProperties;

typedef enum DocSummaryInfoProperties {
    [helpstring("Category - LPSTR")]
    PID_CATEGORY = 0x2,
    [helpstring("Presentation Target - LPSTR")]
    PID_PRESFORMAT = 0x3,
    [helpstring("Byte count - VT_I4")]
    PID_BYTECOUNT = 0x4,
    [helpstring("Line count - VT_I4")]
    PID_LINECOUNT = 0x5,
    [helpstring("Paragraph count - VT_I4")]
    PID_PARCOUNT = 0x6,
    [helpstring("Slide count - VT_I4")]
    PID_SLIDECOUNT = 0x7,
    [helpstring("Note count - VT_I4")]
    PID_NOTECOUNT = 0x8,
    [helpstring("Hidden slide count - VT_I4")]
    PID_HIDDENCOUNT = 0x9,
    [helpstring("MMClip count - VT_I4")]
    PID_MMCLIPCOUNT = 0xA,
    [helpstring("Scale/Crop - VT_BOOL")]
    PID_SCALE = 0xB,
    [helpstring("Heading Pairs - VT_VARIANT | VT_VECTOR")]
    PID_HEADINGPAIR = 0xC,
    [helpstring("Titles Of Parts - VT_LPSTR | VT_VECTOR")]
    PID_DOCPARTS = 0xD,
    [helpstring("Manager - VT_LPSTR")]
    PID_MANAGER = 0xE,
    [helpstring("Company - VT_LPSTR")]
    PID_COMPANY = 0xF,
    [helpstring("LinksUpToDate - VT_BOOL")]
    PID_LINKSDIRTY = 0x10,
} DocSummaryInfoProperties;

/* Storage instantiation modes */
typedef enum STGM {
    STGM_DIRECT = 0x00000000,
    STGM_TRANSACTED = 0x00010000,
    STGM_SIMPLE = 0x08000000,

    STGM_READ = 0x00000000,
    STGM_WRITE = 0x00000001,
    STGM_READWRITE = 0x00000002,

    STGM_SHARE_DENY_NONE = 0x00000040,
    STGM_SHARE_DENY_READ = 0x00000030,
    STGM_SHARE_DENY_WRITE = 0x00000020,
    STGM_SHARE_EXCLUSIVE = 0x00000010,

    STGM_PRIORITY = 0x00040000,
    STGM_DELETEONRELEASE = 0x04000000,
    STGM_NOSCRATCH = 0x00100000,

    STGM_CREATE = 0x00001000,
    STGM_CONVERT = 0x00020000,
    STGM_FAILIFTHERE = 0x00000000,

    STGM_NOSNAPSHOT = 0x00200000,
    STGM_DIRECT_SWMR = 0x00400000,
} STGM;

typedef LPWSTR SNB;

typedef struct SYSTEMTIME {
    short wYear;
    short wMonth;
    short wDayOfWeek;
    short wDay;
    short wHour;
    short wMinute;
    short wSecond;
    short wMilliseconds;
} SYSTEMTIME;

typedef enum STGC {
     STGC_DEFAULT = 0,
     STGC_OVERWRITE = 1,
     STGC_ONLYIFCURRENT = 2,
     STGC_DANGEROUSLYCOMMITMERELYTODISKCACHE = 4,
     STGC_CONSOLIDATE = 8
} STGC;

typedef enum STATFLAG {
    STATFLAG_DEFAULT = 0,
    STATFLAG_NONAME = 1,
    STATFLAG_NOOPEN = 2
} STATFLAG;

typedef enum STGTY {
    STGTY_STORAGE = 1,
    STGTY_STREAM = 2,
    STGTY_LOCKBYTES = 3,
    STGTY_PROPERTY = 4
} STGTY;

typedef enum LOCKTYPE {
    LOCK_WRITE = 1,
    LOCK_EXCLUSIVE = 2,
    LOCK_ONLYONCE = 4
} LOCKTYPE;

typedef struct STATSTG {
    LONG pwcsName;
    STGTY type;
    CURRENCY cbSize;
    CURRENCY mtime;
    CURRENCY ctime;
    CURRENCY atime;
    STGM grfMode;
    LOCKTYPE grfLocksSupported;
    UUID clsid;
    LONG grfStateBits;
    LONG reserved;
} STATSTG;

typedef enum STREAM_SEEK {
    STREAM_SEEK_SET = 0,
    STREAM_SEEK_CUR = 1,
    STREAM_SEEK_END = 2
} STREAM_SEEK;

typedef enum STGMOVE {
    STGMOVE_MOVE = 0,
    STGMOVE_COPY = 1,
    STGMOVE_SHALLOWCOPY = 2
} STGMOVE;

typedef enum OLEGETMONIKER {
    OLEGETMONIKER_ONLYIFTHERE = 1,
    OLEGETMONIKER_FORCEASSIGN = 2,
    OLEGETMONIKER_UNASSIGN = 3,
    OLEGETMONIKER_TEMPFORUSER = 4
} OLEGETMONIKER;

typedef enum OLEWHICHMK {
    OLEWHICHMK_CONTAINER = 1,
    OLEWHICHMK_OBJREL = 2,
    OLEWHICHMK_OBJFULL = 3
} OLEWHICHMK;

typedef enum OLECONTF {
    OLECONTF_EMBEDDINGS = 1,
    OLECONTF_LINKS = 2,
    OLECONTF_OTHERS = 4,
    OLECONTF_ONLYUSER = 8,
    OLECONTF_ONLYIFRUNNING = 16
} OLECONTF;

typedef enum OLECLOSE {
    OLECLOSE_SAVEIFDIRTY = 0,
    OLECLOSE_NOSAVE = 1,
    OLECLOSE_PROMPTSAVE = 2
} OLECLOSE;

typedef enum USERCLASSTYPE {
    USERCLASSTYPE_FULL = 1,
    USERCLASSTYPE_SHORT = 2,
    USERCLASSTYPE_APPNAME = 3
} USERCLASSTYPE;

typedef enum OLEVERBATTRIB {
    OLEVERBATTRIB_NEVERDIRTIES = 1,
    OLEVERBATTRIB_ONCONTAINERMENU = 2
} OLEVERBATTRIB;

typedef enum DVASPECT {
    DVASPECT_CONTENT = 1,
    DVASPECT_THUMBNAIL = 2,
    DVASPECT_ICON = 4,
    DVASPECT_DOCPRINT = 8,
    DVASPECT_OPAQUE = 16,
    DVASPECT_TRANSPARENT = 32
} DVASPECT;

typedef enum TYMED {
    TYMED_HGLOBAL = 1,
    TYMED_FILE = 2,
    TYMED_ISTREAM = 4,
    TYMED_ISTORAGE = 8,
    TYMED_GDI = 16,
    TYMED_MFPICT = 32,
    TYMED_ENHMF = 64,
    TYMED_NULL = 0
} TYMED;

typedef enum OLEMISC {
    OLEMISC_RECOMPOSEONRESIZE = 1,
    OLEMISC_ONLYICONIC = 2,
    OLEMISC_INSERTNOTREPLACE = 4,
    OLEMISC_STATIC = 8,
    OLEMISC_CANTLINKINSIDE = 16,
    OLEMISC_CANLINKBYOLE1 = 32,
    OLEMISC_ISLINKOBJECT = 64,
    OLEMISC_INSIDEOUT = 128,
    OLEMISC_ACTIVATEWHENVISIBLE = 256,
    OLEMISC_RENDERINGISDEVICEINDEPENDENT = 512,
    OLEMISC_INVISIBLEATRUNTIME = 1024,
    OLEMISC_ALWAYSRUN = 2048,
    OLEMISC_ACTSLIKEBUTTON = 4096,
    OLEMISC_ACTSLIKELABEL = 8192,
    OLEMISC_NOUIACTIVATE = 16384,
    OLEMISC_ALIGNABLE = 32768,
    OLEMISC_SIMPLEFRAME = 0x00010000,
    OLEMISC_SETCLIENTSITEFIRST = 0x00020000,
    OLEMISC_IMEMODE = 0x00040000,
    OLEMISC_IGNOREACTIVATEWHENVISIBLE = 0x00080000,
    OLEMISC_WANTSTOMENUMERGE = 0x00100000,
    OLEMISC_SUPPORTSMULTILEVELUNDO = 0x00200000
} OLEMISC;

typedef enum DVASPECTINFOFLAG {
    DVASPECTINFOFLAG_CANOPTIMIZE = 1
} DVASPECTINFOFLAG;

typedef enum DATADIR {
    DATADIR_GET = 1,
    DATADIR_SET = 2
} DATADIR;

typedef enum OLEVERBS {
    OLEIVERB_PRIMARY = 0,
    OLEIVERB_SHOW = -1,
    OLEIVERB_OPEN = -2,
    OLEIVERB_HIDE = -3,
    OLEIVERB_UIACTIVATE = -4,
    OLEIVERB_INPLACEACTIVATE = -5,
    OLEIVERB_DISCARDUNDOSTATE = -6,
    OLEIVERB_PROPERTIES = -7
} OLEVERBS;

typedef enum ADVF {
    ADVF_NODATA = 1,
    ADVF_PRIMEFIRST = 2,
    ADVF_ONLYONCE = 4,
    ADVF_DATAONSTOP = 64,
    ADVFCACHE_NOHANDLER = 8,
    ADVFCACHE_FORCEBUILTIN = 16,
    ADVFCACHE_ONSAVE = 32
} ADVF;

typedef struct OLEVERB {
    long lVerb;
    long lpszVerbName;
    long fuFlags;
    OLEVERBATTRIB grfAttribs;
} OLEVERB;

typedef struct OLEMENUGROUPWIDTHS {
    LONG width[6];
} OLEMENUGROUPWIDTHS;

typedef struct RECT {
    long Left;
    long Top;
    long Right;
    long Bottom;
} RECT;

typedef struct SIZE {
    long cx;
    long cy;
} SIZE;

typedef struct STGMEDIUM {
    TYMED TYMED;
    long data;
    long pUnkForRelease;
} STGMEDIUM;

typedef struct MSG {
    long hwnd;
    long message;
    long wParam;
    long lParam;
    long time;
    long pt_x;
    long pt_y;
} MSG;

typedef struct DVTARGETDEVICE {
    long tdSize;
    short tdDriverNameOffset;
    short tdDeviceNameOffset;
    short tdPortNameOffset;
    short tdExtDevmodeOffset;
    [helpstring("size_is(tdSize - sizeof(long) - 4 * sizeof(SHORT))")]
    BYTE tdData[1];
} DVTARGETDEVICE;

typedef struct FORMATETC {
    short cfFormat;
    long pDVTARGETDEVICE;
    DVASPECT dwAspect;
    long lindex;
    TYMED TYMED;
} FORMATETC;

typedef struct STATDATA {
    FORMATETC FORMATETC;
    ADVF grfAdvf;
    IAdviseSink *pIAdviseSink;
    long dwConnection;
} STATDATA;

typedef struct PALETTEENTRY {
    BYTE peRed;
    BYTE peGreen;
    BYTE peBlue;
    BYTE peFlags;
} PALETTEENTRY;

typedef struct LOGPALETTE {
    short palVersion;
    short palNumEntries;
    PALETTEENTRY* palPalEntry;
} LOGPALETTE;

typedef struct DVASPECTINFO {
    long cb;
    DVASPECTINFOFLAG dwFlags;
} DVASPECTINFO;

typedef struct OLEINPLACEFRAMEINFO {
    long cb;
    long fMDIApp;
    long hwndFrame;
    long haccel;
    long cAccelEntries;
} OLEINPLACEFRAMEINFO;

typedef struct CSPLATFORM {
    LONG dwPlatformId;
    LONG dwVersionHi;
    LONG dwVersionLo;
    LONG dwProcessorArch;
} CSPLATFORM;

typedef struct QUERYCONTEXT {
    LONG dwContext;
    CSPLATFORM Platform;
    LONG Locale;
    LONG dwVersionHi;
    LONG dwVersionLo;
} QUERYCONTEXT;

typedef enum TYSPEC {
    TYSPEC_CLSID = 0,
    TYSPEC_FILEEXT,
    TYSPEC_MIMETYPE,
    TYSPEC_FILENAME,
    TYSPEC_PROGID,
    TYSPEC_PACKAGENAME,
    TYSPEC_OBJECTID
} TYSPEC;

typedef struct uCLSSPEC {
    TYSPEC tyspec;
    BYTE Data[32];
} uCLSSPEC;

// flags for FaultInIEFeature
typedef enum FIEF {
    FIEF_FLAG_FORCE_JITUI = 0x1,  // force JIT ui even if
                                  // previoulsy rejected by
                                  // user in this session or
                                  // marked as Never Ask Again
    FIEF_FLAG_PEEK = 0x2,         // just peek, don't faultin
    FIEF_FLAG_SKIP_INSTALLED_VERSION_CHECK = 0x4 // force JIT without checking local version

} FIEF;

typedef enum BINDVERB {
    BINDVERB_GET    = 0x00000000,       // default action
    BINDVERB_POST   = 0x00000001,       // post verb
    BINDVERB_PUT    = 0x00000002,       // put verb
    BINDVERB_CUSTOM = 0x00000003,       // custom verb
} BINDVERB;

typedef enum BINDINFOF {
    BINDINFOF_URLENCODESTGMEDDATA = 0x00000001,       // url encode the stgmed data
    BINDINFOF_URLENCODEDEXTRAINFO = 0x00000002,       // url encode the extra info
} BINDINFOF;

// flags that describe the type of transaction that caller wants
typedef enum BINDF {
    BINDF_ASYNCHRONOUS             = 0x00000001,
    BINDF_ASYNCSTORAGE             = 0x00000002,
    BINDF_NOPROGRESSIVERENDERING   = 0x00000004,
    BINDF_OFFLINEOPERATION         = 0x00000008,
    BINDF_GETNEWESTVERSION         = 0x00000010,
    BINDF_NOWRITECACHE             = 0x00000020,
    BINDF_NEEDFILE                 = 0x00000040,
    BINDF_PULLDATA                 = 0x00000080,
    BINDF_IGNORESECURITYPROBLEM    = 0x00000100,
    BINDF_RESYNCHRONIZE            = 0x00000200,
    BINDF_HYPERLINK                = 0x00000400,
    BINDF_NO_UI                    = 0x00000800,
    BINDF_SILENTOPERATION          = 0x00001000,
    BINDF_PRAGMA_NO_CACHE          = 0x00002000,
    BINDF_GETCLASSOBJECT           = 0x00004000,
    BINDF_RESERVED_1               = 0x00008000,
    BINDF_FREE_THREADED            = 0x00010000,
    BINDF_DIRECT_READ              = 0x00020000,
    BINDF_FORMS_SUBMIT             = 0x00040000,
    BINDF_GETFROMCACHE_IF_NET_FAIL = 0x00080000,
    BINDF_FROMURLMON               = 0x00100000,
    BINDF_FWD_BACK                 = 0x00200000,
    BINDF_PREFERDEFAULTHANDLER     = 0x00400000,
    BINDF_ENFORCERESTRICTED        = 0x00800000,
} BINDF;

typedef enum URL_ENCODING {
    URL_ENCODING_NONE = 0x00000000,
    URL_ENCODING_ENABLE_UTF8 = 0x10000000,
    URL_ENCODING_DISABLE_UTF8 = 0x20000000,
} URL_ENCODING;

typedef struct SECURITY_ATTRIBUTES {
    LONG nLength;
    LONG lpSecurityDescriptor;
    BOOL bInheritHandle;
} SECURITY_ATTRIBUTES;

typedef struct BINDINFO {
    LONG cbSize;
    LONG szExtraInfo;        // e.g. string added to the URL string [range byte]
    STGMEDIUM stgmedData;    // e.g. put or post data
    LONG grfBindInfoF;       // e.g. how meddata should be treaded
    LONG dwBindVerb;         // e.g. put, post, ...
    LONG szCustomVerb;       // e.g. custom verb
    LONG cbstgmedData;       // size of data in stgmedData
    LONG dwOptions;          // additional option
    LONG dwOptionsFlags;     // options flags
    LONG dwCodePage;         // the code used to convert the url
    SECURITY_ATTRIBUTES securityAttributes;
    UUID iid;                // iid of pUnk
    stdole.IUnknown*pUnk;          // pUnk - whatever
    LONG dwReserved;         // should be null
} BINDINFO;


// Options flag
typedef enum BINDINFO_OPTIONS {
    BINDINFO_OPTIONS_WININETFLAG     = 0x00010000,
    BINDINFO_OPTIONS_ENABLE_UTF8     = 0x00020000,
    BINDINFO_OPTIONS_DISABLE_UTF8    = 0x00040000,
    BINDINFO_OPTIONS_USE_IE_ENCODING = 0x00080000,
    BINDINFO_OPTIONS_BINDTOOBJECT    = 0x00100000,
} BINDINFO_OPTIONS;

// Data notification flags
typedef enum BSCF {
    BSCF_FIRSTDATANOTIFICATION        = 0x00000001,
    BSCF_INTERMEDIATEDATANOTIFICATION = 0x00000002,
    BSCF_LASTDATANOTIFICATION         = 0x00000004,
    BSCF_DATAFULLYAVAILABLE           = 0x00000008,
    BSCF_AVAILABLEDATASIZEUNKNOWN     = 0x00000010,
} BSCF;

typedef enum BINDSTATUS {
    BINDSTATUS_FINDINGRESOURCE = 1,
    BINDSTATUS_CONNECTING,
    BINDSTATUS_REDIRECTING,
    BINDSTATUS_BEGINDOWNLOADDATA,
    BINDSTATUS_DOWNLOADINGDATA,
    BINDSTATUS_ENDDOWNLOADDATA,
    BINDSTATUS_BEGINDOWNLOADCOMPONENTS,
    BINDSTATUS_INSTALLINGCOMPONENTS,
    BINDSTATUS_ENDDOWNLOADCOMPONENTS,
    BINDSTATUS_USINGCACHEDCOPY,
    BINDSTATUS_SENDINGREQUEST,
    BINDSTATUS_CLASSIDAVAILABLE,
    BINDSTATUS_MIMETYPEAVAILABLE,
    BINDSTATUS_CACHEFILENAMEAVAILABLE,
    BINDSTATUS_BEGINSYNCOPERATION,
    BINDSTATUS_ENDSYNCOPERATION,
    BINDSTATUS_BEGINUPLOADDATA,
    BINDSTATUS_UPLOADINGDATA,
    BINDSTATUS_ENDUPLOADDATA,
    BINDSTATUS_PROTOCOLCLASSID,
    BINDSTATUS_ENCODING,
    BINDSTATUS_VERIFIEDMIMETYPEAVAILABLE,
    BINDSTATUS_CLASSINSTALLLOCATION,
    BINDSTATUS_DECODING,
    BINDSTATUS_LOADINGMIMEHANDLER,
    BINDSTATUS_CONTENTDISPOSITIONATTACH,
    BINDSTATUS_FILTERREPORTMIMETYPE,
    BINDSTATUS_CLSIDCANINSTANTIATE,
    BINDSTATUS_IUNKNOWNAVAILABLE,
    BINDSTATUS_DIRECTBIND,
    BINDSTATUS_RAWMIMETYPE,
    BINDSTATUS_PROXYDETECTING,
    BINDSTATUS_ACCEPTRANGES,
    BINDSTATUS_COOKIE_SENT,
    BINDSTATUS_COMPACT_POLICY_RECEIVED,
    BINDSTATUS_COOKIE_SUPPRESSED,
    BINDSTATUS_COOKIE_STATE_UNKNOWN,
    BINDSTATUS_COOKIE_STATE_ACCEPT,
    BINDSTATUS_COOKIE_STATE_REJECT,
    BINDSTATUS_COOKIE_STATE_PROMPT,
    BINDSTATUS_COOKIE_STATE_LEASH,
    BINDSTATUS_COOKIE_STATE_DOWNGRADE,
    BINDSTATUS_POLICY_HREF,
    BINDSTATUS_P3P_HEADER,
    BINDSTATUS_SESSION_COOKIE_RECEIVED,
    BINDSTATUS_PERSISTENT_COOKIE_RECEIVED,
    BINDSTATUS_SESSION_COOKIES_ALLOWED,
} BINDSTATUS;

typedef enum ROTFLAGS {
    ROTFLAGS_REGISTRATIONKEEPSALIVE = 0x1,
    ROTFLAGS_ALLOWANYCLIENT = 0x2
} ROTFLAGS;

typedef enum DROPEFFECTS {
    DROPEFFECT_NONE = 0,
    DROPEFFECT_COPY = 1,
    DROPEFFECT_MOVE = 2,
    DROPEFFECT_LINK = 4,
    DROPEFFECT_SCROLL = 0x80000000,
} DROPEFFECTS;

typedef enum CLSCTX {
    CLSCTX_INPROC_SERVER = 0x1,
    CLSCTX_INPROC_HANDLER = 0x2,
    CLSCTX_LOCAL_SERVER = 0x4,
    CLSCTX_INPROC_SERVER16 = 0x8,
    CLSCTX_REMOTE_SERVER = 0x10,
    CLSCTX_INPROC_HANDLER16 = 0x20,
    CLSCTX_INPROC_SERVERX86 = 0x40,
    CLSCTX_INPROC_HANDLERX86 = 0x80,
    CLSCTX_ESERVER_HANDLER = 0x100,
    CLSCTX_RESERVED = 0x200,
    CLSCTX_NO_CODE_DOWNLOAD = 0x400,
    CLSCTX_NO_WX86_TRANSLATION = 0x800,
    CLSCTX_NO_CUSTOM_MARSHAL = 0x1000,
    CLSCTX_ENABLE_CODE_DOWNLOAD = 0x2000,
    CLSCTX_NO_FAILURE_LOG = 0x4000
} CLSCTX;

typedef enum REGCLS {
    REGCLS_SINGLEUSE = 0,       // class object only generates one instance
    REGCLS_MULTIPLEUSE = 1,     // same class object genereates multiple inst.
                                // and local automatically goes into inproc tbl.
    REGCLS_MULTI_SEPARATE = 2,  // multiple use, but separate control over each
                                // context.
    REGCLS_SUSPENDED = 4,  // register is as suspended, will be activated
                                // when app calls CoResumeClassObjects
    REGCLS_SURROGATE = 8   // must be used when a surrogate process
                                // is registering a class object that will be
                                // loaded in the surrogate
} REGCLS;

typedef enum MEMCTX {
    MEMCTX_TASK = 1,
    MEMCTX_SHARED = 2,
    MEMCTX_MACSYSTEM = 3,
    MEMCTX_UNKNOWN = -1,
    MEMCTX_SAME = -2
} MEMCTX;

typedef enum MSHCTX {
    MSHCTX_LOCAL = 0,
    MSHCTX_NOSHAREDMEM = 1,
    MSHCTX_DIFFERENTMACHINE = 2,
    MSHCTX_INPROC = 3
} MSHCTX;

typedef enum MSHLFLAGS {
    MSHLFLAGS_NORMAL = 0,
    MSHLFLAGS_TABLESTRONG = 1,
    MSHLFLAGS_TABLEWEAK = 2,
    MSHLFLAGS_NOPING = 3
} MSHLFLAGS;

typedef enum STDMSHLFLAGS {
    SMEXF_SERVER = 0x01,  //server side aggregated std marshaler
    SMEXF_HANDLER = 0x02  //client side (handler) agg std marshaler
} STDMSHLFLAGS;

typedef enum OLERENDER {
    OLERENDER_NONE = 0,
    OLERENDER_DRAW = 1,
    OLERENDER_FORMAT = 2,
    OLERENDER_ASIS = 3
} OLERENDER;

typedef enum OLECREATE_UDT {
    OLECREATE_LEAVERUNNING = 1,
} OLECREATE_UDT;

typedef struct POINT {
    LONG x;
    LONG y;
} POINT;


typedef struct DROPFILES {
    LONG pFiles; // offset of file list
    POINT pt;    // drop point (client coords)
    BOOL fNC;    // is it on NonClient area
                 // and pt is in screen coords
    BOOL fWide;  // WIDE character switch
} DROPFILES;

typedef struct  NETRESOURCE {
    LONG dwScope;
    LONG dwType;
    LONG dwDisplayType;
    LONG dwUsage;
    LPSTR lpLocalName;
    LPSTR lpRemoteName;
    LPSTR lpComment;
    LPSTR lpProvider;
}NETRESOURCE;

//
// format of CF_NETRESOURCE
//
typedef struct NRESARRAY {
    LONG cItems;
    NETRESOURCE nr[1];
} NRESARRAY;

//
// format of CF_IDLIST
//
typedef struct CIDA {
    LONG cidl;          // number of relative IDList
    LONG aoffset[1];    // [0]: folder IDList, [1]-[cidl]: item IDList
} CIDA;

//
// FILEDESCRIPTOR.dwFlags field indicate which fields are to be used
//
typedef enum FD_FLAGS {
    FD_CLSID = 0x0001,
    FD_SIZEPOINT = 0x0002,
    FD_ATTRIBUTES = 0x0004,
    FD_CREATETIME = 0x0008,
    FD_ACCESSTIME = 0x0010,
    FD_WRITESTIME = 0x0020,
    FD_FILESIZE = 0x0040,
    FD_PROGRESSUI = 0x4000,       // Show Progress UI w/Drag and Drop
    FD_LINKUI = 0x8000,       // 'link' UI is prefered
} FD_FLAGS;

typedef struct _FILEDESCRIPTORA {
    FD_FLAGS dwFlags;
    UUID clsid;
    SIZE sizel;
    POINT pointl;
    LONG dwFileAttributes;
    CURRENCY ftCreationTime;
    CURRENCY ftLastAccessTime;
    CURRENCY ftLastWriteTime;
    LONG nFileSizeHigh;
    LONG nFileSizeLow;
    BYTE cFileName[260];
} FILEDESCRIPTORA;

typedef struct _FILEDESCRIPTORW {
    FD_FLAGS dwFlags;
    UUID clsid;
    SIZE sizel;
    POINT pointl;
    LONG dwFileAttributes;
    CURRENCY ftCreationTime;
    CURRENCY ftLastAccessTime;
    CURRENCY ftLastWriteTime;
    LONG nFileSizeHigh;
    LONG nFileSizeLow;
    BYTE cFileName[520];
} FILEDESCRIPTORW;

typedef struct FILEGROUPDESCRIPTORA {
     LONG cItems;
     FILEDESCRIPTORA fgd[1];
} FILEGROUPDESCRIPTORA;

typedef struct FILEGROUPDESCRIPTORW {
     LONG cItems;
     FILEDESCRIPTORW fgd[1];
} FILEGROUPDESCRIPTORW;


    typedef enum WBIDM {

        IDM_UNKNOWN = 0,
        IDM_ALIGNBOTTOM = 1,
        IDM_ALIGNHORIZONTALCENTERS = 2,
        IDM_ALIGNLEFT = 3,
        IDM_ALIGNRIGHT = 4,
        IDM_ALIGNTOGRID = 5,
        IDM_ALIGNTOP = 6,
        IDM_ALIGNVERTICALCENTERS = 7,
        IDM_ARRANGEBOTTOM = 8,
        IDM_ARRANGERIGHT = 9,
        IDM_BRINGFORWARD = 10,
        IDM_BRINGTOFRONT = 11,
        IDM_CENTERHORIZONTALLY = 12,
        IDM_CENTERVERTICALLY = 13,
        IDM_CODE = 14,
        IDM_DELETE = 17,
        IDM_FONTNAME = 18,
        IDM_FONTSIZE = 19,
        IDM_GROUP = 20,
        IDM_HORIZSPACECONCATENATE = 21,
        IDM_HORIZSPACEDECREASE = 22,
        IDM_HORIZSPACEINCREASE = 23,
        IDM_HORIZSPACEMAKEEQUAL = 24,
        IDM_INSERTOBJECT = 25,
        IDM_MULTILEVELREDO = 30,
        IDM_SENDBACKWARD = 32,
        IDM_SENDTOBACK = 33,
        IDM_SHOWTABLE = 34,
        IDM_SIZETOCONTROL = 35,
        IDM_SIZETOCONTROLHEIGHT = 36,
        IDM_SIZETOCONTROLWIDTH = 37,
        IDM_SIZETOFIT = 38,
        IDM_SIZETOGRID = 39,
        IDM_SNAPTOGRID = 40,
        IDM_TABORDER = 41,
        IDM_TOOLBOX = 42,
        IDM_MULTILEVELUNDO = 44,
        IDM_UNGROUP = 45,
        IDM_VERTSPACECONCATENATE = 46,
        IDM_VERTSPACEDECREASE = 47,
        IDM_VERTSPACEINCREASE = 48,
        IDM_VERTSPACEMAKEEQUAL = 49,
        IDM_JUSTIFYFULL = 50,
        IDM_BACKCOLOR = 51,
        IDM_BOLD = 52,
        IDM_BORDERCOLOR = 53,
        IDM_FLAT = 54,
        IDM_FORECOLOR = 55,
        IDM_ITALIC = 56,
        IDM_JUSTIFYCENTER = 57,
        IDM_JUSTIFYGENERAL = 58,
        IDM_JUSTIFYLEFT = 59,
        IDM_JUSTIFYRIGHT = 60,
        IDM_RAISED = 61,
        IDM_SUNKEN = 62,
        IDM_UNDERLINE = 63,
        IDM_CHISELED = 64,
        IDM_ETCHED = 65,
        IDM_SHADOWED = 66,
        IDM_FIND = 67,
        IDM_SHOWGRID = 69,
        IDM_OBJECTVERBLIST0 = 72,
        IDM_OBJECTVERBLIST1 = 73,
        IDM_OBJECTVERBLIST2 = 74,
        IDM_OBJECTVERBLIST3 = 75,
        IDM_OBJECTVERBLIST4 = 76,
        IDM_OBJECTVERBLIST5 = 77,
        IDM_OBJECTVERBLIST6 = 78,
        IDM_OBJECTVERBLIST7 = 79,
        IDM_OBJECTVERBLIST8 = 80,
        IDM_OBJECTVERBLIST9 = 81,
        IDM_CONVERTOBJECT = 82,
        IDM_CUSTOMCONTROL = 83,
        IDM_CUSTOMIZEITEM = 84,
        IDM_RENAME = 85,
        IDM_IMPORT = 86,
        IDM_NEWPAGE = 87,
        IDM_MOVE = 88,
        IDM_CANCEL = 89,
        IDM_FONT = 90,
        IDM_STRIKETHROUGH = 91,
        IDM_DELETEWORD = 92,
        IDM_EXECPRINT = 93,
        IDM_JUSTIFYNONE = 94,

        IDM_FOLLOW_ANCHOR = 2008,

        IDM_INSINPUTIMAGE = 2114,
        IDM_INSINPUTBUTTON = 2115,
        IDM_INSINPUTRESET = 2116,
        IDM_INSINPUTSUBMIT = 2117,
        IDM_INSINPUTUPLOAD = 2118,
        IDM_INSFIELDSET = 2119,

        IDM_PASTEINSERT = 2120,
        IDM_REPLACE = 2121,
        IDM_EDITSOURCE = 2122,
        IDM_BOOKMARK = 2123,
        IDM_HYPERLINK = 2124,
        IDM_UNLINK = 2125,
        IDM_BROWSEMODE = 2126,
        IDM_EDITMODE = 2127,
        IDM_UNBOOKMARK = 2128,

        IDM_TOOLBARS = 2130,
        IDM_STATUSBAR = 2131,
        IDM_FORMATMARK = 2132,
        IDM_TEXTONLY = 2133,
        IDM_OPTIONS = 2135,
        IDM_FOLLOWLINKC = 2136,
        IDM_FOLLOWLINKN = 2137,
        IDM_VIEWSOURCE = 2139,
        IDM_ZOOMPOPUP = 2140,

        // IDM_BASELINEFONT1, IDM_BASELINEFONT2, IDM_BASELINEFONT3, IDM_BASELINEFONT4,
        // and IDM_BASELINEFONT5 should be consecutive integers;

        IDM_BASELINEFONT1 = 2141,
        IDM_BASELINEFONT2 = 2142,
        IDM_BASELINEFONT3 = 2143,
        IDM_BASELINEFONT4 = 2144,
        IDM_BASELINEFONT5 = 2145,

        IDM_HORIZONTALLINE = 2150,
        IDM_LINEBREAKNORMAL = 2151,
        IDM_LINEBREAKLEFT = 2152,
        IDM_LINEBREAKRIGHT = 2153,
        IDM_LINEBREAKBOTH = 2154,
        IDM_NONBREAK = 2155,
        IDM_SPECIALCHAR = 2156,
        IDM_HTMLSOURCE = 2157,
        IDM_IFRAME = 2158,
        IDM_HTMLCONTAIN = 2159,
        IDM_TEXTBOX = 2161,
        IDM_TEXTAREA = 2162,
        IDM_CHECKBOX = 2163,
        IDM_RADIOBUTTON = 2164,
        IDM_DROPDOWNBOX = 2165,
        IDM_LISTBOX = 2166,
        IDM_BUTTON = 2167,
        IDM_IMAGE = 2168,
        IDM_OBJECT = 2169,
        IDM_1D = 2170,
        IDM_IMAGEMAP = 2171,
        IDM_FILE = 2172,
        IDM_COMMENT = 2173,
        IDM_SCRIPT = 2174,
        IDM_JAVAAPPLET = 2175,
        IDM_PLUGIN = 2176,
        IDM_PAGEBREAK = 2177,
        IDM_HTMLAREA = 2178,

        IDM_PARAGRAPH = 2180,
        IDM_FORM = 2181,
        IDM_MARQUEE = 2182,
        IDM_LIST = 2183,
        IDM_ORDERLIST = 2184,
        IDM_UNORDERLIST = 2185,
        IDM_INDENT = 2186,
        IDM_OUTDENT = 2187,
        IDM_PREFORMATTED = 2188,
        IDM_ADDRESS = 2189,
        IDM_BLINK = 2190,
        IDM_DIV = 2191,

        IDM_TABLEINSERT = 2200,
        IDM_RCINSERT = 2201,
        IDM_CELLINSERT = 2202,
        IDM_CAPTIONINSERT = 2203,
        IDM_CELLMERGE = 2204,
        IDM_CELLSPLIT = 2205,
        IDM_CELLSELECT = 2206,
        IDM_ROWSELECT = 2207,
        IDM_COLUMNSELECT = 2208,
        IDM_TABLESELECT = 2209,
        IDM_TABLEPROPERTIES = 2210,
        IDM_CELLPROPERTIES = 2211,
        IDM_ROWINSERT = 2212,
        IDM_COLUMNINSERT = 2213,

        IDM_HELP_CONTENT = 2220,
        IDM_HELP_ABOUT = 2221,
        IDM_HELP_README = 2222,

        IDM_REMOVEFORMAT = 2230,
        IDM_PAGEINFO = 2231,
        IDM_TELETYPE = 2232,
        IDM_GETBLOCKFMTS = 2233,
        IDM_BLOCKFMT = 2234,
        IDM_SHOWHIDE_CODE = 2235,
        IDM_TABLE = 2236,

        IDM_COPYFORMAT = 2237,
        IDM_PASTEFORMAT = 2238,
        IDM_GOTO = 2239,

        IDM_CHANGEFONT = 2240,
        IDM_CHANGEFONTSIZE = 2241,
        IDM_INCFONTSIZE = 2242,
        IDM_DECFONTSIZE = 2243,
        IDM_INCFONTSIZE1PT = 2244,
        IDM_DECFONTSIZE1PT = 2245,
        IDM_CHANGECASE = 2246,
        IDM_SUBSCRIPT = 2247,
        IDM_SUPERSCRIPT = 2248,
        IDM_SHOWSPECIALCHAR = 2249,

        IDM_CENTERALIGNPARA = 2250,
        IDM_LEFTALIGNPARA = 2251,
        IDM_RIGHTALIGNPARA = 2252,
        IDM_REMOVEPARAFORMAT = 2253,
        IDM_APPLYNORMAL = 2254,
        IDM_APPLYHEADING1 = 2255,
        IDM_APPLYHEADING2 = 2256,
        IDM_APPLYHEADING3 = 2257,

        IDM_DOCPROPERTIES = 2260,
        IDM_ADDFAVORITES = 2261,
        IDM_COPYSHORTCUT = 2262,
        IDM_SAVEBACKGROUND = 2263,
        IDM_SETWALLPAPER = 2264,
        IDM_COPYBACKGROUND = 2265,
        IDM_CREATESHORTCUT = 2266,
        IDM_PAGE = 2267,
        IDM_SAVETARGET = 2268,
        IDM_SHOWPICTURE = 2269,
        IDM_SAVEPICTURE = 2270,
        IDM_DYNSRCPLAY = 2271,
        IDM_DYNSRCSTOP = 2272,
        IDM_PRINTTARGET = 2273,
        IDM_IMGARTPLAY = 2274,
        IDM_IMGARTSTOP = 2275,
        IDM_IMGARTREWIND = 2276,
        IDM_PRINTQUERYJOBSPENDING = 2277,
        IDM_SETDESKTOPITEM = 2278,

        IDM_CONTEXTMENU = 2280,
        IDM_GOBACKWARD = 2282,
        IDM_GOFORWARD = 2283,
        IDM_PRESTOP = 2284,

        IDM_CREATELINK = 2290,
        IDM_COPYCONTENT = 2291,

        IDM_LANGUAGE = 2292,

        IDM_REFRESH = 2300,
        IDM_STOPDOWNLOAD = 2301,

        IDM_ENABLE_INTERACTION = 2302,

        IDM_LAUNCHDEBUGGER = 2310,
        IDM_BREAKATNEXT = 2311,

        IDM_INSINPUTHIDDEN = 2312,
        IDM_INSINPUTPASSWORD = 2313,

        IDM_OVERWRITE = 2314,

        IDM_PARSECOMPLETE = 2315,

        IDM_HTMLEDITMODE = 2316,

        IDM_REGISTRYREFRESH = 2317,
        IDM_COMPOSESETTINGS = 2318,

        IDM_SHOWALLTAGS = 2327,
        IDM_SHOWALIGNEDSITETAGS = 2321,
        IDM_SHOWSCRIPTTAGS = 2322,
        IDM_SHOWSTYLETAGS = 2323,
        IDM_SHOWCOMMENTTAGS = 2324,
        IDM_SHOWAREATAGS = 2325,
        IDM_SHOWUNKNOWNTAGS = 2326,
        IDM_SHOWMISCTAGS = 2320,
        IDM_SHOWZEROBORDERATDESIGNTIME = 2328,

        IDM_AUTODETECT = 2329,

        IDM_SCRIPTDEBUGGER = 2330,

        IDM_GETBYTESDOWNLOADED = 2331,

        IDM_NOACTIVATENORMALOLECONTROLS = 2332,
        IDM_NOACTIVATEDESIGNTIMECONTROLS = 2333,
        IDM_NOACTIVATEJAVAAPPLETS = 2334,
        IDM_NOFIXUPURLSONPASTE = 2335,

        IDM_EMPTYGLYPHTABLE = 2336,
        IDM_ADDTOGLYPHTABLE = 2337,
        IDM_REMOVEFROMGLYPHTABLE = 2338,
        IDM_REPLACEGLYPHCONTENTS = 2339,

        IDM_SHOWWBRTAGS = 2340,

        IDM_PERSISTSTREAMSYNC = 2341,
        IDM_SETDIRTY = 2342,

        IDM_RUNURLSCRIPT = 2343,

        IDM_ZOOMRATIO = 2344,
        IDM_GETZOOMNUMERATOR = 2345,
        IDM_GETZOOMDENOMINATOR = 2346,

        // COMMANDS FOR COMPLEX TEXT
        IDM_DIRLTR = 2350,
        IDM_DIRRTL = 2351,
        IDM_BLOCKDIRLTR = 2352,
        IDM_BLOCKDIRRTL = 2353,
        IDM_INLINEDIRLTR = 2354,
        IDM_INLINEDIRRTL = 2355,

        // SHDOCVW
        IDM_ISTRUSTEDDLG = 2356,

        // MSHTMLED
        IDM_INSERTSPAN = 2357,
        IDM_LOCALIZEEDITOR = 2358,

        // XML MIMEVIEWER
        IDM_SAVEPRETRANSFORMSOURCE = 2370,
        IDM_VIEWPRETRANSFORMSOURCE = 2371,

        IDM_MIMECSET__FIRST__ = 3609,
        IDM_MIMECSET__LAST__ = 3699,

        IDM_MENUEXT_FIRST__ = 3700,
        IDM_MENUEXT_LAST__ = 3732,
        IDM_MENUEXT_COUNT = 3733,

        // Commands mapped from the standard set.  We should
        // consider deleting them from public header files.

        IDM_OPEN = 2000,
        IDM_NEW = 2001,
        IDM_SAVE = 70,
        IDM_SAVEAS = 71,
        IDM_SAVECOPYAS = 2002,
        IDM_PRINTPREVIEW = 2003,
        IDM_PRINT = 27,
        IDM_PAGESETUP = 2004,
        IDM_SPELL = 2005,
        IDM_PASTESPECIAL = 2006,
        IDM_CLEARSELECTION = 2007,
        IDM_PROPERTIES = 28,
        IDM_REDO = 29,
        IDM_UNDO = 43,
        IDM_SELECTALL = 31,
        IDM_ZOOMPERCENT = 50,
        IDM_GETZOOM = 68,
        IDM_STOP = 2138,
        IDM_COPY = 15,
        IDM_CUT = 16,
        IDM_PASTE = 26,

        // Defines for IDM_ZOOMPERCENT
        CMD_ZOOM_PAGEWIDTH = -1,
        CMD_ZOOM_ONEPAGE = -2,
        CMD_ZOOM_TWOPAGES = -3,
        CMD_ZOOM_SELECTION = -4,
        CMD_ZOOM_FIT = -5,

    } WBIDM;

    typedef enum ContextMenuTarget {
        CONTEXT_MENU_DEFAULT = 0,
        CONTEXT_MENU_IMAGE = 1,
        CONTEXT_MENU_CONTROL = 2,
        CONTEXT_MENU_TABLE = 3,
        CONTEXT_MENU_SELECTION = 4,
        CONTEXT_MENU_LINK = 5,
        CONTEXT_MENU_ANCHOR = 6,
        CONTEXT_MENU_IMGDYNSRC = 7,
        CONTEXT_MENU_IMGART = 8,
    } ContextMenuTarget;

    typedef enum DOCHOSTUIDBLCLK {
        DOCHOSTUIDBLCLK_DEFAULT = 0,
        DOCHOSTUIDBLCLK_SHOWPROPERTIES = 1,
        DOCHOSTUIDBLCLK_SHOWCODE = 2,
    } DOCHOSTUIDBLCLK ;

    typedef enum DOCHOSTUIFLAG {
        DOCHOSTUIFLAG_DIALOG = 1,
        DOCHOSTUIFLAG_DISABLE_HELP_MENU = 2,
        DOCHOSTUIFLAG_NO3DBORDER = 4,
        DOCHOSTUIFLAG_SCROLL_NO = 8,
        DOCHOSTUIFLAG_DISABLE_SCRIPT_INACTIVE = 0x10,
        DOCHOSTUIFLAG_OPENNEWUI = 0x20,
        DOCHOSTUIFLAG_DISABLE_OFFSCREEN = 0x40,
        DOCHOSTUIFLAG_FLAT_SCROLLBAR = 0x80,
        DOCHOSTUIFLAG_DIV_BLOCKDEFAULT = 0x0100,
        DOCHOSTUIFLAG_ACTIVATE_CLIENTHIT_ONLY = 0x0200,
        DOCHOSTUIFLAG_OVERRIDEBEHAVIORFACTORY = 0x0400,
        DOCHOSTUIFLAG_CODEPAGELINKEDFONTS = 0x0800,
        DOCHOSTUIFLAG_URL_ENCODING_DISABLE_UTF8 = 0x1000,
        DOCHOSTUIFLAG_URL_ENCODING_ENABLE_UTF8 = 0x2000,
        DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE = 0x4000,
        DOCHOSTUIFLAG_ENABLE_INPLACE_NAVIGATION = 0x10000,
        DOCHOSTUIFLAG_IME_ENABLE_RECONVERSION = 0x20000,
        DOCHOSTUIFLAG_BROWSER = 0x12,
        DOCHOSTUIFLAG_DESKTOP = 0x2E,
   } DOCHOSTUIFLAG ;

    typedef struct DOCHOSTUIINFO {
        LONG cbSize;
        DOCHOSTUIFLAG dwFlags;
        DOCHOSTUIDBLCLK dwDoubleClick;
        LONG pchHostCss;
        LONG pchHostNS;
    } DOCHOSTUIINFO;

    typedef [helpstring("DISP_E_EXCEPTION Error info")] struct EXCEPINFO {
        [helpstring("Error number")]
        short wCode;
        short wReserved;
        [helpstring("Source")]
        BSTR Source;
        [helpstring("Description")]
        BSTR Description;
        [helpstring("Help file")]
        BSTR HelpFile;
        [helpstring("Help context")]
        LONG dwHelpContext;
        LONG pvReserved;
        [helpstring("HRESULT (__stdcall *pfnDeferredFillIn)(struct EXCEPINFO *);")]
        LONG pfnDeferredFillIn;
        LONG scode;
    } EXCEPINFO;

    typedef struct DISPPARAMS {
        [helpstring("Parameters")]
        LONG rgPointerToVariantArray;
        [helpstring("Named parameters")]
        LONG rgPointerToLONGNamedArgs;
        [helpstring("Parameter count")]
        LONG cArgs;
        [helpstring("Named parameter count")]
        LONG cNamedArgs;
    } DISPPARAMS;

    typedef [helpstring("InvokeFlags")] enum InvokeFlags {
        [helpstring("Sub/Function")]
        DISPATCH_METHOD = 0x1,
        [helpstring("Property Get")]
        DISPATCH_PROPERTYGET = 0x2,
        [helpstring("Property Let")]
        DISPATCH_PROPERTYPUT = 0x4,
        [helpstring("Property Set")]
        DISPATCH_PROPERTYPUTREF = 0x8
    } InvokeFlags;

    typedef [helpstring("DEFAULT_DISPIDS")] enum DEFAULT_DISPIDS {
        DISPID_UNKNOWN = -1,
        DISPID_VALUE = 0,
        DISPID_PROPERTYPUT = -3,
        DISPID_NEWENUM = -4,
        DISPID_EVALUATE = -5,
        DISPID_CONSTRUCTOR = -6,
        DISPID_DESTRUCTOR = -7,
        DISPID_COLLECT = -8,
    } DEFAULT_DISPIDS;

    typedef [helpstring("IDispatch::Invoke results")] enum InvokeResults {
        [helpstring("Unknown Interface")]
        DISP_E_UNKNOWNINTERFACE = 0x80020001,
        [helpstring("Member not found")]
        DISP_E_MEMBERNOTFOUND = 0x80020003,
        [helpstring("Parameter not found")]
        DISP_E_PARAMNOTFOUND = 0x80020004,
        [helpstring("Type Mismatch")]
        DISP_E_TYPEMISMATCH = 0x80020005,
        [helpstring("Unknown Name")]
        DISP_E_UNKNOWNNAME = 0x80020006,
        [helpstring("No named arguments")]
        DISP_E_NONAMEDARGS = 0x80020007,
        [helpstring("Bad variable type")]
        DISP_E_BADVARTYPE = 0x80020008,
        [helpstring("Exception")]
        DISP_E_EXCEPTION = 0x80020009,
        [helpstring("Overflow")]
        DISP_E_OVERFLOW = 0x8002000A,
        [helpstring("Bad Index")]
        DISP_E_BADINDEX = 0x8002000B,
        [helpstring("Unknown LCID")]
        DISP_E_UNKNOWNLCID = 0x8002000C,
        [helpstring("Memory is locked")]
        DISP_E_ARRAYISLOCKED = 0x8002000D,
        [helpstring("Invalid number of parameters")]
        DISP_E_BADPARAMCOUNT = 0x8002000E,
        [helpstring("Parameter not optional")]
        DISP_E_PARAMNOTOPTIONAL = 0x8002000F,
        [helpstring("Invalid callee")]
        DISP_E_BADCALLEE = 0x80020010,
        [helpstring("Does not support a collection")]
        DISP_E_NOTACOLLECTION = 0x80020011,
        [helpstring("Division by zero")]
        DISP_E_DIVBYZERO = 0x80020012
    } InvokeResults;


    typedef enum DISPID_EXTENDER {
        DISPID_EXTENDER_NAME = 0x80010000,
        DISPID_EXTENDER_INDEX = 0x80010001,
        DISPID_EXTENDER_LEFT = 0x80010003,
        DISPID_EXTENDER_TOP = 0x80010004,
        DISPID_EXTENDER_WIDTH = 0x80010005,
        DISPID_EXTENDER_HEIGHT = 0x80010006,
        DISPID_EXTENDER_VISIBLE = 0x80010007,
        DISPID_EXTENDER_PARENT = 0x80010008,
        DISPID_EXTENDER_DRAGMODE = 0x80010009,
        DISPID_EXTENDER_DRAGICON = 0x8001000A,
        DISPID_EXTENDER_TAG = 0x8001000B,
        DISPID_EXTENDER_TABSTOP = 0x8001000E,
        DISPID_EXTENDER_TABINDEX = 0x8001000F,
        DISPID_EXTENDER_OBJECT = 0x8001001C,
        DISPID_EXTENDER_HELPCONTEXTID = 0x80010032,
        DISPID_EXTENDER_DEFAULT = 0x80010037,
        DISPID_EXTENDER_CANCEL = 0x80010038,
        DISPID_EXTENDER_ALIGN = 0x8001003B,
        DISPID_EXTENDER_NEGOTIATE = 0x8001003C,
        DISPID_EXTENDER_WHATSTHISHELP = 0x80010046,
        DISPID_EXTENDER_CONTAINER = 0x80010047,
        DISPID_EXTENDER_TOOLTIPTEXT = 0x8001004A,

        DISPID_EXTENDER_SETFOCUS = 0x80011000,
        DISPID_EXTENDER_ZORDER = 0x80011001,
        DISPID_EXTENDER_MOVE = 0x80011002,
        DISPID_EXTENDER_DRAG = 0x80011003,
        DISPID_EXTENDER_SHOWWHATSTHIS = 0x80011004,
    } DISPID_EXTENDER;

    typedef enum TYPEKIND {
        TKIND_ENUM = 0,
        TKIND_RECORD,
        TKIND_MODULE,
        TKIND_INTERFACE,
        TKIND_DISPATCH,
        TKIND_COCLASS,
        TKIND_ALIAS,
        TKIND_UNION,
        TKIND_MAX
    } TYPEKIND;

    typedef enum SYSKIND {
        SYS_WIN16 = 0,
        SYS_WIN32,
        SYS_MAC
    } SYSKIND;

    typedef enum LIBFLAGS {
        LIBFLAG_FRESTRICTED = 0x01,
        LIBFLAG_FCONTROL = 0x02,
        LIBFLAG_FHIDDEN = 0x04,
        LIBFLAG_FHASDISKIMAGE = 0x08
    } LIBFLAGS;

    typedef struct TLIBATTR {
        UUID IID;
        LONG lcid;
        SYSKIND syskind;
        short wMajorVerNum;
        short wMinorVerNum;
        short wLibFlags;
    } TLIBATTR;

    typedef struct TYPEDESC {
        LONG pTypeDesc;
        short vt;
    } TYPEDESC;

    typedef struct SAFEARRAYBOUND {
        LONG cElements;
        LONG lLbound;
    } SAFEARRAYBOUND;

    typedef struct ARRAYDESC {
        TYPEDESC tdescElem;
        short cDims;
    } ARRAYDESC;

    typedef struct PARAMDESCEX {
        LONG cBytes;
    LONG ignore;
        VARIANT varDefaultValue;
    } PARAMDESCEX;

    typedef struct PARAMDESC {
        [helpstring("valid if PARAMFLAG_FHASDEFAULT bit is set")]
        LONG pPARAMDESCEX;
        short wParamFlags;
    } PARAMDESC;

    typedef enum PARAMFLAGS {
       PARAMFLAG_NONE = 0x00,
       PARAMFLAG_FIN = 0x01,
       PARAMFLAG_FOUT = 0x02,
       PARAMFLAG_FLONG = 0x04,
       PARAMFLAG_FRETVAL = 0x08,
       PARAMFLAG_FOPT = 0x10,
       PARAMFLAG_FHASDEFAULT = 0x20,
    } PARAMFLAGS;

    typedef struct ELEMDESC {
        TYPEDESC tdesc;
        PARAMDESC paramdesc;
    } ELEMDESC;

    typedef struct IDLDESC {
        LONG dwReserved;
        short wIDLFlags;
    } IDLDESC;

    typedef struct TYPEATTR {
        [helpstring("the IID of the TypeInfo")]
        UUID IID;
        [helpstring("locale of member names and doc strings")]
        LONG lcid;
        LONG dwReserved;
        [helpstring("ID of constructor, LONG_NIL if none")]
        LONG memidConstructor;
        [helpstring("ID of destructor, LONG_NIL if none")]
        LONG memidDestructor;
        LONG lpstrSchema;
        [helpstring("the size of an instance of this type")]
        LONG cbSizeInstance;
        [helpstring("the kind of type this typeinfo describes")]
        TYPEKIND typekind;
        [helpstring("number of functions")]
        short cFuncs;
        [helpstring("number of variables / data members")]
        short cVars;
        [helpstring("number of implemented interfaces")]
        short cImplTypes;
        [helpstring("the size of this types virtual func table")]
        short cbSizeVft;
        [helpstring("specifies the alignment requirements for an instance of this type,\n0 = align on 64k boundary\n1 = byte align\n2 = word align\n4 = dword align")]
        short cbAlignment;
        short wTypeFlags;
        [helpstring("major version number")]
        short wMajorVerNum;
        [helpstring("minor version number")]
        short wMinorVerNum;
        [helpstring("if typekind = = TKIND_ALIAS this field specifies the type for which this type is an alias")]
        TYPEDESC tdescAlias;
        [helpstring("IDL attributes of the described type")]
        IDLDESC idldescType;
    } TYPEATTR;

typedef enum CALLCONV {
    CC_FASTCALL = 0,
    CC_CDECL = 1,
    CC_MSCPASCAL,
    CC_PASCAL = 2,
    CC_MACPASCAL,
    CC_STDCALL,
    CC_FPFASTCALL,
    CC_SYSCALL,
    CC_MPWCDECL,
    CC_MPWPASCAL,
    CC_MAX
} CALLCONV;

typedef enum FUNCKIND {
    FUNC_VIRTUAL,
    FUNC_PUREVIRTUAL,
    FUNC_NONVIRTUAL,
    FUNC_STATIC,
    FUNC_DISPATCH
} FUNCKIND;

typedef enum INVOKEKIND {
    INVOKE_FUNC = 1,
    INVOKE_PROPERTYGET = 2,
    INVOKE_PROPERTYPUT = 4,
    INVOKE_PROPERTYPUTREF = 8
} INVOKEKIND;

typedef struct FUNCDESC {
    LONG memid;
    LONG lprgscode;
     [helpstring("array of param types")]
    LONG lprgELEMDESCParam;
    FUNCKIND funckind;
    INVOKEKIND invkind;
    CALLCONV callconv;
    short cParams;
    short cParamsOpt;
    short oVft;
    short cScodes;
    ELEMDESC elemdescFunc;
    short wFuncFlags;
} FUNCDESC;

typedef enum VARKIND {
    VAR_PERINSTANCE,
    VAR_STATIC,
    VAR_CONST,
    VAR_DISPATCH
} VARKIND;

typedef enum IMPLTYPE {
    IMPLTYPEFLAG_FDEFAULT = 0x1,
    IMPLTYPEFLAG_FSOURCE = 0x2,
    IMPLTYPEFLAG_FRESTRICTED = 0x4,
    IMPLTYPEFLAG_FDEFAULTVTABLE= 0x8,
} IMPLTYPE;

typedef struct VARDESC {
    LONG memid;
    LONG lpstrSchema;
    LONG oInst_varValue;
    ELEMDESC elemdescVar;
    short wVarFlags;
    VARKIND varkind;
} VARDESC;

typedef enum TYPEFLAGS {
    TYPEFLAG_FAPPOBJECT = 0x01,
    TYPEFLAG_FCANCREATE = 0x02,
    TYPEFLAG_FLICENSED = 0x04,
    TYPEFLAG_FPREDECLID = 0x08,
    TYPEFLAG_FHIDDEN = 0x10,
    TYPEFLAG_FCONTROL = 0x20,
    TYPEFLAG_FDUAL = 0x40,
    TYPEFLAG_FNONEXTENSIBLE = 0x80,
    TYPEFLAG_FOLEAUTOMATION = 0x100,
    TYPEFLAG_FRESTRICTED = 0x200,
    TYPEFLAG_FAGGREGATABLE = 0x400,
    TYPEFLAG_FREPLACEABLE = 0x800,
    TYPEFLAG_FDISPATCHABLE = 0x1000,
    TYPEFLAG_FREVERSEBIND = 0x2000
} TYPEFLAGS;

typedef enum FUNCFLAGS {
    FUNCFLAG_FRESTRICTED = 0x1,
    FUNCFLAG_FSOURCE = 0x2,
    FUNCFLAG_FBINDABLE = 0x4,
    FUNCFLAG_FREQUESTEDIT = 0x8,
    FUNCFLAG_FDISPLAYBIND = 0x10,
    FUNCFLAG_FDEFAULTBIND = 0x20,
    FUNCFLAG_FHIDDEN = 0x40,
    FUNCFLAG_FUSESGETLASTERROR = 0x80,
    FUNCFLAG_FDEFAULTCOLLELEM = 0x100,
    FUNCFLAG_FUIDEFAULT = 0x200,
    FUNCFLAG_FNONBROWSABLE = 0x400,
    FUNCFLAG_FREPLACEABLE = 0x800,
    FUNCFLAG_FIMMEDIATEBIND = 0x1000
} FUNCFLAGS;

typedef enum VARFLAGS {
    VARFLAG_FREADONLY = 0x1,
    VARFLAG_FSOURCE = 0x2,
    VARFLAG_FBINDABLE = 0x4,
    VARFLAG_FREQUESTEDIT = 0x8,
    VARFLAG_FDISPLAYBIND = 0x10,
    VARFLAG_FDEFAULTBIND = 0x20,
    VARFLAG_FHIDDEN = 0x40,
    VARFLAG_FRESTRICTED = 0x80,
    VARFLAG_FDEFAULTCOLLELEM = 0x100,
    VARFLAG_FUIDEFAULT = 0x200,
    VARFLAG_FNONBROWSABLE = 0x400,
    VARFLAG_FREPLACEABLE = 0x800,
    VARFLAG_FIMMEDIATEBIND = 0x1000
} VARFLAGS;

typedef struct CUSTDATAITEM {
    UUID IID;
    VARIANT varValue;
} CUSTDATAITEM;

typedef struct CUSTDATA {
     [helpstring("number of custom data items in rgCustData")]
    LONG cCustData;
    LONG prgCustData;
} CUSTDATA;

typedef [helpstring("Standard DISPIDs")] enum STANDARD_DISPIDS {
    DISPID_AUTOSIZE = (-500),
    DISPID_BACKCOLOR = (-501),
    DISPID_BACKSTYLE = (-502),
    DISPID_BORDERCOLOR = (-503),
    DISPID_BORDERSTYLE = (-504),
    DISPID_BORDERWIDTH = (-505),
    DISPID_DRAWMODE = (-507),
    DISPID_DRAWSTYLE = (-508),
    DISPID_DRAWWIDTH = (-509),
    DISPID_FILLCOLOR = (-510),
    DISPID_FILLSTYLE = (-511),
    DISPID_FONT = (-512),
    DISPID_FORECOLOR = (-513),
    DISPID_ENABLED = (-514),
    DISPID_HWND = (-515),
    DISPID_TABSTOP = (-516),
    DISPID_TEXT = (-517),
    DISPID_CAPTION = (-518),
    DISPID_BORDERVISIBLE = (-519),
    DISPID_APPEARANCE = (-520),
    DISPID_MOUSEPOINTER = (-521),
    DISPID_MOUSEICON = (-522),
    DISPID_PICTURE = (-523),
    DISPID_VALID = (-524),
    DISPID_READYSTATE = (-525),
    DISPID_LISTINDEX = (-526),
    DISPID_SELECTED = (-527),
    DISPID_LIST = (-528),
    DISPID_COLUMN = (-529),
    DISPID_LISTCOUNT = (-531),
    DISPID_MULTISELECT = (-532),
    DISPID_MAXLENGTH = (-533),
    DISPID_PASSWORDCHAR = (-534),
    DISPID_SCROLLBARS = (-535),
    DISPID_WORDWRAP = (-536),
    DISPID_MULTILINE = (-537),
    DISPID_NUMBEROFROWS = (-538),
    DISPID_NUMBEROFCOLUMNS = (-539),
    DISPID_DISPLAYSTYLE = (-540),
    DISPID_GROUPNAME = (-541),
    DISPID_IMEMODE = (-542),
    DISPID_ACCELERATOR = (-543),
    DISPID_ENTERKEYBEHAVIOR = (-544),
    DISPID_TABKEYBEHAVIOR = (-545),
    DISPID_SELTEXT = (-546),
    DISPID_SELSTART = (-547),
    DISPID_SELLENGTH = (-548),

    DISPID_REFRESH = (-550),
    DISPID_DOCLICK = (-551),
    DISPID_ABOUTBOX = (-552),
    DISPID_ADDITEM = (-553),
    DISPID_CLEAR = (-554),
    DISPID_REMOVEITEM = (-555),

    DISPID_CLICK = (-600),
    DISPID_DBLCLICK = (-601),
    DISPID_KEYDOWN = (-602),
    DISPID_KEYPRESS = (-603),
    DISPID_KEYUP = (-604),
    DISPID_MOUSEDOWN = (-605),
    DISPID_MOUSEMOVE = (-606),
    DISPID_MOUSEUP = (-607),
    DISPID_ERROREVENT = (-608),
    DISPID_READYSTATECHANGE = (-609),
    DISPID_CLICK_VALUE = (-610),
    DISPID_RIGHTTOLEFT = (-611),
    DISPID_TOPTOBOTTOM = (-612),

    DISPID_NAME = (-800),
    DISPID_DELETE = (-801),
    DISPID_OBJECT = (-802),
    DISPID_PARENT = (-803)
} STANDARD_DISPIDS ;

typedef enum AMBIENT_DISPIDS {
    DISPID_AMBIENT_BACKCOLOR = (-701),
    DISPID_AMBIENT_DISPLAYNAME = (-702),
    DISPID_AMBIENT_FONT = (-703),
    DISPID_AMBIENT_FORECOLOR = (-704),
    DISPID_AMBIENT_LOCALEID = (-705),
    DISPID_AMBIENT_MESSAGEREFLECT = (-706),
    DISPID_AMBIENT_SCALEUNITS = (-707),
    DISPID_AMBIENT_TEXTALIGN = (-708),
    DISPID_AMBIENT_USERMODE = (-709),
    DISPID_AMBIENT_UIDEAD = (-710),
    DISPID_AMBIENT_SHOWGRABHANDLES = (-711),
    DISPID_AMBIENT_SHOWHATCHING = (-712),
    DISPID_AMBIENT_DISPLAYASDEFAULT = (-713),
    DISPID_AMBIENT_SUPPORTSMNEMONICS = (-714),
    DISPID_AMBIENT_AUTOCLIP = (-715),
    DISPID_AMBIENT_APPEARANCE = (-716),

    DISPID_AMBIENT_PALETTE = (-726),
    DISPID_AMBIENT_TRANSFERPRIORITY = (-728),

    DISPID_AMBIENT_RIGHTTOLEFT = (-732),
    DISPID_AMBIENT_TOPTOBOTTOM = (-733)

} AMBIENT_DISPIDS;

typedef enum OLEOBJECTS_DISPIDS {
    DISPID_FONT_NAME = 0,
    DISPID_FONT_SIZE = 2,
    DISPID_FONT_BOLD = 3,
    DISPID_FONT_ITALIC = 4,
    DISPID_FONT_UNDER = 5,
    DISPID_FONT_STRIKE = 6,
    DISPID_FONT_WEIGHT = 7,
    DISPID_FONT_CHARSET = 8,

    DISPID_PICT_HANDLE = 0,
    DISPID_PICT_HPAL = 2,
    DISPID_PICT_TYPE = 3,
    DISPID_PICT_WIDTH = 4,
    DISPID_PICT_HEIGHT = 5,
    DISPID_PICT_RENDER = 6
} OLEOBJECTS_DISPIDS ;

typedef enum REGKIND {
    REGKIND_DEFAULT,
    REGKIND_REGISTER,
    REGKIND_NONE
} REGKIND;

typedef enum PARSEACTION {
    PARSE_CANONICALIZE = 1,
    PARSE_FRIENDLY,
    PARSE_SECURITY_URL,
    PARSE_ROOTDOCUMENT,
    PARSE_DOCUMENT,
    PARSE_ANCHOR,
    PARSE_ENCODE,
    PARSE_DECODE,
    PARSE_PATH_FROM_URL,
    PARSE_URL_FROM_PATH,
    PARSE_MIME,
    PARSE_SERVER,
    PARSE_SCHEMA,
    PARSE_SITE,
    PARSE_DOMAIN,
    PARSE_LOCATION,
    PARSE_SECURITY_DOMAIN,
    PARSE_ESCAPE,
    PARSE_UNESCAPE
} PARSEACTION;

typedef enum PSUACTION {
    PSU_DEFAULT = 1,
    PSU_SECURITY_URL_ONLY
} PSUACTION;

typedef enum QUERYOPTION {
    QUERY_EXPIRATION_DATE = 1,
    QUERY_TIME_OF_LAST_CHANGE,
    QUERY_CONTENT_ENCODING,
    QUERY_CONTENT_TYPE,
    QUERY_REFRESH,
    QUERY_RECOMBINE,
    QUERY_CAN_NAVIGATE,
    QUERY_USES_NETWORK,
    QUERY_IS_CACHED,
    QUERY_IS_INSTALLEDENTRY,
    QUERY_IS_CACHED_OR_MAPPED,
    QUERY_USES_CACHE,
    QUERY_IS_SECURE,
    QUERY_IS_SAFE
} QUERYOPTION;

typedef enum VARFORMAT_FLAGS {
    VAR_TIMEVALUEONLY = 0x00000001,    /* return time value */
    VAR_DATEVALUEONLY = 0x00000002,    /* return date value */
    VAR_VALIDDATE = 0x00000004,
    VAR_CALENDAR_HIJRI = 0x00000008,   /* use Hijri calender */
    VAR_LOCALBOOL = 0x00000010,
    VAR_FORMAT_NOSUBSTITUTE = 0x00000020
} VARFORMAT_FLAGS;

typedef enum NAMEDFORMATS {
    GeneralDate,
    LongDate,
    ShortDate,
    LongTime,
    ShortTime
} NAMEDFORMATS;

typedef enum ACTIVEOBJECTFLAGS {
    ACTIVEOBJECT_STRONG = 0x0,
    ACTIVEOBJECT_WEAK = 0x1
} ACTIVEOBJECTFLAGS;

typedef enum NUMPRS {
    NUMPRS_LEADING_WHITE    = 0x0001,
    NUMPRS_TRAILING_WHITE   = 0x0002,
    NUMPRS_LEADING_PLUS     = 0x0004,
    NUMPRS_TRAILING_PLUS    = 0x0008,
    NUMPRS_LEADING_MINUS    = 0x0010,
    NUMPRS_TRAILING_MINUS   = 0x0020,
    NUMPRS_HEX_OCT          = 0x0040,
    NUMPRS_PARENS           = 0x0080,
    NUMPRS_DECIMAL          = 0x0100,
    NUMPRS_THOUSANDS        = 0x0200,
    NUMPRS_CURRENCY         = 0x0400,
    NUMPRS_EXPONENT         = 0x0800,
    NUMPRS_USE_ALL          = 0x1000,
    NUMPRS_STD              = 0x1FFF,
    NUMPRS_NEG              = 0x10000,
    NUMPRS_INEXACT          = 0x20000,
} NUMPRS;

typedef struct NUMPARSE {
    LONG cDig;
    LONG dwInFlags;
    LONG dwOutFlags;
    LONG cchUsed;
    LONG nBaseShift;
    LONG nPwr10;
} NUMPARSE;

typedef enum NORM {
    NORM_IGNORECASE           = 0x00000001,
    NORM_IGNORENONSPACE       = 0x00000002,
    NORM_IGNORESYMBOLS        = 0x00000004,
    NORM_IGNOREWIDTH          = 0x00000008,
    NORM_IGNOREKANATYPE1      = 0x00000040,
    NORM_IGNOREKANATYPE       = 0x00010000,
    NORM_IGNOREWIDTH1         = 0x00020000,
    NORM_IGNOREKASHIDA        = 0x00040000,
} NORM;

typedef enum VARCMP {
    VARCMP_LT = 0,
    VARCMP_EQ = 1,
    VARCMP_GT = 2,
    VARCMP_NULL = 3 
} VARCMP;

typedef enum VCT_FLAGS {
    VARIANT_NOVALUEPROP = 0x01,
    VARIANT_ALPHABOOL = 0x02,      // For VT_BOOL to VT_BSTR conversions
                                   // convert to "True"/"False" instead of
                                   // "-1"/"0"
    VARIANT_NOUSEROVERRIDE = 0x04, // For conversions to/from VT_BSTR,,
                                   // passes LOCALE_NOUSEROVERRIDE
                                   // to core coercion routines
    VARIANT_CALENDAR_HIJRI = 0x08,
    VARIANT_LOCALBOOL = 0x10,      // For VT_BOOL to VT_BSTR and back
                                   // convert to local language rather than
                                   // English
} VCT_FLAGS;

typedef struct tagOBJECTDESCRIPTOR {
    LONG cbSize;              // Size of structure in bytes
    UUID clsid;               // CLSID of data being transferred
    DVASPECT dwDrawAspect;    // Display aspect of the object
                              //     normally DVASPECT_CONTENT or ICON.
                              //     dwDrawAspect will be 0 (which is NOT
                              //     DVASPECT_CONTENT) if the copier or
                              //     dragsource didn't draw the object to
                              //     begin with.
    SIZE sizel;               // size of the object in HIMETRIC
                              //    sizel is opt.: will be (0,0) for apps
                              //    which don't draw the object being
                              //    transferred
    POINT pointl;             // Offset in HIMETRIC units from the
                              //    upper-left corner of the obj where the
                              //    mouse went down for the drag.
                              //    NOTE: y coordinates increase downward.
                              //          x coordinates increase to right
                              //    pointl is opt.; it is only meaningful
                              //    if object is transfered via drag/drop.
                              //    (0, 0) if mouse position is unspecified
                              //    (eg. when obj transfered via clipboard)
    OLEMISC dwStatus;         // Misc. status flags for object. Flags are
                              //    defined by OLEMISC enum. these flags
                              //    are as would be returned
                              //    by IOleObject::GetMiscStatus.
    LONG dwFullUserTypeName;  // Offset from beginning of structure to
                              //    null-terminated string that specifies
                              //    Full User Type Name of the object.
                              //    0 indicates string not present.
    LONG dwSrcOfCopy;         // Offset from beginning of structure to
                              //    null-terminated string that specifies
                              //    source of the transfer.
                              //    dwSrcOfCOpy is normally implemented as
                              //    the display name of the temp-for-user
                              //    moniker which identifies the source of
                              //    the data.
                              //    0 indicates string not present.
                              //    NOTE: moniker assignment is NOT forced.
                              //    see IOleObject::GetMoniker(
                              //                OLEGETMONIKER_TEMPFORUSER)
} OBJECTDESCRIPTOR;

typedef struct NMHDR {
    LONG hwndFrom;
    LONG idFrom;
    LONG code;
} NMHDR;

