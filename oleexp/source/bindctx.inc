//tbShellLib - 100% implemented

[
    odl,
    uuid(0000000e-0000-0000-C000-000000000046),
]
interface IBindCtx : stdole.IUnknown {

    typedef enum BIND_FLAGS {
        BIND_MAYBOTHERUSER = 1,
        BIND_JUSTTESTEXISTENCE = 2
    } BIND_FLAGS;

    typedef struct BIND_OPTS {
        LONG cbStruct;       //  sizeof(BIND_OPTS)
        BIND_FLAGS grfFlags;
        STGM grfMode;
        LONG dwTickCountDeadline;
    } BIND_OPTS;

    typedef struct BIND_OPTS2 {
        LONG cbStruct;       //  sizeof(BIND_OPTS2)
        BIND_FLAGS grfFlags;
        STGM grfMode;
        LONG dwTickCountDeadline;
        LONG dwTrackFlags;
        LONG dwClassContext;
        LONG locale;
        LONG pServerInfo;
    } BIND_OPTS2;

    HRESULT RegisterObjectBound(
        [in] stdole.IUnknown*punk);

    HRESULT RevokeObjectBound(
        [in] stdole.IUnknown*punk);

    HRESULT ReleaseBoundObjects(void);

    HRESULT SetBindOptions(
        [in] BIND_OPTS *pbindopts);

    HRESULT GetBindOptions(
        [in, out] BIND_OPTS *pbindopts);

    HRESULT GetRunningObjectTable(
        [out, retval] IRunningObjectTable **pprot);

    HRESULT RegisterObjectParam(
        [in] LONG pszKey,
        [in] stdole.IUnknown *punk);

    HRESULT GetObjectParam(
        [in] LPWSTR pszKey,
        [out, retval] stdole.IUnknown**ppunk);

    HRESULT EnumObjectParam(
        [out, retval] IEnumString **ppenum);

    HRESULT RevokeObjectParam(
        [in] LPWSTR pszKey);

}

[
    odl,
    uuid(00000010-0000-0000-C000-000000000046)
]
interface IRunningObjectTable : stdole.IUnknown {

    HRESULT Register(
        [in] ROTFLAGS grfFlags,
        [in] stdole.IUnknown*punkObject,
        [in] IMoniker *pmkObjectName,
        [out, retval] LONG *pdwRegister);

    HRESULT Revoke(
        [in] LONG dwRegister);

    HRESULT IsRunning(
        [in] IMoniker *pmkObjectName);

    HRESULT GetObject(
        [in] IMoniker *pmkObjectName,
        [out, retval] stdole.IUnknown**ppunkObject);

    HRESULT NoteChangeTime(
        [in] LONG dwRegister,
        [in] CURRENCY *pfiletime);

    HRESULT GetTimeOfLastChange(
        [in] IMoniker *pmkObjectName,
        [out, retval] CURRENCY *pfiletime);

    HRESULT EnumRunning(
        [out, retval] IEnumMoniker **ppenumMoniker);

}

    [
        helpstring("BindCtx Strings"),
        dllname("bcdummy")
    ]
    module BindCtxSTR {
		const LPSTR STR_BIND_FORCE_FOLDER_SHORTCUT_RESOLVE = "Force Folder Shortcut Resolve";
		const LPSTR STR_AVOID_DRIVE_RESTRICTION_POLICY = "Avoid Drive Restriction Policy";
		const LPSTR STR_SKIP_BINDING_CLSID = "Skip Binding CLSID";
		const LPSTR STR_PARSE_PREFER_FOLDER_BROWSING     = "Parse Prefer Folder Browsing";
		const LPSTR STR_DONT_PARSE_RELATIVE              = "Don't Parse Relative";
		const LPSTR STR_PARSE_TRANSLATE_ALIASES          = "Parse Translate Aliases";
		const LPSTR STR_PARSE_SKIP_NET_CACHE             = "Skip Net Resource Cache";
		const LPSTR STR_PARSE_SHELL_PROTOCOL_TO_FILE_OBJECTS     = "Parse Shell Protocol To File Objects";
		const LPSTR STR_TRACK_CLSID                      = "Track the CLSID";
		const LPSTR STR_INTERNAL_NAVIGATE                = "Internal Navigation";
		const LPSTR STR_PARSE_PROPERTYSTORE             = "DelegateNamedProperties";
		const LPSTR STR_NO_VALIDATE_FILENAME_CHARS      = "NoValidateFilenameChars";
		const LPSTR STR_BIND_DELEGATE_CREATE_OBJECT = "Delegate Object Creation";
		const LPSTR STR_PARSE_ALLOW_INTERNET_SHELL_FOLDERS   = "Allow binding to Internet shell folder handlers and negate STR_PARSE_PREFER_WEB_BROWSING";
		const LPSTR STR_PARSE_PREFER_WEB_BROWSING   = "Do not bind to Internet shell folder handlers";
		const LPSTR STR_PARSE_SHOW_NET_DIAGNOSTICS_UI  = "Show network diagnostics UI";
		const LPSTR STR_PARSE_DONT_REQUIRE_VALIDATED_URLS  = "Do not require validated URLs";
		const LPSTR STR_INTERNETFOLDER_PARSE_ONLY_URLMON_BINDABLE  = "Validate URL";
		const LPSTR STR_BIND_FOLDERS_READ_ONLY = "Folders As Read Only";
		const LPSTR STR_BIND_FOLDER_ENUM_MODE = "Folder Enum Mode";
		const LPSTR STR_PARSE_WITH_EXPLICIT_PROGID = "ExplicitProgid";
		const LPSTR STR_PARSE_WITH_EXPLICIT_ASSOCAPP = "ExplicitAssociationApp";
		const LPSTR STR_PARSE_EXPLICIT_ASSOCIATION_SUCCESSFUL = "ExplicitAssociationSuccessful";
		const LPSTR STR_PARSE_AND_CREATE_ITEM    = "ParseAndCreateItem";
		const LPSTR STR_PROPERTYBAG_PARAM   = "SHBindCtxPropertyBag";
		const LPSTR STR_ENUM_ITEMS_FLAGS    = "SHCONTF";
		const LPSTR STR_ITEM_CACHE_CONTEXT       = "ItemCacheContext";
		const LPSTR STR_DONT_RESOLVE_LINK      = "Don't Resolve Link";
		const LPSTR STR_GET_ASYNC_HANDLER = "GetAsyncHandler";
		const LPSTR STR_GPS_HANDLERPROPERTIESONLY      = "GPS_HANDLERPROPERTIESONLY";
		const LPSTR STR_GPS_FASTPROPERTIESONLY         = "GPS_FASTPROPERTIESONLY";
		const LPSTR STR_GPS_OPENSLOWITEM               = "GPS_OPENSLOWITEM";
		const LPSTR STR_GPS_DELAYCREATION              = "GPS_DELAYCREATION";
		const LPSTR STR_GPS_BESTEFFORT                 = "GPS_BESTEFFORT";
		const LPSTR STR_GPS_NO_OPLOCK                  = "GPS_NO_OPLOCK";
		const LPSTR STR_FILE_SYS_BIND_DATA      = "File System Bind Data";
		const LPSTR STR_FILE_SYS_BIND_DATA_WIN7_FORMAT  = "Win7FileSystemIdList";
		const LPSTR STR_TAB_REUSE_IDENTIFIER   = "Tab Reuse Identifier";
		const LPSTR STR_REFERRER_IDENTIFIER    = "Referrer Identifier";
		const LPSTR STR_PARSE_WITH_PROPERTIES  = "ParseWithProperties";
		const LPSTR STR_PARSE_PARTIAL_IDLIST   = "ParseOriginalItem";
	}