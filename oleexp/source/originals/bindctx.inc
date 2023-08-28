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
        [in] LPWSTR pszKey,
        [in] stdole.IUnknown*punk);

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
