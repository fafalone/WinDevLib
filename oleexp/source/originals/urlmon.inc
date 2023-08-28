[
    dllname ("URLMON.DLL")
]
module urlmon {

    [entry("CreateURLMoniker")]
    HRESULT CreateURLMoniker(
        [in] IMoniker *pMkCtx,
        [in] LPWSTR szURL,
        [out, retval] IMoniker **ppmk);

    [entry("CreateURLMonikerEx")]
    HRESULT CreateURLMonikerEx(
         [in] IMoniker *pMkCtx,
         [in] LPWSTR szURL,
         [out] IMoniker **ppmk,
         [in] LONG dwFlags);

    [entry("GetClassURL")]
    HRESULT GetClassURL(
        [in] LPWSTR szURL,
        [out] UUID *pClsID);

    [entry("CreateAsyncBindCtx")]
    HRESULT CreateAsyncBindCtx(
        [in] LONG reserved,
        [in] IBindStatusCallback *pBSCb,
        [in] IEnumFORMATETC *pEFetc,
        [out, retval] IBindCtx **ppBC);

    [entry("CreateAsyncBindCtxEx")]
    HRESULT CreateAsyncBindCtxEx(
        [in] IBindCtx *pbc,
        [in] LONG dwOptions,
        [in] IBindStatusCallback *pBSCb,
        [in] IEnumFORMATETC *pEnum,
        [out] IBindCtx **ppBC,
        [in] LONG reserved);

    [entry("MkParseDisplayNameEx")]
    HRESULT MkParseDisplayNameEx(
        [in] IBindCtx *pbc,
        [in] LPWSTR szDisplayName,
        [out, defaultvalue(0)] LONG *pchEaten,
        [out, retval] IMoniker **ppmk);

    [entry("RegisterBindStatusCallback")]
    HRESULT RegisterBindStatusCallback(
        [in] IBindCtx *pBC,
        [in] IBindStatusCallback *pBSCb,
        [out] IBindStatusCallback **ppBSCBPrev,
        [in] LONG dwReserved);

    [entry("RevokeBindStatusCallback")]
    HRESULT RevokeBindStatusCallback(
        [in] IBindCtx *pBC,
        [in] IBindStatusCallback *pBSCb);

    [entry("GetClassFileOrMime")]
    HRESULT GetClassFileOrMime(
        [in] IBindCtx *pBC,
        [in] LPWSTR szFilename,
        [in, out] void *pBuffer,
        [in] LONG cbSize,
        [in] LPWSTR szMime,
        [in] LONG dwReserved,
        [out] UUID *pclsid);

    [entry("IsValidURL")]
    LONG IsValidURL(
        [in] IBindCtx *pBC,
        [in] LPWSTR szURL,
        [in, defaultvalue(0)] LONG dwReserved);

    [entry("CoGetClassObjectFromURL")]
    HRESULT CoGetClassObjectFromURL(
        [in] UUID *rCLASSID,
        [in] LPWSTR szCODE,
        [in] LONG dwFileVersionMS,
        [in] LONG dwFileVersionLS,
        [in] LPWSTR szTYPE,
        [in] IBindCtx *pBindCtx,
        [in] LONG dwClsContext,
        [in] void *pvReserved,
        [in] UUID *riid,
        [out] void *ppv);

    [entry("FaultInIEFeature")]
    HRESULT FaultInIEFeature(
        [in] LONG hWnd,
        [in] uCLSSPEC *pClassSpec,
        [in] QUERYCONTEXT *pQuery,
        [in] FIEF dwFlags);

    [entry("GetComponentIDFromCLSSPEC")]
    HRESULT GetComponentIDFromCLSSPEC(
        [in] uCLSSPEC *pClassspec,
        [out, retval] LONG *ppszComponentID);

    [entry("IsAsyncMoniker")]
    HRESULT IsAsyncMoniker(
        [in] IMoniker*pmk);

    [entry("CreateURLBinding")]
    HRESULT CreateURLBinding(
        [in] LPWSTR lpszUrl,
        [in] IBindCtx *pbc,
        [out, retval] IBinding **ppBdg);

    [entry("RegisterMediaTypes")]
    HRESULT RegisterMediaTypes(
        [in] LONG ctypes,
        [in] LPSTR rgszTypes,
        [in] short *rgcfTypes);

    [entry("FindMediaType")]
    HRESULT FindMediaType(
        [in] LPSTR rgszTypes,
        [out, retval] short *rgcfTypes);

    [entry("CreateFormatEnumerator")]
    HRESULT CreateFormatEnumerator(
        [in] LONG cfmtetc,
        [in] FORMATETC *rgfmtetc,
        [out, retval] IEnumFORMATETC **ppenumfmtetc);

    [entry("RegisterFormatEnumerator")]
    HRESULT RegisterFormatEnumerator(
        [in] IBindCtx *pBC,
        [in] IEnumFORMATETC *pEFetc,
        [in] LONG reserved);

    [entry("RevokeFormatEnumerator")]
    HRESULT RevokeFormatEnumerator(
        [in] IBindCtx *pBC,
        [in] IEnumFORMATETC *pEFetc);

    [entry("RegisterMediaTypeClass")]
    HRESULT RegisterMediaTypeClass(
        [in] IBindCtx *pBC,
        [in] LONG ctypes,
        [in] LPSTR rgszTypes,
        [in] UUID *rgclsID,
        [in] LONG reserved);

    [entry("FindMediaTypeClass")]
    HRESULT FindMediaTypeClass(
        [in] IBindCtx *pBC,
        [in] LPSTR szType,
        [in] UUID *pclsID,
        [in] LONG reserved);

    [entry("UrlMkSetSessionOption")]
    HRESULT UrlMkSetSessionOption(
        [in] LONG dwOption,
        [in] void *pBuffer,
        [in] LONG dwBufferLength,
        [in] LONG dwReserved);

    [entry("UrlMkGetSessionOption")]
    HRESULT UrlMkGetSessionOption(
        [in] LONG dwOption,
        [in, out] void *pBuffer,
        [in] LONG dwBufferLength,
        [out] LONG *pdwBufferLength,
        [in] LONG dwReserved);

    [entry("FindMimeFromData")]
    HRESULT FindMimeFromData(
        [in] IBindCtx *pBC,           // bind context - can be NULL
        [in] LPWSTR pwzUrl,           // url - can be null
        [in] void *pBuffer,           // buffer with data to sniff - can be null (pwzUrl must be valid)
        [in] LONG cbSize,             // size of buffer
        [in] LPWSTR pwzMimeProposed,  // proposed mime if - can be null
        [in] LONG dwMimeFlags,        // will be defined
        [out] LONG *ppwzMimeOut,      // the suggested mime
        [in] LONG dwReserved);        // must be 0

    const LONG FMFD_DEFAULT = 0x00000000;
    const LONG FMFD_URLASFILENAME = 0x00000001;

    [entry("ObtainUserAgentString")]
    HRESULT ObtainUserAgentString(
        [in] LONG dwOption,
        [in] LPSTR pszUAOut,
        [in] LONG *cbSize);

    // URLMON-specific defines for UrlMkSetSessionOption() above
    const LONG URLMON_OPTION_USERAGENT = 0x10000001;
    const LONG URLMON_OPTION_USERAGENT_REFRESH = 0x10000002;
    const LONG URLMON_OPTION_URL_ENCODING = 0x10000004;

    [entry("HlinkSimpleNavigateToString")]
    HRESULT HlinkSimpleNavigateToString(
        [in] LPWSTR szTarget,          // required - target document - null if local jump w/in doc
        [in] LPWSTR szLocation,        // optional, for navigation into middle of a doc
        [in] LPWSTR szTargetFrameName, // optional, for targeting frame-sets
        [in] stdole.IUnknown *pUnk,    // required - we'll search this for other necessary interfaces
        [in] IBindCtx *pbc,            // optional. caller may register an IBSC in this
        [in] IBindStatusCallback *pbsc,
        [in] LONG grfHLNF,             // flags
        [in, defaultvalue(0)] LONG dwReserved); // for future use, must be NULL

    [entry("HlinkSimpleNavigateToMoniker")]
    HRESULT HlinkSimpleNavigateToMoniker(
        [in] IMoniker *pmkTarget,      // required - target document - (may be null
        [in] LPWSTR szLocation,        // optional, for navigation into middle of a doc
        [in] LPWSTR szTargetFrameName, // optional, for targeting frame-sets
        [in] stdole.IUnknown *pUnk,    // required - we'll search this for other necessary interfaces
        [in] IBindCtx *pbc,            // optional. caller may register an IBSC in this
        [in] IBindStatusCallback *pbsc,
        [in] LONG grfHLNF,             // flags
        [in, defaultvalue(0)] LONG dwReserved);          // for future use, must be NULL

    [entry("URLOpenStreamA")]
    LONG URLOpenStream(
        [in] stdole.IUnknown*pCaller,
        [in] LPSTR szURL,
        [in] LONG dw,
        [in] IBindStatusCallback *pCallback);

    [entry("URLOpenStreamW")]
    LONG URLOpenStreamW(
        [in] stdole.IUnknown*pCaller,
        [in] LPWSTR szURL,
        [in] LONG dw,
        [in] IBindStatusCallback *pCallback);

    [entry("URLOpenPullStreamA")]
    LONG URLOpenPullStream(
        [in] stdole.IUnknown*pCaller,
        [in] LPSTR szURL,
        [in] LONG dwReserved,
        [in] IBindStatusCallback *pCallback);

    [entry("URLOpenPullStreamW")]
    LONG URLOpenPullStreamW(
        [in] stdole.IUnknown*pCaller,
        [in] LPWSTR szURL,
        [in] LONG dwReserved,
        [in] IBindStatusCallback *pCallback);

    [entry("URLDownloadToCacheFileA")]
    LONG URLDownloadToCacheFile(
        [in] stdole.IUnknown*lpUnkcaller,
        [in] LPSTR szURL,
        [in, out] LPSTR szFileName,
        [in] LONG dwBufLength,
        [in] LONG dwReserved,
        [in] IBindStatusCallback *pBSC);

    [entry("URLDownloadToCacheFileW")]
    LONG URLDownloadToCacheFileW(
        [in] stdole.IUnknown*lpUnkcaller,
        [in] LPWSTR szURL,
        [in] LONG szFileName,
        [in] LONG dwBufLength,
        [in] LONG dwReserved,
        [in] IBindStatusCallback *pBSC);

    [entry("URLOpenBlockingStreamA")]
    LONG URLOpenBlockingStream(
        [in] stdole.IUnknown*pCaller,
        [in] LPSTR szURL,
        [in, out] IStream **Strm,
        [in] LONG dwReserved,
        [in] IBindStatusCallback *lpfnCB);

    [entry("URLOpenBlockingStreamW")]
    LONG URLOpenBlockingStreamW(
        [in] stdole.IUnknown*pCaller,
        [in] LPWSTR szURL,
        [in, out] IStream **Strm,
        [in] LONG dwReserved,
        [in] IBindStatusCallback *lpfnCB);

    [entry("URLDownloadToFileA")]
    LONG URLDownloadToFile(
        [in] stdole.IUnknown*pCaller,
        [in] LPSTR szURL,
        [in] LPSTR szFile,
        [in] LONG dwReserved,
        [in] IBindStatusCallback *lpfnCB);

    [entry("URLDownloadToFileW")]
    LONG URLDownloadToFileW(
        [in] stdole.IUnknown*pCaller,
        [in] LPWSTR szURL,
        [in] LPWSTR szFile,
        [in] LONG dwReserved,
        [in] IBindStatusCallback *lpfnCB);

    [entry("HlinkGoBack")]
    HRESULT HlinkGoBack(
        [in] stdole.IUnknown *pUnk);

    [entry("HlinkGoForward")]
    HRESULT HlinkGoForward(
        [in] stdole.IUnknown *pUnk);

    [entry("HlinkNavigateString")]
    HRESULT HlinkNavigateString(
        [in] stdole.IUnknown *pUnk,
        [in] LPWSTR szTarget);

    [entry("HlinkNavigateMoniker")]
    HRESULT HlinkNavigateMoniker(
        [in] stdole.IUnknown *pUnk,
        [in] IMoniker *pmkTarget);

    [entry("CoInternetParseUrl")]
    HRESULT CoInternetParseUrl(
        [in] LPWSTR pwzUrl,           
        [in] PARSEACTION ParseAction,
        [in] LONG dwFlags,
        [in] LPWSTR pszResult,
        [in] LONG cchResult,
        [in, out] LONG *pcchResult,
        [in, defaultvalue(0)] LONG dwReserved);

    [entry("CoInternetCombineUrl")]
    HRESULT CoInternetCombineUrl(
        [in] LPWSTR pwzBaseUrl,
        [in] LPWSTR pwzRelativeUrl,
        [in] LONG dwCombineFlags,
        [in] LPWSTR pszResult,
        [in] LONG cchResult,
        [in, out] LONG *pcchResult,
        [in, defaultvalue(0)] LONG dwReserved);

    [entry("CoInternetCompareUrl")]
    LONG CoInternetCompareUrl(
        [in] LPWSTR pwzUrl1,
        [in] LPWSTR pwzUrl2,
        [in] LONG dwFlags);

    [entry("CoInternetGetProtocolFlags")]
    HRESULT CoInternetGetProtocolFlags(
        [in] LPWSTR pwzUrl,
        [in, out] LONG *pdwFlags,
        [in, defaultvalue(0)] LONG dwReserved);

    [entry("CoInternetQueryInfo")]
    HRESULT CoInternetQueryInfo(
        [in] LPWSTR pwzUrl,
        [in] QUERYOPTION QueryOptions,
        [in] LONG dwQueryFlags,
        [in] void *pvBuffer,
        [in] LONG cbBuffer,
        [in, out] LONG *pcbBuffer,
        [in, defaultvalue(0)] LONG dwReserved);

    [entry("CoInternetGetSession")]
    HRESULT CoInternetGetSession(
        [in] LONG dwSessionMode,
        [in, out] IInternetSession **ppIInternetSession,
        [in, defaultvalue(0)] LONG dwReserved);

    [entry("CoInternetGetSecurityUrl")]
    HRESULT CoInternetGetSecurityUrl(
        [in] LPWSTR pwzUrl,
        [in] LONG *ppwzSecUrl,
        [in] PSUACTION psuAction,
        [in, defaultvalue(0)] LONG dwReserved);

    [entry("CopyStgMedium")]
    HRESULT CopyStgMedium(
        [in] STGMEDIUM *pcstgmedSrc,
        [out] STGMEDIUM *pstgmedDest);

    [entry("CopyBindInfo")]
    HRESULT CopyBindInfo(
        [in] BINDINFO *pcbiSrc,
        [out] BINDINFO *pbiDest);

    [entry("ReleaseBindInfo")]
    void ReleaseBindInfo(
        [in] BINDINFO *pbindinfo);

    [entry("CoInternetCreateSecurityManager")]
    LONG CoInternetCreateSecurityManager(
        [in] IServiceProvider *pSP,
        [in, out] IInternetSecurityManager **ppSM,
        [in] LONG dwReserved);

    [entry("CoInternetCreateZoneManager")]
    LONG CoInternetCreateZoneManager(
        [in] IServiceProvider *pSP,
        [in, out] IInternetZoneManager **ppZM,
        [in] LONG dwReserved);
/*
    [entry("GetSoftwareUpdateInfo")]
    HRESULT GetSoftwareUpdateInfo(
        [in] LPWSTR szDistUnit,
        [in, out] SOFTDISTINFO *psdi);

    [entry("SetSoftwareUpdateAdvertisementState")]
    HRESULT SetSoftwareUpdateAdvertisementState(
        [in] LPWSTR szDistUnit,
        [in] LONG dwAdState,
        [in] LONG dwAdvertisedVersionMS,
        [in] LONG dwAdvertisedVersionLS);
*/
}

typedef struct PROTOCOLFILTERDATA {
    LONG cbSize;
    IInternetProtocolSink   *pProtocolSink;  // out parameter
    IInternetProtocol       *pProtocol;      // in parameter
    IUnknown                *pUnk;
    LONG dwFilterFlags;
} PROTOCOLFILTERDATA;

[
    odl,
    uuid(79eac9c0-baf9-11ce-8c82-00aa004ba90b),
]
interface IBinding: stdole.IUnknown {

    HRESULT Abort();

    HRESULT Suspend();

    HRESULT Resume();

    HRESULT SetPriority(
        [in] LONG nPriority);

    HRESULT GetPriority(
        [out] LONG *pnPriority);

    HRESULT GetBindResult(
        [out] UUID *pclsidProtocol,
        [out] LONG *pdwResult,
        [out] LONG *pszResult,
        [in, out] LONG *pdwReserved);

}

[
    odl,
    uuid(79eac9c1-baf9-11ce-8c82-00aa004ba90b)
]
interface IBindStatusCallback: stdole.IUnknown {

    HRESULT OnStartBinding(
        [in] LONG dwReserved,
        [in] IBinding *pib);

    HRESULT GetPriority(
        [out, retval] LONG *pnPriority);

    HRESULT OnLowResource(
        [in] LONG reserved);

    HRESULT OnProgress(
        [in] LONG ulProgress,
        [in] LONG ulProgressMax,
        [in] BINDSTATUS ulStatusCode,
        [in] LONG szStatusText);

    HRESULT OnStopBinding(
        [in] LONG hresult,
        [in] LONG szError);

    HRESULT GetBindInfo(
        [in, out] BINDF *grfBINDF,
        [in, out] BINDINFO *pbindinfo);

    HRESULT OnDataAvailable(
        [in] BSCF grfBSCF,
        [in] LONG dwSize,
        [in, out] FORMATETC *pformatetc,
        [in, out] STGMEDIUM *pstgmed);

    HRESULT OnObjectAvailable(
        [in, out] UUID *riid,
        [in] stdole.IUnknown *punk);

}

[
    odl,
    uuid(79eac9d0-baf9-11ce-8c82-00aa004ba90b),
]
interface IAuthenticate : stdole.IUnknown{

    HRESULT Authenticate(
        [in, out] LONG *phwnd,
        [in, out] LONG *pszUsername,
        [in, out] LONG *pszPassword);
}

[
    odl,
    uuid(79eac9ec-baf9-11ce-8c82-00aa004ba90b),
]
interface IInternetProtocolInfo : stdole.IUnknown{

    //
    // returns:
    // S_OK on success
    // S_FALSE on success but buffer is too small
    // INET_E_DEFAULT_ACTION if action is not know should be handled be default implementation
    //
    HRESULT ParseUrl(
       [in] LONG pwzUrl,
       [in] PARSEACTION ParseAction,
       [in] LONG dwParseFlags,
       [in] LONG pwzResult,
       [in] LONG cchResult,
       [in, out] LONG *pcchResult,
       [in] LONG dwReserved);

    //
    // returns:
    // S_OK on success
    // S_FALSE on success but buffer is too small
    // INET_E_DEFAULT_ACTION if action is not know should be handled be default implementation
    //
    HRESULT CombineUrl(
        [in] LONG pwzBaseUrl,
        [in] LONG pwzRelativeUrl,
        [in] LONG dwCombineFlags,
        [in] LONG pwzResult,
        [in] LONG cchResult,
        [in, out] LONG *pcchResult,
        [in] LONG dwReserved);

    //
    // returns:
    // S_OK  if equal
    // S_FALSE otherwise
    //
    HRESULT CompareUrl(
       [in] LONG pwzUrl1,
       [in] LONG pwzUrl2,
       [in] LONG dwCompareFlags);

    //
    // returns:
    // S_OK on success
    // S_FALSE on success but buffer is too small
    // INET_E_QUERYOPTION_UNKNOWN
    //
    HRESULT QueryInfo(
        [in] LONG pwzUrl,
        [in] QUERYOPTION OueryOption,
        [in] LONG dwQueryFlags,
        [in] LONG pBuffer,
        [in] LONG cbBuffer,
        [in, out] LONG *pcbBuf,
        [in] LONG dwReserved);

}

[
    odl,
    uuid(79eac9eb-baf9-11ce-8c82-00aa004ba90b),
]
interface IInternetPriority : stdole.IUnknown{

    HRESULT SetPriority(
        [in] LONG nPriority);

    HRESULT GetPriority(
        [in, out] LONG *pnPriority);

}


[
    odl,
    uuid(79eac9e7-baf9-11ce-8c82-00aa004ba90b),
]
interface IInternetSession : stdole.IUnknown {

   typedef enum OIBDG_FLAGS {
       // default is free threaded
       OIBDG_APARTMENTTHREADED        = 0x00000100,
       OIBDG_DATAONLY                 = 0x00001000
   } OIBDG_FLAGS;

   HRESULT RegisterNameSpace(
       [in] IClassFactory *pCF,
       [in,out] UUID *rclsid,
       [in] LPWSTR pwzProtocol,
       [in] LONG cPatterns,
       [in, out] LONG *ppwzPatterns,
       [in] LONG dwReserved);

   HRESULT UnregisterNameSpace(
       [in] IClassFactory *pCF,
       [in] LONG pszProtocol);

   HRESULT RegisterMimeFilter(
       [in] IClassFactory *pCF,
       [in, out] UUID *rclsid,
       [in] LONG pwzType);

   HRESULT UnregisterMimeFilter(
       [in] IClassFactory *pCF,
       [in] LONG pwzType);

   HRESULT CreateBinding(
       [in] IBindCtx *pBC,
       [in] LPWSTR szUrl,
       [in] stdole.IUnknown *pUnkOuter,
       [in, out] stdole.IUnknown **ppUnk,
       [in, out] IInternetProtocol **ppOInetProt,
       [in] LONG dwOption);

   HRESULT SetSessionOption(
       [in] LONG dwOption,
       [in] void *pBuffer,
       [in] LONG dwBufferLength,
       [in, defaultvalue(0)] LONG dwReserved);

   HRESULT GetSessionOption(
       [in] LONG dwOption,
       [in] void *pBuffer,
       [in, out] LONG *pdwBufferLength,
       [in, defaultvalue(0)] LONG dwReserved);
}

[
    odl,
    uuid(79eac9e3-baf9-11ce-8c82-00aa004ba90b)
]

interface IInternetProtocolRoot : stdole.IUnknown {

    typedef enum PI_FLAGS {
        PI_PARSE_URL               = 0x00000001,
        PI_FILTER_MODE             = 0x00000002,
        PI_FORCE_ASYNC             = 0x00000004,
        PI_USE_WORKERTHREAD        = 0x00000008,
        PI_MIMEVERIFICATION        = 0x00000010,
        PI_CLSIDLOOKUP             = 0x00000020,
        PI_DATAPROGRESS            = 0x00000040,
        PI_SYNCHRONOUS             = 0x00000080,
        PI_APARTMENTTHREADED       = 0x00000100,
        PI_CLASSINSTALL            = 0x00000200,
        PI_PASSONBINDCTX           = 0x00002000,
        PI_NOMIMEHANDLER           = 0x00008000,
        PI_LOADAPPDIRECT           = 0x00004000,
        PD_FORCE_SWITCH            = 0x00010000,
        PI_PREFERDEFAULTHANDLER    = 0x00020000
    } PI_FLAGS;

    typedef struct PROTOCOLDATA {
        LONG grfFlags;
        LONG dwState;
        LONG pData;
        LONG cbData;
    } PROTOCOLDATA;

    typedef struct StartParam {
        UUID iid;
        IBindCtx *pIBindCtx;
        stdole.IUnknown *pItf;
    } StartParam;

    HRESULT Start(
       [in] LONG szUrl,
       [in] IInternetProtocolSink *pOIProtSink,
       [in] IInternetBindInfo *pOIBindInfo,
       [in] PI_FLAGS grfPI,
       [in] PROTOCOLFILTERDATA *dwReserved);

    HRESULT Continue(
       [in] PROTOCOLDATA *pProtocolData);

    HRESULT Abort(
       [in] LONG hrReason,
       [in] LONG dwOptions);

    HRESULT Terminate(
       [in] LONG dwOptions);

    HRESULT Suspend();

    HRESULT Resume();
}

[
    odl,
    uuid(79eac9e4-baf9-11ce-8c82-00aa004ba90b)
]
//interface IInternetProtocol : IInternetProtocolRoot {
interface IInternetProtocol : stdole.IUnknown {

    HRESULT Start(
       [in] LONG szUrl,
       [in] IInternetProtocolSink *pOIProtSink,
       [in] IInternetBindInfo *pOIBindInfo,
       [in] PI_FLAGS grfPI,
       [in] PROTOCOLFILTERDATA *dwReserved);

    HRESULT Continue(
       [in] PROTOCOLDATA *pProtocolData);

    HRESULT Abort(
       [in] LONG hrReason,
       [in] LONG dwOptions);

    HRESULT Terminate(
       [in] LONG dwOptions);

    HRESULT Suspend();

    HRESULT Resume();

    LONG Read(
        [in] LONG pv,
        [in] LONG cb,
        [in, out] LONG *pcbRead);

    HRESULT Seek(
        [in] CURRENCY dlibMove,
        [in] LONG dwOrigin,
        [in, out] CURRENCY *plibNewPosition);

    HRESULT LockRequest(
        [in] LONG dwOptions);

    HRESULT UnlockRequest();

}

[
    odl,
    uuid(79eac9e5-baf9-11ce-8c82-00aa004ba90b),
]
interface IInternetProtocolSink : stdole.IUnknown {

    HRESULT Switch(
       [in] PROTOCOLDATA *pProtocolData);

    HRESULT ReportProgress(
       [in] BINDSTATUS ulStatusCode,
       [in] LPWSTR szStatusText);

    HRESULT ReportData(
       [in] BSCF grfBSCF,
       [in] LONG ulProgress,
       [in] LONG ulProgressMax);

    HRESULT ReportResult(
       [in] LONG hrResult,
       [in] LONG dwError,
       [in] LPWSTR szResult);
}

[
    odl,
    uuid(79eac9e1-baf9-11ce-8c82-00aa004ba90b),
]
interface IInternetBindInfo : stdole.IUnknown {

    typedef enum BINDSTRING {
        BINDSTRING_HEADERS  = 1
       ,BINDSTRING_ACCEPT_MIMES
       ,BINDSTRING_EXTRA_URL
       ,BINDSTRING_LANGUAGE
       ,BINDSTRING_USERNAME
       ,BINDSTRING_PASSWORD
       ,BINDSTRING_UA_PIXELS
       ,BINDSTRING_UA_COLOR
       ,BINDSTRING_OS
       ,BINDSTRING_USER_AGENT
       ,BINDSTRING_ACCEPT_ENCODINGS
       ,BINDSTRING_POST_COOKIE
       ,BINDSTRING_POST_DATA_MIME
       ,BINDSTRING_URL
       ,BINDSTRING_IID
       ,BINDSTRING_FLAG_BIND_TO_OBJECT
       ,BINDSTRING_PTR_BIND_CONTEXT
    } BINDSTRING;

    HRESULT GetBindInfo(
        [in, out] LONG *grfBINDF,
        [in, out] BINDINFO *pbindinfo);

    HRESULT GetBindString(
        [in] LONG ulStringType,
        [in, out] LONG *ppwzStr,
        [in] LONG cEl,
        [in, out] LONG *pcElFetched);
}


[
    odl,
    uuid(79eac9cd-baf9-11ce-8c82-00aa004ba90b),
]
interface IBindProtocol : stdole.IUnknown {

    HRESULT CreateBinding(
        [in] LONG szUrl,
        [in] IBindCtx *pbc,
        [out, retval] IBinding **ppb);

}

[
    odl,
    uuid(79eac9d2-baf9-11ce-8c82-00aa004ba90b),
]
interface IHttpNegotiate : stdole.IUnknown {

    HRESULT BeginningTransaction(
        [in] LONG szURL,
        [in] LONG szHeaders,
        [in] LONG dwReserved,
        [in, out] LONG *pszAdditionalHeaders);

    HRESULT OnResponse(
        [in] LONG dwResponseCode,
        [in] LONG szResponseHeaders,
        [in] LONG szRequestHeaders,
        [in, out] LONG *pszAdditionalRequestHeaders);
}

[
    odl,
    uuid(79eac9d5-bafa-11ce-8c82-00aa004ba90b),
]
interface IWindowForBindingUI : stdole.IUnknown {

    HRESULT GetWindow(
        [in, out] UUID *rguidReason,
        [out, retval] LONG *phwnd);
}

[
    odl,
    uuid(79eac9d6-bafa-11ce-8c82-00aa004ba90b),
]
interface IWinInetInfo : stdole.IUnknown {

    LONG QueryOption(
        [in] QUERYOPTION dwOption,
        [in] void *pBuffer,
        [in, out] LONG *pcbBuf);

}

typedef enum HTTP_QUERY {
    HTTP_QUERY_MIME_VERSION                 = 0,
    HTTP_QUERY_CONTENT_TYPE                 = 1,
    HTTP_QUERY_CONTENT_TRANSFER_ENCODING    = 2,
    HTTP_QUERY_CONTENT_ID                   = 3,
    HTTP_QUERY_CONTENT_DESCRIPTION          = 4,
    HTTP_QUERY_CONTENT_LENGTH               = 5,
    HTTP_QUERY_CONTENT_LANGUAGE             = 6,
    HTTP_QUERY_ALLOW                        = 7,
    HTTP_QUERY_PUBLIC                       = 8,
    HTTP_QUERY_DATE                         = 9,
    HTTP_QUERY_EXPIRES                      = 10,
    HTTP_QUERY_LAST_MODIFIED                = 11,
    HTTP_QUERY_MESSAGE_ID                   = 12,
    HTTP_QUERY_URI                          = 13,
    HTTP_QUERY_DERIVED_FROM                 = 14,
    HTTP_QUERY_COST                         = 15,
    HTTP_QUERY_LINK                         = 16,
    HTTP_QUERY_PRAGMA                       = 17,
    HTTP_QUERY_VERSION                      = 18,
    HTTP_QUERY_STATUS_CODE                  = 19,
    HTTP_QUERY_STATUS_TEXT                  = 20,
    HTTP_QUERY_RAW_HEADERS                  = 21,
    HTTP_QUERY_RAW_HEADERS_CRLF             = 22,
    HTTP_QUERY_CONNECTION                   = 23,
    HTTP_QUERY_ACCEPT                       = 24,
    HTTP_QUERY_ACCEPT_CHARSET               = 25,
    HTTP_QUERY_ACCEPT_ENCODING              = 26,
    HTTP_QUERY_ACCEPT_LANGUAGE              = 27,
    HTTP_QUERY_AUTHORIZATION                = 28,
    HTTP_QUERY_CONTENT_ENCODING             = 29,
    HTTP_QUERY_FORWARDED                    = 30,
    HTTP_QUERY_FROM                         = 31,
    HTTP_QUERY_IF_MODIFIED_SINCE            = 32,
    HTTP_QUERY_LOCATION                     = 33,
    HTTP_QUERY_ORIG_URI                     = 34,
    HTTP_QUERY_REFERER                      = 35,
    HTTP_QUERY_RETRY_AFTER                  = 36,
    HTTP_QUERY_SERVER                       = 37,
    HTTP_QUERY_TITLE                        = 38,
    HTTP_QUERY_USER_AGENT                   = 39,
    HTTP_QUERY_WWW_AUTHENTICATE             = 40,
    HTTP_QUERY_PROXY_AUTHENTICATE           = 41,
    HTTP_QUERY_ACCEPT_RANGES                = 42,
    HTTP_QUERY_SET_COOKIE                   = 43,
    HTTP_QUERY_COOKIE                       = 44,
    HTTP_QUERY_REQUEST_METHOD               = 45,
    HTTP_QUERY_REFRESH                      = 46,
    HTTP_QUERY_CONTENT_DISPOSITION          = 47,
    HTTP_QUERY_AGE                          = 48,
    HTTP_QUERY_CACHE_CONTROL                = 49,
    HTTP_QUERY_CONTENT_BASE                 = 50,
    HTTP_QUERY_CONTENT_LOCATION             = 51,
    HTTP_QUERY_CONTENT_MD5                  = 52,
    HTTP_QUERY_CONTENT_RANGE                = 53,
    HTTP_QUERY_ETAG                         = 54,
    HTTP_QUERY_HOST                         = 55,
    HTTP_QUERY_IF_MATCH                     = 56,
    HTTP_QUERY_IF_NONE_MATCH                = 57,
    HTTP_QUERY_IF_RANGE                     = 58,
    HTTP_QUERY_IF_UNMODIFIED_SINCE          = 59,
    HTTP_QUERY_MAX_FORWARDS                 = 60,
    HTTP_QUERY_PROXY_AUTHORIZATION          = 61,
    HTTP_QUERY_RANGE                        = 62,
    HTTP_QUERY_TRANSFER_ENCODING            = 63,
    HTTP_QUERY_UPGRADE                      = 64,
    HTTP_QUERY_VARY                         = 65,
    HTTP_QUERY_VIA                          = 66,
    HTTP_QUERY_WARNING                      = 67,
    HTTP_QUERY_EXPECT                       = 68,
    HTTP_QUERY_PROXY_CONNECTION             = 69,
    HTTP_QUERY_UNLESS_MODIFIED_SINCE        = 70,     HTTP_QUERY_ECHO_REQUEST                 = 71,
    HTTP_QUERY_ECHO_REPLY                   = 72,
    HTTP_QUERY_ECHO_HEADERS                 = 73,
    HTTP_QUERY_ECHO_HEADERS_CRLF            = 74,     HTTP_QUERY_PROXY_SUPPORT                = 75,
    HTTP_QUERY_AUTHENTICATION_INFO          = 76,
    HTTP_QUERY_PASSPORT_URLS                = 77,
    HTTP_QUERY_PASSPORT_CONFIG              = 78,
    HTTP_QUERY_CUSTOM                       = 65535,
    HTTP_QUERY_FLAG_REQUEST_HEADERS         = 0x80000000,
    HTTP_QUERY_FLAG_SYSTEMTIME              = 0x40000000,
    HTTP_QUERY_FLAG_NUMBER                  = 0x20000000,
    HTTP_QUERY_FLAG_COALESCE                = 0x10000000
} HTTP_QUERY;

[
    odl,
    uuid(79eac9d8-bafa-11ce-8c82-00aa004ba90b),
]
interface IWinInetHttpInfo : IWinInetInfo {

    LONG QueryInfo(
        [in] HTTP_QUERY dwOption,
        [in] void *pBuffer,
        [in, out] LONG *pcbBuf,
        [in, out] LONG *pdwFlags,
        [in, out] LONG *pdwReserved);

}

[
    odl,
    uuid(fc4801a1-2ba9-11cf-a229-00aa003d7352),
]
interface IBindHost : stdole.IUnknown{

    HRESULT CreateMoniker(
        [in] LONG szName,
        [in] IBindCtx *pBC,
        [in, out] IMoniker **ppmk,
        [in] LONG dwReserved);

    HRESULT MonikerBindToStorage(
        [in] IMoniker *pMk,
        [in] IBindCtx *pBC,
        [in] IBindStatusCallback *pBSC,
        [in, out] UUID *riid,
        [in, out] LONG *ppvObj);

    HRESULT MonikerBindToObject(
        [in] IMoniker *pMk,
        [in] IBindCtx *pBC,
        [in] IBindStatusCallback *pBSC,
        [in, out] UUID *riid,
        [in, out] LONG *ppvObj);

}

[
    odl,
    uuid(4F9F9FCB-E0F4-48eb-B7AB-FA2EA9365CB4)
] 
interface IHttpNegotiate2 : stdole.IUnknown {

    HRESULT BeginningTransaction(
        [in] LONG szURL,
        [in] LONG szHeaders,
        [in] LONG dwReserved,
        [in, out] LONG *pszAdditionalHeaders);

    HRESULT OnResponse(
        [in] LONG dwResponseCode,
        [in] LONG szResponseHeaders,
        [in] LONG szRequestHeaders,
        [in, out] LONG *pszAdditionalRequestHeaders);

    HRESULT GetRootSecurityId(
        [out] BYTE *pbSecurityId,
        [in, out] LONG *pcbSecurityId,
        [in] LONG dwReserved);

}

[
    odl,
    uuid(79eac9d7-bafa-11ce-8c82-00aa004ba90b),
]
interface IHttpSecurity : stdole.IUnknown {

    HRESULT GetWindow(
        [in] UUID *rguidReason,
        [out] LONG *phwnd);

    HRESULT OnSecurityProblem(
        [in] LONG dwProblem);

}