[
    odl,
    uuid(00000000-0000-0000-C000-000000000046),
]
interface IUnknown{

    LONG QueryInterface(
        [in, out] UUID *riid,
        [in, out] void *ppvObject);

    LONG AddRef();
    LONG Release();
}

[
    odl,
    uuid(00000001-0000-0000-C000-000000000046),
]
interface IClassFactory : stdole.IUnknown {

    HRESULT CreateInstance(
        [in] stdole.IUnknown*pUnkOuter,
        [in] UUID *riid,
        [out] void *ppvObject);

    HRESULT LockServer(
        [in] BOOL fLock);

}

[
    odl,
    uuid(B196B28F-BAB4-101A-B69C-00AA00341D07),
]
interface IClassFactory2 : IClassFactory {

    typedef struct LICINFO {
        LONG cbLicInfo;
        BOOL fRuntimeKeyAvail;
        BOOL fLicVerified;
    } LICINFO;

    HRESULT GetLicInfo(
        [out] LICINFO *pLicInfo);

    HRESULT RequestLicKey(
        [in] LONG dwReserved,
        [out] BSTR *pBstrKey);

    HRESULT CreateInstanceLic(
        [in] stdole.IUnknown *pUnkOuter,
        [in] stdole.IUnknown *pUnkReserved,
        [in] UUID *riid,
        [in] BSTR bstrKey,
        [out] void *ppvObj);
}

