[
    odl,
    uuid(00000001-0000-0000-C000-000000000046),
]
interface IClassFactory : stdole.IUnknown {

    HRESULT CreateInstance(
        [in] stdole.IUnknown *pUnkOuter,
        [in, out] UUID *riid,
        [in, out] long *ppvObject);

    HRESULT LockServer(
        [in] BOOL fLock);

}

[
    odl,
    uuid(B196B28F-BAB4-101A-B69C-00AA00341D07),
]
interface IClassFactory2 : stdole.IUnknown {

    typedef struct LICINFO {
        LONG cbLicInfo;
        BOOL fRuntimeKeyAvail;
        BOOL fLicVerified;
    } LICINFO;

    HRESULT CreateInstance(
        [in] stdole.IUnknown *pUnkOuter,
        [in, out] UUID *riid,
        [in, out] long *ppvObject);

    HRESULT LockServer(
        [in] BOOL fLock);

    HRESULT GetLicInfo(
        [in, out] LICINFO *pLicInfo);

    HRESULT RequestLicKey(
        [in] LONG dwReserved,
        [in, out] BSTR *pBstrKey);

    HRESULT CreateInstanceLic(
        [in] stdole.IUnknown *pUnkOuter,
        [in] stdole.IUnknown *pUnkReserved,
        [in, out] UUID *riid,
        [in] BSTR bstrKey,
        [in, out] long *ppvObj);
}