//tbShellLib - 100% implemented

[
    odl,
    uuid(FC4801A3-2BA9-11CF-A229-00AA003D7352)
]
interface IObjectWithSite : stdole.IUnknown {

    [helpstring("Sets the object site")] 
    HRESULT SetSite(
        [in] stdole.IUnknown *pUnkSite);

    [helpstring("Gets the object site")] 
    HRESULT GetSite(
        [in, out] UUID *riid,
        [in, out] stdole.IUnknown **ppvSite);
}
