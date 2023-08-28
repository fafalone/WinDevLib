//tbShellLib - 100% implemented

[
    odl,
    uuid(1CF2B120-547D-101B-8E65-08002B2BD119)
]
interface IErrorInfo: stdole.IUnknown {

    HRESULT GetGUID(
        [out] UUID *pGUID);

    HRESULT GetSource(
        [out, retval] BSTR *pBstrSource);

    HRESULT GetDescription(
        [out, retval] BSTR *pBstrDescription);

    HRESULT GetHelpFile(
        [out, retval] BSTR *pBstrHelpFile);

    HRESULT GetHelpContext(
        [out, retval] LONG *pdwHelpContext);
}

[
    odl,
    uuid(22F03340-547D-101B-8E65-08002B2BD119)
]
interface ICreateErrorInfo: stdole.IUnknown {

    HRESULT SetGUID(
        [in] UUID *rguid);

    HRESULT SetSource(
        [in] LPWSTR szSource);

    HRESULT SetDescription(
        [in] LPWSTR szDescription);

    HRESULT SetHelpFile(
        [in] LPWSTR szHelpFile);

    HRESULT SetHelpContext(
        [in] LONG dwHelpContext);
}


[
    odl,
    uuid(DF0B3D60-548F-101B-8E65-08002B2BD119)
]
interface ISupportErrorInfo: stdole.IUnknown {

    HRESULT InterfaceSupportsErrorInfo(
        [in] UUID *riid);

}

