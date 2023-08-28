[
    odl,
    uuid(3127CA40-446E-11CE-8135-00AA004BB851),
    helpstring("IErrorLog interface")
]
interface IErrorLog : stdole.IUnknown {

    HRESULT AddError(
        [in] LONG pszPropName,
        [in, out] olelib.EXCEPINFO *pExcepInfo);
}
