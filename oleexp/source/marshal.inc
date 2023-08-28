//tbShellLib - 100% implemented

[
    odl,
    uuid(00000003-0000-0000-C000-000000000046)
]
interface IMarshal : stdole.IUnknown {

    HRESULT GetUnmarshalClass (
        [in] UUID *riid,
        [in] void *pv,
        [in] MSHCTX dwDestContext,
        [in] void *pvDestContext,
        [in] MSHLFLAGS mshlflags,
        [out] UUID *pCid);

    HRESULT GetMarshalSizeMax(
        [in] UUID *riid,
        [in] void *pv,
        [in] MSHCTX dwDestContext,
        [in] void *pvDestContext,
        [in] MSHLFLAGS mshlflags,
        [out, retval] LONG *pSize);

    HRESULT MarshalInterface(
        [in] IStream *pStm,
        [in] UUID *riid,
        [in] void *pv,
        [in] MSHCTX dwDestContext,
        [in] void *pvDestContext,
        [in] MSHLFLAGS mshlflags);

    HRESULT UnmarshalInterface(
        [in] IStream *pStm,
        [in] UUID *riid,
        [out] void *ppv);

    HRESULT ReleaseMarshalData(
        [in] IStream *pStm);

    HRESULT DisconnectObject(
        [in] LONG dwReserved);

}
