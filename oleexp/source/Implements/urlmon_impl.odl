[
    odl,
    uuid(79eac9e5-baf9-11ce-8c82-00aa004ba90b)
]
interface IInternetProtocolSink : stdole.IUnknown {

    HRESULT Switch(
       [in, out] PROTOCOLDATA *pProtocolData);

    HRESULT ReportProgress(
       [in] BINDSTATUS ulStatusCode,
       [in] LONG szStatusText);

    HRESULT ReportData(
       [in] BSCF grfBSCF,
       [in] LONG ulProgress,
       [in] LONG ulProgressMax);

    HRESULT ReportResult(
       [in] LONG hrResult,
       [in] LONG dwError,
       [in] LONG szResult);
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

    HRESULT Read(
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

