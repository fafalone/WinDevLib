[
  odl,
  uuid(0000010E-0000-0000-C000-000000000046)
]
interface IDataObject : stdole.IUnknown {

    HRESULT GetData(
        [in, out] FORMATETC *pformatetcIn,
        [in, out] STGMEDIUM *pmedium);

    HRESULT GetDataHere(
        [in, out] FORMATETC *pformatetc,
        [in, out] STGMEDIUM *pmedium);

    HRESULT QueryGetData(
        [in, out] FORMATETC* pformatetc);

    HRESULT GetCanonicalFormatEtc(
        [in, out] FORMATETC *pformatectIn,
        [in, out] FORMATETC *pformatetcOut);

    HRESULT SetData(
        [in, out] FORMATETC *pformatetc,
        [in, out] STGMEDIUM *pmedium,
        [in] LONG fRelease);

    HRESULT EnumFormatEtc(
        [in] DATADIR dwDirection,
        [in, out] IEnumFORMATETC **ppenumFormatEtc);

    HRESULT DAdvise(
        [in, out] FORMATETC *pformatetc,
        [in] long advf,
        [in] long pAdvSink,
        [in, out] long *pdwConnection);

    HRESULT DUnadvise(
        [in] long dwConnection);

    HRESULT EnumDAdvise(
        [in, out] long* pIEnumAdvise);

};
