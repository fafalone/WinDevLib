
[
    odl,
    uuid(00000122-0000-0000-C000-000000000046),
]
interface IDropTarget : stdole.IUnknown {

    HRESULT DragEnter(
        [in] IDataObject *pDataObj,
        [in] LONG grfKeyState,
        [in] LONG ptX,
        [in] LONG ptY,
        [in, out] DROPEFFECTS *pdwEffect);

    HRESULT DragOver(
        [in] LONG grfKeyState,
        [in] LONG ptX,
        [in] LONG ptY,
        [in, out] DROPEFFECTS *pdwEffect);

    HRESULT DragLeave();

    HRESULT Drop(
        [in] IDataObject *pDataObj,
        [in] LONG grfKeyState,
        [in] LONG ptX,
        [in] LONG ptY,
        [in, out] DROPEFFECTS *pdwEffect);
}

[
    odl,
    uuid(00000121-0000-0000-C000-000000000046)
]
interface IDropSource : stdole.IUnknown {

    HRESULT QueryContinueDrag(
        [in] BOOL fEscapePressed,
        [in] LONG grfKeyState);

    HRESULT GiveFeedback(
        [in] DROPEFFECTS dwEffect);
}

[
  odl,
  uuid(0000010E-0000-0000-C000-000000000046)
]
interface IDataObject : stdole.IUnknown {

    LONG GetData(
        [in, out] FORMATETC *pformatetcIn,
        [in, out] STGMEDIUM *pmedium);

    LONG GetDataHere(
        [in, out] FORMATETC *pformatetc,
        [in, out] STGMEDIUM *pmedium);

    LONG QueryGetData(
        [in, out] FORMATETC* pformatetc);

    HRESULT GetCanonicalFormatEtc(
        [in, out] FORMATETC *pformatectIn,
        [in, out] FORMATETC *pformatetcOut);

    HRESULT SetData(
        [in, out] FORMATETC *pformatetc,
        [in, out] STGMEDIUM *pmedium,
        [in] BOOL fRelease);

    HRESULT EnumFormatEtc(
        [in] DATADIR dwDirection,
        [out, retval] IEnumFORMATETC **ppenumFormatEtc);

    HRESULT DAdvise(
        [in, out] FORMATETC *pformatetc,
        [in] long advf,
        [in] IAdviseSink *pAdvSink,
        [out, retval] long *pdwConnection);

    HRESULT DUnadvise(
        [in] long dwConnection);

    HRESULT EnumDAdvise(
        [out, retval] IEnumSTATDATA *pIEnumAdvise);

};

[
    odl,
    uuid(0000010f-0000-0000-C000-000000000046),
]
interface IAdviseSink : stdole.IUnknown {

    void OnDataChange(
        [in, out] FORMATETC *pFormatetc,
        [in, out] STGMEDIUM *pStgmed);

    void OnViewChange(
        [in] DVASPECT dwAspect,
        [in] LONG lindex);

    void OnRename(
        [in] LONG pmk);

    void OnSave();

    void OnClose();
}
