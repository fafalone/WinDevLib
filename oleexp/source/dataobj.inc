//tbShellLib - 100% implemented

typedef enum DV_ERROR {
	DV_E_FORMATETC = 0x80040064,
	DV_E_DVTARGETDEVICE = 0x80040065,
	DV_E_STGMEDIUM = 0x80040066,
	DV_E_STATDATA = 0x80040067,
	DV_E_LINDEX = 0x80040068,
	DV_E_TYMED = 0x80040069,
	DV_E_CLIPFORMAT = 0x8004006A,
	DV_E_DVASPECT = 0x8004006B,
	DV_E_DVTARGETDEVICE_SIZE = 0x8004006C,
	DV_E_NOIVIEWOBJECT = 0x8004006D
} DV_ERROR;


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
