[
    odl,
    uuid(0000010D-0000-0000-C000-000000000046)
]
interface IViewObject : stdole.IUnknown {

    HRESULT Draw(
        [in] DVASPECT dwDrawAspect,
        [in] long lindex,
        [in] void *pvAspect,
        [in] void *ptd,
        [in] long hdcTargetDev,
        [in] long hdcDraw,
        [in] RECT *lprcBounds,
        [in] RECT *lprcWBounds,
        [in] long pfnContinue,
        [in] long dwContinue);

    HRESULT GetColorSet(
        [in] DVASPECT dwDrawAspect,
        [in] long lindex,
        [in] void *pvAspect,
        [in] void *ptd,
        [in] long hicTargetDev,
        [in, out] long *ppColorSet); //LOGPALETTE ** to pointer

    HRESULT Freeze(
        [in] DVASPECT dwDrawAspect,
        [in] long lindex,
        [in] void *pvAspect,
        [out, retval] long* pdwFreeze);

    HRESULT Unfreeze([in] long dwFreeze);

    HRESULT SetAdvise(
        [in] DVASPECT aspects,
        [in] long advf,
        [in] long pIAdviseSink);

    HRESULT GetAdvise(
        [out] DVASPECT *pAspects,
        [out] long *pAdvf,
        [out] long *pIAdviseSink);

};

[
    odl,
    uuid(00000127-0000-0000-C000-000000000046)
]
interface IViewObject2 : IViewObject {

    HRESULT GetExtent(
        [in] DVASPECT dwDrawAspect,
        [in] long lindex,
        [in] DVTARGETDEVICE* ptd,
        [out] SIZE* lpSIZE);

};
