    [
      odl,
      uuid(0000010D-0000-0000-C000-000000000046)
    ]
    interface IViewObject : stdole.IUnknown {
        HRESULT Draw(
            [in] DVASPECT dwDrawAspect,
            [in] long lindex,
            [in] long pvAspect,
            [in] long ptd,
            [in] long hdcTargetDev,
            [in] long hdcDraw,
            [in, out] RECT *lprcBounds,
            [in, out] RECT *lprcWBounds,
            [in] long pfnContinue,
            [in] long dwContinue);

        HRESULT GetColorSet(
            [in] DVASPECT dwDrawAspect,
            [in] long lindex,
            [in] long pvAspect,
            [in] long ptd,
            [in] long hicTargetDev,
            [in, out] long *ppColorSet);

        HRESULT Freeze(
            [in] DVASPECT dwDrawAspect,
            [in] long lindex,
            [in] long pvAspect,
            [in, out] long *pdwFreeze);

        HRESULT Unfreeze(
            [in] long dwFreeze);

        HRESULT SetAdvise(
            [in] DVASPECT aspects,
            [in] long advf,
            [in] IAdviseSink *pIAdviseSink);

        HRESULT GetAdvise(
            [in, out] DVASPECT *pAspects,
            [in, out] long *pAdvf,
            [in, out] IAdviseSink **pIAdviseSink);

    };
