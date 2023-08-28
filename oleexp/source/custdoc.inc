
[
    odl,
    uuid(bd3f23c0-d43e-11cf-893b-00aa00bdce1a),
]
interface IDocHostUIHandler : stdole.IUnknown{

    HRESULT ShowContextMenu(
        [in] ContextMenuTarget dwContext,
        [in, out] POINT *pPOINT,
        [in] IOleCommandTarget *pCommandTarget,
        [in] IDispatch *HTMLTagElement);

    HRESULT GetHostInfo(
        [in, out] DOCHOSTUIINFO * pInfo);

    HRESULT ShowUI(
        [in] LONG dwID,
        [in] IOleInPlaceActiveObject *pActiveObject,
        [in] IOleCommandTarget *pCommandTarget,
        [in] IOleInPlaceFrame *pFrame,
        [in] IOleInPlaceUIWindow *pDoc);

    HRESULT HideUI();

    HRESULT UpdateUI();

    HRESULT EnableModeless(
        [in] BOOL fEnable);

    HRESULT OnDocWindowActivate(
        [in] BOOL fActivate);

    HRESULT OnFrameWindowActivate(
        [in] BOOL fActivate);

    HRESULT ResizeBorder(
        [in, out] RECT *prcBorder,
        [in] IOleInPlaceUIWindow *pUIWindow,
        [in] BOOL fRameWindow);

    HRESULT TranslateAccelerator(
        [in, out] MSG *lpMsg,
        [in, out] UUID *pguidCmdGroup,
        [in] LONG nCmdID);

    HRESULT GetOptionKeyPath(
        [in, out] LONG *pOLESTRchKey,
        [in] LONG dw);

    HRESULT GetDropTarget(
        [in] IDropTarget *pDropTarget,
        [out, retval] IDropTarget **ppDropTarget);

    HRESULT GetExternal(
        [out, retval] IDispatch** ppDispatch);

    HRESULT TranslateUrl(
        [in] LONG dwTranslate,
        [in] LONG pchURLIn,
        [out, retval] LONG *ppchURLOut);

    HRESULT FilterDataObject(
        [in] IDataObject *pDO,
        [out, retval] IDataObject **ppDORet);
}

[
    odl,
    uuid(3050f6d0-98b5-11cf-bb82-00aa00bdce0b),
]
interface IDocHostUIHandler2 : stdole.IUnknown {

    HRESULT ShowContextMenu(
        [in] ContextMenuTarget dwContext,
        [in, out] POINT *pPOINT,
        [in] IOleCommandTarget *pCommandTarget,
        [in] IDispatch *HTMLTagElement);

    HRESULT GetHostInfo(
        [in, out] DOCHOSTUIINFO * pInfo);

    HRESULT ShowUI(
        [in] LONG dwID,
        [in] IOleInPlaceActiveObject *pActiveObject,
        [in] IOleCommandTarget *pCommandTarget,
        [in] IOleInPlaceFrame *pFrame,
        [in] IOleInPlaceUIWindow *pDoc);

    HRESULT HideUI();

    HRESULT UpdateUI();

    HRESULT EnableModeless(
        [in] BOOL fEnable);

    HRESULT OnDocWindowActivate(
        [in] BOOL fActivate);

    HRESULT OnFrameWindowActivate(
        [in] BOOL fActivate);

    HRESULT ResizeBorder(
        [in, out] RECT *prcBorder,
        [in] IOleInPlaceUIWindow *pUIWindow,
        [in] BOOL fRameWindow);

    HRESULT TranslateAccelerator(
        [in, out] MSG *lpMsg,
        [in, out] UUID *pguidCmdGroup,
        [in] LONG nCmdID);

    HRESULT GetOptionKeyPath(
        [in, out] LONG *pOLESTRchKey,
        [in] LONG dw);

    HRESULT GetDropTarget(
        [in] IDropTarget *pDropTarget,
        [out, retval] IDropTarget **ppDropTarget);

	HRESULT GetExternal(
        [out, retval] IDispatch** ppDispatch);

   	HRESULT TranslateUrl(
        [in] LONG dwTranslate,
        [in] LONG pchURLIn,
        [out, retval] LONG *ppchURLOut);

    HRESULT FilterDataObject(
        [in] IDataObject *pDO,
        [out, retval] IDataObject **ppDORet);

     HRESULT GetOverrideKeyPath(
         [in, out] LONG *pchKey,
         [in] LONG dw);
}

[
    odl,
    uuid(3050f3f0-98b5-11cf-bb82-00aa00bdce0b)
]
interface ICustomDoc : stdole.IUnknown{

    HRESULT SetUIHandler(
        [in] IDocHostUIHandler *pUIHandler);
}

[
    odl,
    uuid(c4d244b0-d43e-11cf-893b-00aa00bdce1a)
]
interface IDocHostShowUI : stdole.IUnknown {

    HRESULT ShowMessage(
        [in] LONG hwnd,
        [in] LONG lpszText,
        [in] LONG lpszCaption,
        [in] LONG dwType,
        [in] LONG lpszHelpFile,
        [in] LONG dwHelpContext,
        [out, retval] LONG *plResult);

    HRESULT ShowHelp(
        [in] LONG hwnd,
        [in] LONG pszHelpFile,
        [in] LONG uCommand,
        [in] LONG dwData,
        [in] LONG x,
        [in] LONG y,
        [in] IDispatch *pDispatchObjectHit);

}
