[
    odl,
    uuid(00000114-0000-0000-C000-000000000046),
    restricted
]
interface IOleWindow : stdole.IUnknown {

    [restricted]
    HRESULT GetWindow(
        [out, retval] long *phwnd);

    [restricted]
    HRESULT ContextSensitiveHelp(
        [in] BOOL fEnterMode);

};


[
    odl,
    uuid(00000119-0000-0000-C000-000000000046)
]
interface IOleInPlaceSite : stdole.IUnknown {

    HRESULT GetWindow(
        [out, retval] long *phwnd);

    HRESULT ContextSensitiveHelp(
        [in] BOOL fEnterMode);

    HRESULT CanInPlaceActivate();

    HRESULT OnInPlaceActivate();

    HRESULT OnUIActivate();

    HRESULT GetWindowContext(
        [in, out] IOleInPlaceFrame** ppFrame,
        [in, out] IOleInPlaceUIWindow** ppDoc,
        [in, out] RECT* lprcPosRect,
        [in, out] RECT* lprcClipRect,
        [in, out] OLEINPLACEFRAMEINFO* lpFrameInfo);

    HRESULT Scroll(
        [in] LONG scrollX,
        [in] LONG scrollY);

    HRESULT OnUIDeactivate([in] BOOL fUndoable);

    HRESULT OnInPlaceDeactivate();

    HRESULT DiscardUndoState();

    HRESULT DeactivateAndUndo();

    HRESULT OnPosRectChange([in, out] RECT* lprcPosRect);

};

[
    odl,
    uuid(00000117-0000-0000-C000-000000000046)
]
interface IOleInPlaceActiveObject : stdole.IUnknown {

    HRESULT GetWindow(
        [out, retval] long *phwnd);

    HRESULT ContextSensitiveHelp(
        [in] BOOL fEnterMode);

    HRESULT TranslateAccelerator(
        [in, out] MSG* lpmsg);

    HRESULT OnFrameWindowActivate(
        [in] BOOL fActivate);

    HRESULT OnDocWindowActivate(
        [in] BOOL fActivate);

    HRESULT ResizeBorder(
        [in, out] RECT* prcBorder,
        [in] IOleInPlaceUIWindow *pUIWindow,
        [in] BOOL fFrameWindow);

    HRESULT EnableModeless(
        [in] BOOL fEnable);
};
