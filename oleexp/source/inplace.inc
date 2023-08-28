//tbShellLib - 100% implemented

[
    odl,
    uuid(00000115-0000-0000-C000-000000000046)
]
interface IOleInPlaceUIWindow : IOleWindow {

    LONG GetBorder(
        [in, out] RECT* lprectBorder);

    LONG RequestBorderSpace(
        [in, out] RECT* pborderwidths);

    LONG SetBorderSpace(
        [in, out] RECT* pborderwidths);

    LONG SetActiveObject(
        [in] IOleInPlaceActiveObject* pActiveObject,
        [in] LONG pszObjName);
};

[
    odl,
    uuid(00000117-0000-0000-C000-000000000046)
]
interface IOleInPlaceActiveObject : IOleWindow {

    long TranslateAccelerator(
        [in] MSG* lpmsg);

    long OnFrameWindowActivate(
        [in] BOOL fActivate);

    long OnDocWindowActivate(
        [in] BOOL fActivate);

    long ResizeBorder(
        [in] RECT* prcBorder,
        [in] IOleInPlaceUIWindow *pUIWindow,
        [in] BOOL fFrameWindow);

    long EnableModeless(
        [in] BOOL fEnable);
};

[
    odl,
    uuid(00000119-0000-0000-C000-000000000046)
]
interface IOleInPlaceSite : IOleWindow {

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
    uuid(00000116-0000-0000-C000-000000000046),
]
interface IOleInPlaceFrame : IOleInPlaceUIWindow {
    LONG InsertMenus(
        [in] LONG hmenuShared,
        [in, out] OLEMENUGROUPWIDTHS *lpMenuWidths);
    LONG SetMenu(
        [in] LONG hmenuShared,
        [in] LONG holemenu,
        [in] LONG hwndActiveObject);
    LONG RemoveMenus(
        [in] LONG hmenuShared);
    LONG SetStatusText(
        [in] LONG pszStatusText);
    LONG EnableModeless(
        [in] BOOL fEnable);
    LONG TranslateAccelerator(
        [in, out] MSG *lpmsg,
        [in] short wID);
};

[
    odl,
    uuid(00000113-0000-0000-C000-000000000046)
]
interface IOleInPlaceObject : IOleWindow {

    HRESULT InPlaceDeactivate(void);

    HRESULT UIDeactivate(void);

    HRESULT SetObjectRects(
        [in] RECT *lprcPosRect,
        [in] RECT *lprcClipRect);

    HRESULT ReactivateAndUndo(void);
}

[
    odl,
    uuid(B196B289-BAB4-101A-B69C-00AA00341D07)
]
interface IOleControlSite : stdole.IUnknown {

    typedef struct POINTF {
        FLOAT x;
        FLOAT y;
    } POINTF;

    typedef enum XFORMCOORDS {
        XFORMCOORDS_POSITION            = 0x1,
        XFORMCOORDS_SIZE                = 0x2,
        XFORMCOORDS_HIMETRICTOCONTAINER = 0x4,
        XFORMCOORDS_CONTAINERTOHIMETRIC = 0x8,
        XFORMCOORDS_EVENTCOMPAT         = 0x10
    } XFORMCOORDS;

    HRESULT OnControlInfoChanged(void);

    HRESULT LockInPlaceActive(
        [in] BOOL fLock);

    HRESULT GetExtendedControl(
        [out, retval] IDispatch **ppDisp);

    HRESULT TransformCoords(
        [in, out] POINT *pPtlHimetric,
        [in, out] POINTF *pPtfContainer,
        [in] XFORMCOORDS dwFlags);

    typedef enum KEYMODIFIERS {
        KEYMOD_SHIFT      = 0x00000001,
        KEYMOD_CONTROL    = 0x00000002,
        KEYMOD_ALT        = 0x00000004
    } KEYMODIFIERS;

    HRESULT TranslateAccelerator(
        [in, out] MSG *pMsg,
        [in] KEYMODIFIERS grfModifiers);

    HRESULT OnFocus(
        [in] BOOL fGotFocus);

    HRESULT ShowPropertyFrame(void);

}
