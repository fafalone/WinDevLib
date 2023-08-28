typedef enum {
    REO_GETOBJ_NO_INTERFACES = 0,
    REO_GETOBJ_POLEOBJ = 1,
    REO_GETOBJ_PSTG = 2,
    REO_GETOBJ_POLESITE = 4,
    REO_GETOBJ_ALL_INTERFACES = 7
} REO_GETOBJ;

typedef enum {
    REO_IOB_SELECTION = 0xffffffff,
    REO_CP_SELECTION = 0xffffffff,
    REO_IOB_USE_CP = 0xfffffffe
} REO_IOB;

typedef enum {
    REO_NULL = 0,
    REO_RESIZABLE = 1,
    REO_BELOWBASELINE = 2,
    REO_INVERTEDSELECT = 4,
    REO_DYNAMICSIZE = 8,
    REO_BLANK = 16,
    REO_DONTNEEDPALETTE = 32,
    REO_READWRITEMASK = 63,
    REO_GETMETAFILE = 0x00400000,
    REO_LINKAVAILABLE = 0x00800000,
    REO_HILITED = 0x01000000,
    REO_INPLACEACTIVE = 0x02000000,
    REO_OPEN = 0x04000000,
    REO_SELECTED = 0x08000000,
    REO_STATIC = 0x40000000,
    REO_LINK = 0x80000000
} REO_FLAGS;

typedef enum {
    RECO_PASTE = 0,
    RECO_DROP = 1,
    RECO_COPY = 2,
    RECO_CUT = 3,
    RECO_DRAG = 4
} RECO_FLAGS;

typedef struct REOBJECT {
    long cbStruct;
    long cp;
    UUID clsid;
    IOleObject* poleobj;
    IStorage* pstg;
    IOleClientSite* polesite;
    SIZE sizel;
    DVASPECT DVASPECT;
    REO_FLAGS dwFlags;
    long dwUser;
} REOBJECT;

typedef struct CHARRANGE {
    long start;
    long end;
} CHARRANGE;

[
    odl,
    uuid(00020D00-0000-0000-C000-000000000046)
]
interface IRichEditOle : stdole.IUnknown {

    HRESULT GetClientSite(
        [out, retval] IOleClientSite **lplpolesite);

    long GetObjectCount();

    long GetLinkCount();

    long GetObject(
        [in] long iob, 
        [in] REOBJECT *lpreobject, 
        [in] REO_GETOBJ dwFlags);

    long InsertObject(
        [in] REOBJECT *lpreobject);

    long ConvertObject(
        [in] long iob, 
        [in] UUID *rclsidNew, 
        [in] LPSTR lpstrUserTypeNew);

    long ActivateAs(
        [in] UUID *rclsid, 
        [in] UUID *rclsidAs);

    long SetHostNames(
        [in] LPSTR lpstrContainerApp, 
        [in] LPSTR lpstrContainerObj);

    long SetLinkAvailable(
        [in] long iob, 
        [in] long fAvailable);

    long SetDvaspect(
        [in] long iob, 
        [in] DVASPECT DVASPECT);

    long HandsOffStorage(
        [in] long iob);

    long SaveCompleted(
        [in] long iob, 
        [in] IStorage *lpstg);

    long InPlaceDeactivate();

    long ContextSensitiveHelp(
        [in] long fEnterMode);

    long GetClipboardData(
        [in] CHARRANGE *lpchrg, 
        [in] RECO_FLAGS reco, 
        [out] IDataObject **lplpdataobj);

    long ImportDataObject(
        [in] IDataObject* lpdataobj, 
        [in] short cf, 
        [in] long hMetaPict);
};

[
    odl,
    uuid(00020D03-0000-0000-C000-000000000046)
]
interface IRichEditOleCallback : stdole.IUnknown {

    HRESULT GetNewStorage(
        [in, out] IStorage **lplpstg);

    HRESULT GetInPlaceContext(
        [in, out] IOleInPlaceFrame **lplpFrame,
        [in, out] IOleInPlaceUIWindow **lplpDoc,
        [in, out] OLEINPLACEFRAMEINFO *lpFrameInfo);

    HRESULT ShowContainerUI(
        [in] LONG fShow);

    HRESULT QueryInsertObject(
        [in, out] UUID *lpclsid,
        [in] IStorage *lpstg,
        [in] LONG cp);

    HRESULT DeleteObject(
        [in] IOleObject *lpoleobj);

    HRESULT QueryAcceptData(
        [in] IDataObject *lpdataobj,
        [in, out] short *lpcfFormat,
        [in] LONG reco,
        [in] LONG fReally,
        [in] LONG hMetaPict);

    HRESULT ContextSensitiveHelp(
        [in] LONG fEnterMode);

    HRESULT GetClipboardData(
        [in, out] CHARRANGE *lpchrg,
        [in] LONG reco,
        [in, out] IDataObject **lplpdataobj);

    HRESULT GetDragDropEffect(
        [in] LONG fDrag,
        [in] LONG grfKeyState,
        [in, out] LONG *pdwEffect);

    HRESULT GetContextMenu(
        [in] short seltype,
        [in] IOleObject *lpoleobj,
        [in, out] CHARRANGE *lpchrg,
        [in, out] LONG *lphmenu);

};

