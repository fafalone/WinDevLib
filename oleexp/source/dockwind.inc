typedef enum DESKBANDINFO_Masks {
    DBIM_MINSIZE    = 0x0001,
    DBIM_MAXSIZE    = 0x0002,
    DBIM_INTEGRAL   = 0x0004,
    DBIM_ACTUAL     = 0x0008,
    DBIM_TITLE      = 0x0010,
    DBIM_MODEFLAGS  = 0x0020,
    DBIM_BKCOLOR    = 0x0040,
} DESKBANDINFO_Masks;

typedef enum DESKBANDINFO_ModeFlags {
    DBIMF_NORMAL            = 0x0000,
    DBIMF_FIXED             = 0x0001,
    DBIMF_FIXEDBMP          = 0x0004,
    DBIMF_VARIABLEHEIGHT    = 0x0008,
    DBIMF_UNDELETEABLE      = 0x0010,
    DBIMF_DEBOSSED          = 0x0020,
    DBIMF_BKCOLOR           = 0x0040,
    DBIMF_USECHEVRON        = 0x0080,
    DBIMF_BREAK             = 0x0100,
    DBIMF_ADDTOFRONT        = 0x0200,
    DBIMF_TOPALIGN          = 0x0400,
} DESKBANDINFO_ModeFlags;

typedef struct {
    DESKBANDINFO_Masks dwMask;
    POINT ptMinSize;
    POINT ptMaxSize;
    POINT ptIntegral;
    POINT ptActual;
    BYTE wszTitle[512];
    DESKBANDINFO_ModeFlags dwModeFlags;
    LONG crBkgnd;
} DESKBANDINFO;

typedef enum GetBandInfo_ViewModes {
    DBIF_VIEWMODE_NORMAL         = 0x0000,
    DBIF_VIEWMODE_VERTICAL       = 0x0001,
    DBIF_VIEWMODE_FLOATING       = 0x0002,
    DBIF_VIEWMODE_TRANSPARENT    = 0x0004,
} GetBandInfo_ViewModes;

[
    odl,
    uuid(68284faa-6a48-11d0-8c78-00c04fd918b4)
]
interface IInputObject : stdole.IUnknown {

    [helpstring("Activates or deactivates the object.  lpMsg may be NULL.\nReturns S_OK if the activation succeeded.")]
    HRESULT UIActivateIO(
        [in] BOOL fActivate,
        [in] MSG *lpMsg);

    [helpstring("Returns S_OK if the object has the focus, S_FALSE if not")]
    HRESULT HasFocusIO(void);

    [helpstring("Allow the object to process the message.\nReturns S_OK if the message was processed (eaten).")]
    HRESULT TranslateAcceleratorIO(
        [in] MSG *lpMsg);

};

[
    odl,
    uuid(EB0FE172-1A3A-11D0-89B3-00A0C90A90AC)
]
interface IDeskBand : stdole.IUnknown {

    // IOleWindow members
    HRESULT GetWindow(
        [out, retval] LONG *phwnd);

    HRESULT ContextSensitiveHelp(
        [in] BOOL fEnterMode);

    // IDockingWindow members
    HRESULT ShowDW(
        [in] BOOL fShow);

    HRESULT CloseDW(
        [in] LONG dwReserved);

    HRESULT ResizeBorderDW(
        [in, out] RECT *prcBorder,
        [in] LONG punkToolbarSite,
        [in] BOOL fReserved);

    // IDeskBand members
    HRESULT GetBandInfo(
        [in] LONG dwBandID,
        [in] GetBandInfo_ViewModes dwViewMode,
        [in, out] DESKBANDINFO *pdbi);
}

[
    odl,
    uuid(012dd920-7b26-11d0-8ca9-00a0c92dbfe8)
]
interface IDockingWindow : stdole.IUnknown {

    // IOleWindow members
    HRESULT GetWindow(
        [out, retval] LONG *phwnd);

    HRESULT ContextSensitiveHelp(
        [in] BOOL fEnterMode);

    // IDockingWindow members
    HRESULT ShowDW(
        [in] BOOL fShow);

    HRESULT CloseDW(
        [in] LONG dwReserved);

    HRESULT ResizeBorderDW(
        [in, out] RECT *prcBorder,
        [in] LONG punkToolbarSite,
        [in] BOOL fReserved);
}

[
    odl,
    uuid(2a342fc2-7b26-11d0-8ca9-00a0c92dbfe8)
]
interface IDockingWindowSite : IOleWindow {

    HRESULT GetBorderDW(
        [in] stdole.IUnknown *punkObj, 
        [in, out] RECT *prcBorder);

    HRESULT RequestBorderSpaceDW(
        [in] stdole.IUnknown *punkObj,
        [in, out] RECT *pbw);

    HRESULT SetBorderSpaceDW(
        [in] stdole.IUnknown *punkObj, 
        [in, out] RECT *pbw);
}

// Command Target IDs
typedef enum DBID {
    DBID_BANDINFOCHANGED    = 0,
    DBID_SHOWONLY           = 1,
    DBID_MAXIMIZEBAND       = 2,      // Maximize the specified band (VT_UI4 == dwID)
    DBID_PUSHCHEVRON        = 3,
    DBID_DELAYINIT          = 4,      // Note: _bandsite_ calls _band_ with this code
    DBID_FINISHINIT         = 5,      // Note: _bandsite_ calls _band_ with this code
    DBID_SETWINDOWTHEME     = 6,      // Note: _bandsite_ calls _band_ with this code
    DBID_PERMITAUTOHIDE     = 7,
} DBID;

// flags for RemoveToolbar
typedef enum DWFRF {
	DWFRF_NORMAL            = 0x0000,
	DWFRF_DELETECONFIGDATA  = 0x0001,
} DWFRF;

// flags for AddToolbar
typedef enum DWFAF {
	DWFAF_HIDDEN  = 0x0001
} DWFAF;

[
    odl,
    uuid(47d2657a-7b27-11d0-8ca9-00a0c92dbfe8)
]
interface IDockingWindowFrame : IOleWindow {
 
    HRESULT AddToolbar(
        [in] stdole.IUnknown *punkSrc,
        [in] LPWSTR pwszItem,
        [in, defaultvalue(0)] DWFAF dwAddFlags);

    HRESULT RemoveToolbar(
        [in] stdole.IUnknown *punkSrc,
        [in] DWFRF dwRemoveFlags);

    HRESULT FindToolbar(
        [in] LPWSTR pwszItem,
        [in] UUID *riid,
        [out] void **ppv);
}

[
    odl,
    uuid(f1db8392-7331-11d0-8c99-00a0c92dbfe8)
]
interface IInputObjectSite : stdole.IUnknown {

    HRESULT OnFocusChangeIS(
        [in] stdole.IUnknown *punkObj,
        [in] BOOL fSetFocus);

}