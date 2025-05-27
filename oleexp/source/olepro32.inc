//tbShellLib - 100% Implemented

typedef struct tagOCPFIPARAMS {
    LONG cbStructSize;
    LONG hWndOwner;
    LONG x;
    LONG y;
    LONG lpszCaption;
    LONG cObjects;
    LONG lplpUnk;
    LONG cPages;
    LONG lpPagesGUID;
    LONG lcid;
    LONG dispidInitialProperty;
} OCPFIPARAMS;

// #define FONTSIZE(n) { n##0000, 0 }

typedef struct tagFONTDESC {
    LONG cbSizeofstruct;
    LONG lpstrName;
    CURRENCY cySize;
    short sWeight;
    short sCharset;
    BOOL fItalic;
    BOOL fUnderline;
    BOOL fStrikethrough;
} FONTDESC;

typedef enum PICTYPE {
    PICTYPE_UNINITIALIZED   = -1,
    PICTYPE_NONE            = 0,
    PICTYPE_BITMAP          = 1,
    PICTYPE_METAFILE        = 2,
    PICTYPE_ICON            = 3,
    PICTYPE_ENHMETAFILE     = 4,
} PICTYPE;

typedef struct tagPICTDESC {
    LONG cbSizeofstruct;
    PICTYPE picType;
    LONG hbitmap;
    LONG hpal_or_xExt;
    LONG yExt;
} PICTDESC;

[
    dllname("OLEPRO32.DLL")
]
module olepro32 {

    const LONG OLEIVERB_PROPERTIES = -7;

    [entry("OleCreatePropertyFrame")]
    HRESULT OleCreatePropertyFrame(
        [in] LONG hwndOwner,
        [in] LONG x,
        [in] LONG y,
        [in] LPWSTR lpszCaption,
        [in] LONG cObjects,
        [in] IUnknown **ppUnk,
        [in] LONG cPages,
        [in] void *pPageClsID,
        [in] LONG lcid,
        [in, defaultvalue(0)] LONG dwReserved,
        [in, defaultvalue(0)] LONG pvReserved);

    [entry("OleCreatePropertyFrameIndirect")]
    HRESULT OleCreatePropertyFrameIndirect(
        [in] OCPFIPARAMS *lpParams);

    // Standard type APIs
    [entry("OleCreateFontIndirect")]
    HRESULT OleCreateFontIndirect(
        [in] FONTDESC *lpFontDesc,
        [in] UUID *riid,
        [out, retval] IFont **lplpvObj);

    [entry("OleCreatePictureIndirect")]
    HRESULT OleCreatePictureIndirect(
        [in] PICTDESC *lpPictDesc,
        [in] UUID *riid,
        [in] BOOL fOwn,
        [out, retval] IPicture **lplpvObj);

    [entry("OleIconToCursor")]
    LONG OleIconToCursor(
        [in] LONG hinstExe,
        [in] LONG hIcon);
        
}