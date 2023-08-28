[
    odl,
    uuid(b722bcc5-4e68-101b-a2bc-00aa00404770),
]
interface IOleDocument : stdole.IUnknown {
        
	typedef enum {
		DOCMISC_CANCREATEMULTIPLEVIEWS          = 1,
		DOCMISC_SUPPORTCOMPLEXRECTANGLES        = 2,
		DOCMISC_CANTOPENEDIT                    = 4, // fails the IOleDocumentView::Open  method
		DOCMISC_NOFILESUPPORT                   = 8, //  does not support read/writing to a file
	} DOCMISC;

	HRESULT CreateView(
		[in] IOleInPlaceSite *pIPSite,
		[in] IStream *pstm,
		[in] LONG dwReserved,
		[out, retval] IOleDocumentView **ppView);

	HRESULT GetDocMiscStatus(
		[out, retval] LONG *pdwStatus);

	HRESULT EnumViews(
		[out] IEnumOleDocumentViews **ppEnum,
		[out] IOleDocumentView **ppView);
}

[
	odl,
	uuid(b722bcc7-4e68-101b-a2bc-00aa00404770),
]
interface IOleDocumentSite : stdole.IUnknown {

	HRESULT ActivateMe(
		[in] IOleDocumentView *pViewToActivate);
		
}

[
	odl,
	uuid(b722bcc6-4e68-101b-a2bc-00aa00404770),
]
interface IOleDocumentView : stdole.IUnknown {

	HRESULT SetInPlaceSite(
		[in] IOleInPlaceSite *pIPSite);

	HRESULT GetInPlaceSite(
		[out, retval] IOleInPlaceSite **ppIPSite);

	HRESULT GetDocument(
		[out, retval] stdole.IUnknown **ppunk);

	HRESULT SetRect(
		[in] RECT *prcView);

	HRESULT GetRect(
		[out] RECT *prcView);

	HRESULT SetRectComplex(
		[in] RECT *prcView,
		[in] RECT *prcHScroll,
		[in] RECT *prcVScroll,
		[in] RECT *prcSizeBox);

	HRESULT Show(
		[in] BOOL fShow);

	HRESULT UIActivate(
		[in] BOOL fUIActivate);

	HRESULT Open();

	HRESULT CloseView(
		[in, defaultvalue(0)] LONG dwReserved);

	HRESULT SaveViewState(
		[in] IStream *pstm);

	HRESULT ApplyViewState(
		[in] IStream *pstm);

	HRESULT Clone(
		[in] IOleInPlaceSite *pIPSiteNew,
		[out, retval] IOleDocumentView **ppViewNew);
}

[
	odl,
	uuid(b722bcc8-4e68-101b-a2bc-00aa00404770),
]
interface IEnumOleDocumentViews : stdole.IUnknown {

	LONG __stdcall Next(
        [in] LONG cViews,
        [out] IOleDocumentView **rgpView,
		[out, defaultvalue(0)] ULONG *pcFetched);

	HRESULT Skip(
		[in] LONG cViews);

	HRESULT Reset();

	HRESULT Clone(
		[out, retval] IEnumOleDocumentViews **ppEnum);
}

[
	odl,
	uuid(b722bcca-4e68-101b-a2bc-00aa00404770),
]
interface IContinueCallback : stdole.IUnknown {

	HRESULT FContinue();

	HRESULT FContinuePrinting(
		[in] LONG nCntPrinted,
		[in] LONG nCurPage,
		[in] LONG pwszPrintStatus);
}

[
	odl,
    uuid(b722bcc9-4e68-101b-a2bc-00aa00404770),
]
interface IPrint : stdole.IUnknown {

    typedef enum PRINTFLAG {
        PRINTFLAG_MAYBOTHERUSER         = 1,
        PRINTFLAG_PROMPTUSER            = 2,
        PRINTFLAG_USERMAYCHANGEPRINTER  = 4,
        PRINTFLAG_RECOMPOSETODEVICE     = 8,
        PRINTFLAG_DONTACTUALLYPRINT     = 16,
        PRINTFLAG_FORCEPROPERTIES       = 32,
        PRINTFLAG_PRINTTOFILE           = 64
    } PRINTFLAG;

    typedef struct PAGERANGE {
        LONG nFromPage;
        LONG nToPage;
    } PAGERANGE;

	typedef struct PAGESET {
        LONG cbStruct;
        BOOL fOddPages;
        BOOL fEvenPages;
        LONG cPageRange;
        PAGERANGE rgPages[1];
    } PAGESET;

	// #define PAGESET_TOLASTPAGE   ((WORD)(-1L))

	HRESULT SetInitialPageNum(
		[in] LONG nFirstPage);

	HRESULT GetPageInfo(
		[out] LONG *pnFirstPage,
		[out] LONG *pcPages);

    HRESULT __stdcall Print(
        [in] PRINTFLAG grfFlags,
        [in, out] DVTARGETDEVICE **pptd,
        [in, out] PAGESET **ppPageSet,
        [in, out] STGMEDIUM *pstgmOptions,
        [in] IContinueCallback *pcallback,
        [in]  LONG nFirstPage,
        [out] LONG *pcPagesPrinted,
        [out] LONG *pnLastPage);

}
