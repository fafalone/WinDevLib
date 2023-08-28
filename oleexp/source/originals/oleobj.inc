[
  odl,
  uuid(00000118-0000-0000-C000-000000000046)
]
interface IOleClientSite : stdole.IUnknown {

    HRESULT SaveObject();

    HRESULT GetMoniker(
        [in] OLEGETMONIKER dwAssign,
        [in] OLEWHICHMK dwWhichMoniker,
        [out, retval] IMoniker **ppIMoniker);

    HRESULT GetContainer(
        [out, retval] IOleContainer** ppContainer);

    HRESULT ShowObject();

    HRESULT OnShowWindow(
        [in] BOOL fShow);

    HRESULT RequestNewObjectLayout();
};

[
  odl,
  uuid(0000011A-0000-0000-C000-000000000046)
]
interface IParseDisplayName : stdole.IUnknown {

    HRESULT ParseDisplayName(
        [in] IBindCtx *pIBindCtx,
        [in] LPWSTR pszDisplayName,
        [out, defaultvalue(0)] long* pchEaten,
        [out, retval] IMoniker **ppIMoniker);

};

[
  odl,
  uuid(0000011B-0000-0000-C000-000000000046)
]
interface IOleContainer : IParseDisplayName {

    LONG EnumObjects(
        [in] OLECONTF grfFlags,
        [in, out] IEnumUnknown **ppenum);

    HRESULT LockContainer(
        [in] BOOL fLock);

};

[
    odl,
    uuid(00000112-0000-0000-C000-000000000046)
]
interface IOleObject : stdole.IUnknown {

    HRESULT SetClientSite(
        [in] IOleClientSite* ppClientSite);

    HRESULT GetClientSite(
        [out, retval] IOleClientSite** ppClientSite);

    HRESULT SetHostNames(
        [in] LPWSTR szContainerApp,
        [in] LPWSTR szContainerObj);

    HRESULT Close(
        [in] OLECLOSE dwSaveOption);

    HRESULT SetMoniker(
        [in] OLEWHICHMK dwWhichMoniker,
        [in] IMoniker *pmIMoniker);

    HRESULT GetMoniker(
        [in] OLEGETMONIKER dwAssign,
        [in] OLEWHICHMK dwWhichMoniker,
        [out, retval] IMoniker **ppmk);

    HRESULT InitFromData(
        [in] IDataObject *pDataObject,
        [in] BOOL fCreation,
        [in] long dwReserved);

    HRESULT GetClipboardData(
        [in] long dwReserved,
        [out, retval] IDataObject **ppDataObject);

    HRESULT DoVerb(
        [in] long iVerb,
        [in] MSG *lpmsg,
        [in] IOleClientSite *pActiveSite,
        [in] long lindex,
        [in] long hwndParent,
        [in] RECT* lprcPosRect);

    LONG EnumVerbs(
        [in, out] IEnumOLEVERB** ppEnumOleVerb);

    HRESULT Update();

    LONG IsUpToDate();

    HRESULT GetUserClassID(
        [out] UUID* pClsid);

    HRESULT GetUserType(
        [in] USERCLASSTYPE dwFormOfType,
        [out, retval] long *pszUserType);

    HRESULT SetExtent(
        [in] DVASPECT dwDrawAspect,
        [in, out] SIZE *pSIZE);

    HRESULT GetExtent(
        [in] DVASPECT dwDrawAspect,
        [in, out] SIZE* pSIZE);

    HRESULT Advise(
        [in] IAdviseSink *pIAdviseSink,
        [in, out] long* pdwConnection);

    HRESULT Unadvise(
        [in] long dwConnection);

    HRESULT EnumAdvise(
        [out, retval] IEnumSTATDATA** ppenumAdvise);

    HRESULT GetMiscStatus(
        [in] DVASPECT dwAspect,
        [out, retval] OLEMISC* pdwStatus);

    HRESULT SetColorScheme(
        [in] LOGPALETTE* pLogpal);

};

[
    odl,
    uuid(0000011e-0000-0000-C000-000000000046),
]
interface IOleCache : stdole.IUnknown {

    HRESULT Cache (
        [in] FORMATETC *pformatetc,
        [in] LONG advf,
        [out, retval] LONG *pdwConnection);

    HRESULT Uncache(
        [in] LONG dwConnection);

    HRESULT EnumCache(
        [out, retval] IEnumSTATDATA **ppenumSTATDATA);

    HRESULT InitCache(
        [in] IDataObject *pDataObject);

    HRESULT SetData(
        [in] FORMATETC *pformatetc,
        [in] STGMEDIUM *pmedium,
        [in] BOOL fRelease);
};

[
    odl,
    uuid(B196B288-BAB4-101A-B69C-00AA00341D07),
]
interface IOleControl : stdole.IUnknown {

    typedef enum CTRLINFO {
        CTRLINFO_EATS_RETURN = 1,
        CTRLINFO_EATS_ESCAPE = 2
    } CTRLINFO;

    typedef struct CONTROLINFO {
        LONG  cb;
        LONG hAccel;
        short cAccel;
        CTRLINFO dwFlags;
    } CONTROLINFO;

    HRESULT GetControlInfo(
        [out] CONTROLINFO *pCI);

    HRESULT OnMnemonic(
        [in] MSG * pMsg);

    HRESULT OnAmbientPropertyChange(
        [in] LONG dispID);

    HRESULT FreezeEvents(
        [in] BOOL bFreeze);
}

[
    odl,
    uuid(b722bccb-4e68-101b-a2bc-00aa00404770),
]
interface IOleCommandTarget : stdole.IUnknown {

    typedef enum OLECMDF {
        OLECMDF_SUPPORTED       = 0x00000001,
        OLECMDF_ENABLED         = 0x00000002,
        OLECMDF_LATCHED         = 0x00000004,
        OLECMDF_NINCHED         = 0x00000008,
        OLECMDF_INVISIBLE       = 0x00000010,
        OLECMDF_DEFHIDEONCTXTMENU = 0x00000020,
    } OLECMDF;

    typedef struct OLECMD {
        LONG cmdID;
        OLECMDF cmdf;
    } OLECMD;

    typedef enum OLECMDTEXTF {
        OLECMDTEXTF_NONE        = 0,
        OLECMDTEXTF_NAME        = 1,
        OLECMDTEXTF_STATUS      = 2,
    } OLECMDTEXTF;

    typedef struct OLECMDTEXT{
        OLECMDTEXTF cmdtextf;
        LONG cwActual;
        LONG cwBuf;     // size in wide chars of the buffer for text
        BYTE rgwz[512]; // Array into which callee writes the text
    } OLECMDTEXT;

    typedef enum OLECMDEXECOPT {
        OLECMDEXECOPT_DODEFAULT         = 0,
        OLECMDEXECOPT_PROMPTUSER        = 1,
        OLECMDEXECOPT_DONTPROMPTUSER    = 2,
        OLECMDEXECOPT_SHOWHELP          = 3
    } OLECMDEXECOPT;

    typedef enum OLECMDID {
        OLECMDID_OPEN                           = 1,
        OLECMDID_NEW                            = 2,
        OLECMDID_SAVE                           = 3,
        OLECMDID_SAVEAS                         = 4,
        OLECMDID_SAVECOPYAS                     = 5,
        OLECMDID_PRINT                          = 6,
        OLECMDID_PRINTPREVIEW                   = 7,
        OLECMDID_PAGESETUP                      = 8,
        OLECMDID_SPELL                          = 9,
        OLECMDID_PROPERTIES                     = 10,
        OLECMDID_CUT                            = 11,
        OLECMDID_COPY                           = 12,
        OLECMDID_PASTE                          = 13,
        OLECMDID_PASTESPECIAL                   = 14,
        OLECMDID_UNDO                           = 15,
        OLECMDID_REDO                           = 16,
        OLECMDID_SELECTALL                      = 17,
        OLECMDID_CLEARSELECTION                 = 18,
        OLECMDID_ZOOM                           = 19,
        OLECMDID_GETZOOMRANGE                   = 20,
        OLECMDID_UPDATECOMMANDS                 = 21,
        OLECMDID_REFRESH                        = 22,
        OLECMDID_STOP                           = 23,
        OLECMDID_HIDETOOLBARS                   = 24,
        OLECMDID_SETPROGRESSMAX                 = 25,
        OLECMDID_SETPROGRESSPOS                 = 26,
        OLECMDID_SETPROGRESSTEXT                = 27,
        OLECMDID_SETTITLE                       = 28,
        OLECMDID_SETDOWNLOADSTATE               = 29,
        OLECMDID_STOPDOWNLOAD                   = 30,
        OLECMDID_ONTOOLBARACTIVATED             = 31,
        OLECMDID_FIND                           = 32,
        OLECMDID_DELETE                         = 33,
        OLECMDID_HTTPEQUIV                      = 34,
        OLECMDID_HTTPEQUIV_DONE                 = 35,
        OLECMDID_ENABLE_INTERACTION             = 36,
        OLECMDID_ONUNLOAD                       = 37,
        OLECMDID_PROPERTYBAG2                   = 38,
        OLECMDID_PREREFRESH                     = 39,
        OLECMDID_SHOWSCRIPTERROR                = 40,
        OLECMDID_SHOWMESSAGE                    = 41,
        OLECMDID_SHOWFIND                       = 42,
        OLECMDID_SHOWPAGESETUP                  = 43,
        OLECMDID_SHOWPRINT                      = 44,
        OLECMDID_CLOSE                          = 45,
        OLECMDID_ALLOWUILESSSAVEAS              = 46,
        OLECMDID_DONTDOWNLOADCSS                = 47,
    } OLECMDID;

    typedef enum OLECMDERR {
        OLECMDERR_E_FIRST            = 0x80040100,
        OLECMDERR_E_NOTSUPPORTED     = 0x80040100,
        OLECMDERR_E_DISABLED         = 0x80040101,
        OLECMDERR_E_NOHELP           = 0x80040102,
        OLECMDERR_E_CANCELED         = 0x80040103,
        OLECMDERR_E_UNKNOWNGROUP     = 0x80040104,
    } OLECMDERR;

    HRESULT QueryStatus(
        [in] void *pguidCmdGroup,
        [in] LONG cCmds,
        [in, out] OLECMD *prgCmds,
        [in, out] OLECMDTEXT *pCmdText);


    HRESULT Exec(
        [in] void *pguidCmdGroup,
        [in] LONG nCmdID,
        [in, defaultvalue(0)] OLECMDEXECOPT nCmdexecopt,
        [in, defaultvalue(0)] VARIANT *pvaIn,
        [in, out, defaultvalue(0)] VARIANT *pvaOut);
}

[
    odl,
    uuid(6d5140c1-7436-11ce-8034-00aa006009fa),
	helpstring("IServiceProvider interface")
]
interface IServiceProvider : stdole.IUnknown {

    HRESULT QueryService(
        [in, out] UUID *guidService,
        [in, out] UUID *riid,
        [in, out] void *ppvObject);

}

[
    odl,
    uuid(B196B28B-BAB4-101A-B69C-00AA00341D07),
]
interface ISpecifyPropertyPages : stdole.IUnknown {

    typedef [helpstring("Counted Array of UUIDs")]
    struct CAUUID {
        [helpstring("Elements count")]
        LONG cElems;
        [helpstring("Pointer to an array of GUIDs")]
        LONG pElems;
    } CAUUID;

    HRESULT GetPages(
        [in, out] CAUUID *pPages);
}

[
    odl,
    uuid(0000011d-0000-0000-C000-000000000046)
]
interface IOleLink : stdole.IUnknown {

    typedef enum OLEUPDATE {
        OLEUPDATE_ALWAYS=1,
        OLEUPDATE_ONCALL=3
    } OLEUPDATE;

    // for IOleLink::BindToSource
    typedef enum OLELINKBIND {
        OLELINKBIND_EVENIFCLASSDIFF = 1,
    } OLELINKBIND;

    HRESULT SetUpdateOptions(
        [in] OLEUPDATE dwUpdateOpt);

    HRESULT GetUpdateOptions(
        [out, retval] OLEUPDATE *pdwUpdateOpt);

    HRESULT SetSourceMoniker(
        [in] IMoniker *pmk,
        [in] UUID *rclsid);

    HRESULT GetSourceMoniker(
        [out, retval] IMoniker **ppmk);

    HRESULT SetSourceDisplayName(
        [in] LPWSTR pszStatusText);

    HRESULT GetSourceDisplayName(
        [out, retval] LONG *ppszDisplayName);

    HRESULT BindToSource(
        [in] OLELINKBIND bindflags,
        [in] IBindCtx *pbc);

    LONG BindIfRunning(void);

    HRESULT GetBoundSource(
        [out, retval] stdole.IUnknown **ppunk);

    LONG UnbindSource(void);

    HRESULT Update(
        [in] IBindCtx *pbc);
}

/*
[
    odl,
    uuid()
]
interface IOleUILinkContainerA : stdole.IUnknown {

    LONG GetNextLink(
        [in] LONG dwLink);

    HRESULT SetLinkUpdateOptions(
        [in] LONG dwLink,
        [in] LONG dwUpdateOpt);

    HRESULT GetLinkUpdateOptions(
        [in] LONG dwLink,
        [out, retval] LONG *lpdwUpdateOpt);

    HRESULT SetLinkSource(
        [in] LONG dwLink,
        [in] LPSTR lpszDisplayName,
        [in] LONG lenFileName,
        [in, out] LONG *pchEaten,
        [in] BOOL fValidateSource);

    HRESULT GetLinkSource(
        [in] LONG dwLink,
        [in, out] LPSTR *lplpszDisplayName,
        [in, out] LONG *lplenFileName,
        [in, out] LPSTR *lplpszFullLinkType,
        [in, out] LPSTR *lplpszShortLinkType,
        [in, out] BOOL *lpfSourceAvailable,
        [in, out] BOOL *lpfIsSelected);

    HRESULT OpenLinkSource(
        [in] LONG dwLink);

    HRESULT UpdateLink(
        [in] LONG dwLink,
        [in] BOOL fErrorMessage,
        [in] BOOL fReserved);

    HRESULT CancelLink(
        [in] LONG dwLink);
}

*/