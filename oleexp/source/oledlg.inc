[
    dllname("OLEDLG.DLL")
]
module oledlg {

    const LPSTR OLESTDDELIM = "\\";

    // Strings for registered messages
    const LPSTR SZOLEUI_MSG_HELP = "OLEUI_MSG_HELP";
    const LPSTR SZOLEUI_MSG_ENDDIALOG = "OLEUI_MSG_ENDDIALOG";
    const LPSTR SZOLEUI_MSG_BROWSE = "OLEUI_MSG_BROWSE";
    const LPSTR SZOLEUI_MSG_CHANGEICON = "OLEUI_MSG_CHANGEICON";
    const LPSTR SZOLEUI_MSG_CLOSEBUSYDIALOG = "OLEUI_MSG_CLOSEBUSYDIALOG";
    const LPSTR SZOLEUI_MSG_CONVERT = "OLEUI_MSG_CONVERT";
    const LPSTR SZOLEUI_MSG_CHANGESOURCE = "OLEUI_MSG_CHANGESOURCE";
    const LPSTR SZOLEUI_MSG_ADDCONTROL = "OLEUI_MSG_ADDCONTROL";
    const LPSTR SZOLEUI_MSG_BROWSE_OFN = "OLEUI_MSG_BROWSE_OFN";

    // Identifiers for SZOLEUI_MSG_BROWSE_OFN (in wParam)
    const LONG ID_BROWSE_CHANGEICON = 1;
    const LONG ID_BROWSE_INSERTFILE = 2;
    const LONG ID_BROWSE_ADDCONTROL = 3;
    const LONG ID_BROWSE_CHANGESOURCE = 4;

    [entry("OleUIAddVerbMenuW")]
    BOOL OleUIAddVerbMenuW(
        [in] IOleObject *lpOleObj,
        [in] LPWSTR lpszShortType,
        [in] LONG hMenu,
        [in] LONG uPos,
        [in] LONG uIDVerbMin,
        [in] LONG uIDVerbMax,
        [in] BOOL bAddConvert,
        [in] LONG idConvert,
        [in, out] LONG *lphMenu);

    [entry("OleUIAddVerbMenuA")]
    BOOL OleUIAddVerbMenu(
        [in] IOleObject *lpOleObj,
        [in] LPSTR lpszShortType,
        [in] LONG hMenu,
        [in] LONG uPos,
        [in] LONG uIDVerbMin,
        [in] LONG uIDVerbMax,
        [in] BOOL bAddConvert,
        [in] LONG idConvert,
        [in, out] LONG *lphMenu);

    [entry("OleUIInsertObjectW")]
    LONG OleUIInsertObjectW([in] void *any);

    [entry("OleUIInsertObjectA")]
    LONG OleUIInsertObject([in] void *any);

    [entry("OleUIPasteSpecialW")]
    LONG OleUIPasteSpecialW([in] void *any);

    [entry("OleUIPasteSpecialA")]
    LONG OleUIPasteSpecial([in] void *any);

    [entry("OleUIEditLinksW")]
    LONG OleUIEditLinksW([in] void *any);

    [entry("OleUIEditLinksA")]
    LONG OleUIEditLinks([in] void *any);

    [entry("OleUIChangeIconW")]
    LONG OleUIChangeIconW([in] void *any);

    [entry("OleUIChangeIconA")]
    LONG OleUIChangeIcon([in] void *any);

    [entry("OleUIConvertW")]
    LONG OleUIConvertW([in] void *any);

    [entry("OleUIConvertA")]
    LONG OleUIConvert([in] void *any);

    [entry("OleUIBusyW")]
    LONG OleUIBusyW([in] void *any);

    [entry("OleUIBusyA")]
    LONG OleUIBusy([in] void *any);

    [entry("OleUIChangeSourceW")]
    LONG OleUIChangeSourceW([in] void *any);

    [entry("OleUIChangeSourceA")]
    LONG OleUIChangeSource([in] void *any);

    [entry("OleUIUpdateLinksW")]
    BOOL OleUIUpdateLinksW(
        [in] LONG lpOleUILinkCntr,
        [in] LONG hwndParent,
        [in] LPWSTR lpszTitle,
        [in] LONG cLinks);

    [entry("OleUIUpdateLinksA")]
    BOOL OleUIUpdateLinks(
        [in] LONG lpOleUILinkCntr,
        [in] LONG hwndParent,
        [in] LPSTR lpszTitle,
        [in] LONG cLinks);

}

typedef enum OLEUI_ERR {

    // Standard success/error definitions
    OLEUI_FALSE                     = 0,
    OLEUI_SUCCESS                   = 1,     // No error, same as OLEUI_OK
    OLEUI_OK                        = 1,     // OK button pressed
    OLEUI_CANCEL                    = 2,     // Cancel button pressed

    OLEUI_ERR_STANDARDMIN           = 100,
    OLEUI_ERR_OLEMEMALLOC           = 100,
    OLEUI_ERR_STRUCTURENULL         = 101,   // Standard field validation
    OLEUI_ERR_STRUCTUREINVALID      = 102,
    OLEUI_ERR_CBSTRUCTINCORRECT     = 103,
    OLEUI_ERR_HWNDOWNERINVALID      = 104,
    OLEUI_ERR_LPSZCAPTIONINVALID    = 105,
    OLEUI_ERR_LPFNHOOKINVALID       = 106,
    OLEUI_ERR_HINSTANCEINVALID      = 107,
    OLEUI_ERR_LPSZTEMPLATEINVALID   = 108,
    OLEUI_ERR_HRESOURCEINVALID      = 109,
    OLEUI_ERR_FINDTEMPLATEFAILURE   = 110,   // Initialization errors
    OLEUI_ERR_LOADTEMPLATEFAILURE   = 111,
    OLEUI_ERR_DIALOGFAILURE         = 112,
    OLEUI_ERR_LOCALMEMALLOC         = 113,
    OLEUI_ERR_GLOBALMEMALLOC        = 114,
    OLEUI_ERR_LOADSTRING            = 115,
    OLEUI_ERR_STANDARDMAX           = 116,  // Start here for specific errors.

    // Insert Object specific error codes
    OLEUI_IOERR_LPSZFILEINVALID         = 116,
    OLEUI_IOERR_LPSZLABELINVALID,
    OLEUI_IOERR_HICONINVALID,
    OLEUI_IOERR_LPFORMATETCINVALID,
    OLEUI_IOERR_PPVOBJINVALID,
    OLEUI_IOERR_LPIOLECLIENTSITEINVALID,
    OLEUI_IOERR_LPISTORAGEINVALID,
    OLEUI_IOERR_SCODEHASERROR,
    OLEUI_IOERR_LPCLSIDEXCLUDEINVALID,
    OLEUI_IOERR_CCHFILEINVALID,

    // Paste Special specific error codes
    OLEUI_IOERR_SRCDATAOBJECTINVALID = 116,
    OLEUI_IOERR_ARRPASTEENTRIESINVALID,
    OLEUI_IOERR_ARRLINKTYPESINVALID,
    OLEUI_PSERR_CLIPBOARDCHANGED,
    OLEUI_PSERR_GETCLIPBOARDFAILED,

    // Edit Links specific error codes
    OLEUI_ELERR_LINKCNTRNULL = 116,
    OLEUI_ELERR_LINKCNTRINVALID,

    // Change Icon specific error codes
    OLEUI_CIERR_MUSTHAVECLSID = 116,
    OLEUI_CIERR_MUSTHAVECURRENTMETAFILE,
    OLEUI_CIERR_SZICONEXEINVALID,

    // Convert specific error codes
    OLEUI_CTERR_CLASSIDINVALID = 116,
    OLEUI_CTERR_DVASPECTINVALID,
    OLEUI_CTERR_CBFORMATINVALID,
    OLEUI_CTERR_HMETAPICTINVALID,
    OLEUI_CTERR_STRINGINVALID,

    // Busy specific error/return codes
    OLEUI_BZERR_HTASKINVALID = 116,
    OLEUI_BZ_SWITCHTOSELECTED,
    OLEUI_BZ_RETRYSELECTED,
    OLEUI_BZ_CALLUNBLOCKED,

    // Change Source Dialog errors
    OLEUI_CSERR_LINKCNTRNULL = 116,
    OLEUI_CSERR_LINKCNTRINVALID,
    OLEUI_CSERR_FROMNOTNULL,
    OLEUI_CSERR_TONOTNULL,
    OLEUI_CSERR_SOURCENULL,
    OLEUI_CSERR_SOURCEINVALID,
    OLEUI_CSERR_SOURCEPARSERROR,
    OLEUI_CSERR_SOURCEPARSEERROR,

} OLEUI_ERR;

typedef struct OLEUIHDR {
    LONG cbStruct;       // Structure Size
    LONG dwFlags;        // IN-OUT:  Flags
    LONG hWndOwner;      // Owning window
    LONG lpszCaption;    // Dialog caption bar contents
    LONG lpfnHook;       // Hook callback
    LONG lCustData;      // Custom data to pass to hook
    LONG hInstance;      // Instance for customized template name
    LONG lpszTemplate;   // Customized template name
    LONG hResource;      // Customized template handle
} OLEUIHDR;

typedef struct OLEUIINSERTOBJECT {
    OLEUIHDR Hdr;
 
    // Specifics for OLEUIINSERTOBJECT.
    UUID clsid;          // OUT: Return space for class ID
    LONG lpszFile;       // IN-OUT: Filename for inserts or links
    LONG cchFile;        // IN: Size of lpszFile buffer: MAX_PATH
    LONG cClsidExclude;  // IN: CLSIDs in lpClsidExclude
    LONG lpClsidExclude; // IN: List of CLSIDs to exclude from listing.

    // Specific to create objects if flags say so
    UUID iid;             // IN: Requested interface on creation.
    OLERENDER oleRender;  // IN: Rendering option
    LONG lpFormatEtc;     // IN: Desired format
    IOleClientSite *lpIOleClientSite;   // IN: Site to be use for the object.
    IStorage *lpIStorage; // IN: Storage used for the object
    stdole.IUnknown*ppvObj;     // OUT: Where the object is returned.
    LONG sc;              // OUT: Result of creation calls.
    LONG hMetaPict;       // OUT: metafile aspect (METAFILEPICT)
} OLEUIINSERTOBJECT;

// Insert Object flags
typedef enum INSERTOBJECTFLAGS {
    IOF_SHOWHELP                    = 0x00000001,
    IOF_SELECTCREATENEW             = 0x00000002,
    IOF_SELECTCREATEFROMFILE        = 0x00000004,
    IOF_CHECKLINK                   = 0x00000008,
    IOF_CHECKDISPLAYASICON          = 0x00000010,
    IOF_CREATENEWOBJECT             = 0x00000020,
    IOF_CREATEFILEOBJECT            = 0x00000040,
    IOF_CREATELINKOBJECT            = 0x00000080,
    IOF_DISABLELINK                 = 0x00000100,
    IOF_VERIFYSERVERSEXIST          = 0x00000200,
    IOF_DISABLEDISPLAYASICON        = 0x00000400,
    IOF_HIDECHANGEICON              = 0x00000800,
    IOF_SHOWINSERTCONTROL           = 0x00001000,
    IOF_SELECTCREATECONTROL         = 0x00002000,
} INSERTOBJECTFLAGS;

typedef enum OLEUIPASTEFLAG {
   OLEUIPASTE_PASTEONLY     = 0,
   OLEUIPASTE_LINKTYPE1     = 1,
   OLEUIPASTE_LINKTYPE2     = 2,
   OLEUIPASTE_LINKTYPE3     = 4,
   OLEUIPASTE_LINKTYPE4     = 8,
   OLEUIPASTE_LINKTYPE5     = 16,
   OLEUIPASTE_LINKTYPE6     = 32,
   OLEUIPASTE_LINKTYPE7     = 64,
   OLEUIPASTE_LINKTYPE8     = 128,
   OLEUIPASTE_PASTE         = 512,
   OLEUIPASTE_LINKANYTYPE   = 1024,
   OLEUIPASTE_ENABLEICON    = 2048,     // enable display as icon
} OLEUIPASTEFLAG;

typedef struct OLEUIPASTEENTRY {
   FORMATETC fmtetc;      // Format that is acceptable.
   LONG lpstrFormatName;  // String that represents the format
                          // to the user. %s is replaced by the
                          // full user type name of the object.
   LONG lpstrResultText;  // String to customize the result text
                          // of the dialog when the user
                          // selects the format correspoding to
                          // this entry. Any %s in this string
                          // is replaced by the the application
                          // name or FullUserTypeName of the
                          // object on the clipboard.
   OLEUIPASTEFLAG dwFlags;// Values from OLEUIPASTEFLAG enum
   LONG dwScratchSpace;   // Scratch space used internally.
} OLEUIPASTEENTRY;

typedef struct OLEUIPASTESPECIAL {
    OLEUIHDR Hdr;

    // Specifics for OLEUIPASTESPECIAL.
    IDataObject *lpSrcDataObj; // IN-OUT: Source IDataObject* on the clipboard
                               // If lpSrcDataObj is NULL when OleUIPasteSpecial is called, then
                               // OleUIPasteSpecial will attempt to retrieve a pointer to an
                               // IDataObject from the clipboard.  If OleUIPasteSpecial succeeds
                               // then it is the caller's responsibility to free the IDataObject
                               // returned in lpSrcDataObj.
    LONG arrPasteEntries;      // IN: Array of acceptable formats
    LONG cPasteEntries;        // IN: No. of OLEUIPASTEENTRY array entries
    LONG arrLinkTypes;         // IN: List of acceptable link types
    LONG cLinkTypes;           // IN: Number of link types
    LONG cClsidExclude;        // IN: Number of CLSIDs in lpClsidExclude
    LONG lpClsidExclude;       // IN: List of CLSIDs to exclude from list.
    LONG nSelectedIndex;       // OUT: Index that the user selected
    BOOL fLink;                // OUT: Indicates if Paste or PasteLink
    LONG hMetaPict;            // OUT: Handle to Metafile containing icon
    SIZE sizel;                // OUT: size of object/link in its source
                               //  may be 0,0 if different display
                               //  aspect is chosen.
} OLEUIPASTESPECIAL;


// Paste Special specific flags
typedef enum PASTESPECIALFLAGS {
    PSF_SHOWHELP                    = 0x00000001,
    PSF_SELECTPASTE                 = 0x00000002,
    PSF_SELECTPASTELINK             = 0x00000004,
    // NOTE: PSF_CHECKDISPLAYASICON is strictly an output flag.
    //       It is ignored if set when calling OleUIPasteSpecial.
    PSF_CHECKDISPLAYASICON          = 0x00000008,
    PSF_DISABLEDISPLAYASICON        = 0x00000010,
    PSF_HIDECHANGEICON              = 0x00000020,
    PSF_STAYONCLIPBOARDCHANGE       = 0x00000040,
    PSF_NOREFRESHDATAOBJECT         = 0x00000080,
} PASTESPECIALFLAGS;

typedef struct OLEUIEDITLINKS {
    OLEUIHDR Hdr;

    // Specifics for OLEUIEDITLINKS.
    LONG lpOleUILinkContainer;  // IN: Interface to manipulate
} OLEUIEDITLINKS;

// Edit Links flags
typedef enum EDITLINKSFLAGS {
    ELF_SHOWHELP = 0x00000001,
    ELF_DISABLEUPDATENOW = 0x00000002,
    ELF_DISABLEOPENSOURCE = 0x00000004,
    ELF_DISABLECHANGESOURCE = 0x00000008,
    ELF_DISABLECANCELLINK = 0x00000010,
} EDITLINKSFLAGS;

typedef struct OLEUICHANGEICON {
    OLEUIHDR Hdr;

    // Specifics for OLEUICHANGEICON.
    LONG hMetaPict;  // IN-OUT: Current and final image.
                     //  Source of the icon is embedded in
                     //  the metafile itself.
    UUID clsid;      // IN: class used to get Default icon
    BYTE szIconExe[260]; // IN: explicit icon source path
    LONG cchIconExe;     // IN: number of characters in szIconExe
} OLEUICHANGEICON;

// Change Icon flags
typedef enum CHANGEICONFLAGS {
    CIF_SHOWHELP                    = 0x00000001,
    CIF_SELECTCURRENT               = 0x00000002,
    CIF_SELECTDEFAULT               = 0x00000004,
    CIF_SELECTFROMFILE              = 0x00000008,
    CIF_USEICONEXE                  = 0x00000010,
} CHANGEICONFLAGS;

typedef struct OLEUICONVERT {
    OLEUIHDR Hdr;

    // Specifics for OLEUICONVERT.
    UUID clsid;                  // IN: Class ID sent in to dialog: IN only
    UUID clsidConvertDefault;    // IN: use as convert default: IN only
    UUID clsidActivateDefault;   // IN: use as activate default: IN only

    UUID clsidNew;        // OUT: Selected Class ID
    DVASPECT dvAspect;    // IN-OUT: either DVASPECT_CONTENT or
                          //  DVASPECT_ICON
    short wFormat;        // IN" Original data format
    BOOL fIsLinkedObject; // IN: true if object is linked
    LONG hMetaPict;       // IN-OUT: metafile icon image
    LONG lpszUserType;    // IN-OUT: user type name of original class.
                          //      We'll do lookup if NULL.
                          //      This gets freed on exit.
    BOOL fObjectsIconChanged; // OUT: TRUE == ChangeIcon was called
    LPWSTR lpszDefLabel;      //IN-OUT: default label to use for icon.
                              //  if NULL, the short user type name
                              //  will be used. if the object is a
                              //  link, the caller should pass the
                              //  DisplayName of the link source
                              //  This gets freed on exit.
    LONG cClsidExclude;  //IN: No. of CLSIDs in lpClsidExclude
    LONG lpClsidExclude; //IN: List of CLSIDs to exclude from list
} OLEUICONVERT;

// Convert Dialog flags
typedef enum CONVERTFLAGS {
    CF_SHOWHELPBUTTON               = 0x00000001,
    CF_SETCONVERTDEFAULT            = 0x00000002,
    CF_SETACTIVATEDEFAULT           = 0x00000004,
    CF_SELECTCONVERTTO              = 0x00000008,
    CF_SELECTACTIVATEAS             = 0x00000010,
    CF_DISABLEDISPLAYASICON         = 0x00000020,
    CF_DISABLEACTIVATEAS            = 0x00000040,
    CF_HIDECHANGEICON               = 0x00000080,
    CF_CONVERTONLY                  = 0x00000100,
} CONVERTFLAGS;

typedef struct OLEUIBUSY {
    OLEUIHDR Hdr;

    // Specifics for OLEUIBUSY.
    LONG hTask;          // IN: HTask which is blocking
    LONG lphWndDialog;   // OUT: Dialog's HWND is placed here
} OLEUIBUSY;

// Flags for the Busy dialog
typedef enum BUSYFLAGS {
    BZ_DISABLECANCELBUTTON = 0x00000001,
    BZ_DISABLESWITCHTOBUTTON = 0x00000002,
    BZ_DISABLERETRYBUTTON = 0x00000004,
} BUSYFLAGS;

// Data to and from the ChangeSource dialog hook
typedef struct OLEUICHANGESOURCE {
    OLEUIHDR Hdr;

    // INTERNAL ONLY: do not modify these members
    LONG lpOFN;          // pointer OPENFILENAME struct
    LONG dwReserved1[4]; // (reserved for future use)

    // Specifics for OLEUICHANGESOURCE.
    LONG lpOleUILinkContainer;  // IN: used to validate link sources
    LONG dwLink;                // IN: magic# for lpOleUILinkContainer
    LONG lpszDisplayName;       // IN-OUT: complete source display name
    LONG nFileLength;           // IN-OUT: file moniker part of lpszDisplayName
    LONG lpszFrom;              // OUT: prefix of source changed from
    LONG lpszTo;                // OUT: prefix of source changed to
} OLEUICHANGESOURCE;

// Change Source Dialog flags
typedef enum CHANGESOURCEFLAGS {
    CSF_SHOWHELP = 0x00000001,     // IN: enable/show help button
    CSF_VALIDSOURCE = 0x00000002,  // OUT: link was validated
    CSF_ONLYGETSOURCE = 0x00000004,// IN: disables validation of source
    CSF_EXPLORER = 0x00000008,     // IN: use new OFN_EXPLORER custom template behavior
} CHANGESOURCEFLAGS;

