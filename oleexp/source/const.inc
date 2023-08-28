[
    dllname("dummy")  
]
module constants {

    // Clipboard Formats
    // Mime types
	// Those have been moved to exp_cf.odl

    const LONG DVASPECT_SHORTNAME = 2; // use for CF_HDROP to get short name version of file paths
    const LONG DVASPECT_COPY      = 3; // use to indicate format is a "Copy" of the data (FILECONTENTS, FILEDESCRIPTOR, etc)
    const LONG DVASPECT_LINK      = 4; // use to indicate format is a "Shortcut" to the data (FILECONTENTS, FILEDESCRIPTOR, etc)

    // Command group IDs
    const LPSTR CGIDSTR_EditStateCommands     = "{3050f4b6-98b5-11cf-bb82-00aa00bdce0b}";
    const LPSTR CGIDSTR_DocHostCommandHandler = "{f38bc242-b950-11d1-8918-00c04fc2c836}";
    const LPSTR CGIDSTR_DeskBand              = "{EB0FE172-1A3A-11D0-89B3-00A0C90A90AC}";
    const LPSTR CGIDSTR_Explorer              = "{000214D0-0000-0000-C000-000000000046}";
    const LPSTR CGIDSTR_ShellDocView          = "{000214D1-0000-0000-C000-000000000046}";
    const LPSTR CGIDSTR_ShellServiceObject    = "{000214D2-0000-0000-C000-000000000046}";
    const LPSTR CGIDSTR_ExplorerBarDoc        = "{000214D3-0000-0000-C000-000000000046}";
    const LPSTR CGIDSTR_IDeskBand             = "{EB0FE172-1A3A-11D0-89B3-00A0C90A90AC}";
    const LPSTR CGIDSTR_ShortCut              = "{93a68750-951a-11d1-946f-000000000000}";
    const LPSTR CGIDSTR_HTML                  = "{DE4BA900-59CA-11CF-9592-444553540000}";

    // Interface IDs
    const LPSTR IIDSTR_IUnknown                 = "{00000000-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IDispatch                = "{00020400-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IConnectionPoint         = "{B196B286-BAB4-101A-B69C-00AA00341D07}";
    const LPSTR IIDSTR_IConnectionPointContainer= "{B196B284-BAB4-101A-B69C-00AA00341D07}";
    const LPSTR IIDSTR_IOleClientSite           = "{00000118-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IDocHostShowUI           = "{C4D244B0-D43E-11CF-893B-00AA00BDCE1A}";
    const LPSTR IIDSTR_IDocHostUIHandler        = "{BD3F23C0-D43E-11CF-893B-00AA00BDCE1A}";
    const LPSTR IIDSTR_IServiceProvider         = "{6D5140C1-7436-11CE-8034-00AA006009FA}";
    const LPSTR IIDSTR_IWebBrowserApp           = "{0002DF05-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IInternetSecurityManager = "{79eac9ee-baf9-11ce-8c82-00aa004ba90b}";
    const LPSTR IIDSTR_StandardFont             = "{7EBDAAE0-8120-11CF-899F-00AA00688B10}";
    const LPSTR IIDSTR_StandardColor            = "{7EBDAAE1-8120-11CF-899F-00AA00688B10}";
    const LPSTR IIDSTR_StandardPicture          = "{7EBDAAE2-8120-11CF-899F-00AA00688B10}";    
    const LPSTR IIDSTR_IPicture                 = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}";
    const LPSTR IIDSTR_IPictureDisp             = "{7BF80981-BF32-101A-8BBB-00AA00300CAB}";
    const LPSTR IIDSTR_IFont                    = "{BEF6E002-A874-101A-8BBA-00AA00300CAB}";
    const LPSTR IIDSTR_IFontDisp                = "{BEF6E003-A874-101A-8BBA-00AA00300CAB}";
    const LPSTR IIDSTR_IQueryAssociations       = "{c46ca590-3c3f-11d2-bee6-0000f805ca57}";
    const LPSTR IIDSTR_IContextMenu             = "{000214E4-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IContextMenu2            = "{000214F4-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IQueryInfo               = "{000214E7-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IExtractIconA            = "{000214EB-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IExtractIconW            = "{000214FA-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IShellFolder             = "{000214E6-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IShellView               = "{000214E3-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IShellView2              = "{88E39E80-3578-11CF-AE69-08002B2E1262}";
    const LPSTR IIDSTR_IRedbookDiscMaster       = "{E3BC42CD-4E5C-11D3-9144-00104BA11C5E)}";
    const LPSTR IIDSTR_IJolietDiscMaster        = "{E3BC42CE-4E5C-11D3-9144-00104BA11C5E}";

    // CLSIDs
    const LPSTR CLSIDSTR_InternetSecurityManager = "{7B8A2D94-0AC9-11D1-896C-00C04FB6BFC4}";
    const LPSTR CLSIDSTR_InternetZoneManager     = "{7B8A2D95-0AC9-11D1-896C-00C04FB6BFC4}";
    const LPSTR CLSIDSTR_QueryAssociations       = "{c46ca590-3c3f-11d2-bee6-0000f805ca57}";
    const LPSTR CLSIDSTR_ShellDesktop            = "{00021400-0000-0000-C000-000000000046}";
    const LPSTR CLSIDSTR_ShellLink               = "{00021401-0000-0000-C000-000000000046}";
    const LPSTR CLSIDSTR_NetworkPlaces           = "{208D2C60-3AEA-1069-A2D7-08002B30309D}";
    const LPSTR CLSIDSTR_NetworkDomain           = "{46e06680-4bf0-11d1-83ee-00a0c90dc849}";
    const LPSTR CLSIDSTR_NetworkServer           = "{c0542a90-4bf0-11d1-83ee-00a0c90dc849}";
    const LPSTR CLSIDSTR_NetworkShare            = "{54a754c0-4bf1-11d1-83ee-00a0c90dc849}";
    const LPSTR CLSIDSTR_MyComputer              = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}";
    const LPSTR CLSIDSTR_Internet                = "{871C5380-42A0-1069-A2EA-08002B30309D}";
    const LPSTR CLSIDSTR_ShellFSFolder           = "{F3364BA0-65B9-11CE-A9BA-00AA004AE837}";
    const LPSTR CLSIDSTR_RecycleBin              = "{645FF040-5081-101B-9F08-00AA002F954E}";
    const LPSTR CLSIDSTR_ControlPanel            = "{21EC2020-3AEA-1069-A2DD-08002B30309D}";
    const LPSTR CLSIDSTR_Printers                = "{2227A280-3AEA-1069-A2DE-08002B30309D}";
    const LPSTR CLSIDSTR_MyDocuments             = "{450D8FBA-AD25-11D0-98A8-0800361B1103}";

    // Category IDs
    [helpstring("Embeddable Objects")]
    const LPSTR CATIDSTR_Insertable          = "{40FC6ED3-2438-11CF-A3DB-080036F12502}";
    [helpstring("Controls")]
    const LPSTR CATIDSTR_Control             = "{40fc6ed4-2438-11cf-a3db-080036f12502}";
    [helpstring("Automation Objects")]
    const LPSTR CATIDSTR_Programmable        = "{40FC6ED5-2438-11CF-A3DB-080036F12502}";
    [helpstring("Document Objects")]
    const LPSTR CATIDSTR_DocObject           = "{40fc6ed8-2438-11cf-a3db-080036f12502}";
    [helpstring("Printable Objects")]
    const LPSTR CATIDSTR_Printable           = "{40fc6ed9-2438-11cf-a3db-080036f12502}";
    [helpstring("Active Scripting Engine With Authoring")]
    const LPSTR CATIDSTR_ActiveScriptAuthor  = "{0AEE2A92-BCBB-11D0-8C72-00C04FC2B085}";
    [helpstring("Active Scripting Engine With Parsing")]
    const LPSTR CATIDSTR_ActiveScriptParse   = "{F0B7A1A2-9847-11cf-8F20-00805F2CD064}";
    [helpstring("Active Scripting Engine")]
    const LPSTR CATIDSTR_ActiveScript        = "{F0B7A1A1-9847-11cf-8F20-00805F2CD064}";
    [helpstring("Active Scripting Engine With Encoding")]
    const LPSTR CATIDSTR_ActiveScriptEncode  = "{F0B7A1A3-9847-11cf-8F20-00805F2CD064}";
    [helpstring("Controls that are safely scriptable")]
    const LPSTR CATIDSTR_SafeForScripting    = "{7DD95801-9882-11CF-9FA9-00AA006C42C4}";
    [helpstring("Controls safely initializable from persistent data")]
    const LPSTR CATIDSTR_SafeForInitializing = "{7DD95802-9882-11CF-9FA9-00AA006C42C4}";
    const LPSTR CATIDSTR_BrowsableShellExt   = "{00021490-0000-0000-C000-000000000046}";
    const LPSTR CATIDSTR_BrowseInPlace       = "{00021491-0000-0000-C000-000000000046}";
    [helpstring("Desktop Bands")]
    const LPSTR CATIDSTR_DeskBand            = "{00021492-0000-0000-C000-000000000046}";
    [helpstring("Internet Explorer Browser Band")]
    const LPSTR CATIDSTR_InfoBand            = "{00021493-0000-0000-C000-000000000046}";
    [helpstring("Internet Explorer Communication Band")]
    const LPSTR CATIDSTR_CommBand            = "{00021494-0000-0000-C000-000000000046}";

/*
		const LPSTR CATIDSTR_RequiresDataPathHost = "";
		const LPSTR CATIDSTR_PersistsToMoniker = "";
		const LPSTR CATIDSTR_PersistsToStorage = "";
		const LPSTR CATIDSTR_PersistsToStreamInit = "";
		const LPSTR CATIDSTR_PersistsToStream = "";
		const LPSTR CATIDSTR_PersistsToMemory = "";
		const LPSTR CATIDSTR_PersistsToFile = "";
		const LPSTR CATIDSTR_PersistsToPropertyBag = "";
		const LPSTR CATIDSTR_InternetAware = "";
		const LPSTR CATIDSTR_DesignTimeUIActivatableControl = "";
		const LPSTR CATIDSTR_IsShortcut = "{}";
		const LPSTR CATIDSTR_NeverShowExt = "{}";
*/

    // Property Set names
    const LPSTR DocumentSummaryInformation = "\005DocumentSummaryInformation";
    const LPSTR SummaryInformation         = "\005SummaryInformation";

    // Format IDs
    const LPSTR FMTIDSTR_SummaryInformation    = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}";
    const LPSTR FMTIDSTR_DocSummaryInformation = "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}";
    const LPSTR FMTIDSTR_UserProperties        = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
    const LPSTR FMTIDSTR_Intshcut              = "{000214A0-0000-0000-C000-000000000046}";
    const LPSTR FMTIDSTR_InternetSite          = "{000214A1-0000-0000-C000-000000000046}";

    // Service IDs
    const LPSTR SIDSTR_SEditCommandTarget = "{3050f4b5-98b5-11cf-bb82-00aa00bdce0b}";
    const LPSTR SIDSTR_SHTMLEditHost      = "{3050f6a0-98b5-11cf-bb82-00aa00bdce0b}";
    const LPSTR SIDSTR_SHTMLEditServices  = "{3050f7f9-98b5-11cf-bb82-00aa00bdce0b}";
    const LPSTR SIDSTR_SWebBrowserApp     = "{0002DF05-0000-0000-C000-000000000046}";
    const LPSTR SIDSTR_SInternetExplorer  = "{0002DF05-0000-0000-C000-000000000046}";
    const LPSTR SIDSTR_SFolderView        = "{cde725b0-ccc9-4519-917e-325d72fab4ce}";
    const LPSTR SIDSTR_SProfferService    = "{cb728b20-f786-11ce-92ad-00aa00a74cd0}";
    const LPSTR SIDSTR_WebWizardHost      = "{0e6b3f66-98d1-48c0-a222-fbde74e2fbc5}";
    const LPSTR SIDSTR_PublishingWizard   = "{6b33163c-76a5-4b6c-bf21-45de9cd503a1}";
    const LPSTR SIDSTR_SProgressUI        = "{F8383852-FCD3-11d1-A6B9-006097DF5BD4}";
    const LPSTR SIDSTR_SHlinkFrame        = "{79eac9c5-baf9-11ce-8c82-00aa004ba90b}";
    const LPSTR SIDSTR_ShellFolderViewCB  = "{2047E320-F2A9-11CE-AE65-08002B2E1262}";
    const LPSTR SIDSTR_LinkSite           = "{000214F9-0000-0000-C000-000000000046}";
    const LPSTR SIDSTR_SShellBrowser      = "{000214E2-0000-0000-C000-000000000046}";
    const LPSTR SIDSTR_SShellDesktop      = "{00021400-0000-0000-C000-000000000046}";
    const LPSTR SIDSTR_SUrlHistory        = "{3C374A40-BAE4-11CF-BF7D-00AA006946EE}";
    const LPSTR SIDSTR_STopLevelBrowser   = "{4C96BE40-915C-11CF-99D3-00AA004AE837}";

	const long INPLACE_S_TRUNCATED = 0x401A0;
	const long OLEOBJ_S_INVALIDVERB = 0x40180;
	const long OLEOBJ_S_CANNOT_DOVERB_NOW = 0x40181;
	const long OLEOBJ_S_INVALIDHWND = 0x40182;

    const long OLE_E_OLEVERB = 0x80040000;
    const long OLE_E_ADVF = 0x80040001;
    const long OLE_E_ENUM_NOMORE = 0x80040002;
    const long OLE_E_ADVISENOTSUPPORTED = 0x80040003;
    const long OLE_E_NOCONNECTION = 0x80040004;
    const long OLE_E_NOTRUNNING = 0x80040005;
    const long OLE_E_NOCACHE = 0x80040006;
    const long OLE_E_BLANK = 0x80040007;
    const long OLE_E_CLASSDIFF = 0x80040008;
    const long OLE_E_CANT_GETMONIKER = 0x80040009;
    const long OLE_E_CANT_BINDTOSOURCE = 0x8004000A;
    const long OLE_E_STATIC = 0x8004000B;
    const long OLE_E_PROMPTSAVECANCELLED = 0x8004000C;
    const long OLE_E_INVALIDRECT = 0x8004000D;
    const long OLE_E_WRONGCOMPOBJ = 0x8004000E;
    const long OLE_E_INVALIDHWND = 0x8004000F;
    const long OLE_E_NOT_INPLACEACTIVE = 0x80040010;
    const long OLE_E_CANTCONVERT = 0x80040011;
    const long OLE_E_NOSTORAGE = 0x80040012;
    const long OLE_S_USEREG = 0x40000;
    const long OLE_S_STATIC = 0x40001;
    const long OLE_S_MAC_CLIPFORMAT = 0x40002;

	const long REGDB_E_READREGDB = 0x80040150;
	const long REGDB_E_WRITEREGDB = 0x80040151;
	const long REGDB_E_KEYMISSING = 0x80040152;
	const long REGDB_E_INVALIDVALUE = 0x80040153;
	const long REGDB_E_CLASSNOTREG = 0x80040154;
	const long REGDB_E_IIDNOTREG = 0x80040155;

	const long STG_E_INVALIDFUNCTION = 0x80030001;
	const long STG_E_FILENOTFOUND = 0x80030002;
	const long STG_E_PATHNOTFOUND = 0x80030003;
	const long STG_E_TOOMANYOPENFILES = 0x80030004;
	const long STG_E_ACCESSDENIED = 0x80030005;
	const long STG_E_INVALIDHANDLE = 0x80030006;
	const long STG_E_INSUFFICIENTMEMORY = 0x80030008;
	const long STG_E_INVALIDPOINTER = 0x80030009;
	const long STG_E_NOMOREFILES = 0x80030012;
	const long STG_E_DISKISWRITEPROTECTED = 0x80030013;
	const long STG_E_SEEKERROR = 0x80030019;
	const long STG_E_WRITEFAULT = 0x8003001D;
	const long STG_E_READFAULT = 0x8003001E;
	const long STG_E_SHAREVIOLATION = 0x80030020;
	const long STG_E_LOCKVIOLATION = 0x80030021;
	const long STG_E_FILEALREADYEXISTS = 0x80030050;
	const long STG_E_INVALIDPARAMETER = 0x80030057;
	const long STG_E_MEDIUMFULL = 0x80030070;
	const long STG_E_ABNORMALAPIEXIT = 0x800300FA;
	const long STG_E_INVALIDHEADER = 0x800300FB;
	const long STG_E_INVALIDNAME = 0x800300FC;
	const long STG_E_UNKNOWN = 0x800300FD;
	const long STG_E_UNIMPLEMENTEDFUNCTION = 0x800300FE;
	const long STG_E_INVALIDFLAG = 0x800300FF;
	const long STG_E_INUSE = 0x80030100;
	const long STG_E_NOTCURRENT = 0x80030101;
	const long STG_E_REVERTED = 0x80030102;
	const long STG_E_CANTSAVE = 0x80030103;
	const long STG_E_OLDFORMAT = 0x80030104;
	const long STG_E_OLDDLL = 0x80030105;
	const long STG_E_SHAREREQUIRED = 0x80030106;
	const long STG_E_NOTFILEBASEDSTORAGE = 0x80030107;
	const long STG_E_EXTANTMARSHALLINGS = 0x80030108;
	const long STG_S_CONVERTED = 0x30200;

    const long RPC_E_CALL_REJECTED = 0x80010001;
    const long RPC_E_CALL_CANCELED = 0x80010002;
    const long RPC_E_CANTPOST_INSENDCALL = 0x80010003;
    const long RPC_E_CANTCALLOUT_INASYNCCALL = 0x80010004;
    const long RPC_E_CANTCALLOUT_INEXTERNALCALL = 0x80010005;
    const long RPC_E_CONNECTION_TERMINATED = 0x80010006;
    const long RPC_E_SERVER_DIED = 0x80010007;
    const long RPC_E_CLIENT_DIED = 0x80010008;
    const long RPC_E_INVALID_DATAPACKET = 0x80010009;
    const long RPC_E_CANTTRANSMIT_CALL = 0x8001000A;
    const long RPC_E_CLIENT_CANTMARSHAL_DATA = 0x8001000B;
    const long RPC_E_CLIENT_CANTUNMARSHAL_DATA = 0x8001000C;
    const long RPC_E_SERVER_CANTMARSHAL_DATA = 0x8001000D;
    const long RPC_E_SERVER_CANTUNMARSHAL_DATA = 0x8001000E;
    const long RPC_E_INVALID_DATA = 0x8001000F;
    const long RPC_E_INVALID_PARAMETER = 0x80010010;
    const long RPC_E_CANTCALLOUT_AGAIN = 0x80010011;
    const long RPC_E_SERVER_DIED_DNE = 0x80010012;
    const long RPC_E_SYS_CALL_FAILED = 0x80010100;
    const long RPC_E_OUT_OF_RESOURCES = 0x80010101;
    const long RPC_E_ATTEMPTED_MULTITHREAD = 0x80010102;
    const long RPC_E_NOT_REGISTERED = 0x80010103;
    const long RPC_E_FAULT = 0x80010104;
    const long RPC_E_SERVERFAULT = 0x80010105;
    const long RPC_E_CHANGED_MODE = 0x80010106;
    const long RPC_E_INVALIDMETHOD = 0x80010107;
    const long RPC_E_DISCONNECTED = 0x80010108;
    const long RPC_E_RETRY = 0x80010109;
    const long RPC_E_SERVERCALL_RETRYLATER = 0x8001010A;
    const long RPC_E_SERVERCALL_REJECTED = 0x8001010B;
    const long RPC_E_INVALID_CALLDATA = 0x8001010C;
    const long RPC_E_CANTCALLOUT_ININPUTSYNCCALL = 0x8001010D;
    const long RPC_E_WRONG_THREAD = 0x8001010E;
    const long RPC_E_THREAD_NOT_INIT = 0x8001010F;
    const long RPC_E_UNEXPECTED = 0x8001FFFF;


const long TXTBIT_RICHTEXT		=	1;		// rich-text control
const long TXTBIT_MULTILINE	=	2;		// single vs multi-line control
const long TXTBIT_READONLY		=	4;		// read only text
const long TXTBIT_SHOWACCELERATOR	= 8;		// underline accelerator character
const long TXTBIT_USEPASSWORD	=	0x10;	// use password char to display text
const long TXTBIT_HIDESELECTION =	0x20;	// show selection when inactive
const long TXTBIT_SAVESELECTION =	0x40;	// remember selection when inactive
const long TXTBIT_AUTOWORDSEL	=	0x80;	// auto-word selection 
const long TXTBIT_VERTICAL		=	0x100;	// vertical 
const long TXTBIT_SELBARCHANGE 	= 0x200;	// notification that the selection bar width 
										// has changed.
										// FUTURE: move this bit to the end to
										// maintain the division between 
										// properties and notifications.
const long TXTBIT_WORDWRAP  	=	0x400;	// if set, then multi-line controls
										// should wrap words to fit the available
										// display
const long	TXTBIT_ALLOWBEEP		= 0x800;	// enable/disable beeping
const long TXTBIT_DISABLEDRAG    =  0x1000;  // disable/enable dragging
const long TXTBIT_VIEWINSETCHANGE	= 0x2000;	// the inset changed
const long TXTBIT_BACKSTYLECHANGE	= 0x4000 ;
const long TXTBIT_MAXLENGTHCHANGE	= 0x8000;
const long TXTBIT_SCROLLBARCHANGE	 = 0x10000;
const long TXTBIT_CHARFORMATCHANGE = 0x20000;
const long TXTBIT_PARAFORMATCHANGE	= 0x40000;
const long TXTBIT_EXTENTCHANGE		= 0x80000;
const long TXTBIT_CLIENTRECTCHANGE	= 0x100000;	// the client rectangle changed
const long TXTBIT_USECURRENTBKG	= 0x200000;	// tells the renderer to use the current background
											// color rather than the system default for an entire line
const long TXTBIT_NOTHREADREFCOUNT	= 0x400000;	// don't reference TLS data on behalf of this instance
const long TXTBIT_SHOWPASSWORD	=	0x800000;	// Show password string
const long TXTBIT_D2DDWRITE		=	0x1000000;	// Use D2D/DWrite for this instance (and not GDI/Uniscribe).
const long TXTBIT_D2DSIMPLETYPOGRAPHY =	0x2000000;	// Don't glyph all. Only valid if D2DDWRITE is set.
const long TXTBIT_D2DPIXELSNAPPED	=	0x4000000;	// Don't do subpixel. Only valid if D2DDWRITE is set.
const long TXTBIT_D2DSUBPIXELLINES	=	0x8000000;	// Don't pixel-snap text lines and underline, strikethrough
												//  in the secondary text flow direction (usually vertical)
												//  Only valid if D2DDWRITE is set and D2DPIXELSNAPPED is not set.
const long TXTBIT_FLASHLASTPASSWORDCHAR = 0x10000000;	// Show last password char momentarily
const long TXTBIT_ADVANCEDINPUT =	0x20000000;	// Use advanced input features.

}
