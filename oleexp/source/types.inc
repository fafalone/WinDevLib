typedef enum GlobalConstants {
	MAX_PATH = 260
} GlobalConstants;

typedef enum GWL_INDEX {
	GWL_WNDPROC = (-4),
	GWL_HINSTANCE = (-6),
	GWL_HWNDPARENT = (-8),
	GWL_ID = (-12),
	GWL_STYLE = (-16),
	GWL_EXSTYLE = (-20),
	GWL_USERDATA = (-21)
} GWL_INDEX;

typedef struct FILETIME {
	LONG dwLowDateTime;
	LONG dwHighDateTime;
} FILETIME;


typedef struct LUID 
{  
    DWORD LowPart;  
    LONG HighPart;
} LUID;
typedef LUID* PLUID;

typedef enum FINDEX_SEARCH_OPS {
	FindExSearchNameMatch,
	FindExSearchLimitToDirectories,
	FindExSearchLimitToDevices
} FINDEX_SEARCH_OPS;

typedef enum FINDEX_INFO_LEVELS {
	FindExInfoStandard,
	FindExInfoBasic,
	FindExInfoMaxInfoLevel
} FINDEX_INFO_LEVELS;

typedef enum FINDEX_EXTRA_FLAGS {
	FIND_FIRST_EX_CASE_SENSITIVE  = 0x00000001,
	FIND_FIRST_EX_LARGE_FETCH     = 0x00000002

} FINDEX_EXTRA_FLAGS;

typedef enum FILE_ATTRIBUTES
{
	INVALID_FILE_ATTRIBUTES = -1,
	FILE_ATTRIBUTE_READONLY = 0x00000001,
	FILE_ATTRIBUTE_HIDDEN = 0x00000002,
	FILE_ATTRIBUTE_SYSTEM = 0x00000004,
	FILE_ATTRIBUTE_DIRECTORY = 0x00000010, //Read only
	FILE_ATTRIBUTE_ARCHIVE = 0x00000020,
	FILE_ATTRIBUTE_DEVICE = 0x00000040,
	FILE_ATTRIBUTE_NORMAL = 0x00000080,
	FILE_ATTRIBUTE_TEMPORARY = 0x00000100,
	FILE_ATTRIBUTE_SPARSE_FILE = 0x00000200, //Read only
	FILE_ATTRIBUTE_REPARSE_POINT = 0x00000400, //Read only
	FILE_ATTRIBUTE_COMPRESSED = 0x00000800,
	FILE_ATTRIBUTE_OFFLINE = 0x00001000,
	FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = 0x00002000,
	FILE_ATTRIBUTE_ENCRYPTED = 0x00004000,
	FILE_ATTRIBUTE_INTEGRITY_STREAM     = 0x00008000,
	FILE_ATTRIBUTE_PINNED               = 0x00008000,
	FILE_ATTRIBUTE_VIRTUAL           = 0x00010000,
	FILE_ATTRIBUTE_NO_SCRUB_DATA = 0x00020000,
	FILE_ATTRIBUTE_EA = 0x00040000,
	FILE_ATTRIBUTE_RECALL_ON_OPEN = 0x00040000,
	FILE_ATTRIBUTE_UNPINNED       = 0x00100000,
	FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS = 0x400000,
} FILE_ATTRIBUTES;

typedef enum FILE_FLAGS {
	FILE_FLAG_FIRST_PIPE_INSTANCE = 0x00080000,
	FILE_FLAG_OPEN_NO_RECALL = 0x00100000,
	FILE_FLAG_OPEN_REPARSE_POINT = 0x00200000,
	FILE_FLAG_SESSION_AWARE = 0x00800000,
	FILE_FLAG_POSIX_SEMANTICS = 0x01000000,
	FILE_FLAG_BACKUP_SEMANTICS = 0x02000000,
	FILE_FLAG_DELETE_ON_CLOSE = 0x04000000,
	FILE_FLAG_SEQUENTIAL_SCAN = 0x08000000,
	FILE_FLAG_RANDOM_ACCESS = 0x10000000,
	FILE_FLAG_NO_BUFFERING = 0x20000000,
	FILE_FLAG_OVERLAPPED = 0x40000000,
	FILE_FLAG_WRITE_THROUGH = 0x80000000,
} FILE_FLAGS;
typedef struct WIN32_FIND_DATAA {
	FILE_ATTRIBUTES dwFileAttributes;
	CURRENCY ftCreationTime;
	CURRENCY ftLastAccessTime;
	CURRENCY ftLastWriteTime;
	LONG nFileSizeHigh;
	LONG nFileSizeLow;
	LONG dwReserved0;
	LONG dwReserved1;
	BYTE cFileName[260];
	BYTE cAlternateFileName[14];
} WIN32_FIND_DATAA;

typedef struct WIN32_FIND_DATAW {
	FILE_ATTRIBUTES dwFileAttributes;
	CURRENCY ftCreationTime;
	CURRENCY ftLastAccessTime;
	CURRENCY ftLastWriteTime;
	LONG nFileSizeHigh;
	LONG nFileSizeLow;
	LONG dwReserved0;
	LONG dwReserved1;
	BYTE cFileName[520];
	BYTE cAlternateFileName[28];
} WIN32_FIND_DATAW;

typedef enum STREAM_INFO_LEVELS {
	FindStreamInfoStandard = 0,
	FindStreamInfoMaxInfoLevel = 1
} STREAM_INFO_LEVELS;

typedef struct WIN32_FIND_STREAM_DATA {
	LARGE_INTEGER StreamSize;
	BYTE          cStreamName[296]; // [MAX_PATH + 36];
} WIN32_FIND_STREAM_DATA;

typedef enum GET_FILEEX_INFO_LEVELS {
	GetFileExInfoStandard,
	GetFileExMaxInfoLevel
} GET_FILEEX_INFO_LEVELS;
typedef struct WIN32_FILE_ATTRIBUTE_DATA {
	DWORD    dwFileAttributes;
	FILETIME ftCreationTime;
	FILETIME ftLastAccessTime;
	FILETIME ftLastWriteTime;
	DWORD    nFileSizeHigh;
	DWORD    nFileSizeLow;
} WIN32_FILE_ATTRIBUTE_DATA;

typedef struct BY_HANDLE_FILE_INFORMATION {
	DWORD    dwFileAttributes;
	FILETIME ftCreationTime;
	FILETIME ftLastAccessTime;
	FILETIME ftLastWriteTime;
	DWORD    dwVolumeSerialNumber;
	DWORD    nFileSizeHigh;
	DWORD    nFileSizeLow;
	DWORD    nNumberOfLinks;
	DWORD    nFileIndexHigh;
	DWORD    nFileIndexLow;
} BY_HANDLE_FILE_INFORMATION;

typedef enum FILE_INFORMATION_CLASS {
	FileDirectoryInformation = 1,
	FileFullDirectoryInformation,
	FileBothDirectoryInformation,
	FileBasicInformation,
	FileStandardInformation,
	FileInternalInformation,
	FileEaInformation,
	FileAccessInformation,
	FileNameInformation,
	FileRenameInformation,
	FileLinkInformation,
	FileNamesInformation,
	FileDispositionInformation,
	FilePositionInformation,
	FileFullEaInformation,
	FileModeInformation,
	FileAlignmentInformation,
	FileAllInformation,
	FileAllocationInformation,
	FileEndOfFileInformation,
	FileAlternateNameInformation,
	FileStreamInformation,
	FilePipeInformation,
	FilePipeLocalInformation,
	FilePipeRemoteInformation,
	FileMailslotQueryInformation,
	FileMailslotSetInformation,
	FileCompressionInformation,
	FileObjectIdInformation,
	FileCompletionInformation,
	FileMoveClusterInformation,
	FileQuotaInformation,
	FileReparsePointInformation,
	FileNetworkOpenInformation,
	FileAttributeTagInformation,
	FileTrackingInformation,
	FileIdBothDirectoryInformation,
	FileIdFullDirectoryInformation,
	FileValidDataLengthInformation,
	FileShortNameInformation,
	FileIoCompletionNotificationInformation,
	FileIoStatusBlockRangeInformation,
	FileIoPriorityHintInformation,
	FileSfioReserveInformation,
	FileSfioVolumeInformation,
	FileHardLinkInformation,
	FileProcessIdsUsingFileInformation,
	FileNormalizedNameInformation,
	FileNetworkPhysicalNameInformation,
	FileIdGlobalTxDirectoryInformation,
	FileIsRemoteDeviceInformation,
	FileUnusedInformation,
	FileNumaNodeInformation,
	FileStandardLinkInformation,
	FileRemoteProtocolInformation,
	FileRenameInformationBypassAccessCheck,
	FileLinkInformationBypassAccessCheck,
	FileVolumeNameInformation,
	FileIdInformation,
	FileIdExtdDirectoryInformation,
	FileReplaceCompletionInformation,
	FileHardLinkFullIdInformation,
	FileIdExtdBothDirectoryInformation,
	FileMaximumInformation
} FILE_INFORMATION_CLASS;

typedef enum FILE_INFO_BY_HANDLE_CLASS {
	FileBasicInfo = 0,
	FileStandardInfo = 1,
	FileNameInfo = 2,
	FileRenameInfo = 3,
	FileDispositionInfo = 4,
	FileAllocationInfo = 5,
	FileEndOfFileInfo = 6,
	FileStreamInfo = 7,
	FileCompressionInfo = 8,
	FileAttributeTagInfo = 9,
	FileIdBothDirectoryInfo = 10, // 0xA
	FileIdBothDirectoryRestartInfo = 11, // 0xB
	FileIoPriorityHintInfo = 12, // 0xC
	FileRemoteProtocolInfo = 13, // 0xD
	FileFullDirectoryInfo = 14, // 0xE
	FileFullDirectoryRestartInfo = 15, // 0xF
	FileStorageInfo = 16, // 0x10
	FileAlignmentInfo = 17, // 0x11
	FileIdInfo = 18, // 0x12
	FileIdExtdDirectoryInfo = 19, // 0x13
	FileIdExtdDirectoryRestartInfo = 20, // 0x14
	MaximumFileInfoByHandlesClass
} FILE_INFO_BY_HANDLE_CLASS;

typedef struct FILE_BASIC_INFO {
	LARGE_INTEGER CreationTime;
	LARGE_INTEGER LastAccessTime;
	LARGE_INTEGER LastWriteTime;
	LARGE_INTEGER ChangeTime;
	DWORD         FileAttributes;
	BYTE   btAlign[4];
} FILE_BASIC_INFO;

typedef struct FILE_STREAM_INFO {
	DWORD         NextEntryOffset;
	DWORD         StreamNameLength;
	LARGE_INTEGER StreamSize;
	LARGE_INTEGER StreamAllocationSize; //16
	WCHAR         StreamName[1];
} FILE_STREAM_INFO;

typedef struct FILE_FULL_DIR_INFO {
	ULONG         NextEntryOffset;
	ULONG         FileIndex;
	LARGE_INTEGER CreationTime;
	LARGE_INTEGER LastAccessTime;
	LARGE_INTEGER LastWriteTime;
	LARGE_INTEGER ChangeTime;
	LARGE_INTEGER EndOfFile;
	LARGE_INTEGER AllocationSize;
	ULONG         FileAttributes;
	ULONG         FileNameLength;
	ULONG         EaSize;
	BYTE     btAlign[4]; //offset=72
	WCHAR         FileName[1];
} FILE_FULL_DIR_INFO;

typedef struct FILE_DIRECTORY_INFORMATION {
	LONG NextEntryOffset;
	LONG FileIndex;
	LARGE_INTEGER CreationTime;
	LARGE_INTEGER LastAccessTime;
	LARGE_INTEGER LastWriteTime;
	LARGE_INTEGER ChangeTime;
	LARGE_INTEGER EndOfFile;
	LARGE_INTEGER AllocationSize;
	LONG FileAttributes;
	LONG FileNameLength;
	WORD FileName1;
} FILE_DIRECTORY_INFORMATION;

typedef struct COMDLG_FILTERSPEC
{
	BSTR pszName;
	BSTR pszSpec;
} COMDLG_FILTERSPEC;

typedef enum ASSOC_FILTER
{
    ASSOC_FILTER_NONE = 0,
    ASSOC_FILTER_RECOMMENDED
} ASSOC_FILTER;

typedef enum SHOWWINDOW {
	SW_HIDE = 0,
	SW_SHOWNORMAL = 1,
	SW_NORMAL = 1,
	SW_SHOWMINIMIZED = 2,
	SW_SHOWMAXIMIZED = 3,
	SW_MAXIMIZE = 3,
	SW_SHOWNOACTIVATE = 4,
	SW_SHOW = 5,
	SW_MINIMIZE = 6,
	SW_SHOWMINNOACTIVE = 7,
	SW_SHOWNA = 8,
	SW_RESTORE = 9,
	SW_SHOWDEFAULT = 10
} SHOWWINDOW;

typedef enum SHIL
{
	SHIL_LARGE = 0,
	SHIL_SMALL,
	SHIL_EXTRALARGE,
	SHIL_SYSSMALL,
	SHIL_JUMBO,
	SHIL_LAST
} SHIL;

typedef enum SystemMetrics
{
	SM_CXSCREEN            = 0,  // 0x00
	SM_CYSCREEN            = 1,  // 0x01
	SM_CXVSCROLL           = 2,  // 0x02
	SM_CYHSCROLL           = 3,  // 0x03
	SM_CYCAPTION           = 4,  // 0x04
	SM_CXBORDER            = 5,  // 0x05
	SM_CYBORDER            = 6,  // 0x06
	SM_CXDLGFRAME          = 7,  // 0x07
	SM_CXFIXEDFRAME        = 7,  // 0x07
	SM_CYDLGFRAME          = 8,  // 0x08
	SM_CYFIXEDFRAME        = 8,  // 0x08
	SM_CYVTHUMB            = 9,  // 0x09
	SM_CXHTHUMB            = 10, // 0x0A
	SM_CXICON              = 11, // 0x0B
	SM_CYICON              = 12, // 0x0C
	SM_CXCURSOR            = 13, // 0x0D
	SM_CYCURSOR            = 14, // 0x0E
	SM_CYMENU              = 15, // 0x0F
	SM_CXFULLSCREEN        = 16, // 0x10
	SM_CYFULLSCREEN        = 17, // 0x11
	SM_CYKANJIWINDOW       = 18, // 0x12
	SM_MOUSEPRESENT        = 19, // 0x13
	SM_CYVSCROLL           = 20, // 0x14
	SM_CXHSCROLL           = 21, // 0x15
	SM_DEBUG               = 22, // 0x16
	SM_SWAPBUTTON          = 23, // 0x17
	SM_CXMIN               = 28, // 0x1C
	SM_CYMIN               = 29, // 0x1D
	SM_CXSIZE              = 30, // 0x1E
	SM_CYSIZE              = 31, // 0x1F
	SM_CXSIZEFRAME         = 32, // 0x20
	SM_CXFRAME             = 32, // 0x20
	SM_CYSIZEFRAME         = 33, // 0x21
	SM_CYFRAME             = 33, // 0x21
	SM_CXMINTRACK          = 34, // 0x22
	SM_CYMINTRACK          = 35, // 0x23
	SM_CXDOUBLECLK         = 36, // 0x24
	SM_CYDOUBLECLK         = 37, // 0x25
	SM_CXICONSPACING       = 38, // 0x26
	SM_CYICONSPACING       = 39, // 0x27
	SM_MENUDROPALIGNMENT   = 40, // 0x28
	SM_PENWINDOWS          = 41, // 0x29
	SM_DBCSENABLED         = 42, // 0x2A
	SM_CMOUSEBUTTONS       = 43, // 0x2B
	SM_SECURE              = 44, // 0x2C
	SM_CXEDGE              = 45, // 0x2D
	SM_CYEDGE              = 46, // 0x2E
	SM_CXMINSPACING        = 47, // 0x2F
	SM_CYMINSPACING        = 48, // 0x30
	SM_CXSMICON            = 49, // 0x31
	SM_CYSMICON            = 50, // 0x32
	SM_CYSMCAPTION         = 51, // 0x33
	SM_CXSMSIZE            = 52, // 0x34
	SM_CYSMSIZE            = 53, // 0x35
	SM_CXMENUSIZE          = 54, // 0x36
	SM_CYMENUSIZE          = 55, // 0x37
	SM_ARRANGE             = 56, // 0x38
	SM_CXMINIMIZED         = 57, // 0x39
	SM_CYMINIMIZED         = 58, // 0x3A
	SM_CXMAXTRACK          = 59, // 0x3B
	SM_CYMAXTRACK          = 60, // 0x3C
	SM_CXMAXIMIZED         = 61, // 0x3D
	SM_CYMAXIMIZED         = 62, // 0x3E
	SM_NETWORK             = 63, // 0x3F
	SM_CLEANBOOT           = 67, // 0x43
	SM_CXDRAG              = 68, // 0x44
	SM_CYDRAG              = 69, // 0x45
	SM_SHOWSOUNDS          = 70, // 0x46
	SM_CXMENUCHECK         = 71, // 0x47
	SM_CYMENUCHECK         = 72, // 0x48
	SM_SLOWMACHINE         = 73, // 0x49
	SM_MIDEASTENABLED      = 74, // 0x4A
	SM_MOUSEWHEELPRESENT   = 75, // 0x4B
	SM_XVIRTUALSCREEN      = 76, // 0x4C
	SM_YVIRTUALSCREEN      = 77, // 0x4D
	SM_CXVIRTUALSCREEN     = 78, // 0x4E
	SM_CYVIRTUALSCREEN     = 79, // 0x4F
	SM_CMONITORS           = 80, // 0x50
	SM_SAMEDISPLAYFORMAT   = 81, // 0x51
	SM_IMMENABLED          = 82, // 0x52
	SM_CXFOCUSBORDER       = 83, // 0x53
	SM_CYFOCUSBORDER       = 84, // 0x54
	SM_TABLETPC            = 86, // 0x56
	SM_MEDIACENTER         = 87, // 0x57
	SM_STARTER             = 88, // 0x58
	SM_SERVERR2            = 89, // 0x59
	SM_MOUSEHORIZONTALWHEELPRESENT = 91, // 0x5B
	SM_CXPADDEDBORDER      = 92, // 0x5C
	SM_DIGITIZER           = 94, // 0x5E
	SM_MAXIMUMTOUCHES      = 95, // 0x5F

	SM_REMOTESESSION       = 0x1000, // 0x1000
	SM_SHUTTINGDOWN        = 0x2000, // 0x2000
	SM_REMOTECONTROL       = 0x2001, // 0x2001


	SM_CONVERTABLESLATEMODE = 0x2003,
	SM_SYSTEMDOCKED        = 0x2004,
} SystemMetrics;

typedef enum SPI {
     SPI_GETBEEP                 = 0x0001,
     SPI_SETBEEP                 = 0x0002,
     SPI_GETMOUSE                = 0x0003,
     SPI_SETMOUSE                = 0x0004,
     SPI_GETBORDER               = 0x0005,
     SPI_SETBORDER               = 0x0006,
     SPI_GETKEYBOARDSPEED        = 0x000A,
     SPI_SETKEYBOARDSPEED        = 0x000B,
     SPI_LANGDRIVER              = 0x000C,
     SPI_ICONHORIZONTALSPACING   = 0x000D,
     SPI_GETSCREENSAVETIMEOUT    = 0x000E,
     SPI_SETSCREENSAVETIMEOUT    = 0x000F,
     SPI_GETSCREENSAVEACTIVE     = 0x0010,
     SPI_SETSCREENSAVEACTIVE     = 0x0011,
     SPI_GETGRIDGRANULARITY      = 0x0012,
     SPI_SETGRIDGRANULARITY      = 0x0013,
     SPI_SETDESKWALLPAPER        = 0x0014,
     SPI_SETDESKPATTERN          = 0x0015,
     SPI_GETKEYBOARDDELAY        = 0x0016,
     SPI_SETKEYBOARDDELAY        = 0x0017,
     SPI_ICONVERTICALSPACING     = 0x0018,
     SPI_GETICONTITLEWRAP        = 0x0019,
     SPI_SETICONTITLEWRAP        = 0x001A,
     SPI_GETMENUDROPALIGNMENT    = 0x001B,
     SPI_SETMENUDROPALIGNMENT    = 0x001C,
     SPI_SETDOUBLECLKWIDTH       = 0x001D,
     SPI_SETDOUBLECLKHEIGHT      = 0x001E,
     SPI_GETICONTITLELOGFONT     = 0x001F,
     SPI_SETDOUBLECLICKTIME      = 0x0020,
     SPI_SETMOUSEBUTTONSWAP      = 0x0021,
     SPI_SETICONTITLELOGFONT     = 0x0022,
     SPI_GETFASTTASKSWITCH       = 0x0023,
     SPI_SETFASTTASKSWITCH       = 0x0024,
     SPI_SETDRAGFULLWINDOWS      = 0x0025,
     SPI_GETDRAGFULLWINDOWS      = 0x0026,
     SPI_GETNONCLIENTMETRICS     = 0x0029,
     SPI_SETNONCLIENTMETRICS     = 0x002A,
     SPI_GETMINIMIZEDMETRICS     = 0x002B,
     SPI_SETMINIMIZEDMETRICS     = 0x002C,
     SPI_GETICONMETRICS          = 0x002D,
     SPI_SETICONMETRICS          = 0x002E,
     SPI_SETWORKAREA             = 0x002F,
     SPI_GETWORKAREA             = 0x0030,
     SPI_SETPENWINDOWS           = 0x0031,

     SPI_GETHIGHCONTRAST         = 0x0042,
     SPI_SETHIGHCONTRAST         = 0x0043,
     SPI_GETKEYBOARDPREF         = 0x0044,
     SPI_SETKEYBOARDPREF         = 0x0045,
     SPI_GETSCREENREADER         = 0x0046,
     SPI_SETSCREENREADER         = 0x0047,
     SPI_GETANIMATION            = 0x0048,
     SPI_SETANIMATION            = 0x0049,
     SPI_GETFONTSMOOTHING        = 0x004A,
     SPI_SETFONTSMOOTHING        = 0x004B,
     SPI_SETDRAGWIDTH            = 0x004C,
     SPI_SETDRAGHEIGHT           = 0x004D,
     SPI_SETHANDHELD             = 0x004E,
     SPI_GETLOWPOWERTIMEOUT      = 0x004F,
     SPI_GETPOWEROFFTIMEOUT      = 0x0050,
     SPI_SETLOWPOWERTIMEOUT      = 0x0051,
     SPI_SETPOWEROFFTIMEOUT      = 0x0052,
     SPI_GETLOWPOWERACTIVE       = 0x0053,
     SPI_GETPOWEROFFACTIVE       = 0x0054,
     SPI_SETLOWPOWERACTIVE       = 0x0055,
     SPI_SETPOWEROFFACTIVE       = 0x0056,
     SPI_SETCURSORS              = 0x0057,
     SPI_SETICONS                = 0x0058,
     SPI_GETDEFAULTINPUTLANG     = 0x0059,
     SPI_SETDEFAULTINPUTLANG     = 0x005A,
     SPI_SETLANGTOGGLE           = 0x005B,
     SPI_GETWINDOWSEXTENSION     = 0x005C,
     SPI_SETMOUSETRAILS          = 0x005D,
     SPI_GETMOUSETRAILS          = 0x005E,
     SPI_SETSCREENSAVERRUNNING   = 0x0061,
     SPI_SCREENSAVERRUNNING     = 0x0061,   //SPI_SETSCREENSAVERRUNNING
     SPI_GETFILTERKEYS          = 0x0032,
     SPI_SETFILTERKEYS          = 0x0033,
     SPI_GETTOGGLEKEYS          = 0x0034,
     SPI_SETTOGGLEKEYS          = 0x0035,
     SPI_GETMOUSEKEYS           = 0x0036,
     SPI_SETMOUSEKEYS           = 0x0037,
     SPI_GETSHOWSOUNDS          = 0x0038,
     SPI_SETSHOWSOUNDS          = 0x0039,
     SPI_GETSTICKYKEYS          = 0x003A,
     SPI_SETSTICKYKEYS          = 0x003B,
     SPI_GETACCESSTIMEOUT       = 0x003C,
     SPI_SETACCESSTIMEOUT       = 0x003D,
     SPI_GETSERIALKEYS          = 0x003E,
     SPI_SETSERIALKEYS          = 0x003F,
     SPI_GETSOUNDSENTRY         = 0x0040,
     SPI_SETSOUNDSENTRY         = 0x0041,
     SPI_GETSNAPTODEFBUTTON     = 0x005F,
     SPI_SETSNAPTODEFBUTTON     = 0x0060,
     SPI_GETMOUSEHOVERWIDTH     = 0x0062,
     SPI_SETMOUSEHOVERWIDTH     = 0x0063,
     SPI_GETMOUSEHOVERHEIGHT    = 0x0064,
     SPI_SETMOUSEHOVERHEIGHT    = 0x0065,
     SPI_GETMOUSEHOVERTIME      = 0x0066,
     SPI_SETMOUSEHOVERTIME      = 0x0067,
     SPI_GETWHEELSCROLLLINES    = 0x0068,
     SPI_SETWHEELSCROLLLINES    = 0x0069,
     SPI_GETMENUSHOWDELAY       = 0x006A,
     SPI_SETMENUSHOWDELAY       = 0x006B,

     SPI_GETWHEELSCROLLCHARS   = 0x006C,
     SPI_SETWHEELSCROLLCHARS   = 0x006D,

     SPI_GETSHOWIMEUI          = 0x006E,
     SPI_SETSHOWIMEUI          = 0x006F,

     SPI_GETMOUSESPEED         = 0x0070,
     SPI_SETMOUSESPEED         = 0x0071,
     SPI_GETSCREENSAVERRUNNING = 0x0072,
     SPI_GETDESKWALLPAPER      = 0x0073,
	 
     SPI_GETAUDIODESCRIPTION   = 0x0074,
     SPI_SETAUDIODESCRIPTION   = 0x0075,

     SPI_GETSCREENSAVESECURE   = 0x0076,
     SPI_SETSCREENSAVESECURE   = 0x0077,

     SPI_GETHUNGAPPTIMEOUT           = 0x0078,
     SPI_SETHUNGAPPTIMEOUT           = 0x0079,
     SPI_GETWAITTOKILLTIMEOUT        = 0x007A,
     SPI_SETWAITTOKILLTIMEOUT        = 0x007B,
     SPI_GETWAITTOKILLSERVICETIMEOUT = 0x007C,
     SPI_SETWAITTOKILLSERVICETIMEOUT = 0x007D,
     SPI_GETMOUSEDOCKTHRESHOLD       = 0x007E,
     SPI_SETMOUSEDOCKTHRESHOLD       = 0x007F,
     SPI_GETPENDOCKTHRESHOLD         = 0x0080,
     SPI_SETPENDOCKTHRESHOLD         = 0x0081,
     SPI_GETWINARRANGING             = 0x0082,
     SPI_SETWINARRANGING             = 0x0083,
     SPI_GETMOUSEDRAGOUTTHRESHOLD    = 0x0084,
     SPI_SETMOUSEDRAGOUTTHRESHOLD    = 0x0085,
     SPI_GETPENDRAGOUTTHRESHOLD      = 0x0086,
     SPI_SETPENDRAGOUTTHRESHOLD      = 0x0087,
     SPI_GETMOUSESIDEMOVETHRESHOLD   = 0x0088,
     SPI_SETMOUSESIDEMOVETHRESHOLD   = 0x0089,
     SPI_GETPENSIDEMOVETHRESHOLD     = 0x008A,
     SPI_SETPENSIDEMOVETHRESHOLD     = 0x008B,
     SPI_GETDRAGFROMMAXIMIZE         = 0x008C,
     SPI_SETDRAGFROMMAXIMIZE         = 0x008D,
     SPI_GETSNAPSIZING               = 0x008E,
     SPI_SETSNAPSIZING               = 0x008F,
     SPI_GETDOCKMOVING               = 0x0090,
     SPI_SETDOCKMOVING               = 0x0091,

	SPI_GETTOUCHPREDICTIONPARAMETERS = 0x009C,
    SPI_SETTOUCHPREDICTIONPARAMETERS = 0x009D,
    SPI_GETLOGICALDPIOVERRIDE        = 0x009E,
    SPI_SETLOGICALDPIOVERRIDE        = 0x009F,


    SPI_GETMENURECT                     = 0x00A2,
    SPI_SETMENURECT                     = 0x00A3,
    SPI_GETACTIVEWINDOWTRACKING         = 0x1000,
    SPI_SETACTIVEWINDOWTRACKING         = 0x1001,
    SPI_GETMENUANIMATION                = 0x1002,
    SPI_SETMENUANIMATION                = 0x1003,
    SPI_GETCOMBOBOXANIMATION            = 0x1004,
    SPI_SETCOMBOBOXANIMATION            = 0x1005,
    SPI_GETLISTBOXSMOOTHSCROLLING       = 0x1006,
    SPI_SETLISTBOXSMOOTHSCROLLING       = 0x1007,
    SPI_GETGRADIENTCAPTIONS             = 0x1008,
    SPI_SETGRADIENTCAPTIONS             = 0x1009,
    SPI_GETKEYBOARDCUES                 = 0x100A,
    SPI_SETKEYBOARDCUES                 = 0x100B,
    SPI_GETMENUUNDERLINES               = 0x100A,  //SPI_GETKEYBOARDCUES
    SPI_SETMENUUNDERLINES               = 0x100B,  //SPI_SETKEYBOARDCUES
    SPI_GETACTIVEWNDTRKZORDER           = 0x100C,
    SPI_SETACTIVEWNDTRKZORDER           = 0x100D,
    SPI_GETHOTTRACKING                  = 0x100E,
    SPI_SETHOTTRACKING                  = 0x100F,
    SPI_GETMENUFADE                     = 0x1012,
    SPI_SETMENUFADE                     = 0x1013,
    SPI_GETSELECTIONFADE                = 0x1014,
    SPI_SETSELECTIONFADE                = 0x1015,
    SPI_GETTOOLTIPANIMATION             = 0x1016,
    SPI_SETTOOLTIPANIMATION             = 0x1017,
    SPI_GETTOOLTIPFADE                  = 0x1018,
    SPI_SETTOOLTIPFADE                  = 0x1019,
    SPI_GETCURSORSHADOW                 = 0x101A,
    SPI_SETCURSORSHADOW                 = 0x101B,
    SPI_GETMOUSESONAR                   = 0x101C,
    SPI_SETMOUSESONAR                   = 0x101D,
    SPI_GETMOUSECLICKLOCK               = 0x101E,
    SPI_SETMOUSECLICKLOCK               = 0x101F,
    SPI_GETMOUSEVANISH                  = 0x1020,
    SPI_SETMOUSEVANISH                  = 0x1021,
    SPI_GETFLATMENU                     = 0x1022,
    SPI_SETFLATMENU                     = 0x1023,
    SPI_GETDROPSHADOW                   = 0x1024,
    SPI_SETDROPSHADOW                   = 0x1025,
    SPI_GETBLOCKSENDINPUTRESETS         = 0x1026,
    SPI_SETBLOCKSENDINPUTRESETS         = 0x1027,

    SPI_GETUIEFFECTS                    = 0x103E,
    SPI_SETUIEFFECTS                    = 0x103F,

    SPI_GETDISABLEOVERLAPPEDCONTENT     = 0x1040,
    SPI_SETDISABLEOVERLAPPEDCONTENT     = 0x1041,
    SPI_GETCLIENTAREAANIMATION          = 0x1042,
    SPI_SETCLIENTAREAANIMATION          = 0x1043,
    SPI_GETCLEARTYPE                    = 0x1048,
    SPI_SETCLEARTYPE                    = 0x1049,
    SPI_GETSPEECHRECOGNITION            = 0x104A,
    SPI_SETSPEECHRECOGNITION            = 0x104B,

    SPI_GETCARETBROWSING                = 0x104C,
    SPI_SETCARETBROWSING                = 0x104D,
    SPI_GETTHREADLOCALINPUTSETTINGS     = 0x104E,
    SPI_SETTHREADLOCALINPUTSETTINGS     = 0x104F,
    SPI_GETSYSTEMLANGUAGEBAR            = 0x1050,
    SPI_SETSYSTEMLANGUAGEBAR            = 0x1051,

    SPI_GETFOREGROUNDLOCKTIMEOUT        = 0x2000,
    SPI_SETFOREGROUNDLOCKTIMEOUT        = 0x2001,
    SPI_GETACTIVEWNDTRKTIMEOUT          = 0x2002,
    SPI_SETACTIVEWNDTRKTIMEOUT          = 0x2003,
    SPI_GETFOREGROUNDFLASHCOUNT         = 0x2004,
    SPI_SETFOREGROUNDFLASHCOUNT         = 0x2005,
    SPI_GETCARETWIDTH                   = 0x2006,
    SPI_SETCARETWIDTH                   = 0x2007,

    SPI_GETMOUSECLICKLOCKTIME           = 0x2008,
    SPI_SETMOUSECLICKLOCKTIME           = 0x2009,
    SPI_GETFONTSMOOTHINGTYPE            = 0x200A,
    SPI_SETFONTSMOOTHINGTYPE            = 0x200B,

    SPI_GETFONTSMOOTHINGCONTRAST           = 0x200C,
    SPI_SETFONTSMOOTHINGCONTRAST           = 0x200D,

    SPI_GETFOCUSBORDERWIDTH             = 0x200E,
    SPI_SETFOCUSBORDERWIDTH             = 0x200F,
    SPI_GETFOCUSBORDERHEIGHT            = 0x2010,
    SPI_SETFOCUSBORDERHEIGHT            = 0x2011,

    SPI_GETFONTSMOOTHINGORIENTATION           = 0x2012,
    SPI_SETFONTSMOOTHINGORIENTATION           = 0x2013,

    SPI_GETMINIMUMHITRADIUS             = 0x2014,
    SPI_SETMINIMUMHITRADIUS             = 0x2015,
    SPI_GETMESSAGEDURATION              = 0x2016,
    SPI_SETMESSAGEDURATION              = 0x2017,

    SPI_GETCONTACTVISUALIZATION         = 0x2018,
    SPI_SETCONTACTVISUALIZATION         = 0x2019,

    SPI_GETGESTUREVISUALIZATION         = 0x201A,
    SPI_SETGESTUREVISUALIZATION         = 0x201B,

    SPI_GETMOUSEWHEELROUTING            = 0x201C,
    SPI_SETMOUSEWHEELROUTING            = 0x201D,

    SPI_GETPENVISUALIZATION                  = 0x201E,
    SPI_SETPENVISUALIZATION                  = 0x201F

} SPI;

typedef enum SPIF
{
    // <summary>Writes the new system-wide parameter setting to the user profile.</summary>
    SPIF_UPDATEINIFILE = 0x01,
    // <summary>Broadcasts the WM_SETTINGCHANGE message after updating the user profile.</summary>
    SPIF_SENDCHANGE = 0x02,
    // <summary>Same as SPIF_SENDCHANGE.</summary>
    SPIF_SENDWININICHANGE = 0x02

} SPIF;



typedef enum SFS_MASK {
	SSF_SHOWALLOBJECTS          = 0x00000001,
	SSF_SHOWEXTENSIONS          = 0x00000002,
	SSF_HIDDENFILEEXTS          = 0x00000004,
	SSF_SERVERADMINUI           = 0x00000004,
	SSF_SHOWCOMPCOLOR           = 0x00000008,
	SSF_SORTCOLUMNS             = 0x00000010,
	SSF_SHOWSYSFILES            = 0x00000020,
	SSF_DOUBLECLICKINWEBVIEW    = 0x00000080,
	SSF_SHOWATTRIBCOL           = 0x00000100,
	SSF_DESKTOPHTML             = 0x00000200,
	SSF_WIN95CLASSIC            = 0x00000400,
	SSF_DONTPRETTYPATH          = 0x00000800,
	SSF_SHOWINFOTIP             = 0x00002000,
	SSF_MAPNETDRVBUTTON         = 0x00001000,
	SSF_NOCONFIRMRECYCLE        = 0x00008000,
	SSF_HIDEICONS               = 0x00004000,
	SSF_FILTER                  = 0x00010000,
	SSF_WEBVIEW                 = 0x00020000,
	SSF_SHOWSUPERHIDDEN         = 0x00040000,
	SSF_SEPPROCESS              = 0x00080000,
//#if (NTDDI_WINXP)
	SSF_NONETCRAWLING           = 0x00100000,
	SSF_STARTPANELON            = 0x00200000,
	SSF_SHOWSTARTPAGE           = 0x00400000,
//#endif  // NTDDI_WINXP
//#if (NTDDI_VISTA)
	SSF_AUTOCHECKSELECT         = 0x00800000,
	SSF_ICONSONLY               = 0x01000000,
	SSF_SHOWTYPEOVERLAY         = 0x02000000,
//#endif  // NTDDI_VISTA
//#if (NTDDI_WIN8)
	SSF_SHOWSTATUSBAR           = 0x04000000
} SFS_MASK;

typedef struct SHELLFLAGSTATE {
	BOOLEAN fShowAllObjects;
	BOOLEAN fShowExtensions;
	BOOLEAN fNoConfirmRecycle;
	BOOLEAN fShowSysFiles;
	BOOLEAN fShowCompColor;
	BOOLEAN fDoubleClickInWebView;
	BOOLEAN fDesktopHTML;
	BOOLEAN fWin95Classic;
	BOOLEAN fDontPrettyPath;
	BOOLEAN fShowAttribCol;
	BOOLEAN fMapNetDrvBtn;
	BOOLEAN fShowInfoTip;
	BOOLEAN fHideIcons1;
//#if (NTDDI_VERSION >= NTDDI_VISTA)
	BOOLEAN fAutoCheckSelect;
	BOOLEAN fIconsOnly;
	UINT fRestFlags; // when adding additional flags keep SHELLSTATE and SHGetSettings in sync.
//#else
//	UINT fRestFlags : 3; // when adding additional flags keep SHELLSTATE and SHGetSettings in sync.
//#endif
} SHELLFLAGSTATE;

typedef enum SSFLAGS1 {
	fShowAllObjects = 0x1,
	fShowExtensions = 0x2,
	fNoConfirmRecycle = 0x4,
	fShowSysFiles = 0x8,
	fShowCompColor = 0x10,
	fDoubleClickInWebView = 0x20,
	fDesktopHTML = 0x40,
	fWin95Classic = 0x80,
	fDontPrettyPath = 0x100,
	fShowAttribCol = 0x200,
	fMapNetDrvBtn = 0x400,
	fShowInfoTip = 0x800,
	fHideIcons = 0x1000,
	fWebView = 0x2000,
	fFilter = 0x4000,
	fShowSuperHidden = 0x8000,
	fNoNetCrawling = 0x10000
} SSFLAGS1;
typedef enum SSFLAGS2 {
	fSepProcess = 0x1,
	fStartPanelOn = 0x2,
	fShowStartPage = 0x4,
	fAutoCheckSelect = 0x8,
	fIconsOnly = 0x10,
	fShowTypeOverlay = 0x20,
	fShowStatusBar = 0x40
} SSFLAGS2;
typedef struct SHELLSTATE {
	SSFLAGS1 fFlags1;
	DWORD dwWin95Unused;
	UINT uWin95Unused;
	LONG lParamSort;
	LONG iSortDirection;
	LONG Version;
	UINT NotUsed;
	SSFLAGS2 fFlags2;
} SHELLSTATE;

typedef enum STGFMT { 
  STGFMT_STORAGE  = 0,
  STGFMT_FILE     = 3,
  STGFMT_ANY      = 4,
  STGFMT_DOCFILE  = 5
} STGFMT;

typedef enum DATAOBJ_GET_ITEM_FLAGS { 
  DOGIF_DEFAULT        = 0x0000,
  DOGIF_TRAVERSE_LINK  = 0x0001,
  DOGIF_NO_HDROP       = 0x0002,
  DOGIF_NO_URL         = 0x0004,
  DOGIF_ONLY_IF_ONE    = 0x0008
} DATAOBJ_GET_ITEM_FLAGS;

typedef enum KNOWN_FOLDER_FLAG { 
  KF_FLAG_DEFAULT                      = 0x00000000,
  KF_FLAG_SIMPLE_IDLIST                = 0x00000100,
  KF_FLAG_NOT_PARENT_RELATIVE          = 0x00000200,
  KF_FLAG_DEFAULT_PATH                 = 0x00000400,
  KF_FLAG_INIT                         = 0x00000800,
  KF_FLAG_NO_ALIAS                     = 0x00001000,
  KF_FLAG_DONT_UNEXPAND                = 0x00002000,
  KF_FLAG_DONT_VERIFY                  = 0x00004000,
  KF_FLAG_CREATE                       = 0x00008000,
  KF_FLAG_NO_APPCONTAINER_REDIRECTION  = 0x00010000,
  KF_FLAG_NO_PACKAGE_REDIRECTION       = 0x00010000,
  KF_FLAG_FORCE_PACKAGE_REDIRECTION    = 0x00020000,
  KF_FLAG_FORCE_APPCONTAINER_REDIRECTION = 0x00020000,
  KF_FLAG_RETURN_FILTER_REDIRECTION_TARGET = 0x00040000,
  KF_FLAG_FORCE_APP_DATA_REDIRECTION       = 0x00080000,
  KF_FLAG_ALIAS_ONLY                       = 0x80000000
} KNOWN_FOLDER_FLAG;

typedef [helpstring("IShellFolder.GetAttributesOf flags")]
enum SFGAO_Flags {
    [helpstring("Objects can be copied")]
    SFGAO_CANCOPY                    = 0x00000001,
    [helpstring("Objects can be moved")]
    SFGAO_CANMOVE                    = 0x00000002,
    [helpstring("Objects can be linked")]
    SFGAO_CANLINK                    = 0x00000004,
    [helpstring("Supports BindToObject(IID_IStorage")]
    SFGAO_STORAGE                    = 0x00000008,
    [helpstring("Objects can be renamed")]
    SFGAO_CANRENAME                  = 0x00000010,
    [helpstring("Objects can be deleted")]
    SFGAO_CANDELETE                  = 0x00000020,
    [helpstring("Objects have property sheets")]
    SFGAO_HASPROPSHEET               = 0x00000040,
    [helpstring("Objects are drop target")]
    SFGAO_DROPTARGET                 = 0x00000100,
    SFGAO_CAPABILITYMASK             = 0x00000177,
    SFGAO_SYSTEM                     = 0x00001000,
    [helpstring("object is encrypted (use alt color)")]
    SFGAO_ENCRYPTED                  = 0x00002000,
    [helpstring("'slow' object")]
    SFGAO_ISSLOW                     = 0x00004000,
    [helpstring("Ghosted icon")]
    SFGAO_GHOSTED                    = 0x00008000,
    [helpstring("Objects are shortcuts")]
    SFGAO_LINK                       = 0x00010000,
    [helpstring("Objects are shared")]
    SFGAO_SHARE                      = 0x00020000,
    [helpstring("Objects are read-only")]
    SFGAO_READONLY                   = 0x00040000,
    [helpstring("Objects are hidden")]
    SFGAO_HIDDEN                     = 0x00080000,
    SFGAO_DISPLAYATTRMASK            = 0x000FC000,

    [helpstring("It contains file system folder")]
    SFGAO_FILESYSANCESTOR            = 0x10000000,
    [helpstring("Is a folder")]
    SFGAO_FOLDER                     = 0x20000000,
    [helpstring("Is a file system thing (file/folder/root)")]
    SFGAO_FILESYSTEM                = 0x40000000,
    [helpstring("Expandable in the map pane")]
    SFGAO_HASSUBFOLDER              = 0x80000000,
    SFGAO_CONTENTSMASK              = 0x80000000,

    [helpstring("Invalidate cached information")]
    SFGAO_VALIDATE                  = 0x01000000,
    [helpstring("Is this removeable media?")]
    SFGAO_REMOVABLE                 = 0x02000000,
    [helpstring("Object is compressed (use alt color)")]
    SFGAO_COMPRESSED                = 0x04000000,
    [helpstring("Is in-place browsable")]
    SFGAO_BROWSABLE                 = 0x08000000,
    [helpstring("Is a non-enumerated object")]
    SFGAO_NONENUMERATED             = 0x00100000,
    [helpstring("Should show bold in explorer tree")]
    SFGAO_NEWCONTENT                = 0x00200000,

    [helpstring("Obsolete")]
    SFGAO_CANMONIKER                = 0x00400000,
    [helpstring("Obsolete")]
    SFGAO_HASSTORAGE                = 0x00400000,
    [helpstring("supports BindToObject(IID_IStream)")]
    SFGAO_STREAM                    = 0x00400000,     
    [helpstring("may contain children with SFGAO_STORAGE or SFGAO_STREAM")]
    SFGAO_STORAGEANCESTOR           = 0x00800000,
    [helpstring("for determining storage capabilities, ie for open/save semantics")]
    SFGAO_STORAGECAPMASK            = 0x70C50008,
    [helpstring("Attributes that are masked out for PKEY_SFGAOFlags because they are considered to cause slow calculations or lack context (SFGAO_VALIDATE | SFGAO_ISSLOW | SFGAO_HASSUBFOLDER and others")]
    SFGAO_PKEYSFGAOMASK             = 0x81044000,
} SFGAO_Flags;



typedef enum SIGDN { 
  SIGDN_NORMALDISPLAY                = 0x00000000,
  SIGDN_PARENTRELATIVEPARSING        = 0x80018001,
  SIGDN_DESKTOPABSOLUTEPARSING       = 0x80028000,
  SIGDN_PARENTRELATIVEEDITING        = 0x80031001,
  SIGDN_DESKTOPABSOLUTEEDITING       = 0x8004c000,
  SIGDN_FILESYSPATH                  = 0x80058000,
  SIGDN_URL                          = 0x80068000,
  SIGDN_PARENTRELATIVEFORADDRESSBAR  = 0x8007c001,
  SIGDN_PARENTRELATIVE               = 0x80080001,
  SIGDN_PARENTRELATIVEFORUI          = 0x80094001
} SIGDN;
typedef enum SysCommands
{
     SC_SIZE         = 0xF000,
     SC_MOVE         = 0xF010,
     SC_MINIMIZE     = 0xF020,
     SC_MAXIMIZE     = 0xF030,
     SC_NEXTWINDOW   = 0xF040,
     SC_PREVWINDOW   = 0xF050,
     SC_CLOSE        = 0xF060,
     SC_VSCROLL      = 0xF070,
     SC_HSCROLL      = 0xF080,
     SC_MOUSEMENU    = 0xF090,
     SC_KEYMENU      = 0xF100,
     SC_ARRANGE      = 0xF110,
     SC_RESTORE      = 0xF120,
     SC_TASKLIST     = 0xF130,
     SC_SCREENSAVE   = 0xF140,
     SC_HOTKEY       = 0xF150,
//#if(WINVER >= 0x0400) //Win95
     SC_DEFAULT      = 0xF160,
     SC_MONITORPOWER = 0xF170,
     SC_CONTEXTHELP  = 0xF180,
     SC_SEPARATOR    = 0xF00F,
//#endif /* WINVER >= 0x0400 */

//#if(WINVER >= 0x0600) //Vista
     SCF_ISSECURE    = 0x00000001,
//#endif /* WINVER >= 0x0600 */

/*
  * Obsolete names
  */
     SC_ICON     = 0xF020, //SC_MINIMIZE,
     SC_ZOOM     = 0xF030  //SC_MAXIMIZE,
} SysCommands;


typedef enum NTSTATUS {
	STATUS_SUCCESS                                    = 0x00000000,
	STATUS_WAIT_0                                     = 0x00000000,
	STATUS_WAIT_1                                     = 0x00000001,
	STATUS_WAIT_2                                     = 0x00000002,
	STATUS_WAIT_3                                     = 0x00000003,
	STATUS_WAIT_63                                    = 0x0000003F,
	STATUS_ABANDONED                                  = 0x00000080,
	STATUS_ABANDONED_WAIT_0                           = 0x00000080,
	STATUS_ABANDONED_WAIT_63                          = 0x000000BF,
	STATUS_USER_APC                                   = 0x000000C0,
	STATUS_KERNEL_APC                                 = 0x00000100,
	STATUS_ALERTED                                    = 0x00000101,
	STATUS_TIMEOUT                                    = 0x00000102,
	STATUS_PENDING                                    = 0x00000103,
	STATUS_REPARSE                                    = 0x00000104,
	STATUS_MORE_ENTRIES                               = 0x00000105,
	STATUS_NOT_ALL_ASSIGNED                           = 0x00000106,
	STATUS_SOME_NOT_MAPPED                            = 0x00000107,
	STATUS_OPLOCK_BREAK_IN_PROGRESS                   = 0x00000108,
	STATUS_VOLUME_MOUNTED                             = 0x00000109,
	STATUS_RXACT_COMMITTED                            = 0x0000010A,
	STATUS_NOTIFY_CLEANUP                             = 0x0000010B,
	STATUS_NOTIFY_ENUM_DIR                            = 0x0000010C,
	STATUS_NO_QUOTAS_FOR_ACCOUNT                      = 0x0000010D,
	STATUS_PRIMARY_TRANSPORT_CONNECT_FAILED           = 0x0000010E,
	STATUS_PAGE_FAULT_TRANSITION                      = 0x00000110,
	STATUS_PAGE_FAULT_DEMAND_ZERO                     = 0x00000111,
	STATUS_PAGE_FAULT_COPY_ON_WRITE                   = 0x00000112,
	STATUS_PAGE_FAULT_GUARD_PAGE                      = 0x00000113,
	STATUS_PAGE_FAULT_PAGING_FILE                     = 0x00000114,
	STATUS_CACHE_PAGE_LOCKED                          = 0x00000115,
	STATUS_CRASH_DUMP                                 = 0x00000116,
	STATUS_BUFFER_ALL_ZEROS                           = 0x00000117,
	STATUS_REPARSE_OBJECT                             = 0x00000118,
	STATUS_RESOURCE_REQUIREMENTS_CHANGED              = 0x00000119,
	STATUS_TRANSLATION_COMPLETE                       = 0x00000120,
	STATUS_DS_MEMBERSHIP_EVALUATED_LOCALLY            = 0x00000121,
	STATUS_NOTHING_TO_TERMINATE                       = 0x00000122,
	STATUS_PROCESS_NOT_IN_JOB                         = 0x00000123,
	STATUS_PROCESS_IN_JOB                             = 0x00000124,
	STATUS_VOLSNAP_HIBERNATE_READY                    = 0x00000125,
	STATUS_FSFILTER_OP_COMPLETED_SUCCESSFULLY         = 0x00000126,
	DBG_EXCEPTION_HANDLED                             = 0x00010001,
	DBG_CONTINUE                                      = 0x00010002,
	STATUS_OBJECT_NAME_EXISTS                         = 0x40000000,
	STATUS_THREAD_WAS_SUSPENDED                       = 0x40000001,
	STATUS_WORKING_SET_LIMIT_RANGE                    = 0x40000002,
	STATUS_IMAGE_NOT_AT_BASE                          = 0x40000003,
	STATUS_RXACT_STATE_CREATED                        = 0x40000004,
	STATUS_SEGMENT_NOTIFICATION                       = 0x40000005,
	STATUS_LOCAL_USER_SESSION_KEY                     = 0x40000006,
	STATUS_BAD_CURRENT_DIRECTORY                      = 0x40000007,
	STATUS_SERIAL_MORE_WRITES                         = 0x40000008,
	STATUS_REGISTRY_RECOVERED                         = 0x40000009,
	STATUS_FT_READ_RECOVERY_FROM_BACKUP               = 0x4000000A,
	STATUS_FT_WRITE_RECOVERY                          = 0x4000000B,
	STATUS_SERIAL_COUNTER_TIMEOUT                     = 0x4000000C,
	STATUS_NULL_LM_PASSWORD                           = 0x4000000D,
	STATUS_IMAGE_MACHINE_TYPE_MISMATCH                = 0x4000000E,
	STATUS_RECEIVE_PARTIAL                            = 0x4000000F,
	STATUS_RECEIVE_EXPEDITED                          = 0x40000010,
	STATUS_RECEIVE_PARTIAL_EXPEDITED                  = 0x40000011,
	STATUS_EVENT_DONE                                 = 0x40000012,
	STATUS_EVENT_PENDING                              = 0x40000013,
	STATUS_CHECKING_FILE_SYSTEM                       = 0x40000014,
	STATUS_FATAL_APP_EXIT                             = 0x40000015,
	STATUS_PREDEFINED_HANDLE                          = 0x40000016,
	STATUS_WAS_UNLOCKED                               = 0x40000017,
	STATUS_SERVICE_NOTIFICATION                       = 0x40000018,
	STATUS_WAS_LOCKED                                 = 0x40000019,
	STATUS_LOG_HARD_ERROR                             = 0x4000001A,
	STATUS_ALREADY_WIN32                              = 0x4000001B,
	STATUS_WX86_UNSIMULATE                            = 0x4000001C,
	STATUS_WX86_CONTINUE                              = 0x4000001D,
	STATUS_WX86_SINGLE_STEP                           = 0x4000001E,
	STATUS_WX86_BREAKPOINT                            = 0x4000001F,
	STATUS_WX86_EXCEPTION_CONTINUE                    = 0x40000020,
	STATUS_WX86_EXCEPTION_LASTCHANCE                  = 0x40000021,
	STATUS_WX86_EXCEPTION_CHAIN                       = 0x40000022,
	STATUS_IMAGE_MACHINE_TYPE_MISMATCH_EXE            = 0x40000023,
	STATUS_NO_YIELD_PERFORMED                         = 0x40000024,
	STATUS_TIMER_RESUME_IGNORED                       = 0x40000025,
	STATUS_ARBITRATION_UNHANDLED                      = 0x40000026,
	STATUS_CARDBUS_NOT_SUPPORTED                      = 0x40000027,
	STATUS_WX86_CREATEWX86TIB                         = 0x40000028,
	STATUS_MP_PROCESSOR_MISMATCH                      = 0x40000029,
	STATUS_HIBERNATED                                 = 0x4000002A,
	STATUS_RESUME_HIBERNATION                         = 0x4000002B,
	STATUS_FIRMWARE_UPDATED                           = 0x4000002C,
	STATUS_DRIVERS_LEAKING_LOCKED_PAGES               = 0x4000002D,
	DBG_REPLY_LATER                                   = 0x40010001,
	DBG_UNABLE_TO_PROVIDE_HANDLE                      = 0x40010002,
	DBG_TERMINATE_THREAD                              = 0x40010003,
	DBG_TERMINATE_PROCESS                             = 0x40010004,
	DBG_CONTROL_C                                     = 0x40010005,
	DBG_PRINTEXCEPTION_C                              = 0x40010006,
	DBG_RIPEXCEPTION                                  = 0x40010007,
	DBG_CONTROL_BREAK                                 = 0x40010008,
	DBG_COMMAND_EXCEPTION                             = 0x40010009,
	STATUS_GUARD_PAGE_VIOLATION                       = 0x80000001,
	STATUS_DATATYPE_MISALIGNMENT                      = 0x80000002,
	STATUS_BREAKPOINT                                 = 0x80000003,
	STATUS_SINGLE_STEP                                = 0x80000004,
	STATUS_BUFFER_OVERFLOW                            = 0x80000005,
	STATUS_NO_MORE_FILES                              = 0x80000006,
	STATUS_WAKE_SYSTEM_DEBUGGER                       = 0x80000007,
	STATUS_HANDLES_CLOSED                             = 0x8000000A,
	STATUS_NO_INHERITANCE                             = 0x8000000B,
	STATUS_GUID_SUBSTITUTION_MADE                     = 0x8000000C,
	STATUS_PARTIAL_COPY                               = 0x8000000D,
	STATUS_DEVICE_PAPER_EMPTY                         = 0x8000000E,
	STATUS_DEVICE_POWERED_OFF                         = 0x8000000F,
	STATUS_DEVICE_OFF_LINE                            = 0x80000010,
	STATUS_DEVICE_BUSY                                = 0x80000011,
	STATUS_NO_MORE_EAS                                = 0x80000012,
	STATUS_INVALID_EA_NAME                            = 0x80000013,
	STATUS_EA_LIST_INCONSISTENT                       = 0x80000014,
	STATUS_INVALID_EA_FLAG                            = 0x80000015,
	STATUS_VERIFY_REQUIRED                            = 0x80000016,
	STATUS_EXTRANEOUS_INFORMATION                     = 0x80000017,
	STATUS_RXACT_COMMIT_NECESSARY                     = 0x80000018,
	STATUS_NO_MORE_ENTRIES                            = 0x8000001A,
	STATUS_FILEMARK_DETECTED                          = 0x8000001B,
	STATUS_MEDIA_CHANGED                              = 0x8000001C,
	STATUS_BUS_RESET                                  = 0x8000001D,
	STATUS_END_OF_MEDIA                               = 0x8000001E,
	STATUS_BEGINNING_OF_MEDIA                         = 0x8000001F,
	STATUS_MEDIA_CHECK                                = 0x80000020,
	STATUS_SETMARK_DETECTED                           = 0x80000021,
	STATUS_NO_DATA_DETECTED                           = 0x80000022,
	STATUS_REDIRECTOR_HAS_OPEN_HANDLES                = 0x80000023,
	STATUS_SERVER_HAS_OPEN_HANDLES                    = 0x80000024,
	STATUS_ALREADY_DISCONNECTED                       = 0x80000025,
	STATUS_LONGJUMP                                   = 0x80000026,
	STATUS_CLEANER_CARTRIDGE_INSTALLED                = 0x80000027,
	STATUS_PLUGPLAY_QUERY_VETOED                      = 0x80000028,
	STATUS_UNWIND_CONSOLIDATE                         = 0x80000029,
	STATUS_REGISTRY_HIVE_RECOVERED                    = 0x8000002A,
	STATUS_DLL_MIGHT_BE_INSECURE                      = 0x8000002B,
	STATUS_DLL_MIGHT_BE_INCOMPATIBLE                  = 0x8000002C,
	DBG_EXCEPTION_NOT_HANDLED                         = 0x80010001,
	STATUS_CLUSTER_NODE_ALREADY_UP                    = 0x80130001,
	STATUS_CLUSTER_NODE_ALREADY_DOWN                  = 0x80130002,
	STATUS_CLUSTER_NETWORK_ALREADY_ONLINE             = 0x80130003,
	STATUS_CLUSTER_NETWORK_ALREADY_OFFLINE            = 0x80130004,
	STATUS_CLUSTER_NODE_ALREADY_MEMBER                = 0x80130005,
	STATUS_UNSUCCESSFUL                               = 0xC0000001,
	STATUS_NOT_IMPLEMENTED                            = 0xC0000002,
	STATUS_INVALID_INFO_CLASS                         = 0xC0000003,
	STATUS_INFO_LENGTH_MISMATCH                       = 0xC0000004,
	STATUS_ACCESS_VIOLATION                           = 0xC0000005,
	STATUS_IN_PAGE_ERROR                              = 0xC0000006,
	STATUS_PAGEFILE_QUOTA                             = 0xC0000007,
	STATUS_INVALID_HANDLE                             = 0xC0000008,
	STATUS_BAD_INITIAL_STACK                          = 0xC0000009,
	STATUS_BAD_INITIAL_PC                             = 0xC000000A,
	STATUS_INVALID_CID                                = 0xC000000B,
	STATUS_TIMER_NOT_CANCELED                         = 0xC000000C,
	STATUS_INVALID_PARAMETER                          = 0xC000000D,
	STATUS_NO_SUCH_DEVICE                             = 0xC000000E,
	STATUS_NO_SUCH_FILE                               = 0xC000000F,
	STATUS_INVALID_DEVICE_REQUEST                     = 0xC0000010,
	STATUS_END_OF_FILE                                = 0xC0000011,
	STATUS_WRONG_VOLUME                               = 0xC0000012,
	STATUS_NO_MEDIA_IN_DEVICE                         = 0xC0000013,
	STATUS_UNRECOGNIZED_MEDIA                         = 0xC0000014,
	STATUS_NONEXISTENT_SECTOR                         = 0xC0000015,
	STATUS_MORE_PROCESSING_REQUIRED                   = 0xC0000016,
	STATUS_NO_MEMORY                                  = 0xC0000017,
	STATUS_CONFLICTING_ADDRESSES                      = 0xC0000018,
	STATUS_NOT_MAPPED_VIEW                            = 0xC0000019,
	STATUS_UNABLE_TO_FREE_VM                          = 0xC000001A,
	STATUS_UNABLE_TO_DELETE_SECTION                   = 0xC000001B,
	STATUS_INVALID_SYSTEM_SERVICE                     = 0xC000001C,
	STATUS_ILLEGAL_INSTRUCTION                        = 0xC000001D,
	STATUS_INVALID_LOCK_SEQUENCE                      = 0xC000001E,
	STATUS_INVALID_VIEW_SIZE                          = 0xC000001F,
	STATUS_INVALID_FILE_FOR_SECTION                   = 0xC0000020,
	STATUS_ALREADY_COMMITTED                          = 0xC0000021,
	STATUS_ACCESS_DENIED                              = 0xC0000022,
	STATUS_BUFFER_TOO_SMALL                           = 0xC0000023,
	STATUS_OBJECT_TYPE_MISMATCH                       = 0xC0000024,
	STATUS_NONCONTINUABLE_EXCEPTION                   = 0xC0000025,
	STATUS_INVALID_DISPOSITION                        = 0xC0000026,
	STATUS_UNWIND                                     = 0xC0000027,
	STATUS_BAD_STACK                                  = 0xC0000028,
	STATUS_INVALID_UNWIND_TARGET                      = 0xC0000029,
	STATUS_NOT_LOCKED                                 = 0xC000002A,
	STATUS_PARITY_ERROR                               = 0xC000002B,
	STATUS_UNABLE_TO_DECOMMIT_VM                      = 0xC000002C,
	STATUS_NOT_COMMITTED                              = 0xC000002D,
	STATUS_INVALID_PORT_ATTRIBUTES                    = 0xC000002E,
	STATUS_PORT_MESSAGE_TOO_LONG                      = 0xC000002F,
	STATUS_INVALID_PARAMETER_MIX                      = 0xC0000030,
	STATUS_INVALID_QUOTA_LOWER                        = 0xC0000031,
	STATUS_DISK_CORRUPT_ERROR                         = 0xC0000032,
	STATUS_OBJECT_NAME_INVALID                        = 0xC0000033,
	STATUS_OBJECT_NAME_NOT_FOUND                      = 0xC0000034,
	STATUS_OBJECT_NAME_COLLISION                      = 0xC0000035,
	STATUS_PORT_DISCONNECTED                          = 0xC0000037,
	STATUS_DEVICE_ALREADY_ATTACHED                    = 0xC0000038,
	STATUS_OBJECT_PATH_INVALID                        = 0xC0000039,
	STATUS_OBJECT_PATH_NOT_FOUND                      = 0xC000003A,
	STATUS_OBJECT_PATH_SYNTAX_BAD                     = 0xC000003B,
	STATUS_DATA_OVERRUN                               = 0xC000003C,
	STATUS_DATA_LATE_ERROR                            = 0xC000003D,
	STATUS_DATA_ERROR                                 = 0xC000003E,
	STATUS_CRC_ERROR                                  = 0xC000003F,
	STATUS_SECTION_TOO_BIG                            = 0xC0000040,
	STATUS_PORT_CONNECTION_REFUSED                    = 0xC0000041,
	STATUS_INVALID_PORT_HANDLE                        = 0xC0000042,
	STATUS_SHARING_VIOLATION                          = 0xC0000043,
	STATUS_QUOTA_EXCEEDED                             = 0xC0000044,
	STATUS_INVALID_PAGE_PROTECTION                    = 0xC0000045,
	STATUS_MUTANT_NOT_OWNED                           = 0xC0000046,
	STATUS_SEMAPHORE_LIMIT_EXCEEDED                   = 0xC0000047,
	STATUS_PORT_ALREADY_SET                           = 0xC0000048,
	STATUS_SECTION_NOT_IMAGE                          = 0xC0000049,
	STATUS_SUSPEND_COUNT_EXCEEDED                     = 0xC000004A,
	STATUS_THREAD_IS_TERMINATING                      = 0xC000004B,
	STATUS_BAD_WORKING_SET_LIMIT                      = 0xC000004C,
	STATUS_INCOMPATIBLE_FILE_MAP                      = 0xC000004D,
	STATUS_SECTION_PROTECTION                         = 0xC000004E,
	STATUS_EAS_NOT_SUPPORTED                          = 0xC000004F,
	STATUS_EA_TOO_LARGE                               = 0xC0000050,
	STATUS_NONEXISTENT_EA_ENTRY                       = 0xC0000051,
	STATUS_NO_EAS_ON_FILE                             = 0xC0000052,
	STATUS_EA_CORRUPT_ERROR                           = 0xC0000053,
	STATUS_FILE_LOCK_CONFLICT                         = 0xC0000054,
	STATUS_LOCK_NOT_GRANTED                           = 0xC0000055,
	STATUS_DELETE_PENDING                             = 0xC0000056,
	STATUS_CTL_FILE_NOT_SUPPORTED                     = 0xC0000057,
	STATUS_UNKNOWN_REVISION                           = 0xC0000058,
	STATUS_REVISION_MISMATCH                          = 0xC0000059,
	STATUS_INVALID_OWNER                              = 0xC000005A,
	STATUS_INVALID_PRIMARY_GROUP                      = 0xC000005B,
	STATUS_NO_IMPERSONATION_TOKEN                     = 0xC000005C,
	STATUS_CANT_DISABLE_MANDATORY                     = 0xC000005D,
	STATUS_NO_LOGON_SERVERS                           = 0xC000005E,
	STATUS_NO_SUCH_LOGON_SESSION                      = 0xC000005F,
	STATUS_NO_SUCH_PRIVILEGE                          = 0xC0000060,
	STATUS_PRIVILEGE_NOT_HELD                         = 0xC0000061,
	STATUS_INVALID_ACCOUNT_NAME                       = 0xC0000062,
	STATUS_USER_EXISTS                                = 0xC0000063,
	STATUS_NO_SUCH_USER                               = 0xC0000064,
	STATUS_GROUP_EXISTS                               = 0xC0000065,
	STATUS_NO_SUCH_GROUP                              = 0xC0000066,
	STATUS_MEMBER_IN_GROUP                            = 0xC0000067,
	STATUS_MEMBER_NOT_IN_GROUP                        = 0xC0000068,
	STATUS_LAST_ADMIN                                 = 0xC0000069,
	STATUS_WRONG_PASSWORD                             = 0xC000006A,
	STATUS_ILL_FORMED_PASSWORD                        = 0xC000006B,
	STATUS_PASSWORD_RESTRICTION                       = 0xC000006C,
	STATUS_LOGON_FAILURE                              = 0xC000006D,
	STATUS_ACCOUNT_RESTRICTION                        = 0xC000006E,
	STATUS_INVALID_LOGON_HOURS                        = 0xC000006F,
	STATUS_INVALID_WORKSTATION                        = 0xC0000070,
	STATUS_PASSWORD_EXPIRED                           = 0xC0000071,
	STATUS_ACCOUNT_DISABLED                           = 0xC0000072,
	STATUS_NONE_MAPPED                                = 0xC0000073,
	STATUS_TOO_MANY_LUIDS_REQUESTED                   = 0xC0000074,
	STATUS_LUIDS_EXHAUSTED                            = 0xC0000075,
	STATUS_INVALID_SUB_AUTHORITY                      = 0xC0000076,
	STATUS_INVALID_ACL                                = 0xC0000077,
	STATUS_INVALID_SID                                = 0xC0000078,
	STATUS_INVALID_SECURITY_DESCR                     = 0xC0000079,
	STATUS_PROCEDURE_NOT_FOUND                        = 0xC000007A,
	STATUS_INVALID_IMAGE_FORMAT                       = 0xC000007B,
	STATUS_NO_TOKEN                                   = 0xC000007C,
	STATUS_BAD_INHERITANCE_ACL                        = 0xC000007D,
	STATUS_RANGE_NOT_LOCKED                           = 0xC000007E,
	STATUS_DISK_FULL                                  = 0xC000007F,
	STATUS_SERVER_DISABLED                            = 0xC0000080,
	STATUS_SERVER_NOT_DISABLED                        = 0xC0000081,
	STATUS_TOO_MANY_GUIDS_REQUESTED                   = 0xC0000082,
	STATUS_GUIDS_EXHAUSTED                            = 0xC0000083,
	STATUS_INVALID_ID_AUTHORITY                       = 0xC0000084,
	STATUS_AGENTS_EXHAUSTED                           = 0xC0000085,
	STATUS_INVALID_VOLUME_LABEL                       = 0xC0000086,
	STATUS_SECTION_NOT_EXTENDED                       = 0xC0000087,
	STATUS_NOT_MAPPED_DATA                            = 0xC0000088,
	STATUS_RESOURCE_DATA_NOT_FOUND                    = 0xC0000089,
	STATUS_RESOURCE_TYPE_NOT_FOUND                    = 0xC000008A,
	STATUS_RESOURCE_NAME_NOT_FOUND                    = 0xC000008B,
	STATUS_ARRAY_BOUNDS_EXCEEDED                      = 0xC000008C,
	STATUS_FLOAT_DENORMAL_OPERAND                     = 0xC000008D,
	STATUS_FLOAT_DIVIDE_BY_ZERO                       = 0xC000008E,
	STATUS_FLOAT_INEXACT_RESULT                       = 0xC000008F,
	STATUS_FLOAT_INVALID_OPERATION                    = 0xC0000090,
	STATUS_FLOAT_OVERFLOW                             = 0xC0000091,
	STATUS_FLOAT_STACK_CHECK                          = 0xC0000092,
	STATUS_FLOAT_UNDERFLOW                            = 0xC0000093,
	STATUS_INTEGER_DIVIDE_BY_ZERO                     = 0xC0000094,
	STATUS_INTEGER_OVERFLOW                           = 0xC0000095,
	STATUS_PRIVILEGED_INSTRUCTION                     = 0xC0000096,
	STATUS_TOO_MANY_PAGING_FILES                      = 0xC0000097,
	STATUS_FILE_INVALID                               = 0xC0000098,
	STATUS_ALLOTTED_SPACE_EXCEEDED                    = 0xC0000099,
	STATUS_INSUFFICIENT_RESOURCES                     = 0xC000009A,
	STATUS_DFS_EXIT_PATH_FOUND                        = 0xC000009B,
	STATUS_DEVICE_DATA_ERROR                          = 0xC000009C,
	STATUS_DEVICE_NOT_CONNECTED                       = 0xC000009D,
	STATUS_DEVICE_POWER_FAILURE                       = 0xC000009E,
	STATUS_FREE_VM_NOT_AT_BASE                        = 0xC000009F,
	STATUS_MEMORY_NOT_ALLOCATED                       = 0xC00000A0,
	STATUS_WORKING_SET_QUOTA                          = 0xC00000A1,
	STATUS_MEDIA_WRITE_PROTECTED                      = 0xC00000A2,
	STATUS_DEVICE_NOT_READY                           = 0xC00000A3,
	STATUS_INVALID_GROUP_ATTRIBUTES                   = 0xC00000A4,
	STATUS_BAD_IMPERSONATION_LEVEL                    = 0xC00000A5,
	STATUS_CANT_OPEN_ANONYMOUS                        = 0xC00000A6,
	STATUS_BAD_VALIDATION_CLASS                       = 0xC00000A7,
	STATUS_BAD_TOKEN_TYPE                             = 0xC00000A8,
	STATUS_BAD_MASTER_BOOT_RECORD                     = 0xC00000A9,
	STATUS_INSTRUCTION_MISALIGNMENT                   = 0xC00000AA,
	STATUS_INSTANCE_NOT_AVAILABLE                     = 0xC00000AB,
	STATUS_PIPE_NOT_AVAILABLE                         = 0xC00000AC,
	STATUS_INVALID_PIPE_STATE                         = 0xC00000AD,
	STATUS_PIPE_BUSY                                  = 0xC00000AE,
	STATUS_ILLEGAL_FUNCTION                           = 0xC00000AF,
	STATUS_PIPE_DISCONNECTED                          = 0xC00000B0,
	STATUS_PIPE_CLOSING                               = 0xC00000B1,
	STATUS_PIPE_CONNECTED                             = 0xC00000B2,
	STATUS_PIPE_LISTENING                             = 0xC00000B3,
	STATUS_INVALID_READ_MODE                          = 0xC00000B4,
	STATUS_IO_TIMEOUT                                 = 0xC00000B5,
	STATUS_FILE_FORCED_CLOSED                         = 0xC00000B6,
	STATUS_PROFILING_NOT_STARTED                      = 0xC00000B7,
	STATUS_PROFILING_NOT_STOPPED                      = 0xC00000B8,
	STATUS_COULD_NOT_INTERPRET                        = 0xC00000B9,
	STATUS_FILE_IS_A_DIRECTORY                        = 0xC00000BA,
	STATUS_NOT_SUPPORTED                              = 0xC00000BB,
	STATUS_REMOTE_NOT_LISTENING                       = 0xC00000BC,
	STATUS_DUPLICATE_NAME                             = 0xC00000BD,
	STATUS_BAD_NETWORK_PATH                           = 0xC00000BE,
	STATUS_NETWORK_BUSY                               = 0xC00000BF,
	STATUS_DEVICE_DOES_NOT_EXIST                      = 0xC00000C0,
	STATUS_TOO_MANY_COMMANDS                          = 0xC00000C1,
	STATUS_ADAPTER_HARDWARE_ERROR                     = 0xC00000C2,
	STATUS_INVALID_NETWORK_RESPONSE                   = 0xC00000C3,
	STATUS_UNEXPECTED_NETWORK_ERROR                   = 0xC00000C4,
	STATUS_BAD_REMOTE_ADAPTER                         = 0xC00000C5,
	STATUS_PRINT_QUEUE_FULL                           = 0xC00000C6,
	STATUS_NO_SPOOL_SPACE                             = 0xC00000C7,
	STATUS_PRINT_CANCELLED                            = 0xC00000C8,
	STATUS_NETWORK_NAME_DELETED                       = 0xC00000C9,
	STATUS_NETWORK_ACCESS_DENIED                      = 0xC00000CA,
	STATUS_BAD_DEVICE_TYPE                            = 0xC00000CB,
	STATUS_BAD_NETWORK_NAME                           = 0xC00000CC,
	STATUS_TOO_MANY_NAMES                             = 0xC00000CD,
	STATUS_TOO_MANY_SESSIONS                          = 0xC00000CE,
	STATUS_SHARING_PAUSED                             = 0xC00000CF,
	STATUS_REQUEST_NOT_ACCEPTED                       = 0xC00000D0,
	STATUS_REDIRECTOR_PAUSED                          = 0xC00000D1,
	STATUS_NET_WRITE_FAULT                            = 0xC00000D2,
	STATUS_PROFILING_AT_LIMIT                         = 0xC00000D3,
	STATUS_NOT_SAME_DEVICE                            = 0xC00000D4,
	STATUS_FILE_RENAMED                               = 0xC00000D5,
	STATUS_VIRTUAL_CIRCUIT_CLOSED                     = 0xC00000D6,
	STATUS_NO_SECURITY_ON_OBJECT                      = 0xC00000D7,
	STATUS_CANT_WAIT                                  = 0xC00000D8,
	STATUS_PIPE_EMPTY                                 = 0xC00000D9,
	STATUS_CANT_ACCESS_DOMAIN_INFO                    = 0xC00000DA,
	STATUS_CANT_TERMINATE_SELF                        = 0xC00000DB,
	STATUS_INVALID_SERVER_STATE                       = 0xC00000DC,
	STATUS_INVALID_DOMAIN_STATE                       = 0xC00000DD,
	STATUS_INVALID_DOMAIN_ROLE                        = 0xC00000DE,
	STATUS_NO_SUCH_DOMAIN                             = 0xC00000DF,
	STATUS_DOMAIN_EXISTS                              = 0xC00000E0,
	STATUS_DOMAIN_LIMIT_EXCEEDED                      = 0xC00000E1,
	STATUS_OPLOCK_NOT_GRANTED                         = 0xC00000E2,
	STATUS_INVALID_OPLOCK_PROTOCOL                    = 0xC00000E3,
	STATUS_INTERNAL_DB_CORRUPTION                     = 0xC00000E4,
	STATUS_INTERNAL_ERROR                             = 0xC00000E5,
	STATUS_GENERIC_NOT_MAPPED                         = 0xC00000E6,
	STATUS_BAD_DESCRIPTOR_FORMAT                      = 0xC00000E7,
	STATUS_INVALID_USER_BUFFER                        = 0xC00000E8,
	STATUS_UNEXPECTED_IO_ERROR                        = 0xC00000E9,
	STATUS_UNEXPECTED_MM_CREATE_ERR                   = 0xC00000EA,
	STATUS_UNEXPECTED_MM_MAP_ERROR                    = 0xC00000EB,
	STATUS_UNEXPECTED_MM_EXTEND_ERR                   = 0xC00000EC,
	STATUS_NOT_LOGON_PROCESS                          = 0xC00000ED,
	STATUS_LOGON_SESSION_EXISTS                       = 0xC00000EE,
	STATUS_INVALID_PARAMETER_1                        = 0xC00000EF,
	STATUS_INVALID_PARAMETER_2                        = 0xC00000F0,
	STATUS_INVALID_PARAMETER_3                        = 0xC00000F1,
	STATUS_INVALID_PARAMETER_4                        = 0xC00000F2,
	STATUS_INVALID_PARAMETER_5                        = 0xC00000F3,
	STATUS_INVALID_PARAMETER_6                        = 0xC00000F4,
	STATUS_INVALID_PARAMETER_7                        = 0xC00000F5,
	STATUS_INVALID_PARAMETER_8                        = 0xC00000F6,
	STATUS_INVALID_PARAMETER_9                        = 0xC00000F7,
	STATUS_INVALID_PARAMETER_10                       = 0xC00000F8,
	STATUS_INVALID_PARAMETER_11                       = 0xC00000F9,
	STATUS_INVALID_PARAMETER_12                       = 0xC00000FA,
	STATUS_REDIRECTOR_NOT_STARTED                     = 0xC00000FB,
	STATUS_REDIRECTOR_STARTED                         = 0xC00000FC,
	STATUS_STACK_OVERFLOW                             = 0xC00000FD,
	STATUS_NO_SUCH_PACKAGE                            = 0xC00000FE,
	STATUS_BAD_FUNCTION_TABLE                         = 0xC00000FF,
	STATUS_VARIABLE_NOT_FOUND                         = 0xC0000100,
	STATUS_DIRECTORY_NOT_EMPTY                        = 0xC0000101,
	STATUS_FILE_CORRUPT_ERROR                         = 0xC0000102,
	STATUS_NOT_A_DIRECTORY                            = 0xC0000103,
	STATUS_BAD_LOGON_SESSION_STATE                    = 0xC0000104,
	STATUS_LOGON_SESSION_COLLISION                    = 0xC0000105,
	STATUS_NAME_TOO_LONG                              = 0xC0000106,
	STATUS_FILES_OPEN                                 = 0xC0000107,
	STATUS_CONNECTION_IN_USE                          = 0xC0000108,
	STATUS_MESSAGE_NOT_FOUND                          = 0xC0000109,
	STATUS_PROCESS_IS_TERMINATING                     = 0xC000010A,
	STATUS_INVALID_LOGON_TYPE                         = 0xC000010B,
	STATUS_NO_GUID_TRANSLATION                        = 0xC000010C,
	STATUS_CANNOT_IMPERSONATE                         = 0xC000010D,
	STATUS_IMAGE_ALREADY_LOADED                       = 0xC000010E,
	STATUS_ABIOS_NOT_PRESENT                          = 0xC000010F,
	STATUS_ABIOS_LID_NOT_EXIST                        = 0xC0000110,
	STATUS_ABIOS_LID_ALREADY_OWNED                    = 0xC0000111,
	STATUS_ABIOS_NOT_LID_OWNER                        = 0xC0000112,
	STATUS_ABIOS_INVALID_COMMAND                      = 0xC0000113,
	STATUS_ABIOS_INVALID_LID                          = 0xC0000114,
	STATUS_ABIOS_SELECTOR_NOT_AVAILABLE               = 0xC0000115,
	STATUS_ABIOS_INVALID_SELECTOR                     = 0xC0000116,
	STATUS_NO_LDT                                     = 0xC0000117,
	STATUS_INVALID_LDT_SIZE                           = 0xC0000118,
	STATUS_INVALID_LDT_OFFSET                         = 0xC0000119,
	STATUS_INVALID_LDT_DESCRIPTOR                     = 0xC000011A,
	STATUS_INVALID_IMAGE_NE_FORMAT                    = 0xC000011B,
	STATUS_RXACT_INVALID_STATE                        = 0xC000011C,
	STATUS_RXACT_COMMIT_FAILURE                       = 0xC000011D,
	STATUS_MAPPED_FILE_SIZE_ZERO                      = 0xC000011E,
	STATUS_TOO_MANY_OPENED_FILES                      = 0xC000011F,
	STATUS_CANCELLED                                  = 0xC0000120,
	STATUS_CANNOT_DELETE                              = 0xC0000121,
	STATUS_INVALID_COMPUTER_NAME                      = 0xC0000122,
	STATUS_FILE_DELETED                               = 0xC0000123,
	STATUS_SPECIAL_ACCOUNT                            = 0xC0000124,
	STATUS_SPECIAL_GROUP                              = 0xC0000125,
	STATUS_SPECIAL_USER                               = 0xC0000126,
	STATUS_MEMBERS_PRIMARY_GROUP                      = 0xC0000127,
	STATUS_FILE_CLOSED                                = 0xC0000128,
	STATUS_TOO_MANY_THREADS                           = 0xC0000129,
	STATUS_THREAD_NOT_IN_PROCESS                      = 0xC000012A,
	STATUS_TOKEN_ALREADY_IN_USE                       = 0xC000012B,
	STATUS_PAGEFILE_QUOTA_EXCEEDED                    = 0xC000012C,
	STATUS_COMMITMENT_LIMIT                           = 0xC000012D,
	STATUS_INVALID_IMAGE_LE_FORMAT                    = 0xC000012E,
	STATUS_INVALID_IMAGE_NOT_MZ                       = 0xC000012F,
	STATUS_INVALID_IMAGE_PROTECT                      = 0xC0000130,
	STATUS_INVALID_IMAGE_WIN_16                       = 0xC0000131,
	STATUS_LOGON_SERVER_CONFLICT                      = 0xC0000132,
	STATUS_TIME_DIFFERENCE_AT_DC                      = 0xC0000133,
	STATUS_SYNCHRONIZATION_REQUIRED                   = 0xC0000134,
	STATUS_DLL_NOT_FOUND                              = 0xC0000135,
	STATUS_OPEN_FAILED                                = 0xC0000136,
	STATUS_IO_PRIVILEGE_FAILED                        = 0xC0000137,
	STATUS_ORDINAL_NOT_FOUND                          = 0xC0000138,
	STATUS_ENTRYPOINT_NOT_FOUND                       = 0xC0000139,
	STATUS_CONTROL_C_EXIT                             = 0xC000013A,
	STATUS_LOCAL_DISCONNECT                           = 0xC000013B,
	STATUS_REMOTE_DISCONNECT                          = 0xC000013C,
	STATUS_REMOTE_RESOURCES                           = 0xC000013D,
	STATUS_LINK_FAILED                                = 0xC000013E,
	STATUS_LINK_TIMEOUT                               = 0xC000013F,
	STATUS_INVALID_CONNECTION                         = 0xC0000140,
	STATUS_INVALID_ADDRESS                            = 0xC0000141,
	STATUS_DLL_INIT_FAILED                            = 0xC0000142,
	STATUS_MISSING_SYSTEMFILE                         = 0xC0000143,
	STATUS_UNHANDLED_EXCEPTION                        = 0xC0000144,
	STATUS_APP_INIT_FAILURE                           = 0xC0000145,
	STATUS_PAGEFILE_CREATE_FAILED                     = 0xC0000146,
	STATUS_NO_PAGEFILE                                = 0xC0000147,
	STATUS_INVALID_LEVEL                              = 0xC0000148,
	STATUS_WRONG_PASSWORD_CORE                        = 0xC0000149,
	STATUS_ILLEGAL_FLOAT_CONTEXT                      = 0xC000014A,
	STATUS_PIPE_BROKEN                                = 0xC000014B,
	STATUS_REGISTRY_CORRUPT                           = 0xC000014C,
	STATUS_REGISTRY_IO_FAILED                         = 0xC000014D,
	STATUS_NO_EVENT_PAIR                              = 0xC000014E,
	STATUS_UNRECOGNIZED_VOLUME                        = 0xC000014F,
	STATUS_SERIAL_NO_DEVICE_INITED                    = 0xC0000150,
	STATUS_NO_SUCH_ALIAS                              = 0xC0000151,
	STATUS_MEMBER_NOT_IN_ALIAS                        = 0xC0000152,
	STATUS_MEMBER_IN_ALIAS                            = 0xC0000153,
	STATUS_ALIAS_EXISTS                               = 0xC0000154,
	STATUS_LOGON_NOT_GRANTED                          = 0xC0000155,
	STATUS_TOO_MANY_SECRETS                           = 0xC0000156,
	STATUS_SECRET_TOO_LONG                            = 0xC0000157,
	STATUS_INTERNAL_DB_ERROR                          = 0xC0000158,
	STATUS_FULLSCREEN_MODE                            = 0xC0000159,
	STATUS_TOO_MANY_CONTEXT_IDS                       = 0xC000015A,
	STATUS_LOGON_TYPE_NOT_GRANTED                     = 0xC000015B,
	STATUS_NOT_REGISTRY_FILE                          = 0xC000015C,
	STATUS_NT_CROSS_ENCRYPTION_REQUIRED               = 0xC000015D,
	STATUS_DOMAIN_CTRLR_CONFIG_ERROR                  = 0xC000015E,
	STATUS_FT_MISSING_MEMBER                          = 0xC000015F,
	STATUS_ILL_FORMED_SERVICE_ENTRY                   = 0xC0000160,
	STATUS_ILLEGAL_CHARACTER                          = 0xC0000161,
	STATUS_UNMAPPABLE_CHARACTER                       = 0xC0000162,
	STATUS_UNDEFINED_CHARACTER                        = 0xC0000163,
	STATUS_FLOPPY_VOLUME                              = 0xC0000164,
	STATUS_FLOPPY_ID_MARK_NOT_FOUND                   = 0xC0000165,
	STATUS_FLOPPY_WRONG_CYLINDER                      = 0xC0000166,
	STATUS_FLOPPY_UNKNOWN_ERROR                       = 0xC0000167,
	STATUS_FLOPPY_BAD_REGISTERS                       = 0xC0000168,
	STATUS_DISK_RECALIBRATE_FAILED                    = 0xC0000169,
	STATUS_DISK_OPERATION_FAILED                      = 0xC000016A,
	STATUS_DISK_RESET_FAILED                          = 0xC000016B,
	STATUS_SHARED_IRQ_BUSY                            = 0xC000016C,
	STATUS_FT_ORPHANING                               = 0xC000016D,
	STATUS_BIOS_FAILED_TO_CONNECT_INTERRUPT           = 0xC000016E,
	STATUS_PARTITION_FAILURE                          = 0xC0000172,
	STATUS_INVALID_BLOCK_LENGTH                       = 0xC0000173,
	STATUS_DEVICE_NOT_PARTITIONED                     = 0xC0000174,
	STATUS_UNABLE_TO_LOCK_MEDIA                       = 0xC0000175,
	STATUS_UNABLE_TO_UNLOAD_MEDIA                     = 0xC0000176,
	STATUS_EOM_OVERFLOW                               = 0xC0000177,
	STATUS_NO_MEDIA                                   = 0xC0000178,
	STATUS_NO_SUCH_MEMBER                             = 0xC000017A,
	STATUS_INVALID_MEMBER                             = 0xC000017B,
	STATUS_KEY_DELETED                                = 0xC000017C,
	STATUS_NO_LOG_SPACE                               = 0xC000017D,
	STATUS_TOO_MANY_SIDS                              = 0xC000017E,
	STATUS_LM_CROSS_ENCRYPTION_REQUIRED               = 0xC000017F,
	STATUS_KEY_HAS_CHILDREN                           = 0xC0000180,
	STATUS_CHILD_MUST_BE_VOLATILE                     = 0xC0000181,
	STATUS_DEVICE_CONFIGURATION_ERROR                 = 0xC0000182,
	STATUS_DRIVER_INTERNAL_ERROR                      = 0xC0000183,
	STATUS_INVALID_DEVICE_STATE                       = 0xC0000184,
	STATUS_IO_DEVICE_ERROR                            = 0xC0000185,
	STATUS_DEVICE_PROTOCOL_ERROR                      = 0xC0000186,
	STATUS_BACKUP_CONTROLLER                          = 0xC0000187,
	STATUS_LOG_FILE_FULL                              = 0xC0000188,
	STATUS_TOO_LATE                                   = 0xC0000189,
	STATUS_NO_TRUST_LSA_SECRET                        = 0xC000018A,
	STATUS_NO_TRUST_SAM_ACCOUNT                       = 0xC000018B,
	STATUS_TRUSTED_DOMAIN_FAILURE                     = 0xC000018C,
	STATUS_TRUSTED_RELATIONSHIP_FAILURE               = 0xC000018D,
	STATUS_EVENTLOG_FILE_CORRUPT                      = 0xC000018E,
	STATUS_EVENTLOG_CANT_START                        = 0xC000018F,
	STATUS_TRUST_FAILURE                              = 0xC0000190,
	STATUS_MUTANT_LIMIT_EXCEEDED                      = 0xC0000191,
	STATUS_NETLOGON_NOT_STARTED                       = 0xC0000192,
	STATUS_ACCOUNT_EXPIRED                            = 0xC0000193,
	STATUS_POSSIBLE_DEADLOCK                          = 0xC0000194,
	STATUS_NETWORK_CREDENTIAL_CONFLICT                = 0xC0000195,
	STATUS_REMOTE_SESSION_LIMIT                       = 0xC0000196,
	STATUS_EVENTLOG_FILE_CHANGED                      = 0xC0000197,
	STATUS_NOLOGON_INTERDOMAIN_TRUST_ACCOUNT          = 0xC0000198,
	STATUS_NOLOGON_WORKSTATION_TRUST_ACCOUNT          = 0xC0000199,
	STATUS_NOLOGON_SERVER_TRUST_ACCOUNT               = 0xC000019A,
	STATUS_DOMAIN_TRUST_INCONSISTENT                  = 0xC000019B,
	STATUS_FS_DRIVER_REQUIRED                         = 0xC000019C,
	STATUS_NO_USER_SESSION_KEY                        = 0xC0000202,
	STATUS_USER_SESSION_DELETED                       = 0xC0000203,
	STATUS_RESOURCE_LANG_NOT_FOUND                    = 0xC0000204,
	STATUS_INSUFF_SERVER_RESOURCES                    = 0xC0000205,
	STATUS_INVALID_BUFFER_SIZE                        = 0xC0000206,
	STATUS_INVALID_ADDRESS_COMPONENT                  = 0xC0000207,
	STATUS_INVALID_ADDRESS_WILDCARD                   = 0xC0000208,
	STATUS_TOO_MANY_ADDRESSES                         = 0xC0000209,
	STATUS_ADDRESS_ALREADY_EXISTS                     = 0xC000020A,
	STATUS_ADDRESS_CLOSED                             = 0xC000020B,
	STATUS_CONNECTION_DISCONNECTED                    = 0xC000020C,
	STATUS_CONNECTION_RESET                           = 0xC000020D,
	STATUS_TOO_MANY_NODES                             = 0xC000020E,
	STATUS_TRANSACTION_ABORTED                        = 0xC000020F,
	STATUS_TRANSACTION_TIMED_OUT                      = 0xC0000210,
	STATUS_TRANSACTION_NO_RELEASE                     = 0xC0000211,
	STATUS_TRANSACTION_NO_MATCH                       = 0xC0000212,
	STATUS_TRANSACTION_RESPONDED                      = 0xC0000213,
	STATUS_TRANSACTION_INVALID_ID                     = 0xC0000214,
	STATUS_TRANSACTION_INVALID_TYPE                   = 0xC0000215,
	STATUS_NOT_SERVER_SESSION                         = 0xC0000216,
	STATUS_NOT_CLIENT_SESSION                         = 0xC0000217,
	STATUS_CANNOT_LOAD_REGISTRY_FILE                  = 0xC0000218,
	STATUS_DEBUG_ATTACH_FAILED                        = 0xC0000219,
	STATUS_SYSTEM_PROCESS_TERMINATED                  = 0xC000021A,
	STATUS_DATA_NOT_ACCEPTED                          = 0xC000021B,
	STATUS_NO_BROWSER_SERVERS_FOUND                   = 0xC000021C,
	STATUS_VDM_HARD_ERROR                             = 0xC000021D,
	STATUS_DRIVER_CANCEL_TIMEOUT                      = 0xC000021E,
	STATUS_REPLY_MESSAGE_MISMATCH                     = 0xC000021F,
	STATUS_MAPPED_ALIGNMENT                           = 0xC0000220,
	STATUS_IMAGE_CHECKSUM_MISMATCH                    = 0xC0000221,
	STATUS_LOST_WRITEBEHIND_DATA                      = 0xC0000222,
	STATUS_CLIENT_SERVER_PARAMETERS_INVALID           = 0xC0000223,
	STATUS_PASSWORD_MUST_CHANGE                       = 0xC0000224,
	STATUS_NOT_FOUND                                  = 0xC0000225,
	STATUS_NOT_TINY_STREAM                            = 0xC0000226,
	STATUS_RECOVERY_FAILURE                           = 0xC0000227,
	STATUS_STACK_OVERFLOW_READ                        = 0xC0000228,
	STATUS_FAIL_CHECK                                 = 0xC0000229,
	STATUS_DUPLICATE_OBJECTID                         = 0xC000022A,
	STATUS_OBJECTID_EXISTS                            = 0xC000022B,
	STATUS_CONVERT_TO_LARGE                           = 0xC000022C,
	STATUS_RETRY                                      = 0xC000022D,
	STATUS_FOUND_OUT_OF_SCOPE                         = 0xC000022E,
	STATUS_ALLOCATE_BUCKET                            = 0xC000022F,
	STATUS_PROPSET_NOT_FOUND                          = 0xC0000230,
	STATUS_MARSHALL_OVERFLOW                          = 0xC0000231,
	STATUS_INVALID_VARIANT                            = 0xC0000232,
	STATUS_DOMAIN_CONTROLLER_NOT_FOUND                = 0xC0000233,
	STATUS_ACCOUNT_LOCKED_OUT                         = 0xC0000234,
	STATUS_HANDLE_NOT_CLOSABLE                        = 0xC0000235,
	STATUS_CONNECTION_REFUSED                         = 0xC0000236,
	STATUS_GRACEFUL_DISCONNECT                        = 0xC0000237,
	STATUS_ADDRESS_ALREADY_ASSOCIATED                 = 0xC0000238,
	STATUS_ADDRESS_NOT_ASSOCIATED                     = 0xC0000239,
	STATUS_CONNECTION_INVALID                         = 0xC000023A,
	STATUS_CONNECTION_ACTIVE                          = 0xC000023B,
	STATUS_NETWORK_UNREACHABLE                        = 0xC000023C,
	STATUS_HOST_UNREACHABLE                           = 0xC000023D,
	STATUS_PROTOCOL_UNREACHABLE                       = 0xC000023E,
	STATUS_PORT_UNREACHABLE                           = 0xC000023F,
	STATUS_REQUEST_ABORTED                            = 0xC0000240,
	STATUS_CONNECTION_ABORTED                         = 0xC0000241,
	STATUS_BAD_COMPRESSION_BUFFER                     = 0xC0000242,
	STATUS_USER_MAPPED_FILE                           = 0xC0000243,
	STATUS_AUDIT_FAILED                               = 0xC0000244,
	STATUS_TIMER_RESOLUTION_NOT_SET                   = 0xC0000245,
	STATUS_CONNECTION_COUNT_LIMIT                     = 0xC0000246,
	STATUS_LOGIN_TIME_RESTRICTION                     = 0xC0000247,
	STATUS_LOGIN_WKSTA_RESTRICTION                    = 0xC0000248,
	STATUS_IMAGE_MP_UP_MISMATCH                       = 0xC0000249,
	STATUS_INSUFFICIENT_LOGON_INFO                    = 0xC0000250,
	STATUS_BAD_DLL_ENTRYPOINT                         = 0xC0000251,
	STATUS_BAD_SERVICE_ENTRYPOINT                     = 0xC0000252,
	STATUS_LPC_REPLY_LOST                             = 0xC0000253,
	STATUS_IP_ADDRESS_CONFLICT1                       = 0xC0000254,
	STATUS_IP_ADDRESS_CONFLICT2                       = 0xC0000255,
	STATUS_REGISTRY_QUOTA_LIMIT                       = 0xC0000256,
	STATUS_PATH_NOT_COVERED                           = 0xC0000257,
	STATUS_NO_CALLBACK_ACTIVE                         = 0xC0000258,
	STATUS_LICENSE_QUOTA_EXCEEDED                     = 0xC0000259,
	STATUS_PWD_TOO_SHORT                              = 0xC000025A,
	STATUS_PWD_TOO_RECENT                             = 0xC000025B,
	STATUS_PWD_HISTORY_CONFLICT                       = 0xC000025C,
	STATUS_PLUGPLAY_NO_DEVICE                         = 0xC000025E,
	STATUS_UNSUPPORTED_COMPRESSION                    = 0xC000025F,
	STATUS_INVALID_HW_PROFILE                         = 0xC0000260,
	STATUS_INVALID_PLUGPLAY_DEVICE_PATH               = 0xC0000261,
	STATUS_DRIVER_ORDINAL_NOT_FOUND                   = 0xC0000262,
	STATUS_DRIVER_ENTRYPOINT_NOT_FOUND                = 0xC0000263,
	STATUS_RESOURCE_NOT_OWNED                         = 0xC0000264,
	STATUS_TOO_MANY_LINKS                             = 0xC0000265,
	STATUS_QUOTA_LIST_INCONSISTENT                    = 0xC0000266,
	STATUS_FILE_IS_OFFLINE                            = 0xC0000267,
	STATUS_EVALUATION_EXPIRATION                      = 0xC0000268,
	STATUS_ILLEGAL_DLL_RELOCATION                     = 0xC0000269,
	STATUS_LICENSE_VIOLATION                          = 0xC000026A,
	STATUS_DLL_INIT_FAILED_LOGOFF                     = 0xC000026B,
	STATUS_DRIVER_UNABLE_TO_LOAD                      = 0xC000026C,
	STATUS_DFS_UNAVAILABLE                            = 0xC000026D,
	STATUS_VOLUME_DISMOUNTED                          = 0xC000026E,
	STATUS_WX86_INTERNAL_ERROR                        = 0xC000026F,
	STATUS_WX86_FLOAT_STACK_CHECK                     = 0xC0000270,
	STATUS_VALIDATE_CONTINUE                          = 0xC0000271,
	STATUS_NO_MATCH                                   = 0xC0000272,
	STATUS_NO_MORE_MATCHES                            = 0xC0000273,
	STATUS_NOT_A_REPARSE_POINT                        = 0xC0000275,
	STATUS_IO_REPARSE_TAG_INVALID                     = 0xC0000276,
	STATUS_IO_REPARSE_TAG_MISMATCH                    = 0xC0000277,
	STATUS_IO_REPARSE_DATA_INVALID                    = 0xC0000278,
	STATUS_IO_REPARSE_TAG_NOT_HANDLED                 = 0xC0000279,
	STATUS_REPARSE_POINT_NOT_RESOLVED                 = 0xC0000280,
	STATUS_DIRECTORY_IS_A_REPARSE_POINT               = 0xC0000281,
	STATUS_RANGE_LIST_CONFLICT                        = 0xC0000282,
	STATUS_SOURCE_ELEMENT_EMPTY                       = 0xC0000283,
	STATUS_DESTINATION_ELEMENT_FULL                   = 0xC0000284,
	STATUS_ILLEGAL_ELEMENT_ADDRESS                    = 0xC0000285,
	STATUS_MAGAZINE_NOT_PRESENT                       = 0xC0000286,
	STATUS_REINITIALIZATION_NEEDED                    = 0xC0000287,
	STATUS_DEVICE_REQUIRES_CLEANING                   = 0x80000288,
	STATUS_DEVICE_DOOR_OPEN                           = 0x80000289,
	STATUS_ENCRYPTION_FAILED                          = 0xC000028A,
	STATUS_DECRYPTION_FAILED                          = 0xC000028B,
	STATUS_RANGE_NOT_FOUND                            = 0xC000028C,
	STATUS_NO_RECOVERY_POLICY                         = 0xC000028D,
	STATUS_NO_EFS                                     = 0xC000028E,
	STATUS_WRONG_EFS                                  = 0xC000028F,
	STATUS_NO_USER_KEYS                               = 0xC0000290,
	STATUS_FILE_NOT_ENCRYPTED                         = 0xC0000291,
	STATUS_NOT_EXPORT_FORMAT                          = 0xC0000292,
	STATUS_FILE_ENCRYPTED                             = 0xC0000293,
	STATUS_WAKE_SYSTEM                                = 0x40000294,
	STATUS_WMI_GUID_NOT_FOUND                         = 0xC0000295,
	STATUS_WMI_INSTANCE_NOT_FOUND                     = 0xC0000296,
	STATUS_WMI_ITEMID_NOT_FOUND                       = 0xC0000297,
	STATUS_WMI_TRY_AGAIN                              = 0xC0000298,
	STATUS_SHARED_POLICY                              = 0xC0000299,
	STATUS_POLICY_OBJECT_NOT_FOUND                    = 0xC000029A,
	STATUS_POLICY_ONLY_IN_DS                          = 0xC000029B,
	STATUS_VOLUME_NOT_UPGRADED                        = 0xC000029C,
	STATUS_REMOTE_STORAGE_NOT_ACTIVE                  = 0xC000029D,
	STATUS_REMOTE_STORAGE_MEDIA_ERROR                 = 0xC000029E,
	STATUS_NO_TRACKING_SERVICE                        = 0xC000029F,
	STATUS_SERVER_SID_MISMATCH                        = 0xC00002A0,
	STATUS_DS_NO_ATTRIBUTE_OR_VALUE                   = 0xC00002A1,
	STATUS_DS_INVALID_ATTRIBUTE_SYNTAX                = 0xC00002A2,
	STATUS_DS_ATTRIBUTE_TYPE_UNDEFINED                = 0xC00002A3,
	STATUS_DS_ATTRIBUTE_OR_VALUE_EXISTS               = 0xC00002A4,
	STATUS_DS_BUSY                                    = 0xC00002A5,
	STATUS_DS_UNAVAILABLE                             = 0xC00002A6,
	STATUS_DS_NO_RIDS_ALLOCATED                       = 0xC00002A7,
	STATUS_DS_NO_MORE_RIDS                            = 0xC00002A8,
	STATUS_DS_INCORRECT_ROLE_OWNER                    = 0xC00002A9,
	STATUS_DS_RIDMGR_INIT_ERROR                       = 0xC00002AA,
	STATUS_DS_OBJ_CLASS_VIOLATION                     = 0xC00002AB,
	STATUS_DS_CANT_ON_NON_LEAF                        = 0xC00002AC,
	STATUS_DS_CANT_ON_RDN                             = 0xC00002AD,
	STATUS_DS_CANT_MOD_OBJ_CLASS                      = 0xC00002AE,
	STATUS_DS_CROSS_DOM_MOVE_FAILED                   = 0xC00002AF,
	STATUS_DS_GC_NOT_AVAILABLE                        = 0xC00002B0,
	STATUS_DIRECTORY_SERVICE_REQUIRED                 = 0xC00002B1,
	STATUS_REPARSE_ATTRIBUTE_CONFLICT                 = 0xC00002B2,
	STATUS_CANT_ENABLE_DENY_ONLY                      = 0xC00002B3,
	STATUS_FLOAT_MULTIPLE_FAULTS                      = 0xC00002B4,
	STATUS_FLOAT_MULTIPLE_TRAPS                       = 0xC00002B5,
	STATUS_DEVICE_REMOVED                             = 0xC00002B6,
	STATUS_JOURNAL_DELETE_IN_PROGRESS                 = 0xC00002B7,
	STATUS_JOURNAL_NOT_ACTIVE                         = 0xC00002B8,
	STATUS_NOINTERFACE                                = 0xC00002B9,
	STATUS_DS_ADMIN_LIMIT_EXCEEDED                    = 0xC00002C1,
	STATUS_DRIVER_FAILED_SLEEP                        = 0xC00002C2,
	STATUS_MUTUAL_AUTHENTICATION_FAILED               = 0xC00002C3,
	STATUS_CORRUPT_SYSTEM_FILE                        = 0xC00002C4,
	STATUS_DATATYPE_MISALIGNMENT_ERROR                = 0xC00002C5,
	STATUS_WMI_READ_ONLY                              = 0xC00002C6,
	STATUS_WMI_SET_FAILURE                            = 0xC00002C7,
	STATUS_COMMITMENT_MINIMUM                         = 0xC00002C8,
	STATUS_REG_NAT_CONSUMPTION                        = 0xC00002C9,
	STATUS_TRANSPORT_FULL                             = 0xC00002CA,
	STATUS_DS_SAM_INIT_FAILURE                        = 0xC00002CB,
	STATUS_ONLY_IF_CONNECTED                          = 0xC00002CC,
	STATUS_DS_SENSITIVE_GROUP_VIOLATION               = 0xC00002CD,
	STATUS_PNP_RESTART_ENUMERATION                    = 0xC00002CE,
	STATUS_JOURNAL_ENTRY_DELETED                      = 0xC00002CF,
	STATUS_DS_CANT_MOD_PRIMARYGROUPID                 = 0xC00002D0,
	STATUS_SYSTEM_IMAGE_BAD_SIGNATURE                 = 0xC00002D1,
	STATUS_PNP_REBOOT_REQUIRED                        = 0xC00002D2,
	STATUS_POWER_STATE_INVALID                        = 0xC00002D3,
	STATUS_DS_INVALID_GROUP_TYPE                      = 0xC00002D4,
	STATUS_DS_NO_NEST_GLOBALGROUP_IN_MIXEDDOMAIN      = 0xC00002D5,
	STATUS_DS_NO_NEST_LOCALGROUP_IN_MIXEDDOMAIN       = 0xC00002D6,
	STATUS_DS_GLOBAL_CANT_HAVE_LOCAL_MEMBER           = 0xC00002D7,
	STATUS_DS_GLOBAL_CANT_HAVE_UNIVERSAL_MEMBER       = 0xC00002D8,
	STATUS_DS_UNIVERSAL_CANT_HAVE_LOCAL_MEMBER        = 0xC00002D9,
	STATUS_DS_GLOBAL_CANT_HAVE_CROSSDOMAIN_MEMBER     = 0xC00002DA,
	STATUS_DS_LOCAL_CANT_HAVE_CROSSDOMAIN_LOCAL_MEMBER= 0xC00002DB,
	STATUS_DS_HAVE_PRIMARY_MEMBERS                    = 0xC00002DC,
	STATUS_WMI_NOT_SUPPORTED                          = 0xC00002DD,
	STATUS_INSUFFICIENT_POWER                         = 0xC00002DE,
	STATUS_SAM_NEED_BOOTKEY_PASSWORD                  = 0xC00002DF,
	STATUS_SAM_NEED_BOOTKEY_FLOPPY                    = 0xC00002E0,
	STATUS_DS_CANT_START                              = 0xC00002E1,
	STATUS_DS_INIT_FAILURE                            = 0xC00002E2,
	STATUS_SAM_INIT_FAILURE                           = 0xC00002E3,
	STATUS_DS_GC_REQUIRED                             = 0xC00002E4,
	STATUS_DS_LOCAL_MEMBER_OF_LOCAL_ONLY              = 0xC00002E5,
	STATUS_DS_NO_FPO_IN_UNIVERSAL_GROUPS              = 0xC00002E6,
	STATUS_DS_MACHINE_ACCOUNT_QUOTA_EXCEEDED          = 0xC00002E7,
	STATUS_MULTIPLE_FAULT_VIOLATION                   = 0xC00002E8,
	STATUS_CURRENT_DOMAIN_NOT_ALLOWED                 = 0xC00002E9,
	STATUS_CANNOT_MAKE                                = 0xC00002EA,
	STATUS_SYSTEM_SHUTDOWN                            = 0xC00002EB,
	STATUS_DS_INIT_FAILURE_CONSOLE                    = 0xC00002EC,
	STATUS_DS_SAM_INIT_FAILURE_CONSOLE                = 0xC00002ED,
	STATUS_UNFINISHED_CONTEXT_DELETED                 = 0xC00002EE,
	STATUS_NO_TGT_REPLY                               = 0xC00002EF,
	STATUS_OBJECTID_NOT_FOUND                         = 0xC00002F0,
	STATUS_NO_IP_ADDRESSES                            = 0xC00002F1,
	STATUS_WRONG_CREDENTIAL_HANDLE                    = 0xC00002F2,
	STATUS_CRYPTO_SYSTEM_INVALID                      = 0xC00002F3,
	STATUS_MAX_REFERRALS_EXCEEDED                     = 0xC00002F4,
	STATUS_MUST_BE_KDC                                = 0xC00002F5,
	STATUS_STRONG_CRYPTO_NOT_SUPPORTED                = 0xC00002F6,
	STATUS_TOO_MANY_PRINCIPALS                        = 0xC00002F7,
	STATUS_NO_PA_DATA                                 = 0xC00002F8,
	STATUS_PKINIT_NAME_MISMATCH                       = 0xC00002F9,
	STATUS_SMARTCARD_LOGON_REQUIRED                   = 0xC00002FA,
	STATUS_KDC_INVALID_REQUEST                        = 0xC00002FB,
	STATUS_KDC_UNABLE_TO_REFER                        = 0xC00002FC,
	STATUS_KDC_UNKNOWN_ETYPE                          = 0xC00002FD,
	STATUS_SHUTDOWN_IN_PROGRESS                       = 0xC00002FE,
	STATUS_SERVER_SHUTDOWN_IN_PROGRESS                = 0xC00002FF,
	STATUS_NOT_SUPPORTED_ON_SBS                       = 0xC0000300,
	STATUS_WMI_GUID_DISCONNECTED                      = 0xC0000301,
	STATUS_WMI_ALREADY_DISABLED                       = 0xC0000302,
	STATUS_WMI_ALREADY_ENABLED                        = 0xC0000303,
	STATUS_MFT_TOO_FRAGMENTED                         = 0xC0000304,
	STATUS_COPY_PROTECTION_FAILURE                    = 0xC0000305,
	STATUS_CSS_AUTHENTICATION_FAILURE                 = 0xC0000306,
	STATUS_CSS_KEY_NOT_PRESENT                        = 0xC0000307,
	STATUS_CSS_KEY_NOT_ESTABLISHED                    = 0xC0000308,
	STATUS_CSS_SCRAMBLED_SECTOR                       = 0xC0000309,
	STATUS_CSS_REGION_MISMATCH                        = 0xC000030A,
	STATUS_CSS_RESETS_EXHAUSTED                       = 0xC000030B,
	STATUS_PKINIT_FAILURE                             = 0xC0000320,
	STATUS_SMARTCARD_SUBSYSTEM_FAILURE                = 0xC0000321,
	STATUS_NO_KERB_KEY                                = 0xC0000322,
	STATUS_HOST_DOWN                                  = 0xC0000350,
	STATUS_UNSUPPORTED_PREAUTH                        = 0xC0000351,
	STATUS_EFS_ALG_BLOB_TOO_BIG                       = 0xC0000352,
	STATUS_PORT_NOT_SET                               = 0xC0000353,
	STATUS_DEBUGGER_INACTIVE                          = 0xC0000354,
	STATUS_DS_VERSION_CHECK_FAILURE                   = 0xC0000355,
	STATUS_AUDITING_DISABLED                          = 0xC0000356,
	STATUS_PRENT4_MACHINE_ACCOUNT                     = 0xC0000357,
	STATUS_DS_AG_CANT_HAVE_UNIVERSAL_MEMBER           = 0xC0000358,
	STATUS_INVALID_IMAGE_WIN_32                       = 0xC0000359,
	STATUS_INVALID_IMAGE_WIN_64                       = 0xC000035A,
	STATUS_BAD_BINDINGS                               = 0xC000035B,
	STATUS_NETWORK_SESSION_EXPIRED                    = 0xC000035C,
	STATUS_APPHELP_BLOCK                              = 0xC000035D,
	STATUS_ALL_SIDS_FILTERED                          = 0xC000035E,
	STATUS_NOT_SAFE_MODE_DRIVER                       = 0xC000035F,
	STATUS_ACCESS_DISABLED_BY_POLICY_DEFAULT          = 0xC0000361,
	STATUS_ACCESS_DISABLED_BY_POLICY_PATH             = 0xC0000362,
	STATUS_ACCESS_DISABLED_BY_POLICY_PUBLISHER        = 0xC0000363,
	STATUS_ACCESS_DISABLED_BY_POLICY_OTHER            = 0xC0000364,
	STATUS_FAILED_DRIVER_ENTRY                        = 0xC0000365,
	STATUS_DEVICE_ENUMERATION_ERROR                   = 0xC0000366,
	STATUS_WAIT_FOR_OPLOCK                            = 0x00000367,
	STATUS_MOUNT_POINT_NOT_RESOLVED                   = 0xC0000368,
	STATUS_INVALID_DEVICE_OBJECT_PARAMETER            = 0xC0000369,
	STATUS_MCA_OCCURED                                = 0xC000036A,
	STATUS_DRIVER_BLOCKED_CRITICAL                    = 0xC000036B,
	STATUS_DRIVER_BLOCKED                             = 0xC000036C,
	STATUS_DRIVER_DATABASE_ERROR                      = 0xC000036D,
	STATUS_SYSTEM_HIVE_TOO_LARGE                      = 0xC000036E,
	STATUS_INVALID_IMPORT_OF_NON_DLL                  = 0xC000036F,
	STATUS_DS_SHUTTING_DOWN                           = 0x40000370,
	STATUS_SMARTCARD_WRONG_PIN                        = 0xC0000380,
	STATUS_SMARTCARD_CARD_BLOCKED                     = 0xC0000381,
	STATUS_SMARTCARD_CARD_NOT_AUTHENTICATED           = 0xC0000382,
	STATUS_SMARTCARD_NO_CARD                          = 0xC0000383,
	STATUS_SMARTCARD_NO_KEY_CONTAINER                 = 0xC0000384,
	STATUS_SMARTCARD_NO_CERTIFICATE                   = 0xC0000385,
	STATUS_SMARTCARD_NO_KEYSET                        = 0xC0000386,
	STATUS_SMARTCARD_IO_ERROR                         = 0xC0000387,
	STATUS_DOWNGRADE_DETECTED                         = 0xC0000388,
	STATUS_SMARTCARD_CERT_REVOKED                     = 0xC0000389,
	STATUS_ISSUING_CA_UNTRUSTED                       = 0xC000038A,
	STATUS_REVOCATION_OFFLINE_C                       = 0xC000038B,
	STATUS_PKINIT_CLIENT_FAILURE                      = 0xC000038C,
	STATUS_SMARTCARD_CERT_EXPIRED                     = 0xC000038D,
	STATUS_DRIVER_FAILED_PRIOR_UNLOAD                 = 0xC000038E,
	STATUS_SMARTCARD_SILENT_CONTEXT                   = 0xC000038F,
	STATUS_PER_USER_TRUST_QUOTA_EXCEEDED              = 0xC0000401,
	STATUS_ALL_USER_TRUST_QUOTA_EXCEEDED              = 0xC0000402,
	STATUS_USER_DELETE_TRUST_QUOTA_EXCEEDED           = 0xC0000403,
	STATUS_DS_NAME_NOT_UNIQUE                         = 0xC0000404,
	STATUS_DS_DUPLICATE_ID_FOUND                      = 0xC0000405,
	STATUS_DS_GROUP_CONVERSION_ERROR                  = 0xC0000406,
	STATUS_VOLSNAP_PREPARE_HIBERNATE                  = 0xC0000407,
	STATUS_USER2USER_REQUIRED                         = 0xC0000408,
	STATUS_STACK_BUFFER_OVERRUN                       = 0xC0000409,
	STATUS_NO_S4U_PROT_SUPPORT                        = 0xC000040A,
	STATUS_CROSSREALM_DELEGATION_FAILURE              = 0xC000040B,
	STATUS_REVOCATION_OFFLINE_KDC                     = 0xC000040C,
	STATUS_ISSUING_CA_UNTRUSTED_KDC                   = 0xC000040D,
	STATUS_KDC_CERT_EXPIRED                           = 0xC000040E,
	STATUS_KDC_CERT_REVOKED                           = 0xC000040F,
	STATUS_PARAMETER_QUOTA_EXCEEDED                   = 0xC0000410,
	STATUS_HIBERNATION_FAILURE                        = 0xC0000411,
	STATUS_DELAY_LOAD_FAILED                          = 0xC0000412,
	STATUS_AUTHENTICATION_FIREWALL_FAILED             = 0xC0000413,
	STATUS_VDM_DISALLOWED                             = 0xC0000414,
	STATUS_HUNG_DISPLAY_DRIVER_THREAD                 = 0xC0000415,
	STATUS_ASSERTION_FAILURE                          = 0xC0000420,
	STATUS_CALLBACK_POP_STACK                         = 0xC0000423,
	STATUS_INCOMPATIBLE_DRIVER_BLOCKED                = 0xC0000424,
	STATUS_ENCOUNTERED_WRITE_IN_PROGRESS              = 0xC0000433,
	STATUS_WOW_ASSERTION                              = 0xC0009898,
	DBG_NO_STATE_CHANGE                               = 0xC0010001,
	DBG_APP_NOT_IDLE                                  = 0xC0010002,
	RPC_NT_INVALID_STRING_BINDING                     = 0xC0020001,
	RPC_NT_WRONG_KIND_OF_BINDING                      = 0xC0020002,
	RPC_NT_INVALID_BINDING                            = 0xC0020003,
	RPC_NT_PROTSEQ_NOT_SUPPORTED                      = 0xC0020004,
	RPC_NT_INVALID_RPC_PROTSEQ                        = 0xC0020005,
	RPC_NT_INVALID_STRING_UUID                        = 0xC0020006,
	RPC_NT_INVALID_ENDPOINT_FORMAT                    = 0xC0020007,
	RPC_NT_INVALID_NET_ADDR                           = 0xC0020008,
	RPC_NT_NO_ENDPOINT_FOUND                          = 0xC0020009,
	RPC_NT_INVALID_TIMEOUT                            = 0xC002000A,
	RPC_NT_OBJECT_NOT_FOUND                           = 0xC002000B,
	RPC_NT_ALREADY_REGISTERED                         = 0xC002000C,
	RPC_NT_TYPE_ALREADY_REGISTERED                    = 0xC002000D,
	RPC_NT_ALREADY_LISTENING                          = 0xC002000E,
	RPC_NT_NO_PROTSEQS_REGISTERED                     = 0xC002000F,
	RPC_NT_NOT_LISTENING                              = 0xC0020010,
	RPC_NT_UNKNOWN_MGR_TYPE                           = 0xC0020011,
	RPC_NT_UNKNOWN_IF                                 = 0xC0020012,
	RPC_NT_NO_BINDINGS                                = 0xC0020013,
	RPC_NT_NO_PROTSEQS                                = 0xC0020014,
	RPC_NT_CANT_CREATE_ENDPOINT                       = 0xC0020015,
	RPC_NT_OUT_OF_RESOURCES                           = 0xC0020016,
	RPC_NT_SERVER_UNAVAILABLE                         = 0xC0020017,
	RPC_NT_SERVER_TOO_BUSY                            = 0xC0020018,
	RPC_NT_INVALID_NETWORK_OPTIONS                    = 0xC0020019,
	RPC_NT_NO_CALL_ACTIVE                             = 0xC002001A,
	RPC_NT_CALL_FAILED                                = 0xC002001B,
	RPC_NT_CALL_FAILED_DNE                            = 0xC002001C,
	RPC_NT_PROTOCOL_ERROR                             = 0xC002001D,
	RPC_NT_UNSUPPORTED_TRANS_SYN                      = 0xC002001F,
	RPC_NT_UNSUPPORTED_TYPE                           = 0xC0020021,
	RPC_NT_INVALID_TAG                                = 0xC0020022,
	RPC_NT_INVALID_BOUND                              = 0xC0020023,
	RPC_NT_NO_ENTRY_NAME                              = 0xC0020024,
	RPC_NT_INVALID_NAME_SYNTAX                        = 0xC0020025,
	RPC_NT_UNSUPPORTED_NAME_SYNTAX                    = 0xC0020026,
	RPC_NT_UUID_NO_ADDRESS                            = 0xC0020028,
	RPC_NT_DUPLICATE_ENDPOINT                         = 0xC0020029,
	RPC_NT_UNKNOWN_AUTHN_TYPE                         = 0xC002002A,
	RPC_NT_MAX_CALLS_TOO_SMALL                        = 0xC002002B,
	RPC_NT_STRING_TOO_LONG                            = 0xC002002C,
	RPC_NT_PROTSEQ_NOT_FOUND                          = 0xC002002D,
	RPC_NT_PROCNUM_OUT_OF_RANGE                       = 0xC002002E,
	RPC_NT_BINDING_HAS_NO_AUTH                        = 0xC002002F,
	RPC_NT_UNKNOWN_AUTHN_SERVICE                      = 0xC0020030,
	RPC_NT_UNKNOWN_AUTHN_LEVEL                        = 0xC0020031,
	RPC_NT_INVALID_AUTH_IDENTITY                      = 0xC0020032,
	RPC_NT_UNKNOWN_AUTHZ_SERVICE                      = 0xC0020033,
	EPT_NT_INVALID_ENTRY                              = 0xC0020034,
	EPT_NT_CANT_PERFORM_OP                            = 0xC0020035,
	EPT_NT_NOT_REGISTERED                             = 0xC0020036,
	RPC_NT_NOTHING_TO_EXPORT                          = 0xC0020037,
	RPC_NT_INCOMPLETE_NAME                            = 0xC0020038,
	RPC_NT_INVALID_VERS_OPTION                        = 0xC0020039,
	RPC_NT_NO_MORE_MEMBERS                            = 0xC002003A,
	RPC_NT_NOT_ALL_OBJS_UNEXPORTED                    = 0xC002003B,
	RPC_NT_INTERFACE_NOT_FOUND                        = 0xC002003C,
	RPC_NT_ENTRY_ALREADY_EXISTS                       = 0xC002003D,
	RPC_NT_ENTRY_NOT_FOUND                            = 0xC002003E,
	RPC_NT_NAME_SERVICE_UNAVAILABLE                   = 0xC002003F,
	RPC_NT_INVALID_NAF_ID                             = 0xC0020040,
	RPC_NT_CANNOT_SUPPORT                             = 0xC0020041,
	RPC_NT_NO_CONTEXT_AVAILABLE                       = 0xC0020042,
	RPC_NT_INTERNAL_ERROR                             = 0xC0020043,
	RPC_NT_ZERO_DIVIDE                                = 0xC0020044,
	RPC_NT_ADDRESS_ERROR                              = 0xC0020045,
	RPC_NT_FP_DIV_ZERO                                = 0xC0020046,
	RPC_NT_FP_UNDERFLOW                               = 0xC0020047,
	RPC_NT_FP_OVERFLOW                                = 0xC0020048,
	RPC_NT_NO_MORE_ENTRIES                            = 0xC0030001,
	RPC_NT_SS_CHAR_TRANS_OPEN_FAIL                    = 0xC0030002,
	RPC_NT_SS_CHAR_TRANS_SHORT_FILE                   = 0xC0030003,
	RPC_NT_SS_IN_NULL_CONTEXT                         = 0xC0030004,
	RPC_NT_SS_CONTEXT_MISMATCH                        = 0xC0030005,
	RPC_NT_SS_CONTEXT_DAMAGED                         = 0xC0030006,
	RPC_NT_SS_HANDLES_MISMATCH                        = 0xC0030007,
	RPC_NT_SS_CANNOT_GET_CALL_HANDLE                  = 0xC0030008,
	RPC_NT_NULL_REF_POINTER                           = 0xC0030009,
	RPC_NT_ENUM_VALUE_OUT_OF_RANGE                    = 0xC003000A,
	RPC_NT_BYTE_COUNT_TOO_SMALL                       = 0xC003000B,
	RPC_NT_BAD_STUB_DATA                              = 0xC003000C,
	RPC_NT_CALL_IN_PROGRESS                           = 0xC0020049,
	RPC_NT_NO_MORE_BINDINGS                           = 0xC002004A,
	RPC_NT_GROUP_MEMBER_NOT_FOUND                     = 0xC002004B,
	EPT_NT_CANT_CREATE                                = 0xC002004C,
	RPC_NT_INVALID_OBJECT                             = 0xC002004D,
	RPC_NT_NO_INTERFACES                              = 0xC002004F,
	RPC_NT_CALL_CANCELLED                             = 0xC0020050,
	RPC_NT_BINDING_INCOMPLETE                         = 0xC0020051,
	RPC_NT_COMM_FAILURE                               = 0xC0020052,
	RPC_NT_UNSUPPORTED_AUTHN_LEVEL                    = 0xC0020053,
	RPC_NT_NO_PRINC_NAME                              = 0xC0020054,
	RPC_NT_NOT_RPC_ERROR                              = 0xC0020055,
	RPC_NT_UUID_LOCAL_ONLY                            = 0x40020056,
	RPC_NT_SEC_PKG_ERROR                              = 0xC0020057,
	RPC_NT_NOT_CANCELLED                              = 0xC0020058,
	RPC_NT_INVALID_ES_ACTION                          = 0xC0030059,
	RPC_NT_WRONG_ES_VERSION                           = 0xC003005A,
	RPC_NT_WRONG_STUB_VERSION                         = 0xC003005B,
	RPC_NT_INVALID_PIPE_OBJECT                        = 0xC003005C,
	RPC_NT_INVALID_PIPE_OPERATION                     = 0xC003005D,
	RPC_NT_WRONG_PIPE_VERSION                         = 0xC003005E,
	RPC_NT_PIPE_CLOSED                                = 0xC003005F,
	RPC_NT_PIPE_DISCIPLINE_ERROR                      = 0xC0030060,
	RPC_NT_PIPE_EMPTY                                 = 0xC0030061,
	RPC_NT_INVALID_ASYNC_HANDLE                       = 0xC0020062,
	RPC_NT_INVALID_ASYNC_CALL                         = 0xC0020063,
	RPC_NT_SEND_INCOMPLETE                            = 0x400200AF,
	STATUS_ACPI_INVALID_OPCODE                        = 0xC0140001,
	STATUS_ACPI_STACK_OVERFLOW                        = 0xC0140002,
	STATUS_ACPI_ASSERT_FAILED                         = 0xC0140003,
	STATUS_ACPI_INVALID_INDEX                         = 0xC0140004,
	STATUS_ACPI_INVALID_ARGUMENT                      = 0xC0140005,
	STATUS_ACPI_FATAL                                 = 0xC0140006,
	STATUS_ACPI_INVALID_SUPERNAME                     = 0xC0140007,
	STATUS_ACPI_INVALID_ARGTYPE                       = 0xC0140008,
	STATUS_ACPI_INVALID_OBJTYPE                       = 0xC0140009,
	STATUS_ACPI_INVALID_TARGETTYPE                    = 0xC014000A,
	STATUS_ACPI_INCORRECT_ARGUMENT_COUNT              = 0xC014000B,
	STATUS_ACPI_ADDRESS_NOT_MAPPED                    = 0xC014000C,
	STATUS_ACPI_INVALID_EVENTTYPE                     = 0xC014000D,
	STATUS_ACPI_HANDLER_COLLISION                     = 0xC014000E,
	STATUS_ACPI_INVALID_DATA                          = 0xC014000F,
	STATUS_ACPI_INVALID_REGION                        = 0xC0140010,
	STATUS_ACPI_INVALID_ACCESS_SIZE                   = 0xC0140011,
	STATUS_ACPI_ACQUIRE_GLOBAL_LOCK                   = 0xC0140012,
	STATUS_ACPI_ALREADY_INITIALIZED                   = 0xC0140013,
	STATUS_ACPI_NOT_INITIALIZED                       = 0xC0140014,
	STATUS_ACPI_INVALID_MUTEX_LEVEL                   = 0xC0140015,
	STATUS_ACPI_MUTEX_NOT_OWNED                       = 0xC0140016,
	STATUS_ACPI_MUTEX_NOT_OWNER                       = 0xC0140017,
	STATUS_ACPI_RS_ACCESS                             = 0xC0140018,
	STATUS_ACPI_INVALID_TABLE                         = 0xC0140019,
	STATUS_ACPI_REG_HANDLER_FAILED                    = 0xC0140020,
	STATUS_ACPI_POWER_REQUEST_FAILED                  = 0xC0140021,
	STATUS_CTX_WINSTATION_NAME_INVALID                = 0xC00A0001,
	STATUS_CTX_INVALID_PD                             = 0xC00A0002,
	STATUS_CTX_PD_NOT_FOUND                           = 0xC00A0003,
	STATUS_CTX_CDM_CONNECT                            = 0x400A0004,
	STATUS_CTX_CDM_DISCONNECT                         = 0x400A0005,
	STATUS_CTX_CLOSE_PENDING                          = 0xC00A0006,
	STATUS_CTX_NO_OUTBUF                              = 0xC00A0007,
	STATUS_CTX_MODEM_INF_NOT_FOUND                    = 0xC00A0008,
	STATUS_CTX_INVALID_MODEMNAME                      = 0xC00A0009,
	STATUS_CTX_RESPONSE_ERROR                         = 0xC00A000A,
	STATUS_CTX_MODEM_RESPONSE_TIMEOUT                 = 0xC00A000B,
	STATUS_CTX_MODEM_RESPONSE_NO_CARRIER              = 0xC00A000C,
	STATUS_CTX_MODEM_RESPONSE_NO_DIALTONE             = 0xC00A000D,
	STATUS_CTX_MODEM_RESPONSE_BUSY                    = 0xC00A000E,
	STATUS_CTX_MODEM_RESPONSE_VOICE                   = 0xC00A000F,
	STATUS_CTX_TD_ERROR                               = 0xC00A0010,
	STATUS_CTX_LICENSE_CLIENT_INVALID                 = 0xC00A0012,
	STATUS_CTX_LICENSE_NOT_AVAILABLE                  = 0xC00A0013,
	STATUS_CTX_LICENSE_EXPIRED                        = 0xC00A0014,
	STATUS_CTX_WINSTATION_NOT_FOUND                   = 0xC00A0015,
	STATUS_CTX_WINSTATION_NAME_COLLISION              = 0xC00A0016,
	STATUS_CTX_WINSTATION_BUSY                        = 0xC00A0017,
	STATUS_CTX_BAD_VIDEO_MODE                         = 0xC00A0018,
	STATUS_CTX_GRAPHICS_INVALID                       = 0xC00A0022,
	STATUS_CTX_NOT_CONSOLE                            = 0xC00A0024,
	STATUS_CTX_CLIENT_QUERY_TIMEOUT                   = 0xC00A0026,
	STATUS_CTX_CONSOLE_DISCONNECT                     = 0xC00A0027,
	STATUS_CTX_CONSOLE_CONNECT                        = 0xC00A0028,
	STATUS_CTX_SHADOW_DENIED                          = 0xC00A002A,
	STATUS_CTX_WINSTATION_ACCESS_DENIED               = 0xC00A002B,
	STATUS_CTX_INVALID_WD                             = 0xC00A002E,
	STATUS_CTX_WD_NOT_FOUND                           = 0xC00A002F,
	STATUS_CTX_SHADOW_INVALID                         = 0xC00A0030,
	STATUS_CTX_SHADOW_DISABLED                        = 0xC00A0031,
	STATUS_RDP_PROTOCOL_ERROR                         = 0xC00A0032,
	STATUS_CTX_CLIENT_LICENSE_NOT_SET                 = 0xC00A0033,
	STATUS_CTX_CLIENT_LICENSE_IN_USE                  = 0xC00A0034,
	STATUS_CTX_SHADOW_ENDED_BY_MODE_CHANGE            = 0xC00A0035,
	STATUS_CTX_SHADOW_NOT_RUNNING                     = 0xC00A0036,
	STATUS_PNP_BAD_MPS_TABLE                          = 0xC0040035,
	STATUS_PNP_TRANSLATION_FAILED                     = 0xC0040036,
	STATUS_PNP_IRQ_TRANSLATION_FAILED                 = 0xC0040037,
	STATUS_PNP_INVALID_ID                             = 0xC0040038,
	STATUS_SXS_SECTION_NOT_FOUND                      = 0xC0150001,
	STATUS_SXS_CANT_GEN_ACTCTX                        = 0xC0150002,
	STATUS_SXS_INVALID_ACTCTXDATA_FORMAT              = 0xC0150003,
	STATUS_SXS_ASSEMBLY_NOT_FOUND                     = 0xC0150004,
	STATUS_SXS_MANIFEST_FORMAT_ERROR                  = 0xC0150005,
	STATUS_SXS_MANIFEST_PARSE_ERROR                   = 0xC0150006,
	STATUS_SXS_ACTIVATION_CONTEXT_DISABLED            = 0xC0150007,
	STATUS_SXS_KEY_NOT_FOUND                          = 0xC0150008,
	STATUS_SXS_VERSION_CONFLICT                       = 0xC0150009,
	STATUS_SXS_WRONG_SECTION_TYPE                     = 0xC015000A,
	STATUS_SXS_THREAD_QUERIES_DISABLED                = 0xC015000B,
	STATUS_SXS_ASSEMBLY_MISSING                       = 0xC015000C,
	STATUS_SXS_RELEASE_ACTIVATION_CONTEXT             = 0x4015000D,
	STATUS_SXS_PROCESS_DEFAULT_ALREADY_SET            = 0xC015000E,
	STATUS_SXS_EARLY_DEACTIVATION                     = 0xC015000F,
	STATUS_SXS_INVALID_DEACTIVATION                   = 0xC0150010,
	STATUS_SXS_MULTIPLE_DEACTIVATION                  = 0xC0150011,
	STATUS_SXS_SYSTEM_DEFAULT_ACTIVATION_CONTEXT_EMPTY= 0xC0150012,
	STATUS_SXS_PROCESS_TERMINATION_REQUESTED          = 0xC0150013,
	STATUS_SXS_CORRUPT_ACTIVATION_STACK               = 0xC0150014,
	STATUS_SXS_CORRUPTION                             = 0xC0150015,
	STATUS_CLUSTER_INVALID_NODE                       = 0xC0130001,
	STATUS_CLUSTER_NODE_EXISTS                        = 0xC0130002,
	STATUS_CLUSTER_JOIN_IN_PROGRESS                   = 0xC0130003,
	STATUS_CLUSTER_NODE_NOT_FOUND                     = 0xC0130004,
	STATUS_CLUSTER_LOCAL_NODE_NOT_FOUND               = 0xC0130005,
	STATUS_CLUSTER_NETWORK_EXISTS                     = 0xC0130006,
	STATUS_CLUSTER_NETWORK_NOT_FOUND                  = 0xC0130007,
	STATUS_CLUSTER_NETINTERFACE_EXISTS                = 0xC0130008,
	STATUS_CLUSTER_NETINTERFACE_NOT_FOUND             = 0xC0130009,
	STATUS_CLUSTER_INVALID_REQUEST                    = 0xC013000A,
	STATUS_CLUSTER_INVALID_NETWORK_PROVIDER           = 0xC013000B,
	STATUS_CLUSTER_NODE_DOWN                          = 0xC013000C,
	STATUS_CLUSTER_NODE_UNREACHABLE                   = 0xC013000D,
	STATUS_CLUSTER_NODE_NOT_MEMBER                    = 0xC013000E,
	STATUS_CLUSTER_JOIN_NOT_IN_PROGRESS               = 0xC013000F,
	STATUS_CLUSTER_INVALID_NETWORK                    = 0xC0130010,
	STATUS_CLUSTER_NO_NET_ADAPTERS                    = 0xC0130011,
	STATUS_CLUSTER_NODE_UP                            = 0xC0130012,
	STATUS_CLUSTER_NODE_PAUSED                        = 0xC0130013,
	STATUS_CLUSTER_NODE_NOT_PAUSED                    = 0xC0130014,
	STATUS_CLUSTER_NO_SECURITY_CONTEXT                = 0xC0130015,
	STATUS_CLUSTER_NETWORK_NOT_INTERNAL               = 0xC0130016,
	STATUS_CLUSTER_POISONED                           = 0xC0130017,
	STATUS_LOG_SECTOR_INVALID                         = 0xC01A0001,
	STATUS_LOG_SECTOR_PARITY_INVALID                  = 0xC01A0002,
	STATUS_LOG_SECTOR_REMAPPED                        = 0xC01A0003,
	STATUS_LOG_BLOCK_INCOMPLETE                       = 0xC01A0004,
	STATUS_LOG_INVALID_RANGE                          = 0xC01A0005,
	STATUS_LOG_BLOCKS_EXHAUSTED                       = 0xC01A0006,
	STATUS_LOG_READ_CONTEXT_INVALID                   = 0xC01A0007,
	STATUS_LOG_RESTART_INVALID                        = 0xC01A0008,
	STATUS_LOG_BLOCK_VERSION                          = 0xC01A0009,
	STATUS_LOG_BLOCK_INVALID                          = 0xC01A000A,
	STATUS_LOG_READ_MODE_INVALID                      = 0xC01A000B,
	STATUS_LOG_NO_RESTART                             = 0x401A000C,
	STATUS_LOG_METADATA_CORRUPT                       = 0xC01A000D,
	STATUS_LOG_METADATA_INVALID                       = 0xC01A000E,
	STATUS_LOG_METADATA_INCONSISTENT                  = 0xC01A000F,
	STATUS_LOG_RESERVATION_INVALID                    = 0xC01A0010,
	STATUS_LOG_CANT_DELETE                            = 0xC01A0011,
	STATUS_LOG_CONTAINER_LIMIT_EXCEEDED               = 0xC01A0012,
	STATUS_LOG_START_OF_LOG                           = 0xC01A0013,
	STATUS_LOG_POLICY_ALREADY_INSTALLED               = 0xC01A0014,
	STATUS_LOG_POLICY_NOT_INSTALLED                   = 0xC01A0015,
	STATUS_LOG_POLICY_INVALID                         = 0xC01A0016,
	STATUS_LOG_POLICY_CONFLICT                        = 0xC01A0017,
	STATUS_LOG_PINNED_ARCHIVE_TAIL                    = 0xC01A0018,
	STATUS_LOG_RECORD_NONEXISTENT                     = 0xC01A0019,
	STATUS_LOG_RECORDS_RESERVED_INVALID               = 0xC01A001A,
	STATUS_LOG_SPACE_RESERVED_INVALID                 = 0xC01A001B,
	STATUS_LOG_TAIL_INVALID                           = 0xC01A001C,
	STATUS_LOG_FULL                                   = 0xC01A001D,
	STATUS_LOG_MULTIPLEXED                            = 0xC01A001E,
	STATUS_LOG_DEDICATED                              = 0xC01A001F,
	STATUS_LOG_ARCHIVE_NOT_IN_PROGRESS                = 0xC01A0020,
	STATUS_LOG_ARCHIVE_IN_PROGRESS                    = 0xC01A0021,
	STATUS_LOG_EPHEMERAL                              = 0xC01A0022,
	STATUS_LOG_NOT_ENOUGH_CONTAINERS                  = 0xC01A0023,
	STATUS_LOG_CLIENT_ALREADY_REGISTERED              = 0xC01A0024,
	STATUS_LOG_CLIENT_NOT_REGISTERED                  = 0xC01A0025,
	STATUS_LOG_FULL_HANDLER_IN_PROGRESS               = 0xC01A0026,
	STATUS_LOG_CONTAINER_READ_FAILED                  = 0xC01A0027,
	STATUS_LOG_CONTAINER_WRITE_FAILED                 = 0xC01A0028,
	STATUS_LOG_CONTAINER_OPEN_FAILED                  = 0xC01A0029,
	STATUS_LOG_CONTAINER_STATE_INVALID                = 0xC01A002A,
	STATUS_LOG_STATE_INVALID                          = 0xC01A002B,
	STATUS_LOG_PINNED                                 = 0xC01A002C,
	STATUS_LOG_METADATA_FLUSH_FAILED                  = 0xC01A002D,
	STATUS_LOG_INCONSISTENT_SECURITY                  = 0xC01A002E,
	STATUS_COULD_NOT_RESIZE_LOG                       = 0x80190009
} NTSTATUS;


typedef enum WindowMessages {
    WM_ACTIVATE           = 0x0006,
    WM_ACTIVATEAPP        = 0x001C,
    WM_AFXFIRST           = 0x0360,
    WM_AFXLAST        = 0x037F,
    WM_APP            = 0x8000,
    WM_ASKCBFORMATNAME    = 0x030C,
    WM_CANCELJOURNAL      = 0x004B,
    WM_CANCELMODE         = 0x001F,
    WM_CAPTURECHANGED     = 0x0215,
    WM_CHANGECBCHAIN      = 0x030D,
    WM_CHANGEUISTATE      = 0x0127,
    WM_CHAR           = 0x0102,
    WM_CHARTOITEM         = 0x002F,
    WM_CHILDACTIVATE      = 0x0022,
    WM_CLEAR          = 0x0303,
    WM_CLOSE          = 0x0010,
    WM_COMMAND        = 0x0111,
    WM_COMPACTING         = 0x0041,
    WM_COMPAREITEM        = 0x0039,
    WM_CONTEXTMENU        = 0x007B,
    WM_COPY           = 0x0301,
    WM_COPYDATA           = 0x004A,
    WM_CREATE         = 0x0001,
    WM_CTLCOLORBTN        = 0x0135,
    WM_CTLCOLORDLG        = 0x0136,
    WM_CTLCOLOREDIT       = 0x0133,
    WM_CTLCOLORLISTBOX    = 0x0134,
    WM_CTLCOLORMSGBOX     = 0x0132,
    WM_CTLCOLORSCROLLBAR      = 0x0137,
    WM_CTLCOLORSTATIC     = 0x0138,
    WM_CUT            = 0x0300,
    WM_DEADCHAR           = 0x0103,
    WM_DELETEITEM         = 0x002D,
    WM_DESTROY        = 0x0002,
    WM_DESTROYCLIPBOARD       = 0x0307,
    WM_DEVICECHANGE       = 0x0219,
    WM_DEVMODECHANGE      = 0x001B,
    WM_DISPLAYCHANGE      = 0x007E,
    WM_DRAWCLIPBOARD      = 0x0308,
    WM_DRAWITEM           = 0x002B,
    WM_DROPFILES          = 0x0233,
    WM_ENABLE         = 0x000A,
    WM_ENDSESSION         = 0x0016,
    WM_ENTERIDLE          = 0x0121,
    WM_ENTERMENULOOP      = 0x0211,
    WM_ENTERSIZEMOVE      = 0x0231,
    WM_ERASEBKGND         = 0x0014,
    WM_EXITMENULOOP       = 0x0212,
    WM_EXITSIZEMOVE       = 0x0232,
    WM_FONTCHANGE         = 0x001D,
    WM_GETDLGCODE         = 0x0087,
    WM_GETFONT        = 0x0031,
    WM_GETHOTKEY          = 0x0033,
    WM_GETICON        = 0x007F,
    WM_GETMINMAXINFO      = 0x0024,
    WM_GETOBJECT          = 0x003D,
    WM_GETTEXT        = 0x000D,
    WM_GETTEXTLENGTH      = 0x000E,
    WM_HANDHELDFIRST      = 0x0358,
    WM_HANDHELDLAST       = 0x035F,
    WM_HELP           = 0x0053,
    WM_HOTKEY         = 0x0312,
    WM_HSCROLL        = 0x0114,
    WM_HSCROLLCLIPBOARD       = 0x030E,
    WM_ICONERASEBKGND     = 0x0027,
    WM_IME_CHAR           = 0x0286,
    WM_IME_COMPOSITION    = 0x010F,
    WM_IME_COMPOSITIONFULL    = 0x0284,
    WM_IME_CONTROL        = 0x0283,
    WM_IME_ENDCOMPOSITION     = 0x010E,
    WM_IME_KEYDOWN        = 0x0290,
    WM_IME_KEYLAST        = 0x010F,
    WM_IME_KEYUP          = 0x0291,
    WM_IME_NOTIFY         = 0x0282,
    WM_IME_REQUEST        = 0x0288,
    WM_IME_SELECT         = 0x0285,
    WM_IME_SETCONTEXT     = 0x0281,
    WM_IME_STARTCOMPOSITION   = 0x010D,
    WM_INITDIALOG         = 0x0110,
    WM_INITMENU           = 0x0116,
    WM_INITMENUPOPUP      = 0x0117,
    WM_INPUTLANGCHANGE    = 0x0051,
    WM_INPUTLANGCHANGEREQUEST = 0x0050,
    WM_KEYDOWN        = 0x0100,
    WM_KEYFIRST           = 0x0100,
    WM_KEYLAST        = 0x0108,
    WM_KEYUP          = 0x0101,
    WM_KILLFOCUS          = 0x0008,
    WM_LBUTTONDBLCLK      = 0x0203,
    WM_LBUTTONDOWN        = 0x0201,
    WM_LBUTTONUP          = 0x0202,
    WM_MBUTTONDBLCLK      = 0x0209,
    WM_MBUTTONDOWN        = 0x0207,
    WM_MBUTTONUP          = 0x0208,
    WM_MDIACTIVATE        = 0x0222,
    WM_MDICASCADE         = 0x0227,
    WM_MDICREATE          = 0x0220,
    WM_MDIDESTROY         = 0x0221,
    WM_MDIGETACTIVE       = 0x0229,
    WM_MDIICONARRANGE     = 0x0228,
    WM_MDIMAXIMIZE        = 0x0225,
    WM_MDINEXT        = 0x0224,
    WM_MDIREFRESHMENU     = 0x0234,
    WM_MDIRESTORE         = 0x0223,
    WM_MDISETMENU         = 0x0230,
    WM_MDITILE        = 0x0226,
    WM_MEASUREITEM        = 0x002C,
    WM_MENUCHAR           = 0x0120,
    WM_MENUCOMMAND        = 0x0126,
    WM_MENUDRAG           = 0x0123,
    WM_MENUGETOBJECT      = 0x0124,
    WM_MENURBUTTONUP      = 0x0122,
    WM_MENUSELECT         = 0x011F,
    WM_MOUSEACTIVATE      = 0x0021,
    WM_MOUSEFIRST         = 0x0200,
    WM_MOUSEHOVER         = 0x02A1,
    WM_MOUSELAST          = 0x020D,
    WM_MOUSELEAVE         = 0x02A3,
    WM_MOUSEMOVE          = 0x0200,
    WM_MOUSEWHEEL         = 0x020A,
    WM_MOUSEHWHEEL        = 0x020E,
    WM_MOVE           = 0x0003,
    WM_MOVING         = 0x0216,
    WM_NCACTIVATE         = 0x0086,
    WM_NCCALCSIZE         = 0x0083,
    WM_NCCREATE           = 0x0081,
    WM_NCDESTROY          = 0x0082,
    WM_NCHITTEST          = 0x0084,
    WM_NCLBUTTONDBLCLK    = 0x00A3,
    WM_NCLBUTTONDOWN      = 0x00A1,
    WM_NCLBUTTONUP        = 0x00A2,
    WM_NCMBUTTONDBLCLK    = 0x00A9,
    WM_NCMBUTTONDOWN      = 0x00A7,
    WM_NCMBUTTONUP        = 0x00A8,
    WM_NCMOUSEHOVER       = 0x02A0,
    WM_NCMOUSELEAVE       = 0x02A2,
    WM_NCMOUSEMOVE        = 0x00A0,
    WM_NCPAINT        = 0x0085,
    WM_NCRBUTTONDBLCLK    = 0x00A6,
    WM_NCRBUTTONDOWN      = 0x00A4,
    WM_NCRBUTTONUP        = 0x00A5,
    WM_NCXBUTTONDBLCLK    = 0x00AD,
    WM_NCXBUTTONDOWN      = 0x00AB,
    WM_NCXBUTTONUP        = 0x00AC,
    WM_NCUAHDRAWCAPTION       = 0x00AE,
    WM_NCUAHDRAWFRAME     = 0x00AF,
    WM_NEXTDLGCTL         = 0x0028,
    WM_NEXTMENU           = 0x0213,
    WM_NOTIFY         = 0x004E,
    WM_NOTIFYFORMAT       = 0x0055,
    WM_NULL           = 0x0000,
    WM_PAINT          = 0x000F,
    WM_PAINTCLIPBOARD     = 0x0309,
    WM_PAINTICON          = 0x0026,
    WM_PALETTECHANGED     = 0x0311,
    WM_PALETTEISCHANGING      = 0x0310,
    WM_PARENTNOTIFY       = 0x0210,
    WM_PASTE          = 0x0302,
    WM_PENWINFIRST        = 0x0380,
    WM_PENWINLAST         = 0x038F,
    WM_POWER          = 0x0048,
    WM_POWERBROADCAST     = 0x0218,
    WM_PRINT          = 0x0317,
    WM_PRINTCLIENT        = 0x0318,
    WM_QUERYDRAGICON      = 0x0037,
    WM_QUERYENDSESSION    = 0x0011,
    WM_QUERYNEWPALETTE    = 0x030F,
    WM_QUERYOPEN          = 0x0013,
    WM_QUEUESYNC          = 0x0023,
    WM_QUIT           = 0x0012,
    WM_RBUTTONDBLCLK      = 0x0206,
    WM_RBUTTONDOWN        = 0x0204,
    WM_RBUTTONUP          = 0x0205,
    WM_RENDERALLFORMATS       = 0x0306,
    WM_RENDERFORMAT       = 0x0305,
    WM_SETCURSOR          = 0x0020,
    WM_SETFOCUS           = 0x0007,
    WM_SETFONT        = 0x0030,
    WM_SETHOTKEY          = 0x0032,
    WM_SETICON        = 0x0080,
    WM_SETREDRAW          = 0x000B,
    WM_SETTEXT        = 0x000C,
    WM_SETTINGCHANGE      = 0x001A,
    WM_SHOWWINDOW         = 0x0018,
    WM_SIZE           = 0x0005,
    WM_SIZECLIPBOARD      = 0x030B,
    WM_SIZING         = 0x0214,
    WM_SPOOLERSTATUS      = 0x002A,
    WM_STYLECHANGED       = 0x007D,
    WM_STYLECHANGING      = 0x007C,
    WM_SYNCPAINT          = 0x0088,
    WM_SYSCHAR        = 0x0106,
    WM_SYSCOLORCHANGE     = 0x0015,
    WM_SYSCOMMAND         = 0x0112,
    WM_SYSDEADCHAR        = 0x0107,
    WM_SYSKEYDOWN         = 0x0104,
    WM_SYSKEYUP           = 0x0105,
    WM_TCARD          = 0x0052,
    WM_TIMECHANGE         = 0x001E,
    WM_TIMER          = 0x0113,
    WM_UNDO           = 0x0304,
	WM_UNICHAR        = 0x0109,
    WM_UNINITMENUPOPUP    = 0x0125,
    WM_USER           = 0x0400,
    WM_USERCHANGED        = 0x0054,
    WM_VKEYTOITEM         = 0x002E,
    WM_VSCROLL        = 0x0115,
    WM_VSCROLLCLIPBOARD       = 0x030A,
    WM_WINDOWPOSCHANGED       = 0x0047,
    WM_WINDOWPOSCHANGING      = 0x0046,
    WM_WININICHANGE       = 0x001A,
    WM_XBUTTONDBLCLK      = 0x020D,
    WM_XBUTTONDOWN        = 0x020B,
    WM_XBUTTONUP          = 0x020C,
} WindowMessages;
typedef enum WindowStyles {
    WS_OVERLAPPED      = 0x00000000,
    WS_POPUP           = 0x80000000,
    WS_CHILD           = 0x40000000,
    WS_MINIMIZE        = 0x20000000,
    WS_VISIBLE         = 0x10000000,
    WS_DISABLED        = 0x08000000,
    WS_CLIPSIBLINGS    = 0x04000000,
    WS_CLIPCHILDREN    = 0x02000000,
    WS_MAXIMIZE        = 0x01000000,
    WS_BORDER          = 0x00800000,
    WS_DLGFRAME        = 0x00400000,
    WS_VSCROLL         = 0x00200000,
    WS_HSCROLL         = 0x00100000,
    WS_SYSMENU         = 0x00080000,
    WS_THICKFRAME      = 0x00040000,
    WS_GROUP           = 0x00020000,
    WS_TABSTOP         = 0x00010000,

    WS_MINIMIZEBOX     = 0x00020000,
    WS_MAXIMIZEBOX     = 0x00010000,

    WS_CAPTION         = 0x00C00000, //(WS_BORDER | WS_DLGFRAME),
    WS_TILED           = 0x00000000, //WS_OVERLAPPED,
    WS_ICONIC          = 0x20000000, //WS_MINIMIZE,
    WS_SIZEBOX         = 0x00040000, //WS_THICKFRAME,
    WS_TILEDWINDOW     = 0x00CF0000, //WS_OVERLAPPEDWINDOW,

    WS_OVERLAPPEDWINDOW    = 0x00CF0000, //WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX,
    WS_POPUPWINDOW     = 0x80880000, //WS_POPUP | WS_BORDER | WS_SYSMENU,
    WS_CHILDWINDOW     = 0x40000000, //WS_CHILD
} WindowStyles;
    //Extended Window Styles
typedef enum WindowStylesEx {
    WS_EX_DLGMODALFRAME    = 0x00000001,
    WS_EX_NOPARENTNOTIFY   = 0x00000004,
    WS_EX_TOPMOST          = 0x00000008,
    WS_EX_ACCEPTFILES      = 0x00000010,
    WS_EX_TRANSPARENT      = 0x00000020,
    WS_EX_MDICHILD         = 0x00000040,
    WS_EX_TOOLWINDOW       = 0x00000080,
    WS_EX_WINDOWEDGE       = 0x00000100,
    WS_EX_CLIENTEDGE       = 0x00000200,
    WS_EX_CONTEXTHELP      = 0x00000400,

    WS_EX_RIGHT            = 0x00001000,
    WS_EX_LEFT             = 0x00000000,
    WS_EX_RTLREADING       = 0x00002000,
    WS_EX_LTRREADING       = 0x00000000,
    WS_EX_LEFTSCROLLBAR    = 0x00004000,
    WS_EX_RIGHTSCROLLBAR   = 0x00000000,

    WS_EX_CONTROLPARENT    = 0x00010000,
    WS_EX_STATICEDGE       = 0x00020000,
    WS_EX_APPWINDOW        = 0x00040000,

    WS_EX_OVERLAPPEDWINDOW = 0x00000300, //(WS_EX_WINDOWEDGE | WS_EX_CLIENTEDGE),
    WS_EX_PALETTEWINDOW    = 0x00000188, //(WS_EX_WINDOWEDGE | WS_EX_TOOLWINDOW | WS_EX_TOPMOST),
    WS_EX_LAYERED          = 0x00080000,       
    WS_EX_NOINHERITLAYOUT  = 0x00100000, // Disable inheritence of mirroring by children
    WS_EX_LAYOUTRTL        = 0x00400000, // Right to left mirroring
    WS_EX_COMPOSITED       = 0x02000000,
    WS_EX_NOACTIVATE       = 0x08000000
 } WindowStylesEx;
        

typedef enum HRESULTS {
    S_OK = 0,
    S_FALSE = 1,

	NOERROR = - 0,
	
    E_NOTIMPL = 0x80004001,
    E_OUTOFMEMORY = 0x8007000E,
    E_INVALIDARG = 0x80070057,
    E_NOINTERFACE = 0x80004002,
    E_POINTER = 0x80004003,
    E_HANDLE = 0x80070006,
    E_ABORT = 0x80004004,
    E_FAIL = 0x80004005,
	E_UNSPEC = 0x80004005, //E_FAIL
	E_UNEXPECTED = 0x8000FFFF,
    E_ACCESSDENIED = 0x80070005,
	E_PENDING = 0x8000000A,
	E_BOUNDS = 0x8000000B,
	E_CHANGED_STATE = 0x8000000C,
	E_ILLEGAL_STATE_CHANGE = 0x8000000D,
	E_ILLEGAL_METHOD_CALL = 0x8000000E,
    // IConnectionPoint errors
        
    CONNECT_E_FIRST = 0x80040200,
    CONNECT_E_NOCONNECTION,  // there is no connection for this connection id
    CONNECT_E_ADVISELIMIT,   // this implementation's limit for advisory connections has been reached
    CONNECT_E_CANNOTCONNECT, // connection attempt failed
    CONNECT_E_OVERRIDDEN,    // must use a derived interface to connect
    
    // DllRegisterServer/DllUnregisterServer errors
    SELFREG_E_TYPELIB = 0x80040200, // failed to register/unregister type library
    SELFREG_E_CLASS,                // failed to register/unregister class

    // IPerPropertyBrowsing status codes

    PERPROP_E_NOPAGEAVAILABLE = 0x80040200, // no page available for requested property

    // INET errors
    
    INET_E_INVALID_URL = 0x800C0002,
    INET_E_NO_SESSION = 0x800C0003,
    INET_E_CANNOT_CONNECT = 0x800C0004,
    INET_E_RESOURCE_NOT_FOUND = 0x800C0005,
    INET_E_OBJECT_NOT_FOUND = 0x800C0006,
    INET_E_DATA_NOT_AVAILABLE = 0x800C0007,
    INET_E_DOWNLOAD_FAILURE = 0x800C0008,
    INET_E_AUTHENTICATION_REQUIRED = 0x800C0009,
    INET_E_NO_VALID_MEDIA = 0x800C000A,
    INET_E_CONNECTION_TIMEOUT = 0x800C000B,
    INET_E_INVALID_REQUEST = 0x800C000C,
    INET_E_UNKNOWN_PROTOCOL = 0x800C000D,
    INET_E_SECURITY_PROBLEM = 0x800C000E,
    INET_E_CANNOT_LOAD_DATA = 0x800C000F,
    INET_E_CANNOT_INSTANTIATE_OBJECT = 0x800C0010,
    INET_E_USE_DEFAULT_PROTOCOLHANDLER = 0x800C0011,
    INET_E_DEFAULT_ACTION = 0x800C0011,
    INET_E_USE_DEFAULT_SETTING = 0x800C0012,
    INET_E_QUERYOPTION_UNKNOWN = 0x800C0013,
    INET_E_REDIRECTING = 0x800C0014,
    INET_E_REDIRECT_FAILED = 0x800C0014,
    INET_E_REDIRECT_TO_DIR = 0x800C0015,
    INET_E_CANNOT_LOCK_REQUEST = 0x800C0016,
    INET_E_USE_EXTEND_BINDING = 0x800C0017,
    INET_E_ERROR_FIRST = 0x800C0002,
    INET_E_ERROR_LAST = 0x800C0017,
    INET_E_CODE_DOWNLOAD_DECLINED = 0x800C0100,
    INET_E_RESULT_DISPATCHED = 0x800C0200,
    INET_E_CANNOT_REPLACE_SFP_FILE = 0x800C0300,

} HRESULTS;



typedef enum VARENUM {
    VT_EMPTY = 0,
    VT_NULL = 1,
    VT_I2 = 2,
    VT_I4 = 3,
    VT_R4 = 4,
    VT_R8 = 5,
    VT_CY = 6,
    VT_DATE = 7,
    VT_BSTR = 8,
    VT_DISPATCH = 9,
    VT_ERROR = 10,
    VT_BOOL = 11,
    VT_VARIANT = 12,
    VT_UNKNOWN = 13,
    VT_DECIMAL = 14,
    VT_I1 = 16,
    VT_UI1 = 17,
    VT_UI2 = 18,
    VT_UI4 = 19,
    VT_I8 = 20,
    VT_UI8 = 21,
    VT_INT = 22,
    VT_UINT = 23,
    VT_VOID = 24,
    VT_HRESULT = 25,
    VT_PTR = 26,
    VT_SAFEARRAY = 27,
    VT_CARRAY = 28,
    VT_USERDEFINED = 29,
    VT_LPSTR = 30,
    VT_LPWSTR = 31,
	VT_RECORD = 36,
	VT_INT_PTR = 37,
	VT_UINT_PTR = 38,
    VT_FILETIME = 64,
    VT_BLOB = 65,
    VT_STREAM = 66,
    VT_STORAGE = 67,
    VT_STREAMED_OBJECT = 68,
    VT_STORED_OBJECT = 69,
    VT_BLOB_OBJECT = 70,
    VT_CF = 71,
    VT_CLSID = 72,
    VT_STREAMED_PROPSET = 73,
    VT_STORED_PROPSET = 74,
    VT_BLOB_PROPSET = 75,
    VT_VERBOSE_ENUM = 76,
    VT_BSTR_BLOB = 0xFFF,
    VT_VECTOR = 0x1000,
    VT_ARRAY = 0x2000,
    VT_BYREF = 0x4000,
    VT_RESERVED = 0x8000,
    VT_ILLEGAL = 0xffff,
    VT_ILLEGALMASKED = 0x0fff,
    VT_TYPEMASK = 0x0fff
} VARENUM;

typedef enum SummaryInfoProperties {
    [helpstring("Document Title - VT_LPSTR")]
    PIDSI_TITLE = 0x00000002,
    [helpstring("Document Subject - VT_LPSTR")]
    PIDSI_SUBJECT = 0x00000003,
    [helpstring("Document Author - VT_LPSTR")]
    PIDSI_AUTHOR = 0x00000004,
    [helpstring("Keywords - VT_LPSTR")]
    PIDSI_KEYWORDS = 0x00000005,
    [helpstring("Comments - VT_LPSTR")]
    PIDSI_COMMENTS = 0x00000006,
    [helpstring("Template name - VT_LPSTR")]
    PIDSI_TEMPLATE = 0x00000007,
    [helpstring("Last Saved By - VT_LPSTR")]
    PIDSI_LASTAUTHOR = 0x00000008,
    [helpstring("Revision Number (save count) - VT_LPSTR")]
    PIDSI_REVNUMBER = 0x00000009,
    [helpstring("Total editing time - VT_CURRENCY (UTC)")]
    PIDSI_EDITTIME = 0x0000000a,
    [helpstring("Last printed date - VT_CURRENCY (UTC)")]
    PIDSI_LASTPRINTED = 0x0000000b,
    [helpstring("Create Date - VT_CURRENCY (UTC)")]
    PIDSI_CREATE_DTM = 0x0000000c,
    [helpstring("Last save date - VT_CURRENCY (UTC)")]
    PIDSI_LASTSAVE_DTM = 0x0000000d,
    [helpstring("Page count - VT_I4")]
    PIDSI_PAGECOUNT = 0x0000000e,
    [helpstring("Word count - VT_I4")]
    PIDSI_WORDCOUNT = 0x0000000f,
    [helpstring("Char count - VT_I4")]
    PIDSI_CHARCOUNT = 0x00000010,
    [helpstring("Thumbnail - VT_CF")]
    PIDSI_THUMBNAIL = 0x00000011,
    [helpstring("Application Name - VT_LPSTR")]
    PIDSI_APPNAME = 0x00000012,
    [helpstring("Document security - VT_I4")]
    PIDSI_DOC_SECURITY = 0x00000013,
} SummaryInfoProperties;

typedef enum DocSummaryInfoProperties {
    [helpstring("Category - LPSTR")]
    PID_CATEGORY = 0x2,
    [helpstring("Presentation Target - LPSTR")]
    PID_PRESFORMAT = 0x3,
    [helpstring("Byte count - VT_I4")]
    PID_BYTECOUNT = 0x4,
    [helpstring("Line count - VT_I4")]
    PID_LINECOUNT = 0x5,
    [helpstring("Paragraph count - VT_I4")]
    PID_PARCOUNT = 0x6,
    [helpstring("Slide count - VT_I4")]
    PID_SLIDECOUNT = 0x7,
    [helpstring("Note count - VT_I4")]
    PID_NOTECOUNT = 0x8,
    [helpstring("Hidden slide count - VT_I4")]
    PID_HIDDENCOUNT = 0x9,
    [helpstring("MMClip count - VT_I4")]
    PID_MMCLIPCOUNT = 0xA,
    [helpstring("Scale/Crop - VT_BOOL")]
    PID_SCALE = 0xB,
    [helpstring("Heading Pairs - VT_VARIANT | VT_VECTOR")]
    PID_HEADINGPAIR = 0xC,
    [helpstring("Titles Of Parts - VT_LPSTR | VT_VECTOR")]
    PID_DOCPARTS = 0xD,
    [helpstring("Manager - VT_LPSTR")]
    PID_MANAGER = 0xE,
    [helpstring("Company - VT_LPSTR")]
    PID_COMPANY = 0xF,
    [helpstring("LinksUpToDate - VT_BOOL")]
    PID_LINKSDIRTY = 0x10,
} DocSummaryInfoProperties;

/* Storage instantiation modes */
typedef enum STGM {
    STGM_DIRECT = 0x00000000,
    STGM_TRANSACTED = 0x00010000,
    STGM_SIMPLE = 0x08000000,

    STGM_READ = 0x00000000,
    STGM_WRITE = 0x00000001,
    STGM_READWRITE = 0x00000002,

    STGM_SHARE_DENY_NONE = 0x00000040,
    STGM_SHARE_DENY_READ = 0x00000030,
    STGM_SHARE_DENY_WRITE = 0x00000020,
    STGM_SHARE_EXCLUSIVE = 0x00000010,

    STGM_PRIORITY = 0x00040000,
    STGM_DELETEONRELEASE = 0x04000000,
    STGM_NOSCRATCH = 0x00100000,

    STGM_CREATE = 0x00001000,
    STGM_CONVERT = 0x00020000,
    STGM_FAILIFTHERE = 0x00000000,

    STGM_NOSNAPSHOT = 0x00200000,
    STGM_DIRECT_SWMR = 0x00400000,
} STGM;


typedef struct SYSTEMTIME {
    short wYear;
    short wMonth;
    short wDayOfWeek;
    short wDay;
    short wHour;
    short wMinute;
    short wSecond;
    short wMilliseconds;
} SYSTEMTIME;

typedef enum STGC {
     STGC_DEFAULT = 0,
     STGC_OVERWRITE = 1,
     STGC_ONLYIFCURRENT = 2,
     STGC_DANGEROUSLYCOMMITMERELYTODISKCACHE = 4,
     STGC_CONSOLIDATE = 8
} STGC;

typedef enum STATFLAG {
    STATFLAG_DEFAULT = 0,
    STATFLAG_NONAME = 1,
    STATFLAG_NOOPEN = 2
} STATFLAG;

typedef enum STGTY {
    STGTY_STORAGE = 1,
    STGTY_STREAM = 2,
    STGTY_LOCKBYTES = 3,
    STGTY_PROPERTY = 4
} STGTY;

typedef enum LOCKTYPE {
    LOCK_WRITE = 1,
    LOCK_EXCLUSIVE = 2,
    LOCK_ONLYONCE = 4
} LOCKTYPE;

typedef struct STATSTG {
    LONG pwcsName;
    STGTY type;
    CURRENCY cbSize;
    CURRENCY mtime;
    CURRENCY ctime;
    CURRENCY atime;
    STGM grfMode;
    LOCKTYPE grfLocksSupported;
    UUID clsid;
    LONG grfStateBits;
    LONG reserved;
} STATSTG;

typedef enum STREAM_SEEK {
    STREAM_SEEK_SET = 0,
    STREAM_SEEK_CUR = 1,
    STREAM_SEEK_END = 2
} STREAM_SEEK;

typedef enum STGMOVE {
    STGMOVE_MOVE = 0,
    STGMOVE_COPY = 1,
    STGMOVE_SHALLOWCOPY = 2
} STGMOVE;

typedef enum OLEGETMONIKER {
    OLEGETMONIKER_ONLYIFTHERE = 1,
    OLEGETMONIKER_FORCEASSIGN = 2,
    OLEGETMONIKER_UNASSIGN = 3,
    OLEGETMONIKER_TEMPFORUSER = 4
} OLEGETMONIKER;

typedef enum OLEWHICHMK {
    OLEWHICHMK_CONTAINER = 1,
    OLEWHICHMK_OBJREL = 2,
    OLEWHICHMK_OBJFULL = 3
} OLEWHICHMK;

typedef enum OLECONTF {
    OLECONTF_EMBEDDINGS = 1,
    OLECONTF_LINKS = 2,
    OLECONTF_OTHERS = 4,
    OLECONTF_ONLYUSER = 8,
    OLECONTF_ONLYIFRUNNING = 16
} OLECONTF;

typedef enum OLECLOSE {
    OLECLOSE_SAVEIFDIRTY = 0,
    OLECLOSE_NOSAVE = 1,
    OLECLOSE_PROMPTSAVE = 2
} OLECLOSE;

typedef enum USERCLASSTYPE {
    USERCLASSTYPE_FULL = 1,
    USERCLASSTYPE_SHORT = 2,
    USERCLASSTYPE_APPNAME = 3
} USERCLASSTYPE;

typedef enum OLEVERBATTRIB {
    OLEVERBATTRIB_NEVERDIRTIES = 1,
    OLEVERBATTRIB_ONCONTAINERMENU = 2
} OLEVERBATTRIB;

typedef enum DVASPECT {
    DVASPECT_CONTENT = 1,
    DVASPECT_THUMBNAIL = 2,
    DVASPECT_ICON = 4,
    DVASPECT_DOCPRINT = 8,
    DVASPECT_OPAQUE = 16,
    DVASPECT_TRANSPARENT = 32
} DVASPECT;

typedef enum TYMED {
    TYMED_HGLOBAL = 1,
    TYMED_FILE = 2,
    TYMED_ISTREAM = 4,
    TYMED_ISTORAGE = 8,
    TYMED_GDI = 16,
    TYMED_MFPICT = 32,
    TYMED_ENHMF = 64,
    TYMED_NULL = 0
} TYMED;

typedef enum OLEMISC {
    OLEMISC_RECOMPOSEONRESIZE = 1,
    OLEMISC_ONLYICONIC = 2,
    OLEMISC_INSERTNOTREPLACE = 4,
    OLEMISC_STATIC = 8,
    OLEMISC_CANTLINKINSIDE = 16,
    OLEMISC_CANLINKBYOLE1 = 32,
    OLEMISC_ISLINKOBJECT = 64,
    OLEMISC_INSIDEOUT = 128,
    OLEMISC_ACTIVATEWHENVISIBLE = 256,
    OLEMISC_RENDERINGISDEVICEINDEPENDENT = 512,
    OLEMISC_INVISIBLEATRUNTIME = 1024,
    OLEMISC_ALWAYSRUN = 2048,
    OLEMISC_ACTSLIKEBUTTON = 4096,
    OLEMISC_ACTSLIKELABEL = 8192,
    OLEMISC_NOUIACTIVATE = 16384,
    OLEMISC_ALIGNABLE = 32768,
    OLEMISC_SIMPLEFRAME = 0x00010000,
    OLEMISC_SETCLIENTSITEFIRST = 0x00020000,
    OLEMISC_IMEMODE = 0x00040000,
    OLEMISC_IGNOREACTIVATEWHENVISIBLE = 0x00080000,
    OLEMISC_WANTSTOMENUMERGE = 0x00100000,
    OLEMISC_SUPPORTSMULTILEVELUNDO = 0x00200000
} OLEMISC;

typedef enum DVASPECTINFOFLAG {
    DVASPECTINFOFLAG_CANOPTIMIZE = 1
} DVASPECTINFOFLAG;

typedef enum DATADIR {
    DATADIR_GET = 1,
    DATADIR_SET = 2
} DATADIR;

typedef enum OLEVERBS {
    OLEIVERB_PRIMARY = 0,
    OLEIVERB_SHOW = -1,
    OLEIVERB_OPEN = -2,
    OLEIVERB_HIDE = -3,
    OLEIVERB_UIACTIVATE = -4,
    OLEIVERB_INPLACEACTIVATE = -5,
    OLEIVERB_DISCARDUNDOSTATE = -6,
    OLEIVERB_PROPERTIES = -7
} OLEVERBS;

typedef enum ADVF {
    ADVF_NODATA = 1,
    ADVF_PRIMEFIRST = 2,
    ADVF_ONLYONCE = 4,
    ADVF_DATAONSTOP = 64,
    ADVFCACHE_NOHANDLER = 8,
    ADVFCACHE_FORCEBUILTIN = 16,
    ADVFCACHE_ONSAVE = 32
} ADVF;

typedef struct OLEVERB {
    long lVerb;
    long lpszVerbName;
    long fuFlags;
    OLEVERBATTRIB grfAttribs;
} OLEVERB;

typedef struct OLEMENUGROUPWIDTHS {
    LONG width[6];
} OLEMENUGROUPWIDTHS;

typedef struct RECT {
    long Left;
    long Top;
    long Right;
    long Bottom;
} RECT;

typedef struct SIZE {
    long cx;
    long cy;
} SIZE;

typedef struct STGMEDIUM {
    TYMED TYMED;
    long data;
    long pUnkForRelease;
} STGMEDIUM;

typedef struct MSG {
    long hwnd;
    long message;
    long wParam;
    long lParam;
    long time;
    long pt_x;
    long pt_y;
} MSG;

typedef struct DVTARGETDEVICE {
    long tdSize;
    short tdDriverNameOffset;
    short tdDeviceNameOffset;
    short tdPortNameOffset;
    short tdExtDevmodeOffset;
    [helpstring("size_is(tdSize - sizeof(long) - 4 * sizeof(SHORT))")]
    BYTE tdData[1];
} DVTARGETDEVICE;

typedef struct FORMATETC {
    long cfFormat;
    long pDVTARGETDEVICE;
    DVASPECT dwAspect;
    long lindex;
    TYMED TYMED;
} FORMATETC;

typedef struct STATDATA {
    FORMATETC FORMATETC;
    ADVF grfAdvf;
    IAdviseSink *pIAdviseSink;
    long dwConnection;
} STATDATA;

typedef struct PALETTEENTRY {
    BYTE peRed;
    BYTE peGreen;
    BYTE peBlue;
    BYTE peFlags;
} PALETTEENTRY;

typedef struct LOGPALETTE {
    short palVersion;
    short palNumEntries;
    PALETTEENTRY* palPalEntry;
} LOGPALETTE;

typedef struct DVASPECTINFO {
    long cb;
    DVASPECTINFOFLAG dwFlags;
} DVASPECTINFO;

typedef struct OLEINPLACEFRAMEINFO {
    long cb;
    long fMDIApp;
    long hwndFrame;
    long haccel;
    long cAccelEntries;
} OLEINPLACEFRAMEINFO;

typedef struct CSPLATFORM {
    LONG dwPlatformId;
    LONG dwVersionHi;
    LONG dwVersionLo;
    LONG dwProcessorArch;
} CSPLATFORM;

typedef struct QUERYCONTEXT {
    LONG dwContext;
    CSPLATFORM Platform;
    LONG Locale;
    LONG dwVersionHi;
    LONG dwVersionLo;
} QUERYCONTEXT;

typedef enum TYSPEC {
    TYSPEC_CLSID = 0,
    TYSPEC_FILEEXT,
    TYSPEC_MIMETYPE,
    TYSPEC_FILENAME,
    TYSPEC_PROGID,
    TYSPEC_PACKAGENAME,
    TYSPEC_OBJECTID
} TYSPEC;

typedef struct uCLSSPEC {
    TYSPEC tyspec;
    BYTE Data[32];
} uCLSSPEC;

// flags for FaultInIEFeature
typedef enum FIEF {
    FIEF_FLAG_FORCE_JITUI = 0x1,  // force JIT ui even if
                                  // previoulsy rejected by
                                  // user in this session or
                                  // marked as Never Ask Again
    FIEF_FLAG_PEEK = 0x2,         // just peek, don't faultin
    FIEF_FLAG_SKIP_INSTALLED_VERSION_CHECK = 0x4 // force JIT without checking local version

} FIEF;

typedef enum BINDVERB {
    BINDVERB_GET    = 0x00000000,       // default action
    BINDVERB_POST   = 0x00000001,       // post verb
    BINDVERB_PUT    = 0x00000002,       // put verb
    BINDVERB_CUSTOM = 0x00000003,       // custom verb
} BINDVERB;

typedef enum BINDINFOF {
    BINDINFOF_URLENCODESTGMEDDATA = 0x00000001,       // url encode the stgmed data
    BINDINFOF_URLENCODEDEXTRAINFO = 0x00000002,       // url encode the extra info
} BINDINFOF;

// flags that describe the type of transaction that caller wants
typedef enum BINDF {
	BINDF_ASYNCHRONOUS             = 0x00000001,
    BINDF_ASYNCSTORAGE             = 0x00000002,
    BINDF_NOPROGRESSIVERENDERING   = 0x00000004,
    BINDF_OFFLINEOPERATION         = 0x00000008,
    BINDF_GETNEWESTVERSION         = 0x00000010,
    BINDF_NOWRITECACHE             = 0x00000020,
    BINDF_NEEDFILE                 = 0x00000040,
    BINDF_PULLDATA                 = 0x00000080,
    BINDF_IGNORESECURITYPROBLEM    = 0x00000100,
    BINDF_RESYNCHRONIZE            = 0x00000200,
    BINDF_HYPERLINK                = 0x00000400,
    BINDF_NO_UI                    = 0x00000800,
    BINDF_SILENTOPERATION          = 0x00001000,
    BINDF_PRAGMA_NO_CACHE          = 0x00002000,
    BINDF_GETCLASSOBJECT           = 0x00004000,
    BINDF_RESERVED_1               = 0x00008000,
    BINDF_FREE_THREADED            = 0x00010000,
    BINDF_DIRECT_READ              = 0x00020000,
    BINDF_FORMS_SUBMIT             = 0x00040000,
    BINDF_GETFROMCACHE_IF_NET_FAIL = 0x00080000,
    BINDF_FROMURLMON               = 0x00100000,
    BINDF_FWD_BACK                 = 0x00200000,
    BINDF_PREFERDEFAULTHANDLER     = 0x00400000,
    BINDF_ENFORCERESTRICTED        = 0x00800000,
} BINDF;

typedef enum URL_ENCODING {
    URL_ENCODING_NONE = 0x00000000,
    URL_ENCODING_ENABLE_UTF8 = 0x10000000,
    URL_ENCODING_DISABLE_UTF8 = 0x20000000,
} URL_ENCODING;

typedef struct SECURITY_ATTRIBUTES {
    LONG nLength;
    LONG lpSecurityDescriptor;
    BOOL bInheritHandle;
} SECURITY_ATTRIBUTES;

typedef struct BINDINFO {
    LONG cbSize;
    LONG szExtraInfo;        // e.g. string added to the URL string [range byte]
    STGMEDIUM stgmedData;    // e.g. put or post data
    LONG grfBindInfoF;       // e.g. how meddata should be treaded
    LONG dwBindVerb;         // e.g. put, post, ...
    LONG szCustomVerb;       // e.g. custom verb
    LONG cbstgmedData;       // size of data in stgmedData
    LONG dwOptions;          // additional option
    LONG dwOptionsFlags;     // options flags
    LONG dwCodePage;         // the code used to convert the url
    SECURITY_ATTRIBUTES securityAttributes;
    UUID iid;                // iid of pUnk
    stdole.IUnknown*pUnk;          // pUnk - whatever
    LONG dwReserved;         // should be null
} BINDINFO;


// Options flag
typedef enum BINDINFO_OPTIONS {
  BINDINFO_OPTIONS_WININETFLAG     = 0x00010000,
  BINDINFO_OPTIONS_ENABLE_UTF8     = 0x00020000,
  BINDINFO_OPTIONS_DISABLE_UTF8    = 0x00040000,
  BINDINFO_OPTIONS_USE_IE_ENCODING = 0x00080000,
  BINDINFO_OPTIONS_BINDTOOBJECT    = 0x00100000,
  BINDINFO_OPTIONS_SECURITYOPTOUT            = 0x00200000,
  BINDINFO_OPTIONS_IGNOREMIMETEXTPLAIN       = 0x00400000,
  BINDINFO_OPTIONS_USEBINDSTRINGCREDS        = 0x00800000,
  BINDINFO_OPTIONS_IGNOREHTTPHTTPSREDIRECTS  = 0x01000000,
  BINDINFO_OPTIONS_IGNORE_SSLERRORS_ONCE     = 0x02000000,
  BINDINFO_WPC_DOWNLOADBLOCKED               = 0x08000000,
  BINDINFO_WPC_LOGGING_ENABLED               = 0x10000000,
  BINDINFO_OPTIONS_DISABLEAUTOREDIRECTS      = 0x40000000,
  BINDINFO_OPTIONS_SHDOCVW_NAVIGATE          = 0x80000000,
  BINDINFO_OPTIONS_ALLOWCONNECTDATA          = 0x20000000,
} BINDINFO_OPTIONS;

// Data notification flags
typedef enum BSCF {
    BSCF_FIRSTDATANOTIFICATION        = 0x00000001,
    BSCF_INTERMEDIATEDATANOTIFICATION = 0x00000002,
    BSCF_LASTDATANOTIFICATION         = 0x00000004,
    BSCF_DATAFULLYAVAILABLE           = 0x00000008,
    BSCF_AVAILABLEDATASIZEUNKNOWN     = 0x00000010,
    BSCF_SKIPDRAINDATAFORFILEURLS     = 0x00000020,
    BSCF_64BITLENGTHDOWNLOAD          = 0x00000040
} BSCF;

typedef enum BINDSTATUS {
    BINDSTATUS_FINDINGRESOURCE = 1,
    BINDSTATUS_CONNECTING,
    BINDSTATUS_REDIRECTING,
    BINDSTATUS_BEGINDOWNLOADDATA,
    BINDSTATUS_DOWNLOADINGDATA,
    BINDSTATUS_ENDDOWNLOADDATA,
    BINDSTATUS_BEGINDOWNLOADCOMPONENTS,
    BINDSTATUS_INSTALLINGCOMPONENTS,
    BINDSTATUS_ENDDOWNLOADCOMPONENTS,
    BINDSTATUS_USINGCACHEDCOPY,
    BINDSTATUS_SENDINGREQUEST,
    BINDSTATUS_CLASSIDAVAILABLE,
    BINDSTATUS_MIMETYPEAVAILABLE,
    BINDSTATUS_CACHEFILENAMEAVAILABLE,
    BINDSTATUS_BEGINSYNCOPERATION,
    BINDSTATUS_ENDSYNCOPERATION,
    BINDSTATUS_BEGINUPLOADDATA,
    BINDSTATUS_UPLOADINGDATA,
    BINDSTATUS_ENDUPLOADDATA,
    BINDSTATUS_PROTOCOLCLASSID,
    BINDSTATUS_ENCODING,
    BINDSTATUS_VERIFIEDMIMETYPEAVAILABLE,
    BINDSTATUS_CLASSINSTALLLOCATION,
    BINDSTATUS_DECODING,
    BINDSTATUS_LOADINGMIMEHANDLER,
    BINDSTATUS_CONTENTDISPOSITIONATTACH,
    BINDSTATUS_FILTERREPORTMIMETYPE,
    BINDSTATUS_CLSIDCANINSTANTIATE,
    BINDSTATUS_IUNKNOWNAVAILABLE,
    BINDSTATUS_DIRECTBIND,
    BINDSTATUS_RAWMIMETYPE,
    BINDSTATUS_PROXYDETECTING,
    BINDSTATUS_ACCEPTRANGES,
    BINDSTATUS_COOKIE_SENT,
    BINDSTATUS_COMPACT_POLICY_RECEIVED,
    BINDSTATUS_COOKIE_SUPPRESSED,
    BINDSTATUS_COOKIE_STATE_UNKNOWN,
    BINDSTATUS_COOKIE_STATE_ACCEPT,
    BINDSTATUS_COOKIE_STATE_REJECT,
    BINDSTATUS_COOKIE_STATE_PROMPT,
    BINDSTATUS_COOKIE_STATE_LEASH,
    BINDSTATUS_COOKIE_STATE_DOWNGRADE,
    BINDSTATUS_POLICY_HREF,
    BINDSTATUS_P3P_HEADER,
    BINDSTATUS_SESSION_COOKIE_RECEIVED,
    BINDSTATUS_PERSISTENT_COOKIE_RECEIVED,
    BINDSTATUS_SESSION_COOKIES_ALLOWED,
  BINDSTATUS_CACHECONTROL,
  BINDSTATUS_CONTENTDISPOSITIONFILENAME,
  BINDSTATUS_MIMETEXTPLAINMISMATCH,
  BINDSTATUS_PUBLISHERAVAILABLE,
  BINDSTATUS_DISPLAYNAMEAVAILABLE,
  BINDSTATUS_SSLUX_NAVBLOCKED,
  BINDSTATUS_SERVER_MIMETYPEAVAILABLE,
  BINDSTATUS_SNIFFED_CLASSIDAVAILABLE,
  BINDSTATUS_64BIT_PROGRESS,
  //BINDSTATUS_LAST                        = BINDSTATUS_64BIT_PROGRESS,
  BINDSTATUS_RESERVED_0,
  BINDSTATUS_RESERVED_1,
  BINDSTATUS_RESERVED_2,
  BINDSTATUS_RESERVED_3,
  BINDSTATUS_RESERVED_4,
  BINDSTATUS_RESERVED_5,
  BINDSTATUS_RESERVED_6,
  BINDSTATUS_RESERVED_7,
  BINDSTATUS_RESERVED_8,
  BINDSTATUS_RESERVED_9,
  //BINDSTATUS_LAST_PRIVATE                = BINDSTATUS_RESERVED_9
} BINDSTATUS;

typedef enum ROTFLAGS {
    ROTFLAGS_REGISTRATIONKEEPSALIVE = 0x1,
    ROTFLAGS_ALLOWANYCLIENT = 0x2
} ROTFLAGS;

typedef enum DROPEFFECTS {
    DROPEFFECT_NONE = 0,
    DROPEFFECT_COPY = 1,
    DROPEFFECT_MOVE = 2,
    DROPEFFECT_LINK = 4,
    DROPEFFECT_SCROLL = 0x80000000,
} DROPEFFECTS;

typedef enum CLSCTX {
    CLSCTX_INPROC_SERVER = 0x1,
    CLSCTX_INPROC_HANDLER = 0x2,
    CLSCTX_LOCAL_SERVER = 0x4,
    CLSCTX_INPROC_SERVER16 = 0x8,
    CLSCTX_REMOTE_SERVER = 0x10,
    CLSCTX_INPROC_HANDLER16 = 0x20,
    CLSCTX_INPROC_SERVERX86 = 0x40,
    CLSCTX_INPROC_HANDLERX86 = 0x80,
    CLSCTX_ESERVER_HANDLER = 0x100,
    CLSCTX_RESERVED = 0x200,
    CLSCTX_NO_CODE_DOWNLOAD = 0x400,
    CLSCTX_NO_WX86_TRANSLATION = 0x800,
    CLSCTX_NO_CUSTOM_MARSHAL = 0x1000,
    CLSCTX_ENABLE_CODE_DOWNLOAD = 0x2000,
    CLSCTX_NO_FAILURE_LOG = 0x4000,
  CLSCTX_DISABLE_AAA             = 0x8000,
  CLSCTX_ENABLE_AAA              = 0x10000,
  CLSCTX_FROM_DEFAULT_CONTEXT    = 0x20000,
  CLSCTX_ACTIVATE_32_BIT_SERVER  = 0x40000,
  CLSCTX_ACTIVATE_64_BIT_SERVER  = 0x80000,
  CLSCTX_ENABLE_CLOAKING         = 0x100000,
  CLSCTX_APPCONTAINER            = 0x400000,
  CLSCTX_ACTIVATE_AAA_AS_IU      = 0x800000,
  CLSCTX_PS_DLL                  = 0x80000000
} CLSCTX;

typedef enum REGCLS {
    REGCLS_SINGLEUSE = 0,       // class object only generates one instance
    REGCLS_MULTIPLEUSE = 1,     // same class object genereates multiple inst.
                                // and local automatically goes into inproc tbl.
    REGCLS_MULTI_SEPARATE = 2,  // multiple use, but separate control over each
                                // context.
    REGCLS_SUSPENDED = 4,  // register is as suspended, will be activated
                                // when app calls CoResumeClassObjects
    REGCLS_SURROGATE = 8   // must be used when a surrogate process
                                // is registering a class object that will be
                                // loaded in the surrogate
} REGCLS;

typedef enum MEMCTX {
    MEMCTX_TASK = 1,
    MEMCTX_SHARED = 2,
    MEMCTX_MACSYSTEM = 3,
    MEMCTX_UNKNOWN = -1,
    MEMCTX_SAME = -2
} MEMCTX;

typedef enum MSHCTX {
    MSHCTX_LOCAL = 0,
    MSHCTX_NOSHAREDMEM = 1,
    MSHCTX_DIFFERENTMACHINE = 2,
    MSHCTX_INPROC = 3,
	MSHCTX_CROSSCTX          = 4
} MSHCTX;

typedef enum MSHLFLAGS {
    MSHLFLAGS_NORMAL = 0,
    MSHLFLAGS_TABLESTRONG = 1,
    MSHLFLAGS_TABLEWEAK = 2,
    MSHLFLAGS_NOPING = 3
} MSHLFLAGS;

typedef enum STDMSHLFLAGS {
    SMEXF_SERVER = 0x01,  //server side aggregated std marshaler
    SMEXF_HANDLER = 0x02  //client side (handler) agg std marshaler
} STDMSHLFLAGS;

typedef enum OLERENDER {
    OLERENDER_NONE = 0,
    OLERENDER_DRAW = 1,
    OLERENDER_FORMAT = 2,
    OLERENDER_ASIS = 3
} OLERENDER;

typedef enum OLECREATE_UDT {
    OLECREATE_LEAVERUNNING = 1,
} OLECREATE_UDT;


typedef struct POINT {
    LONG x;
    LONG y;
} POINT;


typedef struct DROPFILES {
    LONG pFiles; // offset of file list
    POINT pt;    // drop point (client coords)
    BOOL fNC;    // is it on NonClient area
                 // and pt is in screen coords
    BOOL fWide;  // WIDE character switch
} DROPFILES;

typedef struct  NETRESOURCE {
    LONG dwScope;
    LONG dwType;
    LONG dwDisplayType;
    LONG dwUsage;
    LPSTR lpLocalName;
    LPSTR lpRemoteName;
    LPSTR lpComment;
    LPSTR lpProvider;
}NETRESOURCE;

//
// format of CF_NETRESOURCE
//
typedef struct NRESARRAY {
    LONG cItems;
    NETRESOURCE nr[1];
} NRESARRAY;

//
// format of CF_IDLIST
//
typedef struct CIDA {
    LONG cidl;          // number of relative IDList
    LONG aoffset[1];    // [0]: folder IDList, [1]-[cidl]: item IDList
} CIDA;

//
// FILEDESCRIPTOR.dwFlags field indicate which fields are to be used
//
typedef enum FD_FLAGS {
    FD_CLSID = 0x0001,
    FD_SIZEPOINT = 0x0002,
    FD_ATTRIBUTES = 0x0004,
    FD_CREATETIME = 0x0008,
    FD_ACCESSTIME = 0x0010,
    FD_WRITESTIME = 0x0020,
    FD_FILESIZE = 0x0040,
    FD_PROGRESSUI = 0x4000,       // Show Progress UI w/Drag and Drop
    FD_LINKUI = 0x8000,       // 'link' UI is prefered
    FD_UNICODE = 0x80000000,
} FD_FLAGS;

typedef struct _FILEDESCRIPTORA {
    FD_FLAGS dwFlags;
    UUID clsid;
    SIZE sizel;
    POINT pointl;
    FILE_ATTRIBUTES dwFileAttributes;
    CURRENCY ftCreationTime;
    CURRENCY ftLastAccessTime;
    CURRENCY ftLastWriteTime;
    LONG nFileSizeHigh;
    LONG nFileSizeLow;
    BYTE cFileName[260];
} FILEDESCRIPTORA;

typedef struct _FILEDESCRIPTORW {
    FD_FLAGS dwFlags;
    UUID clsid;
    SIZE sizel;
    POINT pointl;
    FILE_ATTRIBUTES dwFileAttributes;
    CURRENCY ftCreationTime;
    CURRENCY ftLastAccessTime;
    CURRENCY ftLastWriteTime;
    LONG nFileSizeHigh;
    LONG nFileSizeLow;
    BYTE cFileName[520];
} FILEDESCRIPTORW;

typedef struct FILEGROUPDESCRIPTORA {
     LONG cItems;
     FILEDESCRIPTORA fgd[1];
} FILEGROUPDESCRIPTORA;

typedef struct FILEGROUPDESCRIPTORW {
     LONG cItems;
     FILEDESCRIPTORW fgd[1];
} FILEGROUPDESCRIPTORW;


    typedef enum WBIDM {

        IDM_UNKNOWN = 0,
        IDM_ALIGNBOTTOM = 1,
        IDM_ALIGNHORIZONTALCENTERS = 2,
        IDM_ALIGNLEFT = 3,
        IDM_ALIGNRIGHT = 4,
        IDM_ALIGNTOGRID = 5,
        IDM_ALIGNTOP = 6,
        IDM_ALIGNVERTICALCENTERS = 7,
        IDM_ARRANGEBOTTOM = 8,
        IDM_ARRANGERIGHT = 9,
        IDM_BRINGFORWARD = 10,
        IDM_BRINGTOFRONT = 11,
        IDM_CENTERHORIZONTALLY = 12,
        IDM_CENTERVERTICALLY = 13,
        IDM_CODE = 14,
        IDM_DELETE = 17,
        IDM_FONTNAME = 18,
        IDM_FONTSIZE = 19,
        IDM_GROUP = 20,
        IDM_HORIZSPACECONCATENATE = 21,
        IDM_HORIZSPACEDECREASE = 22,
        IDM_HORIZSPACEINCREASE = 23,
        IDM_HORIZSPACEMAKEEQUAL = 24,
        IDM_INSERTOBJECT = 25,
        IDM_MULTILEVELREDO = 30,
        IDM_SENDBACKWARD = 32,
        IDM_SENDTOBACK = 33,
        IDM_SHOWTABLE = 34,
        IDM_SIZETOCONTROL = 35,
        IDM_SIZETOCONTROLHEIGHT = 36,
        IDM_SIZETOCONTROLWIDTH = 37,
        IDM_SIZETOFIT = 38,
        IDM_SIZETOGRID = 39,
        IDM_SNAPTOGRID = 40,
        IDM_TABORDER = 41,
        IDM_TOOLBOX = 42,
        IDM_MULTILEVELUNDO = 44,
        IDM_UNGROUP = 45,
        IDM_VERTSPACECONCATENATE = 46,
        IDM_VERTSPACEDECREASE = 47,
        IDM_VERTSPACEINCREASE = 48,
        IDM_VERTSPACEMAKEEQUAL = 49,
        IDM_JUSTIFYFULL = 50,
        IDM_BACKCOLOR = 51,
        IDM_BOLD = 52,
        IDM_BORDERCOLOR = 53,
        IDM_FLAT = 54,
        IDM_FORECOLOR = 55,
        IDM_ITALIC = 56,
        IDM_JUSTIFYCENTER = 57,
        IDM_JUSTIFYGENERAL = 58,
        IDM_JUSTIFYLEFT = 59,
        IDM_JUSTIFYRIGHT = 60,
        IDM_RAISED = 61,
        IDM_SUNKEN = 62,
        IDM_UNDERLINE = 63,
        IDM_CHISELED = 64,
        IDM_ETCHED = 65,
        IDM_SHADOWED = 66,
        IDM_FIND = 67,
        IDM_SHOWGRID = 69,
        IDM_OBJECTVERBLIST0 = 72,
        IDM_OBJECTVERBLIST1 = 73,
        IDM_OBJECTVERBLIST2 = 74,
        IDM_OBJECTVERBLIST3 = 75,
        IDM_OBJECTVERBLIST4 = 76,
        IDM_OBJECTVERBLIST5 = 77,
        IDM_OBJECTVERBLIST6 = 78,
        IDM_OBJECTVERBLIST7 = 79,
        IDM_OBJECTVERBLIST8 = 80,
        IDM_OBJECTVERBLIST9 = 81,
        IDM_CONVERTOBJECT = 82,
        IDM_CUSTOMCONTROL = 83,
        IDM_CUSTOMIZEITEM = 84,
        IDM_RENAME = 85,
        IDM_IMPORT = 86,
        IDM_NEWPAGE = 87,
        IDM_MOVE = 88,
        IDM_CANCEL = 89,
        IDM_FONT = 90,
        IDM_STRIKETHROUGH = 91,
        IDM_DELETEWORD = 92,
        IDM_EXECPRINT = 93,
        IDM_JUSTIFYNONE = 94,
		IDM_TRISTATEBOLD        =    95,
		IDM_TRISTATEITALIC      =    96,
		IDM_TRISTATEUNDERLINE    =   97,
		IDM_FORWARDDELETE       =    98,

        IDM_FOLLOW_ANCHOR = 2008,

        IDM_INSINPUTIMAGE = 2114,
        IDM_INSINPUTBUTTON = 2115,
        IDM_INSINPUTRESET = 2116,
        IDM_INSINPUTSUBMIT = 2117,
        IDM_INSINPUTUPLOAD = 2118,
        IDM_INSFIELDSET = 2119,

        IDM_PASTEINSERT = 2120,
        IDM_REPLACE = 2121,
        IDM_EDITSOURCE = 2122,
        IDM_BOOKMARK = 2123,
        IDM_HYPERLINK = 2124,
        IDM_UNLINK = 2125,
        IDM_BROWSEMODE = 2126,
        IDM_EDITMODE = 2127,
        IDM_UNBOOKMARK = 2128,

        IDM_TOOLBARS = 2130,
        IDM_STATUSBAR = 2131,
        IDM_FORMATMARK = 2132,
        IDM_TEXTONLY = 2133,
        IDM_OPTIONS = 2135,
        IDM_FOLLOWLINKC = 2136,
        IDM_FOLLOWLINKN = 2137,
        IDM_VIEWSOURCE = 2139,
        IDM_ZOOMPOPUP = 2140,

        // IDM_BASELINEFONT1, IDM_BASELINEFONT2, IDM_BASELINEFONT3, IDM_BASELINEFONT4,
        // and IDM_BASELINEFONT5 should be consecutive integers;

        IDM_BASELINEFONT1 = 2141,
        IDM_BASELINEFONT2 = 2142,
        IDM_BASELINEFONT3 = 2143,
        IDM_BASELINEFONT4 = 2144,
        IDM_BASELINEFONT5 = 2145,

        IDM_HORIZONTALLINE = 2150,
        IDM_LINEBREAKNORMAL = 2151,
        IDM_LINEBREAKLEFT = 2152,
        IDM_LINEBREAKRIGHT = 2153,
        IDM_LINEBREAKBOTH = 2154,
        IDM_NONBREAK = 2155,
        IDM_SPECIALCHAR = 2156,
        IDM_HTMLSOURCE = 2157,
        IDM_IFRAME = 2158,
        IDM_HTMLCONTAIN = 2159,
        IDM_TEXTBOX = 2161,
        IDM_TEXTAREA = 2162,
        IDM_CHECKBOX = 2163,
        IDM_RADIOBUTTON = 2164,
        IDM_DROPDOWNBOX = 2165,
        IDM_LISTBOX = 2166,
        IDM_BUTTON = 2167,
        IDM_IMAGE = 2168,
        IDM_OBJECT = 2169,
        IDM_1D = 2170,
        IDM_IMAGEMAP = 2171,
        IDM_FILE = 2172,
        IDM_COMMENT = 2173,
        IDM_SCRIPT = 2174,
        IDM_JAVAAPPLET = 2175,
        IDM_PLUGIN = 2176,
        IDM_PAGEBREAK = 2177,
        IDM_HTMLAREA = 2178,

        IDM_PARAGRAPH = 2180,
        IDM_FORM = 2181,
        IDM_MARQUEE = 2182,
        IDM_LIST = 2183,
        IDM_ORDERLIST = 2184,
        IDM_UNORDERLIST = 2185,
        IDM_INDENT = 2186,
        IDM_OUTDENT = 2187,
        IDM_PREFORMATTED = 2188,
        IDM_ADDRESS = 2189,
        IDM_BLINK = 2190,
        IDM_DIV = 2191,

        IDM_TABLEINSERT = 2200,
        IDM_RCINSERT = 2201,
        IDM_CELLINSERT = 2202,
        IDM_CAPTIONINSERT = 2203,
        IDM_CELLMERGE = 2204,
        IDM_CELLSPLIT = 2205,
        IDM_CELLSELECT = 2206,
        IDM_ROWSELECT = 2207,
        IDM_COLUMNSELECT = 2208,
        IDM_TABLESELECT = 2209,
        IDM_TABLEPROPERTIES = 2210,
        IDM_CELLPROPERTIES = 2211,
        IDM_ROWINSERT = 2212,
        IDM_COLUMNINSERT = 2213,

        IDM_HELP_CONTENT = 2220,
        IDM_HELP_ABOUT = 2221,
        IDM_HELP_README = 2222,

        IDM_REMOVEFORMAT = 2230,
        IDM_PAGEINFO = 2231,
        IDM_TELETYPE = 2232,
        IDM_GETBLOCKFMTS = 2233,
        IDM_BLOCKFMT = 2234,
        IDM_SHOWHIDE_CODE = 2235,
        IDM_TABLE = 2236,

        IDM_COPYFORMAT = 2237,
        IDM_PASTEFORMAT = 2238,
        IDM_GOTO = 2239,

        IDM_CHANGEFONT = 2240,
        IDM_CHANGEFONTSIZE = 2241,
        IDM_INCFONTSIZE = 2242,
        IDM_DECFONTSIZE = 2243,
        IDM_INCFONTSIZE1PT = 2244,
        IDM_DECFONTSIZE1PT = 2245,
        IDM_CHANGECASE = 2246,
        IDM_SUBSCRIPT = 2247,
        IDM_SUPERSCRIPT = 2248,
        IDM_SHOWSPECIALCHAR = 2249,

        IDM_CENTERALIGNPARA = 2250,
        IDM_LEFTALIGNPARA = 2251,
        IDM_RIGHTALIGNPARA = 2252,
        IDM_REMOVEPARAFORMAT = 2253,
        IDM_APPLYNORMAL = 2254,
        IDM_APPLYHEADING1 = 2255,
        IDM_APPLYHEADING2 = 2256,
        IDM_APPLYHEADING3 = 2257,

        IDM_DOCPROPERTIES = 2260,
        IDM_ADDFAVORITES = 2261,
        IDM_COPYSHORTCUT = 2262,
        IDM_SAVEBACKGROUND = 2263,
        IDM_SETWALLPAPER = 2264,
        IDM_COPYBACKGROUND = 2265,
        IDM_CREATESHORTCUT = 2266,
        IDM_PAGE = 2267,
        IDM_SAVETARGET = 2268,
        IDM_SHOWPICTURE = 2269,
        IDM_SAVEPICTURE = 2270,
        IDM_DYNSRCPLAY = 2271,
        IDM_DYNSRCSTOP = 2272,
        IDM_PRINTTARGET = 2273,
        IDM_IMGARTPLAY = 2274,
        IDM_IMGARTSTOP = 2275,
        IDM_IMGARTREWIND = 2276,
        IDM_PRINTQUERYJOBSPENDING = 2277,
        IDM_SETDESKTOPITEM = 2278,

        IDM_CONTEXTMENU = 2280,
        IDM_GOBACKWARD = 2282,
        IDM_GOFORWARD = 2283,
        IDM_PRESTOP = 2284,
		IDM_MP_MYPICS         =      2287,
		IDM_MP_EMAILPICTURE    =     2288,
		IDM_MP_PRINTPICTURE    =     2289,

        IDM_CREATELINK = 2290,
        IDM_COPYCONTENT = 2291,

        IDM_LANGUAGE = 2292,
     IDM_GETPRINTTEMPLATE  =      2295,
     IDM_SETPRINTTEMPLATE   =     2296,
     IDM_TEMPLATE_PAGESETUP  =    2298,


        IDM_REFRESH = 2300,
        IDM_STOPDOWNLOAD = 2301,

        IDM_ENABLE_INTERACTION = 2302,

        IDM_LAUNCHDEBUGGER = 2310,
        IDM_BREAKATNEXT = 2311,

        IDM_INSINPUTHIDDEN = 2312,
        IDM_INSINPUTPASSWORD = 2313,

        IDM_OVERWRITE = 2314,

        IDM_PARSECOMPLETE = 2315,

        IDM_HTMLEDITMODE = 2316,

        IDM_REGISTRYREFRESH = 2317,
        IDM_COMPOSESETTINGS = 2318,

        IDM_SHOWALLTAGS = 2327,
        IDM_SHOWALIGNEDSITETAGS = 2321,
        IDM_SHOWSCRIPTTAGS = 2322,
        IDM_SHOWSTYLETAGS = 2323,
        IDM_SHOWCOMMENTTAGS = 2324,
        IDM_SHOWAREATAGS = 2325,
        IDM_SHOWUNKNOWNTAGS = 2326,
        IDM_SHOWMISCTAGS = 2320,
        IDM_SHOWZEROBORDERATDESIGNTIME = 2328,

        IDM_AUTODETECT = 2329,

        IDM_SCRIPTDEBUGGER = 2330,

        IDM_GETBYTESDOWNLOADED = 2331,

        IDM_NOACTIVATENORMALOLECONTROLS = 2332,
        IDM_NOACTIVATEDESIGNTIMECONTROLS = 2333,
        IDM_NOACTIVATEJAVAAPPLETS = 2334,
        IDM_NOFIXUPURLSONPASTE = 2335,

        IDM_EMPTYGLYPHTABLE = 2336,
        IDM_ADDTOGLYPHTABLE = 2337,
        IDM_REMOVEFROMGLYPHTABLE = 2338,
        IDM_REPLACEGLYPHCONTENTS = 2339,

        IDM_SHOWWBRTAGS = 2340,

        IDM_PERSISTSTREAMSYNC = 2341,
        IDM_SETDIRTY = 2342,

        IDM_RUNURLSCRIPT = 2343,

        IDM_ZOOMRATIO = 2344,
        IDM_GETZOOMNUMERATOR = 2345,
        IDM_GETZOOMDENOMINATOR = 2346,

        // COMMANDS FOR COMPLEX TEXT
        IDM_DIRLTR = 2350,
        IDM_DIRRTL = 2351,
        IDM_BLOCKDIRLTR = 2352,
        IDM_BLOCKDIRRTL = 2353,
        IDM_INLINEDIRLTR = 2354,
        IDM_INLINEDIRRTL = 2355,

        // SHDOCVW
        IDM_ISTRUSTEDDLG = 2356,

        // MSHTMLED
        IDM_INSERTSPAN = 2357,
        IDM_LOCALIZEEDITOR = 2358,

        // XML MIMEVIEWER
        IDM_SAVEPRETRANSFORMSOURCE = 2370,
        IDM_VIEWPRETRANSFORMSOURCE = 2371,
     IDM_SCROLL_HERE        =     2380,
     IDM_SCROLL_TOP          =    2381,
     IDM_SCROLL_BOTTOM       =    2382,
     IDM_SCROLL_PAGEUP       =    2383,
     IDM_SCROLL_PAGEDOWN     =    2384,
     IDM_SCROLL_UP           =    2385,
     IDM_SCROLL_DOWN         =    2386,
     IDM_SCROLL_LEFTEDGE     =    2387,
     IDM_SCROLL_RIGHTEDGE    =    2388,
     IDM_SCROLL_PAGELEFT     =    2389,
     IDM_SCROLL_PAGERIGHT    =    2390,
     IDM_SCROLL_LEFT         =    2391,
     IDM_SCROLL_RIGHT        =    2392,
     IDM_MULTIPLESELECTION   =    2393,
     IDM_2D_POSITION         =    2394,
     IDM_2D_ELEMENT          =    2395,
     IDM_1D_ELEMENT          =    2396,
     IDM_ABSOLUTE_POSITION   =    2397,
     IDM_LIVERESIZE          =    2398,
     IDM_ATOMICSELECTION     =       2399,
     IDM_AUTOURLDETECT_MODE  =    2400,
     IDM_IE50_PASTE          =    2401,
     IDM_IE50_PASTE_MODE     =    2402,
     IDM_DISABLE_EDITFOCUS_UI =   2404,

     IDM_RESPECTVISIBILITY_INDESIGN = 2405,

     IDM_CSSEDITING_LEVEL       =     2406,

     IDM_UI_OUTDENT              =    2407,

     IDM_UPDATEPAGESTATUS        =    2408,

     IDM_IME_ENABLE_RECONVERSION    =    2409,

        IDM_KEEPSELECTION           =     2410,

     IDM_UNLOADDOCUMENT            =  2411,

     IDM_OVERRIDE_CURSOR           =  2420,

     IDM_PEERHITTESTSAMEINEDIT     =  2423,

     IDM_TRUSTAPPCACHE             =  2425,

     IDM_BACKGROUNDIMAGECACHE     =   2430,

     IDM_GETUSERACTIONTIME         =  2431,
     IDM_BEGINUSERACTION           =  2432,
     IDM_ENDUSERACTION             =  2433,
     IDM_SETCUSTOMCURSOR           =  2434,

     IDM_FOLLOWLINKT               =  2435,

     IDM_CARETBROWSINGMODE         =  2436,

     IDM_STYLEMENU_SETNOSTYLE     =   2437,
     IDM_STYLEMENU_GETNOSTYLE      =  2438,
     IDM_STYLEMENU_GETPREFSTYLE    =  2439,
     IDM_STYLEMENU_CHANGESELECTEDSTYLE  =  2440,

     IDM_MEDIA_PLAYPAUSE            = 2441,
     IDM_MEDIA_MUTEUNMUTE           = 2442,
     IDM_MEDIA_PLAY                 = 2443,
     IDM_MEDIA_PAUSE                = 2444,
     IDM_MEDIA_STOP                 = 2445,
     IDM_MEDIA_FULLSCREEN_TOGGLE    = 2446,
     IDM_MEDIA_FULLSCREEN_EXIT      = 2447,
     IDM_MEDIA_VOLUME_UP            = 2448,
     IDM_MEDIA_VOLUME_DOWN          = 2449,
     IDM_MEDIA_SEEK_TO_START        = 2450,
     IDM_MEDIA_SEEK_TO_END          = 2451,
     IDM_MEDIA_SEEK_FWD_SMALL       = 2452,
     IDM_MEDIA_SEEK_BACK_SMALL      = 2453,
     IDM_MEDIA_SEEK_FWD_LARGE       = 2454,
     IDM_MEDIA_SEEK_BACK_LARGE      = 2455,
     IDM_MEDIA_RATE_FASTER          = 2456,
     IDM_MEDIA_RATE_SLOWER          = 2457,
     IDM_MEDIA_SHOWCONTROLS_TOGGLE  = 2458,
     IDM_MEDIA_ZOOMMODE_TOGGLE      = 2459,
     IDM_MEDIA_FRAMESTEP_FWD        = 2460,
     IDM_MEDIA_FRAMESTEP_BACK       = 2461,
     IDM_MEDIA_MUTE                 = 2462,
     IDM_MEDIA_UNMUTE               = 2463,
     IDM_MEDIA_SHOW_AUDIO_ACCESS    = 2464,
     IDM_MEDIA_SHOW_SUBTITLE_ACCESS = 2465,

     IDM_MEDIA_PLAYRATE0            = 2480,
     IDM_MEDIA_PLAYRATE1            = 2481,
     IDM_MEDIA_PLAYRATE2            = 2482,
     IDM_MEDIA_PLAYRATE3            = 2483,
     IDM_MEDIA_PLAYRATE4            = 2484,
     IDM_MEDIA_PLAYRATE5            = 2485,
     IDM_MEDIA_PLAYRATE6            = 2486,
     IDM_MEDIA_PLAYRATE7            = 2487,
     IDM_MEDIA_PLAYRATE8            = 2488,
     IDM_MEDIA_PLAYRATE9            = 2489,

     IDM_PASTECONTENTONLY           = 2500,
     IDM_PASTETEXTONLY              = 2501,
     IDM_INSERTHTML                 = 2502,
        IDM_MIMECSET__FIRST__ = 3609,
        IDM_MIMECSET__LAST__ = 3699,

        IDM_MENUEXT_FIRST__ = 3700,
        IDM_MENUEXT_LAST__ = 3732,
        IDM_MENUEXT_COUNT = 3733,
     IDM_ADDCONSOLEMESSAGERECEIVER    = 3800,
     IDM_REMOVECONSOLEMESSAGERECEIVER = 3801,
     IDM_STARTDIAGNOSTICSMODE         = 3802,
     IDM_GETSCRIPTENGINE              = 3803,
     IDM_ADDDEBUGCALLBACKRECEIVER     = 3804,
     IDM_REMOVEDEBUGCALLBACKRECEIVER  = 3805,

     IDM_DEFAULTPARAGRAPHSEPARATOR    = 3900,
     IDM_BEGINUNDOUNIT                = 3901,
     IDM_ENDUNDOUNIT                  = 3902,
     IDM_CLEARUNDO                    = 3903,
     IDM_INSPECTELEMENT               = 3904,

     IDM_SHAREPICTURE                 = 3905,
     IDM_ENABLE_OBJECT_RESIZING       = 3906,
     IDM_INSERTTEXT                   = 3907,
     IDM_LAUNCHURICALLBACK            = 3908,

        // Commands mapped from the standard set.  We should
        // consider deleting them from public header files.

        IDM_OPEN = 2000,
        IDM_NEW = 2001,
        IDM_SAVE = 70,
        IDM_SAVEAS = 71,
        IDM_SAVECOPYAS = 2002,
        IDM_PRINTPREVIEW = 2003,
        IDM_PRINT = 27,
        IDM_PAGESETUP = 2004,
        IDM_SPELL = 2005,
        IDM_PASTESPECIAL = 2006,
        IDM_CLEARSELECTION = 2007,
        IDM_PROPERTIES = 28,
        IDM_REDO = 29,
        IDM_UNDO = 43,
        IDM_SELECTALL = 31,
        IDM_ZOOMPERCENT = 50,
        IDM_GETZOOM = 68,
        IDM_STOP = 2138,
        IDM_COPY = 15,
        IDM_CUT = 16,
        IDM_PASTE = 26,
		IDM_SAVEPDF     =            99,
        // Defines for IDM_ZOOMPERCENT
        CMD_ZOOM_PAGEWIDTH = -1,
        CMD_ZOOM_ONEPAGE = -2,
        CMD_ZOOM_TWOPAGES = -3,
        CMD_ZOOM_SELECTION = -4,
        CMD_ZOOM_FIT = -5,
     IDM_CONTEXT                = 1,
     IDM_HWND                   = 2,

     IDM_NEW_TOPLEVELWINDOW     = 7050,


     IDM_PRESERVEUNDOALWAYS     = 6049,
     IDM_PERSISTDEFAULTVALUES   = 7100,
     IDM_PROTECTMETATAGS        = 7101,

     IDM_GETFRAMEZONE           = 6037,

     IDM_REFRESH_THIS           = 6042,

     IDM_MENUEXT_PLACEHOLDER    = 6047,


     IDM_SETPAGEACTIONALLOWEDFLAGS  = 15100,

     IDM_INVOKEFLIPAHEADTARGET      = 15200,
     IDM_ENABLEFLIPAHEADTARGET      = 15201,

     IDM_DEBUGGERDYNAMICATTACH      = 15202,
     IDM_DEBUGGERDYNAMICDETACH      = 15203,
     IDM_DEBUGGERDYNAMICATTACHSOURCERUNDOWN = 15204,
     IDM_GETDEBUGGERSTATE           = 15205,
     IDM_SELECTIONSEARCH            = 15206,
     IDM_SHOWSHAREUI                = 15207,
    } WBIDM;

    typedef enum ContextMenuTarget {
        CONTEXT_MENU_DEFAULT = 0,
        CONTEXT_MENU_IMAGE = 1,
        CONTEXT_MENU_CONTROL = 2,
        CONTEXT_MENU_TABLE = 3,
        CONTEXT_MENU_SELECTION = 4,
        CONTEXT_MENU_LINK = 5,
        CONTEXT_MENU_ANCHOR = 6,
        CONTEXT_MENU_IMGDYNSRC = 7,
        CONTEXT_MENU_IMGART = 8,
    } ContextMenuTarget;

    typedef enum DOCHOSTUIDBLCLK {
        DOCHOSTUIDBLCLK_DEFAULT = 0,
        DOCHOSTUIDBLCLK_SHOWPROPERTIES = 1,
        DOCHOSTUIDBLCLK_SHOWCODE = 2,
    } DOCHOSTUIDBLCLK ;

    typedef enum DOCHOSTUIFLAG {
        DOCHOSTUIFLAG_DIALOG = 1,
        DOCHOSTUIFLAG_DISABLE_HELP_MENU = 2,
        DOCHOSTUIFLAG_NO3DBORDER = 4,
        DOCHOSTUIFLAG_SCROLL_NO = 8,
        DOCHOSTUIFLAG_DISABLE_SCRIPT_INACTIVE = 0x10,
        DOCHOSTUIFLAG_OPENNEWUI = 0x20,
        DOCHOSTUIFLAG_DISABLE_OFFSCREEN = 0x40,
        DOCHOSTUIFLAG_FLAT_SCROLLBAR = 0x80,
        DOCHOSTUIFLAG_DIV_BLOCKDEFAULT = 0x0100,
        DOCHOSTUIFLAG_ACTIVATE_CLIENTHIT_ONLY = 0x0200,
        DOCHOSTUIFLAG_OVERRIDEBEHAVIORFACTORY = 0x0400,
        DOCHOSTUIFLAG_CODEPAGELINKEDFONTS = 0x0800,
        DOCHOSTUIFLAG_URL_ENCODING_DISABLE_UTF8 = 0x1000,
        DOCHOSTUIFLAG_URL_ENCODING_ENABLE_UTF8 = 0x2000,
        DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE = 0x4000,
        DOCHOSTUIFLAG_ENABLE_INPLACE_NAVIGATION = 0x10000,
        DOCHOSTUIFLAG_IME_ENABLE_RECONVERSION = 0x20000,
        DOCHOSTUIFLAG_BROWSER = 0x12,
        DOCHOSTUIFLAG_DESKTOP = 0x2E,
   } DOCHOSTUIFLAG ;

    typedef struct DOCHOSTUIINFO {
        LONG cbSize;
        DOCHOSTUIFLAG dwFlags;
        DOCHOSTUIDBLCLK dwDoubleClick;
        LONG pchHostCss;
        LONG pchHostNS;
    } DOCHOSTUIINFO;

    typedef [helpstring("DISP_E_EXCEPTION Error info")] struct EXCEPINFO {
        [helpstring("Error number")]
        short wCode;
        short wReserved;
        [helpstring("Source")]
        BSTR Source;
        [helpstring("Description")]
        BSTR Description;
        [helpstring("Help file")]
        BSTR HelpFile;
        [helpstring("Help context")]
        LONG dwHelpContext;
        LONG pvReserved;
        [helpstring("HRESULT (__stdcall *pfnDeferredFillIn)(struct EXCEPINFO *);")]
        LONG pfnDeferredFillIn;
        LONG scode;
    } EXCEPINFO;

    typedef struct DISPPARAMS {
        [helpstring("Parameters")]
        LONG rgPointerToVariantArray;
        [helpstring("Named parameters")]
        LONG rgPointerToLONGNamedArgs;
        [helpstring("Parameter count")]
        LONG cArgs;
        [helpstring("Named parameter count")]
        LONG cNamedArgs;
    } DISPPARAMS;

    typedef [helpstring("InvokeFlags")] enum InvokeFlags {
        [helpstring("Sub/Function")]
        DISPATCH_METHOD = 0x1,
        [helpstring("Property Get")]
        DISPATCH_PROPERTYGET = 0x2,
        [helpstring("Property Let")]
        DISPATCH_PROPERTYPUT = 0x4,
        [helpstring("Property Set")]
        DISPATCH_PROPERTYPUTREF = 0x8
    } InvokeFlags;

    typedef [helpstring("DEFAULT_DISPIDS")] enum DEFAULT_DISPIDS {
        DISPID_UNKNOWN = -1,
        DISPID_VALUE = 0,
        DISPID_PROPERTYPUT = -3,
        DISPID_NEWENUM = -4,
        DISPID_EVALUATE = -5,
        DISPID_CONSTRUCTOR = -6,
        DISPID_DESTRUCTOR = -7,
        DISPID_COLLECT = -8,
          DISPID_AUTOSIZE                 = (-500),
          DISPID_BACKCOLOR                = (-501),
          DISPID_BACKSTYLE                = (-502),
          DISPID_BORDERCOLOR              = (-503),
          DISPID_BORDERSTYLE              = (-504),
          DISPID_BORDERWIDTH              = (-505),
          DISPID_DRAWMODE                 = (-507),
          DISPID_DRAWSTYLE                = (-508),
          DISPID_DRAWWIDTH                = (-509),
          DISPID_FILLCOLOR                = (-510),
          DISPID_FILLSTYLE                = (-511),
          DISPID_FONT                     = (-512),
          DISPID_FORECOLOR                = (-513),
          DISPID_ENABLED                  = (-514),
          DISPID_HWND                     = (-515),
          DISPID_TABSTOP                  = (-516),
          DISPID_TEXT                     = (-517),
          DISPID_CAPTION                  = (-518),
          DISPID_BORDERVISIBLE            = (-519),
          DISPID_APPEARANCE               = (-520),
          DISPID_MOUSEPOINTER             = (-521),
          DISPID_MOUSEICON                = (-522),
          DISPID_PICTURE                  = (-523),
          DISPID_VALID                    = (-524),
          DISPID_READYSTATE               = (-525),
          DISPID_LISTINDEX                = (-526),
          DISPID_SELECTED                 = (-527),
          DISPID_LIST                     = (-528),
          DISPID_COLUMN                   = (-529),
          DISPID_LISTCOUNT                = (-531),
          DISPID_MULTISELECT              = (-532),
          DISPID_MAXLENGTH                = (-533),
          DISPID_PASSWORDCHAR             = (-534),
          DISPID_SCROLLBARS               = (-535),
          DISPID_WORDWRAP                 = (-536),
          DISPID_MULTILINE                = (-537),
          DISPID_NUMBEROFROWS             = (-538),
          DISPID_NUMBEROFCOLUMNS          = (-539),
          DISPID_DISPLAYSTYLE             = (-540),
          DISPID_GROUPNAME                = (-541),
          DISPID_IMEMODE                  = (-542),
          DISPID_ACCELERATOR              = (-543),
          DISPID_ENTERKEYBEHAVIOR         = (-544),
          DISPID_TABKEYBEHAVIOR           = (-545),
          DISPID_SELTEXT                  = (-546),
          DISPID_SELSTART                 = (-547),
          DISPID_SELLENGTH                = (-548),
          
          DISPID_REFRESH                  = (-550),
          DISPID_DOCLICK                  = (-551),
          DISPID_ABOUTBOX                 = (-552),
          DISPID_ADDITEM                  = (-553),
          DISPID_CLEAR                    = (-554),
          DISPID_REMOVEITEM               = (-555),

          DISPID_CLICK                    = (-600),
          DISPID_DBLCLICK                 = (-601),
          DISPID_KEYDOWN                  = (-602),
          DISPID_KEYPRESS                 = (-603),
          DISPID_KEYUP                    = (-604),
          DISPID_MOUSEDOWN                = (-605),
          DISPID_MOUSEMOVE                = (-606),
          DISPID_MOUSEUP                  = (-607),
          DISPID_ERROREVENT               = (-608),
          DISPID_READYSTATECHANGE         = (-609),
          DISPID_CLICK_VALUE              = (-610),
          DISPID_RIGHTTOLEFT              = (-611),
          DISPID_TOPTOBOTTOM              = (-612),
          DISPID_THIS                     = (-613),

          DISPID_AMBIENT_BACKCOLOR        = (-701),
          DISPID_AMBIENT_DISPLAYNAME      = (-702),
          DISPID_AMBIENT_FONT             = (-703),
          DISPID_AMBIENT_FORECOLOR        = (-704),
          DISPID_AMBIENT_LOCALEID         = (-705),
          DISPID_AMBIENT_MESSAGEREFLECT   = (-706),
          DISPID_AMBIENT_SCALEUNITS       = (-707),
          DISPID_AMBIENT_TEXTALIGN        = (-708),
          DISPID_AMBIENT_USERMODE         = (-709),
          DISPID_AMBIENT_UIDEAD           = (-710),
          DISPID_AMBIENT_SHOWGRABHANDLES  = (-711),
          DISPID_AMBIENT_SHOWHATCHING     = (-712),
          DISPID_AMBIENT_DISPLAYASDEFAULT = (-713),
          DISPID_AMBIENT_SUPPORTSMNEMONICS = (-714),
          DISPID_AMBIENT_AUTOCLIP         = (-715),
          DISPID_AMBIENT_APPEARANCE       = (-716),

          DISPID_AMBIENT_CODEPAGE         = (-725),
          DISPID_AMBIENT_PALETTE          = (-726),
          DISPID_AMBIENT_CHARSET          = (-727),
          DISPID_AMBIENT_TRANSFERPRIORITY = (-728),

          DISPID_AMBIENT_RIGHTTOLEFT      = (-732),
          DISPID_AMBIENT_TOPTOBOTTOM      = (-733),

          DISPID_Name                     = (-800),
          DISPID_Delete                   = (-801),
          DISPID_Object                   = (-802),
          DISPID_Parent                   = (-803)
    } DEFAULT_DISPIDS;

    typedef [helpstring("IDispatch::Invoke results")] enum InvokeResults {
        [helpstring("Unknown Interface")]
        DISP_E_UNKNOWNINTERFACE = 0x80020001,
        [helpstring("Member not found")]
        DISP_E_MEMBERNOTFOUND = 0x80020003,
        [helpstring("Parameter not found")]
        DISP_E_PARAMNOTFOUND = 0x80020004,
        [helpstring("Type Mismatch")]
        DISP_E_TYPEMISMATCH = 0x80020005,
        [helpstring("Unknown Name")]
        DISP_E_UNKNOWNNAME = 0x80020006,
        [helpstring("No named arguments")]
        DISP_E_NONAMEDARGS = 0x80020007,
        [helpstring("Bad variable type")]
        DISP_E_BADVARTYPE = 0x80020008,
        [helpstring("Exception")]
        DISP_E_EXCEPTION = 0x80020009,
        [helpstring("Overflow")]
        DISP_E_OVERFLOW = 0x8002000A,
        [helpstring("Bad Index")]
        DISP_E_BADINDEX = 0x8002000B,
        [helpstring("Unknown LCID")]
        DISP_E_UNKNOWNLCID = 0x8002000C,
        [helpstring("Memory is locked")]
        DISP_E_ARRAYISLOCKED = 0x8002000D,
        [helpstring("Invalid number of parameters")]
        DISP_E_BADPARAMCOUNT = 0x8002000E,
        [helpstring("Parameter not optional")]
        DISP_E_PARAMNOTOPTIONAL = 0x8002000F,
        [helpstring("Invalid callee")]
        DISP_E_BADCALLEE = 0x80020010,
        [helpstring("Does not support a collection")]
        DISP_E_NOTACOLLECTION = 0x80020011,
        [helpstring("Division by zero")]
        DISP_E_DIVBYZERO = 0x80020012
    } InvokeResults;


    typedef enum DISPID_EXTENDER {
        DISPID_EXTENDER_NAME = 0x80010000,
        DISPID_EXTENDER_INDEX = 0x80010001,
        DISPID_EXTENDER_LEFT = 0x80010003,
        DISPID_EXTENDER_TOP = 0x80010004,
        DISPID_EXTENDER_WIDTH = 0x80010005,
        DISPID_EXTENDER_HEIGHT = 0x80010006,
        DISPID_EXTENDER_VISIBLE = 0x80010007,
        DISPID_EXTENDER_PARENT = 0x80010008,
        DISPID_EXTENDER_DRAGMODE = 0x80010009,
        DISPID_EXTENDER_DRAGICON = 0x8001000A,
        DISPID_EXTENDER_TAG = 0x8001000B,
        DISPID_EXTENDER_TABSTOP = 0x8001000E,
        DISPID_EXTENDER_TABINDEX = 0x8001000F,
        DISPID_EXTENDER_OBJECT = 0x8001001C,
        DISPID_EXTENDER_HELPCONTEXTID = 0x80010032,
        DISPID_EXTENDER_DEFAULT = 0x80010037,
        DISPID_EXTENDER_CANCEL = 0x80010038,
        DISPID_EXTENDER_ALIGN = 0x8001003B,
        DISPID_EXTENDER_NEGOTIATE = 0x8001003C,
        DISPID_EXTENDER_WHATSTHISHELP = 0x80010046,
        DISPID_EXTENDER_CONTAINER = 0x80010047,
        DISPID_EXTENDER_TOOLTIPTEXT = 0x8001004A,

        DISPID_EXTENDER_SETFOCUS = 0x80011000,
        DISPID_EXTENDER_ZORDER = 0x80011001,
        DISPID_EXTENDER_MOVE = 0x80011002,
        DISPID_EXTENDER_DRAG = 0x80011003,
        DISPID_EXTENDER_SHOWWHATSTHIS = 0x80011004,
    } DISPID_EXTENDER;

    typedef enum TYPEKIND {
        TKIND_ENUM = 0,
        TKIND_RECORD,
        TKIND_MODULE,
        TKIND_INTERFACE,
        TKIND_DISPATCH,
        TKIND_COCLASS,
        TKIND_ALIAS,
        TKIND_UNION,
        TKIND_MAX
    } TYPEKIND;

    typedef enum SYSKIND {
        SYS_WIN16 = 0,
        SYS_WIN32,
        SYS_MAC,
        SYS_WIN64
    } SYSKIND;

    typedef enum LIBFLAGS {
        LIBFLAG_FRESTRICTED = 0x01,
        LIBFLAG_FCONTROL = 0x02,
        LIBFLAG_FHIDDEN = 0x04,
        LIBFLAG_FHASDISKIMAGE = 0x08
    } LIBFLAGS;

    typedef struct TLIBATTR {
        UUID IID;
        LONG lcid;
        SYSKIND syskind;
        short wMajorVerNum;
        short wMinorVerNum;
        short wLibFlags;
    } TLIBATTR;

    typedef struct TYPEDESC {
        LONG pTypeDesc;
        short vt;
    } TYPEDESC;

    typedef struct SAFEARRAYBOUND {
        LONG cElements;
        LONG lLbound;
    } SAFEARRAYBOUND;

    typedef struct ARRAYDESC {
        TYPEDESC tdescElem;
        short cDims;
    } ARRAYDESC;

    typedef struct PARAMDESCEX {
        LONG cBytes;
    LONG ignore;
        VARIANT varDefaultValue;
    } PARAMDESCEX;

    typedef struct PARAMDESC {
        [helpstring("valid if PARAMFLAG_FHASDEFAULT bit is set")]
        LONG pPARAMDESCEX;
        short wParamFlags;
    } PARAMDESC;

    typedef enum PARAMFLAGS {
       PARAMFLAG_NONE = 0x00,
       PARAMFLAG_FIN = 0x01,
       PARAMFLAG_FOUT = 0x02,
       PARAMFLAG_FLONG = 0x04,
       PARAMFLAG_FRETVAL = 0x08,
       PARAMFLAG_FOPT = 0x10,
       PARAMFLAG_FHASDEFAULT = 0x20,
    } PARAMFLAGS;

    typedef struct ELEMDESC {
        TYPEDESC tdesc;
        PARAMDESC paramdesc;
    } ELEMDESC;

    typedef struct IDLDESC {
        LONG dwReserved;
        short wIDLFlags;
    } IDLDESC;

    typedef struct TYPEATTR {
        [helpstring("the IID of the TypeInfo")]
        UUID IID;
        [helpstring("locale of member names and doc strings")]
        LONG lcid;
        LONG dwReserved;
        [helpstring("ID of constructor, LONG_NIL if none")]
        LONG memidConstructor;
        [helpstring("ID of destructor, LONG_NIL if none")]
        LONG memidDestructor;
        LONG lpstrSchema;
        [helpstring("the size of an instance of this type")]
        LONG cbSizeInstance;
        [helpstring("the kind of type this typeinfo describes")]
        TYPEKIND typekind;
        [helpstring("number of functions")]
        short cFuncs;
        [helpstring("number of variables / data members")]
        short cVars;
        [helpstring("number of implemented interfaces")]
        short cImplTypes;
        [helpstring("the size of this types virtual func table")]
        short cbSizeVft;
        [helpstring("specifies the alignment requirements for an instance of this type,\n0 = align on 64k boundary\n1 = byte align\n2 = word align\n4 = dword align")]
        short cbAlignment;
        short wTypeFlags;
        [helpstring("major version number")]
        short wMajorVerNum;
        [helpstring("minor version number")]
        short wMinorVerNum;
        [helpstring("if typekind = = TKIND_ALIAS this field specifies the type for which this type is an alias")]
        TYPEDESC tdescAlias;
        [helpstring("IDL attributes of the described type")]
        IDLDESC idldescType;
    } TYPEATTR;

typedef enum CALLCONV_TYPES {
    CC_FASTCALL = 0,
    CC_CDECL = 1,
    CC_MSCPASCAL,
    CC_PASCAL = 2,
    CC_MACPASCAL,
    CC_STDCALL,
    CC_FPFASTCALL,
    CC_SYSCALL,
    CC_MPWCDECL,
    CC_MPWPASCAL,
    CC_MAX
} CALLCONV_TYPES;

typedef enum FUNCKIND {
    FUNC_VIRTUAL,
    FUNC_PUREVIRTUAL,
    FUNC_NONVIRTUAL,
    FUNC_STATIC,
    FUNC_DISPATCH
} FUNCKIND;

typedef enum INVOKEKIND {
    INVOKE_FUNC = 1,
    INVOKE_PROPERTYGET = 2,
    INVOKE_PROPERTYPUT = 4,
    INVOKE_PROPERTYPUTREF = 8
} INVOKEKIND;

typedef struct FUNCDESC {
    LONG memid;
    LONG lprgscode;
     [helpstring("array of param types")]
    LONG lprgELEMDESCParam;
    FUNCKIND funckind;
    INVOKEKIND invkind;
    CALLCONV_TYPES callconv;
    short cParams;
    short cParamsOpt;
    short oVft;
    short cScodes;
    ELEMDESC elemdescFunc;
    short wFuncFlags;
} FUNCDESC;

typedef enum VARKIND {
    VAR_PERINSTANCE,
    VAR_STATIC,
    VAR_CONST,
    VAR_DISPATCH
} VARKIND;

typedef enum IMPLTYPE {
    IMPLTYPEFLAG_FDEFAULT = 0x1,
    IMPLTYPEFLAG_FSOURCE = 0x2,
    IMPLTYPEFLAG_FRESTRICTED = 0x4,
    IMPLTYPEFLAG_FDEFAULTVTABLE= 0x8,
} IMPLTYPE;

typedef struct VARDESC {
    LONG memid;
    LONG lpstrSchema;
    LONG oInst_varValue;
    ELEMDESC elemdescVar;
    short wVarFlags;
    VARKIND varkind;
} VARDESC;

typedef enum TYPEFLAGS {
    TYPEFLAG_FAPPOBJECT = 0x01,
    TYPEFLAG_FCANCREATE = 0x02,
    TYPEFLAG_FLICENSED = 0x04,
    TYPEFLAG_FPREDECLID = 0x08,
    TYPEFLAG_FHIDDEN = 0x10,
    TYPEFLAG_FCONTROL = 0x20,
    TYPEFLAG_FDUAL = 0x40,
    TYPEFLAG_FNONEXTENSIBLE = 0x80,
    TYPEFLAG_FOLEAUTOMATION = 0x100,
    TYPEFLAG_FRESTRICTED = 0x200,
    TYPEFLAG_FAGGREGATABLE = 0x400,
    TYPEFLAG_FREPLACEABLE = 0x800,
    TYPEFLAG_FDISPATCHABLE = 0x1000,
    TYPEFLAG_FREVERSEBIND = 0x2000,
	TYPEFLAG_FPROXY          = 0x4000
} TYPEFLAGS;

typedef enum FUNCFLAGS {
    FUNCFLAG_FRESTRICTED = 0x1,
    FUNCFLAG_FSOURCE = 0x2,
    FUNCFLAG_FBINDABLE = 0x4,
    FUNCFLAG_FREQUESTEDIT = 0x8,
    FUNCFLAG_FDISPLAYBIND = 0x10,
    FUNCFLAG_FDEFAULTBIND = 0x20,
    FUNCFLAG_FHIDDEN = 0x40,
    FUNCFLAG_FUSESGETLASTERROR = 0x80,
    FUNCFLAG_FDEFAULTCOLLELEM = 0x100,
    FUNCFLAG_FUIDEFAULT = 0x200,
    FUNCFLAG_FNONBROWSABLE = 0x400,
    FUNCFLAG_FREPLACEABLE = 0x800,
    FUNCFLAG_FIMMEDIATEBIND = 0x1000
} FUNCFLAGS;

typedef enum VARFLAGS {
    VARFLAG_FREADONLY = 0x1,
    VARFLAG_FSOURCE = 0x2,
    VARFLAG_FBINDABLE = 0x4,
    VARFLAG_FREQUESTEDIT = 0x8,
    VARFLAG_FDISPLAYBIND = 0x10,
    VARFLAG_FDEFAULTBIND = 0x20,
    VARFLAG_FHIDDEN = 0x40,
    VARFLAG_FRESTRICTED = 0x80,
    VARFLAG_FDEFAULTCOLLELEM = 0x100,
    VARFLAG_FUIDEFAULT = 0x200,
    VARFLAG_FNONBROWSABLE = 0x400,
    VARFLAG_FREPLACEABLE = 0x800,
    VARFLAG_FIMMEDIATEBIND = 0x1000
} VARFLAGS;

typedef struct CUSTDATAITEM {
    UUID IID;
    VARIANT varValue;
} CUSTDATAITEM;

typedef struct CUSTDATA {
     [helpstring("number of custom data items in rgCustData")]
    LONG cCustData;
    LONG prgCustData;
} CUSTDATA;

typedef [helpstring("Standard DISPIDs")] enum STANDARD_DISPIDS {
    DISPID_AUTOSIZE = (-500),
    DISPID_BACKCOLOR = (-501),
    DISPID_BACKSTYLE = (-502),
    DISPID_BORDERCOLOR = (-503),
    DISPID_BORDERSTYLE = (-504),
    DISPID_BORDERWIDTH = (-505),
    DISPID_DRAWMODE = (-507),
    DISPID_DRAWSTYLE = (-508),
    DISPID_DRAWWIDTH = (-509),
    DISPID_FILLCOLOR = (-510),
    DISPID_FILLSTYLE = (-511),
    DISPID_FONT = (-512),
    DISPID_FORECOLOR = (-513),
    DISPID_ENABLED = (-514),
    DISPID_HWND = (-515),
    DISPID_TABSTOP = (-516),
    DISPID_TEXT = (-517),
    DISPID_CAPTION = (-518),
    DISPID_BORDERVISIBLE = (-519),
    DISPID_APPEARANCE = (-520),
    DISPID_MOUSEPOINTER = (-521),
    DISPID_MOUSEICON = (-522),
    DISPID_PICTURE = (-523),
    DISPID_VALID = (-524),
    DISPID_READYSTATE = (-525),
    DISPID_LISTINDEX = (-526),
    DISPID_SELECTED = (-527),
    DISPID_LIST = (-528),
    DISPID_COLUMN = (-529),
    DISPID_LISTCOUNT = (-531),
    DISPID_MULTISELECT = (-532),
    DISPID_MAXLENGTH = (-533),
    DISPID_PASSWORDCHAR = (-534),
    DISPID_SCROLLBARS = (-535),
    DISPID_WORDWRAP = (-536),
    DISPID_MULTILINE = (-537),
    DISPID_NUMBEROFROWS = (-538),
    DISPID_NUMBEROFCOLUMNS = (-539),
    DISPID_DISPLAYSTYLE = (-540),
    DISPID_GROUPNAME = (-541),
    DISPID_IMEMODE = (-542),
    DISPID_ACCELERATOR = (-543),
    DISPID_ENTERKEYBEHAVIOR = (-544),
    DISPID_TABKEYBEHAVIOR = (-545),
    DISPID_SELTEXT = (-546),
    DISPID_SELSTART = (-547),
    DISPID_SELLENGTH = (-548),

    DISPID_REFRESH = (-550),
    DISPID_DOCLICK = (-551),
    DISPID_ABOUTBOX = (-552),
    DISPID_ADDITEM = (-553),
    DISPID_CLEAR = (-554),
    DISPID_REMOVEITEM = (-555),

    DISPID_CLICK = (-600),
    DISPID_DBLCLICK = (-601),
    DISPID_KEYDOWN = (-602),
    DISPID_KEYPRESS = (-603),
    DISPID_KEYUP = (-604),
    DISPID_MOUSEDOWN = (-605),
    DISPID_MOUSEMOVE = (-606),
    DISPID_MOUSEUP = (-607),
    DISPID_ERROREVENT = (-608),
    DISPID_READYSTATECHANGE = (-609),
    DISPID_CLICK_VALUE = (-610),
    DISPID_RIGHTTOLEFT = (-611),
    DISPID_TOPTOBOTTOM = (-612),

    DISPID_NAME = (-800),
    DISPID_DELETE = (-801),
    DISPID_OBJECT = (-802),
    DISPID_PARENT = (-803)
} STANDARD_DISPIDS ;

typedef enum AMBIENT_DISPIDS {
    DISPID_AMBIENT_BACKCOLOR = (-701),
    DISPID_AMBIENT_DISPLAYNAME = (-702),
    DISPID_AMBIENT_FONT = (-703),
    DISPID_AMBIENT_FORECOLOR = (-704),
    DISPID_AMBIENT_LOCALEID = (-705),
    DISPID_AMBIENT_MESSAGEREFLECT = (-706),
    DISPID_AMBIENT_SCALEUNITS = (-707),
    DISPID_AMBIENT_TEXTALIGN = (-708),
    DISPID_AMBIENT_USERMODE = (-709),
    DISPID_AMBIENT_UIDEAD = (-710),
    DISPID_AMBIENT_SHOWGRABHANDLES = (-711),
    DISPID_AMBIENT_SHOWHATCHING = (-712),
    DISPID_AMBIENT_DISPLAYASDEFAULT = (-713),
    DISPID_AMBIENT_SUPPORTSMNEMONICS = (-714),
    DISPID_AMBIENT_AUTOCLIP = (-715),
    DISPID_AMBIENT_APPEARANCE = (-716),

    DISPID_AMBIENT_PALETTE = (-726),
    DISPID_AMBIENT_TRANSFERPRIORITY = (-728),

    DISPID_AMBIENT_RIGHTTOLEFT = (-732),
    DISPID_AMBIENT_TOPTOBOTTOM = (-733)

} AMBIENT_DISPIDS;

typedef enum OLEOBJECTS_DISPIDS {
    DISPID_FONT_NAME = 0,
    DISPID_FONT_SIZE = 2,
    DISPID_FONT_BOLD = 3,
    DISPID_FONT_ITALIC = 4,
    DISPID_FONT_UNDER = 5,
    DISPID_FONT_STRIKE = 6,
    DISPID_FONT_WEIGHT = 7,
    DISPID_FONT_CHARSET = 8,

    DISPID_PICT_HANDLE = 0,
    DISPID_PICT_HPAL = 2,
    DISPID_PICT_TYPE = 3,
    DISPID_PICT_WIDTH = 4,
    DISPID_PICT_HEIGHT = 5,
    DISPID_PICT_RENDER = 6
} OLEOBJECTS_DISPIDS ;

typedef enum REGKIND {
    REGKIND_DEFAULT,
    REGKIND_REGISTER,
    REGKIND_NONE
} REGKIND;

typedef enum PARSEACTION {
    PARSE_CANONICALIZE = 1,
    PARSE_FRIENDLY,
    PARSE_SECURITY_URL,
    PARSE_ROOTDOCUMENT,
    PARSE_DOCUMENT,
    PARSE_ANCHOR,
    PARSE_ENCODE,
    PARSE_DECODE,
    PARSE_PATH_FROM_URL,
    PARSE_URL_FROM_PATH,
    PARSE_MIME,
    PARSE_SERVER,
    PARSE_SCHEMA,
    PARSE_SITE,
    PARSE_DOMAIN,
    PARSE_LOCATION,
    PARSE_SECURITY_DOMAIN,
    PARSE_ESCAPE,
    PARSE_UNESCAPE
} PARSEACTION;

typedef enum PSUACTION {
    PSU_DEFAULT = 1,
    PSU_SECURITY_URL_ONLY
} PSUACTION;

typedef enum QUERYOPTION {
    QUERY_EXPIRATION_DATE = 1,
    QUERY_TIME_OF_LAST_CHANGE,
    QUERY_CONTENT_ENCODING,
    QUERY_CONTENT_TYPE,
    QUERY_REFRESH,
    QUERY_RECOMBINE,
    QUERY_CAN_NAVIGATE,
    QUERY_USES_NETWORK,
    QUERY_IS_CACHED,
    QUERY_IS_INSTALLEDENTRY,
    QUERY_IS_CACHED_OR_MAPPED,
    QUERY_USES_CACHE,
    QUERY_IS_SECURE,
    QUERY_IS_SAFE
} QUERYOPTION;

typedef enum VARFORMAT_FLAGS {
    VAR_TIMEVALUEONLY = 0x00000001,    /* return time value */
    VAR_DATEVALUEONLY = 0x00000002,    /* return date value */
    VAR_VALIDDATE = 0x00000004,
    VAR_CALENDAR_HIJRI = 0x00000008,   /* use Hijri calender */
    VAR_LOCALBOOL = 0x00000010,
    VAR_FORMAT_NOSUBSTITUTE = 0x00000020
} VARFORMAT_FLAGS;

typedef enum NAMEDFORMATS {
    GeneralDate,
    LongDate,
    ShortDate,
    LongTime,
    ShortTime
} NAMEDFORMATS;

typedef enum ACTIVEOBJECTFLAGS {
    ACTIVEOBJECT_STRONG = 0x0,
    ACTIVEOBJECT_WEAK = 0x1
} ACTIVEOBJECTFLAGS;

typedef enum NUMPRS {
    NUMPRS_LEADING_WHITE    = 0x0001,
    NUMPRS_TRAILING_WHITE   = 0x0002,
    NUMPRS_LEADING_PLUS     = 0x0004,
    NUMPRS_TRAILING_PLUS    = 0x0008,
    NUMPRS_LEADING_MINUS    = 0x0010,
    NUMPRS_TRAILING_MINUS   = 0x0020,
    NUMPRS_HEX_OCT          = 0x0040,
    NUMPRS_PARENS           = 0x0080,
    NUMPRS_DECIMAL          = 0x0100,
    NUMPRS_THOUSANDS        = 0x0200,
    NUMPRS_CURRENCY         = 0x0400,
    NUMPRS_EXPONENT         = 0x0800,
    NUMPRS_USE_ALL          = 0x1000,
    NUMPRS_STD              = 0x1FFF,
    NUMPRS_NEG              = 0x10000,
    NUMPRS_INEXACT          = 0x20000,
} NUMPRS;

typedef struct NUMPARSE {
    LONG cDig;
    LONG dwInFlags;
    LONG dwOutFlags;
    LONG cchUsed;
    LONG nBaseShift;
    LONG nPwr10;
} NUMPARSE;

typedef enum NORM {
    NORM_IGNORECASE           = 0x00000001,
    NORM_IGNORENONSPACE       = 0x00000002,
    NORM_IGNORESYMBOLS        = 0x00000004,
    NORM_IGNOREWIDTH          = 0x00000008,
    NORM_IGNOREKANATYPE1      = 0x00000040,
    NORM_IGNOREKANATYPE       = 0x00010000,
    NORM_IGNOREWIDTH1         = 0x00020000,
    NORM_IGNOREKASHIDA        = 0x00040000,
} NORM;

typedef enum VARCMP {
    VARCMP_LT = 0,
    VARCMP_EQ = 1,
    VARCMP_GT = 2,
    VARCMP_NULL = 3 
} VARCMP;

typedef enum VCT_FLAGS {
    VARIANT_NOVALUEPROP = 0x01,
    VARIANT_ALPHABOOL = 0x02,      // For VT_BOOL to VT_BSTR conversions
                                   // convert to "True"/"False" instead of
                                   // "-1"/"0"
    VARIANT_NOUSEROVERRIDE = 0x04, // For conversions to/from VT_BSTR,,
                                   // passes LOCALE_NOUSEROVERRIDE
                                   // to core coercion routines
    VARIANT_CALENDAR_HIJRI = 0x08,
    VARIANT_LOCALBOOL = 0x10,      // For VT_BOOL to VT_BSTR and back
                                   // convert to local language rather than
                                   // English
} VCT_FLAGS;

typedef struct tagOBJECTDESCRIPTOR {
    LONG cbSize;              // Size of structure in bytes
    UUID clsid;               // CLSID of data being transferred
    DVASPECT dwDrawAspect;    // Display aspect of the object
                              //     normally DVASPECT_CONTENT or ICON.
                              //     dwDrawAspect will be 0 (which is NOT
                              //     DVASPECT_CONTENT) if the copier or
                              //     dragsource didn't draw the object to
                              //     begin with.
    SIZE sizel;               // size of the object in HIMETRIC
                              //    sizel is opt.: will be (0,0) for apps
                              //    which don't draw the object being
                              //    transferred
    POINT pointl;             // Offset in HIMETRIC units from the
                              //    upper-left corner of the obj where the
                              //    mouse went down for the drag.
                              //    NOTE: y coordinates increase downward.
                              //          x coordinates increase to right
                              //    pointl is opt.; it is only meaningful
                              //    if object is transfered via drag/drop.
                              //    (0, 0) if mouse position is unspecified
                              //    (eg. when obj transfered via clipboard)
    OLEMISC dwStatus;         // Misc. status flags for object. Flags are
                              //    defined by OLEMISC enum. these flags
                              //    are as would be returned
                              //    by IOleObject::GetMiscStatus.
    LONG dwFullUserTypeName;  // Offset from beginning of structure to
                              //    null-terminated string that specifies
                              //    Full User Type Name of the object.
                              //    0 indicates string not present.
    LONG dwSrcOfCopy;         // Offset from beginning of structure to
                              //    null-terminated string that specifies
                              //    source of the transfer.
                              //    dwSrcOfCOpy is normally implemented as
                              //    the display name of the temp-for-user
                              //    moniker which identifies the source of
                              //    the data.
                              //    0 indicates string not present.
                              //    NOTE: moniker assignment is NOT forced.
                              //    see IOleObject::GetMoniker(
                              //                OLEGETMONIKER_TEMPFORUSER)
} OBJECTDESCRIPTOR;

typedef struct NMHDR {
    LONG hwndFrom;
    LONG idFrom;
    LONG code;
} NMHDR;

//
// shellview select item flags
//
typedef enum SVSI_Flags {
    SVSI_DESELECT       = 0x0000,
    SVSI_SELECT         = 0x0001,
    SVSI_EDIT           = 0x0003, // includes select
    SVSI_DESELECTOTHERS = 0x0004,
    SVSI_ENSUREVISIBLE  = 0x0008,
    SVSI_FOCUSED        = 0x0010,
    SVSI_TRANSLATEPT    = 0x0020,
    SVSI_SELECTIONMARK	= 0x40,
	SVSI_POSITIONITEM	= 0x80,
	SVSI_CHECK	= 0x100,
	SVSI_CHECK2	= 0x200,
	SVSI_KEYBOARDSELECT	= 0x401,
	SVSI_NOTAKEFOCUS	= 0x40000000,
} SVSI_Flags;

typedef enum SV2GV_Flags {
    SV2GV_CURRENTVIEW = -1,
    SV2GV_DEFAULTVIEW = -2
} SV2GV_Flags;

//
// shellview get item object flags
//
typedef enum SVGIO_Flags {
    SVGIO_BACKGROUND = 0x00000000,
    SVGIO_SELECTION  = 0x00000001,
    SVGIO_ALLVIEW    = 0x00000002,
	SVGIO_CHECKED	= 0x3,
	SVGIO_TYPE_MASK	= 0xf,
	SVGIO_FLAG_VIEWORDER	= 0x80000000,
} SVGIO_Flags;