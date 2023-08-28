//tbShellLib 100% implemented

[
    odl,
    uuid(79eac9ed-baf9-11ce-8c82-00aa004ba90b),
]
interface IInternetSecurityMgrSite : stdole.IUnknown {

    HRESULT GetWindow (
            [out, retval] LONG *phwnd);

    HRESULT EnableModeless (
            [in] BOOL fEnable);

};



// MapUrlToZone returns the zone index given a URL
//     MAX_SIZE_SECURITY_ID 512 // bytes

typedef enum URLACTIONS {
    URLACTION_DOWNLOAD_MIN                                 = 0x00001000,
    URLACTION_DOWNLOAD_SIGNED_ACTIVEX                      = 0x00001001,
    URLACTION_DOWNLOAD_UNSIGNED_ACTIVEX                    = 0x00001004,
    URLACTION_DOWNLOAD_CURR_MAX                            = 0x00001004,
    URLACTION_DOWNLOAD_MAX                                 = 0x000011FF,

     URLACTION_ACTIVEX_MIN                                 = 0x00001200,
     URLACTION_ACTIVEX_RUN                                 = 0x00001200,
     [helpstring("aggregate next four")]
     URLACTION_ACTIVEX_OVERRIDE_OBJECT_SAFETY              = 0x00001201,
     URLACTION_ACTIVEX_OVERRIDE_DATA_SAFETY                = 0x00001202,
     URLACTION_ACTIVEX_OVERRIDE_SCRIPT_SAFETY              = 0x00001203,
     URLACTION_SCRIPT_OVERRIDE_SAFETY                      = 0x00001401,
     URLACTION_ACTIVEX_CONFIRM_NOOBJECTSAFETY              = 0x00001204,
     URLACTION_ACTIVEX_TREATASUNTRUSTED                    = 0x00001205,
     URLACTION_ACTIVEX_NO_WEBOC_SCRIPT                     = 0x00001206,
     URLACTION_ACTIVEX_OVERRIDE_REPURPOSEDETECTION         = 0x00001207,
     URLACTION_ACTIVEX_OVERRIDE_OPTIN                      = 0x00001208,
     URLACTION_ACTIVEX_SCRIPTLET_RUN                       = 0x00001209,
     URLACTION_ACTIVEX_DYNSRC_VIDEO_AND_ANIMATION          = 0x0000120A,
     URLACTION_ACTIVEX_OVERRIDE_DOMAINLIST                 = 0x0000120B,
     URLACTION_ACTIVEX_CURR_MAX                            = 0x0000120B,
     URLACTION_ACTIVEX_MAX                                 = 0x000013ff,

     URLACTION_SCRIPT_MIN                                  = 0x00001400,
     URLACTION_SCRIPT_RUN                                  = 0x00001400,
     URLACTION_SCRIPT_JAVA_USE                             = 0x00001402,
     URLACTION_SCRIPT_SAFE_ACTIVEX                         = 0x00001405,
     URLACTION_CROSS_DOMAIN_DATA                           = 0x00001406,
     URLACTION_SCRIPT_PASTE                                = 0x00001407,
     URLACTION_ALLOW_XDOMAIN_SUBFRAME_RESIZE               = 0x00001408,
     URLACTION_SCRIPT_XSSFILTER                            = 0x00001409,
     URLACTION_SCRIPT_NAVIGATE                             = 0x0000140A,
     URLACTION_PLUGGABLE_PROTOCOL_XHR                      = 0x0000140B,
     URLACTION_SCRIPT_CURR_MAX                             = 0x0000140B,
     URLACTION_SCRIPT_MAX                                  = 0x000015ff,

     URLACTION_HTML_MIN                                    = 0x00001600,
     [helpstring("aggregate next two")]
     URLACTION_HTML_SUBMIT_FORMS                           = 0x00001601,
     URLACTION_HTML_SUBMIT_FORMS_FROM                      = 0x00001602,
     URLACTION_HTML_SUBMIT_FORMS_TO                        = 0x00001603,
     URLACTION_HTML_FONT_DOWNLOAD                          = 0x00001604,
     [helpstring("derive from Java custom policy")]
     URLACTION_HTML_JAVA_RUN                               = 0x00001605,
     URLACTION_HTML_USERDATA_SAVE                          = 0x00001606,
     URLACTION_HTML_SUBFRAME_NAVIGATE                      = 0x00001607,
     URLACTION_HTML_META_REFRESH                           = 0x00001608,
     URLACTION_HTML_MIXED_CONTENT                          = 0x00001609,
     URLACTION_HTML_INCLUDE_FILE_PATH                      = 0x0000160A,
     URLACTION_HTML_ALLOW_INJECTED_DYNAMIC_HTML            = 0x0000160B,
     URLACTION_HTML_REQUIRE_UTF8_DOCUMENT_CODEPAGE         = 0x0000160C,
     URLACTION_HTML_ALLOW_CROSS_DOMAIN_CANVAS              = 0x0000160D,
     URLACTION_HTML_ALLOW_WINDOW_CLOSE                     = 0x0000160E,
     URLACTION_HTML_ALLOW_CROSS_DOMAIN_WEBWORKER           = 0x0000160F,
     URLACTION_HTML_ALLOW_CROSS_DOMAIN_TEXTTRACK           = 0x00001610,
     URLACTION_HTML_ALLOW_INDEXEDDB                        = 0x00001611,
     URLACTION_HTML_MAX                                    = 0x000017ff,

    URLACTION_SHELL_MIN                                    = 0x00001800,
    URLACTION_SHELL_INSTALL_DTITEMS                        = 0x00001800,
    URLACTION_SHELL_MOVE_OR_COPY                           = 0x00001802,
    URLACTION_SHELL_FILE_DOWNLOAD                          = 0x00001803,
    URLACTION_SHELL_VERB                                   = 0x00001804,
    URLACTION_SHELL_WEBVIEW_VERB                           = 0x00001805,
    URLACTION_SHELL_SHELLEXECUTE                           = 0x00001806,
     URLACTION_SHELL_EXECUTE_HIGHRISK                      = 0x00001806,
     URLACTION_SHELL_EXECUTE_MODRISK                       = 0x00001807,
     URLACTION_SHELL_EXECUTE_LOWRISK                       = 0x00001808,
     URLACTION_SHELL_POPUPMGR                              = 0x00001809,
     URLACTION_SHELL_RTF_OBJECTS_LOAD                      = 0x0000180A,
     URLACTION_SHELL_ENHANCED_DRAGDROP_SECURITY            = 0x0000180B,
     URLACTION_SHELL_EXTENSIONSECURITY                     = 0x0000180C,
     URLACTION_SHELL_SECURE_DRAGSOURCE                     = 0x0000180D,
     URLACTION_SHELL_REMOTEQUERY                           = 0x0000180E,
     URLACTION_SHELL_PREVIEW                               = 0x0000180F,
     URLACTION_SHELL_SHARE                                 = 0x00001810,
     URLACTION_SHELL_ALLOW_CROSS_SITE_SHARE                = 0x00001811,
     URLACTION_SHELL_TOCTOU_RISK                           = 0x00001812,
     URLACTION_SHELL_CURR_MAX                              = 0x00001812,
    URLACTION_SHELL_MAX                                    = 0x000019ff,

    URLACTION_NETWORK_MIN                                  = 0x00001A00,

    URLACTION_CREDENTIALS_USE                              = 0x00001A00,

    URLACTION_AUTHENTICATE_CLIENT                          = 0x00001A01,

    URLACTION_COOKIES                                      = 0x00001A02,
    URLACTION_COOKIES_SESSION                              = 0x00001A03,

    URLACTION_CLIENT_CERT_PROMPT                           = 0x00001A04,

     URLACTION_COOKIES_THIRD_PARTY                         = 0x00001A05,
     URLACTION_COOKIES_SESSION_THIRD_PARTY                 = 0x00001A06,
     URLACTION_COOKIES_ENABLED                             = 0x00001A10,

    URLACTION_NETWORK_CURR_MAX                             = 0x00001A04,
    URLACTION_NETWORK_MAX                                  = 0x00001Bff,

     URLACTION_JAVA_MIN                                    = 0x00001C00,
     URLACTION_JAVA_PERMISSIONS                            = 0x00001C00,
     URLACTION_JAVA_CURR_MAX                               = 0x00001C00,
     URLACTION_JAVA_MAX                                    = 0x00001Cff,

    URLACTION_INFODELIVERY_MIN                           = 0x00001D00,
    URLACTION_INFODELIVERY_NO_ADDING_CHANNELS            = 0x00001D00,
    URLACTION_INFODELIVERY_NO_EDITING_CHANNELS           = 0x00001D01,
    URLACTION_INFODELIVERY_NO_REMOVING_CHANNELS          = 0x00001D02,
    URLACTION_INFODELIVERY_NO_ADDING_SUBSCRIPTIONS       = 0x00001D03,
    URLACTION_INFODELIVERY_NO_EDITING_SUBSCRIPTIONS      = 0x00001D04,
    URLACTION_INFODELIVERY_NO_REMOVING_SUBSCRIPTIONS     = 0x00001D05,
    URLACTION_INFODELIVERY_NO_CHANNEL_LOGGING            = 0x00001D06,
    URLACTION_INFODELIVERY_CURR_MAX                      = 0x00001D06,
    URLACTION_INFODELIVERY_MAX                           = 0x00001Dff,
    URLACTION_CHANNEL_SOFTDIST_MIN                       = 0x00001E00,
    URLACTION_CHANNEL_SOFTDIST_PERMISSIONS               = 0x00001E05,
    URLACTION_CHANNEL_SOFTDIST_MAX                       = 0x00001Eff,
    
     URLACTION_DOTNET_USERCONTROLS                       = 0x00002005,
     URLACTION_BEHAVIOR_MIN                              = 0x00002000,
     URLACTION_BEHAVIOR_RUN                              = 0x00002000,
     URLPOLICY_BEHAVIOR_CHECK_LIST                       = 0x00010000,
     
     URLACTION_FEATURE_MIN                               = 0x00002100,
     URLACTION_FEATURE_MIME_SNIFFING                     = 0x00002100,
     URLACTION_FEATURE_ZONE_ELEVATION                    = 0x00002101,
     URLACTION_FEATURE_WINDOW_RESTRICTIONS               = 0x00002102,
     URLACTION_FEATURE_SCRIPT_STATUS_BAR                 = 0x00002103,
     URLACTION_FEATURE_FORCE_ADDR_AND_STATUS             = 0x00002104,
     URLACTION_FEATURE_BLOCK_INPUT_PROMPTS               = 0x00002105,
     URLACTION_FEATURE_DATA_BINDING                      = 0x00002106,
     URLACTION_FEATURE_CROSSDOMAIN_FOCUS_CHANGE          = 0x00002107,
     
     URLACTION_AUTOMATIC_DOWNLOAD_UI_MIN                 = 0x00002200,
     URLACTION_AUTOMATIC_DOWNLOAD_UI                     = 0x00002200,
     URLACTION_AUTOMATIC_ACTIVEX_UI                      = 0x00002201,
     
     URLACTION_ALLOW_RESTRICTEDPROTOCOLS               = 0x00002300,
     URLACTION_ALLOW_APEVALUATION                      = 0x00002301,
     URLACTION_ALLOW_XHR_EVALUATION                    = 0x00002302,
     
     URLACTION_WINDOWS_BROWSER_APPLICATIONS            = 0x00002400,
     URLACTION_XPS_DOCUMENTS                           = 0x00002401,
     URLACTION_LOOSE_XAML                              = 0x00002402,
     
     URLACTION_LOWRIGHTS                               = 0x00002500,
     
     URLACTION_WINFX_SETUP                             = 0x00002600,
     
     URLACTION_INPRIVATE_BLOCKING                      = 0x00002700,
     URLACTION_ALLOW_AUDIO_VIDEO                       = 0x00002701,
     URLACTION_ALLOW_ACTIVEX_FILTERING                 = 0x00002702,
     URLACTION_ALLOW_STRUCTURED_STORAGE_SNIFFING       = 0x00002703,
     URLACTION_ALLOW_AUDIO_VIDEO_PLUGINS               = 0x00002704,
     URLACTION_ALLOW_ZONE_ELEVATION_VIA_OPT_OUT        = 0x00002705,
     URLACTION_ALLOW_ZONE_ELEVATION_OPT_OUT_ADDITION   = 0x00002706,
     URLACTION_ALLOW_CROSSDOMAIN_DROP_WITHIN_WINDOW    = 0x00002708,
     URLACTION_ALLOW_CROSSDOMAIN_DROP_ACROSS_WINDOWS   = 0x00002709,
     URLACTION_ALLOW_CROSSDOMAIN_APPCACHE_MANIFEST     = 0x0000270A,
     URLACTION_ALLOW_RENDER_LEGACY_DXTFILTERS          = 0x0000270B,
     URLACTION_ALLOW_ANTIMALWARE_SCANNING_OF_ACTIVEX   = 0x0000270C,
     URLACTION_ALLOW_CSS_EXPRESSIONS                   = 0x0000270D,

} URLACTIONS;

typedef enum URLPOLICIES {
    URLPOLICY_ACTIVEX_CHECK_LIST                 = 0x00010000,

    URLPOLICY_CREDENTIALS_SILENT_LOGON_OK        = 0x00000000,
    URLPOLICY_CREDENTIALS_MUST_PROMPT_USER       = 0x00010000,
    URLPOLICY_CREDENTIALS_CONDITIONAL_PROMPT     = 0x00020000,
    URLPOLICY_CREDENTIALS_ANONYMOUS_ONLY         = 0x00030000,

    URLPOLICY_AUTHENTICATE_CLEARTEXT_OK          = 0x00000000,
    URLPOLICY_AUTHENTICATE_CHALLENGE_RESPONSE    = 0x00010000,
    URLPOLICY_AUTHENTICATE_MUTUAL_ONLY           = 0x00030000,

    URLPOLICY_JAVA_PROHIBIT                      = 0x00000000,
    URLPOLICY_JAVA_HIGH                          = 0x00010000,
    URLPOLICY_JAVA_MEDIUM                        = 0x00020000,
    URLPOLICY_JAVA_LOW                           = 0x00030000,
    URLPOLICY_JAVA_CUSTOM                        = 0x00800000,

    URLPOLICY_CHANNEL_SOFTDIST_PROHIBIT          = 0x00010000,
    URLPOLICY_CHANNEL_SOFTDIST_PRECACHE          = 0x00020000,
    URLPOLICY_CHANNEL_SOFTDIST_AUTOINSTALL       = 0x00030000,

    URLPOLICY_ALLOW                = 0x00,
    URLPOLICY_QUERY                = 0x01,
    URLPOLICY_DISALLOW             = 0x03,
    URLPOLICY_NOTIFY_ON_ALLOW      = 0x10,
    URLPOLICY_NOTIFY_ON_DISALLOW   = 0x20,
    URLPOLICY_LOG_ON_ALLOW         = 0x40,
    URLPOLICY_LOG_ON_DISALLOW      = 0x80,
    URLPOLICY_MASK_PERMISSIONS     = 0x0f,
    URLPOLICY_DONTCHECKDLGBOX     = 0x100,

} URLPOLICIES;
typedef enum MUTZ_FLAGS {
	MUTZ_NOSAVEDFILECHECK       = 0x00000001, // don't check file: for saved file comment")
	MUTZ_ISFILE                 = 0x00000002, // Assume URL if File, url does not need file://")
	MUTZ_ACCEPT_WILDCARD_SCHEME = 0x00000080, // Accept a wildcard scheme")
	MUTZ_ENFORCERESTRICTED      = 0x00000100, // enforce restricted zone independent of URL")
	MUTZ_RESERVED               = 0x00000200, // This is same as PUAF_NOSAVEDFILECHECK. However we already have MUTZ_NOSAVEDFILECHECK for this.")
	MUTZ_REQUIRESAVEDFILECHECK  = 0x00000400, // always check the file for MOTW (overriding FEATURE_UNC_SAVEDFILECHECK)")
	MUTZ_DONT_UNESCAPE          = 0x00000800, // Do not unescape the url")
	MUTZ_DONT_USE_CACHE         = 0x00001000, // Do not check the cache")
	MUTZ_FORCE_INTRANET_FLAGS   = 0x00002000, // Force the intranet flags to be active")
	MUTZ_IGNORE_ZONE_MAPPINGS   = 0x00004000 // Don't look up the Zone Mappings
} MUTZ_FLAGS;

[
    odl,
    uuid(79eac9ee-baf9-11ce-8c82-00aa004ba90b),
]
interface IInternetSecurityManager : stdole.IUnknown {

    HRESULT SetSecuritySite(
        [in] IInternetSecurityMgrSite *pSite);

    HRESULT GetSecuritySite(
        [out, retval] IInternetSecurityMgrSite **ppSite);

    HRESULT MapUrlToZone(
        [in] LONG pwszUrl,
        [in, out] LONG *pdwZone,
        [in] LONG dwFlags);

    HRESULT GetSecurityId(
        [in] LONG pwszUrl,
        [in] LONG pbSecurityId,
        [in, out] LONG *pcbSecurityId,
        [in] LONG dwReserved);

    // Flags passed into ProcessUrlAction.
    typedef enum {
        PUAF_DEFAULT                        = 0x00000000,
        PUAF_NOUI                           = 0x00000001,
        PUAF_ISFILE                         = 0x00000002,  // Assume URL if File, url does not need file://
        PUAF_WARN_IF_DENIED                 = 0x00000004,
        PUAF_FORCEUI_FOREGROUND             = 0x00000008,
        PUAF_CHECK_TIFS                     = 0x00000010,
        PUAF_DONTCHECKBOXINDIALOG           = 0x00000020,
        PUAF_TRUSTED                        = 0x00000040,
        PUAF_ACCEPT_WILDCARD_SCHEME         = 0x00000080,
        PUAF_ENFORCERESTRICTED              = 0x00000100,
        [helpstring("don't check file: for saved comment")]
        PUAF_NOSAVEDFILECHECK               = 0x00000200,
        [helpstring("always check the file for MOTW (overriding FEATURE_UNC_SAVEDFILECHECK)")]
        PUAF_REQUIRESAVEDFILECHECK          = 0x00000400,
        [helpstring("don't use the MUTZ zone cache - recompute the zone.")]
        PUAF_DONT_USE_CACHE                 = 0x00001000,
        [helpstring("RESERVED for MUTZ_FORCE_INTRANET_FLAGS.")]
        PUAF_RESERVED1                      = 0x00002000,
        [helpstring("RESERVED for MUTZ_IGNORE_ZONE_MAPPINGS.")]
        PUAF_RESERVED2                      = 0x00004000,
        [helpstring("flag to retrieve local machine policies when LMZL is on")]
        PUAF_LMZ_UNLOCKED                   = 0x00010000,
        [helpstring("flag to force retrieving locked down local machine policies when LMZL is on")]
        PUAF_LMZ_LOCKED                     = 0x00020000,
        [helpstring("flag to retrieve Default Security Zone policies when NPL is on")]
        PUAF_DEFAULTZONEPOL                 = 0x00040000,
        [helpstring("flag to tell urlmon that trident wants to use locked zones if needed.")]
        PUAF_NPL_USE_LOCKED_IF_RESTRICTED   = 0x00080000,
        [helpstring("flag to turn off UI if Urlmon figures it is using the locked zone.")]
        PUAF_NOUIIFLOCKED                   = 0x00100000,
        [helpstring("flag to check for allow-list drag protcols when URLACTION_SHELL_MOVE_OR_COPY is allow")]
        PUAF_DRAGPROTOCOLCHECK              = 0x00200000,
    } PUAF ;

    HRESULT ProcessUrlAction(
        [in] LONG pwszUrl,
        [in] URLACTIONS dwAction,
        [in] LONG pPolicy,
        [in] LONG cbPolicy,
        [in, out] BYTE *pContext,
        [in] LONG cbContext,
        [in] PUAF dwFlags,
        [in] LONG dwReserved);

    HRESULT QueryCustomPolicy(
        [in] LONG pwszUrl,
        [in, out] UUID *guidKey,
        [in, out] LONG *ppPolicy,   // pointer to output buffer pointer
        [in, out] LONG *pcbPolicy,  // pointer to output buffer size
        [in, out] BYTE *pContext,        // context (used by the delegation routines)
        [in] LONG cbContext,  // size of the Context
        [in, defaultvalue(0)] LONG dwReserved);

    typedef enum {
        SZM_CREATE = 0x00000000, // create new mapping, error's if mapping already exists.
        SZM_DELETE = 0x00000001, // delete the mapping
    } SZM_FLAGS;

    HRESULT SetZoneMapping(
        [in] LONG dwZone,        // absolute zone index
        [in] LONG lpszPattern,   // URL pattern with limited wildcarding
        [in] SZM_FLAGS dwFlags);  // add, change, delete

    HRESULT GetZoneMappings (
        [in] LONG dwZone,                        // absolute zone index
        [in, out] IEnumString  **ppenumString,   // output buffer size
        [in] LONG dwFlags);                      // reserved, pass 0
}

[
	odl,
    uuid(F164EDF1-CC7C-4f0d-9A94-34222625C393),
    helpstring("IInternetSecurityManagerEx Interface")
]
interface IInternetSecurityManagerEx : IInternetSecurityManager
{
    HRESULT     ProcessUrlActionEx
    (
        [in]    LPCWSTR     pwszUrl,
        [in]    DWORD       dwAction,
        [in]
                long    pPolicy,    // output buffer pointer
        [in]    DWORD   cbPolicy,   // output buffer size
        [in, out]    BYTE*   pContext,   // context (used by the delegation routines)
        [in]    DWORD   cbContext,  // size of the Context
        [in]    DWORD   dwFlags,    // See enum PUAF for details.
        [in]    DWORD   dwReserved,
        [out]   DWORD   *pdwOutFlags //See enum PUAF for details
    );
}

[
	odl,
    uuid(F1E50292-A795-4117-8E09-2B560A72AC60),
    helpstring("IInternetSecurityManagerEx2 Interface")
]
interface IInternetSecurityManagerEx2 : IInternetSecurityManagerEx
{
    HRESULT MapUrlToZoneEx2
    (
        [in]
                IUri*       pUri,
        [out]   DWORD*      pdwZone,
        [in]    DWORD       dwFlags,
        [out]   long*       ppwszMappedUrl,
        [out]   DWORD*      pdwOutFlags
    );

    HRESULT ProcessUrlActionEx2
    (
        [in]    IUri*       pUri,
        [in]    DWORD       dwAction,
        [in]    long       pPolicy,    // output buffer pointer
        [in]    DWORD       cbPolicy,   // output buffer size
        [in, out]
                BYTE*       pContext,   // context (used by the delegation routines)
        [in]    DWORD       cbContext,  // size of the Context
        [in]    DWORD       dwFlags,    // See enum PUAF for details.
        [in]    DWORD       dwReserved,
        [out]   DWORD*      pdwOutFlags
    );

    HRESULT GetSecurityIdEx2
    (
        [in]    IUri*       pUri,
        [in]    long       pbSecurityId,
        [in, out]
                DWORD*      pcbSecurityId,
        [in]    DWORD_PTR   dwReserved
    );
    HRESULT QueryCustomPolicyEx2
    (
        [in]
                IUri*       pUri,
        [in]    UUID*       guidKey,
        [in]     // allocation via IMemAlloc; caller frees
                long        ppPolicy,   // pointer to output buffer pointer
        [out]   DWORD*      pcbPolicy,  // pointer to output buffer size
        [in, out]    BYTE*       pContext,   // context (used by the delegation routines)
        [in]    DWORD       cbContext,  // size of the Context
        [in]    DWORD      dwReserved
    );
};

[
    odl,
    uuid(3af280b6-cb3f-11d0-891e-00c04fb6bfc4),
]
interface IInternetHostSecurityManager : stdole.IUnknown {

    HRESULT GetSecurityId(
        [in, out] BYTE *pbSecurityId,
        [in, out] LONG *pcbSecurityId,
        [in, defaultvalue(0)] LONG dwReserved);

    HRESULT ProcessUrlAction(
        [in] URLACTIONS dwAction,
        [in, out] BYTE *pPolicy,    // output buffer pointer
        [in] LONG cbPolicy,   // output buffer size
        [in, out] BYTE *pContext,   // context (used by the delegation routines)
        [in] LONG cbContext,  // size of the Context
        [in] PUAF dwFlags,    // See enum PUAF for details.
        [in, defaultvalue(0)] LONG dwReserved);

    HRESULT QueryCustomPolicy (
        [in, out] UUID *guidKey,
        [in, out] LONG *ppPolicy,   // pointer to output buffer pointer
        [in, out] LONG *pcbPolicy,  // pointer to output buffer size
        [in, out] BYTE *pContext,   // context (used by the delegation routines)
        [in] LONG cbContext,  // size of the Context
        [in, defaultvalue(0)] LONG dwReserved);

};

    typedef enum URLZONE {
        URLZONE_PREDEFINED_MIN = 0,
        URLZONE_LOCAL_MACHINE  = 0,         // local machine zone is not exposed in UI")
        URLZONE_INTRANET,                   // My Intranet zone")
        URLZONE_TRUSTED,                    // Trusted Web sites zone")
        URLZONE_INTERNET,                   // The Internet zone")
        URLZONE_UNTRUSTED,                  // Untrusted sites zone
        URLZONE_PREDEFINED_MAX = 999,
        URLZONE_USER_MIN = 1000,
        URLZONE_USER_MAX = 10000,
    }  URLZONE;

    typedef enum URLTEMPLATE {

        // This value is just used to indicate the current set
        // of policies are not based on any template.
        URLTEMPLATE_CUSTOM  =   0x000000,
        URLTEMPLATE_PREDEFINED_MIN =    0x10000,
        URLTEMPLATE_LOW =               0x10000,
        URLTEMPLATE_MEDLOW =            0x10500,
        URLTEMPLATE_MEDIUM =            0x11000,
        URLTEMPLATE_MEDHIGH =           0x11500,
        URLTEMPLATE_HIGH  =             0x12000,
        URLTEMPLATE_PREDEFINED_MAX =    0x20000
    } URLTEMPLATE ;

    typedef enum ZAFLAGS{
        ZAFLAGS_CUSTOM_EDIT                 = 0x00000001,
        ZAFLAGS_ADD_SITES                   = 0x00000002,
        ZAFLAGS_REQUIRE_VERIFICATION        = 0x00000004,
        ZAFLAGS_INCLUDE_PROXY_OVERRIDE      = 0x00000008,  // Intranet only.
        ZAFLAGS_INCLUDE_INTRANET_SITES      = 0x00000010,  // Intranet only.
        ZAFLAGS_NO_UI                       = 0x00000020,  // Don't display UI (used for local machine)
        ZAFLAGS_SUPPORTS_VERIFICATION       = 0x00000040,  // Supports server verification.
        ZAFLAGS_UNC_AS_INTRANET             = 0x00000080,
        ZAFLAGS_DETECT_INTRANET             = 0x00000100,  // Intranet only.

        // Locked/Unlocked state specific flags.
        ZAFLAGS_USE_LOCKED_ZONES            = 0x00010000,
        // Used ONLY in GetZoneAttributes to specify that Template Matching should be done to verify
        // that zone's Current Level is correct.
        ZAFLAGS_VERIFY_TEMPLATE_SETTINGS    = 0x00020000,
        // Bypass the zonemgr cache for this setting
        ZAFLAGS_NO_CACHE                    = 0x00040000,
    } ZAFLAGS;

    typedef struct ZONEATTRIBUTES {
        LONG cbSize;
        BYTE szDisplayName[520];
        BYTE szDescription[400];
        BYTE szIconPath[520];
        URLTEMPLATE dwTemplateMinLevel;
        URLTEMPLATE dwTemplateRecommended;
        URLTEMPLATE dwTemplateCurrentLevel;
        ZAFLAGS dwFlags;                     // ZAFLAGS.
    } ZONEATTRIBUTES;
[
    odl,
    uuid(79eac9ef-baf9-11ce-8c82-00aa004ba90b),
]
interface IInternetZoneManager : stdole.IUnknown {



    HRESULT GetZoneAttributes(
        [in] URLZONE dwZone,
        [in, out] ZONEATTRIBUTES *pZoneAttributes);

    HRESULT SetZoneAttributes(
        [in] URLZONE dwZone,
        [in] ZONEATTRIBUTES *pZoneAttributes);

    typedef enum URLZONEREG {
        URLZONEREG_DEFAULT=0,
        URLZONEREG_HKLM,
        URLZONEREG_HKCU
    } URLZONEREG;

    HRESULT GetZoneCustomPolicy(
        [in] URLZONE dwZone,         // zone index
        [in] UUID *guidKey,          // key to lookup value
        [in, out] void *ppPolicy,    // pointer to output buffer pointer
        [in, out] LONG *pcbPolicy,   // pointer to output buffer size
        [in] URLZONEREG urlZoneReg); // effective, HKCU, or HKLM

    HRESULT SetZoneCustomPolicy(
        [in] URLZONE dwZone,         // zone index
        [in] UUID *guidKey,          // key to lookup value
        [in, out] void *pPolicy,     // input buffer pointer
        [in] LONG cbPolicy,          // input data size
        [in] URLZONEREG urlZoneReg); // default, HKCU, or HKLM

    HRESULT GetZoneActionPolicy(
        [in] URLZONE dwZone,        // zone index
        [in] URLACTIONS dwAction,   // index number of action
        [in, out] void *pPolicy,    // output buffer pointer
        [in] LONG cbPolicy,         // output buffer size
        [in] URLZONEREG urlZoneReg);// effective, HKCU, or HKLM

    HRESULT SetZoneActionPolicy (
        [in] URLZONE dwZone,         // zone index
        [in] URLACTIONS dwAction,    // index number of action
        [in, out] void *pPolicy,     // input buffer pointer
        [in] LONG cbPolicy,          // input data size
        [in] URLZONEREG urlZoneReg); // HKCU, or HKLM

    HRESULT PromptAction(
        [in] LONG dwAction,       // action type
        [in] LONG hwndParent,     // parent window handle
        [in] LPWSTR pwszUrl,      // URL to display
        [in] LPWSTR pwszText,     // dialog text
        [in, defaultvalue(0)] LONG dwPromptFlags); // reserved, pass 0

    HRESULT LogAction(
        [in] URLACTIONS dwAction,  // action type
        [in] LPWSTR pwszUrl,       // URL to log
        [in] LPWSTR pwszText,      // associated text
        [in, defaultvalue(0)] LONG dwLogFlags);   // reserved, pass 0

    HRESULT CreateZoneEnumerator(
        [in, out] LONG *pdwEnumHandle,  // enum handle
        [in, out] LONG *pdwCount, // # of elements in the list.
        [in, defaultvalue(0)] LONG dwFlags);       // reserved, pass 0

    HRESULT GetZoneAt(
        [in] LONG dwEnumHandle,           // returned by CreateZoneEnumerator
        [in] LONG dwIndex,                // 0-based
        [out, retval] URLZONE *pdwZone);  // absolute zone index.

    HRESULT DestroyZoneEnumerator(
        [in] LONG dwEnumHandle);         // enum handle

    HRESULT CopyTemplatePoliciesToZone(
        [in] URLTEMPLATE dwTemplate,       // High, medium or low
        [in] URLZONE dwZone,        // Zone to copy policies to.
        [in, defaultvalue(0)] LONG dwReserved);

}

[
	odl,
    uuid(A4C23339-8E06-431e-9BF4-7E711C085648),
    helpstring("IInternetZoneManagerEx Interface")
]

interface IInternetZoneManagerEx : IInternetZoneManager
{
    HRESULT GetZoneActionPolicyEx
    (
        [in]    DWORD   dwZone,     // zone index
        [in]    DWORD   dwAction,   // index number of action
        [out]   long*   pPolicy,    // output buffer pointer - was BYTE*
        [in]    DWORD   cbPolicy,    // output buffer size
        [in]    URLZONEREG urlZoneReg, // effective, HKCU, or HKLM
        [in]    DWORD    dwFlags   //Lockdown Zones or Normal Zones via ZAFLAGS
    );

    HRESULT SetZoneActionPolicyEx
    (
        [in]    DWORD   dwZone,     // zone index
        [in]    DWORD   dwAction,   // index number of action
        [in]    long    pPolicy,    // input buffer pointer waas BYTE*
        [in]    DWORD   cbPolicy,    // input data size
        [in]    URLZONEREG urlZoneReg, // HKCU, or HKLM
        [in]    DWORD    dwFlags   //Lockdown Zones or Normal Zones via ZAFLAGS
    );
 }
 
[

    odl,
    uuid(EDC17559-DD5D-4846-8EEF-8BECBA5A4ABF),
    helpstring("IInternetZoneManagerEx2 Interface")
]

interface IInternetZoneManagerEx2 : IInternetZoneManagerEx
{
    HRESULT GetZoneAttributesEx
    (
        [in]    DWORD   dwZone,
        [in, out] ZONEATTRIBUTES* pZoneAttributes,
        [in]    DWORD    dwFlags // can only be ZAFLAGS_VERIFY_TEMPLATE_SETTINGS
    );

    HRESULT GetZoneSecurityState
    (
        [in] DWORD dwZoneIndex,
        [in] BOOL fRespectPolicy,
        [in, out] long* pdwState,
        [in, out] BOOL* pfPolicyEncountered
    );

    HRESULT GetIESecurityState
    (
        [in] BOOL fRespectPolicy,
        [in, out] long* pdwState,
        [in, out] BOOL* pfPolicyEncountered,
        [in] BOOL fNoCache
    );

    HRESULT FixUnsecureSettings();
}
