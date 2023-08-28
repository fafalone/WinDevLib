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

typedef enum ISCM_FLAGS {
    MUTZ_NOSAVEDFILECHECK = 0x00000001
} ISCM_FLAGS;

// MapUrlToZone returns the zone index given a URL
//     MAX_SIZE_SECURITY_ID 512 // bytes

typedef enum URLACTIONS {
    URLACTION_DOWNLOAD_MIN                                 = 0x00001000,
    URLACTION_DOWNLOAD_SIGNED_ACTIVEX                      = 0x00001001,
    URLACTION_DOWNLOAD_UNSIGNED_ACTIVEX                    = 0x00001004,
    URLACTION_DOWNLOAD_CURR_MAX                            = 0x00001004,
    URLACTION_DOWNLOAD_MAX                                 = 0x000011FF,

    URLACTION_ACTIVEX_MIN                                  = 0x00001200,
    URLACTION_ACTIVEX_RUN                                  = 0x00001200,
    URLACTION_ACTIVEX_OVERRIDE_OBJECT_SAFETY               = 0x00001201,
    URLACTION_ACTIVEX_OVERRIDE_DATA_SAFETY                 = 0x00001202,
    URLACTION_ACTIVEX_OVERRIDE_SCRIPT_SAFETY               = 0x00001203,
    URLACTION_SCRIPT_OVERRIDE_SAFETY                       = 0x00001401,
    URLACTION_ACTIVEX_CONFIRM_NOOBJECTSAFETY               = 0x00001204,
    URLACTION_ACTIVEX_TREATASUNTRUSTED                     = 0x00001205,
    URLACTION_ACTIVEX_CURR_MAX                             = 0x00001205,
    URLACTION_ACTIVEX_MAX                                  = 0x000013ff,

    URLACTION_SCRIPT_MIN                                   = 0x00001400,
    URLACTION_SCRIPT_RUN                                   = 0x00001400,
    URLACTION_SCRIPT_JAVA_USE                              = 0x00001402,
    URLACTION_SCRIPT_SAFE_ACTIVEX                          = 0x00001405,
    URLACTION_CROSS_DOMAIN_DATA                            = 0x00001406,
    URLACTION_SCRIPT_PASTE                                 = 0x00001407,
    URLACTION_SCRIPT_CURR_MAX                              = 0x00001407,
    URLACTION_SCRIPT_MAX                                   = 0x000015ff,

    URLACTION_HTML_MIN                                     = 0x00001600,
    URLACTION_HTML_SUBMIT_FORMS                            = 0x00001601,
    URLACTION_HTML_SUBMIT_FORMS_FROM                       = 0x00001602,
    URLACTION_HTML_SUBMIT_FORMS_TO                         = 0x00001603,
    URLACTION_HTML_FONT_DOWNLOAD                           = 0x00001604,
    URLACTION_HTML_JAVA_RUN                                = 0x00001605,
    URLACTION_HTML_USERDATA_SAVE                           = 0x00001606,
    URLACTION_HTML_SUBFRAME_NAVIGATE                       = 0x00001607,
    URLACTION_HTML_CURR_MAX                                = 0x00001607,
    URLACTION_HTML_MAX                                     = 0x000017ff,

    URLACTION_SHELL_MIN                                    = 0x00001800,
    URLACTION_SHELL_INSTALL_DTITEMS                        = 0x00001800,
    URLACTION_SHELL_MOVE_OR_COPY                           = 0x00001802,
    URLACTION_SHELL_FILE_DOWNLOAD                          = 0x00001803,
    URLACTION_SHELL_VERB                                   = 0x00001804,
    URLACTION_SHELL_WEBVIEW_VERB                           = 0x00001805,
    URLACTION_SHELL_CURR_MAX                               = 0x00001805,
    URLACTION_SHELL_MAX                                    = 0x000019ff,

    URLACTION_NETWORK_MIN                                  = 0x00001A00,

    URLACTION_CREDENTIALS_USE                              = 0x00001A00,

    URLACTION_AUTHENTICATE_CLIENT                          = 0x00001A01,

    URLACTION_COOKIES                                      = 0x00001A02,
    URLACTION_COOKIES_SESSION                              = 0x00001A03,

    URLACTION_CLIENT_CERT_PROMPT                           = 0x00001A04,

    URLACTION_NETWORK_CURR_MAX                             = 0x00001A04,
    URLACTION_NETWORK_MAX                                  = 0x00001Bff,

    URLACTION_JAVA_MIN                                     = 0x00001C00,
    URLACTION_JAVA_PERMISSIONS                             = 0x00001C00,
    URLACTION_JAVA_CURR_MAX                                = 0x00001C00,
    URLACTION_JAVA_MAX                                     = 0x00001Cff,

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


[
    odl,
    uuid(79eac9ef-baf9-11ce-8c82-00aa004ba90b),
]
interface IInternetZoneManager : stdole.IUnknown {

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
