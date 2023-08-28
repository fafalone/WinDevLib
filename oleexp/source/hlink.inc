//tbShellLib - 100% iplemented
// hlink special reference ids used in HlinkSetSpecialReference and HlinkGetSpecialReference
typedef enum HLSR {
    HLSR_HOME = 0,
    HLSR_SEARCHPAGE = 1,
    HLSR_HISTORYFOLDER = 2,
} HLSR;

typedef enum HLSHORTCUTF {
    HLSHORTCUTF_DEFAULT             = 0x00000000,
    HLSHORTCUTF_DONTACTUALLYCREATE  = 0x00000001,
    HLSHORTCUTF_USEFILENAMEFROMFRIENDLYNAME = 0x00000002,
    HLSHORTCUTF_USEUNIQUEFILENAME   = 0x00000004,
    HLSHORTCUTF_MAYUSEEXISTINGSHORTCUT = 0x00000008,
} HLSHORTCUTF;

typedef enum HLTRANSLATEF {
    HLTRANSLATEF_DEFAULT                = 0x00000000,
    HLTRANSLATEF_DONTAPPLYDEFAULTPREFIX = 0x00000001
} HLTRANSLATEF;

[
	odl,
        uuid(79eac9c3-baf9-11ce-8c82-00aa004ba90b),
]
interface IHlink: stdole.IUnknown {

        // Navigation flags
        typedef enum HLNF {
            HLNF_INTERNALJUMP           = 0x00000001,
            HLNF_OPENINNEWWINDOW        = 0x00000002,
            HLNF_NAVIGATINGBACK         = 0x00000004,
            HLNF_NAVIGATINGFORWARD      = 0x00000008,
            HLNF_NAVIGATINGTOSTACKITEM  = 0x00000010,
            HLNF_CREATENOHISTORY        = 0x00000020,
        } HLNF;

        // GetXXXReference values indicating reference to be retrieved
        typedef enum HLINKGETREF {
            HLINKGETREF_DEFAULT  = 0,
            HLINKGETREF_ABSOLUTE = 1,
            HLINKGETREF_RELATIVE = 2,
        } HLINKGETREF;

        // GetFriendlyName flags
        typedef enum HLFNAMEF {
            HLFNAMEF_DEFAULT            = 0x00000000,
            HLFNAMEF_TRYCACHE           = 0x00000001,
            HLFNAMEF_TRYPRETTYTARGET    = 0x00000002,
            HLFNAMEF_TRYFULLTARGET      = 0x00000004,
            HLFNAMEF_TRYWIN95SHORTCUT   = 0x00000008,
        } HLFNAMEF;

        // Status flags
        typedef enum HLINKMISC {
            HLINKMISC_RELATIVE = 0x00000001,
        } HLINKMISC;

        // SetXXXReference flags, indicating which part of the reference
        // need to be updated
        typedef enum HLINKSETF {
            HLINKSETF_TARGET              = 0x00000001,
            HLINKSETF_LOCATION            = 0x00000002,
        } HLINKSETF;

        HRESULT SetHlinkSite(
                [in] IHlinkSite *pihlSite,
                [in] LONG dwSiteData);

        HRESULT GetHlinkSite(
                [out] IHlinkSite **ppihlSite,
                [out] LONG *pdwSiteData);

        HRESULT SetMonikerReference(
                [in] HLINKSETF grfHLSETF,
                [in] IMoniker *pimkTarget,
                [in] LPWSTR pwzLocation);

        HRESULT GetMonikerReference(
                [in] HLINKGETREF dwWhichRef,
                [out] IMoniker **ppimkTarget,
                [out] LONG *ppwzLocation);

        HRESULT SetStringReference(
                [in] HLINKSETF grfHLSETF,
                [in] LPWSTR pwzTarget,
                [in] LPWSTR pwzLocation);

        HRESULT GetStringReference(
                [in] HLINKGETREF dwWhichRef,
                [out] LONG *ppwzTarget,
                [out] LONG *ppwzLocation);

        HRESULT SetFriendlyName(
                [in] LPWSTR pwzFriendlyName);

        HRESULT GetFriendlyName(
                [in] HLFNAMEF grfHLFNAMEF,
                [out, retval] LONG *ppwzFriendlyName);

        HRESULT SetTargetFrameName(
                [in] LPWSTR pwzTargetFrameName);

        HRESULT GetTargetFrameName(
                [out, retval] LONG *ppwzTargetFrameName);

        HRESULT GetMiscStatus(
                [out, retval] HLINKMISC *pdwStatus);

        HRESULT Navigate(
                [in] HLNF grfHLNF,
                [in] IBindCtx *pibc,
                [in] IBindStatusCallback *pibsc,
                [in] IHlinkBrowseContext *pihlbc);

        HRESULT SetAdditionalParams(
                [in] LPWSTR pwzAdditionalParams);

        HRESULT GetAdditionalParams(
                [out, retval] LONG *ppwzAdditionalParams);
}

[
	odl,
        uuid(79eac9c2-baf9-11ce-8c82-00aa004ba90b),
]
interface IHlinkSite: stdole.IUnknown {

        // values indicating the moniker to be retrieved
        typedef enum HLINKWHICHMK {
            HLINKWHICHMK_CONTAINER  = 1,
            HLINKWHICHMK_BASE       = 2,
        } HLINKWHICHMK;

        HRESULT QueryService(
                [in] LONG dwSiteData,
	        [in, out] UUID *guidService,
		[in, out] UUID *riid,
                [out] void *ppiunk);

        HRESULT GetMoniker(
                [in] LONG dwSiteData,
                [in] OLEGETMONIKER dwAssign,
                [in] OLEWHICHMK dwWhich,
                [out, retval] IMoniker ** ppimk);

        HRESULT ReadyToNavigate(
                [in] LONG dwSiteData,
                [in] LONG dwReserved);

        HRESULT OnNavigationComplete(
                [in] LONG dwSiteData,
                [in] LONG dwreserved,
                [in] HRESULT hrError,
                [in] LONG pwzError);
}

[
        odl,
        uuid(79eac9c4-baf9-11ce-8c82-00aa004ba90b),
]
interface IHlinkTarget : stdole.IUnknown {

        HRESULT SetBrowseContext(
                [in] IHlinkBrowseContext *pihlbc);

        HRESULT GetBrowseContext(
                [out, retval] IHlinkBrowseContext **ppihlbc);

        HRESULT Navigate(
                [in] HLNF grfHLNF,
                [in] LPWSTR pwzJumpLocation);

        HRESULT GetMoniker(
                [in] LPWSTR pwzLocation,
                [in] OLEGETMONIKER dwAssign,
                [out, retval] IMoniker ** ppimkLocation);

        HRESULT GetFriendlyName(
                [in] LPWSTR pwzLocation,
                [out, retval] LONG *ppwzFriendlyName);
}

[
        odl,
        uuid(79eac9c5-baf9-11ce-8c82-00aa004ba90b),
]
interface IHlinkFrame : stdole.IUnknown {

        HRESULT SetBrowseContext(
                [in] IHlinkBrowseContext * pihlbc);

        HRESULT GetBrowseContext(
                [out, retval] IHlinkBrowseContext ** ppihlbc);

        HRESULT Navigate(
                [in] HLNF grfHLNF,
                [in] IBindCtx *pbc,
                [in] IBindStatusCallback *pibsc,
                [in] IHlink *pihlNavigate);

        HRESULT OnNavigate(
                [in] HLNF grfHLNF,
                [in] IMoniker *pimkTarget,
                [in] LPWSTR pwzLocation,
                [in] LPWSTR pwzFriendlyName,
                [in] LONG dwreserved);

        HRESULT UpdateHlink(
                [in] LONG uHLID,
                [in] IMoniker *pimkTarget,
                [in] LPWSTR pwzLocation,
                [in] LPWSTR pwzFriendlyName);
}

[
        odl,
        uuid(79eac9c6-baf9-11ce-8c82-00aa004ba90b),
]
interface IEnumHLITEM : stdole.IUnknown {

        typedef struct HLITEM {
                LONG uHLID;
                LONG pwzFriendlyName;
        } HLITEM;

        LONG __stdcall Next(
                [in] LONG celt,
                [out] HLITEM *rgelt,
                [out, defaultvalue(0)] LONG *pceltFetched);

        HRESULT Skip(
                [in] LONG celt);

        HRESULT Reset();

        HRESULT Clone(
                [out, retval] IEnumHLITEM ** ppienumhlitem);
}

[
        odl,
        uuid(79eac9c7-baf9-11ce-8c82-00aa004ba90b),
]
interface IHlinkBrowseContext : stdole.IUnknown {

        // Information about the toolbars
        typedef enum HLTB {
            HLTB_DOCKEDLEFT     = 0,
            HLTB_DOCKEDTOP      = 1,
            HLTB_DOCKEDRIGHT    = 2,
            HLTB_DOCKEDBOTTOM   = 3,
            HLTB_FLOATING       = 4,
        } HLTB;

        typedef struct HLTBINFO {
            HLTB uDockType;
            RECT rcTbPos;
        } HLTBINFO;

        // information about the browse window
        typedef enum HLBWIF {
            HLBWIF_HASFRAMEWNDINFO      = 0x00000001,
            HLBWIF_HASDOCWNDINFO        = 0x00000002,
            HLBWIF_FRAMEWNDMAXIMIZED    = 0x00000004,
            HLBWIF_DOCWNDMAXIMIZED      = 0x00000008,
            HLBWIF_HASWEBTOOLBARINFO    = 0x00000010,
            HLBWIF_WEBTOOLBARHIDDEN     = 0x00000020
        } HLBWIF;

        typedef struct HLBWINFO {
            LONG cbSize;
            HLBWIF grfHLBWIF;
            RECT rcFramePos;
            RECT rcDocPos;
            HLTBINFO hltbinfo;
        } HLBWINFO;

        typedef enum HLID {
            HLID_INVALID        = 0x0,
            HLID_PREVIOUS       = 0xFFFFFFFF,
            HLID_NEXT           = 0xFFFFFFFE,
            HLID_CURRENT        = 0xFFFFFFFD,
            HLID_STACKBOTTOM    = 0xFFFFFFFC,
            HLID_STACKTOP       = 0xFFFFFFFB,
	} HLID;

        typedef enum HLQF {
            HLQF_ISVALID    = 0x00000001,
            HLQF_ISCURRENT  = 0x00000002
        } HLQF;

        HRESULT Register(
                [in] LONG reserved,
                [in] stdole.IUnknown *piunk,
                [in] IMoniker *pimk,
                [in, out] LONG *pdwRegister);

        HRESULT GetObject(
                [in] IMoniker *pimk,
                [in] BOOL fBindIfRootRegistered,
                [in, out] stdole.IUnknown **ppiunk);

        HRESULT Revoke(
                [in] LONG dwRegister);

        HRESULT SetBrowseWindowInfo(
                [in, out] HLBWINFO *phlbwi);

        HRESULT GetBrowseWindowInfo(
                [in, out] HLBWINFO *phlbwi);

        HRESULT SetInitialHlink(
                [in] IMoniker *pimkTarget,
                [in] LONG pwzLocation,
                [in] LONG pwzFriendlyName);

        HRESULT OnNavigateHlink(
                [in] HLNF grfHLNF,
                [in] IMoniker *pimkTarget,
                [in] LONG pwzLocation,
                [in] LONG pwzFriendlyName,
                [out, retval] HLID *puHLID);

        HRESULT UpdateHlink(
                [in] HLID uHLID,
                [in] IMoniker *pimkTarget,
                [in] LONG pwzLocation,
                [in] LONG pwzFriendlyName);

        HRESULT EnumNavigationStack(
                [in] LONG dwReserved,
                [in] HLFNAMEF grfHLFNAMEF,
                [in, out] IEnumHLITEM **ppienumhlitem);

        HRESULT QueryHlink(
                [in] HLQF grfHLQF,
                [in] HLID uHLID);

        HRESULT GetHlink(
                [in] HLID uHLID,
                [in, out] IHlink **ppihl);

        HRESULT SetCurrentHlink(
                [in] HLID uHLID);

        HRESULT Clone(
                [in] stdole.IUnknown *piunkOuter,
                [in, out] UUID *riid,
                [in, out] LONG *ppiunkObj);

        HRESULT Close(
                [in] LONG reserved);
}

[
	dllname("HLINK.DLL")
]
module hlink {

[entry("HlinkCreateFromMoniker")]
HRESULT HlinkCreateFromMoniker(                                                
	[in] IMoniker * pimkTrgt,                                             
	[in] LPWSTR pwzLocation,                                             
	[in] LPWSTR pwzFriendlyName,                                         
	[in] IHlinkSite * pihlsite,                                           
	[in] LONG dwSiteData,                                                
	[in] IUnknown * piunkOuter,                                           
	[in] UUID *riid,                                                     
	[out] void *ppvObj);                                                 
                                                                              
[entry("HlinkCreateFromString")]
HRESULT HlinkCreateFromString(                                                 
	[in] LPWSTR pwzTarget,                                               
	[in] LPWSTR pwzLocation,                                             
	[in] LPWSTR pwzFriendlyName,                                         
	[in] IHlinkSite *pihlsite,                                           
	[in] LONG dwSiteData,                                                
	[in] IUnknown *piunkOuter,                                           
	[in] UUID *riid,                                                     
	[out] void *ppvObj);                                                 
                                                                              
[entry("HlinkCreateFromData")]
HRESULT HlinkCreateFromData(                                                   
	[in] IDataObject *piDataObj,                                          
	[in] IHlinkSite * pihlsite,                                           
	[in] LONG dwSiteData,                                                
	[in] IUnknown * piunkOuter,                                           
	[in] UUID *riid,                                                     
	[out] void *ppvObj);                                                 
                                                                              
[entry("HlinkQueryCreateFromData")]
HRESULT HlinkQueryCreateFromData(
		[in] IDataObject *piDataObj);
                                                                              
[entry("HlinkClone")]
HRESULT HlinkClone(                                                            
	[in] IHlink * pihl,                                                   
	[in] UUID *riid,                                                     
	[in] IHlinkSite * pihlsiteForClone,                                   
	[in] LONG dwSiteData,                                                
	[out] void *ppvObj);                                                 
                                                                              
[entry("HlinkCreateBrowseContext")]
HRESULT HlinkCreateBrowseContext(                                              
	[in] IUnknown *piunkOuter,                                           
	[in] UUID *riid,                                                     
	[out] void *ppvObj);                                                 
                                                                              
[entry("HlinkNavigateToStringReference")]
HRESULT HlinkNavigateToStringReference(                                        
	[in] LPWSTR pwzTarget,                                               
	[in] LPWSTR pwzLocation,                                             
	[in] IHlinkSite *pihlsite,                                           
	[in] LONG dwSiteData,                                                
	[in] IHlinkFrame *pihlframe,                                          
	[in] HLNF grfHLNF,                                                   
	[in] IBindCtx *pibc,                                                       
	[in] IBindStatusCallback * pibsc,                                     
	[in] IHlinkBrowseContext *pihlbc);                                    
                                                                              
[entry("HlinkNavigate")]
HRESULT HlinkNavigate(                                                         
	[in] IHlink * pihl,                                                   
	[in] IHlinkFrame * pihlframe,                                         
	[in] HLNF grfHLNF,                                                   
	[in] IBindCtx *pbc,                                                        
	[in] IBindStatusCallback * pibsc,                                     
	[in] IHlinkBrowseContext *pihlbc);                                    
                                                                              
[entry("HlinkOnNavigate")]
HRESULT HlinkOnNavigate(                                                       
	[in] IHlinkFrame * pihlframe,                                         
	[in] IHlinkBrowseContext * pihlbc,                                    
	[in] HLNF grfHLNF,                                                   
	[in] IMoniker *pimkTarget,                                           
	[in] LPWSTR pwzLocation,                                             
	[in] LPWSTR pwzFriendlyName,                                         
	[out, retval] HLID *puHLID);                                                 
                                                                              
[entry("HlinkUpdateStackItem")]
HRESULT HlinkUpdateStackItem(                                                  
	[in] IHlinkFrame *pihlframe,                                         
	[in] IHlinkBrowseContext *pihlbc,                                    
	[in] HLID uHLID,                                                     
	[in] IMoniker * pimkTrgt,                                             
	[in] LPWSTR pwzLocation,                                             
	[in] LPWSTR pwzFriendlyName);                                        
                                                                              
[entry("HlinkOnRenameDocument")]
HRESULT HlinkOnRenameDocument(                                                 
	[in] LONG dwReserved,                                                
	[in] IHlinkBrowseContext *pihlbc,                                    
	[in] IMoniker *pimkOld,                                              
	[in] IMoniker *pimkNew);                                             
                                                                              
[entry("HlinkResolveMonikerForData")]
HRESULT HlinkResolveMonikerForData(                                            
	[in] IMoniker *pimkReference,                                         
	[in] LONG reserved,                                                  
	[in] IBindCtx *pibc,                                                       
	[in] LONG cFmtetc,                                                   
	[in] FORMATETC *rgFmtetc,                                            
	[in] IBindStatusCallback * pibsc,                                     
	[in] IMoniker *pimkBase);                                             
                                                                              
[entry("HlinkResolveStringForData")]
HRESULT HlinkResolveStringForData(                                             
	[in] LPWSTR pwzReference,                                            
	[in] LONG reserved,                                                  
	[in] IBindCtx *pibc,                                                       
	[in] LONG cFmtetc,                                                   
	[in] FORMATETC * rgFmtetc,                                            
	[in] IBindStatusCallback * pibsc,                                     
	[in] IMoniker *pimkBase);                                             
                                                                              
[entry("HlinkParseDisplayName")]
HRESULT HlinkParseDisplayName(                                                 
	[in] IBindCtx *pibc,                                                       
	[in] LPWSTR pwzDisplayName,                                          
	[in] BOOL fNoForceAbs,                                                
	[in] LONG *pcchEaten,                                               
	[out, retval] IMoniker **ppimk);                                              
                                                                              
[entry("HlinkCreateExtensionServices")]
HRESULT HlinkCreateExtensionServices(                                          
	[in] LPWSTR pwzAdditionalHeaders,                                    
	[in] LONG phwnd,                                                      
	[in] LPWSTR pszUsername,                                             
	[in] LPWSTR pszPassword,                                             
	[in] IUnknown *piunkOuter,                                           
	[in] UUID *riid,                                                     
	[out] void *ppvObj);                                                 
                                                                              
[entry("HlinkPreprocessMoniker")]
HRESULT HlinkPreprocessMoniker(                                                
	[in] IBindCtx *pibc,                                                       
	[in] IMoniker *pimkIn,                                                
	[out, retval] IMoniker **ppimkOut);                                            
                                                                              
[entry("OleSaveToStreamEx")]
HRESULT OleSaveToStreamEx(                                                     
	[in] IUnknown *piunk,                                                
	[in] IStream *pistm,                                                 
	[in] BOOL fClearDirty);                                               
                                                                              
[entry("HlinkSetSpecialReference")]
HRESULT HlinkSetSpecialReference(                                              
	[in] HLSR uReference,                                                
	[in] LPWSTR pwzReference);                                           
                                                                              
[entry("HlinkGetSpecialReference")]
HRESULT HlinkGetSpecialReference(                                              
	[in] HLSR uReference,                                                
	[out, retval] LONG *ppwzReference);                                          

[entry("HlinkCreateShortcut")]
HRESULT HlinkCreateShortcut(                                                   
	[in] HLSHORTCUTF grfHLSHORTCUTF,
	[in] IHlink *pihl,                                                    
	[in] LPWSTR pwzDir,                                                  
	[in] LPWSTR pwzFileName,                                             
	[in] LPWSTR *ppwzShortcutFile,                                        
	[in] LONG dwReserved);                                               
                                                                              
[entry("HlinkCreateShortcutFromMoniker")]
HRESULT HlinkCreateShortcutFromMoniker(                                        
	[in] HLSHORTCUTF grfHLSHORTCUTF,                                            
	[in] IMoniker *pimkTarget,                                            
	[in] LPWSTR pwzLocation,                                             
	[in] LPWSTR pwzDir,                                                  
	[in] LPWSTR pwzFileName,                                             
	[out] LONG *ppwzShortcutFile,                                        
	[in] LONG dwReserved);                                               
                                                                              
[entry("HlinkCreateShortcutFromString")]
HRESULT HlinkCreateShortcutFromString(                                         
	[in] HLSHORTCUTF grfHLSHORTCUTF,                                            
	[in] LPWSTR pwzTarget,                                               
	[in] LPWSTR pwzLocation,                                             
	[in] LPWSTR pwzDir,                                                  
	[in] LPWSTR pwzFileName,                                             
	[out] LONG *ppwzShortcutFile,                                        
	[in] LONG dwReserved);                                               

[entry("HlinkResolveShortcut")]
HRESULT HlinkResolveShortcut(                                                  
	[in] LPWSTR pwzShortcutFileName,                                     
	[in] IHlinkSite * pihlsite,                                           
	[in] LONG dwSiteData,                                                
	[in] IUnknown * piunkOuter,                                           
	[in] UUID *riid,                                                     
	[out] void *ppvObj);                                                 
                                                                              
[entry("HlinkResolveShortcutToMoniker")]
HRESULT HlinkResolveShortcutToMoniker(                                         
	[in] LPWSTR pwzShortcutFileName,                                     
	[out] IMoniker **ppimkTarget,                                          
	[out] LONG *ppwzLocation);                                           
                                                                              
[entry("HlinkResolveShortcutToString")]
HRESULT HlinkResolveShortcutToString(                                          
	[in] LPWSTR pwzShortcutFileName,                                     
	[out] LONG *ppwzTarget,                                              
	[out] LONG *ppwzLocation);                                           
                                                                              
[entry("HlinkIsShortcut")]
HRESULT HlinkIsShortcut(
	[in] LPWSTR pwzFileName);                                 
                                                                              
[entry("HlinkGetValueFromParams")]
HRESULT HlinkGetValueFromParams(                                               
	[in] LPWSTR pwzParams,                                               
	[in] LPWSTR pwzName,                                                 
	[out] LONG *ppwzValue);                                              

[entry("HlinkTranslateURL")]
HRESULT HlinkTranslateURL(                                                     
	[in] LPWSTR pwzURL,                                                  
	[in] LONG grfFlags,                                                  
	[out, retval] LONG *ppwzTranslatedURL);
	
}