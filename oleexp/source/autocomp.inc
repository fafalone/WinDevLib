//tbShellLib - 100% implemented

[
    odl,
    uuid(00bb2762-6a77-11d0-a535-00c04fd7d062)
]
interface IAutoComplete : stdole.IUnknown {

    HRESULT Init(
        [in] LONG hwndEdit,
        [in] IEnumString *punkACL,
        [in] LPWSTR pwszRegKeyPath,
        [in] LPWSTR pwszQuickComplete);

    HRESULT Enable(
        [in] BOOL fEnable);

};
    
typedef enum {
        [helpstring("No AutoComplete")]
        ACO_NONE               = 0x00,
        [helpstring("enable autosuggest dropdown")]
        ACO_AUTOSUGGEST        = 0x01,
        [helpstring("enable autoappend")]
        ACO_AUTOAPPEND         = 0x02,
        [helpstring("add search entry to completion list")]
        ACO_SEARCH             = 0x04,
        [helpstring("don't match common prefixes (www., http://, etc)")]
        ACO_FILTERPREFIXES     = 0x08,
        [helpstring("use tab to select autosuggest entries")]
        ACO_USETAB             = 0x10,
        [helpstring("up/down arrow key invokes autosuggest dropdown (if enabled)")]
        ACO_UPDOWNKEYDROPSLIST = 0x20,
        [helpstring("enable RTL reading order for dropdown")]
        ACO_RTLREADING         = 0x40,
        [helpstring("Enable AND-ing of results for the search ux.")]
        ACO_WORD_FILTER        = 0x80,
        [helpstring("Disable prefix filtering when displaying autosuggest dropdown.  Always display all suggestions.")]
        ACO_NOPREFIXFILTERING  = 0x100,
} AUTOCOMPLETEOPTIONS;

[
    odl,
    uuid(EAC04BC0-3791-11d2-BB95-0060977B464C)
]
interface IAutoComplete2 : IAutoComplete {

    HRESULT SetOptions( 
        [in] AUTOCOMPLETEOPTIONS dwFlag);
        
    HRESULT GetOptions( 
        [out, retval] AUTOCOMPLETEOPTIONS *pdwFlag);

};

typedef enum ACENUMOPTION {
    ACEO_NONE = 0,
    ACEO_MOSTRECENTFIRST = 0x1,
    ACEO_FIRSTUNUSED = 0x10000
} ACENUMOPTION;

[
    odl,
    uuid(8E74C210-CF9D-4eaf-A403-7356428F0A5A)
]
interface IEnumACString : stdole.IUnknown {

    HRESULT Next(
        [in] LONG celt,
        [in, out] LONG *rgelt,
        [in, out, defaultvalue(0)] LONG *pceltFetched);

    HRESULT Skip(
        [in] LONG celt);

    HRESULT Reset();

    HRESULT Clone(
        [in, out] IEnumString **ppenum);

    HRESULT NextItem( 
        [in] LONG pszUrl,
        [in] LONG cchMax,
        [in, out] LONG *pulSortIndex);
        
    HRESULT SetEnumOptions( 
        [in] ACENUMOPTION dwOptions);
        
    HRESULT GetEnumOptions( 
        [in, out] ACENUMOPTION *pdwOptions);

};

[
    odl,
    uuid(77A130B0-94FD-11D0-A544-00C04FD7d062)
]
interface IACList : stdole.IUnknown {
    HRESULT Expand(
        [in] LONG pszExpand);
};

typedef enum AUTOCOMPLETELISTOPTIONS {
    [helpstring("don't enumerate anything")]
    ACLO_NONE            = 0,
    [helpstring("enumerate current directory")]
    ACLO_CURRENTDIR      = 1,
    [helpstring("enumerate MyComputer")]
    ACLO_MYCOMPUTER      = 2,
    [helpstring("enumerate Desktop Folder")]
    ACLO_DESKTOP         = 4,
    [helpstring("enumerate Favorites Folder")]
    ACLO_FAVORITES       = 8,
    [helpstring("enumerate only the file system")]
    ACLO_FILESYSONLY     = 16,
    [helpstring("enumerate only the file system dirs, UNC shares, and UNC servers.")]
    ACLO_FILESYSDIRS     = 32,
    [helpstring("enumereate on the virual namespace")]
    ACLO_VIRTUALNAMESPACE = 64,
} AUTOCOMPLETELISTOPTIONS;

[
    odl,
    uuid(470141a0-5186-11d2-bbb6-0060977b464c)
]
interface IACList2 : stdole.IUnknown {
    
    HRESULT Expand(
        [in] LONG pszExpand);
        
    HRESULT SetOptions(
        [in] AUTOCOMPLETELISTOPTIONS dwFlag);
        
    HRESULT GetOptions(
        [in, out] AUTOCOMPLETELISTOPTIONS *pdwFlag);
        
};

[
    uuid(00BB2763-6A77-11D0-A535-00C04FD7D062),
    helpstring("AutoComplete class")
]
coclass AutoComplete {
    [default] interface IAutoComplete2;
    interface IAutoComplete;
}

[
    uuid(00BB2764-6A77-11D0-A535-00C04FD7D062),
]
coclass ACLHistory {
    interface IEnumString;
    interface IACList;
}

[
    uuid(03C036F1-A186-11D0-824A-00AA005B4383)
]
coclass ACListISF {
    interface IEnumString;
    interface IACList;
}

[
    uuid(6756A641-DE71-11d0-831B-00AA005B4383)
]
coclass ACLMRU {
    interface IEnumString;
    interface IACList;
}

[
    uuid(00BB2765-6A77-11D0-A535-00C04FD7D062)
]
coclass ACLMulti {
    interface IEnumString;
    interface IACList;
}

