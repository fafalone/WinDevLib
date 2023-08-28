typedef enum STATURL_QUERYFLAGS {
    [helpstring("Return cached items")]
    STATURL_QUERYFLAG_ISCACHED = 0x00010000,
    [helpstring("Don't return URL")]
    STATURL_QUERYFLAG_NOURL    = 0x00020000,
    [helpstring("Don't return URL title")]
    STATURL_QUERYFLAG_NOTITLE  = 0x00040000
} STATURL_QUERYFLAGS;

typedef enum STATURLFLAGS {
    [helpstring("The page is cached.")]
    STATURLFLAG_ISCACHED	= 0x00000001
} STATURLFLAGS;

typedef struct STATURL {
    [helpstring("Struct size")]
    LONG cbSize;
    [helpstring("URL")]
    LONG pwcsUrl;
    [helpstring("Page title")]
    LONG pwcsTitle;
    [helpstring("Last visited date (UTC)")]
    CURRENCY ftLastVisited;
    [helpstring("Last updated date (UTC)")]
    CURRENCY ftLastUpdated;
    [helpstring("Expiration date (UTC)")]
    CURRENCY ftExpires;
    [helpstring("Flags")]
    STATURLFLAGS dwFlags;
} STATURL;

[
    odl,
    uuid(3C374A42-BAE4-11CF-BF7D-00AA006946EE),
    helpstring("Enumerates the cached URLs")
]
interface IEnumSTATURL : stdole.IUnknown {

    [helpstring("Returns the next \"celt\" URLS from the cache")]
    LONG Next (
        [in] LONG celt,
        [in, out] STATURL *rgelt,
        [out, defaultvalue(0)] LONG *pceltFetched);

    [helpstring("Skips the next \"celt\" URLS from the cache")]
    HRESULT Skip (
        [in] LONG celt);

    [helpstring("Resets the enumeration")]
    HRESULT Reset (void);

    [helpstring("Clones this object")]
    HRESULT Clone(
        [out, retval] IEnumSTATURL ** ppenum);

    [helpstring("Sets the enumeration filter")]
    HRESULT SetFilter (
        [in] LPWSTR poszFilter,
        [in] STATURL_QUERYFLAGS dwFlags);
}

[
    odl,
    uuid(3C374A41-BAE4-11CF-BF7D-00AA006946EE),
]
interface IUrlHistoryStg : stdole.IUnknown {

    [helpstring("Adds a new history entry")]
    HRESULT AddUrl (
        [in] LPWSTR pocsUrl,
        [in] LPWSTR pocsTitle,
        [in] STATURLFLAGS dwFlags);

    [helpstring("Deletes an entry by its URL")]
    HRESULT DeleteUrl(
        [in] LPWSTR pocsUrl,
        [in] LONG dwFlags);

    [helpstring("Returns a STATURL for a given URL")]
    HRESULT QueryUrl (
        [in] LPWSTR pocsUrl,
        [in] LONG dwFlags,
        [in, out] STATURL *lpSTATURL);

    [helpstring("Binds to an object")]	
    HRESULT BindToObject (
        [in] LPWSTR pocsUrl,
        [in] UUID *riid,
        [in, out] void *ppvOut);

    [helpstring("Returns an enumerator for URLs")]
    HRESULT EnumUrls (
        [out, retval] IEnumSTATURL **ppEnum);

}

[
    odl,
    uuid(AFA0DC11-C313-11d0-831A-00C04FD5AE38)
]
interface IUrlHistoryStg2 : IUrlHistoryStg {

    [helpstring("Adds an URLS\n IOleCommandTarget *poctNotify\nIShellFolder *punkISFolder")]
    HRESULT AddUrlAndNotify (
        [in] LPWSTR pocsUrl,
        [in] LPWSTR pocsTitle,
        [in] LONG dwFlags,
        [in] BOOL fWriteHistory,
        [in] stdole.IUnknown*poctNotify,
        [in] stdole.IUnknown*punkISFolder );

    [helpstring("Removes all history items")]
    HRESULT ClearHistory();
}

[
	uuid(3C374A40-BAE4-11CF-BF7D-00AA006946EE),
	helpstring("UrlHistory class")
]
coclass UrlHistory {
	[default] interface IUrlHistoryStg2;
}

