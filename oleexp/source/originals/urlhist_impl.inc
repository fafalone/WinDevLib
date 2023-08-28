[
	odl,
	uuid(3C374A41-BAE4-11CF-BF7D-00AA006946EE),
]
interface IUrlHistoryStg : stdole.IUnknown {

	HRESULT AddUrl (
		[in] LONG pocsUrl,
		[in] LONG pocsTitle,
		[in] olelib.STATURLFLAGS dwFlags);

	HRESULT DeleteUrl(
		[in] LONG pocsUrl,
		[in] LONG dwFlags);

	HRESULT QueryUrl (
		[in] LONG pocsUrl,
		[in] LONG dwFlags,
		[in, out] olelib.STATURL *lpSTATURL);

	HRESULT BindToObject (
		[in] LONG pocsUrl,
		[in, out] UUID *riid,
		[in] long ppvOut);

	HRESULT EnumUrls (
		[out, retval] IEnumSTATURL **ppEnum);

}
