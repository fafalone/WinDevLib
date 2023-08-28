//tbShellLib - 100% implemented	

typedef struct CALPOLESTR {
		LONG cElems;
		LONG pElems;
	} CALPOLESTR;

	typedef struct CADWORD {
		LONG cElems;
		LONG pElems;
	} CADWORD;

	typedef LONG DISPID;

	[
		odl,
		uuid(376BD3AA-3845-101B-84ED-08002B2EC713),
	 	helpstring("IPerPropertyBrowsing interface")
	]
	interface IPerPropertyBrowsing : stdole.IUnknown{

		[helpstring("Get the display string for the property")]
		LONG GetDisplayString(
			[in] DISPID DispID,
			[in, out] BSTR * DisplayName);

		[helpstring("MapPropertyToPage")]
		LONG MapPropertyToPage(
			[in] DISPID DispID,
			[in, out] UUID *pClsID);

		[helpstring("Get the predefined strings for the property")]
		LONG GetPredefinedStrings(
			[in] DISPID DispID,
	                [in, out] CALPOLESTR * pCaStringsOut,
        	        [in, out] CADWORD * pCaCookiesOut);
		
		[helpstring("Get the predefined value for the property")]
		LONG GetPredefinedValue(
			[in] DISPID DispID,
			[in] LONG dwCookie,
			[in, out] VARIANT * pVarOut);
	}

