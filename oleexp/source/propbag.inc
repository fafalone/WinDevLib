//tbShellLib::100% complete

[
	odl,
	uuid(3127CA40-446E-11CE-8135-00AA004BB851),
	helpstring("IErrorLog interface")
]
interface IErrorLog	: stdole.IUnknown {

	HRESULT	AddError(
		[in] LPWSTR	pszPropName,
		[in] EXCEPINFO *pExcepInfo);

}

[
    odl,
    uuid(55272A00-42CB-11CE-8135-00AA004BB851),
    helpstring("IPropertyBag")
]
interface IPropertyBag : stdole.IUnknown {

    HRESULT Read(
            [in] LONG pszPropName,
            [in, out] VARIANT *pVar,
            [in] IErrorLog *pErrorLog);

    HRESULT Write(
            [in] LONG pszPropName,
            [in] VARIANT *pVar);
}

[
    odl,
    uuid(22F55882-280B-11d0-A8A9-00A0C90C2004),
]
interface IPropertyBag2 : stdole.IUnknown {

    typedef enum PROPBAG2_TYPE {
        PROPBAG2_TYPE_UNDEFINED = 0,
        PROPBAG2_TYPE_DATA      = 1,    // Value is simple data
        PROPBAG2_TYPE_URL       = 2,    // Value is a URL reference
        PROPBAG2_TYPE_OBJECT    = 3,    // Value is an object
        PROPBAG2_TYPE_STREAM    = 4,    // Value is a stream
        PROPBAG2_TYPE_STORAGE   = 5,    // Value is a storage
        PROPBAG2_TYPE_MONIKER   = 6     // Value is a moniker
    } PROPBAG2_TYPE;

    typedef struct PROPBAG2 {
        LONG dwType;             // Property type (from PROPBAG2_TYPE)
        short vt;                // VARIANT property type
        short cfType;            // Clipboard format (aka MIME-type)
        LONG dwHint;             // Property name hint
        LONG pstrName;           // Property name
        UUID clsid;              // CLSID (for PROPBAG2_TYPE_OBJECT)
    } PROPBAG2;

    HRESULT Read(
        [in] LONG cProperties,
        [in] PROPBAG2 *pPropBag,
        [in] IErrorLog *pErrLog,
        [in] LONG pvarValue,
        [in] LONG* phrError);

    HRESULT Write(
        [in] LONG cProperties,
        [in, out] PROPBAG2 *pPropBag,
        [in] VARIANT* pvarValue);

    HRESULT CountProperties(
        [out, retval] LONG *pcProperties);

    HRESULT GetPropertyInfo(
        [in] LONG iProperty,
        [in] LONG cProperties,
        [in, out] PROPBAG2 *pPropBag,
        [in, out] LONG *pcProperties);

    HRESULT LoadObject(
        [in] LONG pstrName,
        [in] LONG dwHint,
        [in] stdole.IUnknown *pUnkObject,
        [in] IErrorLog *pErrLog);
}

[
	odl,
	uuid(9BFBBC02-EFF1-101A-84ED-00AA00341D07)
]
interface IPropertyNotifySink : stdole.IUnknown {

	HRESULT OnChanged(
		[in] long dispID);
    
	HRESULT OnRequestEdit(
		[in] long dispID);
    
}
