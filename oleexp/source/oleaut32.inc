//tbShellLib - 100% Implemented

[
    dllname("OLEAUT32.DLL")
]
module oleaut32 {


    [entry("DispCallFunc")]int DispCallFunc([in] void* pvInstance,
                                            [in] int oVft,
                                            [in] short cc,
                                            [in] short vtReturn,
                                            [in] int cActuals,
                                            [in] void* prgvt,
                                            [in] void* prgpvarg,
                                            [in] void* pvargResult);
    [entry("SystemTimeToVariantTime")]
    LONG SystemTimeToVariantTime(
        [in] SYSTEMTIME *lpSystemTime,
        [out] DOUBLE *pvtime);

    [entry("VariantTimeToSystemTime")]
    LONG VariantTimeToSystemTime(
        [in] DOUBLE vtime,
        [in] SYSTEMTIME *lpSystemTime);

    [entry("DosDateTimeToVariantTime")]
    LONG DosDateTimeToVariantTime(
        [in] short wDosDate,
        [in] short wDosTime,
        [out] DOUBLE *pvtime);

    [entry("VariantTimeToDosDateTime")]
    LONG VariantTimeToDosDateTime(
        [in] DOUBLE vtime,
        [out] short *pwDosDate,
        [out] short *pwDosTime);

    [entry("LHashValOfName")]
    LONG LHashValOfName(
        [in] LONG lcid,
        [in, string] LPWSTR szName);

    [entry("LHashValOfNameSys")]
    LONG LHashValOfNameSys(
        [in] SYSKIND syskind,
        [in] LONG  lcid,
        [in, string] LPWSTR szName);

    [entry("RegisterTypeLib")]
    HRESULT RegisterTypeLib(
        [in] ITypeLib * ptlib,
        [in] LPWSTR szFullPath,
        [in] LPWSTR szHelpDir);

    [entry("UnRegisterTypeLib")]
    HRESULT UnRegisterTypeLib(
        [in] UUID *libID,
        [in] short wVerMajor,
        [in] short wVerMinor,
        [in] LONG lcid,
        [in] SYSKIND syskind);

    [entry("CreateTypeLib")]
    HRESULT CreateTypeLib(
        [in] SYSKIND syskind,
        [in] LPWSTR szFile,
        [out, retval] ICreateTypeLib **ppctlib);

    [entry("CreateTypeLib2")]
    HRESULT CreateTypeLib2(
        [in] SYSKIND syskind,
        [in] LPWSTR szFile,
        [out, retval] ICreateTypeLib2 **ppctlib);

    [entry("LoadRegTypeLib")]
    HRESULT LoadRegTypeLib(
        [in] UUID *rguid,
        [in] short wVerMajor,
        [in] short wVerMinor,
        [in] LONG lcid,
        [out, retval] ITypeLib ** pptlib);

    [entry("LoadTypeLibEx")]
    HRESULT LoadTypeLibEx(
        [in, string] LPWSTR szFile,
        [in] REGKIND regkind,
        [out, retval] ITypeLib **pptlib);

    [entry("LoadTypeLib")]
    HRESULT LoadTypeLib(
        [in, string] LPWSTR szFile,
        [out, retval] ITypeLib **pptlib);

    [entry("QueryPathOfRegTypeLib")]
    HRESULT QueryPathOfRegTypeLib(
        [in] UUID *GUID,
        [in] short wVerMajor,
        [in] short wVerMinor,
        [in] LONG LCID,
        [out, retval] BSTR *lpbstrPathName);

    [entry("SysAllocString")]
    BSTR SysAllocString(
        [in] LONG pStr);

    [entry("SysAllocString")]
    LONG SysAllocStringPtrPtr(
        [in] LONG pStr);

    [entry("SysAllocString")]
    LONG SysAllocStringPtr(
        [in] LPWSTR pStr);

    [entry("SysAllocString"),helpstring("This is an alias for SysAllocString")]
    BSTR RTrimNull(
        [in] LPWSTR pStr);

    [entry("SysReAllocString")]
    LONG SysReAllocString(
        [in] BSTR *in,
        [in] LONG pStr);

    [entry("SysReAllocString"),helpstring("Alias for SysReAllocString supporting Unicode.")]
    LONG SysReAllocStringW(
        [in] LONG pBSTR,
        [in] LONG pStr);

    [entry("SysAllocStringLen")]
    BSTR SysAllocStringLen(
        [in] LONG pStr,
        [in] LONG cb);

    [entry("SysReAllocStringLen")]
    LONG SysReAllocStringLen(
        [in] BSTR *bstr,
        [in] LONG pStr,
        [in] lONG cb);

    [entry("SysFreeString")]
    void SysFreeString(
        [in] BSTR *bstr);
 
    [entry("SysStringLen")]
    LONG SysStringLen(
        [in] BSTR str);

    [entry("SysStringByteLen")]
    LONG SysStringByteLen(
        [in] BSTR bstr);

    [entry("SysAllocStringByteLen")]
    BSTR SysAllocStringByteLen(
        [in] void *psz,
        [in] LONG len);

    [entry("VarR8Round")]
    HRESULT VarR8Round(
        [in] double dblIn,
        [in] int cDecimals,
        [out, retval] double *pdblResult);

    [entry("GetAltMonthNames")]
    HRESULT GetAltMonthNames(
        [in] LONG lcid,
        [out] LONG *prgp);

    [entry("VarFormat")]
    HRESULT VarFormat(
        [in] VARIANT *pvarIn,
        [in] LPWSTR pstrFormat,
        [in, defaultvalue(0)] LONG iFirstDay,
        [in, defaultvalue(0)] LONG iFirstWeek,
        [in, defaultvalue(0)] VARFORMAT_FLAGS dwFlags,
        [out, retval] BSTR *pbstrOut);

    [entry("VarFormatDateTime")]
    HRESULT VarFormatDateTime(
        [in] VARIANT *pvarIn,
        [in, defaultvalue(0)] NAMEDFORMATS iNamedFormat,
        [in, defaultvalue(0)] VARFORMAT_FLAGS dwFlags,
        [out, retval] BSTR *pbstrOut);

    [entry("VarFormatNumber")]
    HRESULT VarFormatNumber(
        [in] VARIANT *pvarIn,
        [in, defaultvalue(-1)] LONG iNumDig,
        [in, defaultvalue(-2)] LONG iIncLead,
        [in, defaultvalue(-2)] LONG iUseParens,
        [in, defaultvalue(-2)] LONG iGroup,
        [in, defaultvalue(0)] VARFORMAT_FLAGS dwFlags,
        [out, retval] BSTR *pbstrOut);

    [entry("VarFormatPercent")]
    HRESULT VarFormatPercent(
        [in] VARIANT *pvarIn,
        [in, defaultvalue(-1)] LONG iNumDig,
        [in, defaultvalue(-2)] LONG iIncLead,
        [in, defaultvalue(-2)] LONG iUseParens,
        [in, defaultvalue(-2)] LONG iGroup,
        [in, defaultvalue(0)] VARFORMAT_FLAGS dwFlags,
        [out, retval] BSTR *pbstrOut);

    [entry("VarFormatCurrency")]
    HRESULT VarFormatCurrency(
        [in] VARIANT *pvarIn,
        [in, defaultvalue(-1)] LONG iNumDig,
        [in, defaultvalue(-2)] LONG iIncLead,
        [in, defaultvalue(-2)] LONG iUseParens,
        [in, defaultvalue(-2)] LONG iGroup,
        [in, defaultvalue(0)] VARFORMAT_FLAGS dwFlags,
        [out, retval] BSTR *pbstrOut);

    [entry("VarWeekdayName")]
    HRESULT VarWeekdayName(
        [in] LONG iWeekday,
        [in, defaultvalue(0)] BOOL fAbbrev,
        [in, defaultvalue(0)] LONG iFirstDay,
        [in, defaultvalue(0)] VARFORMAT_FLAGS dwFlags,
        [out, retval] BSTR *pbstrOut);

    [entry("VarMonthName")]
    HRESULT VarMonthName(
        [in] LONG iMonth, 
        [in, defaultvalue(0)] BOOL fAbbrev,
        [in, defaultvalue(0)] VARFORMAT_FLAGS dwFlags,
        [out, retval] BSTR *pbstrOut);

    [entry("VarFormatFromTokens")]
    HRESULT VarFormatFromTokens(
        [in] VARIANT *pvarIn,
        [in] LPWSTR pstrFormat,
        [in] void *pbTokCur,
        [in] VARFORMAT_FLAGS dwFlags,
        [out] BSTR *pbstrOut, 
        [in] LONG lcid);

    [entry("VarTokenizeFormatString")]
    HRESULT VarTokenizeFormatString(
        [in] LPWSTR pstrFormat,
        [in] void *rgbTok,
        [in] LONG cbTok,
        [in, defaultvalue(0)] LONG iFirstDay,
        [in, defaultvalue(0)] LONG iFirstWeek,
        [in, defaultvalue(0)] LONG lcid,
        [in, defaultvalue(0)] LONG *pcbActual);

    [entry("DispGetParam")]
    HRESULT DispGetParam(
        [in] DISPPARAMS *pdispparams,
        [in] LONG position,
        [in] short vtTarg,
        [out] VARIANT *pvarResult,
        [out] LONG *puArgErr);

    [entry("DispGetIDsOfNames")]
    HRESULT DispGetIDsOfNames(
        [in] ITypeInfo * ptinfo,
        [in] LPWSTR *rgszNames,
        [in] LONG cNames,
        [out] LONG *rgdispid);

    [entry("DispInvoke")]
    HRESULT DispInvoke(
        [in] void *this,
        [in] ITypeInfo *ptinfo,
        [in] LONG dispidMember,
        [in] short wFlags,
        [in, out] DISPPARAMS *pparams,
        [out] VARIANT *pvarResult,
        [in, out] EXCEPINFO *pexcepinfo,
        [out] LONG *puArgErr);

    [entry("CreateStdDispatch")]
    HRESULT CreateStdDispatch(
        [in] stdole.IUnknown*punkOuter,
        [in] void *pvThis,
        [in] ITypeInfo *ptinfo,
        [out, retval] stdole.IUnknown **ppunkStdDisp);

    [entry("RegisterActiveObject")]
    HRESULT RegisterActiveObject(
        [in] stdole.IUnknown *punk,
        [in] UUID *rclsid,
        [in, defaultvalue(0)] ACTIVEOBJECTFLAGS dwFlags,
        [out, retval] LONG *pdwRegister);

    [entry("RevokeActiveObject")]
    HRESULT RevokeActiveObject(
        [in] LONG dwRegister,
        [in, defaultvalue(0)] LONG pvReserved);

    [entry("GetActiveObject")]
    HRESULT GetActiveObject(
        [in] UUID *rclsid,
        [in, defaultvalue(0)] LONG pvReserved,
        [out, retval] stdole.IUnknown **ppunk);

    [entry("SetErrorInfo")]
    HRESULT SetErrorInfo(
        [in] LONG dwReserved,
        [in] IErrorInfo *perrinfo);

    [entry("GetErrorInfo")]
    HRESULT GetErrorInfo(
        [in, defaultvalue(0)] LONG dwReserved,
        [out, retval] IErrorInfo **pperrinfo);

    [entry("CreateErrorInfo")]
    HRESULT CreateErrorInfo(
        [out, retval] ICreateErrorInfo **pperrinfo);

    [entry("GetRecordInfoFromTypeInfo")]
    HRESULT GetRecordInfoFromTypeInfo(
        [in] ITypeInfo *pTypeInfo,
        [out, retval] IRecordInfo **ppRecInfo);

    [entry("GetRecordInfoFromGuids")]
    HRESULT GetRecordInfoFromGuids(
        [in] UUID *rGuidTypeLib,
        [in] LONG uVerMajor,
        [in] LONG uVerMinor,
        [in] LONG lcid,
        [in] UUID *rGuidTypeInfo,
        [out, retval] IRecordInfo **ppRecInfo);

    [entry("VarParseNumFromStr")]
    HRESULT VarParseNumFromStr(
        [in] LPWSTR strIn,
        [in] LONG lcid,
        [in] NUMPRS dwFlags,
        [in] NUMPARSE *pnumprs,
        [in] BYTE *rgbDig);

    [entry("VarNumFromParseNum")]
    HRESULT VarNumFromParseNum(
        [in] NUMPARSE *pnumprs,
        [in] BYTE *rgbDig,
        [in] LONG dwVtBits,
        [out, retval] VARIANT *pvar);

    [entry("VarBstrCat")]
    HRESULT VarBstrCat( 
        [in] BSTR bstrLeft,         
        [in] BSTR bstrRight,  
        [out, retval] BSTR *pbstrResult);

    [entry("VarBstrCmp")]
    VARCMP VarBstrCmp( 
        [in] BSTR bstrLeft,         
        [in] BSTR bstrRight,  
        [in] LONG lcid,
        [in] NORM dwFlags);

    [entry("VariantChangeType")]
    HRESULT VariantChangeType( 
        [in] VARIANT *pvargDest,  
        [in] VARIANT *pvarSrc,  
        [in] short wFlags,  
        [in] short vt);

    [entry("VariantChangeTypeEx")]
    HRESULT VariantChangeTypeEx( 
        [in] VARIANT *pvargDest,  
        [in] VARIANT *pvarSrc,  
        [in] LONG lcid,             
        [in] short wFlags,  
        [in] short vt);
        
    [entry("VariantClear")]
    HRESULT VariantClear( 
        [in] VARIANT *pvarg);

    [entry("VariantCopy")]
    HRESULT VariantCopy( 
        [in, out] VARIANT *pvargDest,
        [in] VARIANT *pvargSrc);

    [entry("VariantCopy")]
    HRESULT VariantCopyPtrVar( 
        [in] LONG pvargDest,
        [in] VARIANT *pvargSrc);

    [entry("VariantCopy")]
    HRESULT VariantCopyVarPtr( 
        [in, out] VARIANT *pvargDest,
        [in] LONG pvargSrc);

    [entry("VariantCopy")]
    HRESULT VariantCopyPtr( 
        [in] LONG pvargDest,
        [in] LONG pvargSrc);

    [entry("VariantCopyInd")]
    HRESULT VariantCopyInd( 
        [in, out] VARIANT *pvargDest,
        [in] VARIANT *pvargSrc);

    [entry("VariantCopyInd")]
    HRESULT VariantCopyIndPtrVar( 
        [in] LONG pvargDest,
        [in] VARIANT *pvargSrc);

    [entry("VariantCopyInd")]
    HRESULT VariantCopyIndVarPtr( 
        [in, out] VARIANT *pvargDest,
        [in] LONG pvargSrc);

    [entry("VariantCopyInd")]
    HRESULT VariantCopyIndPtr( 
        [in] LONG pvargDest,
        [in] LONG pvargSrc);

    [entry("VariantInit")]
    void VariantInit(
        [in] VARIANT *pvarg);
        
    [entry("OleLoadPicture")]
    HRESULT OleLoadPicture(
        [in] IStream *lpstream,
        [in] LONG lSize,
        [in] BOOL fRunmode,
        [in] UUID *riid,
        [out, retval] IPicture **lplpvObj);

    [entry("OleLoadPictureEx")]
    HRESULT OleLoadPictureEx(
        [in] IStream *lpstream,
        [in] LONG lSize,
        [in] BOOL fRunmode,
        [in] UUID *riid,
        [in] LONG xSizeDesired,
        [in] LONG ySizeDesired,
        [in] stdole.LoadPictureConstants dwFlags,
        [out, retval] IPicture **lplpvObj);

    [entry("OleLoadPicturePath")]
    HRESULT OleLoadPicturePath(
        [in] LPWSTR szURLorPath,
	    [in] IUnknown *punkCaller,
        [in] LONG dwReserved,
	    [in] OLE_COLOR clrReserved,
        [in] UUID *riid,
        [out, retval] IPicture **ppvRet);

    [entry("OleLoadPictureFile")]
    HRESULT OleLoadPictureFile(
        [in] VARIANT varFileName,
        [out, retval] IPictureDisp **lplpdispPicture);

    [entry("OleLoadPictureFileEx")]
    HRESULT OleLoadPictureFileEx(
        [in] VARIANT varFileName,
        [in] LONG xSizeDesired,
        [in] LONG ySizeDesired,
        [in] stdole.LoadPictureConstants dwFlags,
        [out, retval] IPictureDisp **lplpdispPicture);

    [entry("OleSavePictureFile")]
    HRESULT OleSavePictureFile(
        [in] IPictureDisp *lpdispPicture,
        [in] BSTR bstrFileName);

    [entry("OleTranslateColor")]
    HRESULT OleTranslateColor(
        [in] OLE_COLOR clr,
        [in, defaultvalue(0)] LONG hpal,
        [out, retval] LONG *lpcolorref);

}