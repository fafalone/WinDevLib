//tbShellLib - 100% implemented

[
    odl,
    uuid(00020405-0000-0000-C000-000000000046)
]
interface ICreateTypeInfo: stdole.IUnknown {

    HRESULT SetGuid(
        [in] UUID *guid);

    HRESULT SetTypeFlags(
        [in] LONG uTypeFlags);

    HRESULT SetDocString(
        [in] LPWSTR pStrDoc);

    HRESULT SetHelpContext(
        [in] LONG dwHelpContext);

    HRESULT SetVersion(
        [in] short wMajorVerNum,
        [in] short wMinorVerNum);

    HRESULT AddRefTypeInfo(
        [in] ITypeInfo *pTInfo,
        [in] LONG *phRefType);

    HRESULT AddFuncDesc(
        [in] LONG index,
        [in] FUNCDESC * pFuncDesc);

    HRESULT AddImplType(
        [in] LONG index,
        [in] LONG hRefType);

    HRESULT SetImplTypeFlags(
        [in] LONG index,
        [in] INT implTypeFlags);

    HRESULT SetAlignment(
        [in] short cbAlignment);

    HRESULT SetSchema(
        [in] LPWSTR pStrSchema);

    HRESULT AddVarDesc(
        [in] LONG index,
        [in] VARDESC * pVarDesc);

    HRESULT SetFuncAndParamNames(
        [in] LONG index,
        [in] LPWSTR * rgszNames,
        [in] LONG cNames);

    HRESULT SetVarName(
        [in] LONG index,
        [in] LPWSTR szName);

    HRESULT SetTypeDescAlias(
        [in] TYPEDESC * pTDescAlias);

    HRESULT DefineFuncAsDllEntry(
        [in] LONG index,
        [in] LPWSTR szDllName,
        [in] LPWSTR szProcName);

    HRESULT SetFuncDocString(
        [in] LONG index,
        [in] LPWSTR szDocString);

    HRESULT SetVarDocString(
        [in] LONG index,
        [in] LPWSTR szDocString);

    HRESULT SetFuncHelpContext(
        [in] LONG index,
        [in] LONG dwHelpContext);

    HRESULT SetVarHelpContext(
        [in] LONG index,
        [in] LONG dwHelpContext);

    HRESULT SetMops(
        [in] LONG index,
        [in] BSTR bstrMops);

    HRESULT SetTypeIdldesc(
        [in] IDLDESC * pIdlDesc);

    HRESULT LayOut(void);

}

[
    odl,
    uuid(0002040E-0000-0000-C000-000000000046)
]
interface ICreateTypeInfo2: ICreateTypeInfo {

    HRESULT DeleteFuncDesc(
        [in] LONG index);

    HRESULT DeleteFuncDescByMemId(
        [in] LONG memid,
        [in] INVOKEKIND invKind);

    HRESULT DeleteVarDesc(
        [in] LONG index);

    HRESULT DeleteVarDescByMemId(
        [in] LONG memid);

    HRESULT DeleteImplType(
        [in] LONG index);

    HRESULT SetCustData(
        [in] UUID *guid,
        [in] VARIANT * pVarVal);

    HRESULT SetFuncCustData(
        [in] LONG index,
        [in] UUID *guid,
        [in] VARIANT * pVarVal);

    HRESULT SetParamCustData(
        [in] LONG indexFunc,
        [in] LONG indexParam,
        [in] UUID *guid,
        [in] VARIANT * pVarVal);

    HRESULT SetVarCustData(
        [in] LONG index,
        [in] UUID *guid,
        [in] VARIANT * pVarVal);

    HRESULT SetImplTypeCustData(
        [in] LONG index,
        [in] UUID *guid,
        [in] VARIANT * pVarVal);

    HRESULT SetHelpStringContext(
        [in] LONG dwHelpStringContext);

    HRESULT SetFuncHelpStringContext(
        [in] LONG index,
        [in] LONG dwHelpStringContext);

    HRESULT SetVarHelpStringContext(
        [in] LONG index,
        [in] LONG dwHelpStringContext);

    HRESULT Invalidate(void);

    HRESULT SetName(
        [in]  LPWSTR szName);

}

[
    odl,
    uuid(00020406-0000-0000-C000-000000000046)
]
interface ICreateTypeLib : stdole.IUnknown {

    HRESULT CreateTypeInfo(
        [in] LPWSTR szName,
        [in] TYPEKIND tkind,
        [out] ICreateTypeInfo **ppCTInfo);

    HRESULT SetName(
        [in] LPWSTR szName);

    HRESULT SetVersion(
        [in] short wMajorVerNum,
        [in] short wMinorVerNum);

    HRESULT SetGuid(
        [in] UUID *guid);

    HRESULT SetDocString(
        [in] LPWSTR szDoc);

    HRESULT SetHelpFileName(
        [in] LPWSTR szHelpFileName);

    HRESULT SetHelpContext(
        [in] LONG dwHelpContext);

    HRESULT SetLcid(
        [in] LONG lcid);

    HRESULT SetLibFlags(
        [in] LIBFLAGS uLibFlags);

    HRESULT SaveAllChanges(void);
}

[
    odl,
    uuid(0002040F-0000-0000-C000-000000000046)
]
interface ICreateTypeLib2 : ICreateTypeLib
{

    HRESULT DeleteTypeInfo(
        [in] LPWSTR szName);

    HRESULT SetCustData(
        [in] UUID *guid,
        [in] VARIANT * pVarVal);

    HRESULT SetHelpStringContext(
        [in] LONG dwHelpStringContext);

    HRESULT SetHelpStringDll(
        [in] LPWSTR szFileName);

}
