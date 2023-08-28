    [
        odl,
        uuid(00020401-0000-0000-C000-000000000146),
    ]
    interface ITypeInfo : stdole.IUnknown {

        HRESULT GetTypeAttr(
                [out, retval] LONG * ppTypeAttr);

        HRESULT GetTypeComp(
                [out, retval] ITypeComp ** ppTComp);

        HRESULT GetFuncDesc(
                [in] LONG index,
                [out, retval] LONG * ppFuncDesc);

        HRESULT GetVarDesc(
                [in] LONG index,
                [out, retval] LONG * ppVarDesc);

        HRESULT GetNames(
                [in] LONG memid,
                [in, out] BSTR * rgBstrNames,
                [in] LONG cMaxNames,
                [out, retval] LONG * pcNames);

        HRESULT GetRefTypeOfImplType(
                [in] LONG index,
                [out, retval] LONG *pRefType);

        HRESULT GetImplTypeFlags(
                [in] LONG index,
                [out, retval] LONG *pImplTypeFlags);

        HRESULT GetIDsOfNames(
                [in] LPWSTR *rgszNames,
                [in] long cNames,
                [in, out] long *pMemId);

        HRESULT Invoke(
                [in] stdole.IDispatch *pvInstance,
                [in] LONG memid,
                [in] short wFlags,
                [in, out] DISPPARAMS * pDispParams,
                [in] LONG pVarResult,
                [in, out] EXCEPINFO * pExcepInfo,
                [in, out] LONG * puArgErr);

        HRESULT GetDocumentation(
                [in] LONG memid,
                [in, out] BSTR * pBstrName,
                [in, out] BSTR * pBstrDocString,
                [in, out] LONG * pdwHelpContext,
                [in, out] BSTR * pBstrHelpFile);

        HRESULT GetDllEntry(
                [in] LONG memid,
                [in] INVOKEKIND invKind,
                [in, out] BSTR * pBstrDllName,
                [in, out] BSTR * pBstrName,
                [in, out] short * pwOrdinal);

        HRESULT GetRefTypeInfo(
                [in] LONG hRefType,
                [out, retval] ITypeInfo **ppTInfo);

        HRESULT AddressOfMember(
                [in] LONG memid,
                [in] INVOKEKIND invKind,
                [out, retval] LONG *ppv);

        HRESULT CreateInstance(
                [in] void *pUnkOuter,
                [in] UUID *riid,
                [out, retval] IDispatch **ppvObj);

        HRESULT GetMops(
                [in] LONG memid,
                [in, out] BSTR * pBstrMops);

        HRESULT GetContainingTypeLib(
                [out] ITypeLib **ppTLib,
                [out, retval] LONG * pIndex);

        void ReleaseTypeAttr(
                [in] LONG pTypeAttr);

        void ReleaseFuncDesc(
                [in] LONG pFuncDesc);

        void ReleaseVarDesc(
                [in] LONG pVarDesc);

    }

    [
        uuid(00020412-0000-0000-C000-000000000046),
        helpstring("ITypeInfo2 interface"),
        odl
    ]
    interface ITypeInfo2 : ITypeInfo {

        HRESULT GetTypeKind(
                [out, retval] TYPEKIND * pTypeKind);

        HRESULT GetTypeFlags(
                [out, retval] TYPEFLAGS *pTypeFlags);

        HRESULT GetFuncIndexOfMemId(
                [in] LONG memid,
                [in] INVOKEKIND invKind,
                [out, retval] LONG *pFuncIndex);

        HRESULT GetVarIndexOfMemId(
                [in] LONG memid,
                [out, retval] LONG * pVarIndex);

        HRESULT GetCustData(
                [in] UUID *IID,
                [out, retval] VARIANT * pVarVal);

        HRESULT GetFuncCustData(
                [in] LONG index,
                [in] UUID *riid,
                [out, retval] VARIANT * pVarVal);

        HRESULT GetParamCustData(
                [in] LONG indexFunc,
                [in] LONG indexParam,
                [in] UUID *riid,
                [out, retval] VARIANT * pVarVal);

        HRESULT GetVarCustData(
                [in] LONG index,
                [in] UUID * riid,
                [out, retval] VARIANT * pVarVal);

       HRESULT GetImplTypeCustData(
                [in] LONG index,
                [in] UUID *riid,
                [out, retval] VARIANT * pVarVal);

       HRESULT GetDocumentation2(
                [in] LONG memid,
                [in] LONG lcid,
                [out] BSTR *pbstrHelpString,
                [out] LONG *pdwHelpStringContext,
                [out] BSTR *pbstrHelpStringDll);

       HRESULT GetAllCustData(
                [in, out] CUSTDATA * pCustData);

       HRESULT GetAllFuncCustData(
                [in] LONG index,
                [in, out] CUSTDATA * pCustData);

       HRESULT GetAllParamCustData(
                [in] LONG indexFunc,
                [in] LONG indexParam,
                [in, out] CUSTDATA * pCustData);

       HRESULT GetAllVarCustData(
                [in] LONG index,
                [in, out] CUSTDATA * pCustData);

       HRESULT GetAllImplTypeCustData(
                [in] LONG index,
                [in, out] CUSTDATA * pCustData);
    }

    [
        uuid(00020402-0000-0000-C000-000000000046),
        helpstring("ITypeLib interface"),
        odl
    ]
    interface ITypeLib : stdole.IUnknown {

        LONG GetTypeInfoCount(
                void);

        HRESULT GetTypeInfo(
            [in]  LONG index,
            [out, retval] ITypeInfo ** ppTInfo);

        HRESULT GetTypeInfoType(
            [in]  LONG index,
            [out, retval] TYPEKIND * pTKind);

        HRESULT GetTypeInfoOfIID(
            [in] UUID *IID,
            [out, retval] ITypeInfo ** ppTinfo);

        HRESULT GetLibAttr(
            [out, retval] LONG *ppTLibAttr);

        HRESULT GetTypeComp(
            [out, retval] ITypeComp ** ppTComp);

        HRESULT GetDocumentation(
            [in] LONG index,
            [out] BSTR * pBstrName,
            [out] BSTR * pBstrDocString,
            [out] LONG * pdwHelpContext,
            [out] BSTR * pBstrHelpFile);

        HRESULT IsName(
            [in, out] LPWSTR szNameBuf,
            [in] LONG lHashVal,
            [out, retval] BOOL * pfName);

       HRESULT FindName(
            [in, out] LPWSTR szNameBuf,
            [in] LONG lHashVal,
            [in, out] ITypeInfo **ppTInfo,
            [in, out] LONG *rgMemId,
            [in, out] short *pcFound);

       void ReleaseTLibAttr(
            [in]  LONG pTLibAttr);

    }

    [
        uuid(00020411-0000-0000-C000-000000000046),
        odl
    ]
    interface ITypeLib2 : ITypeLib {

        HRESULT GetCustData(
            [in] UUID *IID,
            [out, retval] VARIANT * pVarVal);

        HRESULT GetLibStatistics(
            [out] LONG * pcUniqueNames,
            [out] LONG * pcchUniqueNames);

        HRESULT GetDocumentation2(
            [in]  INT index,
            [in]  LONG lcid,
            [out] BSTR *pbstrHelpString,
            [out] LONG *pdwHelpStringContext,
            [out] BSTR *pbstrHelpStringDll);

        HRESULT GetAllCustData(
            [out] CUSTDATA * pCustData);
    }

    [
        odl,
        uuid(00020403-0000-0000-C000-000000000046)
    ]
    interface ITypeComp : stdole.IUnknown {

        typedef enum DESCKIND {
            DESCKIND_NONE = 0,
            DESCKIND_FUNCDESC,
            DESCKIND_VARDESC,
            DESCKIND_TYPECOMP,
            DESCKIND_IMPLICITAPPOBJ,
            DESCKIND_MAX
        } DESCKIND;

        typedef [public] LONG BINDPTR;

        HRESULT Bind(
            [in] LPWSTR szName,
            [in] LONG lHashVal,
            [in] short wFlags,
            [out] ITypeInfo ** ppTInfo,
            [out] DESCKIND * pDescKind,
            [out] BINDPTR * pBindPtr);

        HRESULT BindType(
            [in] LPWSTR szName,
            [in] LONG lHashVal,
            [out] ITypeInfo ** ppTInfo,
            [out] ITypeComp ** ppTComp);

    }

    [
        odl,
        uuid(B196B283-BAB4-101A-B69C-00AA00341D07),
    ]
    interface IProvideClassInfo : stdole.IUnknown {

        HRESULT GetClassInfo(
                [out, retval] ITypeInfo ** ppTI);
    }

    [
        odl,
        uuid(B196B284-BAB4-101A-B69C-00AA00341D07),
    ]
    interface IConnectionPointContainer : stdole.IUnknown {

        HRESULT EnumConnectionPoints(
            [out, retval] IEnumConnectionPoints ** ppEnum);

        HRESULT FindConnectionPoint(
            [in] UUID *riid,
            [out, retval]   IConnectionPoint ** ppCP);
    }

    [
        odl,
        uuid(B196B286-BAB4-101A-B69C-00AA00341D07),
    ]
    interface IConnectionPoint : stdole.IUnknown {

        HRESULT GetConnectionInterface(
            [in, out] UUID *piid);

        HRESULT GetConnectionPointContainer(
            [out, retval] IConnectionPointContainer ** ppCPC);

        HRESULT Advise(
            [in] stdole.IUnknown *pObjetPtr,
            [out, retval] LONG *pdwCookie);

        HRESULT Unadvise(
            [in] LONG dwCookie);

        HRESULT EnumConnections(
            [out, retval] IEnumConnections **ppEnum);
    }

    [
      odl,
      uuid(00020400-0000-0000-C000-000000000046),
    ]
    interface IDispatch : IUnknown {

        HRESULT GetTypeInfoCount(
            [out, retval] int* pctinfo);

        HRESULT GetTypeInfo(
            [in, defaultvalue(0)] int itinfo,
            [in, defaultvalue(0)] long lcid,
            [out, retval] ITypeInfo **pptinfo);

        LONG GetIDsOfNames(
            [in] UUID* riid,
            [in] LPWSTR *rgszNames,
            [in] int cNames,
            [in] long lcid,
            [out] long *rgdispid);

        LONG Invoke(
            [in] long dispidMember,
            [in] UUID* riid,
            [in] long lcid,
            [in] short wFlags,
            [in] DISPPARAMS *pdispparams,
            [in] long pvarResult,
            [out] EXCEPINFO *pexcepinfo,
            [out] int *puArgErr);
    };

