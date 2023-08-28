    [
        odl,
        uuid(0000010c-0000-0000-C000-000000000046)
    ]
    interface IPersist : stdole.IUnknown {

        HRESULT GetClassID(
            [in, out] UUID *pClassID);
    }

    [
        odl,
        uuid(00000109-0000-0000-C000-000000000046),
    ]
    interface IPersistStream : IPersist {

        LONG IsDirty(void);

        HRESULT Load(
            [in] IStream *pStm);

        HRESULT Save(
            [in] IStream *pStm,
            [in] BOOL fClearDirty);

        HRESULT GetSizeMax(
            [out, retval] CURRENCY *pcbSize);

    }

    [
        odl,
        uuid(7FD52380-4E07-101B-AE2D-08002B2EC713),
        helpstring("IPersistStreamInit interface")
    ]
    interface IPersistStreamInit : IPersist {

        HRESULT IsDirty();

        HRESULT Load(
            [in] IStream *pStm);

        HRESULT Save(
            [in] IStream *pStm,
            [in] BOOL fClearDirty);

        HRESULT GetSizeMax(
            [out, retval] CURRENCY *pCbSize);

        HRESULT InitNew();

    }

    [
        odl,
        uuid(0000010b-0000-0000-C000-000000000046),
    ]
    interface IPersistFile : IPersist {

        LONG IsDirty();

        HRESULT Load(
            [in] LPWSTR pszFileName,
            [in] STGM dwMode);

        HRESULT Save(
        [in] LPWSTR pszFileName,
        [in] BOOL fRemember);

        HRESULT SaveCompleted(
            [in] LPWSTR pszFileName);

        HRESULT GetCurFile(
            [out, retval] LONG *ppszFileName);
    }

    [
        odl,
        uuid(0000010a-0000-0000-C000-000000000046),
    ]
    interface IPersistStorage : IPersist {

        LONG IsDirty();

        HRESULT InitNew(
            [in] IStorage *pStg);

        HRESULT Load(
            [in] IStorage *pStg);

        HRESULT Save(
            [in] IStorage *pStgSave,
            [in] BOOL fSameAsLoad);

        HRESULT SaveCompleted(
            [in] IStorage *pStgNew);

        HRESULT HandsOffStorage();

    }

    [
        odl,
        uuid(37D84F60-42CB-11CE-8135-00AA004BB851),
        helpstring("IPersistPropertyBag interface")
    ]
    interface IPersistPropertyBag : IPersist {

        HRESULT InitNew();

        HRESULT Load(
            [in] IPropertyBag *pPropBag,
            [in] IErrorLog *pErrorLog);

        HRESULT Save(
            [in] IPropertyBag *pPropBag,
            [in] LONG fClearDirty,
            [in] LONG fSaveAllProperties);
    }

    [
        odl,
        uuid(22F55881-280B-11d0-A8A9-00A0C90C2004),
    ]
    interface IPersistPropertyBag2 : IPersist {

        HRESULT InitNew(void);

        HRESULT Load(
            [in] IPropertyBag2 * pPropBag,
            [in] IErrorLog * pErrLog);

        HRESULT Save(
            [in] IPropertyBag2 * pPropBag,
            [in] BOOL fClearDirty,
            [in] BOOL fSaveAllProperties);

        LONG IsDirty(void);

    }

    [
        odl,
        uuid(BD1AE5E0-A6AE-11CE-BD37-504200C10000),
    ]
    interface IPersistMemory : IPersist {

        LONG IsDirty(void);

        HRESULT Load(
            [in] void *pMem,
            [in] LONG cbSize);

        HRESULT Save(
            [out] void *pMem,
            [in] BOOL fClearDirty,
            [in] LONG cbSize);

        HRESULT GetSizeMax(
            [out, retval] LONG *pCbSize);

        HRESULT InitNew(void);

    }



    [
        odl,
        uuid(79eac9c9-baf9-11ce-8c82-00aa004ba90b),
    ]
    interface IPersistMoniker : stdole.IUnknown {

        HRESULT GetClassID(
            [out] UUID *pClassID);

        LONG IsDirty(void);

        HRESULT Load(
            [in] BOOL fFullyAvailable,
            [in] IMoniker *pimkName,
            [in] IBindCtx *pibc,
            [in] LONG grfMode);

        HRESULT Save(
            [in] IMoniker *pimkName,
            [in] IBindCtx *pbc,
            [in] BOOL fRemember);

        HRESULT SaveCompleted(
            [in] IMoniker *pimkName,
            [in] IBindCtx *pibc);

        HRESULT GetCurMoniker(
            [out, retval] IMoniker ** ppimkName);

    }
