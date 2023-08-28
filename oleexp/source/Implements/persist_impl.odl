[
    odl,
    uuid(0000010b-0000-0000-C000-000000000046),
]
interface IPersistFile : stdole.IUnknown {

    HRESULT GetClassID(
        [in, out] UUID *pClassID);

    HRESULT IsDirty();

    HRESULT Load(
        [in] LONG pszFileName,
        [in] STGM dwMode);

    HRESULT Save(
        [in] LONG pszFileName,
        [in] BOOL fRemember);

    HRESULT SaveCompleted(
        [in] LONG pszFileName);

    HRESULT GetCurFile(
        [in, out] LONG *ppszFileName);
}

[
    odl,
    uuid(00000109-0000-0000-C000-000000000046),
]
interface IPersistStream : stdole.IUnknown {

    HRESULT GetClassID(
        [in, out] UUID *pClassID);

    HRESULT IsDirty(void);

    HRESULT Load(
        [in] IStream *pStm);

    HRESULT Save(
        [in] IStream *pStm,
        [in] BOOL fClearDirty);

    HRESULT GetSizeMax(
        [out, retval] CURRENCY *pcbSize);

}
