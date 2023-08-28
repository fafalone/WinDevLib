//tbShellLib - 100% implemented

[
    odl,
    uuid(0000002F-0000-0000-C000-000000000046)
]
interface IRecordInfo: stdole.IUnknown {

    HRESULT RecordInit(
        [in, out] void *pvNew);

    HRESULT RecordClear(
        [in] void *pvExisting);

    HRESULT RecordCopy(
        [in] void *pvExisting,
        [out] void *pvNew);

    HRESULT GetGuid(
        [out] UUID *pguid);

    HRESULT GetName(
        [out, retval] BSTR *pbstrName);

    HRESULT GetSize(
        [out, retval] LONG *pcbSize);

    HRESULT GetTypeInfo(
        [out, retval] ITypeInfo **ppTypeInfo);

    HRESULT GetField(
        [in] void *pvData,
        [in] LPWSTR szFieldName,
	[out, retval] VARIANT *pvarField);

    HRESULT GetFieldNoCopy(
        [in] void *pvData,
        [in] LPWSTR szFieldName,
        [out] VARIANT *pvarField,
        [out] void *ppvDataCArray);

    HRESULT PutField(
        [in] INVOKEKIND wFlags,
        [in, out] void *pvData,
        [in] LPWSTR szFieldName,
        [in] VARIANT *pvarField);

    HRESULT PutFieldNoCopy(
        [in] INVOKEKIND wFlags,
        [in,out] void *pvData,
        [in] LPWSTR szFieldName,
        [in] VARIANT * pvarField);

    HRESULT GetFieldNames(
        [in,out] LONG *pcNames,
        [out] BSTR *rgBstrNames);

    BOOL IsMatchingType(
        [in] IRecordInfo *pRecordInfo);

    LONG RecordCreate();

    HRESULT RecordCreateCopy(
        [in] void *pvSource,
        [out] void *ppvDest);

    HRESULT RecordDestroy(
        [in] void *pvRecord);

}
