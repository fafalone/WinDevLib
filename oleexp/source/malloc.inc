//tbShellLib - 100% implemented

[
    odl,
    uuid(00000002-0000-0000-C000-000000000046)
]

interface IMalloc : stdole.IUnknown {

    LONG Alloc(
        [in] LONG cb);

    LONG Realloc(
        [in] void *pv,
        [in] LONG cb);

    LONG Free(
        [in] LONG pv);

    LONG GetSize(
        [in] LONG pv);

    LONG DidAlloc(
        [in] LONG pv);

    void HeapMinimize(void);

}

