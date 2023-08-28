//tbShellLib - 100% implemented

[
    odl,
    uuid(0000000a-0000-0000-C000-000000000046),
]
interface ILockBytes : stdole.IUnknown {

    HRESULT ReadAt(
        [in] CURRENCY ulOffset,
        [out] void *pv,
        [in] LONG cb,
        [out, retval] LONG *pcbRead);

    HRESULT WriteAt(
        [in] CURRENCY ulOffset,
        [in] void *pv,
        [in] LONG cb,
        [out, retval] LONG *pcbWritten);

    HRESULT Flush();

    HRESULT SetSize(
        [in] CURRENCY cb);

    HRESULT LockRegion(
        [in] CURRENCY libOffset,
        [in] CURRENCY cb,
        [in] LONG dwLockType);

    HRESULT UnlockRegion(
        [in] CURRENCY libOffset,
        [in] CURRENCY cb,
        [in] LONG dwLockType);

    HRESULT Stat(
        [out] STATSTG *pstatstg,
        [in, defaultvalue(0)] STATFLAG grfStatFlag);

}

[
    odl,
    uuid(99caf010-415e-11cf-8814-00aa00b569f5),
]
interface IFillLockBytes: stdole.IUnknown {

    HRESULT FillAppend(
        [in] void *pv,
        [in] LONG cb,
        [out, retval] LONG *pcbWritten);

    HRESULT FillAt(
        [in] CURRENCY ulOffset,
        [in] void *pv,
        [in] LONG cb,
        [out, retval] LONG *pcbWritten);

    HRESULT SetFillSize(
        [in] CURRENCY ulSize);

    HRESULT Terminate(
        [in] BOOL bCanceled);
}
