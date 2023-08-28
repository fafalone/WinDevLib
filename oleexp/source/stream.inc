//tbShellLib - 100% implemented   

[
        odl,
        uuid(0c733a30-2a1c-11ce-ade5-00aa0044773d),
    ]
    interface ISequentialStream : stdole.IUnknown {

        HRESULT Read(
            [out] void *pv,
            [in] LONG cb,
            [out, retval] LONG *pcbRead);

        HRESULT Write(
            [in] void *pv,
            [in] LONG cb,
            [out, retval] LONG *pcbWritten);

    }

    [
        odl,
        uuid(0000000c-0000-0000-C000-000000000046),
    ]
    interface IStream : ISequentialStream {

        HRESULT Seek(
            [in] CURRENCY dlibMove,
            [in] LONG dwOrigin,
            [out, retval] CURRENCY *plibNewPosition);

        HRESULT SetSize(
            [in] CURRENCY libNewSize);

        HRESULT CopyTo(
            [in] IStream *pstm,
            [in] CURRENCY cb,
            [out] CURRENCY *pcbRead,
            [out] CURRENCY *pcbWritten);

        HRESULT Commit(
            [in, defaultvalue(0)] STGC grfCommitFlags);

        HRESULT Revert();

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

        HRESULT Clone(
            [out, retval] IStream **ppstm);

    }
