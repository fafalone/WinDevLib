    [
        odl,
        uuid(0c733a30-2a1c-11ce-ade5-00aa0044773d),
    ]
    interface ISequentialStream : stdole.IUnknown {

        HRESULT Read(
            [in] LONG pv,
            [in] LONG cb,
            [in, out] LONG *pcbRead);

        HRESULT Write(
            [in] LONG pv,
            [in] LONG cb,
            [in, out] LONG *pcbWritten);

    }

    [
        odl,
        uuid(0000000c-0000-0000-C000-000000000046)
    ]
    interface IStream : stdole.IUnknown {

        HRESULT Read(
            [in] LONG pv,
            [in] LONG cb,
            [in, out] LONG *pcbRead);

        HRESULT Write(
            [in] LONG pv,
            [in] LONG cb,
            [in, out] LONG *pcbWritten);

        HRESULT Seek(
            [in] CURRENCY dlibMove,
            [in] STREAM_SEEK dwOrigin,
            [in, out] CURRENCY *plibNewPosition);

        HRESULT SetSize(
            [in] CURRENCY libNewSize);

        HRESULT CopyTo(
            [in] IStream *pstm,
            [in] CURRENCY cb,
            [in, out] CURRENCY *pcbRead,
            [in, out] CURRENCY *pcbWritten);

        HRESULT Commit(
            [in] STGC grfCommitFlags);

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
            [in, out] STATSTG *pstatstg,
            [in] STATFLAG grfStatFlag);

        HRESULT Clone(
            [out, retval] IStream **ppstm);


    }
