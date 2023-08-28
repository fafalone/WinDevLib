//tbShellLib - 100% implemented

    [
        odl,
        uuid(00000012-0000-0000-C000-000000000046),
    ]
    interface IRootStorage : stdole.IUnknown {

        HRESULT SwitchToFile(
            [in] LPWSTR pszFile);
    }

    [
        odl,
        uuid(0000000b-0000-0000-C000-000000000046),
    ]
    interface IStorage : stdole.IUnknown {

        HRESULT CreateStream(
            [in] LPWSTR pwcsName,
            [in] STGM grfMode,
            [in, defaultvalue(0)] LONG reserved1,
            [in, defaultvalue(0)] LONG reserved2,
            [out, retval] IStream **ppstm);

        HRESULT OpenStream(
            [in] LPWSTR pwcsName,
            [in] LONG reserved1,
            [in] STGM grfMode,
            [in, defaultvalue(0)] LONG reserved2,
            [out, retval] IStream **ppstm);

        HRESULT CreateStorage(
            [in] LPWSTR pwcsName,
            [in] STGM grfMode,
            [in, defaultvalue(0)] LONG reserved1,
            [in, defaultvalue(0)] LONG reserved2,
            [out, retval] IStorage **ppstg);

        HRESULT OpenStorage(
            [in] LPWSTR pwcsName,
            [in] LONG pstgPriority,
            [in] STGM grfMode,
            [in, defaultvalue(0)] LONG snbExclude,
            [in, defaultvalue(0)] LONG reserved,
            [out, retval] IStorage **ppstg);

        HRESULT CopyTo(
            [in] LONG ciidExclude,
            [in] void *rgiidExclude,
            [in] SNB snbExclude,
            [in] IStorage *pstgDest);

        HRESULT MoveElementTo(
            [in] LPWSTR pwcsName,
            [in] IStorage *pstgDest,
            [in] LPWSTR *pwcsNewName,
            [in] STGMOVE grfFlags);

        HRESULT Commit(
            [in, defaultvalue(0)] STGC grfCommitFlags);

        HRESULT Revert();

        HRESULT EnumElements(
            [in, defaultvalue(0)] LONG reserved1,
            [in, defaultvalue(0)] LONG reserved2,
            [in, defaultvalue(0)] LONG reserved3,
            [out, retval] IEnumSTATSTG **ppenum);

        HRESULT DestroyElement(
            [in] LPWSTR pwcsName);

        HRESULT RenameElement(
            [in] LPWSTR pwcsOldName,
            [in] LPWSTR pwcsNewName);

        HRESULT SetElementTimes(
            [in] LPWSTR pwcsName,
            [in] CURRENCY *pctime,
            [in] CURRENCY *patime,
            [in] CURRENCY *pmtime);

        HRESULT SetClass(
            [in] UUID *clsid);

        HRESULT SetStateBits(
            [in] LONG grfStateBits,
            [in] LONG grfMask);

        HRESULT Stat(
            [out] STATSTG *pstatstg,
            [in, defaultvalue(0)] STATFLAG grfStatFlag);
    }

    [
        odl,
        uuid(0e6d4d90-6738-11cf-9608-00aa00680db4),
    ]
    interface ILayoutStorage: stdole.IUnknown {

        typedef struct StorageLayout {
            LONG LayoutType;
            LONG pwcsElementName;
            CURRENCY cOffset;
            CURRENCY cBytes;
        } StorageLayout;

        HRESULT LayoutScript (
            [in] StorageLayout *pStorageLayout,
            [in] LONG nEntries,
            [in] LONG glfInterleavedFlag);

        HRESULT BeginMonitor(void);

        HRESULT EndMonitor(void);

        HRESULT ReLayoutDocfile (
            [in] LPWSTR pwcsNewDfName);

        HRESULT ReLayoutDocfileOnILockBytes(
            [in] ILockBytes *pILockBytes);

    }
