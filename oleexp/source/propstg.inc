//tbShellLib: 100% complete    
[
        odl,
        uuid(00000138-0000-0000-C000-000000000046),
        helpstring("IPropertyStorage interface")
    ]
    interface IPropertyStorage : stdole.IUnknown {

        typedef enum PROPSETFLAG {
            PROPSETFLAG_DEFAULT    = 0,
            PROPSETFLAG_NONSIMPLE  = 1,
            PROPSETFLAG_ANSI       = 2,
            PROPSETFLAG_UNBUFFERED = 4,
            PROPSETFLAG_CASE_SENSITIVE = 8,
        } PROPSETFLAG;

        typedef enum PRPSPEC {
            PRSPEC_LPWSTR  = 0,
            PRSPEC_PROPID  = 1,
            PRSPEC_INVALID = 0xffffffff,
        } PRPSPEC;

        typedef struct PROPSPEC {
            PRPSPEC ulKind;
            LONG    ID_or_LPWSTR;
        } PROPSPEC;

        typedef struct STATPROPSTG {
            LONG  lpwstrName;
            LONG  propid;
            short vt;
        } STATPROPSTG;

        typedef struct STATPROPSETSTG {
            UUID     fmtid;
            UUID     clsid;
            LONG     grfFlags;
            CURRENCY mtime;
            CURRENCY ctime;
            CURRENCY atime;
            LONG     dwOSVersion;
        } STATPROPSETSTG;

        HRESULT ReadMultiple(
            [in] LONG cpspec,
            [in] PROPSPEC *rgpspec,
            [in] VARIANT *rgpropvar);

        HRESULT WriteMultiple(
            [in] LONG cpspec,
            [in] PROPSPEC *rgpspec,
            [in] VARIANT *rgPropvar,
            [in, defaultvalue(2)] LONG propidNameFirst);

        HRESULT DeleteMultiple(
            [in] LONG cpspec,
            [in] PROPSPEC *rgpspec);

        HRESULT ReadPropertyNames(
            [in] LONG cpropid,
            [in] LONG *rgpropid,
            [out, retval] LONG *rglpwstrName);

        HRESULT WritePropertyNames(
            [in] LONG cpropid,
            [in] LONG *rgpropid,
            [in] LPWSTR *rglpwstrName);

        HRESULT DeletePropertyNames(
            [in] LONG cpropid,
            [in] LONG *rgpropid);

        HRESULT Commit(
            [in, defaultvalue(0)] STGC grfCommitFlags);

        HRESULT Revert();

        HRESULT Enum(
            [out, retval] IEnumSTATPROPSTG **ppenum);

        HRESULT SetTimes(
            [in] CURRENCY *pctime,
            [in] CURRENCY *patime,
            [in] CURRENCY *pmtime);

        HRESULT SetClass(
            [in] UUID *clsid);

        HRESULT Stat(
            [out] STATPROPSETSTG *pstatpsstg);
    }

    [
        odl,
        uuid(0000013A-0000-0000-C000-000000000046),
        helpstring("IPropertySetStorage interface")
    ]
    interface IPropertySetStorage : stdole.IUnknown {

        HRESULT Create(
            [in] UUID *rfmtid,
            [in] UUID *pclsid,
            [in] PROPSETFLAG grfFlags,
            [in, defaultvalue(0x12)] STGM grfMode,
            [out, retval] IPropertyStorage **ppprstg);

        HRESULT Open(
            [in] UUID *rfmtid,
            [in, defaultvalue(0x12)] STGM grfMode,
            [out, retval] IPropertyStorage **ppprstg);

        HRESULT Delete(
            [in] UUID *rfmtid);

        HRESULT Enum(
            [out, retval] IEnumSTATPROPSETSTG **ppenum);
    }
