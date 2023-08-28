
//tbShellLib - 100% implemented

//****** IEnumSTATPROPSTG

[
    odl,
    uuid(00000139-0000-0000-C000-000000000046),
    helpstring("IEnumSTATPROPSTG interface")
]
interface IEnumSTATPROPSTG : stdole.IUnknown {

    LONG Next(
        [in] LONG celt,
        [in, out] STATPROPSTG *pSTATPROPSTG,
        [out, defaultvalue(0)] LONG *pceltFetched);

    HRESULT Skip(
            [in] LONG celt);

    HRESULT Reset();

    HRESULT Clone(
        [out, retval] IEnumSTATPROPSTG **ppenum);
}

//****** IEnumSTATPROPSETSTG

[
    odl,
    uuid(0000013B-0000-0000-C000-000000000046),
    helpstring("IEnumSTATPROPSETSTG interface")
]
interface IEnumSTATPROPSETSTG : stdole.IUnknown {

    LONG Next(
        [in] LONG celt,
        [in, out] STATPROPSETSTG *pSTATPROPSETSTG,
        [out, defaultvalue(0)] LONG *pceltFetched);

    HRESULT Skip(
            [in] LONG celt);

    HRESULT Reset();

    HRESULT Clone(
        [out, retval] IEnumSTATPROPSETSTG **ppenum);
}

//****** IEnumSTATSTG

[
    odl,
    uuid(0000000d-0000-0000-C000-000000000046),
]
interface IEnumSTATSTG : stdole.IUnknown {

    LONG Next(
        [in] LONG celt,
        [out] STATSTG *rgelt,
        [out, defaultvalue(0)] LONG *pceltFetched);

    HRESULT Skip(
        [in] LONG celt);

    HRESULT Reset();

    HRESULT Clone(
        [out, retval] IEnumSTATSTG **ppenum);
}

// ****** IEnumSTATDATA

[
  odl,
  uuid(00000105-0000-0000-C000-000000000046)
]
interface IEnumSTATDATA : stdole.IUnknown {

    LONG Next(
        [in] long celt,
        [out] STATDATA* rgelt,
        [out, defaultvalue(0)] long* pceltFetched);

    HRESULT Skip(
		[in] long celt);

    HRESULT Reset();

    HRESULT Clone(
		[out, retval] IEnumSTATDATA** ppenum);

};

//****** IEnumString

[
    odl,
    uuid(00000101-0000-0000-C000-000000000046),
]
interface IEnumString : stdole.IUnknown {

    LONG Next(
        [in] LONG celt,
        [out] LONG *rgelt,
        [out, defaultvalue(0)] LONG *pceltFetched);

    HRESULT Skip(
        [in] LONG celt);

    HRESULT Reset();

    HRESULT Clone(
        [out, retval] IEnumString **ppenum);

}

//****** IEnumMoniker

[
    odl,
    uuid(00000102-0000-0000-C000-000000000046),
]
interface IEnumMoniker : stdole.IUnknown {

    LONG Next(
        [in] LONG celt,
        [out] IMoniker **rgelt,
        [out, defaultvalue(0)] LONG *pceltFetched);

    HRESULT Skip(
        [in] LONG celt);

    HRESULT Reset();

    HRESULT Clone(
        [out, retval] IEnumMoniker **ppenum);
}

//****** IEnumFORMATETC

[
    odl,
    uuid(00000103-0000-0000-C000-000000000046),
]
interface IEnumFORMATETC : stdole.IUnknown {

    LONG Next(
        [in] LONG celt,
        [in, out] FORMATETC *rgelt,
        [out, defaultvalue(0)] LONG *pceltFetched);

    HRESULT Skip(
        [in] LONG celt);

    HRESULT Reset();

    HRESULT Clone(
        [out, retval] IEnumFORMATETC **ppenum);
}


[
    odl,
    uuid(00000100-0000-0000-C000-000000000046)
]
interface IEnumUnknown : stdole.IUnknown {

    LONG Next(
        [in] long celt,
        [out] IUnknown** rgelt,
        [out, defaultvalue(0)] long* pceltFetched);

    HRESULT Skip([in] long celt);

    HRESULT Reset();

    HRESULT Clone([out, retval] IEnumUnknown** ppenum);

};

[
    odl,
    uuid(00000104-0000-0000-C000-000000000046)
]
interface IEnumOLEVERB : stdole.IUnknown {

    long Next(
        [in] long celt,
        [in, out] OLEVERB *rgelt,
        [out, defaultvalue(0)] long* pceltFetched);

    HRESULT Skip(
        [in] long celt);

    HRESULT Reset();

    HRESULT Clone(
        [out, retval] IEnumOLEVERB** ppenum);
};

[
    odl,
    uuid(0002E000-0000-0000-C000-000000000046)
]
interface IEnumGUID : stdole.IUnknown {

    LONG Next(
        [in] LONG celt,
        [out] UUID *rgelt,
        [out, defaultvalue(0)] LONG *pceltFetched);

    HRESULT Skip(
        [in] LONG celt);

    HRESULT Reset();

    HRESULT Clone(
        [out, retval ] IEnumGUID **ppenum);
}

[
    odl,
    uuid(0002E011-0000-0000-C000-000000000046)
]
interface IEnumCATEGORYINFO : stdole.IUnknown {

    typedef struct CATEGORYINFO {
        UUID catid;
        LONG lcid;
        BYTE szDescription[256];
    } CATEGORYINFO;

    LONG Next(
        [in] LONG celt,
        [out] CATEGORYINFO* rgelt,
        [out, defaultvalue(0)] LONG* pceltFetched);

    HRESULT Skip(
        [in] LONG celt);

    HRESULT Reset();

    HRESULT Clone(
        [out, retval] IEnumCATEGORYINFO** ppenum);
}


[
    odl,
    uuid(00020404-0000-0000-C000-000000000046)
]
interface IEnumVARIANT : stdole.IUnknown 
{
    LONG Next(
        [in] long celt,
        [in, out] VARIANT* rgvar,
        [in, out, defaultvalue(0)] long* pceltFetched);

    HRESULT Skip([in] long celt);

    HRESULT Reset();

    HRESULT Clone(
        [out] IEnumVARIANT **ppenum);

}


    [
        odl,
        uuid(B196B287-BAB4-101A-B69C-00AA00341D07),
    ]
    interface IEnumConnections : stdole.IUnknown {

        typedef struct CONNECTDATA {
            IDispatch *pUnk;
            LONG dwCookie;
        } CONNECTDATA;

        LONG Next(
            [in] LONG cConnections,
            [in, out] CONNECTDATA *rgcd,
            [in, out, defaultvalue(0)] LONG *lpcFetched);

        HRESULT Skip(
            [in] LONG cConnections);

        HRESULT Reset(void);

        HRESULT Clone(
            [out, retval] IEnumConnections ** ppEnum);
    }

    [
        odl,
        uuid(B196B285-BAB4-101A-B69C-00AA00341D07),
    ]
    interface IEnumConnectionPoints : stdole.IUnknown {

        LONG Next(
            [in] LONG cConnections,
            [in, out] IConnectionPoint **rgpcn,
            [out, defaultvalue(0)] LONG *lpcFetched);

        HRESULT Skip(
            [in] LONG cConnections);

        HRESULT Reset(void);

        HRESULT Clone(
            [out, retval]   IEnumConnectionPoints **ppEnum);
    }

