    [
        odl,
        uuid(00000104-0000-0000-C000-000000000046),
        helpstring("Implementable IEnumOLEVERB interface")
    ]
    interface IEnumOLEVERB : stdole.IUnknown {

        HRESULT Next(
            [in] long celt, 
            [in, out] OLEVERB* rgelt, 
            [out, retval] long* pceltFetched);

        HRESULT Skip([in] long celt);

        HRESULT Reset();

        HRESULT Clone([out, retval] IEnumOLEVERB** ppenum);

    };

[
    odl,
    uuid(00000101-0000-0000-C000-000000000046),
]
interface IEnumString : stdole.IUnknown {

    HRESULT Next(
        [in] LONG celt,
        [in, out] LONG *rgelt,
        [in, out] LONG *pceltFetched);

    HRESULT Skip(
        [in] LONG celt);

    HRESULT Reset();

    HRESULT Clone(
        [in, out] IEnumString **ppenum);

}

