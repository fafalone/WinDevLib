//tbShellLib - 100% implemented

[
    odl,
    uuid(00000114-0000-0000-C000-000000000046)
]
interface IOleWindow : stdole.IUnknown {

    HRESULT GetWindow(
        [out, retval] long *phwnd);

    HRESULT ContextSensitiveHelp(
        [in] BOOL fEnterMode);

};

