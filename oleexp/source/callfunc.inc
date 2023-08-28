[uuid(9fb58518-92ec-4bf6-bc61-ff4e59df7369), odl]
interface ICallQI : stdole.IUnknown {
	long QueryInterface([in] int This, [in] UUID *riid, [out] void *ppvObj);
}

[uuid(9fb58519-92ec-4bf6-bc61-ff4e59df7369), odl]
interface ICallAddRelease : stdole.IUnknown {
	LONG AddRelease([in] int This);
}

[uuid(9fb58520-92ec-4bf6-bc61-ff4e59df7369), odl]
interface ICallGION : stdole.IUnknown {
    LONG GetIDsOfNames(
        [in] int This,
        [in] UUID *riid,
        [in] LONG *rgszNames,
        [in] int cNames,
        [in] long lcid,
        [out] long *rgdispid);
}

[uuid(9fb58521-92ec-4bf6-bc61-ff4e59df7369), odl]
interface ICallInvoke : stdole.IUnknown {
    LONG Invoke(
        [in] int This,
        [in] long dispidMember,
        [in] UUID* riid,
        [in] long lcid,
        [in] short wFlags,
        [in] DISPPARAMS *pdispparams,
        [in] long pvarResult,
        [out] EXCEPINFO *pexcepinfo,
        [out] int *puArgErr);
}