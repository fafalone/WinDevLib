
    [
        odl,
        uuid(000214F2-0000-0000-C000-000000000046),
        helpstring("IEnumIDList interface")
    ]
    interface IEnumIDList : stdole.IUnknown {

        LONG Next(
            [in] LONG celt,
            [in, out] LONG *rgelt,
            [in, out, defaultvalue(0)] LONG *pceltFetched);

        HRESULT Skip(
            [in] LONG celt);

        HRESULT Reset();

        HRESULT Clone(
            [out, retval] IEnumIDList **ppenum);
    }

    [
      odl,
      uuid(000214E6-0000-0000-C000-000000000046),
      helpstring("IShellFolder interface")
    ]
    interface IShellFolder : stdole.IUnknown {

        HRESULT ParseDisplayName(
            [in] long hwndOwner,
            [in] long pbcReserved,
            [in] LPWSTR lpszDisplayName,
            [in, out] long* pchEaten,
            [in, out] long* ppidl,
            [in, out] long* pdwAttributes);

        HRESULT EnumObjects(
            [in] long hwndOwner,
            [in] SHCONTF grfFlags,
            [out, retval] IEnumIDList** ppenumIDList);

        HRESULT BindToObject(
            [in] long pidl,
            [in] long pbcReserved,
            [in, out] UUID *riid,
            [out, retval] IUnknown **ppvOut);

        HRESULT BindToStorage(
            [in] long pidl,
            [in] long pbcReserved,
            [in, out] UUID *riid,
            [out, retval] IStorage **ppvObj);

        LONG CompareIDs(
            [in] long lparam,
            [in] long pidl1,
            [in] long pidl2);

        HRESULT CreateViewObject(
            [in] long hwndOwner,
            [in, out] UUID* riid,
            [out, retval] stdole.IUnknown **ppvOut);

        HRESULT GetAttributesOf(
            [in] long cidl,
            [in, out] long* apidl,
            [in, out] long* rgfInOut);

        HRESULT GetUIObjectOf(
            [in] long hwndOwner,
            [in] long cidl,
            [in, out] long *apidl,
            [in, out] UUID  *riid,
            [in, out] long *prgfInOut,
            [out, retval] stdole.IUnknown **ppvOut);

        HRESULT GetDisplayNameOf(
            [in] long pidl,
            [in] SHGNO_Flags uFlags,
            [in, out] STRRET* lpName);

        HRESULT SetNameOf(
            [in] long hwndOwner,
            [in] long pidl,
            [in] LPWSTR lpszName,
            [in] SHGNO_Flags uFlags,
            [out, retval] long* ppidlOut);
    };
