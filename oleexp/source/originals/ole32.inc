[
    dllname("OLE32.DLL"),
    helpstring("Storage functions")
]
module ole32 {

    [entry("CoCreateGuid")]
    HRESULT CoCreateGuid(
        [out] UUID *pguid);

    [entry("CoTaskMemAlloc")]
    LONG CoTaskMemAlloc(
        [in] LONG cb);

    [entry("CoTaskMemRealloc")]
    LONG CoTaskMemRealloc(
        [in] LONG pv,
        [in] LONG cb);

    [entry("CoTaskMemFree")]
    LONG CoTaskMemFree(
        [in] LONG Ptr);

    [entry("StgIsStorageFile")]
    LONG StgIsStorageFile(
        [in] LPWSTR lpwcsFileName);

    [entry("StgOpenStorage")]
    HRESULT StgOpenStorage(
        [in] LPWSTR pwcsName,
        [in] IStorage *pstgPriority,
        [in] STGM grfMode,
        [in] SNB snbExclude,
        [in, defaultvalue(0)] LONG reserved,
        [out, retval] IStorage **ppstgOpen);

    [entry("StgCreateDocfile")]
    HRESULT StgCreateDocfile(
        [in] LPWSTR pwcsName,
        [in] STGM grfMode,
        [in, defaultvalue(0)] LONG reserved,
        [out, retval] IStorage **ppstgOpen);

    [entry("GetHGlobalFromILockBytes")]
    HRESULT GetHGlobalFromILockBytes(
        [in] ILockBytes *plkbyt,
        [out, retval] LONG *phglobal);

    [entry("CreateILockBytesOnHGlobal")]
    HRESULT CreateILockBytesOnHGlobal(
        [in] LONG hGlobal,
        [in] BOOL fDeleteOnRelease,
        [out, retval] IlockBytes **pplkbyt);

    [entry("GetHGlobalFromStream")]
    HRESULT GetHGlobalFromStream(
        [in] IStream *pstm,
        [out, retval] LONG *phglobal);

    [entry("CreateStreamOnHGlobal")]
    HRESULT CreateStreamOnHGlobal(
        [in] LONG hGlobal,
        [in] BOOL fDeleteOnRelease,
        [out, retval] IStream **ppstm);

    [entry("StgCreateDocfileOnILockBytes")]
    HRESULT StgCreateDocfileOnILockBytes(
        [in] ILockBytes *plkbyt,
        [in] STGM grfMode,
        [in] LONG reserved,
        [out, retval] IStorage **ppstgOpen);

    [entry("StgOpenStorageOnILockBytes")]
    HRESULT StgOpenStorageOnILockBytes(
        [in] ILockBytes *plkbyt,
        [in] IStorage *pstgPriority,
        [in] STGM grfMode,
        [in] SNB snbExclude,
        [in, defaultvalue(0)] LONG reserved,
        [out, retval] IStorage **ppstgOpen);

    [entry("StgIsStorageILockBytes")]
    LONG StgIsStorageILockBytes(
        [in] ILockBytes *plkbyt);

    [entry("ReadClassStg")]
    HRESULT ReadClassStg(
        [in] IStorage *pStg,
        [out] UUID *pclsid);

    [entry("WriteClassStg")]
    HRESULT WriteClassStg(
        [in] IStorage *pStg,
        [in] UUID *rclsid);

    [entry("ReadClassStm")]
    HRESULT ReadClassStm(
        [in] IStream *pStg,
        [out] UUID *rclsid);

    [entry("WriteClassStm")]
    HRESULT WriteClassStm(
        [in] IStream *pStg,
        [in] UUID *rclsid);

    [entry("PropVariantClear")]
    HRESULT PropVariantClear(
        [in] VARIANT* pvarg);

    [entry("CLSIDFromProgID")]
    HRESULT CLSIDFromProgID(
        [in] LPWSTR lpszProgID,
        [in, out] UUID *lpclsid);

    [entry("CLSIDFromString")]
    HRESULT CLSIDFromString(
        [in] LPWSTR lpszProgID,
        [in, out] UUID *lpclsid);

    [entry("ProgIDFromCLSID")]
    void ProgIDFromCLSID(
        [in] UUID *clsid,
        [in, out] LONG *lplpszProgID);

    [entry("StringFromGUID2")]
    LONG StringFromGUID2(
        [in] UUID *UUID,
        [in, out] LPWSTR lplpsz,
        [in] LONG cbMax);

    [entry("BindMoniker")]
    HRESULT BindMoniker(
        [in] IMoniker *pmk,
        [in] LONG grfOpt,
        [in] UUID *iidResult,
        [out] void *ppvResult);

    [entry("CoGetObject")]
    HRESULT CoGetObject(
        [in] LPWSTR pszName,
        [in] BIND_OPTS *pBindOptions,
        [in] UUID *riid,
        [out] void *ppv);

    [entry("MkParseDisplayName")]
    HRESULT MkParseDisplayName(
        [in] IBindCtx *pbc,
        [in] LPWSTR szUserName,
        [out, defaultvalue(0)] LONG *pchEaten,
        [out, retval] IMoniker **ppmk);

    [entry("MonikerRelativePathTo")]
    HRESULT MonikerRelativePathTo(
        [in] IMoniker *pmkSrc,
        [in] IMoniker *pmkDest,
        [out] IMoniker **ppmkRelPath,
        [in] BOOL dwReserved);

    [entry("MonikerCommonPrefixWith")]
    HRESULT MonikerCommonPrefixWith(
        [in] IMoniker *pmkThis,
        [in] IMoniker *pmkOther,
        [out, retval] IMoniker **ppmkCommon);

    [entry("CreateBindCtx")]
    HRESULT CreateBindCtx(
        [in, defaultvalue(0)] LONG reserved,
        [out, retval] IBindCtx **ppbc);

    [entry("CreateGenericComposite")]
    HRESULT CreateGenericComposite(
        [in] IMoniker *pmkFirst,
        [in] IMoniker *pmkRest,
        [out, retval] IMoniker **ppmkComposite);

    [entry("GetClassFile")]
    HRESULT GetClassFile(
        [in] LPWSTR szFilename,
        [out] UUID *pclsid);

    [entry("CreateClassMoniker")]
    HRESULT CreateClassMoniker(
        [in] UUID *rclsid,
        [out, retval] IMoniker **ppmk);

    [entry("CreateFileMoniker")]
    HRESULT CreateFileMoniker(
        [in] LPWSTR lpszPathName,
        [out, retval] IMoniker **ppmk);

    [entry("CreateItemMoniker")]
    HRESULT CreateItemMoniker(
        [in] LPWSTR lpszDelim,
        [in] LPWSTR lpszItem,
        [out, retval] IMoniker **ppmk);

    [entry("CreateAntiMoniker")]
    HRESULT CreateAntiMoniker(
        [out, retval] IMoniker **ppmk);

    [entry("CreatePointerMoniker")]
    HRESULT CreatePointerMoniker(
        [in] stdole.IUnknown *punk,
        [out, retval] IMoniker **ppmk);

    [entry("CreateObjrefMoniker")]
    HRESULT CreateObjrefMoniker(
        [in] stdole.IUnknown *punk,
        [out, retval] IMoniker **ppmk);

    [entry("GetRunningObjectTable")]
    HRESULT GetRunningObjectTable(
        [in, defaultvalue(0)] LONG reserved,
        [out, retval] IRunningObjectTable **pprot);

    [entry("CoGetClassObject")]
    HRESULT CoGetClassObject(
        [in] UUID *rclsid,
        [in] LONG dwClsContext,
        [in] void *pvReserved,
        [in] UUID *riid,
        [out] void *ppv);

    [entry("CoRegisterClassObject")]
    HRESULT CoRegisterClassObject(
        [in] UUID *rclsid,
        [in] stdole.IUnknown *pUnk,
        [in] LONG dwClsContext,
        [in] LONG flags,
        [out, retval] LONG *lpdwRegister);

    [entry("CoRevokeClassObject")]
    HRESULT CoRevokeClassObject(
        [in] LONG dwRegister);

    [entry("CoResumeClassObjects")]
    HRESULT CoResumeClassObjects(void);

    [entry("CoSuspendClassObjects")]
    HRESULT CoSuspendClassObjects(void);

/*        [entry("CoRegisterSurrogate")]
    HRESULT CoRegisterSurrogate(
			[in] ISurrogate *pSurrogate); */

    // Apartment model inter-thread interface passing helpers
    [entry("CoMarshalInterThreadInterfaceInStream")]
    HRESULT CoMarshalInterThreadInterfaceInStream(
        [in] UUID *riid,
        [in] stdole.IUnknown *pUnk,
        [out, retval] IStream **ppStm);

    [entry("CoGetInterfaceAndReleaseStream")]
    HRESULT CoGetInterfaceAndReleaseStream(
        [in] IStream *pStm,
        [in] UUID *iid,
        [out] void *ppv);

    [entry("CoLoadLibrary")]
    LONG CoLoadLibrary(
        [in] LPWSTR lpszLibName,
        [in] BOOL bAutoFree);

    [entry("CoFreeLibrary")]
    void CoFreeLibrary(
	        [in] LONG hInst);

    [entry("CoFreeAllLibraries")]
    void CoFreeAllLibraries(void);

    [entry("CoFreeUnusedLibraries")]
    void CoFreeUnusedLibraries(void);

    [entry("CoFileTimeNow")]
    HRESULT CoFileTimeNow(
        [out, retval] CURRENCY *lpFileTime);

    [entry("CreateStdProgressIndicator")]
    HRESULT CreateStdProgressIndicator(
        [in] LONG hwndParent,
        [in] LPWSTR pszTitle,
        [in] IBindStatusCallback *pIbscCaller,
        [out, retval] IBindStatusCallback ** ppIbsc);

    [entry("IsEqualGUID")]
    BOOL IsEqualGUID(
			[in] UUID *rguid1,
			[in] UUID *rguid2);

    [entry("CoBuildVersion")]
    LONG CoBuildVersion();

    [entry("CoInitialize")]
    HRESULT CoInitialize(
        [in] void *pvReserved);

    [entry("CoUninitialize")]
    void CoUninitialize();

    [entry("CoGetMalloc")]
    HRESULT CoGetMalloc(
        [in] MEMCTX dwMemContext,
        [out, retval] IMalloc **ppMalloc);

    [entry("CoCreateInstance")]
    LONG CoCreateInstance(
        [in] UUID *CLSID,
        [in] stdole.IUnknown *pUnkOuter,
        [in] CLSCTX dwClsContext,
        [in] UUID *IID,
        [out] void *ppv);

    [entry("CoGetCurrentProcess")]
    LONG CoGetCurrentProcess();

/*        [entry("CoRegisterMallocSpy")]
    HRESULT CoRegisterMallocSpy(
        [in] IMallocSpy *pMallocSpy);

    [entry("CoRevokeMallocSpy")]
    HRESULT CoRevokeMallocSpy(); */

    [entry("CoCreateStandardMalloc")]
    HRESULT CoCreateStandardMalloc(
        [in] MEMCTX memctx,
        [out, retval] IMalloc ** ppMalloc);

    [entry("CoGetMarshalSizeMax")]
    HRESULT CoGetMarshalSizeMax(
        [out] LONG *pulSize,
        [in] UUID *riid,
        [in] stdole.IUnknown *pUnk,
        [in] MSHCTX dwDestContext,        [in] void *pvDestContext,
        [in] MSHLFLAGS mshlflags);

    [entry("CoMarshalInterface")]
    HRESULT CoMarshalInterface(
        [in] IStream *pStm,
        [in] UUID *riid,
        [in] stdole.IUnknown *pUnk,
        [in] MSHCTX dwDestContext,
        [in] void *pvDestContext,
        [in] MSHLFLAGS mshlflags);

    [entry("CoUnmarshalInterface")]
    HRESULT CoUnmarshalInterface(
        [in] IStream *pStm,
        [in] UUID *riid,
        [out] void *ppv);

    [entry("CoMarshalHresult")]
    HRESULT CoMarshalHresult(
			[in] IStream *pstm,
			[in] LONG hresult);
			
    [entry("CoUnmarshalHresult")]
    HRESULT CoUnmarshalHresult(
		    [in] IStream *pstm,
		    [out, retval] LONG *phresult);
		
    [entry("CoReleaseMarshalData")]
    HRESULT CoReleaseMarshalData(
        [in] IStream *pStm);

    [entry("CoDisconnectObject")]
    HRESULT CoDisconnectObject(
        [in] stdole.IUnknown *pUnk,
        [in] LONG dwReserved);

    [entry("CoLockObjectExternal")]
    HRESULT CoLockObjectExternal(
        [in] stdole.IUnknown *pUnk,
        [in] BOOL fLock,
        [in] BOOL fLastUnlockReleases);

    [entry("CoGetStandardMarshal")]
    HRESULT CoGetStandardMarshal(
        [in] UUID *riid,
        [in] stdole.IUnknown *pUnk,
        [in] MSHCTX dwDestContext,
        [in] void *pvDestContext,
        [in] MSHLFLAGS  mshlflags,
        [out, retval] IMarshal **ppMarshal);

    [entry("CoGetStdMarshalEx")]
    HRESULT CoGetStdMarshalEx(
        [in] stdole.IUnknown *pUnkOuter,
        [in] STDMSHLFLAGS smexflags,
        [out, retval] stdole.IUnknown **ppUnkInner);

    [entry("CoCreateFreeThreadedMarshaler")]
    HRESULT CoCreateFreeThreadedMarshaler(
        [in] stdole.IUnknown*punkOuter,
        [out, retval] stdole.IUnknown **ppunkMarshal);

    [entry("ReleaseStgMedium")]
    void ReleaseStgMedium(
        [in] STGMEDIUM* STM);

    [entry("OleGetIconOfClass")]
    long OleGetIconOfClass(
        [in] UUID *rclsid,
        [in] LPWSTR lpszLabel,
        [in] long fUseTypeAsLabel);

    [entry("OleMetafilePictFromIconAndLabel")]
    long OleMetafilePictFromIconAndLabel(
        [in] long hIcon,
        [in] LPWSTR lpszLabel,
        [in] LPWSTR lpszSourceFile,
        [in, defaultvalue(0)] long iIconIndex);

    [entry("OleSetContainedObject")]
    long OleSetContainedObject(
        [in] stdole.IUnknown *pUnk,
        [in, defaultvalue(1)] long fContained);

    [entry("OleCreate")]
    HRESULT OleCreate(
        [in] UUID *rclsid,
        [in] UUID *riid,
        [in] OLERENDER renderopt,
        [in] long pFormatEtc,
        [in] IOleClientSite *pClientSite,
        [in] IStorage *pStg,
        [out] void *ppvObject);

    [entry("OleCreateEx")]
    HRESULT OleCreateEx(
        [in] UUID *rclsid,
        [in] UUID *riid,
        [in] OLECREATE_UDT dwFlags,
        [in] OLERENDER renderopt,
        [in] LONG cFormats,
        [in] LONG* rgAdvf,
        [in] FORMATETC *rgFormatEtc,
        [in] IAdviseSink *lpAdviseSink,
        [out] LONG *rgdwConnection,
        [in] IOleClientSite *pClientSite,
        [in] IStorage *pStg,
        [out] void *ppvObj);

    [entry("OleCreateFromData")]
    HRESULT OleCreateFromData(
        [in] IdataObject *pSrcDataObj,
        [in] UUID *riid,
        [in] OLERENDER renderopt,
        [in] FORMATETC *pFormatEtc,
        [in] IOleClientSite *pClientSite,
        [in] IStorage *pStg,
        [out] void *ppvObj);

    [entry("OleCreateFromDataEx")]
    HRESULT OleCreateFromDataEx(
        [in] IDataObject *pSrcDataObj,
        [in] UUID *riid,
        [in] OLECREATE_UDT dwFlags,
        [in] OLERENDER renderopt,
        [in] LONG cFormats,
        [in] LONG *rgAdvf,
        [in] FORMATETC *rgFormatEtc,
        [in] IAdviseSink *lpAdviseSink,
        [out] LONG *rgdwConnection,
        [in] IOleClientSite *pClientSite,
        [in] IStorage *pStg,
        [out] void *ppvObj);

    [entry("OleCreateLinkFromData")]
    HRESULT OleCreateLinkFromData(
        [in] IDataObject *pSrcDataObj,
        [in] UUID *riid,
        [in] OLERENDER renderopt,
        [in] void *pFormatEtc,
        [in] IOleClientSite *pClientSite,
        [in] IStorage *pStg,
        [out] void *ppvObj);

    [entry("OleCreateLinkFromDataEx")]
    HRESULT OleCreateLinkFromDataEx(
        [in] IDataObject *pSrcDataObj,
        [in] UUID *riid,
        [in] OLECREATE_UDT dwFlags,
        [in] OLERENDER renderopt,
        [in] LONG cFormats,
        [in] LONG* rgAdvf,
        [in] void *rgFormatEtc,
        [in] IAdviseSink *lpAdviseSink,
        [in, out] LONG *rgdwConnection,
        [in] IOleClientSite *pClientSite,
        [in] IStorage *pStg,
        [out] void *ppvObj);

    [entry("OleCreateStaticFromData")]
    HRESULT OleCreateStaticFromData(
        [in] IDataObject *pSrcDataObj,
        [in] UUID *iid,
        [in] OLERENDER renderopt,
        [in] void *pFormatEtc,
        [in] IOleClientSite *pClientSite,
        [in] IStorage *pStg,
        [out] void *ppvObj);

    [entry("OleCreateLink")]
    HRESULT OleCreateLink(
        [in] IMoniker *pmkLinkSrc,
        [in] UUID *riid,
        [in] OLERENDER renderopt,
        [in] void *rgFormatEtc,
        [in] IOleClientSite *pClientSite,
        [in] IStorage *pStg,
        [out] void *ppvObj);

    [entry("OleCreateLinkEx")]
    HRESULT OleCreateLinkEx(
        [in] IMoniker *pmkLinkSrc,
        [in] UUID *riid,
        [in] OLECREATE_UDT dwFlags,
        [in] OLERENDER renderopt,
        [in] LONG cFormats,
        [in] LONG* rgAdvf,
        [in] void *rgFormatEtc,
        [in] IAdviseSink *lpAdviseSink,
        [out] LONG *rgdwConnection,
        [in] IOleClientSite *pClientSite,
        [in] IStorage *pStg,
        [out] void *ppvObj);

    [entry("OleCreateLinkToFile")]
    HRESULT OleCreateLinkToFile(
        [in] LPWSTR lpszFileName,
        [in] UUID *riid,
        [in] OLERENDER renderopt,
        [in] void *rgFormatEtc,
        [in] IOleClientSite *pClientSite,
        [in] IStorage *pStg,
        [out] void *ppvObj);

    [entry("OleCreateLinkToFileEx")]
    HRESULT OleCreateLinkToFileEx(
        [in] LPWSTR lpszFileName,
        [in] UUID *riid,
        [in] OLECREATE_UDT dwFlags,
        [in] OLERENDER renderopt,
        [in] LONG cFormats,
        [in] LONG* rgAdvf,
        [in] void *rgFormatEtc,
        [in] IAdviseSink *lpAdviseSink,
        [out] LONG *rgdwConnection,
        [in] IOleClientSite *pClientSite,
        [in] IStorage *pStg,
        [out] void *ppvObj);

    [entry("OleCreateFromFile")]
    HRESULT OleCreateFromFile(
        [in] void *rclsid,
        [in] LPWSTR lpszFileName,
        [in] UUID *riid,
        [in] OLERENDER renderopt,
        [in] void *rgFormatEtc,
        [in] IOleClientSite *pClientSite,
        [in] IStorage *pStg,
        [out] void *ppvObj);

    [entry("OleCreateFromFileEx")]
    HRESULT OleCreateFromFileEx(
        [in] UUID *rclsid,
        [in] LPWSTR lpszFileName,
        [in] UUID *riid,
        [in] OLECREATE_UDT dwFlags,
        [in] OLERENDER renderopt,
        [in] LONG cFormats,
        [in] LONG* rgAdvf,
        [in] FORMATETC *rgFormatEtc,
        [in] IAdviseSink *lpAdviseSink,
        [out] LONG *rgdwConnection,
        [in] IOleClientSite *pClientSite,
        [in] IStorage *pStg,
        [out] void *ppvObj);

    [entry("OleLoad")]
    HRESULT OleLoad(
        [in] IStorage *pStg,
        [in] UUID *riid,
        [in] IOleClientSite *pClientSite,
        [out] void *ppvObj);

    [entry("OleSave")]
    HRESULT OleSave(
        [in] IPersistStorage *pPS,
        [in] IStorage *pStg,
        [in] BOOL fSameAsLoad);

    [entry("OleLoadFromStream")]
    HRESULT OleLoadFromStream(
        [in] IStream *pStm,
        [in] UUID *iidInterface,
        [out] void *ppvObj);

    [entry("OleSaveToStream")]
    HRESULT OleSaveToStream(
        [in] IPersistStream * pPStm,
        [in] IStream *pStm );

    [entry("OleNoteObjectVisible")]
    HRESULT OleNoteObjectVisible(
        [in] stdole.IUnknown *pUnknown,
        [in] BOOL fVisible);

/* Drag/Drop APIs */

	[entry("RegisterDragDrop")]
	HRESULT RegisterDragDrop(
		[in] LONG hwnd,
		[in] IDropTarget *pDropTarget);
		
	[entry("RevokeDragDrop")]
	HRESULT RevokeDragDrop(
		[in] LONG hwnd);
		
	[entry("DoDragDrop")]
	HRESULT DoDragDrop(
		[in] IDataObject *pDataObj,
		[in] IDropSource *pDropSource,
		[in] LONG dwOKEffects, 
		[in] LONG *pdwEffect);

/* Clipboard APIs */

    [entry("OleSetClipboard")]
    HRESULT OleSetClipboard(
        [in] IDataObject *pDataObj);

    [entry("OleGetClipboard")]
    HRESULT OleGetClipboard(
        [out, retval] IDataObject **ppDataObj);

    [entry("OleFlushClipboard")]
    HRESULT OleFlushClipboard(void);

    [entry("OleIsCurrentClipboard")]
    HRESULT OleIsCurrentClipboard(
        [in] IDataObject *pDataObj);

/* InPlace Editing APIs */

    [entry("OleCreateMenuDescriptor")]
    LONG OleCreateMenuDescriptor(
        [in] LONG hmenuCombined,
        [in] OLEMENUGROUPWIDTHS *lpMenuWidths);

    [entry("OleSetMenuDescriptor")]
    HRESULT OleSetMenuDescriptor(
        [in] LONG holemenu,
        [in] LONG hwndFrame,
        [in] LONG hwndActiveObject,
        [in] IOleInPlaceFrame *lpFrame,
        [in] IOleInPlaceActiveObject *lpActiveObj);

    [entry("OleDestroyMenuDescriptor")]
    HRESULT OleDestroyMenuDescriptor(
        [in] LONG holemenu);

    [entry("OleTranslateAccelerator")]
    HRESULT OleTranslateAccelerator(
        [in] IOleInPlaceFrame *lpFrame,
        [in] OLEINPLACEFRAMEINFO *lpFrameInfo,
        [in] MSG *lpmsg);

/* Helper APIs */

    [entry("OleDuplicateData")]
    LONG OleDuplicateData(
        [in] LONG hSrc,
        [in] short cfFormat,
        [in] LONG uiFlags);

    [entry("OleDraw")]
    HRESULT OleDraw(
        [in] stdole.IUnknown *pUnknown,
        [in] DVASPECT dwAspect,
        [in] LONG hdcDraw,
        [in] RECT *lprcBounds);

    [entry("OleRun")]
    HRESULT OleRun(
        [in] stdole.IUnknown *pUnknown);

    [entry("OleIsRunning")]
    LONG OleIsRunning(
        [in] IOleObject *pObject);

    [entry("OleLockRunning")]
    HRESULT OleLockRunning(
        [in] stdole.IUnknown *pUnknown,
        [in] BOOL fLock,
        [in] BOOL fLastUnlockCloses);

//    [entry("CreateOleAdviseHolder")]
//    HRESULT CreateOleAdviseHolder(
//        [out, retval] IOleAdviseHolder *ppOAHolder);

    [entry("CreateDataCache")]
    HRESULT CreateDataCache(
        [in] stdole.IUnknown *pUnkOuter,
        [in] UUID *rclsid,
	[in] UUID *iid, 
        [out, retval] stdole.IUnknown **ppv);


    [entry("OleCreateDefaultHandler")]
    HRESULT OleCreateDefaultHandler(
        [in] UUID *clsid,
        [in] stdole.IUnknown *pUnkOuter,
        [in] UUID *riid,
        [out] void *lplpObj);

//    HRESULT OleCreateEmbeddingHelper(
//          [in] UUID *clsid,
//          [in] stdole.IUnknown *pUnkOuter,
//          [in] LONG flags,
//          [in] IClassFactory *pCF,
//          [in] UUID *riid,
//          [out] void *lplpObj);

    [entry("IsAccelerator")]
    LONG IsAccelerator(
        [in] LONG hAccel,
        [in] LONG cAccelEntries,
        [in] MSG *lpMsg,
        [out] short *lpwCmd);


}
