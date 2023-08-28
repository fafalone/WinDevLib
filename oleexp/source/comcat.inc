//tbShellLib - 100% implemented	

[
	    odl,
	    uuid(0002E012-0000-0000-C000-000000000046)
	]
	interface ICatRegister : stdole.IUnknown{

		HRESULT RegisterCategories(
			[in] LONG cCategories,
			[in] CATEGORYINFO *rgCategoryInfo);

		HRESULT UnRegisterCategories(
			[in] LONG cCategories,
			[in] UUID *rgcatid);

		HRESULT RegisterClassImplCategories(
			[in] UUID *rclsid,
			[in] LONG cCategories,
			[in] UUID *rgcatid);

		HRESULT UnRegisterClassImplCategories(
			[in] UUID *rclsid,
			[in] LONG cCategories,
			[in] UUID *rgcatid);

		HRESULT RegisterClassReqCategories(
			[in] UUID *rclsid,
			[in] LONG cCategories,
			[in] UUID *rgcatid);

		HRESULT UnRegisterClassReqCategories(
			[in] UUID *rclsid,
			[in] LONG cCategories,
			[in] UUID *rgcatid);
	}

	[
		odl,
		uuid(0002E013-0000-0000-C000-000000000046)
	]
	interface ICatInformation : stdole.IUnknown{

		HRESULT EnumCategories(
			[in] LONG lcid,
			[out, retval] IEnumCATEGORYINFO** ppenumCategoryInfo);

		HRESULT GetCategoryDesc(
			[in] UUID *rcatid,
			[in] LONG lcid,
			[out, retval] LONG *pszDesc);

		HRESULT EnumClassesOfCategories(
			[in] LONG cImplemented,
			[in] UUID *rgcatidImpl,
			[in] LONG cRequired,
			[in] UUID *rgcatidReq,
			[out, retval] IEnumGUID** ppenumClsid);

		HRESULT IsClassOfCategories(
			[in] UUID *rclsid,
			[in] LONG cImplemented,
			[in] UUID *rgcatidImpl,
			[in] LONG cRequired,
			[in] UUID *rgcatidReq);

		HRESULT EnumImplCategoriesOfClass(
			[in] UUID *rclsid,
			[out, retval] IEnumGUID** ppenumCatid);

		HRESULT EnumReqCategoriesOfClass(
			[in] UUID *rclsid,
			[out, retval] IEnumGUID** ppenumCatid);
	}

	[
		uuid(0002e005-0000-0000-c000-000000000046),
		helpstring("Standard Categories Manager")
	]
	coclass StdComponentCategoriesMgr {
		[default] interface ICatInformation;
		interface ICatRegister;
	}
