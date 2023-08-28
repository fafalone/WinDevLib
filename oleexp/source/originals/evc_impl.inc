    	[
		odl,
		uuid(6E793361-73C6-11D0-8469-00AA00442901)
	]
	interface IEmptyVolumeCacheCallBack : IUnknown {

		HRESULT ScanProgress( 
			[in] CURRENCY dwlSpaceUsed,
			[in] olelib.IEmptyVolumeCacheCallBackFlags dwFlags,
			[in] LPWSTR pcwszStatus);
        
		HRESULT PurgeProgress( 
			[in] CURRENCY dwlSpaceFreed,
			[in] CURRENCY dwlSpaceToFree,
			[in] olelib.IEmptyVolumeCacheCallBackFlags dwFlags,
			[in] LPWSTR pcwszStatus);
	};

	[
		odl,
		uuid(8FCE5227-04DA-11d1-A004-00805F8ABE06)
	]
	interface IEmptyVolumeCache : stdole.IUnknown {

		HRESULT Initialize( 
            		[in] LONG hkRegKey,
			[in] LONG pcwszVolume,
			[in, out] LONG *ppwszDisplayName,
			[in, out] LONG *ppwszDescription,
			[in, out] IEmptyVolumeCacheFlags *pdwFlags);
        
		HRESULT GetSpaceUsed( 
			[in, out] CURRENCY *pdwlSpaceUsed,
			[in] IEmptyVolumeCacheCallBack *picb);
        
		HRESULT Purge( 
			[in] CURRENCY dwlSpaceToFree,
			[in] IEmptyVolumeCacheCallBack *picb);
        
		HRESULT ShowProperties( 
			[in] LONG hwnd);
        
		HRESULT Deactivate( 
			[in, out] IEmptyVolumeCacheFlags *pdwFlags);
        
	};

	[
		odl,
		uuid(02b7e3ba-4db3-11d2-b2d9-00c04f8eec8c)
	]
	interface IEmptyVolumeCache2 : stdole.IUnknown {

		HRESULT Initialize( 
            		[in] LONG hkRegKey,
			[in] LONG pcwszVolume,
			[in, out] LONG *ppwszDisplayName,
			[in, out] LONG *ppwszDescription,
			[in, out] IEmptyVolumeCacheFlags *pdwFlags);
        
		HRESULT GetSpaceUsed( 
			[in, out] CURRENCY *pdwlSpaceUsed,
			[in] IEmptyVolumeCacheCallBack *picb);
        
		HRESULT Purge( 
			[in] CURRENCY dwlSpaceToFree,
			[in] IEmptyVolumeCacheCallBack *picb);
        
		HRESULT ShowProperties( 
			[in] LONG hwnd);
        
		HRESULT Deactivate( 
			[in, out] IEmptyVolumeCacheFlags *pdwFlags);
        
	        HRESULT InitializeEx( 
			[in] LONG hkRegKey,
			[in] LONG pcwszVolume,
			[in] LONG pcwszKeyName,
			[in, out] LONG *ppwszDisplayName,
			[in, out] LONG *ppwszDescription,
			[in, out] LONG *ppwszBtnText,
			[in, out] IEmptyVolumeCacheFlags *pdwFlags);
	};
