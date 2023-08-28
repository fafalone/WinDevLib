//tbShellLib - 100% implemented

typedef enum IEmptyVolumeCacheFlags {
	EVCF_HASSETTINGS             =0x0001,
	EVCF_ENABLEBYDEFAULT         =0x0002,
	EVCF_REMOVEFROMLIST          =0x0004,
	EVCF_ENABLEBYDEFAULT_AUTO    =0x0008,
	EVCF_DONTSHOWIFZERO          =0x0010,
	EVCF_SETTINGSMODE            =0x0020,
	EVCF_OUTOFDISKSPACE          =0x0040,
} IEmptyVolumeCacheFlags;

typedef enum IEmptyVolumeCacheCallBackFlags {
	EVCCBF_LASTNOTIFICATION  = 0x0001
} IEmptyVolumeCacheCallBackFlags;

[
	odl,
	uuid(6E793361-73C6-11D0-8469-00AA00442901)
]
interface IEmptyVolumeCacheCallBack : stdole.IUnknown {

	HRESULT ScanProgress( 
		[in] CURRENCY dwlSpaceUsed,
		[in] IEmptyVolumeCacheCallBackFlags dwFlags,
		[in] LONG pcwszStatus);
        
	HRESULT PurgeProgress( 
		[in] CURRENCY dwlSpaceFreed,
		[in] CURRENCY dwlSpaceToFree,
		[in] IEmptyVolumeCacheCallBackFlags dwFlags,
		[in] LONG pcwszStatus);
};

[
	odl,
	uuid(8FCE5227-04DA-11d1-A004-00805F8ABE06)
]
interface IEmptyVolumeCache : stdole.IUnknown {

	HRESULT Initialize( 
		[in] LONG hkRegKey,
		[in] LPWSTR pcwszVolume,
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
interface IEmptyVolumeCache2 : IEmptyVolumeCache {

	HRESULT InitializeEx( 
		[in] LONG hkRegKey,
		[in] LPWSTR  pcwszVolume,
		[in] LPWSTR  pcwszKeyName,
		[in, out] LPWSTR  *ppwszDisplayName,
		[in, out] LPWSTR  *ppwszDescription,
		[in, out] LPWSTR  *ppwszBtnText,
		[in, out] IEmptyVolumeCacheFlags *pdwFlags);
};
