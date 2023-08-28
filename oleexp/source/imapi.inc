//tbShellLib - 100% complete

//IMAPI

typedef enum MEDIA_TYPES {
	MEDIA_CDDA_CDROM	= 1,
	MEDIA_CD_ROM_XA		= 2,
	MEDIA_CD_I			= 3,
	MEDIA_CD_EXTRA		= 4,
	MEDIA_CD_OTHER		= 5,
	MEDIA_SPECIAL		= 6
} MEDIA_TYPES;

typedef enum MEDIA_FLAGS {
	MEDIA_BLANK			= 0x1,
	MEDIA_RW			= 0x2,
	MEDIA_WRITABLE		= 0x4,
	MEDIA_FORMAT_UNUSABLE_BY_IMAPI	= 0x8
} MEDIA_FLAGS;

typedef enum RECORDER_TYPES {
	RECORDER_CDR	= 0x1,
	RECORDER_CDRW	= 0x2
} RECORDER_TYPES;

typedef enum RECORDER_STATE {
	RECORDER_DOING_NOTHING,
	RECORDER_OPENED	= 1,
	RECORDER_BURNING = 2
} RECORDER_STATE;

[
	odl,
	uuid(85AC9776-CA88-4cf2-894E-09598C078A41)
]
interface IDiscRecorder : stdole.IUnknown {

	HRESULT Init( 
		[in] byte *pbyUniqueID,
		[in] LONG nulIDSize,
		[in] LONG nulDriveNumber);
        
	HRESULT GetRecorderGUID( 
		[in, out] byte *pbyUniqueID,
		[in] LONG ulBufferSize,
		[out, retval] LONG *pulReturnSizeRequired);
        
	HRESULT GetRecorderType( 
		[out, retval] RECORDER_TYPES *fTypeCode);
        
	HRESULT GetDisplayNames( 
		[in, out] BSTR *pbstrVendorID,
		[in, out] BSTR *pbstrProductID,
		[in, out] BSTR *pbstrRevision);
        
	HRESULT GetBasePnPID( 
		[out, retval] BSTR *pbstrBasePnPID);
        
	HRESULT GetPath( 
		[out, retval] BSTR *pbstrPath);
        
	HRESULT GetRecorderProperties( 
		[out, retval] IPropertyStorage **ppPropStg);
        
	HRESULT SetRecorderProperties( 
		[in] IPropertyStorage *pPropStg);
        
	HRESULT GetRecorderState( 
		[out, retval] RECORDER_STATE *pulDevStateFlags);
        
	HRESULT OpenExclusive();
        
	HRESULT QueryMediaType( 
		[in, out] MEDIA_TYPES *fMediaType,
		[in, out] MEDIA_FLAGS *fMediaFlags);
        
	HRESULT QueryMediaInfo( 
		[in, out] byte *pbSessions,
		[in, out] byte *pbLastTrack,
		[in, out] LONG *ulStartAddress,
		[in, out] LONG *ulNextWritable,
		[in, out] LONG *ulFreeBlocks);
        
	HRESULT Eject();
        
	HRESULT Erase( 
		[in] boolean bFullErase);
        
	HRESULT Close();
        
};

[
	odl,
	uuid(9B1921E1-54AC-11d3-9144-00104BA11C5E)
]
interface IEnumDiscRecorders : stdole.IUnknown {

	LONG Next( 
		[in] LONG cRecorders,
		[in, out] IDiscRecorder **ppRecorder,
		[in, out, defaultvalue(0)] LONG *pcFetched);
        
	HRESULT Skip( 
		[in] LONG cRecorders);
        
	HRESULT Reset();
        
	HRESULT Clone( 
		[out, retval] IEnumDiscRecorders **ppEnum);
        
};

[
	odl,    
    uuid(DDF445E1-54BA-11d3-9144-00104BA11C5E)
]
interface IEnumDiscMasterFormats : stdole.IUnknown {

	LONG Next( 
		[in] LONG cFormats,
		[in, out]  UUID *lpiidFormatID,
		[in, out, defaultvalue(0)] LONG *pcFetched);
        
	HRESULT Skip( 
		[in] LONG cFormats);
        
	HRESULT Reset();
        
	HRESULT Clone( 
		[out, retval] IEnumDiscMasterFormats **ppEnum);
        
};

[
	odl,
	uuid(E3BC42CD-4E5C-11D3-9144-00104BA11C5E)
]
interface IRedbookDiscMaster : stdole.IUnknown {

	HRESULT GetTotalAudioTracks( 
		[out, retval] long *pnTracks);
        
	HRESULT GetTotalAudioBlocks( 
		[out, retval] long *pnBlocks);
        
	HRESULT GetUsedAudioBlocks( 
		[out, retval]  long *pnBlocks);
        
	HRESULT GetAvailableAudioTrackBlocks( 
		[out, retval]  long *pnBlocks);
        
	HRESULT GetAudioBlockSize( 
		[out, retval]  long *pnBlockBytes);
        
	HRESULT CreateAudioTrack( 
		[in] long nBlocks);
        
	HRESULT AddAudioTrackBlocks( 
		[in, out] byte *pby,
		[in] long cb);
        
	HRESULT CloseAudioTrack();
        
};

[
	odl,
	uuid(E3BC42CE-4E5C-11D3-9144-00104BA11C5E)
]
interface IJolietDiscMaster : stdole.IUnknown {

	HRESULT GetTotalDataBlocks( 
		[out, retval] long *pnBlocks);
        
	HRESULT GetUsedDataBlocks( 
		[out, retval] long *pnBlocks);
        
	HRESULT GetDataBlockSize( 
		[out, retval] long *pnBlockBytes);
        
	HRESULT AddData( 
		[in] IStorage *pStorage,
		[in] long lFileOverwrite);
        
	HRESULT GetJolietProperties( 
		[out, retval] IPropertyStorage **ppPropStg);
        
	HRESULT SetJolietProperties( 
		[in] IPropertyStorage *pPropStg);
        
};

[
	odl,
	uuid(EC9E51C1-4E5D-11D3-9144-00104BA11C5E)
]
interface IDiscMasterProgressEvents : stdole.IUnknown {

	HRESULT QueryCancel( 
		[out, retval] boolean *pbCancel);
        
	HRESULT NotifyPnPActivity();
        
	HRESULT NotifyAddProgress( 
		[in] long nCompletedSteps,
		[in] long nTotalSteps);
        
	HRESULT NotifyBlockProgress( 
		[in] long nCompleted,
		[in] long nTotal);
        
	HRESULT NotifyTrackProgress( 
		[in] long nCurrentTrack,
		[in] long nTotalTracks);
        
	HRESULT NotifyPreparingBurn( 
		[in] long nEstimatedSeconds);
        
	HRESULT NotifyClosingDisc( 
		[in] long nEstimatedSeconds);
        
	HRESULT NotifyBurnComplete( 
        [in] long status);
        
	HRESULT NotifyEraseComplete( 
		[in] long status);
        
};
    
[
	odl,   
	uuid(520CCA62-51A5-11D3-9144-00104BA11C5E)
]
interface IDiscMaster : stdole.IUnknown {

	HRESULT Open();
        
	HRESULT EnumDiscMasterFormats( 
		[out, retval] IEnumDiscMasterFormats **ppEnum);
        
	HRESULT GetActiveDiscMasterFormat( 
		[in, out] UUID *lpiid);
        
	HRESULT SetActiveDiscMasterFormat( 
		[in, out]  UUID *riid,
        [out, retval] stdole.IUnknown **ppUnk);
        
	HRESULT EnumDiscRecorders( 
		[out, retval] IEnumDiscRecorders **ppEnum);
        
	HRESULT GetActiveDiscRecorder( 
		[out, retval] IDiscRecorder **ppRecorder);
        
	HRESULT SetActiveDiscRecorder( 
		[in] IDiscRecorder *pRecorder);
        
	HRESULT ClearFormatContent();
        
	HRESULT ProgressAdvise( 
		[in] IDiscMasterProgressEvents *pEvents,
		[out, retval] LONG *pvCookie);
        
	HRESULT ProgressUnadvise( 
		[in] LONG vCookie);
        
	HRESULT RecordDisc( 
		[in] boolean bSimulate,
		[in] boolean bEjectAfterBurn);
        
	HRESULT Close();
        
};

[
	uuid(520CCA61-51A5-11D3-9144-00104BA11C5E)
]
coclass MSDiscRecorderObj {
	interface IDiscRecorder;
}

[
	uuid(520CCA63-51A5-11D3-9144-00104BA11C5E)
]
coclass MSDiscMasterObj {
	interface IDiscMaster;
}

[
	uuid(8A03567A-63CB-4BA8-BAF6-52119816D1EF)
]
coclass MSEnumDiscRecordersObj {
	interface IEnumDiscRecorders;
}