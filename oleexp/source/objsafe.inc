//tbShellLib - 100% implemented

typedef [helpstring("Option bit definitions for IObjectSafety")]
enum OBJSAFE_Flags {

	[helpstring("Caller of interface may be untrusted")]
	INTERFACESAFE_FOR_UNTRUSTED_CALLER	=	0x00000001,

	[helpstring("Data passed into interface may be untrusted")]
	INTERFACESAFE_FOR_UNTRUSTED_DATA	=	0x00000002,

	[helpstring("Object knows to use IDispatchEx")]
	INTERFACE_USES_DISPEX			=	0x00000004,

	[helpstring("Object knows to use IInternetHostSecurityManager")]
	INTERFACE_USES_SECURITY_MANAGER		=	0x00000008,
} OBJSAFE_Flags;

[
	odl,
	uuid(CB5BDC81-93C1-11cf-8F20-00805F2CD064)	
]
interface IObjectSafety : stdole.IUnknown {

	HRESULT GetInterfaceSafetyOptions(
		[in] UUID * riid,				// Interface that we want options for
		[in, out] OBJSAFE_Flags *pdwSupportedOptions,	// Options meaningful on this interface
		[in, out] OBJSAFE_Flags *pdwEnabledOptions);	// current option values on this interface

	HRESULT SetInterfaceSafetyOptions(
		[in]  UUID * riid,			// Interface to set options for
		[in]  OBJSAFE_Flags dwOptionSetMask,	// Options to change
		[in]  OBJSAFE_Flags dwEnabledOptions);	// New option values
}
