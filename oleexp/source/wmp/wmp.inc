[
    odl,
    uuid(CBB92747-741F-44fe-AB5B-F1A48F3B2A59),
]
interface IWMPRemoteMediaServices : stdole.IUnknown {

    HRESULT GetServiceType(
        [in, out] BSTR * pbstrType);

    HRESULT GetApplicationName(
        [in, out] BSTR * pbstrName);

    HRESULT GetScriptableObject(
        [in, out] BSTR * pbstrName,
        [in, out] IDispatch ** ppDispatch);

    HRESULT GetCustomUIMode(
        [in, out] BSTR * pbstrFile);

}

// The following variables should be written to the following key:
//      HLKM\PLUGIN_INSTALLREGKEY\{GUID}
// Where {GUID} is the identifier used to create the plugin, and PLUGIN_INSTALLREGKEY is the
// registry path listed below.

[
    dllname("wmp.dll")
]
module wmp {

	const LPSTR PLUGIN_INSTALLREGKEY = "Software\\Microsoft\\MediaPlayer\\UIPlugins";
	const LPSTR PLUGIN_INSTALLREGKEY_FRIENDLYNAME = "FriendlyName";
	const LPSTR PLUGIN_INSTALLREGKEY_DESCRIPTION  = "Description";
	const LPSTR PLUGIN_INSTALLREGKEY_CAPABILITIES = "Capabilities";
	const LPSTR PLUGIN_INSTALLREGKEY_UNINSTALL    = "UninstallPath";

	// Plugin Types (stored in the PLUGIN_INSTALLREGKEY_CAPABILITIES registry entry)

	const LONG PLUGIN_TYPE_BACKGROUND          =   0x00000001;         // plugin has no UI
	const LONG PLUGIN_TYPE_SEPARATEWINDOW      =   0x00000002;         // plugin runs in its own window
	const LONG PLUGIN_TYPE_DISPLAYAREA         =   0x00000003;         // plugin takes the full visuzalization area
	const LONG PLUGIN_TYPE_SETTINGSAREA        =   0x00000004;         // plugin takes the settings area
	const LONG PLUGIN_TYPE_METADATAAREA        =   0x00000005;         // plugin takes the metadata area

	// Capability Flags (also stored in the PLUGIN_INSTALLREGKEY_CAPABILITIES registry entry)

	const LONG PLUGIN_FLAGS_HASPROPERTYPAGE    =   0x80000000;         // does the plugin have a property page?
	const LONG PLUGIN_FLAGS_INSTALLAUTORUN     =   0x40000000;         // autorun after installation?
	const LONG PLUGIN_FLAGS_LAUNCHPROPERTYPAGE =   0x20000000;         // auto-launch property page on first load?
	const LONG PLUGIN_FLAGS_ACCEPTSMEDIA       =   0x10000000;         // allow media 'send-to' context item?
	const LONG PLUGIN_FLAGS_ACCEPTSPLAYLISTS   =   0x08000000;         // allow playlist 'send-to' context item?
	const LONG PLUGIN_FLAGS_HASPRESETS         =   0x04000000;         // does the plugin have presets? (Display, Settings, Metadata only)
	const LONG PLUGIN_FLAGS_HIDDEN             =   0x02000000;         // should the plugin not be listed in menus?(Background plugins only)

	// Variables requested from the GetProperty API for plugins that have PLUGIN_FLAGS_HASPRESETS flag set

	const LPSTR PLUGIN_MISC_PRESETCOUNT      = "PresetCount"; // (VT_I4)      returns total number of presets
	const LPSTR PLUGIN_MISC_PRESETNAMES      = "PresetNames"; // (VT_ARRAY)   returns BSTR array of preset titles

	// Variables requested from the GetProperty/SetProperty API for plugins that have PLUGIN_FLAGS_HASPRESETS flag set

	const LPSTR PLUGIN_MISC_CURRENTPRESET    = "CurrentPreset"; // (VT_I4)      returns/sets the current preset

	// Variables requested from the GetProperty API for PLUGIN_TYPE_SEPARATEWINDOW plugins

	const LPSTR PLUGIN_SEPARATEWINDOW_RESIZABLE     = "Resizable";       // (VT_BOOL)    when true, makes the window resizable
	const LPSTR PLUGIN_SEPARATEWINDOW_DEFAULTWIDTH  = "DefaultWidth";    // (VT_I4)      the plugin's initial launch width
	const LPSTR PLUGIN_SEPARATEWINDOW_DEFAULTHEIGHT = "DefaultHeight";   // (VT_I4)      the plugin's initial launch height
	const LPSTR PLUGIN_SEPARATEWINDOW_MINWIDTH      = "MinWidth";        // (VT_I4)      the plugin's minimum resize width
	const LPSTR PLUGIN_SEPARATEWINDOW_MINHEIGHT     = "MinHeight";       // (VT_I4)      the plugin's minimum resize height
	const LPSTR PLUGIN_SEPARATEWINDOW_MAXWIDTH      = "MaxWidth";        // (VT_I4)      the plugin's maximum resize width
	const LPSTR PLUGIN_SEPARATEWINDOW_MAXHEIGHT     = "MaxHeight";       // (VT_I4)      the plugin's maximum resize height

	// Variables requested from the GetProperty API for PLUGIN_TYPE_SEPARATEWINDOW and PLUGIN_TYPE_BACKGROUND plugins

	const LPSTR PLUGIN_MISC_QUERYDESTROY            = "QueryDestroy";    // (VT_BSTR)    when not blank, displays returned string and doesn't close plugin

	// Variables given to the SetProperty API for any plugin type exposing PLUGIN_FLAGS_ACCEPTSMEDIA or PLUGIN_FLAGS_ACCEPTSPLAYLISTS

	const LPSTR PLUGIN_ALL_MEDIASENDTO              = "MediaSendTo";     // (VT_ARRAY)   array of IUnknown pointers (QI for IWMPMedia)
	const LPSTR PLUGIN_ALL_PLAYLISTSENDTO           = "PlaylistSendTo";  // (VT_ARRAY)   array of IUnknown pointers (QI for IWMPPlaylist)

	// Notification function that plugin installed/uninstalled
    /*
	cpp_quote( "__inline BOOL WMPNotifyPluginAddRemove()" )
	cpp_quote( "{" )
	cpp_quote( "    return( ::PostMessage( HWND_BROADCAST, ::RegisterWindowMessageA( \"WMPlayer_PluginAddRemove\" ), 0, 0 ) );" )
	cpp_quote( "}" )
    */

}

[
    odl,
    uuid(4C5E8F9F-AD3E-4bf9-9753-FCD30D6D38DD),
]
interface IWMPPluginUI : stdole.IUnknown {

    HRESULT SetCore(
		[in] IWMPCore *pCore);
	
    HRESULT Create(
		[in] LONG hwndParent,
		[out, retval] LONG *phwndWindow);
	
    HRESULT Destroy();

    HRESULT DisplayPropertyPage(
		[in] LONG hwndParent);

    HRESULT GetProperty(
		[in] LONG pwszName,
		[in, out] VARIANT *pvarProperty);
	
    HRESULT SetProperty(
		[in] LONG pwszName,
		[in, out] VARIANT *pvarProperty);
	
    HRESULT TranslateAccelerator(
		[in, out] MSG *lpmsg);
}
