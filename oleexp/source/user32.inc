//tbShellLib - 100% implemented    


typedef enum
    {
        PM_NOREMOVE = 0,
        PM_REMOVE = 1,
        PM_NOYIELD = 2
    } PMOptions;

[
	dllname("USER32.DLL")
]
module user32 {
	//begin oleexp additions
    [entry("PeekMessageA")] BOOL PeekMessage([in] MSG* lpMsg, [in] HWND hWnd, [in] UINT wMsgFilterMin, [in] UINT wMsgFilterMax, [in] PMOptions wRemoveMsg);
    [entry("PeekMessageW")] BOOL PeekMessageW([in] MSG* lpMsg, [in] HWND hWnd, [in] UINT wMsgFilterMin, [in] UINT wMsgFilterMax, [in] PMOptions wRemoveMsg);
    [entry("TranslateMessage")] BOOL TranslateMessage([in] MSG* lpMsg);
    [entry("DispatchMessageA")] long DispatchMessage([in] MSG* lpMsg);
    [entry("DispatchMessageW")] long DispatchMessageW([in] MSG* lpMsg);
    [entry("EnumThreadWindows")] BOOL EnumThreadWindows([in] DWORD dwThreadId, [in] long lpfn, [in] long lParam);
    [entry("GetClassNameA")] long GetClassNameA([in] HWND hWnd, [in] LPSTR lpClassName, [in]long nMaxCount);
    [entry("GetClassNameW")] long GetClassNameW([in] HWND hWnd, [in] LPWSTR lpClassName, [in]long nMaxCount);
    [entry("SetTimer")] UINT SetTimer([in] HWND hWnd, [in] UINT nIDEvent, [in] UINT uEleapse, [in] long lpTimerFunc);
    [entry("KillTimer")] BOOL KillTimer([in] HWND hWnd, [in] UINT nIDEvent);

	[entry("SystemParametersInfo")]
	LONG SystemParametersInfoA(
		[in] SPI uiAction,
		[in] LONG uiParam,
		[in,out] LPVOID pvParam,
		[in] SPIF fWinIni);

	[entry("GetWindowLongW")]
	LONG GetWindowLong(
		[in] LONG hWnd,
		[in] GWL_INDEX nIndex);

	[entry("SetWindowLongW")]
	LONG SetWindowLong(
		[in] LONG hWnd,
		[in] GWL_INDEX nIndex,
		[in] LONG dwNewLong);
	
	[entry("GetSystemMetrics")]
	LONG GetSystemMetrics(
		[in] SystemMetrics nIndex);

	//End oleexp additions

	[entry("RegisterClipboardFormatA")]
	LONG RegisterClipboardFormat(
		[in] LPSTR lpszFormat);

	[entry("SendMessageA")]
	LONG SendMessage(
                [in] LONG hWnd,
                [in] LONG wMsg,
                [in] LONG wParam,
                [in] void *lParam);

	[entry("SendMessageW")]
	LONG SendMessageW(
                [in] LONG hWnd,
                [in] LONG wMsg,
                [in] LONG wParam,
                [in] void *lParam);

	[entry("PostMessageA")]
	LONG PostMessage(
				[in] LONG hWnd,
				[in] LONG wMsg,
				[in] LONG wParam,
				[in] void *lParam);
				
	[entry("PostMessageW")]
	LONG PostMessageW(
				[in] LONG hWnd,
				[in] LONG wMsg,
				[in] LONG wParam,
				[in] void *lParam);

	[entry("FindWindowA")]
	LONG FindWindow(
                [in] LPSTR Class,
                [in] LPSTR Title);

	[entry("FindWindowW")]
	LONG FindWindowW(
                [in] LPWSTR Class,
                [in] LPWSTR Title);

}