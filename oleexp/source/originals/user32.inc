typedef [public] unsigned long ULONG;
typedef [public] char CHAR;

[
	dllname("USER32.DLL")
]
module user32 {

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