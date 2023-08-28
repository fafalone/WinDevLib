//tbShellLib - 100% implemented

typedef struct MEMORYSTATUSEX {
    LONG dwLength;
    LONG dwMemoryLoad;
    CURRENCY ullTotalPhys;
    CURRENCY ullAvailPhys;
    CURRENCY ullTotalPageFile;
    CURRENCY ullAvailPageFile;
    CURRENCY ullTotalVirtual;
    CURRENCY ullAvailVirtual;
    CURRENCY ullAvailExtendedVirtual;
} MEMORYSTATUSEX;

typedef struct MEMORYSTATUS { 
    LONG dwLength; 
    LONG dwMemoryLoad; 
    LONG dwTotalPhys; 
    LONG dwAvailPhys; 
    LONG dwTotalPageFile; 
    LONG dwAvailPageFile; 
    LONG dwTotalVirtual; 
    LONG dwAvailVirtual; 
} MEMORYSTATUS;

/* Global Memory Flags */
typedef enum GMEM {
     GMEM_FIXED          = 0x0000,
     GMEM_MOVEABLE       = 0x0002,
     GMEM_NOCOMPACT      = 0x0010,
     GMEM_NODISCARD      = 0x0020,
     GMEM_ZEROINIT       = 0x0040,
     GMEM_MODIFY         = 0x0080,
     GMEM_DISCARDABLE    = 0x0100,
     GMEM_NOT_BANKED     = 0x1000,
     GMEM_SHARE          = 0x2000,
     GMEM_DDESHARE       = 0x2000,
     GMEM_NOTIFY         = 0x4000,
     GMEM_LOWER          = 0x1000,
     GMEM_VALID_FLAGS    = 0x7F72,
     GMEM_INVALID_HANDLE = 0x8000,
     GHND = 0x42,
     GPTR = 0x40,

     /* Flags returned by GlobalFlags (in addition to GMEM_DISCARDABLE) */
     GMEM_DISCARDED      = 0x4000,
     GMEM_LOCKCOUNT      = 0x00FF,
} GMEM;

    typedef struct 
    {
      long DebugInfo;
      long LockCount;
      long RecursionCount;
      long OwningThread;
      long LockSemaphore;
      long SpinCount;
    } CRITICAL_SECTION;
    typedef [public] long LPCRITICAL_SECTION;

[
    dllname("KERNEL32.DLL")
]
module kernel32 {

	//oleexp additions
	[entry("TerminateThread")]
	long _stdcall TerminateThread([in, out] HANDLE *hThread, [in] DWORD dwExitCode);
      const long CREATE_SUSPENDED = 0x00000004;
      const long INFINITE = 0xFFFFFFFF;
      const long DUPLICATE_SAME_ACCESS = 0x00000002;
      const long STILL_ACTIVE = 0x00000103;
      [entry("WaitForSingleObject")] long WaitForSingleObject ([in] HANDLE hHandle, [in] long dwMilliseconds);
      [entry("CreateEventA"), usesgetlasterror] HANDLE CreateEvent ([in] long lpEventAttributes, [in] long bManualReset, [in] long bInitialState, [in] LPSTR lpName);
      [entry("CloseHandle")] long CloseHandle ([in] HANDLE hOjbect);
      [entry("PulseEvent")] long PulseEvent ([in] HANDLE hEvent);
      [entry("SetEvent")] long SetEvent ([in] HANDLE hEvent);
      [entry("ResetEvent")] long ResetEvent ([in] HANDLE hEvent);
      [entry("ResumeThread")] long ResumeThread ([in] HANDLE hThread);
      [entry("SuspendThread")] long SuspendThread ([in] HANDLE hThread);
      [entry("InterlockedIncrement")] long InterlockedIncrement([in] long pLong);
      [entry("InterlockedDecrement")] long InterlockedDecrement([in] long pLong);
      [entry("WaitForMultipleObjects")] long WaitForMultipleObjects ([in] long nCount, [in,out] HANDLE* lpHandles, [in] BOOL bWaitAll, [in] long dwMilliseconds);
      [entry("GetCurrentThreadId")] long GetCurrentThreadId ();
      [entry("GetCurrentProcessId")] long GetCurrentProcessId ();
      [entry("GetCurrentProcess")] HANDLE GetCurrentProcess();
      [entry("InitializeCriticalSection")] void InitializeCriticalSection([in] LPCRITICAL_SECTION lpCriticalSection);
      [entry("DeleteCriticalSection")] void DeleteCriticalSection([in] LPCRITICAL_SECTION lpCriticalSection);
      [entry("EnterCriticalSection")] void EnterCriticalSection([in] LPCRITICAL_SECTION lpCriticalSection);
      [entry("LeaveCriticalSection")] void LeaveCriticalSection([in] LPCRITICAL_SECTION lpCriticalSection);
      [entry("DuplicateHandle")] BOOL DuplicateHandle([in] HANDLE hSourceProcessHandle, [in] HANDLE hSourceHandle, [in] HANDLE hTargetProcessHandle, [out] HANDLE* lpTargetHandle, [in,defaultvalue(0)] DWORD dwDesiredAccess, [in, defaultvalue(0)] BOOL bInheritHandle, [in,defaultvalue(2)] DWORD dwOptions);
      [entry("GetExitCodeThread")]BOOL GetExitCodeThread([in] HANDLE hThread, [out] DWORD *lpExitCode);


        const int PAGE_EXECUTE              = 0x10;
        const int PAGE_EXECUTE_READ         = 0x20;
        const int PAGE_EXECUTE_READWRITE    = 0x40;
        const int PAGE_EXECUTE_WRITECOPY    = 0x80;
        const int PAGE_NOACCESS             = 0x1;
        const int PAGE_READONLY             = 0x2;
        const int PAGE_READWRITE            = 0x4;
        const int PAGE_WRITECOPY            = 0x8;

        const int MEM_COMMIT                = 0x1000;
        const int MEM_RESERVE               = 0x2000;

        [entry("RtlMoveMemory")]void CopyMemory([out] void* Dst,
                                                [in] void* Src,
                                                [in] int Length);

        [entry("CreateThread")] int CreateThread([in] void* lpThreadAttributes,
                                                 [in] int dwStackSize,
                                                 [in] int lpStartAddress,
                                                 [in] void* lpParameter,
                                                 [in] int dwCreationFlags,
                                                 [out] int* lpThreadId);

        [entry("HeapAlloc")] int HeapAlloc([in] int hHeap,
                                           [in] int dwFlags,
                                           [in] int dwBytes);

        [entry("HeapFree")] int HeapFree([in] int hHeap,
                                         [in] int dwFlags,
                                         [in] void* lpMem);

        [entry("GetProcessHeap")]int GetProcessHeap();

        [entry("VirtualProtect")] int VirtualProtect([in] void* lpAddress,
                                                     [in] int dwSize,
                                                     [in] int flNewProtect,
                                                     [in] int* lpflOldProtect);

        [entry("VirtualAlloc")]int VirtualAlloc([in] void* lpAddress,
                                                [in] int dwSize,
                                                [in] int flAllocationType,
                                                [in] int flProtect);

        [entry("TlsAlloc")]int TlsAlloc();

        [entry("TlsSetValue")]int TlsSetValue([in] int dwTlsIndex,
                                              [in] void* lpTlsValue);

        [entry("TlsGetValue")]int TlsGetValue([in] int dwTlsIndex);

        [entry("TlsFree")]int TlsFree([in] int dwTlsIndex);

        [entry("GetModuleHandleW")]int GetModuleHandle([in] void* lpModuleName);

        [entry("GetProcAddress")]int GetProcAddress([in] int hModule,
                                                    [in] lpStr lpProcName);



	[entry("GetLastError")]
	long GetLastError(void);

	[entry("Sleep")]
	void Sleep([in] LONG dwMilliseconds);

	[entry("GetShortPathNameW")]
	LONG GetShortPathNameW(
		[in] LONG lpszLongPath,
		[in] LONG lpszShortPath,
		[in] LONG cchBuffer);
		
	[entry("EncryptFileW")]
	BOOL EncryptFile(
		[in] LONG lpFileName);
		
	[entry("DecryptFileW")]
	BOOL DecryptFile(
		[in] LONG lpFileName,
		[in] LONG dwReserved);
	
	[entry("FindFirstFileExW")]
	int FindFirstFileEx(
		[in] LONG lpFileName,
		[in] FINDEX_INFO_LEVELS fInfoLevelId,
		[out] WIN32_FIND_DATAW *lpFindFileData,
		[in] FINDEX_SEARCH_OPS fSearchOp,
		[in] LONG lpSearchFilter,
		[in] FINDEX_EXTRA_FLAGS dwAdditionalFlags);

	[entry("FindFirstFileW")]
	int FindFirstFile(
		[in] LONG lpFileName,
		[out] WIN32_FIND_DATAW *lpFindFileData);

	[entry("FindFirstStreamW")]
	int FindFirstStreamW(
		[in] LONG                     lpFileName,
		[in] STREAM_INFO_LEVELS       InfoLevel,
		[out] WIN32_FIND_STREAM_DATA *lpFindStreamData,
		[in] DWORD                    dwFlags);

	[entry("FindNextFileW")]
	BOOL FindNextFile(
		[in] LONG hFindFile,
		[out] WIN32_FIND_DATAW *lpFindFileData);

	[entry("FindNextStreamW")]
	BOOL FindNextStreamW(
		[in] LONG hFindStream,
		[out] WIN32_FIND_STREAM_DATA *lpFindStreamData);

	[entry("FindClose")]
	BOOL FindClose(
		[in] LONG hFindFile);

	[entry("SetFileAttributesW")]
	BOOL SetFileAttributes(
		[in] LONG lpFileName,
		[in] FILE_ATTRIBUTES dwFileAttributes);

	[entry("GetFileAttributesW")]
	FILE_ATTRIBUTES GetFileAttributes(
		[in] LONG lpFileName);

	[entry("GetFileAttributesExW")]
	BOOL GetFileAttributesEx(
		[in] LONG lpFileName,
		[in] GET_FILEEX_INFO_LEVELS fInfoLevelId,
		[out] LPVOID lpFileInformation);

	[entry("GetFileInformationByHandle")]
	BOOL GetFileInformationByHandle(
		[in] LONG hFile,
		[out] BY_HANDLE_FILE_INFORMATION *lpFileInformation);

	
	[entry("GetFileInformationByHandleEx")]
	BOOL GetFileInformationByHandleEx(
		[in] LONG hFile,
		[in] FILE_INFO_BY_HANDLE_CLASS FileInformationClass,
		[in] LONG lpFileInformation,
		[in] DWORD dwBufferSize);




	
    [entry("lstrlenW")]
    LONG lstrlenW (
        [in] LONG StringPtr);

    [entry("lstrlenA")]
    LONG lstrlenA (
        [in] LONG StringPtr);

    [entry("lstrcpyA")]
    LONG lstrcpyA (
        [in] LPSTR Dest,
        [in] LONG Src);

    [entry("FileTimeToLocalFileTime")]
    BOOL FileTimeToLocalFileTime(
        [in] CURRENCY *lpFileTime,
        [in, out] CURRENCY *lpLocalFileTime);

    [entry("LocalFileTimeToFileTime")]
    BOOL LocalFileTimeToFileTime(
        [in] CURRENCY *lpFileTime,
        [in, out] CURRENCY *lpLocalFileTime);

    [entry("FileTimeToSystemTime")]
    BOOL FileTimeToSystemTime(
        [in] CURRENCY *lpFileTime,
        [in, out] SYSTEMTIME *lpSystemTime);

    [entry("SystemTimeToFileTime")]
    BOOL SystemTimeToFileTime(
            [in] SYSTEMTIME *lpSystemTime,
            [in, out] CURRENCY *lpFileTime);

    [entry("RtlMoveMemory")]
    void MoveMemory(
        [in] void *pDest,
        [in] void *pSource,
        [in] LONG ByteLen);

    [entry("RtlMoveMemory")]
    void MoveMemoryToLPSTR(
        [in] LPSTR pDest,
        [in] void *pSource,
        [in] LONG ByteLen);

    [entry("RtlMoveMemory")]
    void MoveMemoryFromLPSTR(
        [in] void *pDest,
        [in] LPSTR pSource,
        [in] LONG ByteLen);

    /* Global memory API */

    [entry("GlobalAlloc")]
    LONG GlobalAlloc(
        [in] GMEM uFlags,
        [in] LONG dwBytes);

    [entry("GlobalReAlloc")]
    LONG GlobalReAlloc(
        [in] LONG hMem,
        [in] LONG dwBytes,
        [in] GMEM uFlags);

    [entry("GlobalSize")]
    LONG GlobalSize(
        [in] LONG hMem);

    [entry("GlobalFlags")]
    GMEM GlobalFlags(
        [in] LONG hMem);

    [entry("GlobalLock")]
    LONG GlobalLock(
        [in] LONG hMem);

    [entry("GlobalHandle")]
    LONG GlobalHandle(
        [in] LONG pMem);

    [entry("GlobalUnlock")]
    BOOL GlobalUnlock(
        [in] LONG hMem);

    [entry("GlobalFree")]
    LONG GlobalFree(
        [in] LONG hMem);

    [entry("GlobalCompact")]
    LONG GlobalCompact(
        [in] LONG dwMinFree);

    [entry("GlobalFix")]
    void GlobalFix(
        [in] LONG hMem);

    [entry("GlobalUnfix")]
    void GlobalUnfix(
        [in] LONG hMem);

    [entry("GlobalWire")]
    long GlobalWire(
        [in] LONG hMem);

    [entry("GlobalUnWire")]
    BOOL GlobalUnWire(
        [in] LONG hMem);

    [entry("GlobalMemoryStatus")]
    void GlobalMemoryStatus(
        [in, out] MEMORYSTATUS *lpBuffer);

    [entry("GlobalMemoryStatusEx")]
    BOOL GlobalMemoryStatusEx(
        [in, out] MEMORYSTATUSEX *lpBuffer);
}
