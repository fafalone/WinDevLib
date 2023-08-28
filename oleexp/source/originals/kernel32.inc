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

[
    dllname("KERNEL32.DLL")
]
module kernel32 {

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
    void GlobalWire(
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
