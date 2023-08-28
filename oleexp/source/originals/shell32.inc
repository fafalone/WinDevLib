[
    dllname("SHELL32.DLL")
]
module shell32 {

    [entry("DragQueryFileA")]
    int DragQueryFile(
        [in] LONG hDrop,
        [in] LONG iFile,
        [in] LPSTR lpszFile,
        [in] LONG cch);

    [entry("DragQueryFileW")]
    int DragQueryFileW(
        [in] LONG hDrop,
        [in] LONG iFile,
        [in] LPWSTR lpszFile,
        [in] LONG cch);

    [entry("DragQueryPoint")]
    BOOL DragQueryPoint(
        [in] LONG hDrop,
        [in, out] POINT *lppt);

    [entry("SHGetDesktopFolder")]
    HRESULT SHGetDesktopFolder(
        [out, retval] IShellFolder **ppshf);

    [entry("SHGetMalloc")]
    HRESULT SHGetMalloc(
        [out, retval] IMalloc ** ppMalloc);

    [entry("SHGetPathFromIDListA")]
    BOOL SHGetPathFromIDList(
        [in] LONG pidl,
        [in] LPSTR pszPath);

    [entry("SHGetPathFromIDListW")]
    BOOL SHGetPathFromIDListW(
        [in] LONG pidl,
        [in] LPWSTR pszPath);

    [entry("SHGetSpecialFolderLocation")]
    HRESULT SHGetSpecialFolderLocation(
        [in] LONG hwndOwner,
        [in] LONG nFolder,
        [out, retval] LONG *ppidl);

    [entry("SHGetSpecialFolderPathA")]
    BOOL SHGetSpecialFolderPath(
        [in] LONG hwndOwner,
        [in, out] LPSTR lpszPath,
        [in] LONG nFolder,
        [in] BOOL fCreate);

    [entry("SHGetSpecialFolderPathW")]
    BOOL SHGetSpecialFolderPathW(
        [in] LONG hwndOwner,
        [in, out] LPWSTR lpszPath,
        [in] LONG nFolder,
        [in] BOOL fCreate);

    [entry("SHGetInstanceExplorer")]
    HRESULT SHGetInstanceExplorer(
        [out, retval] stdole.IUnknown **ppunk);

}
