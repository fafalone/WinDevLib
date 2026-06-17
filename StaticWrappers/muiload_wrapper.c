// muiload.c
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <muiload.h>

#pragma comment(lib, "muiload.lib")

__declspec(dllexport) HINSTANCE __stdcall Wrap_LoadMUILibraryA(
    PCSTR pszFullModuleName,
    DWORD dwLangConvention,
    LANGID LangID)
{
    return LoadMUILibraryA(pszFullModuleName, dwLangConvention, LangID);
}

__declspec(dllexport) HINSTANCE __stdcall Wrap_LoadMUILibraryW(
    PCWSTR pszFullModuleName,
    DWORD dwLangConvention,
    LANGID LangID)
{
    return LoadMUILibraryW(pszFullModuleName, dwLangConvention, LangID);
}

__declspec(dllexport) BOOL __stdcall Wrap_FreeMUILibrary(
    HMODULE hResModule)
{
    return FreeMUILibrary(hResModule);
}

__declspec(dllexport) BOOL __stdcall Wrap_GetUILanguageFallbackList(
    PWSTR pFallbackList,
    ULONG cchFallbackList,
    PULONG pcchFallbackOut)
{
    return GetUILanguageFallbackList(pFallbackList, cchFallbackList, pcchFallbackOut);
}