// mqrt.c
#include <windows.h>
#include <Mq.h>

#pragma comment(lib, "mqrt.lib")

__declspec(dllexport) HRESULT __stdcall Wrap_MQCreateQueue(
    _In_opt_ PSECURITY_DESCRIPTOR pSecurityDescriptor,
    _Inout_ MQQUEUEPROPS* pQueueProps,
    _Out_writes_opt_(*lpdwFormatNameLength)LPWSTR lpwcsFormatName,
    _Inout_ LPDWORD lpdwFormatNameLength
    )
{
    return MQCreateQueue(pSecurityDescriptor, pQueueProps, lpwcsFormatName, lpdwFormatNameLength);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQDeleteQueue(
    _In_ LPCWSTR lpwcsFormatName
    )
{
    return MQDeleteQueue(lpwcsFormatName);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQLocateBegin(
    _In_opt_ LPCWSTR lpwcsContext,
    _In_opt_ MQRESTRICTION* pRestriction,
    _In_ MQCOLUMNSET* pColumns,
    _In_ MQSORTSET* pSort,
    _Out_ PHANDLE phEnum
    )
{
    return MQLocateBegin(lpwcsContext, pRestriction, pColumns, pSort, phEnum);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQLocateNext(
    _In_ HANDLE hEnum,
    _Inout_ DWORD* pcProps,
    _Out_ MQPROPVARIANT aPropVar[]
    )
{
    return MQLocateNext(hEnum, pcProps, aPropVar);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQLocateEnd(
    _In_ HANDLE hEnum
    )
{
    return MQLocateEnd(hEnum);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQOpenQueue(
    _In_ LPCWSTR lpwcsFormatName,
    _In_ DWORD dwAccess,
    _In_ DWORD dwShareMode,
    _Out_ QUEUEHANDLE* phQueue
    )
{
    return MQOpenQueue(lpwcsFormatName, dwAccess, dwShareMode, phQueue);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQSendMessage(
    _In_ QUEUEHANDLE hDestinationQueue,
    _In_ MQMSGPROPS* pMessageProps,
    _In_opt_ ITransaction *pTransaction
    )
{
    return MQSendMessage(hDestinationQueue, pMessageProps, pTransaction);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQReceiveMessage(
    _In_ QUEUEHANDLE hSource,
    _In_ DWORD dwTimeout,
    _In_ DWORD dwAction,
    _Inout_opt_ MQMSGPROPS* pMessageProps,
    _Inout_opt_ LPOVERLAPPED lpOverlapped,
    _In_opt_ PMQRECEIVECALLBACK fnReceiveCallback,
    _In_opt_ HANDLE hCursor,
    _In_opt_ ITransaction* pTransaction
    )
{
    return MQReceiveMessage(hSource, dwTimeout, dwAction, pMessageProps, lpOverlapped, fnReceiveCallback, hCursor, pTransaction);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQReceiveMessageByLookupId(
    _In_ QUEUEHANDLE hSource,
    _In_ ULONGLONG ullLookupId,
    _In_ DWORD dwLookupAction,
    _Inout_opt_ MQMSGPROPS* pMessageProps,
    _Inout_opt_ LPOVERLAPPED lpOverlapped,
    _In_opt_ PMQRECEIVECALLBACK fnReceiveCallback,
    _In_opt_ ITransaction *pTransaction
    )
{
    return MQReceiveMessageByLookupId(hSource, ullLookupId, dwLookupAction, pMessageProps, lpOverlapped, fnReceiveCallback, pTransaction);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQCreateCursor(
    _In_ QUEUEHANDLE hQueue,
    _Out_ PHANDLE phCursor
    )
{
    return MQCreateCursor(hQueue, phCursor);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQCloseCursor(
    _In_ HANDLE hCursor
    )
{
    return MQCloseCursor(hCursor);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQCloseQueue(
    _In_ QUEUEHANDLE hQueue
    )
{
    return MQCloseQueue(hQueue);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQSetQueueProperties(
    _In_ LPCWSTR lpwcsFormatName,
    _Inout_ MQQUEUEPROPS* pQueueProps
    )
{
    return MQSetQueueProperties(lpwcsFormatName, pQueueProps);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQGetQueueProperties(
    _In_ LPCWSTR lpwcsFormatName,
    _Inout_ MQQUEUEPROPS* pQueueProps
    )
{
    return MQGetQueueProperties(lpwcsFormatName, pQueueProps);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQGetQueueSecurity(
    _In_ LPCWSTR lpwcsFormatName,
    _In_ SECURITY_INFORMATION RequestedInformation,
    _Out_writes_bytes_(nLength)  PSECURITY_DESCRIPTOR pSecurityDescriptor,
    _In_ DWORD nLength,
    _Out_ LPDWORD lpnLengthNeeded
    )
{
    return MQGetQueueSecurity(lpwcsFormatName, RequestedInformation, pSecurityDescriptor, nLength, lpnLengthNeeded);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQSetQueueSecurity(
    _In_ LPCWSTR lpwcsFormatName,
    _In_ SECURITY_INFORMATION SecurityInformation,
    _In_opt_ PSECURITY_DESCRIPTOR pSecurityDescriptor
    )
{
    return MQSetQueueSecurity(lpwcsFormatName, SecurityInformation, pSecurityDescriptor);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQPathNameToFormatName(
    _In_ LPCWSTR lpwcsPathName,
    _Out_writes_(*lpdwFormatNameLength)  LPWSTR lpwcsFormatName,
    _Inout_ LPDWORD lpdwFormatNameLength
    )
{
    return MQPathNameToFormatName(lpwcsPathName, lpwcsFormatName, lpdwFormatNameLength);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQHandleToFormatName(
    _In_ QUEUEHANDLE hQueue,
    _Out_writes_(*lpdwFormatNameLength) LPWSTR lpwcsFormatName,
    _Inout_ LPDWORD lpdwFormatNameLength
    )
{
    return MQHandleToFormatName(hQueue, lpwcsFormatName, lpdwFormatNameLength);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQInstanceToFormatName(
    _In_ GUID* pGuid,
    _Out_writes_(*lpdwFormatNameLength)  LPWSTR lpwcsFormatName,
    _Inout_ LPDWORD lpdwFormatNameLength
    )
{
    return MQInstanceToFormatName(pGuid, lpwcsFormatName, lpdwFormatNameLength);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQADsPathToFormatName(
    _In_ LPCWSTR lpwcsADsPath,
    _Out_writes_(*lpdwFormatNameLength) LPWSTR lpwcsFormatName,
    _Inout_ LPDWORD lpdwFormatNameLength
    )
{
    return MQADsPathToFormatName(lpwcsADsPath, lpwcsFormatName, lpdwFormatNameLength);
};

__declspec(dllexport) VOID __stdcall Wrap_MQFreeMemory(
    _In_ PVOID pvMemory
    )
{
    MQFreeMemory(pvMemory);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQGetMachineProperties(
    _In_opt_ LPCWSTR lpwcsMachineName,
    _In_opt_ const GUID* pguidMachineId,
    _Inout_ MQQMPROPS* pQMProps
    )
{
    return MQGetMachineProperties(lpwcsMachineName, pguidMachineId, pQMProps);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQGetSecurityContext(
    _In_reads_bytes_opt_(dwCertBufferLength) PVOID lpCertBuffer,
    _In_ DWORD dwCertBufferLength,
    _Out_ HANDLE* phSecurityContext
    )
{
    return MQGetSecurityContext(lpCertBuffer, dwCertBufferLength, phSecurityContext);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQGetSecurityContextEx(
    _In_reads_bytes_opt_(dwCertBufferLength) PVOID lpCertBuffer,
    _In_ DWORD dwCertBufferLength,
    _Out_ HANDLE* phSecurityContext
    )
{
    return MQGetSecurityContextEx(lpCertBuffer, dwCertBufferLength, phSecurityContext);
};

__declspec(dllexport) VOID __stdcall Wrap_MQFreeSecurityContext(
    _In_ HANDLE hSecurityContext
    )
{
    MQFreeSecurityContext(hSecurityContext);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQRegisterCertificate(
    _In_ DWORD dwFlags,
    _In_ PVOID lpCertBuffer,
    _In_ DWORD dwCertBufferLength
    )
{
    return MQRegisterCertificate(dwFlags, lpCertBuffer, dwCertBufferLength);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQBeginTransaction(
    _Out_ ITransaction **ppTransaction
    )
{
    return MQBeginTransaction(ppTransaction);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQGetOverlappedResult(
    _In_ LPOVERLAPPED lpOverlapped
    )
{
    return MQGetOverlappedResult(lpOverlapped);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQGetPrivateComputerInformation(
    _In_opt_ LPCWSTR lpwcsComputerName,
    _Inout_ MQPRIVATEPROPS* pPrivateProps
    )
{
    return MQGetPrivateComputerInformation(lpwcsComputerName, pPrivateProps);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQPurgeQueue(
    _In_ QUEUEHANDLE hQueue
    )
{
    return MQPurgeQueue(hQueue);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQMgmtGetInfo(
    _In_opt_ LPCWSTR pComputerName,
    _In_ LPCWSTR pObjectName,
    _Inout_ MQMGMTPROPS* pMgmtProps
    )
{
    return MQMgmtGetInfo(pComputerName, pObjectName, pMgmtProps);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQMgmtAction(
    _In_opt_ LPCWSTR pComputerName,
    _In_ LPCWSTR pObjectName,
    _In_ LPCWSTR pAction
    )
{
    return MQMgmtAction(pComputerName, pObjectName, pAction);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQMarkMessageRejected(
    _In_ HANDLE hQueue,
    _In_ ULONGLONG ullLookupId
    )
{
    return MQMarkMessageRejected(hQueue, ullLookupId);
};

__declspec(dllexport) HRESULT __stdcall Wrap_MQMoveMessage(
    _In_ QUEUEHANDLE hSourceQueue,
    _In_ QUEUEHANDLE hDestinationQueue,
    _In_ ULONGLONG ullLookupId,
    _In_opt_ ITransaction *pTransaction
    )
{
    return MQMoveMessage(hSourceQueue, hDestinationQueue, ullLookupId, pTransaction);
};

 