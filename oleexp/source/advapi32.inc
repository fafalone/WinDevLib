[
    dllname("ADVAPI32.DLL")
]
module advapi32 {
	//Public Declare Function ProcessTrace Lib "advapi32" (ByVal HandleArray As Long, ByVal HandleCount As Long, ByVal pStartTime As Long, ByVal pEndTime As Long) As Long 'StartTime As FILETIME, EndTime As FILETIME) As Long
	//ProcessTrace is usually done in another thread. Defining it in a typelib allows a simple CreateThread() call.
	[entry("ProcessTrace")]
	HRESULT ProcessTrace(
		[in] long HandleArray,
		[in] long HandleCount,
		[in] long pStartTime,
		[in] long pEndTime);
};