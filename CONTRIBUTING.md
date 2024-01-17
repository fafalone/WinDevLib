## Contributing to the Windows Development Library for twinBASIC

Contributions are welcome and accepted: The Windows API surface is vast, and even as large as this project is, it barely scratches the surface. There's a long tail of less commonly used stuff I simply haven't had the time for. 

### Scope of project

Contributions must be from Windows system components: DLL declares and COM interfaces/coclasses. Undocumented functions are welcome. Coverage of the Native API is welcome; I've prioritized regular the Win32 API and only included native APIs which have no standard equivalent at all, or offer significant benefits over it. But that's not a requirement now that twinBASIC has the performance to handle it. Code which is unimplemented on Windows Vista and newer is not accepted. Code for 3rd party components is not accpted unless it ships preinstalled in Windows. 

> [!NOTE]
> OpenGL APIs are on hold until tB supports aliases in-language. This is out of a desire to retain compatibility with modules and samples by NeHe, who's already done extensive work bringing OpenGL to VB6. You may submit them, with the extra requirement that don't lose the type aliases from NeHe's library, but they won't appear until supported.

### Standards

- Much work has been put in to ensuring consistency. If an argument or type member has a set of numeric constants associated it, and they're `Byte/Integer/Long` compatible, they **must** be converted to enums associated with their argument or type member. If only some are reused by other arguments or members, you can use separate enums, and associate both with `[ TypeHint(EnumA, EnumB, ...) ]`. Note that this syntax also allows associating an enum with types other than `Long`.

- `BOOL` types must be preserved. WinDevLib defines `BOOL` as an enum with CFALSE/CTRUE. All other types are converted to VBx/tB native types when they're not UDTs. Use `LongLong` where appropriate instead of `Currency` (unless the SDK explicitly defines it as such). 

- If there are separate ANSI and Unicode versions for DLL declares, like -A and -W, they must be broken down where the A variant uses `Declare` and `String` (this one is optional), a W variant that uses `LongPtr` for strings, and an aliased version that uses `DeclareWide` and `String`, linked to the Unicode version. Examples:

```vba
Public Declare PtrSafe Function GetUserNameA Lib "advapi32" (ByVal lpBuffer As String, nSize As Long) As BOOL
Public Declare PtrSafe Function GetUserNameW Lib "advapi32" (ByVal lpBuffer As LongPtr, nSize As Long) As BOOL
Public DeclareWide PtrSafe Function GetUserName Lib "advapi32" Alias "GetUserNameW" (ByVal lpBuffer As String, nSize As Long) As BOOL

Public Declare PtrSafe Function OpenFileMappingA Lib "kernel32" ([TypeHint (GenericRights, FileMapFlags) ] ByVal dwDesiredAccess As Long, ByVal bInheritHandle As BOOL, ByVal lpName As String) As LongPtr
Public Declare PtrSafe Function OpenFileMappingW Lib "kernel32" ([TypeHint (GenericRights, FileMapFlags) ] ByVal dwDesiredAccess As Long, ByVal bInheritHandle As BOOL, ByVal lpName As LongPtr) As LongPtr
Public DeclareWide PtrSafe Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingW" ([ TypeHint (GenericRights, FileMapFlags) ] ByVal dwDesiredAccess As Long, ByVal bInheritHandle As BOOL, ByVal lpName As String) As LongPtr

```

- If a declare changes from Windows version to Windows version, declares targeting Windows 7 and above, if they exist, must be included. Vista and XP are optional. Do not include defs that only work on pre-XP versions. Windows 7 is the preferred primary target, but Windows 10 is also acceptable.

- Descriptions are welcomed but not required unless there's special instructions for an API. Such as if a C-style array is handled as buffered or as a `SAFEARRAY` with special copying requirements, or if there's alternate declares for different Windows versions.

### Waiver of copyright

Contributors to this project waive any and all claims of copyright and other rights over contributions to this project, other than an acknowledgement in the readme and in comments by their code. You also certify that, to the best of your knowledge, you're not submitting any content over which other parties hold rights that would prevent inclusion in this work.

### Submitting

You can submit contributions any way you want; via PR, or by posting in issues, posting in the WinDevLib updates thread on VBForums or tB Discord, or emailing me. 


### Most Wanted

If you're looking for ideas on what to do, my current top interest is all the modern shell UI things done through undocumented interfaces related to CLSID_ImmersiveShell. 
