----------------------------------------------------------------------------------------
## WinDevLib
## Windows Development Library for twinBASIC
### (c) 2022-2023 Jon Johnson (fafalone)
----------------------------------------------------------------------------------------

This project is a comprehensive twinBASIC replacement for oleexp.tlb: http://www.vbforums.com/showthread.php?786079, 
my Modern Shell Interfaces Type Library project for VB6, that is x64 compatible, due to the many problems using midl 
to create a 64bit tlb.

This and oleexp are projects to supply Windows shell and component interfaces in a format consumable by VB6/VBA/tB. 
This involves not only defining interfaces, but using VB/tB compatible types-- so in some cases, even though there 
may be an existing way to import references to interfaces, they may be unusable due to e.g. the use of unsigned types, 
C-style arrays, double pointers, etc.

All interfaces, types, consts, and APIs from oleexp are covered, and there's additional API coverage not included in 
oleexp. For a full list of interfaces, see https://github.com/fafalone/tbShellLib/blob/main/INTERFACES.md

This project is implemented purely in tB native code, as unlike VB6 there's language support for defining interfaces and 
coclasses. As a twinPACKAGE, regular code is supported in addition to the definitions, so the regular addin modules have 
been built in (mIID.bas, mPKEY.bas, etc). Does it still make sense to use a project like this when interfaces can be 
defined in-language? I'd say yes, because for a large number of interfaces, there's deep dependency chains with other 
interfaces and the types they rely on. It makes more sense to drop this in and be done with it than constantly have to 
define the interfaces you want and then stubs for their dependencies, especially when you might need those later on. This
project is even more useful now with the API coverage; it should cover about 99% of your needs for system DLLS. 

Please report any bugs via the Issues feature here on GitHub.

----------------------------------------------------------------------------------------
Requirements
----------------------------------------------------------------------------------------
-twinBASIC Beta 269 or newer: https://github.com/twinbasic/twinbasic/releases is required.

----------------------------------------------------------------------------------------
Adding tbShellLib to your project
----------------------------------------------------------------------------------------
You have 2 options for this:

1) Via the Package Server
    twinBASIC has an online package server and tbShellLib is published on it. Open your project settings and scroll to the 
    **COM Type Library / ActiveX References**, then click **TWINPACK PACKAGES**. Add "twinBASIC Shell Library v3.4.46" (or
    whatever the newest version is). "twinBASIC Shell Library for Implements" contains `Implements` compatible versions of a 
    small number of common interfaces not defined in a compatible way in the main project; you normally don't need this. For 
    more details, including illustrations, see this post: https://github.com/fafalone/tbShellLib/issues/9#issuecomment-1416767019.

2) From a local file
    You can download the project from this repository and use the .twinpack file. Navigate to the same area as above, and click on 
    the "Import from file" button. 

----------------------------------------------------------------------------------------
Guide to switching from oleexp.tlb
----------------------------------------------------------------------------------------
tbShellLib presented the best opportunity there would be to ditch some olelib legacy baggage. It's fairly simple to move your VB6
 projects to tbShellLib, just follow these steps:

1) Replace public aliases: It's important to do this first. Run a Replace All changing oleexp.LONG_PTR to LongPtr, 
   oleexp.REFERENCE_TIME to LongLong, oleexp.HNSTIME to LongLong, oleexp.KNOWNFOLDERID to UUID, oleexp.EventRegistrationToken to 
   LongLong, oleexp.BINDPTR to LongPtr, and oleexp.LPCRITICAL_SECTION to LongPtr. If you've used them without the oleexp. prefix, 
   you'll also need to replace those, but if you've imported into tB they should be tagged.

2) Replace oleexp.IUnknown with IUnknownUnrestricted. tbShellLib keeps this separate due to the major issues with conflicts with the 
   former approach. If your project has IUnknown *without* oleexp. in front of it, **do not** replace- it's not referring to oleexp. 

3) After you've done those two, you can now go ahead and simply delete all remaining instances of `oleexp.` (including the .). 

4) Convert `Currency` to `LongLong` for interfaces and APIs: It's no longer neccessary to worry about multiplying and dividing by 
   10,000 since tB supports a true 64bit type in both 32bit and 64bit mode. So this change is ultimately for the better, but existing 
   codebases will have had to have used `Currency` for all interfaces and oleexp APIs expecting a 64bit integer.

5) Manually address any errors remaining. Interfaces should be mostly fine at this point, but if you've made use of the APIs in oleexp,
   many of them have syntax differences, mainly not being able to rewrite an ending [out] argument as the return value, and changing
   String arguments to LongPtr you'll need StrPtr with. Another major difference is that the default for almost all APIs with ANSI plus
   Unicode (A/W) versions, is now the Unicode version. A notable exception is `SendMessage` due to the overwhelming amount of VBx code 
   expecting it to mean `SendMessageA`. In most cases, the W version is declared with `LongPtr` for strings, and the untagged alias 
   version uses tB's new `DeclareWide` keyword to disable ANSI conversion while using `String`.
   Finally, a very small number of APIs and interfaces use ByVal UDTs. Since VB cannot do this, nor can tB yet, a typical workaround was 
   to pass each member as an individual argument. This worked when arguments were 4 bytes each, but the x64 calling convention aligns 
   arguments at 8 bytes. So the two options were to follow that convention, which also works for 32bit allowing a single call for both, 
   or require two different calls for 32 and 64bit. Since one of the main points of twinBASIC is 64bit support, tbShellLib uses the former 
   option. The downside of this is that VB-style calls will have to be rewritten. If you see, for example, `ByVal ptX As Long, ByVal ptY 
   As Long` replaced with `ByVal pt As LongLong`, this was an unsupported `ByVal POINT`. You'd declare a LongLong, and use `CopyMemory` to 
   fill it: `Dim pt As POINT: Dim ptt As Long: ...: CopyMemory ptt, pt, 8`.

Note that this is just for using tbShellLib-- you'll likely have a lot more changes to make if you want to make your project x64 compatible.

----------------------------------------------------------------------------------------
Guide to switching from oleexpimp.tlb
----------------------------------------------------------------------------------------
There's 'twinBASIC Shell Library for Implements' (tbShellLibImpl.twinpack/.twinproj) as well, but you'll note it has substantially fewer
interfaces than oleexpimp.tlb. This is because there's two reasons for an interface to have an alternate version: It uses `[ Preservesig ]` on 
one or more methods, or it uses `As Any`. twinBASIC allows using `Implements` with `As Any` by replacing it with `As LongPtr` (which is what 
the alternate versions do). So many interfaces were in oleexpimp.tlb for this latter reason, and subsequently are *not* included in tbShellLibImpl
as it's not neccessary.

If you find an oleexpimp.tlb interface is not in tbShellLibImpl, you will be able to use the one from tbShellLib, simply make sure `As Any` is 
changed to `As LongPtr`. 

tB has announced plans to support `[ PreserveSig ]` in implemented interfaces in the future; when that happens tbShellLibImpl will be deprecated.

----------------------------------------------------------------------------------------
tbShellLib API standards
----------------------------------------------------------------------------------------
This was mentioned above, but it's worth going into more detail. In addition to the COM interfaces, tbShellLib has a large selection of common 
Windows APIs; this is a much larger set than oleexp. tbShellLib and twinBASIC represented the best opportunity there would be to modernize 
standards... most VB programs are written with ANSI versions of APIs being the default. **This is not the case with tbShellLib**. With very few 
exceptions, APIs are Unicode by default-- i.e. they use the W, rather than A, version of APIs e.g. `DeleteFile` maps to `DeleteFileW` rather than 
`DeleteFileA`. The A and W variants use String/LongPtr, and in almost all cases, the mapped version uses `String` with twinBASIC's `DeclareWide` 
keyword-- this disables Unicode-ANSI conversion, so you can still use `String` without `StrPtr` or any Unicode <-> ANSI conversion. Note this 
usually only applies to strings passed as input, APIs passing a LPWSTR that's allocated externally will still be LongPtr, as they're not in the 
same BSTR format as VBx/TB strings.

All APIs are provided, as a minimum, as the explicit W variant, and an untagged version that maps to the W version. Some, but not all, APIs also 
have an explicit A variant defined that will perform the normal ANSI conversion for compatibility purposes. This is decided on a case by case basis 
depending on my impression of how much legacy code is around that needs the ANSI version. All new code should use the Unicode versions.

UDTs used by these calls are also supplied in the same manner, the W variant, an untagged variant that's the same as the W version, and in some
cases, an A version. UDTs always use `LongPtr` for strings, even the untagged versions for `DeclareWide`. 

As noted before, an exception to the rule is `SendMessage`, due to the enourmous volume of existing code expecting SendMessage to map to SendMessageA.

If you have any doubts about which API is being called, twinBASIC will show the full declaration when you hover your cursor over the API in your code.

----------------------------------------------------------------------------------------
A note on seeing UDTs where before they were As Any
----------------------------------------------------------------------------------------
The best example of this is many APIs, like file APIs, where in traditional VB declarations, you see 'As Any' and in tbShellLib you see e.g. 
`SECURITY_ATTRIBUTES` or `OVERLAPPED`. These are the correct the definitions, but VB6 had no facility to specify 'NULL', which is what they usually 
would be set to. So the VB6 way was a workaround, where you could pass ByVal 0. 

twinBASIC has direct support for passing a null pointer instead of a UDT. You can pass `vbNullPtr` to these arguments where previously you would have used ByVal 0 on an `As Any` argument that you've found is now a UDT. You can also pass a non-null pointer; simply pass a `LongPtr` *without* `ByVal` (for now, twinBASIC will be changing this to require `ByVal` as that makes it far more clear you intend this kind of substitution and doesn't imply you're passing ByRef LongPtr). 

Example:

VB6:
```vb6
Public Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

hFile = CreateFileW(StrPtr("name"), 0, 0, ByVal 0, ...)
```
twinBASIC:
```vb6
Public Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr

hFile = CreateFileW(StrPtr("name"), 0, 0, vbNullPtr, ...)
'---or---
Dim pSec As SECURITY_ATTRIBUTES
Dim lPtr As LongPtr = VarPtr(pSec)
hFile = CreateFileW(StrPtr("name"), 0, 0, lPtr, ...)
```
