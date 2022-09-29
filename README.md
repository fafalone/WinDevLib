# tbShellLib
twinBASIC Shell Library

Version 1.0.3
(c) 2022 Jon Johnson (fafalone)

This project aims to be a twinBASIC replacement for oleexp.tlb that is x64 compatible, due to the many problems using midl to create a 64bit tlb.

oleexp is a massive project and this currently represents just a small fraction of the interfaces, but I'm expanding it all the time.

As a twinPACKAGE, the regular addin modules can be built in.


## Known Issues 

There's a bug in the current version for IShellItemImageFactory under x64. It takes a ByVal SIZE, which under x86 is done as 2xByVal Long, but on x64 must be a single ByVal LongLong for which you then use CopyMemory.
Corrected definition:

```
[ InterfaceId ("bcc18b79-ba16-442f-80c4-8a59c30c463b") ]
Interface IShellItemImageFactory Extends stdole.iunknown
    #If Win64 Then
    Sub GetImage(ByVal cxy As LongLong, ByVal flags As SIIGBF, phbm As LongPtr)
    #Else
    Sub GetImage(ByVal cx As Long, ByVal cy As Long, ByVal flags As SIIGBF, phbm As LongPtr)
    #End If
End Interface
```
