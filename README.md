# tbShellLib
twinBASIC Shell Library

Current Version 1.2.10 (October 5th, 2022)
(c) 2022 Jon Johnson (fafalone)

This project aims to be a twinBASIC replacement for oleexp.tlb that is x64 compatible, due to the many problems using midl to create a 64bit tlb.

oleexp is a massive project and this currently represents just a small fraction of the interfaces, but I'm expanding it all the time.

As a twinPACKAGE, the regular addin modules can be built in.

**Update (v1.2.10):** twinBASIC now supports [ PreserveSig ] as an attribute to have HRESULT values as a function return instead of only available via Err.LastHResult; tbShellLib now has this implemented wherever it is in oleexp. Like in VB, this means they're not Implements compatible, and at some point in the next few weeks, there will also be a tbShellLibImpl as a counterpart to oleexpimp.tlb. This update also adds all DirectShow interfaces (and mDirectShow.bas), most remaining oleexp interfaces, and several additional olelib interfaces.

Update (v1.1.8): Added class factory/typelib interfaces from olelib plus oleexp extensions; added manipulation.idl stuff (internial scrolling), and a few misc others.

Update (v1.1.6): Added a small number of interfaces I shouldn't have left out of the first major release... IShellExtInit, IShellExtPropPage (and related structs/apis), IQueryAssociations, IItemNameLimits, IObjectWithSite, a few others. 


For more information and a list of available interfaces, visit the [VB Forums thread](https://www.vbforums.com/showthread.php?897883-twinBASIC-tbShellLib-Shell-Interface-Library-(x64-compatible-successor-to-oleexp)) for this project.
