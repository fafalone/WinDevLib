# tbShellLib
twinBASIC Shell Library

Version 1.1.6
(c) 2022 Jon Johnson (fafalone)

This project aims to be a twinBASIC replacement for oleexp.tlb that is x64 compatible, due to the many problems using midl to create a 64bit tlb.

oleexp is a massive project and this currently represents just a small fraction of the interfaces, but I'm expanding it all the time.

As a twinPACKAGE, the regular addin modules can be built in.


**Update (v1.1.6)**: Added a small number of interfaces I shouldn't have left out of the first major release... IShellExtInit, IShellExtPropPage (and related structs/apis), IQueryAssociations, IItemNameLimits, IObjectWithSite, a few others. 
