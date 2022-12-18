# tbShellLib
**twinBASIC Shell Library

Current Version: 2.3.38 (December 17th, 2022)

(c) 2022 Jon Johnson (fafalone)

This project aims to be a twinBASIC replacement for [oleexp.tlb](http://www.vbforums.com/showthread.php?786079-VB6-Modern-Shell-Interface-Type-Library-oleexp-tlb), my Modern Shell Interfaces Type Library project for VB6, that is x64 compatible, due to the many problems using midl to create a 64bit tlb.

It's now essentially complete; all interfaces are covered except ones useless on anything newer than Win9x, and all APIs are covered.

As a twinPACKAGE, the regular addin modules can be built in.

This project is also available via the twinBASIC Package Manager, so you can simply check the box for it there rather than download and import it manually. The versions are always kept in sync so this repo won't have anything newer. This repo is mainly for if you wish to modify it.

Please report any bugs via the Issues feature here on GitHub.

## Requirements

[twinBASIC Beta 167 or newer](https://github.com/twinbasic/twinbasic/releases) is required.


## Updates

**Update (v2.3.38):** ICategorizer::GetCategory had apidl argument incorrectly defined as ByVal.

**Update (v2.3.35):** IShellIconOverlay had incorrect pIndex params in both methods. This didn't effect 32bit projects as pointers were the same size as the index. 

**Update (v2.3.32):** Fixed GWL_* duplicate error and LARGE_INTEGER restored to hipart/lowpart for compatibility; ULARGE_INTEGER still uses quadpart if desired.

**Update (v2.3.30):** Fixed CM_COLUMNINFO bug since it was causing SetColumnInfo to trigger an automation error.

**Update (v2.3.26):** Minor bug fixes.

**Update (v2.2.24):** Added IWebBrowser2 interface I thought was already there.

**Update (v2.1.24):** twinBASIC now supports in-project `CoClass` syntax! All coclasses from oleexp have been added (I think, if you find one missing please create an issue), and can once again be used with the New keyword. The prior sCLSID constants have been left in. Also greatly expanded the API declare coverage to match what was in oleexp, though a few DLLs are still pending. Finally, tbShellLib now declares an compiler constant, `TB_SHELLLIB_DEFINED`, to help avoid conflicts with other projects (chiefly, my upcoming Common Controls 64-bit compatible library). *tbShellLib now requires twinBASIC Beta 167 or newer*.

**Major Update (v2.0.20):**  The project has reached it's initial goal of implementing all but the most obsolete oleexp.tlb interfaces. In addition, with similar exception of a small set of highly obsolete items, the API coverage is now available. Note that this was subject to extensive cleanup; native-language declares can't use the last param as retval on APIs, so all of those were converted, and TLB APIs pass Strings as BSTR, while native language passes ANSI strings, so there's currently a mix of either using LongPtr or tB's DeclareWide for BSTR/LPWSTR support (if it says String you can use a String without StrPtr). 

Update (v1.9.17): Extensive new interface additions; all remaining oleexp additions have been added including WIC and NetCon, and the majority of remaining original olelib interfaces have been added as well. 

Update (v1.2.10): twinBASIC now supports [ PreserveSig ] as an attribute to have HRESULT values as a function return instead of only available via Err.LastHResult; tbShellLib now has this implemented wherever it is in oleexp. Like in VB, this means they're not Implements compatible, and at some point in the next few weeks, there will also be a tbShellLibImpl as a counterpart to oleexpimp.tlb. This update also adds all DirectShow interfaces (and mDirectShow.bas), most remaining oleexp interfaces, and several additional olelib interfaces.

Update (v1.1.8): Added class factory/typelib interfaces from olelib plus oleexp extensions; added manipulation.idl stuff (internial scrolling), and a few misc others.

Update (v1.1.6): Added a small number of interfaces I shouldn't have left out of the first major release... IShellExtInit, IShellExtPropPage (and related structs/apis), IQueryAssociations, IItemNameLimits, IObjectWithSite, a few others. 


For more information and a list of available interfaces, visit the [VB Forums thread](https://www.vbforums.com/showthread.php?897883-twinBASIC-tbShellLib-Shell-Interface-Library-(x64-compatible-successor-to-oleexp)) for this project.
