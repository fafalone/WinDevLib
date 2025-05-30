Attribute VB_Name = "mIID"
Option Explicit

'mIID.bas by fafalone, an add-on module for oleexp.tlb.
'Revision 28 (Updated with oleexp v6.3)

'This module contains the IIDs of all oleexp.tlb interfaces, as well as all BHID_ values, FOLDERID_ values, SID_ values,
'EP_ values, GUID_DEVCLASS_ values, and GUID_DEVINTERFACE_ values.
'These can be used directly in calls as an riid/GUID argument; no need for CLSIDFromString; for example:
'SHCreateItemFromIDList(pidl, IID_IShellItem2, isi)
'
'All values in this module are directly usable in that matter, no need to convert them.
'
'
'NOTE: While this module is large, only values that your project actually uses are compiled into the exe. Unused values
'      do not appear in the exe, so despite the large size of this module, typically only a few KB's worth make it in.
'
'
'*****************
'Revision History:
'
'Rev. 5
'Added all remaining BHID_ GUID's for IShellItem.BindToHandler

'Rev. 6
'Added UUID_NULL

'Rev. 7
'Added API declare for IsEqualIID

'Rev. 8
'A number of missing IIDs were added; a small error in the automatic conversion script

'Rev. 9
'Major IID additions for oleexp 4.0

'Rev. 10
'IID additions for oleexp 4.2
'GUIDToString function added since the API doesn't seem to work

'Rev. 11
'Fixed IsEqualIID
'Fixed IID_IContextMenu/IID_IContextMenu2
'Added FreeKnownFolderDefinitionFields macro from shobjidl.h; for IKnownFolder.GetDescription

'Rev. 12
'IID additions for oleexp 4.4

'Rev. 13
'Missing IIDs ICall____

'Rev. 14
'IID additions for oleexp 4.42, 4.43

'Rev. 15
'IID additions for oleexp 4.5
'Added new FOLDERID_ values from Win10

'Rev. 16
'IID additions for oleexp 4.51

'Rev. 17
'IID additions for oleexp 4.6

'Rev. 18
'IID additions for oleexp 4.61

'Rev. 19
'IID additions for oleexp 4.7

'Rev. 20
'IID additions for oleexp 4.8

'Rev. 21
'Added FOLDERTYPEID UUIDs
'IID additions for oleexp 5.0

'Rev. 22
'IID additions for oleexp 5.01

'Rev. 23
'IID additions for oleexp 5.02

'Rev. 26
'IID additions for oleexp 5.1

'Rev 27
'IID additions for oleexp 5.2/5.3


Public Declare Function IsEqualIID Lib "ole32" Alias "IsEqualGUID" (riid1 As UUID, riid2 As UUID) As Long

Public Sub DEFINE_UUID(Name As UUID, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
  With Name
    .Data1 = L: .Data2 = w1: .Data3 = w2: .Data4(0) = B0: .Data4(1) = b1: .Data4(2) = b2: .Data4(3) = B3: .Data4(4) = b4: .Data4(5) = b5: .Data4(6) = b6: .Data4(7) = b7
  End With
End Sub
Public Sub DEFINE_OLEGUID(Name As UUID, L As Long, w1 As Integer, w2 As Integer)
  DEFINE_UUID Name, L, w1, w2, &HC0, 0, 0, 0, 0, 0, 0, &H46
End Sub
Public Function UUID_NULL() As UUID
Static bSet As Boolean
Static iid As UUID
If bSet = False Then
  With iid
    .Data1 = 0: .Data2 = 0: .Data3 = 0
    .Data4(0) = 0: .Data4(1) = 0: .Data4(2) = 0: .Data4(3) = 0: .Data4(4) = 0: .Data4(5) = 0: .Data4(6) = 0: .Data4(7) = 0
  End With
End If
bSet = True
UUID_NULL = iid
End Function
Public Function GUID_NULL() As UUID
Static bSet As Boolean
Static iid As UUID
If bSet = False Then
  With iid
    .Data1 = 0: .Data2 = 0: .Data3 = 0
    .Data4(0) = 0: .Data4(1) = 0: .Data4(2) = 0: .Data4(3) = 0: .Data4(4) = 0: .Data4(5) = 0: .Data4(6) = 0: .Data4(7) = 0
  End With
End If
bSet = True
GUID_NULL = iid
End Function
Public Sub DEFINE_PROPERTYKEY(Name As PROPERTYKEY, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte, pid As Long)
  With Name.fmtid
    .Data1 = L: .Data2 = w1: .Data3 = w2: .Data4(0) = B0: .Data4(1) = b1: .Data4(2) = b2: .Data4(3) = B3: .Data4(4) = b4: .Data4(5) = b5: .Data4(6) = b6: .Data4(7) = b7
  End With
  Name.pid = pid
End Sub



Public Function GUIDToString(tg As UUID, Optional bBrack As Boolean = True) As String
'StringFromGUID2 never works, even "working" code from vbaccelerator AND MSDN
GUIDToString = Right$("00000000" & Hex$(tg.Data1), 8) & "-" & Right$("0000" & Hex$(tg.Data2), 4) & "-" & Right$("0000" & Hex$(tg.Data3), 4) & _
"-" & Right$("00" & Hex$(CLng(tg.Data4(0))), 2) & Right$("00" & Hex$(CLng(tg.Data4(1))), 2) & "-" & Right$("00" & Hex$(CLng(tg.Data4(2))), 2) & _
Right$("00" & Hex$(CLng(tg.Data4(3))), 2) & Right$("00" & Hex$(CLng(tg.Data4(4))), 2) & Right$("00" & Hex$(CLng(tg.Data4(5))), 2) & _
Right$("00" & Hex$(CLng(tg.Data4(6))), 2) & Right$("00" & Hex$(CLng(tg.Data4(7))), 2)
If bBrack Then GUIDToString = "{" & GUIDToString & "}"
End Function

'====================================================
Public Function IID_FolderItem() As UUID
'{FAC32C80-CBE4-11CE-8350-444553540000}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFAC32C80, CInt(&HCBE4), CInt(&H11CE), &H83, &H50, &H44, &H45, &H53, &H54, &H0, &H0)
 IID_FolderItem = iid
End Function
Public Function IID_FolderItem2() As UUID
'{edc817aa-92b8-11d1-b075-00c04fc33aa5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEDC817AA, CInt(&H92B8), CInt(&H11D1), &HB0, &H75, &H0, &HC0, &H4F, &HC3, &H3A, &HA5)
 IID_FolderItem2 = iid
End Function
Public Function IID_FolderItems() As UUID
'{744129E0-CBE5-11CE-8350-444553540000}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H744129E0, CInt(&HCBE5), CInt(&H11CE), &H83, &H50, &H44, &H45, &H53, &H54, &H0, &H0)
 IID_FolderItems = iid
End Function
Public Function IID_FolderItems2() As UUID
'{C94F0AD0-F363-11d2-A327-00C04F8EEC7F}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC94F0AD0, CInt(&HF363), CInt(&H11D2), &HA3, &H27, &H0, &HC0, &H4F, &H8E, &HEC, &H7F)
 IID_FolderItems2 = iid
End Function
Public Function IID_Folder() As UUID
'{BBCBDE60-C3FF-11CE-8350-444553540000}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBBCBDE60, CInt(&HC3FF), CInt(&H11CE), &H83, &H50, &H44, &H45, &H53, &H54, &H0, &H0)
 IID_Folder = iid
End Function
Public Function IID_Folder2() As UUID
'{f0d2d8ef-3890-11d2-bf8b-00c04fb93661}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF0D2D8EF, CInt(&H3890), CInt(&H11D2), &HBF, &H8B, &H0, &HC0, &H4F, &HB9, &H36, &H61)
 IID_Folder2 = iid
End Function
Public Function IID_Folder3() As UUID
'{A7AE5F64-C4D7-4d7f-9307-4D24EE54B841}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA7AE5F64, CInt(&HC4D7), CInt(&H4D7F), &H93, &H7, &H4D, &H24, &HEE, &H54, &HB8, &H41)
 IID_Folder3 = iid
End Function
Public Function IID_FolderItemVerb() As UUID
'{08EC3E00-50B0-11CF-960C-0080C7F4EE85}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8EC3E00, CInt(&H50B0), CInt(&H11CF), &H96, &HC, &H0, &H80, &HC7, &HF4, &HEE, &H85)
 IID_FolderItemVerb = iid
End Function
Public Function IID_FolderItemVerbs() As UUID
'{1F8352C0-50B0-11CF-960C-0080C7F4EE85}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1F8352C0, CInt(&H50B0), CInt(&H11CF), &H96, &HC, &H0, &H80, &HC7, &HF4, &HEE, &H85)
 IID_FolderItemVerbs = iid
End Function
Public Function IID_IShellFolderViewDual() As UUID
'{E7A1AF80-4D96-11CF-960C-0080C7F4EE85}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE7A1AF80, CInt(&H4D96), CInt(&H11CF), &H96, &HC, &H0, &H80, &HC7, &HF4, &HEE, &H85)
 IID_IShellFolderViewDual = iid
End Function
Public Function IID_IShellFolderViewDual2() As UUID
'{31C147b6-0ADE-4A3C-B514-DDF932EF6D17}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H31C147B6, CInt(&HADE), CInt(&H4A3C), &HB5, &H14, &HDD, &HF9, &H32, &HEF, &H6D, &H17)
 IID_IShellFolderViewDual2 = iid
End Function
Public Function IID_IShellFolderViewDual3() As UUID
'{29EC8E6C-46D3-411f-BAAA-611A6C9CAC66}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H29EC8E6C, CInt(&H46D3), CInt(&H411F), &HBA, &HAA, &H61, &H1A, &H6C, &H9C, &HAC, &H66)
 IID_IShellFolderViewDual3 = iid
End Function
Public Function IID_IWebBrowser() As UUID
'{EAB22AC1-30C1-11CF-A7EB-0000C05BAE0B}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEAB22AC1, CInt(&H30C1), CInt(&H11CF), &HA7, &HEB, &H0, &H0, &HC0, &H5B, &HAE, &HB)
 IID_IWebBrowser = iid
End Function
Public Function IID_IWebBrowserApp() As UUID
'{0002DF05-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2DF05, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IWebBrowserApp = iid
End Function
Public Function IID_IWebBrowser2() As UUID
'{D30C1661-CDAF-11d0-8A3E-00C04FC9E26E}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD30C1661, CInt(&HCDAF), CInt(&H11D0), &H8A, &H3E, &H0, &HC0, &H4F, &HC9, &HE2, &H6E)
 IID_IWebBrowser2 = iid
End Function

Public Function IID_IFolderViewOC() As UUID
'{9BA05970-F6A8-11CF-A442-00A0C90A8F39}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9BA05970, CInt(&HF6A8), CInt(&H11CF), &HA4, &H42, &H0, &HA0, &HC9, &HA, &H8F, &H39)
 IID_IFolderViewOC = iid
End Function
Public Function IID_DShellFolderViewEvents() As UUID
'{62112AA2-EBE4-11cf-A5FB-0020AFE7292D}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H62112AA2, CInt(&HEBE4), CInt(&H11CF), &HA5, &HFB, &H0, &H20, &HAF, &HE7, &H29, &H2D)
 IID_DShellFolderViewEvents = iid
End Function
Public Function IID_DFConstraint() As UUID
'{4a3df050-23bd-11d2-939f-00a0c91eedba}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4A3DF050, CInt(&H23BD), CInt(&H11D2), &H93, &H9F, &H0, &HA0, &HC9, &H1E, &HED, &HBA)
 IID_DFConstraint = iid
End Function
Public Function IID_IShellLinkDual() As UUID
'{88A05C00-F000-11CE-8350-444553540000}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H88A05C00, CInt(&HF000), CInt(&H11CE), &H83, &H50, &H44, &H45, &H53, &H54, &H0, &H0)
 IID_IShellLinkDual = iid
End Function
Public Function IID_IShellLinkDual2() As UUID
'{317EE249-F12E-11d2-B1E4-00C04F8EEB3E}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H317EE249, CInt(&HF12E), CInt(&H11D2), &HB1, &HE4, &H0, &HC0, &H4F, &H8E, &HEB, &H3E)
 IID_IShellLinkDual2 = iid
End Function
Public Function IID_IShellDispatch() As UUID
'{D8F015C0-C278-11CE-A49E-444553540000}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD8F015C0, CInt(&HC278), CInt(&H11CE), &HA4, &H9E, &H44, &H45, &H53, &H54, &H0, &H0)
 IID_IShellDispatch = iid
End Function
Public Function IID_IShellDispatch2() As UUID
'{A4C6892C-3BA9-11d2-9DEA-00C04FB16162}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA4C6892C, CInt(&H3BA9), CInt(&H11D2), &H9D, &HEA, &H0, &HC0, &H4F, &HB1, &H61, &H62)
 IID_IShellDispatch2 = iid
End Function
Public Function IID_IShellDispatch3() As UUID
'{177160ca-bb5a-411c-841d-bd38facdeaa0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H177160CA, CInt(&HBB5A), CInt(&H411C), &H84, &H1D, &HBD, &H38, &HFA, &HCD, &HEA, &HA0)
 IID_IShellDispatch3 = iid
End Function
Public Function IID_IShellDispatch4() As UUID
'{efd84b2d-4bcf-4298-be25-eb542a59fbda}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEFD84B2D, CInt(&H4BCF), CInt(&H4298), &HBE, &H25, &HEB, &H54, &H2A, &H59, &HFB, &HDA)
 IID_IShellDispatch4 = iid
End Function
Public Function IID_IShellDispatch5() As UUID
'{866738b9-6cf2-4de8-8767-f794ebe74f4e}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H866738B9, CInt(&H6CF2), CInt(&H4DE8), &H87, &H67, &HF7, &H94, &HEB, &HE7, &H4F, &H4E)
 IID_IShellDispatch5 = iid
End Function
Public Function IID_IShellDispatch6() As UUID
'{286e6f1b-7113-4355-9562-96b7e9d64c5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H286E6F1B, CInt(&H7113), CInt(&H4355), &H95, &H62, &H96, &HB7, &HE9, &HD6, &H4C, &H5)
 IID_IShellDispatch6 = iid
End Function
Public Function IID_IShellUIHelper() As UUID
'{729FE2F8-1EA8-11d1-8F85-00C04FC2FBE1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H729FE2F8, CInt(&H1EA8), CInt(&H11D1), &H8F, &H85, &H0, &HC0, &H4F, &HC2, &HFB, &HE1)
 IID_IShellUIHelper = iid
End Function
Public Function IID_IShellUIHelper2() As UUID
'{a7fe6eda-1932-4281-b881-87b31b8bc52c}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA7FE6EDA, CInt(&H1932), CInt(&H4281), &HB8, &H81, &H87, &HB3, &H1B, &H8B, &HC5, &H2C)
 IID_IShellUIHelper2 = iid
End Function
Public Function IID_IShellUIHelper3() As UUID
'{528DF2EC-D419-40bc-9B6D-DCDBF9C1B25D}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H528DF2EC, CInt(&HD419), CInt(&H40BC), &H9B, &H6D, &HDC, &HDB, &HF9, &HC1, &HB2, &H5D)
 IID_IShellUIHelper3 = iid
End Function
Public Function IID_IShellUIHelper4() As UUID
'{B36E6A53-8073-499E-824C-D776330A333E}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB36E6A53, CInt(&H8073), CInt(&H499E), &H82, &H4C, &HD7, &H76, &H33, &HA, &H33, &H3E)
 IID_IShellUIHelper4 = iid
End Function
Public Function IID_IShellUIHelper5() As UUID
'{A2A08B09-103D-4D3F-B91C-EA455CA82EFA}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA2A08B09, CInt(&H103D), CInt(&H4D3F), &HB9, &H1C, &HEA, &H45, &H5C, &HA8, &H2E, &HFA)
 IID_IShellUIHelper5 = iid
End Function
Public Function IID_IShellUIHelper6() As UUID
'{987A573E-46EE-4E89-96AB-DDF7F8FDC98C}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H987A573E, CInt(&H46EE), CInt(&H4E89), &H96, &HAB, &HDD, &HF7, &HF8, &HFD, &HC9, &H8C)
 IID_IShellUIHelper6 = iid
End Function
Public Function IID_IShellUIHelper7() As UUID
'{60E567C8-9573-4AB2-A264-637C6C161CB1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H60E567C8, CInt(&H9573), CInt(&H4AB2), &HA2, &H64, &H63, &H7C, &H6C, &H16, &H1C, &HB1)
 IID_IShellUIHelper7 = iid
End Function
Public Function IID_IShellUIHelper8() As UUID
'{66DEBCF2-05B0-4F07-B49B-B96241A65DB2}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H66DEBCF2, CInt(&H5B0), CInt(&H4F07), &HB4, &H9B, &HB9, &H62, &H41, &HA6, &H5D, &HB2)
 IID_IShellUIHelper8 = iid
End Function
Public Function IID_IShellUIHelper9() As UUID
'{6cdf73b0-7f2f-451f-bc0f-63e0f3284e54}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6CDF73B0, CInt(&H7F2F), CInt(&H451F), &HBC, &HF, &H63, &HE0, &HF3, &H28, &H4E, &H54)
 IID_IShellUIHelper9 = iid
End Function
Public Function IID_IShellFavoritesNameSpace() As UUID
'{55136804-B2DE-11D1-B9F2-00A0C98BC547}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H55136804, CInt(&HB2DE), CInt(&H11D1), &HB9, &HF2, &H0, &HA0, &HC9, &H8B, &HC5, &H47)
 IID_IShellFavoritesNameSpace = iid
End Function
Public Function IID_IShellNameSpace() As UUID
'{e572d3c9-37be-4ae2-825d-d521763e3108}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE572D3C9, CInt(&H37BE), CInt(&H4AE2), &H82, &H5D, &HD5, &H21, &H76, &H3E, &H31, &H8)
 IID_IShellNameSpace = iid
End Function
Public Function IID_IScriptErrorList() As UUID
'{F3470F24-15FD-11d2-BB2E-00805FF7EFCA}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF3470F24, CInt(&H15FD), CInt(&H11D2), &HBB, &H2E, &H0, &H80, &H5F, &HF7, &HEF, &HCA)
 IID_IScriptErrorList = iid
End Function



'====================================================
'IIDs added in Rev. 8
'====================================================
Public Function IID_IShellExtInit() As UUID
'{000214E8-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H214E8, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IShellExtInit = iid
End Function
Public Function IID_IShellExecuteHookA() As UUID
'{000214F5-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H214F5, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IShellExecuteHookA = iid
End Function
Public Function IID_IShellExecuteHookW() As UUID
'{000214FB-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H214FB, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IShellExecuteHookW = iid
End Function
Public Function IID_IEnumExtraSearch() As UUID
'{0E700BE1-9DB6-11d1-A1CE-00C04FD75D13}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE700BE1, CInt(&H9DB6), CInt(&H11D1), &HA1, &HCE, &H0, &HC0, &H4F, &HD7, &H5D, &H13)
IID_IEnumExtraSearch = iid
End Function
Public Function IID_IFolderFilterSite() As UUID
'{C0A651F5-B48B-11d2-B5ED-006097C686F6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC0A651F5, CInt(&HB48B), CInt(&H11D2), &HB5, &HED, &H0, &H60, &H97, &HC6, &H86, &HF6)
IID_IFolderFilterSite = iid
End Function
Public Function IID_IFileSystemBindData() As UUID
'{01E18D10-4D8B-11d2-855D-006008059367}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1E18D10, CInt(&H4D8B), CInt(&H11D2), &H85, &H5D, &H0, &H60, &H8, &H5, &H93, &H67)
 IID_IFileSystemBindData = iid
End Function
Public Function IID_IFileSystemBindData2() As UUID
'{3acf075f-71db-4afa-81f0-3fc4fdf2a5b8}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3ACF075F, CInt(&H71DB), CInt(&H4AFA), &H81, &HF0, &H3F, &HC4, &HFD, &HF2, &HA5, &HB8)
 IID_IFileSystemBindData2 = iid
End Function
Public Function IID_IObjectWithFolderEnumMode() As UUID
'{6a9d9026-0e6e-464c-b000-42ecc07de673}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6A9D9026, CInt(&HE6E), CInt(&H464C), &HB0, &H0, &H42, &HEC, &HC0, &H7D, &HE6, &H73)
 IID_IObjectWithFolderEnumMode = iid
End Function
Public Function IID_IProfferService() As UUID
'{cb728b20-f786-11ce-92ad-00aa00a74cd0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCB728B20, CInt(&HF786), CInt(&H11CE), &H92, &HAD, &H0, &HAA, &H0, &HA7, &H4C, &HD0)
IID_IProfferService = iid
End Function
Public Function IID_IPropertyUI() As UUID
'{757a7d9f-919a-4118-99d7-dbb208c8cc66}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H757A7D9F, CInt(&H919A), CInt(&H4118), &H99, &HD7, &HDB, &HB2, &H8, &HC8, &HCC, &H66)
IID_IPropertyUI = iid
End Function
Public Function IID_ICategoryProvider() As UUID
'{9af64809-5864-4c26-a720-c1f78c086ee3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9AF64809, CInt(&H5864), CInt(&H4C26), &HA7, &H20, &HC1, &HF7, &H8C, &H8, &H6E, &HE3)
IID_ICategoryProvider = iid
End Function
Public Function IID_ICategorizer() As UUID
'{a3b14589-9174-49a8-89a3-06a1ae2b9ba7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA3B14589, CInt(&H9174), CInt(&H49A8), &H89, &HA3, &H6, &HA1, &HAE, &H2B, &H9B, &HA7)
IID_ICategorizer = iid
End Function
Public Function IID_IUserEventTimerCallback() As UUID
'{e9ead8e6-2a25-410e-9b58-a9fbef1dd1a2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE9EAD8E6, CInt(&H2A25), CInt(&H410E), &H9B, &H58, &HA9, &HFB, &HEF, &H1D, &HD1, &HA2)
IID_IUserEventTimerCallback = iid
End Function
Public Function IID_IUserEventTimer() As UUID
'{0F504B94-6E42-42E6-99E0-E20FAFE52AB4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF504B94, CInt(&H6E42), CInt(&H42E6), &H99, &HE0, &HE2, &HF, &HAF, &HE5, &H2A, &HB4)
IID_IUserEventTimer = iid
End Function
Public Function IID_IWebWizardExtension() As UUID
'{0e6b3f66-98d1-48c0-a222-fbde74e2fbc5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE6B3F66, CInt(&H98D1), CInt(&H48C0), &HA2, &H22, &HFB, &HDE, &H74, &HE2, &HFB, &HC5)
IID_IWebWizardExtension = iid
End Function
Public Function IID_IPublishingWizard() As UUID
'{aa9198bb-ccec-472d-beed-19a4f6733f7a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAA9198BB, CInt(&HCCEC), CInt(&H472D), &HBE, &HED, &H19, &HA4, &HF6, &H73, &H3F, &H7A)
IID_IPublishingWizard = iid
End Function
Public Function IID_INetCrawler() As UUID
''{49c929ee-a1b7-4c58-b539-e63be392b6f3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H49C929EE, CInt(&HA1B7), CInt(&H4C58), &HB5, &H39, &HE6, &H3B, &HE3, &H92, &HB6, &HF3)
IID_INetCrawler = iid
End Function
Public Function IID_IAsyncOperation() As UUID
'{3D8B0590-F691-11d2-8EA9-006097DF5BD4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3D8B0590, CInt(&HF691), CInt(&H11D2), &H8E, &HA9, &H0, &H60, &H97, &HDF, &H5B, &HD4)
IID_IAsyncOperation = iid
End Function
Public Function IID_ITypeInfo2() As UUID
'{00020412-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20412, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ITypeInfo2 = iid
End Function
Public Function IID_ITypeLib() As UUID
'{00020402-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20402, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ITypeLib = iid
End Function
'==================================================================
'End Rev. 8 Update
'==================================================================
Public Function IID_ITypeMarshal() As UUID
'{0000002D-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2D, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_ITypeMarshal = iid
End Function
Public Function IID_ITypeFactory() As UUID
'{0000002E-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2E, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_ITypeFactory = iid
End Function
Public Function IID_ITypeChangeEvents() As UUID
'{00020410-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H20410, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_ITypeChangeEvents = iid
End Function
Public Function IID_ITypeLibRegistrationReader() As UUID
'{ED6A8A2A-B160-4E77-8F73-AA7435CD5C27}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HED6A8A2A, CInt(&HB160), CInt(&H4E77), &H8F, &H73, &HAA, &H74, &H35, &HCD, &H5C, &H27)
 IID_ITypeLibRegistrationReader = iid
End Function
Public Function IID_ITypeLibRegistration() As UUID
'{76A3E735-02DF-4A12-98EB-043AD3600AF3}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H76A3E735, CInt(&H2DF), CInt(&H4A12), &H98, &HEB, &H4, &H3A, &HD3, &H60, &HA, &HF3)
 IID_ITypeLibRegistration = iid
End Function
Public Function IID_IFrameworkInputPaneHandler() As UUID
'{226C537B-1E76-4D9E-A760-33DB29922F18}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H226C537B, CInt(&H1E76), CInt(&H4D9E), &HA7, &H60, &H33, &HDB, &H29, &H92, &H2F, &H18)
 IID_IFrameworkInputPaneHandler = iid
End Function
Public Function IID_IFrameworkInputPane() As UUID
'{5752238B-24F0-495A-82F1-2FD593056796}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5752238B, CInt(&H24F0), CInt(&H495A), &H82, &HF1, &H2F, &HD5, &H93, &H5, &H67, &H96)
 IID_IFrameworkInputPane = iid
End Function
Public Function IID_ISharingConfigurationManager() As UUID
'{B4CD448A-9C86-4466-9201-2E62105B87AE}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB4CD448A, CInt(&H9C86), CInt(&H4466), &H92, &H1, &H2E, &H62, &H10, &H5B, &H87, &HAE)
 IID_ISharingConfigurationManager = iid
End Function
Public Function IID_IRunnableTask() As UUID
'{85788d00-6807-11d0-b810-00c04fd706ec}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H85788D00, CInt(&H6807), CInt(&H11D0), &HB8, &H10, &H0, &HC0, &H4F, &HD7, &H6, &HEC)
 IID_IRunnableTask = iid
End Function
Public Function IID_IShellTaskScheduler() As UUID
'{6CCB7BE0-6807-11d0-B810-00C04FD706EC}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6CCB7BE0, CInt(&H6807), CInt(&H11D0), &HB8, &H10, &H0, &HC0, &H4F, &HD7, &H6, &HEC)
 IID_IShellTaskScheduler = iid
End Function
Public Function IID_IAccessible() As UUID
'{618736e0-3c3d-11cf-810c-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H618736E0, CInt(&H3C3D), CInt(&H11CF), &H81, &HC, &H0, &HAA, &H0, &H38, &H9B, &H71)
IID_IAccessible = iid
End Function
Public Function IID_IAccessibleHandler() As UUID
'{03022430-ABC4-11d0-BDE2-00AA001A1953}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3022430, CInt(&HABC4), CInt(&H11D0), &HBD, &HE2, &H0, &HAA, &H0, &H1A, &H19, &H53)
IID_IAccessibleHandler = iid
End Function
Public Function IID_IAccessibleWindowlessSite() As UUID
'{BF3ABD9C-76DA-4389-9EB6-1427D25ABAB7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBF3ABD9C, CInt(&H76DA), CInt(&H4389), &H9E, &HB6, &H14, &H27, &HD2, &H5A, &HBA, &HB7)
IID_IAccessibleWindowlessSite = iid
End Function
Public Function IID_IAccIdentity() As UUID
'{7852b78d-1cfd-41c1-a615-9c0c85960b5f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7852B78D, CInt(&H1CFD), CInt(&H41C1), &HA6, &H15, &H9C, &HC, &H85, &H96, &HB, &H5F)
IID_IAccIdentity = iid
End Function
Public Function IID_IAccPropServer() As UUID
'{76c0dbbb-15e0-4e7b-b61b-20eeea2001e0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H76C0DBBB, CInt(&H15E0), CInt(&H4E7B), &HB6, &H1B, &H20, &HEE, &HEA, &H20, &H1, &HE0)
IID_IAccPropServer = iid
End Function
Public Function IID_IAccPropServices() As UUID
'{6e26e776-04f0-495d-80e4-3330352e3169}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6E26E776, CInt(&H4F0), CInt(&H495D), &H80, &HE4, &H33, &H30, &H35, &H2E, &H31, &H69)
IID_IAccPropServices = iid
End Function
Public Function IID_IExplorerBrowserEvents() As UUID
'{361bbdc7-e6ee-4e13-be58-58e2240c810f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H361BBDC7, CInt(&HE6EE), CInt(&H4E13), &HBE, &H58, &H58, &HE2, &H24, &HC, &H81, &HF)
IID_IExplorerBrowserEvents = iid
End Function
Public Function IID_IExplorerBrowser() As UUID
'{dfd3b6b5-c10c-4be9-85f6-a66969f402f6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDFD3B6B5, CInt(&HC10C), CInt(&H4BE9), &H85, &HF6, &HA6, &H69, &H69, &HF4, &H2, &HF6)
IID_IExplorerBrowser = iid
End Function
Public Function IID_IExplorerPaneVisibility() As UUID
'{e07010ec-bc17-44c0-97b0-46c7c95b9edc}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE07010EC, CInt(&HBC17), CInt(&H44C0), &H97, &HB0, &H46, &HC7, &HC9, &H5B, &H9E, &HDC)
IID_IExplorerPaneVisibility = iid
End Function
Public Function IID_INameSpaceTreeControl() As UUID
'{028212A3-B627-47e9-8856-C14265554E4F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H28212A3, CInt(&HB627), CInt(&H47E9), &H88, &H56, &HC1, &H42, &H65, &H55, &H4E, &H4F)
IID_INameSpaceTreeControl = iid
End Function
Public Function IID_INameSpaceTreeControl2() As UUID
'{7cc7aed8-290e-49bc-8945-c1401cc9306c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7CC7AED8, CInt(&H290E), CInt(&H49BC), &H89, &H45, &HC1, &H40, &H1C, &HC9, &H30, &H6C)
IID_INameSpaceTreeControl2 = iid
End Function
Public Function IID_INameSpaceTreeControlEvents() As UUID
'{93D77985-B3D8-4484-8318-672CDDA002CE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H93D77985, CInt(&HB3D8), CInt(&H4484), &H83, &H18, &H67, &H2C, &HDD, &HA0, &H2, &HCE)
IID_INameSpaceTreeControlEvents = iid
End Function
Public Function IID_INameSpaceTreeControlDropHandler() As UUID
'{F9C665D6-C2F2-4c19-BF33-8322D7352F51}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF9C665D6, CInt(&HC2F2), CInt(&H4C19), &HBF, &H33, &H83, &H22, &HD7, &H35, &H2F, &H51)
IID_INameSpaceTreeControlDropHandler = iid
End Function
Public Function IID_INameSpaceTreeAccessible() As UUID
'{71f312de-43ed-4190-8477-e9536b82350b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H71F312DE, CInt(&H43ED), CInt(&H4190), &H84, &H77, &HE9, &H53, &H6B, &H82, &H35, &HB)
IID_INameSpaceTreeAccessible = iid
End Function
Public Function IID_INameSpaceTreeControlCustomDraw() As UUID
'{2D3BA758-33EE-42d5-BB7B-5F3431D86C78}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2D3BA758, CInt(&H33EE), CInt(&H42D5), &HBB, &H7B, &H5F, &H34, &H31, &HD8, &H6C, &H78)
IID_INameSpaceTreeControlCustomDraw = iid
End Function
Public Function IID_INameSpaceTreeControlFolderCapabilities() As UUID
'{e9701183-e6b3-4ff2-8568-813615fec7be}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE9701183, CInt(&HE6B3), CInt(&H4FF2), &H85, &H68, &H81, &H36, &H15, &HFE, &HC7, &HBE)
IID_INameSpaceTreeControlFolderCapabilities = iid
End Function
Public Function IID_IShellWindows() As UUID
'{85CB6900-4D95-11CF-960C-0080C7F4EE85}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H85CB6900, CInt(&H4D95), CInt(&H11CF), &H96, &HC, &H0, &H80, &HC7, &HF4, &HEE, &H85)
IID_IShellWindows = iid
End Function
Public Function IID_IStreamAsync() As UUID
'{fe0b6665-e0ca-49b9-a178-2b5cb48d92a5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFE0B6665, CInt(&HE0CA), CInt(&H49B9), &HA1, &H78, &H2B, &H5C, &HB4, &H8D, &H92, &HA5)
IID_IStreamAsync = iid
End Function
Public Function IID_IEnumFullIDList() As UUID
'{d0191542-7954-4908-bc06-b2360bbe45ba}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD0191542, CInt(&H7954), CInt(&H4908), &HBC, &H6, &HB2, &H36, &HB, &HBE, &H45, &HBA)
IID_IEnumFullIDList = iid
End Function
Public Function IID_IShellView3() As UUID
'{ec39fa88-f8af-41c5-8421-38bed28f4673}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEC39FA88, CInt(&HF8AF), CInt(&H41C5), &H84, &H21, &H38, &HBE, &HD2, &H8F, &H46, &H73)
IID_IShellView3 = iid
End Function
Public Function IID_ICommDlgBrowser() As UUID
'{000214F1-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214F1, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICommDlgBrowser = iid
End Function
Public Function IID_ICommDlgBrowser2() As UUID
'{10339516-2894-11d2-9039-00C04F8EEB3E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10339516, CInt(&H2894), CInt(&H11D2), &H90, &H39, &H0, &HC0, &H4F, &H8E, &HEB, &H3E)
IID_ICommDlgBrowser2 = iid
End Function
Public Function IID_ICommDlgBrowser3() As UUID
'{c8ad25a1-3294-41ee-8165-71174bd01c57}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC8AD25A1, CInt(&H3294), CInt(&H41EE), &H81, &H65, &H71, &H17, &H4B, &HD0, &H1C, &H57)
IID_ICommDlgBrowser3 = iid
End Function
Public Function IID_IColumnManager() As UUID
'{d8ec27bb-3f3b-4042-b10a-4acfd924d453}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD8EC27BB, CInt(&H3F3B), CInt(&H4042), &HB1, &HA, &H4A, &HCF, &HD9, &H24, &HD4, &H53)
IID_IColumnManager = iid
End Function
Public Function IID_ITaskbarList3() As UUID
'{ea1afb91-9e28-4b86-90e9-9e9f8a5eefaf}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEA1AFB91, CInt(&H9E28), CInt(&H4B86), &H90, &HE9, &H9E, &H9F, &H8A, &H5E, &HEF, &HAF)
IID_ITaskbarList3 = iid
End Function
Public Function IID_ITaskbarList4() As UUID
'{c43dc798-95d1-4bea-9030-bb99e2983a1a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC43DC798, CInt(&H95D1), CInt(&H4BEA), &H90, &H30, &HBB, &H99, &HE2, &H98, &H3A, &H1A)
IID_ITaskbarList4 = iid
End Function
Public Function IID_IThumbnailProvider() As UUID
'{e357fccd-a995-4576-b01f-234630154e96}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE357FCCD, CInt(&HA995), CInt(&H4576), &HB0, &H1F, &H23, &H46, &H30, &H15, &H4E, &H96)
IID_IThumbnailProvider = iid
End Function
Public Function IID_IOperationsProgressDialog() As UUID
'{0C9FB851-E5C9-43EB-A370-F0677B13874C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC9FB851, CInt(&HE5C9), CInt(&H43EB), &HA3, &H70, &HF0, &H67, &H7B, &H13, &H87, &H4C)
IID_IOperationsProgressDialog = iid
End Function
Public Function IID_IFileOperationProgressSink() As UUID
'{04b0f1a7-9490-44bc-96e1-4296a31252e2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4B0F1A7, CInt(&H9490), CInt(&H44BC), &H96, &HE1, &H42, &H96, &HA3, &H12, &H52, &HE2)
IID_IFileOperationProgressSink = iid
End Function
Public Function IID_IFileOperation() As UUID
'{947aab5f-0a5c-4c13-b4d6-4bf7836fc9f8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H947AAB5F, CInt(&HA5C), CInt(&H4C13), &HB4, &HD6, &H4B, &HF7, &H83, &H6F, &HC9, &HF8)
IID_IFileOperation = iid
End Function
Public Function IID_IObjectCollection() As UUID
'{5632b1a4-e38a-400a-928a-d4cd63230295}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5632B1A4, CInt(&HE38A), CInt(&H400A), &H92, &H8A, &HD4, &HCD, &H63, &H23, &H2, &H95)
IID_IObjectCollection = iid
End Function
Public Function IID_IApplicationDestinations() As UUID
'{12337d35-94c6-48a0-bce7-6a9c69d4d600}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H12337D35, CInt(&H94C6), CInt(&H48A0), &HBC, &HE7, &H6A, &H9C, &H69, &HD4, &HD6, &H0)
IID_IApplicationDestinations = iid
End Function
Public Function IID_ICustomDestinationList() As UUID
'{6332debf-87b5-4670-90c0-5e57b408a49e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6332DEBF, CInt(&H87B5), CInt(&H4670), &H90, &HC0, &H5E, &H57, &HB4, &H8, &HA4, &H9E)
IID_ICustomDestinationList = iid
End Function
Public Function IID_IModalWindow() As UUID
'{b4db1657-70d7-485e-8e3e-6fcb5a5c1802}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB4DB1657, CInt(&H70D7), CInt(&H485E), &H8E, &H3E, &H6F, &HCB, &H5A, &H5C, &H18, &H2)
IID_IModalWindow = iid
End Function
Public Function IID_IFileDialogEvents() As UUID
'{973510db-7d7f-452b-8975-74a85828d354}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H973510DB, CInt(&H7D7F), CInt(&H452B), &H89, &H75, &H74, &HA8, &H58, &H28, &HD3, &H54)
IID_IFileDialogEvents = iid
End Function
Public Function IID_IShellItemFilter() As UUID
'{2659B475-EEB8-48b7-8F07-B378810F48CF}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2659B475, CInt(&HEEB8), CInt(&H48B7), &H8F, &H7, &HB3, &H78, &H81, &HF, &H48, &HCF)
IID_IShellItemFilter = iid
End Function
Public Function IID_IFileDialog() As UUID
'{42f85136-db7e-439c-85f1-e4075d135fc8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H42F85136, CInt(&HDB7E), CInt(&H439C), &H85, &HF1, &HE4, &H7, &H5D, &H13, &H5F, &HC8)
IID_IFileDialog = iid
End Function
Public Function IID_IFileSaveDialog() As UUID
'{84bccd23-5fde-4cdb-aea4-af64b83d78ab}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H84BCCD23, CInt(&H5FDE), CInt(&H4CDB), &HAE, &HA4, &HAF, &H64, &HB8, &H3D, &H78, &HAB)
IID_IFileSaveDialog = iid
End Function
Public Function IID_IFileOpenDialog() As UUID
'{d57c7288-d4ad-4768-be02-9d969532d960}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD57C7288, CInt(&HD4AD), CInt(&H4768), &HBE, &H2, &H9D, &H96, &H95, &H32, &HD9, &H60)
IID_IFileOpenDialog = iid
End Function
Public Function IID_IFileDialogControlEvents() As UUID
'{36116642-D713-4b97-9B83-7484A9D00433}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H36116642, CInt(&HD713), CInt(&H4B97), &H9B, &H83, &H74, &H84, &HA9, &HD0, &H4, &H33)
IID_IFileDialogControlEvents = iid
End Function
Public Function IID_IFileDialog2() As UUID
'{61744fc7-85b5-4791-a9b0-272276309b13}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H61744FC7, CInt(&H85B5), CInt(&H4791), &HA9, &HB0, &H27, &H22, &H76, &H30, &H9B, &H13)
IID_IFileDialog2 = iid
End Function
Public Function IID_IShellMenuCallback() As UUID
'{4CA300A1-9B8D-11d1-8B22-00C04FD918D0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4CA300A1, CInt(&H9B8D), CInt(&H11D1), &H8B, &H22, &H0, &HC0, &H4F, &HD9, &H18, &HD0)
IID_IShellMenuCallback = iid
End Function
Public Function IID_IAssocHandlerInvoker() As UUID
'{92218CAB-ECAA-4335-8133-807FD234C2EE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H92218CAB, CInt(&HECAA), CInt(&H4335), &H81, &H33, &H80, &H7F, &HD2, &H34, &HC2, &HEE)
IID_IAssocHandlerInvoker = iid
End Function
Public Function IID_IAssocHandler() As UUID
'{F04061AC-1659-4a3f-A954-775AA57FC083}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF04061AC, CInt(&H1659), CInt(&H4A3F), &HA9, &H54, &H77, &H5A, &HA5, &H7F, &HC0, &H83)
IID_IAssocHandler = iid
End Function
Public Function IID_IEnumAssocHandlers() As UUID
'{973810ae-9599-4b88-9e4d-6ee98c9552da}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H973810AE, CInt(&H9599), CInt(&H4B88), &H9E, &H4D, &H6E, &HE9, &H8C, &H95, &H52, &HDA)
IID_IEnumAssocHandlers = iid
End Function
Public Function IID_INamespaceWalkCB() As UUID
'{d92995f8-cf5e-4a76-bf59-ead39ea2b97e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD92995F8, CInt(&HCF5E), CInt(&H4A76), &HBF, &H59, &HEA, &HD3, &H9E, &HA2, &HB9, &H7E)
IID_INamespaceWalkCB = iid
End Function
Public Function IID_INamespaceWalkCB2() As UUID
'{7ac7492b-c38e-438a-87db-68737844ff70}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7AC7492B, CInt(&HC38E), CInt(&H438A), &H87, &HDB, &H68, &H73, &H78, &H44, &HFF, &H70)
IID_INamespaceWalkCB2 = iid
End Function
Public Function IID_INamespaceWalk() As UUID
'{57ced8a7-3f4a-432c-9350-30f24483f74f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H57CED8A7, CInt(&H3F4A), CInt(&H432C), &H93, &H50, &H30, &HF2, &H44, &H83, &HF7, &H4F)
IID_INamespaceWalk = iid
End Function
Public Function IID_IUserNotificationCallback() As UUID
'{19108294-0441-4AFF-8013-FA0A730B0BEA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H19108294, CInt(&H441), CInt(&H4AFF), &H80, &H13, &HFA, &HA, &H73, &HB, &HB, &HEA)
IID_IUserNotificationCallback = iid
End Function
Public Function IID_IUserNotification2() As UUID
'{215913CC-57EB-4FAB-AB5A-E5FA7BEA2A6C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H215913CC, CInt(&H57EB), CInt(&H4FAB), &HAB, &H5A, &HE5, &HFA, &H7B, &HEA, &H2A, &H6C)
IID_IUserNotification2 = iid
End Function
Public Function IID_ITransferAdviseSink() As UUID
'{d594d0d8-8da7-457b-b3b4-ce5dbaac0b88}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD594D0D8, CInt(&H8DA7), CInt(&H457B), &HB3, &HB4, &HCE, &H5D, &HBA, &HAC, &HB, &H88)
IID_ITransferAdviseSink = iid
End Function
Public Function IID_IObjectWithPropertyKey() As UUID
'{fc0ca0a7-c316-4fd2-9031-3e628e6d4f23}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFC0CA0A7, CInt(&HC316), CInt(&H4FD2), &H90, &H31, &H3E, &H62, &H8E, &H6D, &H4F, &H23)
IID_IObjectWithPropertyKey = iid
End Function
Public Function IID_IPropertyChange() As UUID
'{f917bc8a-1bba-4478-a245-1bde03eb9431}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF917BC8A, CInt(&H1BBA), CInt(&H4478), &HA2, &H45, &H1B, &HDE, &H3, &HEB, &H94, &H31)
IID_IPropertyChange = iid
End Function
Public Function IID_IPropertyChangeArray() As UUID
'{380f5cad-1b5e-42f2-805d-637fd392d31e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H380F5CAD, CInt(&H1B5E), CInt(&H42F2), &H80, &H5D, &H63, &H7F, &HD3, &H92, &HD3, &H1E)
IID_IPropertyChangeArray = iid
End Function
Public Function IID_IPropertyDescription2() As UUID
'{57d2eded-5062-400e-b107-5dae79fe57a6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H57D2EDED, CInt(&H5062), CInt(&H400E), &HB1, &H7, &H5D, &HAE, &H79, &HFE, &H57, &HA6)
IID_IPropertyDescription2 = iid
End Function
Public Function IID_IPropertyDescriptionSearchInfo() As UUID
'{078f91bd-29a2-440f-924e-46a291524520}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H78F91BD, CInt(&H29A2), CInt(&H440F), &H92, &H4E, &H46, &HA2, &H91, &H52, &H45, &H20)
IID_IPropertyDescriptionSearchInfo = iid
End Function
Public Function IID_IPropertyDescriptionRelatedPropertyInfo() As UUID
'{507393f4-2a3d-4a60-b59e-d9c75716c2dd}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H507393F4, CInt(&H2A3D), CInt(&H4A60), &HB5, &H9E, &HD9, &HC7, &H57, &H16, &HC2, &HDD)
IID_IPropertyDescriptionRelatedPropertyInfo = iid
End Function
Public Function IID_IPropertyEnumType() As UUID
'{11e1fbf9-2d56-4a6b-8db3-7cd193a471f2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H11E1FBF9, CInt(&H2D56), CInt(&H4A6B), &H8D, &HB3, &H7C, &HD1, &H93, &HA4, &H71, &HF2)
IID_IPropertyEnumType = iid
End Function
Public Function IID_IPropertyEnumType2() As UUID
'{9b6e051c-5ddd-4321-9070-fe2acb55e794}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9B6E051C, CInt(&H5DDD), CInt(&H4321), &H90, &H70, &HFE, &H2A, &HCB, &H55, &HE7, &H94)
IID_IPropertyEnumType2 = iid
End Function
Public Function IID_IPropertyEnumTypeList() As UUID
'{a99400f4-3d84-4557-94ba-1242fb2cc9a6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA99400F4, CInt(&H3D84), CInt(&H4557), &H94, &HBA, &H12, &H42, &HFB, &H2C, &HC9, &HA6)
IID_IPropertyEnumTypeList = iid
End Function
Public Function IID_IPropertyStoreFactory() As UUID
'{bc110b6d-57e8-4148-a9c6-91015ab2f3a5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBC110B6D, CInt(&H57E8), CInt(&H4148), &HA9, &HC6, &H91, &H1, &H5A, &HB2, &HF3, &HA5)
IID_IPropertyStoreFactory = iid
End Function
Public Function IID_IDelayedPropertyStoreFactory() As UUID
'{40d4577f-e237-4bdb-bd69-58f089431b6a}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H40D4577F, CInt(&HE237), CInt(&H4BDB), &HBD, &H69, &H58, &HF0, &H89, &H43, &H1B, &H6A)
 IID_IDelayedPropertyStoreFactory = iid
End Function
Public Function IID_IPropertyStoreCapabilities() As UUID
'{c8e2d566-186e-4d49-bf41-6909ead56acc}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC8E2D566, CInt(&H186E), CInt(&H4D49), &HBF, &H41, &H69, &H9, &HEA, &HD5, &H6A, &HCC)
IID_IPropertyStoreCapabilities = iid
End Function
Public Function IID_IPropertyStoreCache() As UUID
'{3017056d-9a91-4e90-937d-746c72abbf4f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3017056D, CInt(&H9A91), CInt(&H4E90), &H93, &H7D, &H74, &H6C, &H72, &HAB, &HBF, &H4F)
IID_IPropertyStoreCache = iid
End Function
Public Function IID_INamedPropertyStore() As UUID
'{71604b0f-97b0-4764-8577-2f13e98a1422}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H71604B0F, CInt(&H97B0), CInt(&H4764), &H85, &H77, &H2F, &H13, &HE9, &H8A, &H14, &H22)
 IID_INamedPropertyStore = iid
End Function
Public Function IID_IPropertyDescriptionAliasInfo() As UUID
'{f67104fc-2af9-46fd-b32d-243c1404f3d1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF67104FC, CInt(&H2AF9), CInt(&H46FD), &HB3, &H2D, &H24, &H3C, &H14, &H4, &HF3, &HD1)
 IID_IPropertyDescriptionAliasInfo = iid
End Function
Public Function IID_IAutoComplete() As UUID
'{00bb2762-6a77-11d0-a535-00c04fd7d062}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBB2762, CInt(&H6A77), CInt(&H11D0), &HA5, &H35, &H0, &HC0, &H4F, &HD7, &HD0, &H62)
IID_IAutoComplete = iid
End Function
Public Function IID_IAutoComplete2() As UUID
'{EAC04BC0-3791-11d2-BB95-0060977B464C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEAC04BC0, CInt(&H3791), CInt(&H11D2), &HBB, &H95, &H0, &H60, &H97, &H7B, &H46, &H4C)
IID_IAutoComplete2 = iid
End Function
Public Function IID_IEnumACString() As UUID
'{8E74C210-CF9D-4eaf-A403-7356428F0A5A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8E74C210, CInt(&HCF9D), CInt(&H4EAF), &HA4, &H3, &H73, &H56, &H42, &H8F, &HA, &H5A)
IID_IEnumACString = iid
End Function
Public Function IID_IACList() As UUID
'{77A130B0-94FD-11D0-A544-00C04FD7d062}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H77A130B0, CInt(&H94FD), CInt(&H11D0), &HA5, &H44, &H0, &HC0, &H4F, &HD7, &HD0, &H62)
IID_IACList = iid
End Function
Public Function IID_IACList2() As UUID
'{470141a0-5186-11d2-bbb6-0060977b464c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H470141A0, CInt(&H5186), CInt(&H11D2), &HBB, &HB6, &H0, &H60, &H97, &H7B, &H46, &H4C)
IID_IACList2 = iid
End Function
Public Function IID_IBindCtx() As UUID
'{0000000e-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IBindCtx = iid
End Function
Public Function IID_IRunningObjectTable() As UUID
'{00000010-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IRunningObjectTable = iid
End Function
Public Function IID_ICatRegister() As UUID
'{0002E012-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2E012, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICatRegister = iid
End Function
Public Function IID_ICatInformation() As UUID
'{0002E013-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2E013, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICatInformation = iid
End Function
Public Function IID_ICreateTypeInfo() As UUID
'{00020405-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20405, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICreateTypeInfo = iid
End Function
Public Function IID_ICreateTypeInfo2() As UUID
'{0002040E-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2040E, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICreateTypeInfo2 = iid
End Function
Public Function IID_ICreateTypeLib() As UUID
'{00020406-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20406, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICreateTypeLib = iid
End Function
Public Function IID_ICreateTypeLib2() As UUID
'{0002040F-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2040F, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICreateTypeLib2 = iid
End Function
Public Function IID_IDocHostUIHandler() As UUID
'{bd3f23c0-d43e-11cf-893b-00aa00bdce1a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBD3F23C0, CInt(&HD43E), CInt(&H11CF), &H89, &H3B, &H0, &HAA, &H0, &HBD, &HCE, &H1A)
IID_IDocHostUIHandler = iid
End Function
Public Function IID_IDocHostUIHandler2() As UUID
'{3050f6d0-98b5-11cf-bb82-00aa00bdce0b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3050F6D0, CInt(&H98B5), CInt(&H11CF), &HBB, &H82, &H0, &HAA, &H0, &HBD, &HCE, &HB)
IID_IDocHostUIHandler2 = iid
End Function
Public Function IID_ICustomDoc() As UUID
'{3050f3f0-98b5-11cf-bb82-00aa00bdce0b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3050F3F0, CInt(&H98B5), CInt(&H11CF), &HBB, &H82, &H0, &HAA, &H0, &HBD, &HCE, &HB)
IID_ICustomDoc = iid
End Function
Public Function IID_IDocHostShowUI() As UUID
'{c4d244b0-d43e-11cf-893b-00aa00bdce1a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC4D244B0, CInt(&HD43E), CInt(&H11CF), &H89, &H3B, &H0, &HAA, &H0, &HBD, &HCE, &H1A)
IID_IDocHostShowUI = iid
End Function
Public Function IID_IAdviseSink() As UUID
'{0000010f-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10F, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IAdviseSink = iid
End Function
Public Function IID_IInputObject() As UUID
'{68284faa-6a48-11d0-8c78-00c04fd918b4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H68284FAA, CInt(&H6A48), CInt(&H11D0), &H8C, &H78, &H0, &HC0, &H4F, &HD9, &H18, &HB4)
IID_IInputObject = iid
End Function
Public Function IID_IDeskBand() As UUID
'{EB0FE172-1A3A-11D0-89B3-00A0C90A90AC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEB0FE172, CInt(&H1A3A), CInt(&H11D0), &H89, &HB3, &H0, &HA0, &HC9, &HA, &H90, &HAC)
IID_IDeskBand = iid
End Function
Public Function IID_IDockingWindow() As UUID
'{012dd920-7b26-11d0-8ca9-00a0c92dbfe8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H12DD920, CInt(&H7B26), CInt(&H11D0), &H8C, &HA9, &H0, &HA0, &HC9, &H2D, &HBF, &HE8)
IID_IDockingWindow = iid
End Function
Public Function IID_IDockingWindowSite() As UUID
'{2a342fc2-7b26-11d0-8ca9-00a0c92dbfe8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2A342FC2, CInt(&H7B26), CInt(&H11D0), &H8C, &HA9, &H0, &HA0, &HC9, &H2D, &HBF, &HE8)
IID_IDockingWindowSite = iid
End Function
Public Function IID_IDockingWindowFrame() As UUID
'{47d2657a-7b27-11d0-8ca9-00a0c92dbfe8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H47D2657A, CInt(&H7B27), CInt(&H11D0), &H8C, &HA9, &H0, &HA0, &HC9, &H2D, &HBF, &HE8)
IID_IDockingWindowFrame = iid
End Function
Public Function IID_ITrayDeskBand() As UUID
'{6D67E846-5B9C-4db8-9CBC-DDE12F4254F1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6D67E846, CInt(&H5B9C), CInt(&H4DB8), &H9C, &HBC, &HDD, &HE1, &H2F, &H42, &H54, &HF1)
IID_ITrayDeskBand = iid
End Function
Public Function IID_IBandHost() As UUID
'{B9075C7C-D48E-403f-AB99-D6C77A1084AC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9075C7C, CInt(&HD48E), CInt(&H403F), &HAB, &H99, &HD6, &HC7, &H7A, &H10, &H84, &HAC)
IID_IBandHost = iid
End Function
Public Function IID_IDeskBand2() As UUID
'{79D16DE4-ABEE-4021-8D9D-9169B261D657}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79D16DE4, CInt(&HABEE), CInt(&H4021), &H8D, &H9D, &H91, &H69, &HB2, &H61, &HD6, &H57)
IID_IDeskBand2 = iid
End Function
Public Function IID_IDeskBandInfo() As UUID
'{77E425FC-CBF9-4307-BA6A-BB5727745661}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H77E425FC, CInt(&HCBF9), CInt(&H4307), &HBA, &H6A, &HBB, &H57, &H27, &H74, &H56, &H61)
IID_IDeskBandInfo = iid
End Function
Public Function IID_IAccessibleObject() As UUID
'{95A391C5-9ED4-4c28-8401-AB9E06719E11}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H95A391C5, CInt(&H9ED4), CInt(&H4C28), &H84, &H1, &HAB, &H9E, &H6, &H71, &H9E, &H11)
IID_IAccessibleObject = iid
End Function
Public Function IID_IFolderBandPriv() As UUID
'{47c01f95-e185-412c-b5c5-4f27df965aea}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H47C01F95, CInt(&HE185), CInt(&H412C), &HB5, &HC5, &H4F, &H27, &HDF, &H96, &H5A, &HEA)
IID_IFolderBandPriv = iid
End Function
Public Function IID_IWinEventHandler() As UUID
'{47c01f95-e185-412c-b5c5-4f27df965aea}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H47C01F95, CInt(&HE185), CInt(&H412C), &HB5, &HC5, &H4F, &H27, &HDF, &H96, &H5A, &HEA)
IID_IWinEventHandler = iid
End Function
Public Function IID_INotificationCB() As UUID
'{D782CCBA-AFB0-43F1-94DB-FDA3779EACCB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD782CCBA, CInt(&HAFB0), CInt(&H43F1), &H94, &HDB, &HFD, &HA3, &H77, &H9E, &HAC, &HCB)
IID_INotificationCB = iid
End Function
Public Function IID_ITrayNotify() As UUID
'{FB852B2C-6BAD-4605-9551-F15F87830935}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFB852B2C, CInt(&H6BAD), CInt(&H4605), &H95, &H51, &HF1, &H5F, &H87, &H83, &H9, &H35)
IID_ITrayNotify = iid
End Function
Public Function IID_IPinnedList() As UUID
'{C3C6EB6D-C837-4EAE-B172-5FEC52A2A4FD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC3C6EB6D, CInt(&HC837), CInt(&H4EAE), &HB1, &H72, &H5F, &HEC, &H52, &HA2, &HA4, &HFD)
IID_IPinnedList = iid
End Function
Public Function IID_IPinnedList2() As UUID
'{BBD20037-BC0E-42F1-913F-E2936BB0EA0C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBBD20037, CInt(&HBC0E), CInt(&H42F1), &H91, &H3F, &HE2, &H93, &H6B, &HB0, &HEA, &HC)
IID_IPinnedList2 = iid
End Function
Public Function IID_IPinnedList3() As UUID
'{0DD79AE2-D156-45D4-9EEB-3B549769E940}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDD79AE2, CInt(&HD156), CInt(&H45D4), &H9E, &HEB, &H3B, &H54, &H97, &H69, &HE9, &H40)
IID_IPinnedList3 = iid
End Function
Public Function IID_ICurrentWorkingDirectory() As UUID
'{91956D21-9276-11d1-921A-006097DF5BD4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H91956D21, CInt(&H9276), CInt(&H11D1), &H92, &H1A, &H0, &H60, &H97, &HDF, &H5B, &HD4)
IID_ICurrentWorkingDirectory = iid
End Function
Public Function IID_IShellFolderBand() As UUID
'{7FE80CC8-C247-11d0-B93A-00A0C90312E1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7FE80CC8, CInt(&HC247), CInt(&H11D0), &HB9, &H3A, &H0, &HA0, &HC9, &H3, &H12, &HE1)
IID_IShellFolderBand = iid
End Function
Public Function IID_IDeskBarClient() As UUID
'{EB0FE175-1A3A-11D0-89B3-00A0C90A90AC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEB0FE175, CInt(&H1A3A), CInt(&H11D0), &H89, &HB3, &H0, &HA0, &HC9, &HA, &H90, &HAC)
IID_IDeskBarClient = iid
End Function
Public Function IID_IDeskBar() As UUID
'{EB0FE173-1A3A-11D0-89B3-00A0C90A90AC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEB0FE173, CInt(&H1A3A), CInt(&H11D0), &H89, &HB3, &H0, &HA0, &HC9, &HA, &H90, &HAC)
IID_IDeskBar = iid
End Function
Public Function IID_IHandlerInfo2() As UUID
'{31cca04c-04d3-4ea9-90de-97b15e87a532}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H31CCA04C, CInt(&H4D3), CInt(&H4EA9), &H90, &HDE, &H97, &HB1, &H5E, &H87, &HA5, &H32)
IID_IHandlerInfo2 = iid
End Function
Public Function IID_IBannerNotificationHandler() As UUID
'{8d7b2ba7-db05-46a8-823c-d2b6de08ee91}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8D7B2BA7, CInt(&HDB05), CInt(&H46A8), &H82, &H3C, &HD2, &HB6, &HDE, &H8, &HEE, &H91)
IID_IBannerNotificationHandler = iid
End Function
Public Function IID_ISortColumnArray() As UUID
'{6dfc60fb-f2e9-459b-beb5-288f1a7c7d54}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6DFC60FB, CInt(&HF2E9), CInt(&H459B), &HBE, &HB5, &H28, &H8F, &H1A, &H7C, &H7D, &H54)
IID_ISortColumnArray = iid
End Function
Public Function IID_IPropertyKeyStore() As UUID
'{75BD59AA-F23B-4963-ABA4-0B355752A91B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H75BD59AA, CInt(&HF23B), CInt(&H4963), &HAB, &HA4, &HB, &H35, &H57, &H52, &HA9, &H1B)
IID_IPropertyKeyStore = iid
End Function
Public Function IID_IInputObjectSite() As UUID
'{f1db8392-7331-11d0-8c99-00a0c92dbfe8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF1DB8392, CInt(&H7331), CInt(&H11D0), &H8C, &H99, &H0, &HA0, &HC9, &H2D, &HBF, &HE8)
IID_IInputObjectSite = iid
End Function
Public Function IID_IEnumSTATPROPSTG() As UUID
'{00000139-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H139, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumSTATPROPSTG = iid
End Function
Public Function IID_IEnumSTATPROPSETSTG() As UUID
'{0000013B-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H13B, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumSTATPROPSETSTG = iid
End Function
Public Function IID_IEnumSTATSTG() As UUID
'{0000000d-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumSTATSTG = iid
End Function
Public Function IID_IEnumSTATDATA() As UUID
'{00000105-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H105, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumSTATDATA = iid
End Function
Public Function IID_IEnumString() As UUID
'{00000101-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H101, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumString = iid
End Function
Public Function IID_IEnumMoniker() As UUID
'{00000102-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H102, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumMoniker = iid
End Function
Public Function IID_IEnumFORMATETC() As UUID
'{00000103-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H103, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumFORMATETC = iid
End Function
Public Function IID_IEnumUnknown() As UUID
'{00000100-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H100, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumUnknown = iid
End Function
Public Function IID_IEnumOLEVERB() As UUID
'{00000104-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H104, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumOLEVERB = iid
End Function
Public Function IID_IEnumGUID() As UUID
'{0002E000-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2E000, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumGUID = iid
End Function
Public Function IID_IEnumCATEGORYINFO() As UUID
'{0002E011-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2E011, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumCATEGORYINFO = iid
End Function
Public Function IID_IEnumVARIANT() As UUID
'{00020404-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20404, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumVARIANT = iid
End Function
Public Function IID_IEnumConnections() As UUID
'{B196B287-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B287, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IEnumConnections = iid
End Function
Public Function IID_IEnumConnectionPoints() As UUID
'{B196B285-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B285, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IEnumConnectionPoints = iid
End Function
Public Function IID_IErrorInfo() As UUID
'{1CF2B120-547D-101B-8E65-08002B2BD119}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1CF2B120, CInt(&H547D), CInt(&H101B), &H8E, &H65, &H8, &H0, &H2B, &H2B, &HD1, &H19)
IID_IErrorInfo = iid
End Function
Public Function IID_ICreateErrorInfo() As UUID
'{22F03340-547D-101B-8E65-08002B2BD119}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H22F03340, CInt(&H547D), CInt(&H101B), &H8E, &H65, &H8, &H0, &H2B, &H2B, &HD1, &H19)
IID_ICreateErrorInfo = iid
End Function
Public Function IID_ISupportErrorInfo() As UUID
'{DF0B3D60-548F-101B-8E65-08002B2BD119}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDF0B3D60, CInt(&H548F), CInt(&H101B), &H8E, &H65, &H8, &H0, &H2B, &H2B, &HD1, &H19)
IID_ISupportErrorInfo = iid
End Function
Public Function IID_IEmptyVolumeCacheCallBack() As UUID
'{6E793361-73C6-11D0-8469-00AA00442901}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6E793361, CInt(&H73C6), CInt(&H11D0), &H84, &H69, &H0, &HAA, &H0, &H44, &H29, &H1)
IID_IEmptyVolumeCacheCallBack = iid
End Function
Public Function IID_IEmptyVolumeCache() As UUID
'{8FCE5227-04DA-11d1-A004-00805F8ABE06}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8FCE5227, CInt(&H4DA), CInt(&H11D1), &HA0, &H4, &H0, &H80, &H5F, &H8A, &HBE, &H6)
IID_IEmptyVolumeCache = iid
End Function
Public Function IID_IEmptyVolumeCache2() As UUID
'{02b7e3ba-4db3-11d2-b2d9-00c04f8eec8c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2B7E3BA, CInt(&H4DB3), CInt(&H11D2), &HB2, &HD9, &H0, &HC0, &H4F, &H8E, &HEC, &H8C)
IID_IEmptyVolumeCache2 = iid
End Function
Public Function IID_IPublishedApp() As UUID
'{1BC752E0-9046-11D1-B8B3-006008059382}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1BC752E0, CInt(&H9046), CInt(&H11D1), &HB8, &HB3, &H0, &H60, &H8, &H5, &H93, &H82)
IID_IPublishedApp = iid
End Function
Public Function IID_IPublishedApp2() As UUID
'{12B81347-1B3A-4A04-AA61-3F768B67FD7E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H12B81347, CInt(&H1B3A), CInt(&H4A04), &HAA, &H61, &H3F, &H76, &H8B, &H67, &HFD, &H7E)
IID_IPublishedApp2 = iid
End Function
Public Function IID_IEnumPublishedApps() As UUID
'{0B124F8C-91F0-11D1-B8B5-006008059382}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB124F8C, CInt(&H91F0), CInt(&H11D1), &HB8, &HB5, &H0, &H60, &H8, &H5, &H93, &H82)
IID_IEnumPublishedApps = iid
End Function
Public Function IID_IShellBrowser() As UUID
'{000214E2-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214E2, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IShellBrowser = iid
End Function
Public Function IID_IProgressDialog() As UUID
'{EBBC7C04-315E-11d2-B62F-006097DF5BD4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEBBC7C04, CInt(&H315E), CInt(&H11D2), &HB6, &H2F, &H0, &H60, &H97, &HDF, &H5B, &HD4)
IID_IProgressDialog = iid
End Function
Public Function IID_IMoniker() As UUID
'{0000000f-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IMoniker = iid
End Function
Public Function IID_IHlink() As UUID
'{79eac9c3-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C3, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IHlink = iid
End Function
Public Function IID_IHlinkSite() As UUID
'{79eac9c2-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C2, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IHlinkSite = iid
End Function
Public Function IID_IHlinkTarget() As UUID
'{79eac9c4-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C4, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IHlinkTarget = iid
End Function
Public Function IID_IHlinkFrame() As UUID
'{79eac9c5-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C5, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IHlinkFrame = iid
End Function
Public Function IID_IEnumHLITEM() As UUID
'{79eac9c6-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C6, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IEnumHLITEM = iid
End Function
Public Function IID_IHlinkBrowseContext() As UUID
'{79eac9c7-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C7, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IHlinkBrowseContext = iid
End Function
Public Function IID_IDiscRecorder() As UUID
'{85AC9776-CA88-4cf2-894E-09598C078A41}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H85AC9776, CInt(&HCA88), CInt(&H4CF2), &H89, &H4E, &H9, &H59, &H8C, &H7, &H8A, &H41)
IID_IDiscRecorder = iid
End Function
Public Function IID_IEnumDiscRecorders() As UUID
'{9B1921E1-54AC-11d3-9144-00104BA11C5E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9B1921E1, CInt(&H54AC), CInt(&H11D3), &H91, &H44, &H0, &H10, &H4B, &HA1, &H1C, &H5E)
IID_IEnumDiscRecorders = iid
End Function
Public Function IID_IEnumDiscMasterFormats() As UUID
'{DDF445E1-54BA-11d3-9144-00104BA11C5E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDDF445E1, CInt(&H54BA), CInt(&H11D3), &H91, &H44, &H0, &H10, &H4B, &HA1, &H1C, &H5E)
IID_IEnumDiscMasterFormats = iid
End Function
Public Function IID_IRedbookDiscMaster() As UUID
'{E3BC42CD-4E5C-11D3-9144-00104BA11C5E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE3BC42CD, CInt(&H4E5C), CInt(&H11D3), &H91, &H44, &H0, &H10, &H4B, &HA1, &H1C, &H5E)
IID_IRedbookDiscMaster = iid
End Function
Public Function IID_IJolietDiscMaster() As UUID
'{E3BC42CE-4E5C-11D3-9144-00104BA11C5E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE3BC42CE, CInt(&H4E5C), CInt(&H11D3), &H91, &H44, &H0, &H10, &H4B, &HA1, &H1C, &H5E)
IID_IJolietDiscMaster = iid
End Function
Public Function IID_IDiscMasterProgressEvents() As UUID
'{EC9E51C1-4E5D-11D3-9144-00104BA11C5E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEC9E51C1, CInt(&H4E5D), CInt(&H11D3), &H91, &H44, &H0, &H10, &H4B, &HA1, &H1C, &H5E)
IID_IDiscMasterProgressEvents = iid
End Function
Public Function IID_IDiscMaster() As UUID
'{520CCA62-51A5-11D3-9144-00104BA11C5E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H520CCA62, CInt(&H51A5), CInt(&H11D3), &H91, &H44, &H0, &H10, &H4B, &HA1, &H1C, &H5E)
IID_IDiscMaster = iid
End Function
Public Function IID_IOleInPlaceUIWindow() As UUID
'{00000115-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H115, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleInPlaceUIWindow = iid
End Function
Public Function IID_IOleInPlaceActiveObject() As UUID
'{00000117-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H117, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleInPlaceActiveObject = iid
End Function
Public Function IID_IOleInPlaceSite() As UUID
'{00000119-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H119, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleInPlaceSite = iid
End Function
Public Function IID_IOleInPlaceFrame() As UUID
'{00000116-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H116, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleInPlaceFrame = iid
End Function
Public Function IID_IOleInPlaceObject() As UUID
'{00000113-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H113, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleInPlaceObject = iid
End Function
Public Function IID_IOleControlSite() As UUID
'{B196B289-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B289, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IOleControlSite = iid
End Function
Public Function IID_ILockBytes() As UUID
'{0000000a-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ILockBytes = iid
End Function
Public Function IID_IFillLockBytes() As UUID
'{99caf010-415e-11cf-8814-00aa00b569f5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H99CAF010, CInt(&H415E), CInt(&H11CF), &H88, &H14, &H0, &HAA, &H0, &HB5, &H69, &HF5)
IID_IFillLockBytes = iid
End Function
Public Function IID_IMalloc() As UUID
'{00000002-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IMalloc = iid
End Function
Public Function IID_IMarshal() As UUID
'{00000003-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IMarshal = iid
End Function
Public Function IID_IObjectSafety() As UUID
'{CB5BDC81-93C1-11cf-8F20-00805F2CD064}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCB5BDC81, CInt(&H93C1), CInt(&H11CF), &H8F, &H20, &H0, &H80, &H5F, &H2C, &HD0, &H64)
IID_IObjectSafety = iid
End Function
Public Function IID_IOleDocument() As UUID
'{b722bcc5-4e68-101b-a2bc-00aa00404770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB722BCC5, CInt(&H4E68), CInt(&H101B), &HA2, &HBC, &H0, &HAA, &H0, &H40, &H47, &H70)
IID_IOleDocument = iid
End Function
Public Function IID_IOleDocumentSite() As UUID
'{b722bcc7-4e68-101b-a2bc-00aa00404770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB722BCC7, CInt(&H4E68), CInt(&H101B), &HA2, &HBC, &H0, &HAA, &H0, &H40, &H47, &H70)
IID_IOleDocumentSite = iid
End Function
Public Function IID_IOleDocumentView() As UUID
'{b722bcc6-4e68-101b-a2bc-00aa00404770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB722BCC6, CInt(&H4E68), CInt(&H101B), &HA2, &HBC, &H0, &HAA, &H0, &H40, &H47, &H70)
IID_IOleDocumentView = iid
End Function
Public Function IID_IEnumOleDocumentViews() As UUID
'{b722bcc8-4e68-101b-a2bc-00aa00404770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB722BCC8, CInt(&H4E68), CInt(&H101B), &HA2, &HBC, &H0, &HAA, &H0, &H40, &H47, &H70)
IID_IEnumOleDocumentViews = iid
End Function
Public Function IID_IContinueCallback() As UUID
'{b722bcca-4e68-101b-a2bc-00aa00404770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB722BCCA, CInt(&H4E68), CInt(&H101B), &HA2, &HBC, &H0, &HAA, &H0, &H40, &H47, &H70)
IID_IContinueCallback = iid
End Function
Public Function IID_IPrint() As UUID
'{b722bcc9-4e68-101b-a2bc-00aa00404770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB722BCC9, CInt(&H4E68), CInt(&H101B), &HA2, &HBC, &H0, &HAA, &H0, &H40, &H47, &H70)
IID_IPrint = iid
End Function
Public Function IID_IPrintDialogCallback() As UUID
'{5852A2C3-6530-11D1-B6A3-0000F8757BF9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5852A2C3, CInt(&H6530), CInt(&H11D1), &HB6, &HA3, &H0, &H0, &HF8, &H75, &H7B, &HF9)
IID_IPrintDialogCallback = iid
End Function
Public Function IID_IPrintDialogServices() As UUID
'{509AAEDA-5639-11D1-B6A1-0000F8757BF9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H509AAEDA, CInt(&H5639), CInt(&H11D1), &HB6, &HA1, &H0, &H0, &HF8, &H75, &H7B, &HF9)
IID_IPrintDialogServices = iid
End Function
Public Function IID_IOleClientSite() As UUID
'{00000118-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H118, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleClientSite = iid
End Function
Public Function IID_IParseDisplayName() As UUID
'{0000011A-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H11A, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IParseDisplayName = iid
End Function
Public Function IID_IOleContainer() As UUID
'{0000011B-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H11B, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleContainer = iid
End Function
Public Function IID_IOleObject() As UUID
'{00000112-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H112, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleObject = iid
End Function
Public Function IID_IOleCache() As UUID
'{0000011e-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H11E, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleCache = iid
End Function
Public Function IID_IOleControl() As UUID
'{B196B288-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B288, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IOleControl = iid
End Function
Public Function IID_IOleCommandTarget() As UUID
'{b722bccb-4e68-101b-a2bc-00aa00404770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB722BCCB, CInt(&H4E68), CInt(&H101B), &HA2, &HBC, &H0, &HAA, &H0, &H40, &H47, &H70)
IID_IOleCommandTarget = iid
End Function
Public Function IID_IServiceProvider() As UUID
'{6d5140c1-7436-11ce-8034-00aa006009fa}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6D5140C1, CInt(&H7436), CInt(&H11CE), &H80, &H34, &H0, &HAA, &H0, &H60, &H9, &HFA)
IID_IServiceProvider = iid
End Function
Public Function IID_ISpecifyPropertyPages() As UUID
'{B196B28B-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B28B, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_ISpecifyPropertyPages = iid
End Function
Public Function IID_IOleWindow() As UUID
'{00000114-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H114, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleWindow = iid
End Function
Public Function IID_IObjectWithSite() As UUID
'{FC4801A3-2BA9-11CF-A229-00AA003D7352}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFC4801A3, CInt(&H2BA9), CInt(&H11CF), &HA2, &H29, &H0, &HAA, &H0, &H3D, &H73, &H52)
IID_IObjectWithSite = iid
End Function
Public Function IID_IPersist() As UUID
'{0000010c-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10C, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IPersist = iid
End Function
Public Function IID_IPersistStream() As UUID
'{00000109-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H109, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IPersistStream = iid
End Function
Public Function IID_IPersistStreamInit() As UUID
'{7FD52380-4E07-101B-AE2D-08002B2EC713}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7FD52380, CInt(&H4E07), CInt(&H101B), &HAE, &H2D, &H8, &H0, &H2B, &H2E, &HC7, &H13)
IID_IPersistStreamInit = iid
End Function
Public Function IID_IPersistFile() As UUID
'{0000010b-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10B, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IPersistFile = iid
End Function
Public Function IID_IPersistStorage() As UUID
'{0000010a-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10A, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IPersistStorage = iid
End Function
Public Function IID_IPersistPropertyBag() As UUID
'{37D84F60-42CB-11CE-8135-00AA004BB851}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H37D84F60, CInt(&H42CB), CInt(&H11CE), &H81, &H35, &H0, &HAA, &H0, &H4B, &HB8, &H51)
IID_IPersistPropertyBag = iid
End Function
Public Function IID_IPersistPropertyBag2() As UUID
'{22F55881-280B-11d0-A8A9-00A0C90C2004}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H22F55881, CInt(&H280B), CInt(&H11D0), &HA8, &HA9, &H0, &HA0, &HC9, &HC, &H20, &H4)
IID_IPersistPropertyBag2 = iid
End Function
Public Function IID_IPersistMemory() As UUID
'{BD1AE5E0-A6AE-11CE-BD37-504200C10000}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBD1AE5E0, CInt(&HA6AE), CInt(&H11CE), &HBD, &H37, &H50, &H42, &H0, &HC1, &H0, &H0)
IID_IPersistMemory = iid
End Function
Public Function IID_IPersistMoniker() As UUID
'{79eac9c9-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C9, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IPersistMoniker = iid
End Function
Public Function IID_IPerPropertyBrowsing() As UUID
'{376BD3AA-3845-101B-84ED-08002B2EC713}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H376BD3AA, CInt(&H3845), CInt(&H101B), &H84, &HED, &H8, &H0, &H2B, &H2E, &HC7, &H13)
IID_IPerPropertyBrowsing = iid
End Function
Public Function IID_IErrorLog() As UUID
'{3127CA40-446E-11CE-8135-00AA004BB851}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3127CA40, CInt(&H446E), CInt(&H11CE), &H81, &H35, &H0, &HAA, &H0, &H4B, &HB8, &H51)
IID_IErrorLog = iid
End Function
Public Function IID_IPropertyBag2() As UUID
'{22F55882-280B-11d0-A8A9-00A0C90C2004}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H22F55882, CInt(&H280B), CInt(&H11D0), &HA8, &HA9, &H0, &HA0, &HC9, &HC, &H20, &H4)
IID_IPropertyBag2 = iid
End Function
Public Function IID_IPropertyNotifySink() As UUID
'{9BFBBC02-EFF1-101A-84ED-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9BFBBC02, CInt(&HEFF1), CInt(&H101A), &H84, &HED, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IPropertyNotifySink = iid
End Function
Public Function IID_IRecordInfo() As UUID
'{0000002F-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2F, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IRecordInfo = iid
End Function
Public Function IID_IRichEditOle() As UUID
'{00020D00-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20D00, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IRichEditOle = iid
End Function
Public Function IID_IRichEditOleCallback() As UUID
'{00020D03-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20D03, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IRichEditOleCallback = iid
End Function
Public Function IID_IInternetSecurityMgrSite() As UUID
'{79eac9ed-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9ED, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetSecurityMgrSite = iid
End Function
Public Function IID_IInternetSecurityManager() As UUID
'{79eac9ee-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9EE, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetSecurityManager = iid
End Function
Public Function IID_IInternetHostSecurityManager() As UUID
'{3af280b6-cb3f-11d0-891e-00c04fb6bfc4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3AF280B6, CInt(&HCB3F), CInt(&H11D0), &H89, &H1E, &H0, &HC0, &H4F, &HB6, &HBF, &HC4)
IID_IInternetHostSecurityManager = iid
End Function
Public Function IID_IInternetZoneManager() As UUID
'{79eac9ef-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9EF, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetZoneManager = iid
End Function
Public Function IID_IInternetZoneManagerEx() As UUID
'{A4C23339-8E06-431e-9BF4-7E711C085648}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA4C23339, CInt(&H8E06), CInt(&H431E), &H9B, &HF4, &H7E, &H71, &H1C, &H8, &H56, &H48)
IID_IInternetZoneManagerEx = iid
End Function
Public Function IID_IInternetZoneManagerEx2() As UUID
'{EDC17559-DD5D-4846-8EEF-8BECBA5A4ABF}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEDC17559, CInt(&HDD5D), CInt(&H4846), &H8E, &HEF, &H8B, &HEC, &HBA, &H5A, &H4A, &HBF)
IID_IInternetZoneManagerEx2 = iid
End Function
Public Function IID_IInternetSecurityManagerEx() As UUID
'{F164EDF1-CC7C-4f0d-9A94-34222625C393}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF164EDF1, CInt(&HCC7C), CInt(&H4F0D), &H9A, &H94, &H34, &H22, &H26, &H25, &HC3, &H93)
IID_IInternetSecurityManagerEx = iid
End Function
Public Function IID_IInternetSecurityManagerEx2() As UUID
'{F1E50292-A795-4117-8E09-2B560A72AC60}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF1E50292, CInt(&HA795), CInt(&H4117), &H8E, &H9, &H2B, &H56, &HA, &H72, &HAC, &H60)
IID_IInternetSecurityManagerEx2 = iid
End Function
Public Function IID_IZoneIdentifier2() As UUID
'{EB5E760C-09EF-45C0-B510-70830CE31E6A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEB5E760C, CInt(&H9EF), CInt(&H45C0), &HB5, &H10, &H70, &H83, &HC, &HE3, &H1E, &H6A)
IID_IZoneIdentifier2 = iid
End Function
Public Function IID_IUri() As UUID
'{A39EE748-6A27-4817-A6F2-13914BEF5890}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA39EE748, CInt(&H6A27), CInt(&H4817), &HA6, &HF2, &H13, &H91, &H4B, &HEF, &H58, &H90)
IID_IUri = iid
End Function
Public Function IID_IPersistFolder() As UUID
'{000214EA-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214EA, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IPersistFolder = iid
End Function
Public Function IID_IPersistFolder2() As UUID
'{1AC3D9F0-175C-11d1-95BE-00609797EA4F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1AC3D9F0, CInt(&H175C), CInt(&H11D1), &H95, &HBE, &H0, &H60, &H97, &H97, &HEA, &H4F)
IID_IPersistFolder2 = iid
End Function
Public Function IID_IPersistFolder3() As UUID
'{CEF04FDF-FE72-11d2-87a5-00c04f6837cf}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCEF04FDF, CInt(&HFE72), CInt(&H11D2), &H87, &HA5, &H0, &HC0, &H4F, &H68, &H37, &HCF)
 IID_IPersistFolder3 = iid
End Function
Public Function IID_IPersistIDList() As UUID
'{1079acfc-29bd-11d3-8e0d-00c04f6837d5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1079ACFC, CInt(&H29BD), CInt(&H11D3), &H8E, &HD, &H0, &HC0, &H4F, &H68, &H37, &HD5)
IID_IPersistIDList = iid
End Function
Public Function IID_IShellView2() As UUID
'{88E39E80-3578-11CF-AE69-08002B2E1262}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H88E39E80, CInt(&H3578), CInt(&H11CF), &HAE, &H69, &H8, &H0, &H2B, &H2E, &H12, &H62)
IID_IShellView2 = iid
End Function
Public Function IID_IEnumIDList() As UUID
'{000214F2-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214F2, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumIDList = iid
End Function
Public Function IID_IShellIcon() As UUID
'{000214E5-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214E5, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IShellIcon = iid
End Function
Public Function IID_IShellLinkA() As UUID
'{000214EE-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214EE, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IShellLinkA = iid
End Function
Public Function IID_IShellLinkW() As UUID
'{000214F9-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214F9, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IShellLinkW = iid
End Function
Public Function IID_IActionProgressDialog() As UUID
'{49ff1172-eadc-446d-9285-156453a6431c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H49FF1172, CInt(&HEADC), CInt(&H446D), &H92, &H85, &H15, &H64, &H53, &HA6, &H43, &H1C)
IID_IActionProgressDialog = iid
End Function
Public Function IID_IHWEventHandler() As UUID
'{C1FB73D0-EC3A-4ba2-B512-8CDB9187B6D1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC1FB73D0, CInt(&HEC3A), CInt(&H4BA2), &HB5, &H12, &H8C, &HDB, &H91, &H87, &HB6, &HD1)
IID_IHWEventHandler = iid
End Function
Public Function IID_IQueryCancelAutoPlay() As UUID
'{DDEFE873-6997-4e68-BE26-39B633ADBE12}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDDEFE873, CInt(&H6997), CInt(&H4E68), &HBE, &H26, &H39, &HB6, &H33, &HAD, &HBE, &H12)
IID_IQueryCancelAutoPlay = iid
End Function
Public Function IID_IActionProgress() As UUID
'{49ff1173-eadc-446d-9285-156453a6431c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H49FF1173, CInt(&HEADC), CInt(&H446D), &H92, &H85, &H15, &H64, &H53, &HA6, &H43, &H1C)
IID_IActionProgress = iid
End Function
Public Function IID_IQueryContinue() As UUID
'{7307055c-b24a-486b-9f25-163e597a28a9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7307055C, CInt(&HB24A), CInt(&H486B), &H9F, &H25, &H16, &H3E, &H59, &H7A, &H28, &HA9)
IID_IQueryContinue = iid
End Function
Public Function IID_IUserNotification() As UUID
'{ba9711ba-5893-4787-a7e1-41277151550b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBA9711BA, CInt(&H5893), CInt(&H4787), &HA7, &HE1, &H41, &H27, &H71, &H51, &H55, &HB)
IID_IUserNotification = iid
End Function
Public Function IID_ITaskbarList() As UUID
'{56FDF342-FD6D-11d0-958A-006097C9A090}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56FDF342, CInt(&HFD6D), CInt(&H11D0), &H95, &H8A, &H0, &H60, &H97, &HC9, &HA0, &H90)
IID_ITaskbarList = iid
End Function
Public Function IID_ITaskbarList2() As UUID
'{602D4995-B13A-429b-A66E-1935E44F4317}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H602D4995, CInt(&HB13A), CInt(&H429B), &HA6, &H6E, &H19, &H35, &HE4, &H4F, &H43, &H17)
IID_ITaskbarList2 = iid
End Function
Public Function IID_IActiveDesktop() As UUID
'{F490EB00-1240-11D1-9888-006097DEACF9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF490EB00, CInt(&H1240), CInt(&H11D1), &H98, &H88, &H0, &H60, &H97, &HDE, &HAC, &HF9)
IID_IActiveDesktop = iid
End Function
Public Function IID_ICDBurn() As UUID
'{3d73a659-e5d0-4d42-afc0-5121ba425c8d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3D73A659, CInt(&HE5D0), CInt(&H4D42), &HAF, &HC0, &H51, &H21, &HBA, &H42, &H5C, &H8D)
IID_ICDBurn = iid
End Function
Public Function IID_ICDBurnExt() As UUID
'{2271dcca-74fc-4414-8fb7-c56b05ace2d7}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2271DCCA, CInt(&H74FC), CInt(&H4414), &H8F, &HB7, &HC5, &H6B, &H5, &HAC, &HE2, &HD7)
 IID_ICDBurnExt = iid
End Function
Public Function IID_IAddressBarParser() As UUID
'{C9D81948-443A-40C7-945C-5E171B8C66B4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC9D81948, CInt(&H443A), CInt(&H40C7), &H94, &H5C, &H5E, &H17, &H1B, &H8C, &H66, &HB4)
IID_IAddressBarParser = iid
End Function
Public Function IID_IWizardSite() As UUID
'{88960f5b-422f-4e7b-8013-73415381c3c3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H88960F5B, CInt(&H422F), CInt(&H4E7B), &H80, &H13, &H73, &H41, &H53, &H81, &HC3, &HC3)
IID_IWizardSite = iid
End Function
Public Function IID_IWizardExtension() As UUID
'{c02ea696-86cc-491e-9b23-74394a0444a8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC02EA696, CInt(&H86CC), CInt(&H491E), &H9B, &H23, &H74, &H39, &H4A, &H4, &H44, &HA8)
IID_IWizardExtension = iid
End Function
Public Function IID_IFolderViewHost() As UUID
'{1ea58f02-d55a-411d-b09e-9e65ac21605b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1EA58F02, CInt(&HD55A), CInt(&H411D), &HB0, &H9E, &H9E, &H65, &HAC, &H21, &H60, &H5B)
IID_IFolderViewHost = iid
End Function
Public Function IID_IExtractIconA() As UUID
'{000214EB-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214EB, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IExtractIconA = iid
End Function
Public Function IID_IExtractIconW() As UUID
'{000214FA-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214FA, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IExtractIconW = iid
End Function
Public Function IID_IShellPropSheetExt() As UUID
'{000214E9-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214E9, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IShellPropSheetExt = iid
End Function
Public Function IID_IQueryInfo() As UUID
'{00021500-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H21500, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IQueryInfo = iid
End Function
Public Function IID_ICustomizeInfoTip() As UUID
'{da22171f-70b4-43db-b38f-296741d1494c}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDA22171F, CInt(&H70B4), CInt(&H43DB), &HB3, &H8F, &H29, &H67, &H41, &HD1, &H49, &H4C)
 IID_ICustomizeInfoTip = iid
End Function
Public Function IID_IExtractImage2() As UUID
'{953BB1EE-93B4-11d1-98A3-00C04FB687DA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H953BB1EE, CInt(&H93B4), CInt(&H11D1), &H98, &HA3, &H0, &HC0, &H4F, &HB6, &H87, &HDA)
IID_IExtractImage2 = iid
End Function
Public Function IID_ICopyHookA() As UUID
'{000214EF-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214EF, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICopyHookA = iid
End Function
Public Function IID_ICopyHookW() As UUID
'{000214FC-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214FC, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICopyHookW = iid
End Function
Public Function IID_IColumnProvider() As UUID
'{E8025004-1C42-11d2-BE2C-00A0C9A83DA1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE8025004, CInt(&H1C42), CInt(&H11D2), &HBE, &H2C, &H0, &HA0, &HC9, &HA8, &H3D, &HA1)
IID_IColumnProvider = iid
End Function
Public Function IID_IURLSearchHook() As UUID
'{ac60f6a0-0fd9-11d0-99cb-00c04fd64497}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAC60F6A0, CInt(&HFD9), CInt(&H11D0), &H99, &HCB, &H0, &HC0, &H4F, &HD6, &H44, &H97)
IID_IURLSearchHook = iid
End Function
Public Function IID_ISearchContext() As UUID
'{09F656A2-41AF-480C-88F7-16CC0D164615}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9F656A2, CInt(&H41AF), CInt(&H480C), &H88, &HF7, &H16, &HCC, &HD, &H16, &H46, &H15)
IID_ISearchContext = iid
End Function
Public Function IID_IURLSearchHook2() As UUID
'{5ee44da4-6d32-46e3-86bc-07540dedd0e0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5EE44DA4, CInt(&H6D32), CInt(&H46E3), &H86, &HBC, &H7, &H54, &HD, &HED, &HD0, &HE0)
IID_IURLSearchHook2 = iid
End Function
Public Function IID_INewShortcutHookA() As UUID
'{000214e1-0000-0000-c000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214E1, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_INewShortcutHookA = iid
End Function
Public Function IID_INewShortcutHookW() As UUID
'{000214f7-0000-0000-c000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214F7, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_INewShortcutHookW = iid
End Function
Public Function IID_ILayoutStorage() As UUID
'{0e6d4d90-6738-11cf-9608-00aa00680db4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE6D4D90, CInt(&H6738), CInt(&H11CF), &H96, &H8, &H0, &HAA, &H0, &H68, &HD, &HB4)
IID_ILayoutStorage = iid
End Function
Public Function IID_ISequentialStream() As UUID
'{0c733a30-2a1c-11ce-ade5-00aa0044773d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC733A30, CInt(&H2A1C), CInt(&H11CE), &HAD, &HE5, &H0, &HAA, &H0, &H44, &H77, &H3D)
IID_ISequentialStream = iid
End Function
Public Function IID_ITaskTrigger() As UUID
'{148BD52B-A2AB-11CE-B11F-00AA00530503}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H148BD52B, CInt(&HA2AB), CInt(&H11CE), &HB1, &H1F, &H0, &HAA, &H0, &H53, &H5, &H3)
IID_ITaskTrigger = iid
End Function
Public Function IID_IScheduledWorkItem() As UUID
'{a6b952f0-a4b1-11d0-997d-00aa006887ec}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA6B952F0, CInt(&HA4B1), CInt(&H11D0), &H99, &H7D, &H0, &HAA, &H0, &H68, &H87, &HEC)
IID_IScheduledWorkItem = iid
End Function
Public Function IID_ITask() As UUID
'{148BD524-A2AB-11CE-B11F-00AA00530503}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H148BD524, CInt(&HA2AB), CInt(&H11CE), &HB1, &H1F, &H0, &HAA, &H0, &H53, &H5, &H3)
IID_ITask = iid
End Function
Public Function IID_IEnumWorkItems() As UUID
'{148BD528-A2AB-11CE-B11F-00AA00530503}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H148BD528, CInt(&HA2AB), CInt(&H11CE), &HB1, &H1F, &H0, &HAA, &H0, &H53, &H5, &H3)
IID_IEnumWorkItems = iid
End Function
Public Function IID_ISchedulingAgent() As UUID
'{148BD527-A2AB-11CE-B11F-00AA00530503}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H148BD527, CInt(&HA2AB), CInt(&H11CE), &HB1, &H1F, &H0, &HAA, &H0, &H53, &H5, &H3)
IID_ISchedulingAgent = iid
End Function
Public Function IID_IResultsFolder() As UUID
'{96E5AE6D-6AE1-4b1c-900C-C6480EAA8828}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H96E5AE6D, CInt(&H6AE1), CInt(&H4B1C), &H90, &HC, &HC6, &H48, &HE, &HAA, &H88, &H28)
 IID_IResultsFolder = iid
End Function
Public Function IID_IVirtualDesktopManager() As UUID
'{a5cd92ff-29be-454c-8d04-d82879fb3f1b}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA5CD92FF, CInt(&H29BE), CInt(&H454C), &H8D, &H4, &HD8, &H28, &H79, &HFB, &H3F, &H1B)
 IID_IVirtualDesktopManager = iid
End Function
Public Function IID_IInitializeNetworkFolder() As UUID
'{6e0f9881-42a8-4f2a-97f8-8af4e026d92d}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6E0F9881, CInt(&H42A8), CInt(&H4F2A), &H97, &HF8, &H8A, &HF4, &HE0, &H26, &HD9, &H2D)
 IID_IInitializeNetworkFolder = iid
End Function
Public Function IID_IProvideTaskPage() As UUID
'{4086658a-cbbb-11cf-b604-00c04fd8d565}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4086658A, CInt(&HCBBB), CInt(&H11CF), &HB6, &H4, &H0, &HC0, &H4F, &HD8, &HD5, &H65)
IID_IProvideTaskPage = iid
End Function
Public Function IID_ITextDocument() As UUID
'{8CC497C0-A1DF-11CE-8098-00AA0047BE5D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CC497C0, CInt(&HA1DF), CInt(&H11CE), &H80, &H98, &H0, &HAA, &H0, &H47, &HBE, &H5D)
IID_ITextDocument = iid
End Function
Public Function IID_ITextRange() As UUID
'{8CC497C2-A1DF-11CE-8098-00AA0047BE5D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CC497C2, CInt(&HA1DF), CInt(&H11CE), &H80, &H98, &H0, &HAA, &H0, &H47, &HBE, &H5D)
IID_ITextRange = iid
End Function
Public Function IID_ITextSelection() As UUID
'{8CC497C1-A1DF-11CE-8098-00AA0047BE5D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CC497C1, CInt(&HA1DF), CInt(&H11CE), &H80, &H98, &H0, &HAA, &H0, &H47, &HBE, &H5D)
IID_ITextSelection = iid
End Function
Public Function IID_ITextFont() As UUID
'{8CC497C3-A1DF-11CE-8098-00AA0047BE5D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CC497C3, CInt(&HA1DF), CInt(&H11CE), &H80, &H98, &H0, &HAA, &H0, &H47, &HBE, &H5D)
IID_ITextFont = iid
End Function
Public Function IID_ITextPara() As UUID
'{8CC497C4-A1DF-11CE-8098-00AA0047BE5D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CC497C4, CInt(&HA1DF), CInt(&H11CE), &H80, &H98, &H0, &HAA, &H0, &H47, &HBE, &H5D)
IID_ITextPara = iid
End Function
Public Function IID_ITextStoryRanges() As UUID
'{8CC497C5-A1DF-11CE-8098-00AA0047BE5D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CC497C5, CInt(&HA1DF), CInt(&H11CE), &H80, &H98, &H0, &HAA, &H0, &H47, &HBE, &H5D)
IID_ITextStoryRanges = iid
End Function
Public Function IID_ITextDocument2() As UUID
'{C241F5E0-7206-11D8-A2C7-00A0D1D6C6B3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC241F5E0, CInt(&H7206), CInt(&H11D8), &HA2, &HC7, &H0, &HA0, &HD1, &HD6, &HC6, &HB3)
IID_ITextDocument2 = iid
End Function
Public Function IID_ITextDisplays() As UUID
'{C241F5F2-7206-11D8-A2C7-00A0D1D6C6B3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC241F5F2, CInt(&H7206), CInt(&H11D8), &HA2, &HC7, &H0, &HA0, &HD1, &HD6, &HC6, &HB3)
IID_ITextDisplays = iid
End Function
Public Function IID_ITextFont2() As UUID
'{C241F5E3-7206-11D8-A2C7-00A0D1D6C6B3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC241F5E3, CInt(&H7206), CInt(&H11D8), &HA2, &HC7, &H0, &HA0, &HD1, &HD6, &HC6, &HB3)
IID_ITextFont2 = iid
End Function
Public Function IID_ITextPara2() As UUID
'{C241F5E4-7206-11D8-A2C7-00A0D1D6C6B3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC241F5E4, CInt(&H7206), CInt(&H11D8), &HA2, &HC7, &H0, &HA0, &HD1, &HD6, &HC6, &HB3)
IID_ITextPara2 = iid
End Function
Public Function IID_ITextSelection2() As UUID
'{C241F5E1-7206-11D8-A2C7-00A0D1D6C6B3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC241F5E1, CInt(&H7206), CInt(&H11D8), &HA2, &HC7, &H0, &HA0, &HD1, &HD6, &HC6, &HB3)
IID_ITextSelection2 = iid
End Function
Public Function IID_ITextRange2() As UUID
'{C241F5E2-7206-11D8-A2C7-00A0D1D6C6B3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC241F5E2, CInt(&H7206), CInt(&H11D8), &HA2, &HC7, &H0, &HA0, &HD1, &HD6, &HC6, &HB3)
IID_ITextRange2 = iid
End Function
Public Function IID_ITextRow() As UUID
'{C241F5EF-7206-11D8-A2C7-00A0D1D6C6B3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC241F5EF, CInt(&H7206), CInt(&H11D8), &HA2, &HC7, &H0, &HA0, &HD1, &HD6, &HC6, &HB3)
IID_ITextRow = iid
End Function
Public Function IID_ITextServices() As UUID
'{8D33F740-CF58-11CE-A89D-00AA006CADC5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8D33F740, CInt(&HCF58), CInt(&H11CE), &HA8, &H9D, &H0, &HAA, &H0, &H6C, &HAD, &HC5)
 IID_ITextServices = iid
End Function
Public Function IID_ITextServices2() As UUID
'{8D33F741-CF58-11CE-A89D-00AA006CADC5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8D33F741, CInt(&HCF58), CInt(&H11CE), &HA8, &H9D, &H0, &HAA, &H0, &H6C, &HAD, &HC5)
 IID_ITextServices2 = iid
End Function
Public Function IID_ITextHost() As UUID
'{13E670F4-1A5A-11CF-ABEB-00AA00B65EA1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H13E670F4, CInt(&H1A5A), CInt(&H11CF), &HAB, &HEB, &H0, &HAA, &H0, &HB6, &H5E, &HA1)
 IID_ITextHost = iid
End Function
Public Function IID_ITextHost2() As UUID
'{13E670F5-1A5A-11CF-ABEB-00AA00B65EA1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H13E670F5, CInt(&H1A5A), CInt(&H11CF), &HAB, &HEB, &H0, &HAA, &H0, &HB6, &H5E, &HA1)
 IID_ITextHost2 = iid
End Function
Public Function IID_IRicheditWindowlessAccessibility() As UUID
'{983E572D-20CD-460B-9104-83111592DD10}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H983E572D, CInt(&H20CD), CInt(&H460B), &H91, &H4, &H83, &H11, &H15, &H92, &HDD, &H10)
 IID_IRicheditWindowlessAccessibility = iid
End Function
Public Function IID_IRicheditUiaOverrides() As UUID
'{F2FB5CC0-B5A9-437F-9BA2-47632082269F}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF2FB5CC0, CInt(&HB5A9), CInt(&H437F), &H9B, &HA2, &H47, &H63, &H20, &H82, &H26, &H9F)
 IID_IRicheditUiaOverrides = iid
End Function
Public Function IID_ITypeInfo() As UUID
'{00020401-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20401, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ITypeInfo = iid
End Function
Public Function IID_ITypeLib2() As UUID
'{00020411-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20411, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ITypeLib2 = iid
End Function
Public Function IID_ITypeComp() As UUID
'{00020403-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20403, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ITypeComp = iid
End Function
Public Function IID_IProvideClassInfo() As UUID
'{B196B283-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B283, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IProvideClassInfo = iid
End Function
Public Function IID_IConnectionPointContainer() As UUID
'{B196B284-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B284, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IConnectionPointContainer = iid
End Function
Public Function IID_IConnectionPoint() As UUID
'{B196B286-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B286, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IConnectionPoint = iid
End Function
Public Function IID_IAdviseSink2() As UUID
'{00000125-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H125, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IAdviseSink2 = iid
End Function
Public Function IID_IClientSecurity() As UUID
'{0000013D-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H13D, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IClientSecurity = iid
End Function
Public Function IID_IServerSecurity() As UUID
'{0000013E-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H13E, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IServerSecurity = iid
End Function
Public Function IID_IClassActivator() As UUID
'{00000140-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H140, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IClassActivator = iid
End Function
Public Function IID_IProgressNotify() As UUID
'{a9d758a0-4617-11cf-95fc-00aa00680db4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA9D758A0, CInt(&H4617), CInt(&H11CF), &H95, &HFC, &H0, &HAA, &H0, &H68, &HD, &HB4)
IID_IProgressNotify = iid
End Function
Public Function IID_IMallocSpy() As UUID
'{0000001d-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1D, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IMallocSpy = iid
End Function
Public Function IID_IStdMarshalInfo() As UUID
'{00000018-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H18, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IStdMarshalInfo = iid
End Function
Public Function IID_IExternalConnection() As UUID
'{00000019-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H19, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IExternalConnection = iid
End Function
Public Function IID_IThumbnailExtractor() As UUID
'{969dc708-5c76-11d1-8d86-0000f804b057}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H969DC708, CInt(&H5C76), CInt(&H11D1), &H8D, &H86, &H0, &H0, &HF8, &H4, &HB0, &H57)
IID_IThumbnailExtractor = iid
End Function
Public Function IID_IHWDevice() As UUID
'{99BC7510-0A96-43fa-8BB1-C928A0302EFB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H99BC7510, CInt(&HA96), CInt(&H43FA), &H8B, &HB1, &HC9, &H28, &HA0, &H30, &H2E, &HFB)
IID_IHWDevice = iid
End Function
Public Function IID_IHWDeviceCustomProperties() As UUID
'{77D5D69C-D6CE-4026-B625-26964EEC733F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H77D5D69C, CInt(&HD6CE), CInt(&H4026), &HB6, &H25, &H26, &H96, &H4E, &HEC, &H73, &H3F)
IID_IHWDeviceCustomProperties = iid
End Function
Public Function IID_IEnumAutoplayHandler() As UUID
'{66057ABA-FFDB-4077-998E-7F131C3F8157}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H66057ABA, CInt(&HFFDB), CInt(&H4077), &H99, &H8E, &H7F, &H13, &H1C, &H3F, &H81, &H57)
IID_IEnumAutoplayHandler = iid
End Function
Public Function IID_IAutoplayHandler() As UUID
'{335E9E5D-37FC-4d73-8BA8-FD4E16B28134}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H335E9E5D, CInt(&H37FC), CInt(&H4D73), &H8B, &HA8, &HFD, &H4E, &H16, &HB2, &H81, &H34)
IID_IAutoplayHandler = iid
End Function
Public Function IID_IAutoplayHandlerProperties() As UUID
'{557730F6-41FA-4d11-B9FD-F88AB155347F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H557730F6, CInt(&H41FA), CInt(&H4D11), &HB9, &HFD, &HF8, &H8A, &HB1, &H55, &H34, &H7F)
IID_IAutoplayHandlerProperties = iid
End Function
Public Function IID_IHardwareDevicesEnum() As UUID
'{553A4A55-681C-440e-B109-597B9219CFB2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H553A4A55, CInt(&H681C), CInt(&H440E), &HB1, &H9, &H59, &H7B, &H92, &H19, &HCF, &HB2)
IID_IHardwareDevicesEnum = iid
End Function
Public Function IID_IHardwareDevicesVolumesEnum() As UUID
'{3342BDE1-50AF-4c5d-9A19-DABD01848DAE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3342BDE1, CInt(&H50AF), CInt(&H4C5D), &H9A, &H19, &HDA, &HBD, &H1, &H84, &H8D, &HAE)
IID_IHardwareDevicesVolumesEnum = iid
End Function
Public Function IID_IHardwareDevicesMountPointsEnum() As UUID
'{EE93D145-9B4E-480c-8385-1E8119A6F7B2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEE93D145, CInt(&H9B4E), CInt(&H480C), &H83, &H85, &H1E, &H81, &H19, &HA6, &HF7, &HB2)
IID_IHardwareDevicesMountPointsEnum = iid
End Function
Public Function IID_IHardwareDevices() As UUID
'{CC271F08-E1DD-49bf-87CC-CD6DCF3F3D9F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCC271F08, CInt(&HE1DD), CInt(&H49BF), &H87, &HCC, &HCD, &H6D, &HCF, &H3F, &H3D, &H9F)
IID_IHardwareDevices = iid
End Function
Public Function IID_IDispatch() As UUID
'{00020400-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20400, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IDispatch = iid
End Function
Public Function IID_IClassFactory() As UUID
'{00000001-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IClassFactory = iid
End Function
Public Function IID_IClassFactory2() As UUID
'{B196B28F-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B28F, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IClassFactory2 = iid
End Function
Public Function IID_IUniformResourceLocatorA() As UUID
'{FBF23B80-E3F0-101B-8488-00AA003E56F8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFBF23B80, CInt(&HE3F0), CInt(&H101B), &H84, &H88, &H0, &HAA, &H0, &H3E, &H56, &HF8)
IID_IUniformResourceLocatorA = iid
End Function
Public Function IID_IUniformResourceLocatorW() As UUID
'{CABB0DA0-DA57-11CF-9974-0020AFD79762}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCABB0DA0, CInt(&HDA57), CInt(&H11CF), &H99, &H74, &H0, &H20, &HAF, &HD7, &H97, &H62)
IID_IUniformResourceLocatorW = iid
End Function
Public Function IID_IEnumSTATURL() As UUID
'{3C374A42-BAE4-11CF-BF7D-00AA006946EE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3C374A42, CInt(&HBAE4), CInt(&H11CF), &HBF, &H7D, &H0, &HAA, &H0, &H69, &H46, &HEE)
IID_IEnumSTATURL = iid
End Function
Public Function IID_IUrlHistoryStg() As UUID
'{3C374A41-BAE4-11CF-BF7D-00AA006946EE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3C374A41, CInt(&HBAE4), CInt(&H11CF), &HBF, &H7D, &H0, &HAA, &H0, &H69, &H46, &HEE)
IID_IUrlHistoryStg = iid
End Function
Public Function IID_IUrlHistoryStg2() As UUID
'{AFA0DC11-C313-11d0-831A-00C04FD5AE38}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAFA0DC11, CInt(&HC313), CInt(&H11D0), &H83, &H1A, &H0, &HC0, &H4F, &HD5, &HAE, &H38)
IID_IUrlHistoryStg2 = iid
End Function
Public Function IID_IBinding() As UUID
'{79eac9c0-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C0, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IBinding = iid
End Function
Public Function IID_IBindStatusCallback() As UUID
'{79eac9c1-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C1, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IBindStatusCallback = iid
End Function
Public Function IID_IAuthenticate() As UUID
'{79eac9d0-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9D0, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IAuthenticate = iid
End Function
Public Function IID_IInternetProtocolInfo() As UUID
'{79eac9ec-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9EC, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetProtocolInfo = iid
End Function
Public Function IID_IInternetPriority() As UUID
'{79eac9eb-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9EB, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetPriority = iid
End Function
Public Function IID_IInternetSession() As UUID
'{79eac9e7-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9E7, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetSession = iid
End Function
Public Function IID_IInternetProtocolRoot() As UUID
'{79eac9e3-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9E3, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetProtocolRoot = iid
End Function
Public Function IID_IInternetProtocol() As UUID
'{79eac9e4-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9E4, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetProtocol = iid
End Function
Public Function IID_IInternetProtocolSink() As UUID
'{79eac9e5-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9E5, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetProtocolSink = iid
End Function
Public Function IID_IInternetBindInfo() As UUID
'{79eac9e1-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9E1, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetBindInfo = iid
End Function
Public Function IID_IBindProtocol() As UUID
'{79eac9cd-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9CD, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IBindProtocol = iid
End Function
Public Function IID_IHttpNegotiate() As UUID
'{79eac9d2-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9D2, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IHttpNegotiate = iid
End Function
Public Function IID_IWindowForBindingUI() As UUID
'{79eac9d5-bafa-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9D5, CInt(&HBAFA), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IWindowForBindingUI = iid
End Function
Public Function IID_IWinInetInfo() As UUID
'{79eac9d6-bafa-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9D6, CInt(&HBAFA), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IWinInetInfo = iid
End Function
Public Function IID_IWinInetHttpInfo() As UUID
'{79eac9d8-bafa-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9D8, CInt(&HBAFA), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IWinInetHttpInfo = iid
End Function
Public Function IID_IBindHost() As UUID
'{fc4801a1-2ba9-11cf-a229-00aa003d7352}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFC4801A1, CInt(&H2BA9), CInt(&H11CF), &HA2, &H29, &H0, &HAA, &H0, &H3D, &H73, &H52)
IID_IBindHost = iid
End Function
Public Function IID_IHttpNegotiate2() As UUID
'{4F9F9FCB-E0F4-48eb-B7AB-FA2EA9365CB4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4F9F9FCB, CInt(&HE0F4), CInt(&H48EB), &HB7, &HAB, &HFA, &H2E, &HA9, &H36, &H5C, &HB4)
IID_IHttpNegotiate2 = iid
End Function
Public Function IID_IHttpSecurity() As UUID
'{79eac9d7-bafa-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9D7, CInt(&HBAFA), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IHttpSecurity = iid
End Function
Public Function IID_IHttpNegotiate3() As UUID
'{57b6c80a-34c2-4602-bc26-66a02fc57153}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57B6C80A, CInt(&H34C2), CInt(&H4602), &HBC, &H26, &H66, &HA0, &H2F, &HC5, &H71, &H53)
 IID_IHttpNegotiate3 = iid
End Function
Public Function IID_IViewObject() As UUID
'{0000010D-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10D, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IViewObject = iid
End Function
Public Function IID_IViewObject2() As UUID
'{00000127-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H127, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IViewObject2 = iid
End Function
Public Function IID_IWMPRemoteMediaServices() As UUID
'{CBB92747-741F-44fe-AB5B-F1A48F3B2A59}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCBB92747, CInt(&H741F), CInt(&H44FE), &HAB, &H5B, &HF1, &HA4, &H8F, &H3B, &H2A, &H59)
IID_IWMPRemoteMediaServices = iid
End Function
Public Function IID_IWMPPluginUI() As UUID
'{4C5E8F9F-AD3E-4bf9-9753-FCD30D6D38DD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4C5E8F9F, CInt(&HAD3E), CInt(&H4BF9), &H97, &H53, &HFC, &HD3, &HD, &H6D, &H38, &HDD)
IID_IWMPPluginUI = iid
End Function
'IID_IShellView =    { 0x000214E3, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
Public Function IID_IShellView() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214E3, 0, 0)
 IID_IShellView = iid
End Function
Public Function IID_IFolderView() As UUID
'{cde725b0-ccc9-4519-917e-325d72fab4ce}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCDE725B0, CInt(&HCCC9), CInt(&H4519), &H91, &H7E, &H32, &H5D, &H72, &HFA, &HB4, &HCE)
 IID_IFolderView = iid
End Function
Public Function IID_IFolderView2() As UUID
'{1af3a467-214f-4298-908e-06b03e0b39f9}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1AF3A467, CInt(&H214F), CInt(&H4298), &H90, &H8E, &H6, &HB0, &H3E, &HB, &H39, &HF9)
 IID_IFolderView2 = iid
End Function

' Returns the IShellFolder interface ID, {000214E6-0000-0000-C000-000000046}
Public Function IID_IShellFolder() As UUID
  Static iid As UUID
  If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214E6, 0, 0)
  IID_IShellFolder = iid
End Function

' Returns the IShellDetails interface ID,

Public Function IID_IShellDetails() As UUID
'{000214EC-0000-0000-C000-000000000046}
  Static iid As UUID
  If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214EC, 0, 0)
  IID_IShellDetails = iid
End Function
Public Function IID_IExtractImage() As UUID
'{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBB2E617C, CInt(&H920), CInt(&H11D1), &H9A, &HB, &H0, &HC0, &H4F, &HC2, &HD6, &HC1)
  IID_IExtractImage = iid

End Function
Public Function IID_IShellFolder2() As UUID
'{93F2F68C-1D1B-11D3-A30E-00C04F79ABD1}
    Static iid As UUID
    If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H93F2F68C, CInt(&H1D1B), CInt(&H11D3), &HA3, &HE, 0, &HC0, &H4F, &H79, &HAB, &HD1)
    IID_IShellFolder2 = iid
End Function

Public Function IID_IStorage() As UUID
'({0000000B-0000-0000-C000-000000000046})
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &HB, 0, 0)
 IID_IStorage = iid
End Function
Public Function IID_IRootStorage() As UUID
'({00000012-0000-0000-C000-000000000046})
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H12, 0, 0)
 IID_IRootStorage = iid
End Function
Public Function IID_IPropertyStorage() As UUID
'({00000138-0000-0000-C000-000000000046})
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H12, 0, 0)
 IID_IPropertyStorage = iid
End Function
Public Function IID_IShellItem() As UUID
Static iid As UUID
If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43826D1E, CInt(&HE718), CInt(&H42EE), &HBC, &H55, &HA1, &HE2, &H61, &HC3, &H7B, &HFE)
IID_IShellItem = iid
End Function
Public Function IID_IShellItem2() As UUID
'7e9fb0d3-919f-4307-ab2e-9b1860310c93
Static iid As UUID
If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7E9FB0D3, CInt(&H919F), CInt(&H4307), &HAB, &H2E, &H9B, &H18, &H60, &H31, &HC, &H93)
IID_IShellItem2 = iid
End Function
Public Function IID_IEnumShellItems() As UUID
'{70629033-e363-4a28-a567-0db78006e6d7}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H70629033, CInt(&HE363), CInt(&H4A28), &HA5, &H67, &HD, &HB7, &H80, &H6, &HE6, &HD7)
 IID_IEnumShellItems = iid
End Function
Public Function IID_IShellLibrary() As UUID
'{0x11a66efa, 0x382e, 0x451a, {0x92, 0x34, 0x1e, 0xe, 0x12, 0xef, 0x30, 0x85}}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H11A66EFA, CInt(&H382E), CInt(&H451A), &H92, &H34, &H1E, &HE, &H12, &HEF, &H30, &H85)
  IID_IShellLibrary = iid

End Function
Public Function IID_IShellItemArray() As UUID
'{b63ea76d-1f85-456f-a19c-48159efa858b}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB63EA76D, CInt(&H1F85), CInt(&H456F), &HA1, &H9C, &H48, &H15, &H9E, &HFA, &H85, &H8B)
  IID_IShellItemArray = iid

End Function
Public Function IID_IObjectArray() As UUID
'0x92ca9dcd, 0x5622, 0x4bba, 0xa8,0x05, 0x5e,0x9f,0x54,0x1b,0xd8,0xc9
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H92CA9DCD, CInt(&H5622), CInt(&H4BBA), &HA8, &H5, &H5E, &H9F, &H54, &H1B, &HD8, &HC9)
  IID_IObjectArray = iid

End Function
Public Function IID_IShellItemImageFactory() As UUID
'{BCC18B79-BA16-442F-80C4-8A59C30C463B}
Static iid As UUID
If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBCC18B79, CInt(&HBA16), CInt(&H442F), &H80, &HC4, &H8A, &H59, &HC3, &HC, &H46, &H3B)
IID_IShellItemImageFactory = iid
End Function
Public Function IID_IOleLink() As UUID
'{0000011d-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H11D, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IOleLink = iid
End Function
Public Function IID_IPropertySetStorage() As UUID
'{0000013A-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H13A, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IPropertySetStorage = iid
End Function
Public Function IID_ICondition() As UUID
'{0FC988D4-C935-4b97-A973-46282EA175C8}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFC988D4, CInt(&HC935), CInt(&H4B97), &HA9, &H73, &H46, &H28, &H2E, &HA1, &H75, &HC8)
 IID_ICondition = iid
End Function

Public Function IID_IDataObject() As UUID
'0000010e-0000-0000-C000-000000000046
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H10E, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
  IID_IDataObject = iid

End Function

Public Function IID_IFileDialogCustomize() As UUID
'IID_IFileDialogCustomize "{8016b7b3-3d49-4504-a0aa-2a37494e606f}"
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8016B7B3, CInt(&H3D49), CInt(&H4504), &HA0, &HAA, &H2A, &H37, &H49, &H4E, &H60, &H6F)
  IID_IFileDialogCustomize = iid

End Function
Public Function IID_IShellMenu() As UUID
'{0x1FEAEBFA,0x3C7A,0x4BB6,{0xB0,0xD2,0xF1,0xB8,0x1B,0x8F,0x27,0xED}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1FEAEBFA, CInt(&H3C7A), CInt(&H4BB6), &HB0, &HD2, &HF1, &HB8, &H1B, &H8F, &H27, &HED)
  IID_IShellMenu = iid
  
End Function
Public Function IID_IPropertyDescriptionList() As UUID
'IID_IPropertyDescriptionList, 0x1f9fc1d0, 0xc39b, 0x4b26, 0x81,0x7f, 0x01,0x19,0x67,0xd3,0x44,0x0e
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1F9FC1D0, CInt(&HC39B), CInt(&H4B26), &H81, &H7F, &H1, &H19, &H67, &HD3, &H44, &HE)
  IID_IPropertyDescriptionList = iid

End Function

Public Function IID_IPropertyDescription() As UUID
'(IID_IPropertyDescription, 0x6f79d558, 0x3e96, 0x4549, 0xa1,0xd1, 0x7d,0x75,0xd2,0x28,0x88,0x14
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6F79D558, CInt(&H3E96), CInt(&H4549), &HA1, &HD1, &H7D, &H75, &HD2, &H28, &H88, &H14)
  IID_IPropertyDescription = iid
  
End Function

Public Function IID_IPropertyStore() As UUID
'DEFINE_GUID(IID_IPropertyStore,0x886d8eeb, 0x8cf2, 0x4446, 0x8d,0x02,0xcd,0xba,0x1d,0xbd,0xcf,0x99);
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H886D8EEB, CInt(&H8CF2), CInt(&H4446), &H8D, &H2, &HCD, &HBA, &H1D, &HBD, &HCF, &H99)
  IID_IPropertyStore = iid
  
End Function
Public Function IID_IPropertySystem() As UUID
'IID_IPropertySystem, 0xca724e8a, 0xc3e6, 0x442b, 0x88,0xa4, 0x6f,0xb0,0xdb,0x80,0x35,0xa3
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCA724E8A, CInt(&HC3E6), CInt(&H442B), &H88, &HA4, &H6F, &HB0, &HDB, &H80, &H35, &HA3)
  IID_IPropertySystem = iid
  
End Function
Public Function IID_IPersistSerializedPropStorage() As UUID
'{e318ad57-0aa0-450f-aca5-6fab7103d917}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE318AD57, CInt(&HAA0), CInt(&H450F), &HAC, &HA5, &H6F, &HAB, &H71, &H3, &HD9, &H17)
IID_IPersistSerializedPropStorage = iid
End Function
Public Function IID_IPersistSerializedPropStorage2() As UUID
'{77effa68-4f98-4366-ba72-573b3d880571}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H77EFFA68, CInt(&H4F98), CInt(&H4366), &HBA, &H72, &H57, &H3B, &H3D, &H88, &H5, &H71)
IID_IPersistSerializedPropStorage2 = iid
End Function
Public Function IID_IPropertySystemChangeNotify() As UUID
'{fa955fd9-38be-4879-a6ce-824cf52d609f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA955FD9, CInt(&H38BE), CInt(&H4879), &HA6, &HCE, &H82, &H4C, &HF5, &H2D, &H60, &H9F)
IID_IPropertySystemChangeNotify = iid
End Function
Public Function IID_ISyncMgrHandlerCollection() As UUID
'{a7f337a3-d20b-45cb-9ed7-87d094ca5045}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA7F337A3, CInt(&HD20B), CInt(&H45CB), &H9E, &HD7, &H87, &HD0, &H94, &HCA, &H50, &H45)
IID_ISyncMgrHandlerCollection = iid
End Function
Public Function IID_ISyncMgrHandler() As UUID
'{04ec2e43-ac77-49f9-9b98-0307ef7a72a2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4EC2E43, CInt(&HAC77), CInt(&H49F9), &H9B, &H98, &H3, &H7, &HEF, &H7A, &H72, &HA2)
IID_ISyncMgrHandler = iid
End Function
Public Function IID_ISyncMgrHandlerInfo() As UUID
'{4ff1d798-ecf7-4524-aa81-1e362a0aef3a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4FF1D798, CInt(&HECF7), CInt(&H4524), &HAA, &H81, &H1E, &H36, &H2A, &HA, &HEF, &H3A)
IID_ISyncMgrHandlerInfo = iid
End Function
Public Function IID_ISyncMgrSyncItemContainer() As UUID
'{90701133-be32-4129-a65c-99e616cafff4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H90701133, CInt(&HBE32), CInt(&H4129), &HA6, &H5C, &H99, &HE6, &H16, &HCA, &HFF, &HF4)
IID_ISyncMgrSyncItemContainer = iid
End Function
Public Function IID_ISyncMgrSyncItem() As UUID
'{b20b24ce-2593-4f04-bd8b-7ad6c45051cd}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB20B24CE, CInt(&H2593), CInt(&H4F04), &HBD, &H8B, &H7A, &HD6, &HC4, &H50, &H51, &HCD)
IID_ISyncMgrSyncItem = iid
End Function
Public Function IID_ISyncMgrSyncItemInfo() As UUID
'{e7fd9502-be0c-4464-90a1-2b5277031232}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE7FD9502, CInt(&HBE0C), CInt(&H4464), &H90, &HA1, &H2B, &H52, &H77, &H3, &H12, &H32)
IID_ISyncMgrSyncItemInfo = iid
End Function
Public Function IID_IEnumSyncMgrSyncItems() As UUID
'{54b3abf3-f085-4181-b546-e29c403c726b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H54B3ABF3, CInt(&HF085), CInt(&H4181), &HB5, &H46, &HE2, &H9C, &H40, &H3C, &H72, &H6B)
IID_IEnumSyncMgrSyncItems = iid
End Function
Public Function IID_ISyncMgrSessionCreator() As UUID
'{17f48517-f305-4321-a08d-b25a834918fd}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H17F48517, CInt(&HF305), CInt(&H4321), &HA0, &H8D, &HB2, &H5A, &H83, &H49, &H18, &HFD)
IID_ISyncMgrSessionCreator = iid
End Function
Public Function IID_ISyncMgrSyncCallback() As UUID
'{884ccd87-b139-4937-a4ba-4f8e19513fbe}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H884CCD87, CInt(&HB139), CInt(&H4937), &HA4, &HBA, &H4F, &H8E, &H19, &H51, &H3F, &HBE)
IID_ISyncMgrSyncCallback = iid
End Function
Public Function IID_ISyncMgrUIOperation() As UUID
'{fc7cfa47-dfe1-45b5-a049-8cfd82bec271}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFC7CFA47, CInt(&HDFE1), CInt(&H45B5), &HA0, &H49, &H8C, &HFD, &H82, &HBE, &HC2, &H71)
IID_ISyncMgrUIOperation = iid
End Function
Public Function IID_ISyncMgrEventLinkUIOperation() As UUID
'{64522e52-848b-4015-89ce-5a36f00b94ff}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H64522E52, CInt(&H848B), CInt(&H4015), &H89, &HCE, &H5A, &H36, &HF0, &HB, &H94, &HFF)
IID_ISyncMgrEventLinkUIOperation = iid
End Function
Public Function IID_ISyncMgrScheduleWizardUIOperation() As UUID
'{459a6c84-21d2-4ddc-8a53-f023a46066f2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H459A6C84, CInt(&H21D2), CInt(&H4DDC), &H8A, &H53, &HF0, &H23, &HA4, &H60, &H66, &HF2)
IID_ISyncMgrScheduleWizardUIOperation = iid
End Function
Public Function IID_ISyncMgrSyncResult() As UUID
'{2b90f17e-5a3e-4b33-bb7f-1bc48056b94d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2B90F17E, CInt(&H5A3E), CInt(&H4B33), &HBB, &H7F, &H1B, &HC4, &H80, &H56, &HB9, &H4D)
IID_ISyncMgrSyncResult = iid
End Function
Public Function IID_ISyncMgrControl() As UUID
'{9B63616C-36B2-46BC-959F-C1593952D19B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9B63616C, CInt(&H36B2), CInt(&H46BC), &H95, &H9F, &HC1, &H59, &H39, &H52, &HD1, &H9B)
IID_ISyncMgrControl = iid
End Function
Public Function IID_ISyncMgrEventStore() As UUID
'{37e412f9-016e-44c2-81ff-db3add774266}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H37E412F9, CInt(&H16E), CInt(&H44C2), &H81, &HFF, &HDB, &H3A, &HDD, &H77, &H42, &H66)
IID_ISyncMgrEventStore = iid
End Function
Public Function IID_ISyncMgrEvent() As UUID
'{fee0ef8b-46bd-4db4-b7e6-ff2c687313bc}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFEE0EF8B, CInt(&H46BD), CInt(&H4DB4), &HB7, &HE6, &HFF, &H2C, &H68, &H73, &H13, &HBC)
IID_ISyncMgrEvent = iid
End Function
Public Function IID_IEnumSyncMgrEvents() As UUID
'{c81a1d4e-8cf7-4683-80e0-bcae88d677b6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC81A1D4E, CInt(&H8CF7), CInt(&H4683), &H80, &HE0, &HBC, &HAE, &H88, &HD6, &H77, &HB6)
IID_IEnumSyncMgrEvents = iid
End Function
Public Function IID_ISyncMgrConflictStore() As UUID
'{cf8fc579-c396-4774-85f1-d908a831156e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCF8FC579, CInt(&HC396), CInt(&H4774), &H85, &HF1, &HD9, &H8, &HA8, &H31, &H15, &H6E)
IID_ISyncMgrConflictStore = iid
End Function
Public Function IID_IEnumSyncMgrConflict() As UUID
'{82705914-dda3-4893-ba99-49de6c8c8036}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H82705914, CInt(&HDDA3), CInt(&H4893), &HBA, &H99, &H49, &HDE, &H6C, &H8C, &H80, &H36)
IID_IEnumSyncMgrConflict = iid
End Function
Public Function IID_ISyncMgrConflict() As UUID
'{9c204249-c443-4ba4-85ed-c972681db137}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9C204249, CInt(&HC443), CInt(&H4BA4), &H85, &HED, &HC9, &H72, &H68, &H1D, &HB1, &H37)
IID_ISyncMgrConflict = iid
End Function
Public Function IID_ISyncMgrConflictPresenter() As UUID
'{0b4f5353-fd2b-42cd-8763-4779f2d508a3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB4F5353, CInt(&HFD2B), CInt(&H42CD), &H87, &H63, &H47, &H79, &HF2, &HD5, &H8, &HA3)
IID_ISyncMgrConflictPresenter = iid
End Function
Public Function IID_ISyncMgrConflictResolveInfo() As UUID
'{c405a219-25a2-442e-8743-b845a2cee93f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC405A219, CInt(&H25A2), CInt(&H442E), &H87, &H43, &HB8, &H45, &HA2, &HCE, &HE9, &H3F)
IID_ISyncMgrConflictResolveInfo = iid
End Function
Public Function IID_ISyncMgrConflictFolder() As UUID
'{59287f5e-bc81-4fca-a7f1-e5a8ecdb1d69}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H59287F5E, CInt(&HBC81), CInt(&H4FCA), &HA7, &HF1, &HE5, &HA8, &HEC, &HDB, &H1D, &H69)
IID_ISyncMgrConflictFolder = iid
End Function
Public Function IID_ISyncMgrConflictItems() As UUID
'{9C7EAD52-8023-4936-A4DB-D2A9A99E436A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9C7EAD52, CInt(&H8023), CInt(&H4936), &HA4, &HDB, &HD2, &HA9, &HA9, &H9E, &H43, &H6A)
IID_ISyncMgrConflictItems = iid
End Function
Public Function IID_ISyncMgrConflictResolutionItems() As UUID
'{458725B9-129D-4135-A998-9CEAFEC27007}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H458725B9, CInt(&H129D), CInt(&H4135), &HA9, &H98, &H9C, &HEA, &HFE, &HC2, &H70, &H7)
IID_ISyncMgrConflictResolutionItems = iid
End Function
Public Function IID_ITransferConfirmation() As UUID
'{14cc750c-7b0b-43dc-910e-b687f84e7c3b}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H14CC750C, CInt(&H7B0B), CInt(&H43DC), &H91, &HE, &HB6, &H87, &HF8, &H4E, &H7C, &H3B)
 IID_ITransferConfirmation = iid
End Function
Public Function IID_IDropTarget() As UUID
'{00000122-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H122, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IDropTarget = iid
End Function
Public Function IID_IDropSource() As UUID
'{00000121-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H121, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IDropSource = iid
End Function
Public Function IID_IDragSourceHelper() As UUID
'{de5bf786-477a-11d2-839d-00c04fd918d0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDE5BF786, CInt(&H477A), CInt(&H11D2), &H83, &H9D, &H0, &HC0, &H4F, &HD9, &H18, &HD0)
  IID_IDragSourceHelper = iid
  
End Function

Public Function IID_IDragSourceHelper2() As UUID
'{83E07D0D-0C5F-4163-BF1A-60B274051E40}"
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H83E07D0D, CInt(&HC5F), CInt(&H4163), &HBF, &H1A, &H60, &HB2, &H74, &H5, &H1E, &H40)
  IID_IDragSourceHelper2 = iid
  
End Function
Public Function IID_IDropTargetHelper() As UUID
'{4657278B-411B-11D2-839A-00C04FD918D0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4657278B, CInt(&H411B), CInt(&H11D2), &H83, &H9A, &H0, &HC0, &H4F, &HD9, &H18, &HD0)
 IID_IDropTargetHelper = iid
End Function

Public Function CLSID_QueryAssociations() As UUID
'{a07034fd-6caa-4954-ac3f-97a27216f98a}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA07034FD, CInt(&H6CAA), CInt(&H4954), &HAC, &H3F, &H97, &HA2, &H72, &H16, &HF9, &H8A)
 CLSID_QueryAssociations = iid
End Function
Public Function CLSID_DiskQuotaControl() As UUID
'{7988B571-EC89-11cf-9C00-00AA00A14F56}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7988B571, CInt(&HEC89), CInt(&H11CF), &H9C, &H0, &H0, &HAA, &H0, &HA1, &H4F, &H56)
 CLSID_DiskQuotaControl = iid
End Function
Public Function CLSID_ImageList() As UUID
'{7C476BA2-02B1-48f4-8048-B24619DDC058}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7C476BA2, CInt(&H2B1), CInt(&H48F4), &H80, &H48, &HB2, &H46, &H19, &HDD, &HC0, &H58)
 CLSID_ImageList = iid
End Function

Public Function IID_IQueryAssociations() As UUID
'{c46ca590-3c3f-11d2-bee6-0000f805ca57}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC46CA590, CInt(&H3C3F), CInt(&H11D2), &HBE, &HE6, &H0, &H0, &HF8, &H5, &HCA, &H57)
 IID_IQueryAssociations = iid
End Function
Public Function IID_IEnumShellReminder() As UUID
'{6c6d9735-2d86-40e1-b348-08706b9908c0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6C6D9735, CInt(&H2D86), CInt(&H40E1), &HB3, &H48, &H8, &H70, &H6B, &H99, &H8, &HC0)
IID_IEnumShellReminder = iid
End Function
Public Function IID_IShellReminderManager() As UUID
'{968edb91-8a70-4930-8332-5f15838a64f9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H968EDB91, CInt(&H8A70), CInt(&H4930), &H83, &H32, &H5F, &H15, &H83, &H8A, &H64, &HF9)
IID_IShellReminderManager = iid
End Function
Public Function IID_IACLCustomMRU() As UUID
'{F729FC5E-8769-4f3e-BDB2-D7B50FD2275B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF729FC5E, CInt(&H8769), CInt(&H4F3E), &HBD, &HB2, &HD7, &HB5, &HF, &HD2, &H27, &H5B)
IID_IACLCustomMRU = iid
End Function
Public Function IID_IAssociationElement() As UUID
'{e58b1abf-9596-4dba-8997-89dcdef46992}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE58B1ABF, CInt(&H9596), CInt(&H4DBA), &H89, &H97, &H89, &HDC, &HDE, &HF4, &H69, &H92)
IID_IAssociationElement = iid
End Function
Public Function IID_IEnumAssociationElements() As UUID
'{a6b0fb57-7523-4439-9425-ebe99823b828}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA6B0FB57, CInt(&H7523), CInt(&H4439), &H94, &H25, &HEB, &HE9, &H98, &H23, &HB8, &H28)
IID_IEnumAssociationElements = iid
End Function
Public Function IID_IAssociationArrayInitialize() As UUID
'{ee9165bf-a4d9-474b-8236-6735cb7e28b6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEE9165BF, CInt(&HA4D9), CInt(&H474B), &H82, &H36, &H67, &H35, &HCB, &H7E, &H28, &HB6)
IID_IAssociationArrayInitialize = iid
End Function
Public Function IID_IAssociationArray() As UUID
'{3b877e3c-67de-4f9a-b29b-17d0a1521c6a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3B877E3C, CInt(&H67DE), CInt(&H4F9A), &HB2, &H9B, &H17, &HD0, &HA1, &H52, &H1C, &H6A)
IID_IAssociationArray = iid
End Function
Public Function IID_IFadeTask() As UUID
'{fadb55b4-d382-4fc4-81d7-abb325c7f12a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFADB55B4, CInt(&HD382), CInt(&H4FC4), &H81, &HD7, &HAB, &HB3, &H25, &HC7, &HF1, &H2A)
IID_IFadeTask = iid
End Function
Public Function IID_IPreviewHandler() As UUID
'{8895b1c6-b41f-4c1c-a562-0d564250836f}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8895B1C6, CInt(&HB41F), CInt(&H4C1C), &HA5, &H62, &HD, &H56, &H42, &H50, &H83, &H6F)
 IID_IPreviewHandler = iid
End Function
Public Function IID_IPreviewHandlerVisuals() As UUID
'{196bf9a5-b346-4ef0-aa1e-5dcdb76768b1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H196BF9A5, CInt(&HB346), CInt(&H4EF0), &HAA, &H1E, &H5D, &HCD, &HB7, &H67, &H68, &HB1)
 IID_IPreviewHandlerVisuals = iid
End Function
Public Function IID_IInitializeWithStream() As UUID
'{b824b49d-22ac-4161-ac8a-9916e8fa3f7f}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB824B49D, CInt(&H22AC), CInt(&H4161), &HAC, &H8A, &H99, &H16, &HE8, &HFA, &H3F, &H7F)
 IID_IInitializeWithStream = iid
End Function
Public Function IID_IInitializeWithFile() As UUID
'{b7d14566-0509-4cce-a71f-0a554233bd9b}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB7D14566, CInt(&H509), CInt(&H4CCE), &HA7, &H1F, &HA, &H55, &H42, &H33, &HBD, &H9B)
 IID_IInitializeWithFile = iid
End Function
Public Function IID_IInitializeWithItem() As UUID
'{7f73be3f-fb79-493c-a6c7-7ee14e245841}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7F73BE3F, CInt(&HFB79), CInt(&H493C), &HA6, &HC7, &H7E, &HE1, &H4E, &H24, &H58, &H41)
 IID_IInitializeWithItem = iid
End Function
Public Function IID_IInitializeWithPropertyStore() As UUID
'{C3E12EB5-7D8D-44f8-B6DD-0E77B34D6DE4}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC3E12EB5, CInt(&H7D8D), CInt(&H44F8), &HB6, &HDD, &HE, &H77, &HB3, &H4D, &H6D, &HE4)
 IID_IInitializeWithPropertyStore = iid
End Function
Public Function IID_IInitializeWithWindow() As UUID
'{3E68D4BD-7135-4D10-8018-9FB6D9F33FA1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3E68D4BD, CInt(&H7135), CInt(&H4D10), &H80, &H18, &H9F, &HB6, &HD9, &HF3, &H3F, &HA1)
 IID_IInitializeWithWindow = iid
End Function
Public Function IID_ICreateObject() As UUID
'{75121952-e0d0-43e5-9380-1d80483acf72}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H75121952, CInt(&HE0D0), CInt(&H43E5), &H93, &H80, &H1D, &H80, &H48, &H3A, &HCF, &H72)
 IID_ICreateObject = iid
End Function

Public Function IID_IPropertyBag() As UUID
'{55272A00-42CB-11CE-8135-00AA004BB851}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H55272A00, CInt(&H42CB), CInt(&H11CE), &H81, &H35, &H0, &HAA, &H0, &H4B, &HB8, &H51)
 IID_IPropertyBag = iid
End Function

Public Function IID_IImageList() As UUID
'{46EB5926-582E-4017-9FDF-E8998DAA0950}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H46EB5926, CInt(&H582E), CInt(&H4017), &H9F, &HDF, &HE8, &H99, &H8D, &HAA, &H9, &H50)
 IID_IImageList = iid
End Function
Public Function IID_IImageList2() As UUID
'{192b9d83-50fc-457b-90a0-2b82a8b5dae1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H192B9D83, CInt(&H50FC), CInt(&H457B), &H90, &HA0, &H2B, &H82, &HA8, &HB5, &HDA, &HE1)
 IID_IImageList2 = iid
End Function
Public Function IID_IContextMenu() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214E4, 0, 0)
 IID_IContextMenu = iid
End Function
Public Function IID_IContextMenu2() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214F4, 0, 0)
 IID_IContextMenu2 = iid
End Function
Public Function IID_IContextMenu3() As UUID
'{BCFCE0A0-EC17-11d0-8D10-00A0C90F2719}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBCFCE0A0, CInt(&HEC17), CInt(&H11D0), &H8D, &H10, &H0, &HA0, &HC9, &HF, &H27, &H19)
 IID_IContextMenu3 = iid
End Function
Public Function IID_IContextMenuCB() As UUID
'{3409E930-5A39-11d1-83FA-00A0C90DC849}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3409E930, CInt(&H5A39), CInt(&H11D1), &H83, &HFA, &H0, &HA0, &HC9, &HD, &HC8, &H49)
 IID_IContextMenuCB = iid
End Function
Public Function IID_IContextMenuSite() As UUID
'{0811AEBE-0B87-4C54-9E72-548CF649016B}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H811AEBE, CInt(&HB87), CInt(&H4C54), &H9E, &H72, &H54, &H8C, &HF6, &H49, &H1, &H6B)
 IID_IContextMenuSite = iid
End Function
Public Function IID_IHomeGroup() As UUID
'{7a3bd1d9-35a9-4fb3-a467-f48cac35e2d0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7A3BD1D9, CInt(&H35A9), CInt(&H4FB3), &HA4, &H67, &HF4, &H8C, &HAC, &H35, &HE2, &HD0)
 IID_IHomeGroup = iid
End Function
Public Function IID_ICallQI() As UUID
'{9fb58518-92ec-4bf6-bc61-ff4e59df7369}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9FB58518, CInt(&H92EC), CInt(&H4BF6), &HBC, &H61, &HFF, &H4E, &H59, &HDF, &H73, &H69)
 IID_ICallQI = iid
End Function
Public Function IID_IMultiQI() As UUID
'{00000020-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H20, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IMultiQI = iid
End Function
Public Function IID_ICallAddRelease() As UUID
'{9fb58519-92ec-4bf6-bc61-ff4e59df7369}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9FB58519, CInt(&H92EC), CInt(&H4BF6), &HBC, &H61, &HFF, &H4E, &H59, &HDF, &H73, &H69)
 IID_ICallAddRelease = iid
End Function
Public Function IID_ICallGION() As UUID
'{9fb58520-92ec-4bf6-bc61-ff4e59df7369}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9FB58520, CInt(&H92EC), CInt(&H4BF6), &HBC, &H61, &HFF, &H4E, &H59, &HDF, &H73, &H69)
 IID_ICallGION = iid
End Function
Public Function IID_ICallInvoke() As UUID
'{9fb58521-92ec-4bf6-bc61-ff4e59df7369}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9FB58521, CInt(&H92EC), CInt(&H4BF6), &HBC, &H61, &HFF, &H4E, &H59, &HDF, &H73, &H69)
 IID_ICallInvoke = iid
End Function
Public Function IID_IDefaultExtractIconInit() As UUID
'{41ded17d-d6b3-4261-997d-88c60e4b1d58}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H41DED17D, CInt(&HD6B3), CInt(&H4261), &H99, &H7D, &H88, &HC6, &HE, &H4B, &H1D, &H58)
 IID_IDefaultExtractIconInit = iid
End Function
Public Function IID_IExecuteCommand() As UUID
'{7F9185B0-CB92-43c5-80A9-92277A4F7B54}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7F9185B0, CInt(&HCB92), CInt(&H43C5), &H80, &HA9, &H92, &H27, &H7A, &H4F, &H7B, &H54)
 IID_IExecuteCommand = iid
End Function
Public Function IID_IExecuteCommandHost() As UUID
'{4b6832a2-5f04-4c9d-b89d-727a15d103e7}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4B6832A2, CInt(&H5F04), CInt(&H4C9D), &HB8, &H9D, &H72, &H7A, &H15, &HD1, &H3, &HE7)
 IID_IExecuteCommandHost = iid
End Function
Public Function IID_IExplorerCommandProvider() As UUID
'{64961751-0835-43c0-8ffe-d57686530e64}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H64961751, CInt(&H835), CInt(&H43C0), &H8F, &HFE, &HD5, &H76, &H86, &H53, &HE, &H64)
 IID_IExplorerCommandProvider = iid
End Function
Public Function IID_IEnumExplorerCommand() As UUID
'{a88826f8-186f-4987-aade-ea0cef8fbfe8}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA88826F8, CInt(&H186F), CInt(&H4987), &HAA, &HDE, &HEA, &HC, &HEF, &H8F, &HBF, &HE8)
 IID_IEnumExplorerCommand = iid
End Function
Public Function IID_IInitializeCommand() As UUID
'{85075acf-231f-40ea-9610-d26b7b58f638}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H85075ACF, CInt(&H231F), CInt(&H40EA), &H96, &H10, &HD2, &H6B, &H7B, &H58, &HF6, &H38)
 IID_IInitializeCommand = iid
End Function
Public Function IID_IExplorerCommandState() As UUID
'{bddacb60-7657-47ae-8445-d23e1acf82ae}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBDDACB60, CInt(&H7657), CInt(&H47AE), &H84, &H45, &HD2, &H3E, &H1A, &HCF, &H82, &HAE)
 IID_IExplorerCommandState = iid
End Function
Public Function IID_IExplorerCommand() As UUID
'{a08ce4d0-fa25-44ab-b57c-c7b1c323e0b9}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA08CE4D0, CInt(&HFA25), CInt(&H44AB), &HB5, &H7C, &HC7, &HB1, &HC3, &H23, &HE0, &HB9)
 IID_IExplorerCommand = iid
End Function
Public Function IID_IMessageFilter() As UUID
'{00000016-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H16, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IMessageFilter = iid
End Function
Public Function IID_IApplicationDocumentLists() As UUID
'{3c594f9f-9f30-47a1-979a-c9e83d3d0a06}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3C594F9F, CInt(&H9F30), CInt(&H47A1), &H97, &H9A, &HC9, &HE8, &H3D, &H3D, &HA, &H6)
 IID_IApplicationDocumentLists = iid
End Function
Public Function IID_IShellChangeNotify() As UUID
'{D82BE2B1-5764-11D0-A96E-00C04FD705A2}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD82BE2B1, CInt(&H5764), CInt(&H11D0), &HA9, &H6E, &H0, &HC0, &H4F, &HD7, &H5, &HA2)
 IID_IShellChangeNotify = iid
End Function
Public Function IID_ITransferSource() As UUID
'{00adb003-bde9-45c6-8e29-d09f9353e108}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HADB003, CInt(&HBDE9), CInt(&H45C6), &H8E, &H29, &HD0, &H9F, &H93, &H53, &HE1, &H8)
IID_ITransferSource = iid
End Function
Public Function IID_IEnumResources() As UUID
'{2dd81fe3-a83c-4da9-a330-47249d345ba1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2DD81FE3, CInt(&HA83C), CInt(&H4DA9), &HA3, &H30, &H47, &H24, &H9D, &H34, &H5B, &HA1)
IID_IEnumResources = iid
End Function
Public Function IID_IShellItemResources() As UUID
'{ff5693be-2ce0-4d48-b5c5-40817d1acdb9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFF5693BE, CInt(&H2CE0), CInt(&H4D48), &HB5, &HC5, &H40, &H81, &H7D, &H1A, &HCD, &HB9)
IID_IShellItemResources = iid
End Function
Public Function IID_ITransferDestination() As UUID
'{48addd32-3ca5-4124-abe3-b5a72531b207}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H48ADDD32, CInt(&H3CA5), CInt(&H4124), &HAB, &HE3, &HB5, &HA7, &H25, &H31, &HB2, &H7)
IID_ITransferDestination = iid
End Function
Public Function IID_IKnownFolder() As UUID
'{3AA7AF7E-9B36-420c-A8E3-F77D4674A488}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3AA7AF7E, CInt(&H9B36), CInt(&H420C), &HA8, &HE3, &HF7, &H7D, &H46, &H74, &HA4, &H88)
IID_IKnownFolder = iid
End Function
Public Function IID_IKnownFolderManager() As UUID
'{8BE2D872-86AA-4d47-B776-32CCA40C7018}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8BE2D872, CInt(&H86AA), CInt(&H4D47), &HB7, &H76, &H32, &HCC, &HA4, &HC, &H70, &H18)
IID_IKnownFolderManager = iid
End Function
Public Function IID_IInitializeWithBindCtx() As UUID
'{71c0d2bc-726d-45cc-a6c0-2e31c1db2159}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H71C0D2BC, CInt(&H726D), CInt(&H45CC), &HA6, &HC0, &H2E, &H31, &HC1, &HDB, &H21, &H59)
IID_IInitializeWithBindCtx = iid
End Function
Public Function IID_IPreviewHandlerFrame() As UUID
'{fec87aaf-35f9-447a-adb7-20234491401a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFEC87AAF, CInt(&H35F9), CInt(&H447A), &HAD, &HB7, &H20, &H23, &H44, &H91, &H40, &H1A)
IID_IPreviewHandlerFrame = iid
End Function
Public Function IID_IVisualProperties() As UUID
'{e693cf68-d967-4112-8763-99172aee5e5a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE693CF68, CInt(&HD967), CInt(&H4112), &H87, &H63, &H99, &H17, &H2A, &HEE, &H5E, &H5A)
IID_IVisualProperties = iid
End Function
Public Function IID_ISpellingError() As UUID
'{B7C82D61-FBE8-4B47-9B27-6C0D2E0DE0A3}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB7C82D61, CInt(&HFBE8), CInt(&H4B47), &H9B, &H27, &H6C, &HD, &H2E, &HD, &HE0, &HA3)
 IID_ISpellingError = iid
End Function
Public Function IID_IEnumSpellingError() As UUID
'{803E3BD4-2828-4410-8290-418D1D73C762}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H803E3BD4, CInt(&H2828), CInt(&H4410), &H82, &H90, &H41, &H8D, &H1D, &H73, &HC7, &H62)
 IID_IEnumSpellingError = iid
End Function
Public Function IID_IOptionDescription() As UUID
'{432E5F85-35CF-4606-A801-6F70277E1D7A}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H432E5F85, CInt(&H35CF), CInt(&H4606), &HA8, &H1, &H6F, &H70, &H27, &H7E, &H1D, &H7A)
 IID_IOptionDescription = iid
End Function
Public Function IID_ISpellCheckerChangedEventHandler() As UUID
'{0B83A5B0-792F-4EAB-9799-ACF52C5ED08A}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB83A5B0, CInt(&H792F), CInt(&H4EAB), &H97, &H99, &HAC, &HF5, &H2C, &H5E, &HD0, &H8A)
 IID_ISpellCheckerChangedEventHandler = iid
End Function
Public Function IID_ISpellChecker() As UUID
'{B6FD0B71-E2BC-4653-8D05-F197E412770B}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB6FD0B71, CInt(&HE2BC), CInt(&H4653), &H8D, &H5, &HF1, &H97, &HE4, &H12, &H77, &HB)
 IID_ISpellChecker = iid
End Function
Public Function IID_ISpellChecker2() As UUID
'{E7ED1C71-87F7-4378-A840-C9200DACEE47}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE7ED1C71, CInt(&H87F7), CInt(&H4378), &HA8, &H40, &HC9, &H20, &HD, &HAC, &HEE, &H47)
 IID_ISpellChecker2 = iid
End Function
Public Function IID_ISpellCheckerFactory() As UUID
'{8E018A9D-2415-4677-BF08-794EA61F94BB}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8E018A9D, CInt(&H2415), CInt(&H4677), &HBF, &H8, &H79, &H4E, &HA6, &H1F, &H94, &HBB)
 IID_ISpellCheckerFactory = iid
End Function
Public Function IID_IUserDictionariesRegistrar() As UUID
'{AA176B85-0E12-4844-8E1A-EEF1DA77F586}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAA176B85, CInt(&HE12), CInt(&H4844), &H8E, &H1A, &HEE, &HF1, &HDA, &H77, &HF5, &H86)
 IID_IUserDictionariesRegistrar = iid
End Function
Public Function IID_ISpellCheckProvider() As UUID
'{73E976E0-8ED4-4EB1-80D7-1BE0A16B0C38}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H73E976E0, CInt(&H8ED4), CInt(&H4EB1), &H80, &HD7, &H1B, &HE0, &HA1, &H6B, &HC, &H38)
 IID_ISpellCheckProvider = iid
End Function
Public Function IID_IComprehensiveSpellCheckProvider() As UUID
'{0C58F8DE-8E94-479E-9717-70C42C4AD2C3}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC58F8DE, CInt(&H8E94), CInt(&H479E), &H97, &H17, &H70, &HC4, &H2C, &H4A, &HD2, &HC3)
 IID_IComprehensiveSpellCheckProvider = iid
End Function
Public Function IID_ISpellCheckProviderFactory() As UUID
'{9F671E11-77D6-4C92-AEFB-615215E3A4BE}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9F671E11, CInt(&H77D6), CInt(&H4C92), &HAE, &HFB, &H61, &H52, &H15, &HE3, &HA4, &HBE)
 IID_ISpellCheckProviderFactory = iid
End Function
Public Function IID_IRichChunk() As UUID
'{4FDEF69C-DBC9-454e-9910-B34F3C64B510}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4FDEF69C, CInt(&HDBC9), CInt(&H454E), &H99, &H10, &HB3, &H4F, &H3C, &H64, &HB5, &H10)
IID_IRichChunk = iid
End Function
Public Function IID_ICondition2() As UUID
'{0DB8851D-2E5B-47eb-9208-D28C325A01D7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDB8851D, CInt(&H2E5B), CInt(&H47EB), &H92, &H8, &HD2, &H8C, &H32, &H5A, &H1, &HD7)
IID_ICondition2 = iid
End Function
Public Function IID_IConditionFactory() As UUID
'{A5EFE073-B16F-474f-9F3E-9F8B497A3E08}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA5EFE073, CInt(&HB16F), CInt(&H474F), &H9F, &H3E, &H9F, &H8B, &H49, &H7A, &H3E, &H8)
IID_IConditionFactory = iid
End Function
Public Function IID_IConditionFactory2() As UUID
'{71D222E1-432F-429e-8C13-B6DAFDE5077A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H71D222E1, CInt(&H432F), CInt(&H429E), &H8C, &H13, &HB6, &HDA, &HFD, &HE5, &H7, &H7A)
IID_IConditionFactory2 = iid
End Function
Public Function IID_IQueryParser() As UUID
'{2EBDEE67-3505-43f8-9946-EA44ABC8E5B0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2EBDEE67, CInt(&H3505), CInt(&H43F8), &H99, &H46, &HEA, &H44, &HAB, &HC8, &HE5, &HB0)
IID_IQueryParser = iid
End Function
Public Function IID_IQuerySolution() As UUID
'{D6EBC66B-8921-4193-AFDD-A1789FB7FF57}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD6EBC66B, CInt(&H8921), CInt(&H4193), &HAF, &HDD, &HA1, &H78, &H9F, &HB7, &HFF, &H57)
IID_IQuerySolution = iid
End Function
Public Function IID_IConditionGenerator() As UUID
'{92D2CC58-4386-45a3-B98C-7E0CE64A4117}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H92D2CC58, CInt(&H4386), CInt(&H45A3), &HB9, &H8C, &H7E, &HC, &HE6, &H4A, &H41, &H17)
IID_IConditionGenerator = iid
End Function
Public Function IID_IInterval() As UUID
'{6BF0A714-3C18-430b-8B5D-83B1C234D3DB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6BF0A714, CInt(&H3C18), CInt(&H430B), &H8B, &H5D, &H83, &HB1, &HC2, &H34, &HD3, &HDB)
IID_IInterval = iid
End Function
Public Function IID_IMetaData() As UUID
'{780102B0-C43B-4876-BC7B-5E9BA5C88794}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H780102B0, CInt(&HC43B), CInt(&H4876), &HBC, &H7B, &H5E, &H9B, &HA5, &HC8, &H87, &H94)
IID_IMetaData = iid
End Function
Public Function IID_IEntity() As UUID
'{24264891-E80B-4fd3-B7CE-4FF2FAE8931F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H24264891, CInt(&HE80B), CInt(&H4FD3), &HB7, &HCE, &H4F, &HF2, &HFA, &HE8, &H93, &H1F)
IID_IEntity = iid
End Function
Public Function IID_IRelationship() As UUID
'{2769280B-5108-498c-9C7F-A51239B63147}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2769280B, CInt(&H5108), CInt(&H498C), &H9C, &H7F, &HA5, &H12, &H39, &HB6, &H31, &H47)
IID_IRelationship = iid
End Function
Public Function IID_INamedEntity() As UUID
'{ABDBD0B1-7D54-49fb-AB5C-BFF4130004CD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HABDBD0B1, CInt(&H7D54), CInt(&H49FB), &HAB, &H5C, &HBF, &HF4, &H13, &H0, &H4, &HCD)
IID_INamedEntity = iid
End Function
Public Function IID_ISchemaProvider() As UUID
'{8CF89BCB-394C-49b2-AE28-A59DD4ED7F68}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CF89BCB, CInt(&H394C), CInt(&H49B2), &HAE, &H28, &HA5, &H9D, &HD4, &HED, &H7F, &H68)
IID_ISchemaProvider = iid
End Function
Public Function IID_ITokenCollection() As UUID
'{22D8B4F2-F577-4adb-A335-C2AE88416FAB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H22D8B4F2, CInt(&HF577), CInt(&H4ADB), &HA3, &H35, &HC2, &HAE, &H88, &H41, &H6F, &HAB)
IID_ITokenCollection = iid
End Function
Public Function IID_INamedEntityCollector() As UUID
'{AF2440F6-8AFC-47d0-9A7F-396A0ACFB43D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAF2440F6, CInt(&H8AFC), CInt(&H47D0), &H9A, &H7F, &H39, &H6A, &HA, &HCF, &HB4, &H3D)
IID_INamedEntityCollector = iid
End Function
Public Function IID_ISchemaLocalizerSupport() As UUID
'{CA3FDCA2-BFBE-4eed-90D7-0CAEF0A1BDA1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCA3FDCA2, CInt(&HBFBE), CInt(&H4EED), &H90, &HD7, &HC, &HAE, &HF0, &HA1, &HBD, &HA1)
IID_ISchemaLocalizerSupport = iid
End Function
Public Function IID_IQueryParserManager() As UUID
'{A879E3C4-AF77-44fb-8F37-EBD1487CF920}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA879E3C4, CInt(&HAF77), CInt(&H44FB), &H8F, &H37, &HEB, &HD1, &H48, &H7C, &HF9, &H20)
IID_IQueryParserManager = iid
End Function
Public Function IID_ISearchFolderItemFactory() As UUID
'{a0ffbc28-5482-4366-be27-3e81e78e06c2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA0FFBC28, CInt(&H5482), CInt(&H4366), &HBE, &H27, &H3E, &H81, &HE7, &H8E, &H6, &HC2)
IID_ISearchFolderItemFactory = iid
End Function
Public Function IID_IThumbnailHandlerFactory() As UUID
'{e35b4b2e-00da-4bc1-9f13-38bc11f5d417}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE35B4B2E, CInt(&HDA), CInt(&H4BC1), &H9F, &H13, &H38, &HBC, &H11, &HF5, &HD4, &H17)
IID_IThumbnailHandlerFactory = iid
End Function
Public Function IID_ISharedBitmap() As UUID
'{091162a4-bc96-411f-aae8-c5122cd03363}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H91162A4, CInt(&HBC96), CInt(&H411F), &HAA, &HE8, &HC5, &H12, &H2C, &HD0, &H33, &H63)
 IID_ISharedBitmap = iid
End Function
Public Function IID_IThumbnailCache() As UUID
'{F676C15D-596A-4ce2-8234-33996F445DB1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF676C15D, CInt(&H596A), CInt(&H4CE2), &H82, &H34, &H33, &H99, &H6F, &H44, &H5D, &HB1)
 IID_IThumbnailCache = iid
End Function
Public Function IID_IThumbnailSettings() As UUID
'{F4376F00-BEF5-4d45-80F3-1E023BBF1209}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF4376F00, CInt(&HBEF5), CInt(&H4D45), &H80, &HF3, &H1E, &H2, &H3B, &HBF, &H12, &H9)
 IID_IThumbnailSettings = iid
End Function
Public Function IID_ITrackShellMenu() As UUID
'{8278F932-2A3E-11d2-838F-00C04FD918D0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8278F932, CInt(&H2A3E), CInt(&H11D2), &H83, &H8F, &H0, &HC0, &H4F, &HD9, &H18, &HD0)
 IID_ITrackShellMenu = iid
End Function
Public Function IID_IApplicationActivationManager() As UUID
'{2e941141-7f97-4756-ba1d-9decde894a3d}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2E941141, CInt(&H7F97), CInt(&H4756), &HBA, &H1D, &H9D, &HEC, &HDE, &H89, &H4A, &H3D)
 IID_IApplicationActivationManager = iid
End Function
Public Function IID_IAppVisibilityEvents() As UUID
'{6584CE6B-7D82-49C2-89C9-C6BC02BA8C38}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6584CE6B, CInt(&H7D82), CInt(&H49C2), &H89, &HC9, &HC6, &HBC, &H2, &HBA, &H8C, &H38)
 IID_IAppVisibilityEvents = iid
End Function
Public Function IID_IAppVisibility() As UUID
'{2246EA2D-CAEA-4444-A3C4-6DE827E44313}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2246EA2D, CInt(&HCAEA), CInt(&H4444), &HA3, &HC4, &H6D, &HE8, &H27, &HE4, &H43, &H13)
 IID_IAppVisibility = iid
End Function
Public Function IID_IImageRecompress() As UUID
'{505f1513-6b3e-4892-a272-59f8889a4d3e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H505F1513, CInt(&H6B3E), CInt(&H4892), &HA2, &H72, &H59, &HF8, &H88, &H9A, &H4D, &H3E)
IID_IImageRecompress = iid
End Function
Public Function IID_ITranscodeImage() As UUID
'{BAE86DDD-DC11-421c-B7AB-CC55D1D65C44}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBAE86DDD, CInt(&HDC11), CInt(&H421C), &HB7, &HAB, &HCC, &H55, &HD1, &HD6, &H5C, &H44)
IID_ITranscodeImage = iid
End Function
Public Function IID_IParentAndItem() As UUID
'{b3a4b685-b685-4805-99d9-5dead2873236}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB3A4B685, CInt(&HB685), CInt(&H4805), &H99, &HD9, &H5D, &HEA, &HD2, &H87, &H32, &H36)
IID_IParentAndItem = iid
End Function
Public Function IID_ISearchBoxInfo() As UUID
'{6af6e03f-d664-4ef4-9626-f7e0ed36755e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6AF6E03F, CInt(&HD664), CInt(&H4EF4), &H96, &H26, &HF7, &HE0, &HED, &H36, &H75, &H5E)
IID_ISearchBoxInfo = iid
End Function
Public Function IID_IShellFolderViewCB() As UUID
'{2047E320-F2A9-11CE-AE65-08002B2E1262}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2047E320, CInt(&HF2A9), CInt(&H11CE), &HAE, &H65, &H8, &H0, &H2B, &H2E, &H12, &H62)
IID_IShellFolderViewCB = iid
End Function
Public Function IID_IPreviousVersionsInfo() As UUID
'{76e54780-ad74-48e3-a695-3ba9a0aff10d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H76E54780, CInt(&HAD74), CInt(&H48E3), &HA6, &H95, &H3B, &HA9, &HA0, &HAF, &HF1, &HD)
IID_IPreviousVersionsInfo = iid
End Function
Public Function IID_IZoneIdentifier() As UUID
'{cd45f185-1b21-48e2-967b-ead743a8914e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCD45F185, CInt(&H1B21), CInt(&H48E2), &H96, &H7B, &HEA, &HD7, &H43, &HA8, &H91, &H4E)
IID_IZoneIdentifier = iid
End Function
Public Function IID_IApplicationAssociationRegistration() As UUID
'{4e530b0a-e611-4c77-a3ac-9031d022281b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4E530B0A, CInt(&HE611), CInt(&H4C77), &HA3, &HAC, &H90, &H31, &HD0, &H22, &H28, &H1B)
IID_IApplicationAssociationRegistration = iid
End Function
Public Function IID_IApplicationAssociationRegistrationUI() As UUID
'{1f76a169-f994-40ac-8fc8-0959e8874710}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1F76A169, CInt(&HF994), CInt(&H40AC), &H8F, &HC8, &H9, &H59, &HE8, &H87, &H47, &H10)
IID_IApplicationAssociationRegistrationUI = iid
End Function
Public Function IID_ISystemInformation() As UUID
'{ADE87BF7-7B56-4275-8FAB-B9B0E591844B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HADE87BF7, CInt(&H7B56), CInt(&H4275), &H8F, &HAB, &HB9, &HB0, &HE5, &H91, &H84, &H4B)
IID_ISystemInformation = iid
End Function
Public Function IID_IFolderViewSettings() As UUID
'{ae8c987d-8797-4ed3-be72-2a47dd938db0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAE8C987D, CInt(&H8797), CInt(&H4ED3), &HBE, &H72, &H2A, &H47, &HDD, &H93, &H8D, &HB0)
IID_IFolderViewSettings = iid
End Function
Public Function IID_IFolderViewOptions() As UUID
'{3cc974d2-b302-4d36-ad3e-06d93f695d3f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3CC974D2, CInt(&HB302), CInt(&H4D36), &HAD, &H3E, &H6, &HD9, &H3F, &H69, &H5D, &H3F)
IID_IFolderViewOptions = iid
End Function
Public Function IID_IResolveShellLink() As UUID
'{5cd52983-9449-11d2-963a-00c04f79adf0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5CD52983, CInt(&H9449), CInt(&H11D2), &H96, &H3A, &H0, &HC0, &H4F, &H79, &HAD, &HF0)
IID_IResolveShellLink = iid
End Function
Public Function IID_IStartMenuPinnedList() As UUID
'{4CD19ADA-25A5-4A32-B3B7-347BEE5BE36B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4CD19ADA, CInt(&H25A5), CInt(&H4A32), &HB3, &HB7, &H34, &H7B, &HEE, &H5B, &HE3, &H6B)
IID_IStartMenuPinnedList = iid
End Function
Public Function IID_IObjMgr() As UUID
'{00BB2761-6A77-11D0-A535-00C04FD7D062}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBB2761, CInt(&H6A77), CInt(&H11D0), &HA5, &H35, &H0, &HC0, &H4F, &HD7, &HD0, &H62)
IID_IObjMgr = iid
End Function
Public Function IID_IAutoCompleteDropDown() As UUID
'{3CD141F4-3C6A-11d2-BCAA-00C04FD929DB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3CD141F4, CInt(&H3C6A), CInt(&H11D2), &HBC, &HAA, &H0, &HC0, &H4F, &HD9, &H29, &HDB)
IID_IAutoCompleteDropDown = iid
End Function
Public Function IID_IFolderFilter() As UUID
'{9CC22886-DC8E-11d2-B1D0-00C04F8EEB3E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9CC22886, CInt(&HDC8E), CInt(&H11D2), &HB1, &HD0, &H0, &HC0, &H4F, &H8E, &HEB, &H3E)
IID_IFolderFilter = iid
End Function
Public Function IID_IShellLinkDataList() As UUID
'{45e2b4ae-b1c3-11d0-b92f-00a0c90312e1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H45E2B4AE, CInt(&HB1C3), CInt(&H11D0), &HB9, &H2F, &H0, &HA0, &HC9, &H3, &H12, &HE1)
IID_IShellLinkDataList = iid
End Function
Public Function IID_IDataObjectAsyncCapability() As UUID
'{3D8B0590-F691-11d2-8EA9-006097DF5BD4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3D8B0590, CInt(&HF691), CInt(&H11D2), &H8E, &HA9, &H0, &H60, &H97, &HDF, &H5B, &HD4)
IID_IDataObjectAsyncCapability = iid
End Function
Public Function IID_IPortableDeviceManager() As UUID
'{A1567595-4C2F-4574-A6FA-ECEF917B9A40}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA1567595, CInt(&H4C2F), CInt(&H4574), &HA6, &HFA, &HEC, &HEF, &H91, &H7B, &H9A, &H40)
IID_IPortableDeviceManager = iid
End Function
Public Function IID_IPortableDeviceValuesCollection() As UUID
'{6E3F2D79-4E07-48C4-8208-D8C2E5AF4A99}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6E3F2D79, CInt(&H4E07), CInt(&H48C4), &H82, &H8, &HD8, &HC2, &HE5, &HAF, &H4A, &H99)
IID_IPortableDeviceValuesCollection = iid
End Function
Public Function IID_IPortableDevicePropVariantCollection() As UUID
'{89B2E422-4F1B-4316-BCEF-A44AFEA83EB3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H89B2E422, CInt(&H4F1B), CInt(&H4316), &HBC, &HEF, &HA4, &H4A, &HFE, &HA8, &H3E, &HB3)
IID_IPortableDevicePropVariantCollection = iid
End Function
Public Function IID_IPortableDeviceKeyCollection() As UUID
'{DADA2357-E0AD-492E-98DB-DD61C53BA353}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDADA2357, CInt(&HE0AD), CInt(&H492E), &H98, &HDB, &HDD, &H61, &HC5, &H3B, &HA3, &H53)
IID_IPortableDeviceKeyCollection = iid
End Function
Public Function IID_IPortableDeviceValues() As UUID
'{6848F6F2-3155-4F86-B6F5-263EEEAB3143}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6848F6F2, CInt(&H3155), CInt(&H4F86), &HB6, &HF5, &H26, &H3E, &HEE, &HAB, &H31, &H43)
IID_IPortableDeviceValues = iid
End Function
Public Function IID_IPortableDevice() As UUID
'{625E2DF8-6392-4CF0-9AD1-3CFA5F17775C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H625E2DF8, CInt(&H6392), CInt(&H4CF0), &H9A, &HD1, &H3C, &HFA, &H5F, &H17, &H77, &H5C)
IID_IPortableDevice = iid
End Function
Public Function IID_IPortableDeviceContent() As UUID
'{6A96ED84-7C73-4480-9938-BF5AF477D426}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6A96ED84, CInt(&H7C73), CInt(&H4480), &H99, &H38, &HBF, &H5A, &HF4, &H77, &HD4, &H26)
IID_IPortableDeviceContent = iid
End Function
Public Function IID_IEnumPortableDeviceObjectIDs() As UUID
'{10ECE955-CF41-4728-BFA0-41EEDF1BBF19}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10ECE955, CInt(&HCF41), CInt(&H4728), &HBF, &HA0, &H41, &HEE, &HDF, &H1B, &HBF, &H19)
IID_IEnumPortableDeviceObjectIDs = iid
End Function
Public Function IID_IPortableDeviceProperties() As UUID
'{7F6D695C-03DF-4439-A809-59266BEEE3A6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7F6D695C, CInt(&H3DF), CInt(&H4439), &HA8, &H9, &H59, &H26, &H6B, &HEE, &HE3, &HA6)
IID_IPortableDeviceProperties = iid
End Function
Public Function IID_IPortableDeviceResources() As UUID
'{FD8878AC-D841-4D17-891C-E6829CDB6934}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFD8878AC, CInt(&HD841), CInt(&H4D17), &H89, &H1C, &HE6, &H82, &H9C, &HDB, &H69, &H34)
IID_IPortableDeviceResources = iid
End Function
Public Function IID_IPortableDeviceCapabilities() As UUID
'{2C8C6DBF-E3DC-4061-BECC-8542E810D126}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2C8C6DBF, CInt(&HE3DC), CInt(&H4061), &HBE, &HCC, &H85, &H42, &HE8, &H10, &HD1, &H26)
IID_IPortableDeviceCapabilities = iid
End Function
Public Function IID_IPortableDeviceService() As UUID
'{D3BD3A44-D7B5-40A9-98B7-2FA4D01DEC08}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD3BD3A44, CInt(&HD7B5), CInt(&H40A9), &H98, &HB7, &H2F, &HA4, &HD0, &H1D, &HEC, &H8)
IID_IPortableDeviceService = iid
End Function
Public Function IID_IPortableDeviceServiceCapabilities() As UUID
'{24DBD89D-413E-43E0-BD5B-197F3C56C886}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H24DBD89D, CInt(&H413E), CInt(&H43E0), &HBD, &H5B, &H19, &H7F, &H3C, &H56, &HC8, &H86)
IID_IPortableDeviceServiceCapabilities = iid
End Function
Public Function IID_IPortableDeviceContent2() As UUID
'{9B4ADD96-F6BF-4034-8708-ECA72BF10554}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9B4ADD96, CInt(&HF6BF), CInt(&H4034), &H87, &H8, &HEC, &HA7, &H2B, &HF1, &H5, &H54)
IID_IPortableDeviceContent2 = iid
End Function
Public Function IID_IPortableDeviceServiceMethods() As UUID
'{E20333C9-FD34-412D-A381-CC6F2D820DF7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE20333C9, CInt(&HFD34), CInt(&H412D), &HA3, &H81, &HCC, &H6F, &H2D, &H82, &HD, &HF7)
IID_IPortableDeviceServiceMethods = iid
End Function
Public Function IID_IPortableDeviceDispatchFactory() As UUID
'{5E1EAFC3-E3D7-4132-96FA-759C0F9D1E0F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5E1EAFC3, CInt(&HE3D7), CInt(&H4132), &H96, &HFA, &H75, &H9C, &HF, &H9D, &H1E, &HF)
IID_IPortableDeviceDispatchFactory = iid
End Function
Public Function IID_IWpdSerializer() As UUID
'{B32F4002-BB27-45FF-AF4F-06631C1E8DAD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB32F4002, CInt(&HBB27), CInt(&H45FF), &HAF, &H4F, &H6, &H63, &H1C, &H1E, &H8D, &HAD)
IID_IWpdSerializer = iid
End Function
Public Function IID_IPortableDeviceDataStream() As UUID
'{88e04db3-1012-4d64-9996-f703a950d3f4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H88E04DB3, CInt(&H1012), CInt(&H4D64), &H99, &H96, &HF7, &H3, &HA9, &H50, &HD3, &HF4)
IID_IPortableDeviceDataStream = iid
End Function
Public Function IID_IPortableDeviceUnitsStream() As UUID
'{5e98025f-bfc4-47a2-9a5f-bc900a507c67}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5E98025F, CInt(&HBFC4), CInt(&H47A2), &H9A, &H5F, &HBC, &H90, &HA, &H50, &H7C, &H67)
IID_IPortableDeviceUnitsStream = iid
End Function
Public Function IID_IPortableDevicePropertiesBulk() As UUID
'{482b05c0-4056-44ed-9e0f-5e23b009da93}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H482B05C0, CInt(&H4056), CInt(&H44ED), &H9E, &HF, &H5E, &H23, &HB0, &H9, &HDA, &H93)
IID_IPortableDevicePropertiesBulk = iid
End Function
Public Function IID_IPortableDeviceServiceActivation() As UUID
'{e56b0534-d9b9-425c-9b99-75f97cb3d7c8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE56B0534, CInt(&HD9B9), CInt(&H425C), &H9B, &H99, &H75, &HF9, &H7C, &HB3, &HD7, &HC8)
IID_IPortableDeviceServiceActivation = iid
End Function
Public Function IID_IPortableDeviceWebControl() As UUID
'{94fc7953-5ca1-483a-8aee-df52e7747d00}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H94FC7953, CInt(&H5CA1), CInt(&H483A), &H8A, &HEE, &HDF, &H52, &HE7, &H74, &H7D, &H0)
IID_IPortableDeviceWebControl = iid
End Function
Public Function IID_IPortableDeviceServiceMethodCallback() As UUID
'{C424233C-AFCE-4828-A756-7ED7A2350083}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC424233C, CInt(&HAFCE), CInt(&H4828), &HA7, &H56, &H7E, &HD7, &HA2, &H35, &H0, &H83)
IID_IPortableDeviceServiceMethodCallback = iid
End Function
Public Function IID_IPortableDeviceServiceOpenCallback() As UUID
'{bced49c8-8efe-41ed-960b-61313abd47a9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBCED49C8, CInt(&H8EFE), CInt(&H41ED), &H96, &HB, &H61, &H31, &H3A, &HBD, &H47, &HA9)
IID_IPortableDeviceServiceOpenCallback = iid
End Function
Public Function IID_IPortableDeviceEventCallback() As UUID
'{A8792A31-F385-493C-A893-40F64EB45F6E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA8792A31, CInt(&HF385), CInt(&H493C), &HA8, &H93, &H40, &HF6, &H4E, &HB4, &H5F, &H6E)
IID_IPortableDeviceEventCallback = iid
End Function
Public Function IID_IConnectionRequestCallback() As UUID
'{272C9AE0-7161-4AE0-91BD-9F448EE9C427}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H272C9AE0, CInt(&H7161), CInt(&H4AE0), &H91, &HBD, &H9F, &H44, &H8E, &HE9, &HC4, &H27)
IID_IConnectionRequestCallback = iid
End Function
Public Function IID_IPortableDevicePropertiesBulkCallback() As UUID
'{9deacb80-11e8-40e3-a9f3-f557986a7845}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9DEACB80, CInt(&H11E8), CInt(&H40E3), &HA9, &HF3, &HF5, &H57, &H98, &H6A, &H78, &H45)
IID_IPortableDevicePropertiesBulkCallback = iid
End Function
Public Function IID_IPortableDeviceConnector() As UUID
'{625E2DF8-6392-4CF0-9AD1-3CFA5F17775C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H625E2DF8, CInt(&H6392), CInt(&H4CF0), &H9A, &HD1, &H3C, &HFA, &H5F, &H17, &H77, &H5C)
IID_IPortableDeviceConnector = iid
End Function
Public Function IID_IEnumPortableDeviceConnectors() As UUID
'{BFDEF549-9247-454F-BD82-06FE80853FAA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBFDEF549, CInt(&H9247), CInt(&H454F), &HBD, &H82, &H6, &HFE, &H80, &H85, &H3F, &HAA)
IID_IEnumPortableDeviceConnectors = iid
End Function
Public Function IID_IEnumNetConnection() As UUID
'{C08956A0-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956A0, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_IEnumNetConnection = iid
End Function
Public Function IID_INetConnection() As UUID
'{C08956A1-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956A1, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_INetConnection = iid
End Function
Public Function IID_INetConnectionManager() As UUID
'{C08956A2-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956A2, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_INetConnectionManager = iid
End Function
Public Function IID_INetConnectionConnectUi() As UUID
'{C08956A3-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956A3, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_INetConnectionConnectUi = iid
End Function
Public Function IID_IEnumNetSharingPortMapping() As UUID
'{C08956B0-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956B0, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_IEnumNetSharingPortMapping = iid
End Function
Public Function IID_INetSharingPortMapping() As UUID
'{C08956B1-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956B1, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_INetSharingPortMapping = iid
End Function
Public Function IID_INetSharingPortMappingProps() As UUID
'{24B7E9B5-E38F-4685-851B-00892CF5F940}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H24B7E9B5, CInt(&HE38F), CInt(&H4685), &H85, &H1B, &H0, &H89, &H2C, &HF5, &HF9, &H40)
IID_INetSharingPortMappingProps = iid
End Function
Public Function IID_IEnumNetSharingEveryConnection() As UUID
'{C08956B8-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956B8, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_IEnumNetSharingEveryConnection = iid
End Function
Public Function IID_IEnumNetSharingPublicConnection() As UUID
'{C08956B4-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956B4, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_IEnumNetSharingPublicConnection = iid
End Function
Public Function IID_IEnumNetSharingPrivateConnection() As UUID
'{C08956B5-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956B5, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_IEnumNetSharingPrivateConnection = iid
End Function
Public Function IID_INetSharingPortMappingCollection() As UUID
'{02E4A2DE-DA20-4E34-89C8-AC22275A010B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2E4A2DE, CInt(&HDA20), CInt(&H4E34), &H89, &HC8, &HAC, &H22, &H27, &H5A, &H1, &HB)
IID_INetSharingPortMappingCollection = iid
End Function
Public Function IID_INetConnectionProps() As UUID
'{F4277C95-CE5B-463D-8167-5662D9BCAA72}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF4277C95, CInt(&HCE5B), CInt(&H463D), &H81, &H67, &H56, &H62, &HD9, &HBC, &HAA, &H72)
IID_INetConnectionProps = iid
End Function
Public Function IID_INetSharingConfiguration() As UUID
'{C08956B6-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956B6, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_INetSharingConfiguration = iid
End Function
Public Function IID_INetSharingEveryConnectionCollection() As UUID
'{33C4643C-7811-46FA-A89A-768597BD7223}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H33C4643C, CInt(&H7811), CInt(&H46FA), &HA8, &H9A, &H76, &H85, &H97, &HBD, &H72, &H23)
IID_INetSharingEveryConnectionCollection = iid
End Function
Public Function IID_INetSharingPublicConnectionCollection() As UUID
'{7D7A6355-F372-4971-A149-BFC927BE762A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7D7A6355, CInt(&HF372), CInt(&H4971), &HA1, &H49, &HBF, &HC9, &H27, &HBE, &H76, &H2A)
IID_INetSharingPublicConnectionCollection = iid
End Function
Public Function IID_INetSharingPrivateConnectionCollection() As UUID
'{38AE69E0-4409-402A-A2CB-E965C727F840}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H38AE69E0, CInt(&H4409), CInt(&H402A), &HA2, &HCB, &HE9, &H65, &HC7, &H27, &HF8, &H40)
IID_INetSharingPrivateConnectionCollection = iid
End Function
Public Function IID_INetSharingManager() As UUID
'{C08956B7-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956B7, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_INetSharingManager = iid
End Function
Public Function IID_IEnumReadyCallback() As UUID
'{61E00D45-8FFF-4e60-924E-6537B61612DD}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H61E00D45, CInt(&H8FFF), CInt(&H4E60), &H92, &H4E, &H65, &H37, &HB6, &H16, &H12, &HDD)
 IID_IEnumReadyCallback = iid
End Function
Public Function IID_IEnumerableView() As UUID
'{8C8BF236-1AEC-495f-9894-91D57C3C686F}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8C8BF236, CInt(&H1AEC), CInt(&H495F), &H98, &H94, &H91, &HD5, &H7C, &H3C, &H68, &H6F)
 IID_IEnumerableView = iid
End Function
Public Function IID_IPreviewItem() As UUID
'{36149969-0A8F-49c8-8B00-4AECB20222FB}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H36149969, CInt(&HA8F), CInt(&H49C8), &H8B, &H0, &H4A, &HEC, &HB2, &H2, &H22, &HFB)
 IID_IPreviewItem = iid
End Function
Public Function IID_IViewStateIdentityItem() As UUID
'{9D264146-A94F-4195-9F9F-3BB12CE0C955}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9D264146, CInt(&HA94F), CInt(&H4195), &H9F, &H9F, &H3B, &HB1, &H2C, &HE0, &HC9, &H55)
 IID_IViewStateIdentityItem = iid
End Function
Public Function IID_IDisplayItem() As UUID
'{c6fd5997-9f6b-4888-8703-94e80e8cde3f}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC6FD5997, CInt(&H9F6B), CInt(&H4888), &H87, &H3, &H94, &HE8, &HE, &H8C, &HDE, &H3F)
 IID_IDisplayItem = iid
End Function
Public Function IID_IUseToBrowseItem() As UUID
'{05edda5c-98a3-4717-8adb-c5e7da991eb1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5EDDA5C, CInt(&H98A3), CInt(&H4717), &H8A, &HDB, &HC5, &HE7, &HDA, &H99, &H1E, &HB1)
 IID_IUseToBrowseItem = iid
End Function
Public Function IID_ITransferMedium() As UUID
'{77f295d5-2d6f-4e19-b8ae-322f3e721ab5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H77F295D5, CInt(&H2D6F), CInt(&H4E19), &HB8, &HAE, &H32, &H2F, &H3E, &H72, &H1A, &HB5)
 IID_ITransferMedium = iid
End Function
Public Function IID_ICurrentItem() As UUID
'{240a7174-d653-4a1d-a6d3-d4943cfbfe3d}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H240A7174, CInt(&HD653), CInt(&H4A1D), &HA6, &HD3, &HD4, &H94, &H3C, &HFB, &HFE, &H3D)
 IID_ICurrentItem = iid
End Function
Public Function IID_IDelegateItem() As UUID
'{3c5a1c94-c951-4cb7-bb6d-3b93f30cce9}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3C5A1C94, CInt(&HC951), CInt(&H4CB7), &HBB, &H6D, &H3B, &H93, &HF3, &HC, &HCE, &H9)
 IID_IDelegateItem = iid
End Function
Public Function IID_IIdentityName() As UUID
'{7d903fca-d6f9-4810-8332-946c0177e247}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7D903FCA, CInt(&HD6F9), CInt(&H4810), &H83, &H32, &H94, &H6C, &H1, &H77, &HE2, &H47)
 IID_IIdentityName = iid
End Function
Public Function IID_IRelatedItem() As UUID
'{a73ce67a-8ab1-44f1-8d43-d2fcbf6b1cd0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA73CE67A, CInt(&H8AB1), CInt(&H44F1), &H8D, &H43, &HD2, &HFC, &HBF, &H6B, &H1C, &HD0)
 IID_IRelatedItem = iid
End Function
Public Function IID_IFilterCondition() As UUID
'{FCA2857D-1760-4AD3-8C63-C9B602FCBAEA}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFCA2857D, CInt(&H1760), CInt(&H4AD3), &H8C, &H63, &HC9, &HB6, &H2, &HFC, &HBA, &HEA)
 IID_IFilterCondition = iid
End Function
Public Function IID_IItemFilter() As UUID
'{7FCBEB25-ED60-45C9-9F5E-57B48493C4DD}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7FCBEB25, CInt(&HED60), CInt(&H45C9), &H9F, &H5E, &H57, &HB4, &H84, &H93, &HC4, &HDD)
 IID_IItemFilter = iid
End Function
Public Function IID_INewMenuClient() As UUID
'{dcb07fdc-3bb5-451c-90be-966644fed7b0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDCB07FDC, CInt(&H3BB5), CInt(&H451C), &H90, &HBE, &H96, &H66, &H44, &HFE, &HD7, &HB0)
 IID_INewMenuClient = iid
End Function
Public Function IID_IItemNameLimits() As UUID
'{1df0d7f1-b267-4d28-8b10-12e23202a5c4}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1DF0D7F1, CInt(&HB267), CInt(&H4D28), &H8B, &H10, &H12, &HE2, &H32, &H2, &HA5, &HC4)
 IID_IItemNameLimits = iid
End Function
Public Function IID_ITaskFolderCollection() As UUID
'{79184A66-8664-423F-97F1-637356A5D812}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79184A66, CInt(&H8664), CInt(&H423F), &H97, &HF1, &H63, &H73, &H56, &HA5, &HD8, &H12)
IID_ITaskFolderCollection = iid
End Function
Public Function IID_ITaskFolder() As UUID
'{8CFAC062-A080-4C15-9A88-AA7C2AF80DFC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CFAC062, CInt(&HA080), CInt(&H4C15), &H9A, &H88, &HAA, &H7C, &H2A, &HF8, &HD, &HFC)
IID_ITaskFolder = iid
End Function
Public Function IID_IRegisteredTask() As UUID
'{9C86F320-DEE3-4DD1-B972-A303F26B061E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9C86F320, CInt(&HDEE3), CInt(&H4DD1), &HB9, &H72, &HA3, &H3, &HF2, &H6B, &H6, &H1E)
IID_IRegisteredTask = iid
End Function
Public Function IID_IRunningTask() As UUID
'{653758FB-7B9A-4F1E-A471-BEEB8E9B834E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H653758FB, CInt(&H7B9A), CInt(&H4F1E), &HA4, &H71, &HBE, &HEB, &H8E, &H9B, &H83, &H4E)
IID_IRunningTask = iid
End Function
Public Function IID_IRunningTaskCollection() As UUID
'{6A67614B-6828-4FEC-AA54-6D52E8F1F2DB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6A67614B, CInt(&H6828), CInt(&H4FEC), &HAA, &H54, &H6D, &H52, &HE8, &HF1, &HF2, &HDB)
IID_IRunningTaskCollection = iid
End Function
Public Function IID_ITaskDefinition() As UUID
'{F5BC8FC5-536D-4F77-B852-FBC1356FDEB6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF5BC8FC5, CInt(&H536D), CInt(&H4F77), &HB8, &H52, &HFB, &HC1, &H35, &H6F, &HDE, &HB6)
IID_ITaskDefinition = iid
End Function
Public Function IID_IRegistrationInfo() As UUID
'{416D8B73-CB41-4EA1-805C-9BE9A5AC4A74}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H416D8B73, CInt(&HCB41), CInt(&H4EA1), &H80, &H5C, &H9B, &HE9, &HA5, &HAC, &H4A, &H74)
IID_IRegistrationInfo = iid
End Function
Public Function IID_ITriggerCollection() As UUID
'{85DF5081-1B24-4F32-878A-D9D14DF4CB77}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H85DF5081, CInt(&H1B24), CInt(&H4F32), &H87, &H8A, &HD9, &HD1, &H4D, &HF4, &HCB, &H77)
IID_ITriggerCollection = iid
End Function
Public Function IID_ITrigger() As UUID
'{09941815-EA89-4B5B-89E0-2A773801FAC3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9941815, CInt(&HEA89), CInt(&H4B5B), &H89, &HE0, &H2A, &H77, &H38, &H1, &HFA, &HC3)
IID_ITrigger = iid
End Function
Public Function IID_IRepetitionPattern() As UUID
'{7FB9ACF1-26BE-400E-85B5-294B9C75DFD6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7FB9ACF1, CInt(&H26BE), CInt(&H400E), &H85, &HB5, &H29, &H4B, &H9C, &H75, &HDF, &HD6)
IID_IRepetitionPattern = iid
End Function
Public Function IID_ITaskSettings() As UUID
'{8FD4711D-2D02-4C8C-87E3-EFF699DE127E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8FD4711D, CInt(&H2D02), CInt(&H4C8C), &H87, &HE3, &HEF, &HF6, &H99, &HDE, &H12, &H7E)
IID_ITaskSettings = iid
End Function
Public Function IID_IIdleSettings() As UUID
'{84594461-0053-4342-A8FD-088FABF11F32}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H84594461, CInt(&H53), CInt(&H4342), &HA8, &HFD, &H8, &H8F, &HAB, &HF1, &H1F, &H32)
IID_IIdleSettings = iid
End Function
Public Function IID_INetworkSettings() As UUID
'{9F7DEA84-C30B-4245-80B6-00E9F646F1B4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9F7DEA84, CInt(&HC30B), CInt(&H4245), &H80, &HB6, &H0, &HE9, &HF6, &H46, &HF1, &HB4)
IID_INetworkSettings = iid
End Function
Public Function IID_IPrincipal() As UUID
'{D98D51E5-C9B4-496A-A9C1-18980261CF0F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD98D51E5, CInt(&HC9B4), CInt(&H496A), &HA9, &HC1, &H18, &H98, &H2, &H61, &HCF, &HF)
IID_IPrincipal = iid
End Function
Public Function IID_IActionCollection() As UUID
'{02820E19-7B98-4ED2-B2E8-FDCCCEFF619B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2820E19, CInt(&H7B98), CInt(&H4ED2), &HB2, &HE8, &HFD, &HCC, &HCE, &HFF, &H61, &H9B)
IID_IActionCollection = iid
End Function
Public Function IID_IAction() As UUID
'{BAE54997-48B1-4CBE-9965-D6BE263EBEA4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBAE54997, CInt(&H48B1), CInt(&H4CBE), &H99, &H65, &HD6, &HBE, &H26, &H3E, &HBE, &HA4)
IID_IAction = iid
End Function
Public Function IID_IRegisteredTaskCollection() As UUID
'{86627EB4-42A7-41E4-A4D9-AC33A72F2D52}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H86627EB4, CInt(&H42A7), CInt(&H41E4), &HA4, &HD9, &HAC, &H33, &HA7, &H2F, &H2D, &H52)
IID_IRegisteredTaskCollection = iid
End Function
Public Function IID_ITaskService() As UUID
'{2FABA4C7-4DA9-4013-9697-20CC3FD40F85}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2FABA4C7, CInt(&H4DA9), CInt(&H4013), &H96, &H97, &H20, &HCC, &H3F, &HD4, &HF, &H85)
IID_ITaskService = iid
End Function
Public Function IID_ITaskHandler() As UUID
'{839D7762-5121-4009-9234-4F0D19394F04}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H839D7762, CInt(&H5121), CInt(&H4009), &H92, &H34, &H4F, &HD, &H19, &H39, &H4F, &H4)
IID_ITaskHandler = iid
End Function
Public Function IID_ITaskHandlerStatus() As UUID
'{EAEC7A8F-27A0-4DDC-8675-14726A01A38A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEAEC7A8F, CInt(&H27A0), CInt(&H4DDC), &H86, &H75, &H14, &H72, &H6A, &H1, &HA3, &H8A)
IID_ITaskHandlerStatus = iid
End Function
Public Function IID_ITaskVariables() As UUID
'{3E4C9351-D966-4B8B-BB87-CEBA68BB0107}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3E4C9351, CInt(&HD966), CInt(&H4B8B), &HBB, &H87, &HCE, &HBA, &H68, &HBB, &H1, &H7)
IID_ITaskVariables = iid
End Function
Public Function IID_ITaskNamedValuePair() As UUID
'{39038068-2B46-4AFD-8662-7BB6F868D221}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H39038068, CInt(&H2B46), CInt(&H4AFD), &H86, &H62, &H7B, &HB6, &HF8, &H68, &HD2, &H21)
IID_ITaskNamedValuePair = iid
End Function
Public Function IID_ITaskNamedValueCollection() As UUID
'{B4EF826B-63C3-46E4-A504-EF69E4F7EA4D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB4EF826B, CInt(&H63C3), CInt(&H46E4), &HA5, &H4, &HEF, &H69, &HE4, &HF7, &HEA, &H4D)
IID_ITaskNamedValueCollection = iid
End Function
Public Function IID_IIdleTrigger() As UUID
'{D537D2B0-9FB3-4D34-9739-1FF5CE7B1EF3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD537D2B0, CInt(&H9FB3), CInt(&H4D34), &H97, &H39, &H1F, &HF5, &HCE, &H7B, &H1E, &HF3)
IID_IIdleTrigger = iid
End Function
Public Function IID_ILogonTrigger() As UUID
'{72DADE38-FAE4-4B3E-BAF4-5D009AF02B1C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H72DADE38, CInt(&HFAE4), CInt(&H4B3E), &HBA, &HF4, &H5D, &H0, &H9A, &HF0, &H2B, &H1C)
IID_ILogonTrigger = iid
End Function
Public Function IID_ISessionStateChangeTrigger() As UUID
'{754DA71B-4385-4475-9DD9-598294FA3641}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H754DA71B, CInt(&H4385), CInt(&H4475), &H9D, &HD9, &H59, &H82, &H94, &HFA, &H36, &H41)
IID_ISessionStateChangeTrigger = iid
End Function
Public Function IID_IEventTrigger() As UUID
'{D45B0167-9653-4EEF-B94F-0732CA7AF251}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD45B0167, CInt(&H9653), CInt(&H4EEF), &HB9, &H4F, &H7, &H32, &HCA, &H7A, &HF2, &H51)
IID_IEventTrigger = iid
End Function
Public Function IID_ITimeTrigger() As UUID
'{B45747E0-EBA7-4276-9F29-85C5BB300006}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB45747E0, CInt(&HEBA7), CInt(&H4276), &H9F, &H29, &H85, &HC5, &HBB, &H30, &H0, &H6)
IID_ITimeTrigger = iid
End Function
Public Function IID_IDailyTrigger() As UUID
'{126C5CD8-B288-41D5-8DBF-E491446ADC5C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H126C5CD8, CInt(&HB288), CInt(&H41D5), &H8D, &HBF, &HE4, &H91, &H44, &H6A, &HDC, &H5C)
IID_IDailyTrigger = iid
End Function
Public Function IID_IWeeklyTrigger() As UUID
'{5038FC98-82FF-436D-8728-A512A57C9DC1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5038FC98, CInt(&H82FF), CInt(&H436D), &H87, &H28, &HA5, &H12, &HA5, &H7C, &H9D, &HC1)
IID_IWeeklyTrigger = iid
End Function
Public Function IID_IMonthlyTrigger() As UUID
'{97C45EF1-6B02-4A1A-9C0E-1EBFBA1500AC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H97C45EF1, CInt(&H6B02), CInt(&H4A1A), &H9C, &HE, &H1E, &HBF, &HBA, &H15, &H0, &HAC)
IID_IMonthlyTrigger = iid
End Function
Public Function IID_IMonthlyDOWTrigger() As UUID
'{77D025A3-90FA-43AA-B52E-CDA5499B946A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H77D025A3, CInt(&H90FA), CInt(&H43AA), &HB5, &H2E, &HCD, &HA5, &H49, &H9B, &H94, &H6A)
IID_IMonthlyDOWTrigger = iid
End Function
Public Function IID_IBootTrigger() As UUID
'{2A9C35DA-D357-41F4-BBC1-207AC1B1F3CB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2A9C35DA, CInt(&HD357), CInt(&H41F4), &HBB, &HC1, &H20, &H7A, &HC1, &HB1, &HF3, &HCB)
IID_IBootTrigger = iid
End Function
Public Function IID_IRegistrationTrigger() As UUID
'{4C8FEC3A-C218-4E0C-B23D-629024DB91A2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4C8FEC3A, CInt(&HC218), CInt(&H4E0C), &HB2, &H3D, &H62, &H90, &H24, &HDB, &H91, &HA2)
IID_IRegistrationTrigger = iid
End Function
Public Function IID_IExecAction() As UUID
'{4C3D624D-FD6B-49A3-B9B7-09CB3CD3F047}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4C3D624D, CInt(&HFD6B), CInt(&H49A3), &HB9, &HB7, &H9, &HCB, &H3C, &HD3, &HF0, &H47)
IID_IExecAction = iid
End Function
Public Function IID_IExecAction2() As UUID
'{F2A82542-BDA5-4E6B-9143-E2BF4F8987B6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF2A82542, CInt(&HBDA5), CInt(&H4E6B), &H91, &H43, &HE2, &HBF, &H4F, &H89, &H87, &HB6)
IID_IExecAction2 = iid
End Function
Public Function IID_IShowMessageAction() As UUID
'{505E9E68-AF89-46B8-A30F-56162A83D537}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H505E9E68, CInt(&HAF89), CInt(&H46B8), &HA3, &HF, &H56, &H16, &H2A, &H83, &HD5, &H37)
IID_IShowMessageAction = iid
End Function
Public Function IID_IComHandlerAction() As UUID
'{6D2FD252-75C5-4F66-90BA-2A7D8CC3039F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6D2FD252, CInt(&H75C5), CInt(&H4F66), &H90, &HBA, &H2A, &H7D, &H8C, &HC3, &H3, &H9F)
IID_IComHandlerAction = iid
End Function
Public Function IID_IEmailAction() As UUID
'{10F62C64-7E16-4314-A0C2-0C3683F99D40}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10F62C64, CInt(&H7E16), CInt(&H4314), &HA0, &HC2, &HC, &H36, &H83, &HF9, &H9D, &H40)
IID_IEmailAction = iid
End Function
Public Function IID_IPrincipal2() As UUID
'{248919AE-E345-4A6D-8AEB-E0D3165C904E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H248919AE, CInt(&HE345), CInt(&H4A6D), &H8A, &HEB, &HE0, &HD3, &H16, &H5C, &H90, &H4E)
IID_IPrincipal2 = iid
End Function
Public Function IID_ITaskSettings2() As UUID
'{2C05C3F0-6EED-4C05-A15F-ED7D7A98A369}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2C05C3F0, CInt(&H6EED), CInt(&H4C05), &HA1, &H5F, &HED, &H7D, &H7A, &H98, &HA3, &H69)
IID_ITaskSettings2 = iid
End Function
Public Function IID_ITaskSettings3() As UUID
'{0AD9D0D7-0C7F-4EBB-9A5F-D1C648DCA528}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAD9D0D7, CInt(&HC7F), CInt(&H4EBB), &H9A, &H5F, &HD1, &HC6, &H48, &HDC, &HA5, &H28)
IID_ITaskSettings3 = iid
End Function
Public Function IID_IMaintenanceSettings() As UUID
'{A6024FA8-9652-4ADB-A6BF-5CFCD877A7BA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA6024FA8, CInt(&H9652), CInt(&H4ADB), &HA6, &HBF, &H5C, &HFC, &HD8, &H77, &HA7, &HBA)
IID_IMaintenanceSettings = iid
End Function
Public Function IID_IDefaultFolderMenuInitialize() As UUID
'{7690aa79-f8fc-4615-a327-36f7d18f5d91}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7690AA79, CInt(&HF8FC), CInt(&H4615), &HA3, &H27, &H36, &HF7, &HD1, &H8F, &H5D, &H91)
 IID_IDefaultFolderMenuInitialize = iid
End Function
Public Function IID_IInfoBarMessage() As UUID
'{819d1334-9d74-4254-9ac8-dc745ebc5386}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H819D1334, CInt(&H9D74), CInt(&H4254), &H9A, &HC8, &HDC, &H74, &H5E, &HBC, &H53, &H86)
IID_IInfoBarMessage = iid
End Function
Public Function IID_IInfoBarHost() As UUID
'{e38fe0f3-3db0-47ee-a314-25cf7f4bf521}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE38FE0F3, CInt(&H3DB0), CInt(&H47EE), &HA3, &H14, &H25, &HCF, &H7F, &H4B, &HF5, &H21)
IID_IInfoBarHost = iid
End Function
Public Function IID_IBrowserProgressSessionProvider() As UUID
'{18140CBD-AA23-4384-A38D-6A8D3E2BE505}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H18140CBD, CInt(&HAA23), CInt(&H4384), &HA3, &H8D, &H6A, &H8D, &H3E, &H2B, &HE5, &H5)
IID_IBrowserProgressSessionProvider = iid
End Function
Public Function IID_IShellFolder3() As UUID
'{711B2CFD-93D1-422B-BDF4-69BE923F2449}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H711B2CFD, CInt(&H93D1), CInt(&H422B), &HBD, &HF4, &H69, &HBE, &H92, &H3F, &H24, &H49)
IID_IShellFolder3 = iid
End Function
Public Function IID_IBrowserProgressConnection() As UUID
'{20174539-b2c7-4ec7-970b-04201f9cdbad}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H20174539, CInt(&HB2C7), CInt(&H4EC7), &H97, &HB, &H4, &H20, &H1F, &H9C, &HDB, &HAD)
 IID_IBrowserProgressConnection = iid
End Function
Public Function IID_IBrowserProgressAggregator() As UUID
'{5EA8EEC4-C34B-4DE0-9B56-0E15FD8C8F80}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5EA8EEC4, CInt(&HC34B), CInt(&H4DE0), &H9B, &H56, &HE, &H15, &HFD, &H8C, &H8F, &H80)
 IID_IBrowserProgressAggregator = iid
End Function
Public Function IID_IGlobalInterfaceTable() As UUID
'{00000146-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H146, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IGlobalInterfaceTable = iid
End Function
Public Function IID_IManipulationProcessor() As UUID
'{A22AC519-8300-48a0-BEF4-F1BE8737DBA4}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA22AC519, CInt(&H8300), CInt(&H48A0), &HBE, &HF4, &HF1, &HBE, &H87, &H37, &HDB, &HA4)
 IID_IManipulationProcessor = iid
End Function
Public Function IID_IInertiaProcessor() As UUID
'{18b00c6d-c5ee-41b1-90a9-9d4a929095ad}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H18B00C6D, CInt(&HC5EE), CInt(&H41B1), &H90, &HA9, &H9D, &H4A, &H92, &H90, &H95, &HAD)
 IID_IInertiaProcessor = iid
End Function
Public Function IID_IManipulationEvents() As UUID
'{4f62c8da-9c53-4b22-93df-927a862bbb03}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4F62C8DA, CInt(&H9C53), CInt(&H4B22), &H93, &HDF, &H92, &H7A, &H86, &H2B, &HBB, &H3)
 IID_IManipulationEvents = iid
End Function
Public Function IID_IFilter() As UUID
'{89BCB740-6119-101A-BCB7-00DD010655AF}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H89BCB740, CInt(&H6119), CInt(&H101A), &HBC, &HB7, &H0, &HDD, &H1, &H6, &H55, &HAF)
IID_IFilter = iid
End Function
Public Function IID_ILoadFilter() As UUID
'{c7310722-ac80-11d1-8df3-00c04fb6ef4f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC7310722, CInt(&HAC80), CInt(&H11D1), &H8D, &HF3, &H0, &HC0, &H4F, &HB6, &HEF, &H4F)
IID_ILoadFilter = iid
End Function
Public Function IID_ILoadFilterWithPrivateComActivation() As UUID
'{40BDBD34-780B-48D3-9BB6-12EBD4AD2E75}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H40BDBD34, CInt(&H780B), CInt(&H48D3), &H9B, &HB6, &H12, &HEB, &HD4, &HAD, &H2E, &H75)
IID_ILoadFilterWithPrivateComActivation = iid
End Function
Public Function IID_IUrlAccessor() As UUID
'{0b63e318-9ccc-11d0-bcdb-00805fccce04}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB63E318, CInt(&H9CCC), CInt(&H11D0), &HBC, &HDB, &H0, &H80, &H5F, &HCC, &HCE, &H4)
IID_IUrlAccessor = iid
End Function
Public Function IID_IUrlAccessor2() As UUID
'{c7310734-ac80-11d1-8df3-00c04fb6ef4f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC7310734, CInt(&HAC80), CInt(&H11D1), &H8D, &HF3, &H0, &HC0, &H4F, &HB6, &HEF, &H4F)
IID_IUrlAccessor2 = iid
End Function
Public Function IID_IUrlAccessor3() As UUID
'{6FBC7005-0455-4874-B8FF-7439450241A3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FBC7005, CInt(&H455), CInt(&H4874), &HB8, &HFF, &H74, &H39, &H45, &H2, &H41, &HA3)
IID_IUrlAccessor3 = iid
End Function
Public Function IID_IUrlAccessor4() As UUID
'{5CC51041-C8D2-41d7-BCA3-9E9E286297DC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5CC51041, CInt(&HC8D2), CInt(&H41D7), &HBC, &HA3, &H9E, &H9E, &H28, &H62, &H97, &HDC)
IID_IUrlAccessor4 = iid
End Function
Public Function IID_ISearchProtocolThreadContext() As UUID
'{c73106e1-ac80-11d1-8df3-00c04fb6ef4f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC73106E1, CInt(&HAC80), CInt(&H11D1), &H8D, &HF3, &H0, &HC0, &H4F, &HB6, &HEF, &H4F)
IID_ISearchProtocolThreadContext = iid
End Function
Public Function IID_IOpLockStatus() As UUID
'{c731065d-ac80-11d1-8df3-00c04fb6ef4f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC731065D, CInt(&HAC80), CInt(&H11D1), &H8D, &HF3, &H0, &HC0, &H4F, &HB6, &HEF, &H4F)
IID_IOpLockStatus = iid
End Function
Public Function IID_ISearchProtocol() As UUID
'{c73106ba-ac80-11d1-8df3-00c04fb6ef4f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC73106BA, CInt(&HAC80), CInt(&H11D1), &H8D, &HF3, &H0, &HC0, &H4F, &HB6, &HEF, &H4F)
IID_ISearchProtocol = iid
End Function
Public Function IID_ISearchProtocol2() As UUID
'{7789F0B2-B5B2-4722-8B65-5DBD150697A9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7789F0B2, CInt(&HB5B2), CInt(&H4722), &H8B, &H65, &H5D, &HBD, &H15, &H6, &H97, &HA9)
IID_ISearchProtocol2 = iid
End Function
Public Function IID_IProtocolHandlerSite() As UUID
'{0b63e385-9ccc-11d0-bcdb-00805fccce04}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB63E385, CInt(&H9CCC), CInt(&H11D0), &HBC, &HDB, &H0, &H80, &H5F, &HCC, &HCE, &H4)
IID_IProtocolHandlerSite = iid
End Function
Public Function IID_ISearchRoot() As UUID
'{04C18CCF-1F57-4CBD-88CC-3900F5195CE3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4C18CCF, CInt(&H1F57), CInt(&H4CBD), &H88, &HCC, &H39, &H0, &HF5, &H19, &H5C, &HE3)
IID_ISearchRoot = iid
End Function
Public Function IID_IEnumSearchRoots() As UUID
'{AB310581-AC80-11D1-8DF3-00C04FB6EF52}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB310581, CInt(&HAC80), CInt(&H11D1), &H8D, &HF3, &H0, &HC0, &H4F, &HB6, &HEF, &H52)
IID_IEnumSearchRoots = iid
End Function
Public Function IID_ISearchScopeRule() As UUID
'{AB310581-AC80-11D1-8DF3-00C04FB6EF53}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB310581, CInt(&HAC80), CInt(&H11D1), &H8D, &HF3, &H0, &HC0, &H4F, &HB6, &HEF, &H53)
IID_ISearchScopeRule = iid
End Function
Public Function IID_IEnumSearchScopeRules() As UUID
'{AB310581-AC80-11D1-8DF3-00C04FB6EF54}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB310581, CInt(&HAC80), CInt(&H11D1), &H8D, &HF3, &H0, &HC0, &H4F, &HB6, &HEF, &H54)
IID_IEnumSearchScopeRules = iid
End Function
Public Function IID_ISearchCrawlScopeManager() As UUID
'{AB310581-AC80-11D1-8DF3-00C04FB6EF55}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB310581, CInt(&HAC80), CInt(&H11D1), &H8D, &HF3, &H0, &HC0, &H4F, &HB6, &HEF, &H55)
IID_ISearchCrawlScopeManager = iid
End Function
Public Function IID_ISearchCrawlScopeManager2() As UUID
'{6292F7AD-4E19-4717-A534-8FC22BCD5CCD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6292F7AD, CInt(&H4E19), CInt(&H4717), &HA5, &H34, &H8F, &HC2, &H2B, &HCD, &H5C, &HCD)
IID_ISearchCrawlScopeManager2 = iid
End Function
Public Function IID_ISearchItemsChangedSink() As UUID
'{AB310581-AC80-11D1-8DF3-00C04FB6EF58}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB310581, CInt(&HAC80), CInt(&H11D1), &H8D, &HF3, &H0, &HC0, &H4F, &HB6, &HEF, &H58)
IID_ISearchItemsChangedSink = iid
End Function
Public Function IID_ISearchPersistentItemsChangedSink() As UUID
'{A2FFDF9B-4758-4F84-B729-DF81A1A0612F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA2FFDF9B, CInt(&H4758), CInt(&H4F84), &HB7, &H29, &HDF, &H81, &HA1, &HA0, &H61, &H2F)
IID_ISearchPersistentItemsChangedSink = iid
End Function
Public Function IID_ISearchViewChangedSink() As UUID
'{AB310581-AC80-11D1-8DF3-00C04FB6EF65}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB310581, CInt(&HAC80), CInt(&H11D1), &H8D, &HF3, &H0, &HC0, &H4F, &HB6, &HEF, &H65)
IID_ISearchViewChangedSink = iid
End Function
Public Function IID_ISearchNotifyInlineSite() As UUID
'{B5702E61-E75C-4B64-82A1-6CB4F832FCCF}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB5702E61, CInt(&HE75C), CInt(&H4B64), &H82, &HA1, &H6C, &HB4, &HF8, &H32, &HFC, &HCF)
IID_ISearchNotifyInlineSite = iid
End Function
Public Function IID_ISearchCatalogManager() As UUID
'{AB310581-AC80-11D1-8DF3-00C04FB6EF50}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB310581, CInt(&HAC80), CInt(&H11D1), &H8D, &HF3, &H0, &HC0, &H4F, &HB6, &HEF, &H50)
IID_ISearchCatalogManager = iid
End Function
Public Function IID_ISearchCatalogManager2() As UUID
'{7AC3286D-4D1D-4817-84FC-C1C85E3AF0D9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7AC3286D, CInt(&H4D1D), CInt(&H4817), &H84, &HFC, &HC1, &HC8, &H5E, &H3A, &HF0, &HD9)
IID_ISearchCatalogManager2 = iid
End Function
Public Function IID_ISearchQueryHelper() As UUID
'{AB310581-AC80-11D1-8DF3-00C04FB6EF63}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB310581, CInt(&HAC80), CInt(&H11D1), &H8D, &HF3, &H0, &HC0, &H4F, &HB6, &HEF, &H63)
IID_ISearchQueryHelper = iid
End Function
Public Function IID_IRowsetPrioritization() As UUID
'{42811652-079D-481B-87A2-09A69ECC5F44}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H42811652, CInt(&H79D), CInt(&H481B), &H87, &HA2, &H9, &HA6, &H9E, &HCC, &H5F, &H44)
IID_IRowsetPrioritization = iid
End Function
Public Function IID_IRowsetEvents() As UUID
'{1551AEA5-5D66-4B11-86F5-D5634CB211B9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1551AEA5, CInt(&H5D66), CInt(&H4B11), &H86, &HF5, &HD5, &H63, &H4C, &HB2, &H11, &HB9)
IID_IRowsetEvents = iid
End Function
Public Function IID_ISearchManager() As UUID
'{AB310581-AC80-11D1-8DF3-00C04FB6EF69}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB310581, CInt(&HAC80), CInt(&H11D1), &H8D, &HF3, &H0, &HC0, &H4F, &HB6, &HEF, &H69)
IID_ISearchManager = iid
End Function
Public Function IID_ISearchManager2() As UUID
'{DBAB3F73-DB19-4A79-BFC0-A61A93886DDF}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDBAB3F73, CInt(&HDB19), CInt(&H4A79), &HBF, &HC0, &HA6, &H1A, &H93, &H88, &H6D, &HDF)
IID_ISearchManager2 = iid
End Function
Public Function IID_ISearchLanguageSupport() As UUID
'{24C3CBAA-EBC1-491a-9EF1-9F6D8DEB1B8F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H24C3CBAA, CInt(&HEBC1), CInt(&H491A), &H9E, &HF1, &H9F, &H6D, &H8D, &HEB, &H1B, &H8F)
IID_ISearchLanguageSupport = iid
End Function
Public Function IID_ISecurityInformation() As UUID
'{965FC360-16FF-11d0-91CB-00AA00BBB723}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H965FC360, CInt(&H16FF), CInt(&H11D0), &H91, &HCB, &H0, &HAA, &H0, &HBB, &HB7, &H23)
IID_ISecurityInformation = iid
End Function
Public Function IID_ISecurityInformation2() As UUID
'{c3ccfdb4-6f88-11d2-a3ce-00c04fb1782a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC3CCFDB4, CInt(&H6F88), CInt(&H11D2), &HA3, &HCE, &H0, &HC0, &H4F, &HB1, &H78, &H2A)
IID_ISecurityInformation2 = iid
End Function
Public Function IID_IEffectivePermission() As UUID
'{3853DC76-9F35-407c-88A1-D19344365FBC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3853DC76, CInt(&H9F35), CInt(&H407C), &H88, &HA1, &HD1, &H93, &H44, &H36, &H5F, &HBC)
IID_IEffectivePermission = iid
End Function
Public Function IID_ISecurityObjectTypeInfo() As UUID
'{FC3066EB-79EF-444b-9111-D18A75EBF2FA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFC3066EB, CInt(&H79EF), CInt(&H444B), &H91, &H11, &HD1, &H8A, &H75, &HEB, &HF2, &HFA)
IID_ISecurityObjectTypeInfo = iid
End Function
Public Function IID_ISecurityInformation3() As UUID
'{E2CDC9CC-31BD-4f8f-8C8B-B641AF516A1A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE2CDC9CC, CInt(&H31BD), CInt(&H4F8F), &H8C, &H8B, &HB6, &H41, &HAF, &H51, &H6A, &H1A)
IID_ISecurityInformation3 = iid
End Function
Public Function IID_ISecurityInformation4() As UUID
'{EA961070-CD14-4621-ACE4-F63C03E583E4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEA961070, CInt(&HCD14), CInt(&H4621), &HAC, &HE4, &HF6, &H3C, &H3, &HE5, &H83, &HE4)
IID_ISecurityInformation4 = iid
End Function
Public Function IID_IEffectivePermission2() As UUID
'{941FABCA-DD47-4FCA-90BB-B0E10255F20D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H941FABCA, CInt(&HDD47), CInt(&H4FCA), &H90, &HBB, &HB0, &HE1, &H2, &H55, &HF2, &HD)
IID_IEffectivePermission2 = iid
End Function
Public Function IID_IWscProduct() As UUID
'{8C38232E-3A45-4A27-92B0-1A16A975F669}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8C38232E, CInt(&H3A45), CInt(&H4A27), &H92, &HB0, &H1A, &H16, &HA9, &H75, &HF6, &H69)
IID_IWscProduct = iid
End Function
Public Function IID_IWscProduct2() As UUID
'{F896CA54-FE09-4403-86D4-23CB488D81D8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF896CA54, CInt(&HFE09), CInt(&H4403), &H86, &HD4, &H23, &HCB, &H48, &H8D, &H81, &HD8)
IID_IWscProduct2 = iid
End Function
Public Function IID_IWscProduct3() As UUID
'{55536524-D1D1-4726-8C7C-04996A1904E7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H55536524, CInt(&HD1D1), CInt(&H4726), &H8C, &H7C, &H4, &H99, &H6A, &H19, &H4, &HE7)
IID_IWscProduct3 = iid
End Function
Public Function IID_IWSCProductList() As UUID
'{722A338C-6E8E-4E72-AC27-1417FB0C81C2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H722A338C, CInt(&H6E8E), CInt(&H4E72), &HAC, &H27, &H14, &H17, &HFB, &HC, &H81, &HC2)
IID_IWSCProductList = iid
End Function
Public Function IID_IWSCDefaultProduct() As UUID
'{0476d69c-f21a-11e5-9ce9-5e5517507c66}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H476D69C, CInt(&HF21A), CInt(&H11E5), &H9C, &HE9, &H5E, &H55, &H17, &H50, &H7C, &H66)
IID_IWSCDefaultProduct = iid
End Function
Public Function IID_IEnumBackgroundCopyFiles() As UUID
'{ca51e165-c365-424c-8d41-24aaa4ff3c40}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCA51E165, CInt(&HC365), CInt(&H424C), &H8D, &H41, &H24, &HAA, &HA4, &HFF, &H3C, &H40)
IID_IEnumBackgroundCopyFiles = iid
End Function
Public Function IID_IBackgroundCopyError() As UUID
'{19c613a0-fcb8-4f28-81ae-897c3d078f81}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H19C613A0, CInt(&HFCB8), CInt(&H4F28), &H81, &HAE, &H89, &H7C, &H3D, &H7, &H8F, &H81)
IID_IBackgroundCopyError = iid
End Function
Public Function IID_IBackgroundCopyJob() As UUID
'{37668d37-507e-4160-9316-26306d150b12}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H37668D37, CInt(&H507E), CInt(&H4160), &H93, &H16, &H26, &H30, &H6D, &H15, &HB, &H12)
IID_IBackgroundCopyJob = iid
End Function
Public Function IID_IEnumBackgroundCopyJobs() As UUID
'{1af4f612-3b71-466f-8f58-7b6f73ac57ad}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1AF4F612, CInt(&H3B71), CInt(&H466F), &H8F, &H58, &H7B, &H6F, &H73, &HAC, &H57, &HAD)
IID_IEnumBackgroundCopyJobs = iid
End Function
Public Function IID_IBackgroundCopyCallback() As UUID
'{97ea99c7-0186-4ad4-8df9-c5b4e0ed6b22}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H97EA99C7, CInt(&H186), CInt(&H4AD4), &H8D, &HF9, &HC5, &HB4, &HE0, &HED, &H6B, &H22)
IID_IBackgroundCopyCallback = iid
End Function
Public Function IID_IBackgroundCopyManager() As UUID
'{5ce34c0d-0dc9-4c1f-897c-daa1b78cee7c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5CE34C0D, CInt(&HDC9), CInt(&H4C1F), &H89, &H7C, &HDA, &HA1, &HB7, &H8C, &HEE, &H7C)
IID_IBackgroundCopyManager = iid
End Function
Public Function IID_IBackgroundCopyJob2() As UUID
'{54b50739-686f-45eb-9dff-d6a9a0faa9af}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H54B50739, CInt(&H686F), CInt(&H45EB), &H9D, &HFF, &HD6, &HA9, &HA0, &HFA, &HA9, &HAF)
IID_IBackgroundCopyJob2 = iid
End Function
Public Function IID_IBitsPeerCacheRecord() As UUID
'{659cdeaf-489e-11d9-a9cd-000d56965251}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H659CDEAF, CInt(&H489E), CInt(&H11D9), &HA9, &HCD, &H0, &HD, &H56, &H96, &H52, &H51)
IID_IBitsPeerCacheRecord = iid
End Function
Public Function IID_IEnumBitsPeerCacheRecords() As UUID
'{659cdea4-489e-11d9-a9cd-000d56965251}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H659CDEA4, CInt(&H489E), CInt(&H11D9), &HA9, &HCD, &H0, &HD, &H56, &H96, &H52, &H51)
IID_IEnumBitsPeerCacheRecords = iid
End Function
Public Function IID_IBitsPeer() As UUID
'{659cdea2-489e-11d9-a9cd-000d56965251}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H659CDEA2, CInt(&H489E), CInt(&H11D9), &HA9, &HCD, &H0, &HD, &H56, &H96, &H52, &H51)
IID_IBitsPeer = iid
End Function
Public Function IID_IEnumBitsPeers() As UUID
'{659cdea5-489e-11d9-a9cd-000d56965251}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H659CDEA5, CInt(&H489E), CInt(&H11D9), &HA9, &HCD, &H0, &HD, &H56, &H96, &H52, &H51)
IID_IEnumBitsPeers = iid
End Function
Public Function IID_IBitsPeerCacheAdministration() As UUID
'{659cdead-489e-11d9-a9cd-000d56965251}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H659CDEAD, CInt(&H489E), CInt(&H11D9), &HA9, &HCD, &H0, &HD, &H56, &H96, &H52, &H51)
IID_IBitsPeerCacheAdministration = iid
End Function
Public Function IID_IBackgroundCopyJob4() As UUID
'{659cdeae-489e-11d9-a9cd-000d56965251}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H659CDEAE, CInt(&H489E), CInt(&H11D9), &HA9, &HCD, &H0, &HD, &H56, &H96, &H52, &H51)
IID_IBackgroundCopyJob4 = iid
End Function
Public Function IID_IBackgroundCopyFile3() As UUID
'{659cdeaa-489e-11d9-a9cd-000d56965251}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H659CDEAA, CInt(&H489E), CInt(&H11D9), &HA9, &HCD, &H0, &HD, &H56, &H96, &H52, &H51)
IID_IBackgroundCopyFile3 = iid
End Function
Public Function IID_IBackgroundCopyCallback2() As UUID
'{659cdeac-489e-11d9-a9cd-000d56965251}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H659CDEAC, CInt(&H489E), CInt(&H11D9), &HA9, &HCD, &H0, &HD, &H56, &H96, &H52, &H51)
IID_IBackgroundCopyCallback2 = iid
End Function
Public Function IID_IBitsTokenOptions() As UUID
'{9a2584c3-f7d2-457a-9a5e-22b67bffc7d2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9A2584C3, CInt(&HF7D2), CInt(&H457A), &H9A, &H5E, &H22, &HB6, &H7B, &HFF, &HC7, &HD2)
IID_IBitsTokenOptions = iid
End Function
Public Function IID_IBackgroundCopyFile4() As UUID
'{ef7e0655-7888-4960-b0e5-730846e03492}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEF7E0655, CInt(&H7888), CInt(&H4960), &HB0, &HE5, &H73, &H8, &H46, &HE0, &H34, &H92)
IID_IBackgroundCopyFile4 = iid
End Function
Public Function IID_IBackgroundCopyJob5() As UUID
'{E847030C-BBBA-4657-AF6D-484AA42BF1FE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE847030C, CInt(&HBBBA), CInt(&H4657), &HAF, &H6D, &H48, &H4A, &HA4, &H2B, &HF1, &HFE)
IID_IBackgroundCopyJob5 = iid
End Function
Public Function IID_IBackgroundCopyFile5() As UUID
'{85C1657F-DAFC-40E8-8834-DF18EA25717E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H85C1657F, CInt(&HDAFC), CInt(&H40E8), &H88, &H34, &HDF, &H18, &HEA, &H25, &H71, &H7E)
IID_IBackgroundCopyFile5 = iid
End Function
Public Function IID_IBackgroundCopyCallback3() As UUID
'{98c97bd2-e32b-4ad8-a528-95fd8b16bd42}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H98C97BD2, CInt(&HE32B), CInt(&H4AD8), &HA5, &H28, &H95, &HFD, &H8B, &H16, &HBD, &H42)
IID_IBackgroundCopyCallback3 = iid
End Function
Public Function IID_IBackgroundCopyFile6() As UUID
'{CF6784F7-D677-49FD-9368-CB47AEE9D1AD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCF6784F7, CInt(&HD677), CInt(&H49FD), &H93, &H68, &HCB, &H47, &HAE, &HE9, &HD1, &HAD)
IID_IBackgroundCopyFile6 = iid
End Function
Public Function IID_IBackgroundCopyJobHttpOptions2() As UUID
'{B591A192-A405-4FC3-8323-4C5C542578FC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB591A192, CInt(&HA405), CInt(&H4FC3), &H83, &H23, &H4C, &H5C, &H54, &H25, &H78, &HFC)
IID_IBackgroundCopyJobHttpOptions2 = iid
End Function
Public Function IID_IBackgroundCopyServerCertificateValidationCallback() As UUID
'{4CEC0D02-DEF7-4158-813A-C32A46945FF7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4CEC0D02, CInt(&HDEF7), CInt(&H4158), &H81, &H3A, &HC3, &H2A, &H46, &H94, &H5F, &HF7)
IID_IBackgroundCopyServerCertificateValidationCallback = iid
End Function
Public Function IID_IBackgroundCopyJobHttpOptions3() As UUID
'{8A9263D3-FD4C-4EDA-9B28-30132A4D4E3C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8A9263D3, CInt(&HFD4C), CInt(&H4EDA), &H9B, &H28, &H30, &H13, &H2A, &H4D, &H4E, &H3C)
IID_IBackgroundCopyJobHttpOptions3 = iid
End Function
Public Function IID_IBITSExtensionSetup() As UUID
'{29cfbbf7-09e4-4b97-b0bc-f2287e3d8eb3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H29CFBBF7, CInt(&H9E4), CInt(&H4B97), &HB0, &HBC, &HF2, &H28, &H7E, &H3D, &H8E, &HB3)
IID_IBITSExtensionSetup = iid
End Function
Public Function IID_IBITSExtensionSetupFactory() As UUID
'{d5d2d542-5503-4e64-8b48-72ef91a32ee1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD5D2D542, CInt(&H5503), CInt(&H4E64), &H8B, &H48, &H72, &HEF, &H91, &HA3, &H2E, &HE1)
IID_IBITSExtensionSetupFactory = iid
End Function
Public Function IID_IEnumBackgroundCopyJobs1() As UUID
'{8baeba9d-8f1c-42c4-b82c-09ae79980d25}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8BAEBA9D, CInt(&H8F1C), CInt(&H42C4), &HB8, &H2C, &H9, &HAE, &H79, &H98, &HD, &H25)
IID_IEnumBackgroundCopyJobs1 = iid
End Function
Public Function IID_IBackgroundCopyGroup() As UUID
'{1ded80a7-53ea-424f-8a04-17fea9adc4f5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1DED80A7, CInt(&H53EA), CInt(&H424F), &H8A, &H4, &H17, &HFE, &HA9, &HAD, &HC4, &HF5)
IID_IBackgroundCopyGroup = iid
End Function
Public Function IID_IEnumBackgroundCopyGroups() As UUID
'{d993e603-4aa4-47c5-8665-c20d39c2ba4f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD993E603, CInt(&H4AA4), CInt(&H47C5), &H86, &H65, &HC2, &HD, &H39, &HC2, &HBA, &H4F)
IID_IEnumBackgroundCopyGroups = iid
End Function
Public Function IID_IBackgroundCopyCallback1() As UUID
'{084f6593-3800-4e08-9b59-99fa59addf82}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H84F6593, CInt(&H3800), CInt(&H4E08), &H9B, &H59, &H99, &HFA, &H59, &HAD, &HDF, &H82)
IID_IBackgroundCopyCallback1 = iid
End Function
Public Function IID_IBackgroundCopyQMgr() As UUID
'{16f41c69-09f5-41d2-8cd8-3c08c47bc8a8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H16F41C69, CInt(&H9F5), CInt(&H41D2), &H8C, &HD8, &H3C, &H8, &HC4, &H7B, &HC8, &HA8)
IID_IBackgroundCopyQMgr = iid
End Function
Public Function IID_IQMgr() As UUID
'{16f41c69-09f5-41d2-8cd8-3c08c47bc8a8}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H16F41C69, CInt(&H9F5), CInt(&H41D2), &H8C, &HD8, &H3C, &H8, &HC4, &H7B, &HC8, &HA8)
 IID_IQMgr = iid
End Function
Public Function IID_IWMDMMetaData() As UUID
'{EC3B0663-0951-460a-9A80-0DCEED3C043C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEC3B0663, CInt(&H951), CInt(&H460A), &H9A, &H80, &HD, &HCE, &HED, &H3C, &H4, &H3C)
IID_IWMDMMetaData = iid
End Function
Public Function IID_IWMDeviceManager() As UUID
'{1DCB3A00-33ED-11d3-8470-00C04F79DBC0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1DCB3A00, CInt(&H33ED), CInt(&H11D3), &H84, &H70, &H0, &HC0, &H4F, &H79, &HDB, &HC0)
IID_IWMDeviceManager = iid
End Function
Public Function IID_IWMDeviceManager2() As UUID
'{923E5249-8731-4c5b-9B1C-B8B60B6E46AF}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H923E5249, CInt(&H8731), CInt(&H4C5B), &H9B, &H1C, &HB8, &HB6, &HB, &H6E, &H46, &HAF)
IID_IWMDeviceManager2 = iid
End Function
Public Function IID_IWMDeviceManager3() As UUID
'{af185c41-100d-46ed-be2e-9ce8c44594ef}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAF185C41, CInt(&H100D), CInt(&H46ED), &HBE, &H2E, &H9C, &HE8, &HC4, &H45, &H94, &HEF)
IID_IWMDeviceManager3 = iid
End Function
Public Function IID_IWMDMStorageGlobals() As UUID
'{1DCB3A07-33ED-11d3-8470-00C04F79DBC0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1DCB3A07, CInt(&H33ED), CInt(&H11D3), &H84, &H70, &H0, &HC0, &H4F, &H79, &HDB, &HC0)
IID_IWMDMStorageGlobals = iid
End Function
Public Function IID_IWMDMStorage() As UUID
'{1DCB3A06-33ED-11d3-8470-00C04F79DBC0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1DCB3A06, CInt(&H33ED), CInt(&H11D3), &H84, &H70, &H0, &HC0, &H4F, &H79, &HDB, &HC0)
IID_IWMDMStorage = iid
End Function
Public Function IID_IWMDMStorage2() As UUID
'{1ED5A144-5CD5-4683-9EFF-72CBDB2D9533}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1ED5A144, CInt(&H5CD5), CInt(&H4683), &H9E, &HFF, &H72, &HCB, &HDB, &H2D, &H95, &H33)
IID_IWMDMStorage2 = iid
End Function
Public Function IID_IWMDMStorage3() As UUID
'{97717EEA-926A-464e-96A4-247B0216026E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H97717EEA, CInt(&H926A), CInt(&H464E), &H96, &HA4, &H24, &H7B, &H2, &H16, &H2, &H6E)
IID_IWMDMStorage3 = iid
End Function
Public Function IID_IWMDMStorage4() As UUID
'{c225bac5-a03a-40b8-9a23-91cf478c64a6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC225BAC5, CInt(&HA03A), CInt(&H40B8), &H9A, &H23, &H91, &HCF, &H47, &H8C, &H64, &HA6)
IID_IWMDMStorage4 = iid
End Function
Public Function IID_IWMDMOperation() As UUID
'{1DCB3A0B-33ED-11d3-8470-00C04F79DBC0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1DCB3A0B, CInt(&H33ED), CInt(&H11D3), &H84, &H70, &H0, &HC0, &H4F, &H79, &HDB, &HC0)
IID_IWMDMOperation = iid
End Function
Public Function IID_IWMDMOperation2() As UUID
'{33445B48-7DF7-425c-AD8F-0FC6D82F9F75}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H33445B48, CInt(&H7DF7), CInt(&H425C), &HAD, &H8F, &HF, &HC6, &HD8, &H2F, &H9F, &H75)
IID_IWMDMOperation2 = iid
End Function
Public Function IID_IWMDMOperation3() As UUID
'{d1f9b46a-9ca8-46d8-9d0f-1ec9bae54919"}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD1F9B46A, CInt(&H9CA8), CInt(&H46D8), &H9D, &HF, &H1E, &HC9, &HBA, &HE5, &H49, &H19)
IID_IWMDMOperation3 = iid
End Function
Public Function IID_IWMDMProgress() As UUID
'{1DCB3A0C-33ED-11d3-8470-00C04F79DBC0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1DCB3A0C, CInt(&H33ED), CInt(&H11D3), &H84, &H70, &H0, &HC0, &H4F, &H79, &HDB, &HC0)
IID_IWMDMProgress = iid
End Function
Public Function IID_IWMDMProgress2() As UUID
'{3A43F550-B383-4e92-B04A-E6BBC660FEFC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3A43F550, CInt(&HB383), CInt(&H4E92), &HB0, &H4A, &HE6, &HBB, &HC6, &H60, &HFE, &HFC)
IID_IWMDMProgress2 = iid
End Function
Public Function IID_IWMDMProgress3() As UUID
'{21DE01CB-3BB4-4929-B21A-17AF3F80F658}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H21DE01CB, CInt(&H3BB4), CInt(&H4929), &HB2, &H1A, &H17, &HAF, &H3F, &H80, &HF6, &H58)
IID_IWMDMProgress3 = iid
End Function
Public Function IID_IWMDMDevice() As UUID
'{1DCB3A02-33ED-11d3-8470-00C04F79DBC0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1DCB3A02, CInt(&H33ED), CInt(&H11D3), &H84, &H70, &H0, &HC0, &H4F, &H79, &HDB, &HC0)
IID_IWMDMDevice = iid
End Function
Public Function IID_IWMDMDevice2() As UUID
'{E34F3D37-9D67-4fc1-9252-62D28B2F8B55}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE34F3D37, CInt(&H9D67), CInt(&H4FC1), &H92, &H52, &H62, &HD2, &H8B, &H2F, &H8B, &H55)
IID_IWMDMDevice2 = iid
End Function
Public Function IID_IWMDMDevice3() As UUID
'{6c03e4fe-05db-4dda-9e3c-06233a6d5d65}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6C03E4FE, CInt(&H5DB), CInt(&H4DDA), &H9E, &H3C, &H6, &H23, &H3A, &H6D, &H5D, &H65)
IID_IWMDMDevice3 = iid
End Function
Public Function IID_IWMDMDeviceSession() As UUID
'{82af0a65-9d96-412c-83e5-3c43e4b06cc7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H82AF0A65, CInt(&H9D96), CInt(&H412C), &H83, &HE5, &H3C, &H43, &HE4, &HB0, &H6C, &HC7)
IID_IWMDMDeviceSession = iid
End Function
Public Function IID_IWMDMEnumDevice() As UUID
'{1DCB3A01-33ED-11d3-8470-00C04F79DBC0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1DCB3A01, CInt(&H33ED), CInt(&H11D3), &H84, &H70, &H0, &HC0, &H4F, &H79, &HDB, &HC0)
IID_IWMDMEnumDevice = iid
End Function
Public Function IID_IWMDMDeviceControl() As UUID
'{1DCB3A04-33ED-11d3-8470-00C04F79DBC0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1DCB3A04, CInt(&H33ED), CInt(&H11D3), &H84, &H70, &H0, &HC0, &H4F, &H79, &HDB, &HC0)
IID_IWMDMDeviceControl = iid
End Function
Public Function IID_IWMDMEnumStorage() As UUID
'{1DCB3A05-33ED-11d3-8470-00C04F79DBC0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1DCB3A05, CInt(&H33ED), CInt(&H11D3), &H84, &H70, &H0, &HC0, &H4F, &H79, &HDB, &HC0)
IID_IWMDMEnumStorage = iid
End Function
Public Function IID_IWMDMStorageControl() As UUID
'{1DCB3A08-33ED-11d3-8470-00C04F79DBC0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1DCB3A08, CInt(&H33ED), CInt(&H11D3), &H84, &H70, &H0, &HC0, &H4F, &H79, &HDB, &HC0)
IID_IWMDMStorageControl = iid
End Function
Public Function IID_IWMDMStorageControl2() As UUID
'{972C2E88-BD6C-4125-8E09-84F837E637B6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H972C2E88, CInt(&HBD6C), CInt(&H4125), &H8E, &H9, &H84, &HF8, &H37, &HE6, &H37, &HB6)
IID_IWMDMStorageControl2 = iid
End Function
Public Function IID_IWMDMStorageControl3() As UUID
'{B3266365-D4F3-4696-8D53-BD27EC60993A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB3266365, CInt(&HD4F3), CInt(&H4696), &H8D, &H53, &HBD, &H27, &HEC, &H60, &H99, &H3A)
IID_IWMDMStorageControl3 = iid
End Function
Public Function IID_IWMDMObjectInfo() As UUID
'{1DCB3A09-33ED-11d3-8470-00C04F79DBC0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1DCB3A09, CInt(&H33ED), CInt(&H11D3), &H84, &H70, &H0, &HC0, &H4F, &H79, &HDB, &HC0)
IID_IWMDMObjectInfo = iid
End Function
Public Function IID_IWMDMRevoked() As UUID
'{EBECCEDB-88EE-4e55-B6A4-8D9F07D696AA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEBECCEDB, CInt(&H88EE), CInt(&H4E55), &HB6, &HA4, &H8D, &H9F, &H7, &HD6, &H96, &HAA)
IID_IWMDMRevoked = iid
End Function
Public Function IID_IWMDMNotification() As UUID
'{3F5E95C0-0F43-4ed4-93D2-C89A45D59B81}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3F5E95C0, CInt(&HF43), CInt(&H4ED4), &H93, &HD2, &HC8, &H9A, &H45, &HD5, &H9B, &H81)
IID_IWMDMNotification = iid
End Function



Public Function IID_IStream() As UUID
'{0000000C-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IStream = iid
End Function

Public Function IID_IUnknown() As UUID
'"{00000000-0000-0000-C000-000000000046}"
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H0, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
  IID_IUnknown = iid

End Function



Public Function BHID_AssociationArray() As UUID
'DEFINE_GUID(BHID_AssociationArray, 0xBEA9EF17, 0x82F1, 0x4F60, 0x92,0x84, 0x4F,0x8D,0xB7,0x5C,0x3B,0xE9)
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBEA9EF17, &H82F1, &H4F60, &H92, &H84, &H4F, &H8D, &HB7, &H5C, &H3B, &HE9)
  BHID_AssociationArray = iid
End Function

Public Function BHID_SFUIObject() As UUID
'DEFINE_GUID(BHID_SFUIObject,  0x3981E225, 0xF559, 0x11D3, 0x8E,0x3A, 0x00,0xC0,0x4F,0x68,0x37,0xD5);
'{3981e225-f559-11d3-8e3a-00c04f6837d5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3981E225, &HF559, &H11D3, &H8E, &H3A, &H0, &HC0, &H4F, &H68, &H37, &HD5)
  BHID_SFUIObject = iid
End Function
Public Function BHID_DataObject() As UUID
'{0xB8C0BD9F, 0xED24, 0x455C, 0x83,0xE6, 0xD5,0x39,0x0C,0x4F,0xE8,0xC4}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB8C0BD9F, &HED24, &H455C, &H83, &HE6, &HD5, &H39, &HC, &H4F, &HE8, &HC4)
 BHID_DataObject = iid
End Function
Public Function BHID_SFObject() As UUID
'{0x3981E224, 0xF559, 0x11D3, 0x8E,0x3A, 0x00,0xC0,0x4F,0x68,0x37,0xD5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3981E224, &HF559, &H11D3, &H8E, &H3A, &H0, &HC0, &H4F, &H68, &H37, &HD5)
 BHID_SFObject = iid
End Function
Public Function BHID_SFViewObject() As UUID
'{0x3981E226, 0xF559, 0x11D3, 0x8E,0x3A, 0x00,0xC0,0x4F,0x68,0x37,0xD5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3981E226, &HF559, &H11D3, &H8E, &H3A, &H0, &HC0, &H4F, &H68, &H37, &HD5)
 BHID_SFViewObject = iid
End Function
Public Function BHID_Storage() As UUID
'{0x3981E227, 0xF559, 0x11D3, 0x8E,0x3A, 0x00,0xC0,0x4F,0x68,0x37,0xD5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3981E227, &HF559, &H11D3, &H8E, &H3A, &H0, &HC0, &H4F, &H68, &H37, &HD5)
 BHID_Storage = iid
End Function
Public Function BHID_Stream() As UUID
'{0x1CEBB3AB, 0x7C10, 0x499A, 0xA4,0x17, 0x92,0xCA,0x16,0xC4,0xCB,0x83}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1CEBB3AB, &H7C10, &H499A, &HA4, &H17, &H92, &HCA, &H16, &HC4, &HCB, &H83)
 BHID_Stream = iid
End Function
Public Function BHID_StorageEnum() As UUID
'{0x4621A4E3, 0xF0D6, 0x4773, 0x8A,0x9C, 0x46,0xE7,0x7B,0x17,0x48,0x40}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4621A4E3, &HF0D6, &H4773, &H8A, &H9C, &H46, &HE7, &H7B, &H17, &H48, &H40)
 BHID_StorageEnum = iid
End Function
Public Function BHID_Transfer() As UUID
'{0xD5E346A1, 0xF753, 0x4932, 0xB4,0x03, 0x45,0x74,0x80,0x0E,0x24,0x98}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD5E346A1, &HF753, &H4932, &HB4, &H3, &H45, &H74, &H80, &HE, &H24, &H98)
 BHID_Transfer = iid
End Function
Public Function BHID_Filter() As UUID
'{0x38D08778, 0xF557, 0x4690, 0x9E,0xBF, 0xBA,0x54,0x70,0x6A,0xD8,0xF7}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H38D08778, &HF557, &H4690, &H9E, &HBF, &HBA, &H54, &H70, &H6A, &HD8, &HF7)
 BHID_Filter = iid
End Function
Public Function BHID_LinkTargetItem() As UUID
'{0x3981E228, 0xF559, 0x11D3, 0x8E,0x3A, 0x00,0xC0,0x4F,0x68,0x37,0xD5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3981E228, &HF559, &H11D3, &H8E, &H3A, &H0, &HC0, &H4F, &H68, &H37, &HD5)
 BHID_LinkTargetItem = iid
End Function
Public Function BHID_PropertyStore() As UUID
'{0x0384E1A4, 0x1523, 0x439C, 0xA4,0xC8, 0xAB,0x91,0x10,0x52,0xF5,0x86}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H384E1A4, &H1523, &H439C, &HA4, &HC8, &HAB, &H91, &H10, &H52, &HF5, &H86)
 BHID_PropertyStore = iid
End Function
Public Function BHID_EnumAssocHandlers() As UUID
'{0xB8AB0B9C, 0xC2EC, 0x4F7A, 0x91,0x8D, 0x31,0x49,0x00,0xE6,0x28,0x0A}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB8AB0B9C, &HC2EC, &H4F7A, &H91, &H8D, &H31, &H49, &H0, &HE6, &H28, &HA)
 BHID_EnumAssocHandlers = iid
End Function
Public Function BHID_ThumbnailHandler() As UUID
'{0x7B2E650A, 0x8E20, 0x4F4A, 0xB0,0x9E, 0x65,0x97,0xAF,0xC7,0x2F,0xB0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7B2E650A, &H8E20, &H4F4A, &HB0, &H9E, &H65, &H97, &HAF, &HC7, &H2F, &HB0)
 BHID_ThumbnailHandler = iid
End Function
Public Function BHID_EnumItems() As UUID
'{0x94F60519, 0x2850, 0x4924, 0xAA,0x5A, 0xD1,0x5E,0x84,0x86,0x80,0x39}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H94F60519, &H2850, &H4924, &HAA, &H5A, &HD1, &H5E, &H84, &H86, &H80, &H39)
 BHID_EnumItems = iid
End Function
Public Function BHID_RandomAccessStream() As UUID
'0xf16fc93b, 0x77ae, 0x4cfe, 0xbd, 0xa7, 0xa8, 0x66, 0xee, 0xa6, 0x87, 0x8d
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF16FC93B, &H77AE, &H4CFE, &HBD, &HA7, &HA8, &H66, &HEE, &HA6, &H87, &H8D)
 BHID_RandomAccessStream = iid
End Function
Public Function BHID_FilePlaceholder() As UUID
'0x8677dceb, 0xaae0, 0x4005, 0x8d, 0x3d, 0x54, 0x7f, 0xa8, 0x52, 0xf8, 0x25)
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8677DCEB, &HAAE0, &H4005, &H8D, &H3D, &H54, &H7F, &HA8, &H52, &HF8, &H25)
 BHID_FilePlaceholder = iid
End Function
Public Function IID_IShellIconOverlay() As UUID
'{7d688a70-c613-11d0-999b-00c04fd655e1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7D688A70, CInt(&HC613), CInt(&H11D0), &H99, &H9B, &H0, &HC0, &H4F, &HD6, &H55, &HE1)
 IID_IShellIconOverlay = iid
End Function
Public Function IID_IShellIconOverlayIdentifier() As UUID
'{0c6c4200-c589-11d0-999a-00c04fd655e1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC6C4200, CInt(&HC589), CInt(&H11D0), &H99, &H9A, &H0, &HC0, &H4F, &HD6, &H55, &HE1)
 IID_IShellIconOverlayIdentifier = iid
End Function
Public Function IID_IShellIconOverlayManager() As UUID
'{f10b5e34-dd3b-42a7-aa7d-2f4ec54bb09b}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF10B5E34, CInt(&HDD3B), CInt(&H42A7), &HAA, &H7D, &H2F, &H4E, &HC5, &H4B, &HB0, &H9B)
 IID_IShellIconOverlayManager = iid
End Function
Public Function IID_IListView() As UUID
'{E5B16AF2-3990-4681-A609-1F060CD14269}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE5B16AF2, CInt(&H3990), CInt(&H4681), &HA6, &H9, &H1F, &H6, &HC, &HD1, &H42, &H69)
 IID_IListView = iid
End Function
Public Function IID_IListViewVista() As UUID
'{2FFE2979-5928-4386-9CDB-8E1F15B72FB4}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2FFE2979, CInt(&H5928), CInt(&H4386), &H9C, &HDB, &H8E, &H1F, &H15, &HB7, &H2F, &HB4)
 IID_IListViewVista = iid
End Function
Public Function IID_IListViewFooter() As UUID
'{F0034DA8-8A22-4151-8F16-2EBA76565BCC}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF0034DA8, CInt(&H8A22), CInt(&H4151), &H8F, &H16, &H2E, &HBA, &H76, &H56, &H5B, &HCC)
 IID_IListViewFooter = iid
End Function
Public Function IID_IListViewFooterCallback() As UUID
'{88EB9442-913B-4AB4-A741-DD99DCB7558B}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H88EB9442, CInt(&H913B), CInt(&H4AB4), &HA7, &H41, &HDD, &H99, &HDC, &HB7, &H55, &H8B)
 IID_IListViewFooterCallback = iid
End Function
Public Function IID_IOwnerDataCallback() As UUID
'{44C09D56-8D3B-419D-A462-7B956B105B47}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H44C09D56, CInt(&H8D3B), CInt(&H419D), &HA4, &H62, &H7B, &H95, &H6B, &H10, &H5B, &H47)
 IID_IOwnerDataCallback = iid
End Function
Public Function IID_IPropertyControlBase() As UUID
'{6E71A510-732A-4557-9596-A827E36DAF8F}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6E71A510, CInt(&H732A), CInt(&H4557), &H95, &H96, &HA8, &H27, &HE3, &H6D, &HAF, &H8F)
 IID_IPropertyControlBase = iid
End Function
Public Function IID_IPropertyControl() As UUID
'{5E82A4DD-9561-476A-8634-1BEBACBA4A38}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5E82A4DD, CInt(&H9561), CInt(&H476A), &H86, &H34, &H1B, &HEB, &HAC, &HBA, &H4A, &H38)
 IID_IPropertyControl = iid
End Function
Public Function IID_IDrawPropertyControl() As UUID
'{E6DFF6FD-BCD5-4162-9C65-A3B18C616FDB}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE6DFF6FD, CInt(&HBCD5), CInt(&H4162), &H9C, &H65, &HA3, &HB1, &H8C, &H61, &H6F, &HDB)
 IID_IDrawPropertyControl = iid
End Function
Public Function IID_IPropertyValue() As UUID
'{7AF7F355-1066-4E17-B1F2-19FE2F099CD2}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7AF7F355, CInt(&H1066), CInt(&H4E17), &HB1, &HF2, &H19, &HFE, &H2F, &H9, &H9C, &HD2)
 IID_IPropertyValue = iid
End Function
Public Function IID_ISubItemCallback() As UUID
'{11A66240-5489-42C2-AEBF-286FC831524C}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H11A66240, CInt(&H5489), CInt(&H42C2), &HAE, &HBF, &H28, &H6F, &HC8, &H31, &H52, &H4C)
 IID_ISubItemCallback = iid
End Function

Public Function IID_IShellApp() As UUID
'{A3E14960-935F-11D1-B8B8-006008059382}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA3E14960, CInt(&H935F), CInt(&H11D1), &HB8, &HB8, &H0, &H60, &H8, &H5, &H93, &H82)
 IID_IShellApp = iid
End Function
Public Function IID_IAppPublisher() As UUID
'{07250A10-9CF9-11D1-9076-006008059382}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7250A10, CInt(&H9CF9), CInt(&H11D1), &H90, &H76, &H0, &H60, &H8, &H5, &H93, &H82)
 IID_IAppPublisher = iid
End Function
Public Function IID_IBandSite() As UUID
'{4CF504B0-DE96-11D0-8B3F-00A0C911E8E5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4CF504B0, CInt(&HDE96), CInt(&H11D0), &H8B, &H3F, &H0, &HA0, &HC9, &H11, &HE8, &HE5)
 IID_IBandSite = iid
End Function
Public Function IID_INewWindowManager() As UUID
'{4CF504B0-DE96-11D0-8B3F-00A0C911E8E5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4CF504B0, CInt(&HDE96), CInt(&H11D0), &H8B, &H3F, &H0, &HA0, &HC9, &H11, &HE8, &HE5)
 IID_INewWindowManager = iid
End Function
Public Function IID_IDelegateFolder() As UUID
'{ADD8BA80-002B-11D0-8F0F-00C04FD7D062}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HADD8BA80, CInt(&H2B), CInt(&H11D0), &H8F, &HF, &H0, &HC0, &H4F, &HD7, &HD0, &H62)
 IID_IDelegateFolder = iid
End Function
Public Function IID_IBrowserFrameOptions() As UUID
'{10DF43C8-1DBE-11d3-8B34-006097DF5BD4}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H10DF43C8, CInt(&H1DBE), CInt(&H11D3), &H8B, &H34, &H0, &H60, &H97, &HDF, &H5B, &HD4)
 IID_IBrowserFrameOptions = iid
End Function
Public Function IID_IFileIsInUse() As UUID
'{64a1cbf0-3a1a-4461-9158-376969693950}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H64A1CBF0, CInt(&H3A1A), CInt(&H4461), &H91, &H58, &H37, &H69, &H69, &H69, &H39, &H50)
 IID_IFileIsInUse = iid
End Function
Public Function IID_IOpenControlPanel() As UUID
'{D11AD862-66DE-4DF4-BF6C-1F5621996AF1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD11AD862, CInt(&H66DE), CInt(&H4DF4), &HBF, &H6C, &H1F, &H56, &H21, &H99, &H6A, &HF1)
 IID_IOpenControlPanel = iid
End Function
Public Function IID_IDesktopWallpaper() As UUID
'{B92B56A9-8B55-4E14-9A89-0199BBB6F93B}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB92B56A9, CInt(&H8B55), CInt(&H4E14), &H9A, &H89, &H1, &H99, &HBB, &HB6, &HF9, &H3B)
 IID_IDesktopWallpaper = iid
End Function
Public Function IID_IContactManagerInterop() As UUID
'{99eacba7-e073-43b6-a896-55afe48a0833}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H99EACBA7, CInt(&HE073), CInt(&H43B6), &HA8, &H96, &H55, &HAF, &HE4, &H8A, &H8, &H33)
 IID_IContactManagerInterop = iid
End Function
Public Function IID_IAppActivationUIInfo() As UUID
'{abad189d-9fa3-4278-b3ca-8ca448a88dcb}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HABAD189D, CInt(&H9FA3), CInt(&H4278), &HB3, &HCA, &H8C, &HA4, &H48, &HA8, &H8D, &HCB)
 IID_IAppActivationUIInfo = iid
End Function
Public Function IID_IHandlerActivationHost() As UUID
'{35094a87-8bb1-4237-96c6-c417eebdb078}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H35094A87, CInt(&H8BB1), CInt(&H4237), &H96, &HC6, &HC4, &H17, &HEE, &HBD, &HB0, &H78)
 IID_IHandlerActivationHost = iid
End Function
Public Function IID_IHandlerInfo() As UUID
'{997706ef-f880-453b-8118-39e1a2d2655a}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H997706EF, CInt(&HF880), CInt(&H453B), &H81, &H18, &H39, &HE1, &HA2, &HD2, &H65, &H5A)
 IID_IHandlerInfo = iid
End Function
Public Function IID_ILaunchSourceAppUserModelId() As UUID
'{989191AC-28FF-4CF0-9584-E0D078BC2396}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H989191AC, CInt(&H28FF), CInt(&H4CF0), &H95, &H84, &HE0, &HD0, &H78, &HBC, &H23, &H96)
 IID_ILaunchSourceAppUserModelId = iid
End Function
Public Function IID_ILaunchTargetViewSizePreference() As UUID
'{2F0666C6-12F7-4360-B511-A394A0553725}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2F0666C6, CInt(&H12F7), CInt(&H4360), &HB5, &H11, &HA3, &H94, &HA0, &H55, &H37, &H25)
 IID_ILaunchTargetViewSizePreference = iid
End Function
Public Function IID_ILaunchSourceViewSizePreference() As UUID
'{E5AA01F7-1FB8-4830-8720-4E6734CBD5F3}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE5AA01F7, CInt(&H1FB8), CInt(&H4830), &H87, &H20, &H4E, &H67, &H34, &HCB, &HD5, &HF3)
 IID_ILaunchSourceViewSizePreference = iid
End Function
Public Function IID_ILaunchTargetMonitor() As UUID
'{266FBC7E-490D-46ED-A96B-2274DB252003}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H266FBC7E, CInt(&H490D), CInt(&H46ED), &HA9, &H6B, &H22, &H74, &HDB, &H25, &H20, &H3)
 IID_ILaunchTargetMonitor = iid
End Function
Public Function IID_IApplicationDesignModeSettings2() As UUID
'{490514E1-675A-4D6E-A58D-E54901B4CA2F}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H490514E1, CInt(&H675A), CInt(&H4D6E), &HA5, &H8D, &HE5, &H49, &H1, &HB4, &HCA, &H2F)
 IID_IApplicationDesignModeSettings2 = iid
End Function
Public Function IID_IApplicationDesignModeSettings() As UUID
'{2A3DEE9A-E31D-46D6-8508-BCC597DB3557}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2A3DEE9A, CInt(&HE31D), CInt(&H46D6), &H85, &H8, &HBC, &HC5, &H97, &HDB, &H35, &H57)
 IID_IApplicationDesignModeSettings = iid
End Function
Public Function IID_IExecuteCommandApplicationHostEnvironment() As UUID
'{18B21AA9-E184-4FF0-9F5E-F882D03771B3}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H18B21AA9, CInt(&HE184), CInt(&H4FF0), &H9F, &H5E, &HF8, &H82, &HD0, &H37, &H71, &HB3)
 IID_IExecuteCommandApplicationHostEnvironment = iid
End Function
Public Function IID_ISuspensionDependencyManager() As UUID
'{52B83A42-2543-416A-81D9-C0DE7969C8B3}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H52B83A42, CInt(&H2543), CInt(&H416A), &H81, &HD9, &HC0, &HDE, &H79, &H69, &HC8, &HB3)
 IID_ISuspensionDependencyManager = iid
End Function
Public Function IID_IPackageDebugSettings2() As UUID
'{6E3194BB-AB82-4D22-93F5-FABDA40E7B16}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6E3194BB, CInt(&HAB82), CInt(&H4D22), &H93, &HF5, &HFA, &HBD, &HA4, &HE, &H7B, &H16)
 IID_IPackageDebugSettings2 = iid
End Function
Public Function IID_IPackageDebugSettings() As UUID
'{F27C3930-8029-4AD1-94E3-3DBA417810C1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF27C3930, CInt(&H8029), CInt(&H4AD1), &H94, &HE3, &H3D, &HBA, &H41, &H78, &H10, &HC1)
 IID_IPackageDebugSettings = iid
End Function
Public Function IID_IPackageExecutionStateChangeNotification() As UUID
'{1BB12A62-2AD8-432B-8CCF-0C2C52AFCD5B}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1BB12A62, CInt(&H2AD8), CInt(&H432B), &H8C, &HCF, &HC, &H2C, &H52, &HAF, &HCD, &H5B)
 IID_IPackageExecutionStateChangeNotification = iid
End Function
Public Function IID_IDataTransferManagerInterop() As UUID
'{3A3DCD6C-3EAB-43DC-BCDE-45671CE800C8}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3A3DCD6C, CInt(&H3EAB), CInt(&H43DC), &HBC, &HDE, &H45, &H67, &H1C, &HE8, &H0, &HC8)
 IID_IDataTransferManagerInterop = iid
End Function
Public Function IID_IDataObjectProvider() As UUID
'{3D25F6D6-4B2A-433c-9184-7C33AD35D001}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3D25F6D6, CInt(&H4B2A), CInt(&H433C), &H91, &H84, &H7C, &H33, &HAD, &H35, &HD0, &H1)
 IID_IDataObjectProvider = iid
End Function
Public Function IID_IUpdateIDList() As UUID
'{6589b6d2-5f8d-4b9e-b7e0-23cdd9717d8c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6589B6D2, CInt(&H5F8D), CInt(&H4B9E), &HB7, &HE0, &H23, &HCD, &HD9, &H71, &H7D, &H8C)
IID_IUpdateIDList = iid
End Function
Public Function IID_IObjectWithAppUserModelID() As UUID
'{36db0196-9665-46d1-9ba7-d3709eecf9ed}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H36DB0196, CInt(&H9665), CInt(&H46D1), &H9B, &HA7, &HD3, &H70, &H9E, &HEC, &HF9, &HED)
IID_IObjectWithAppUserModelID = iid
End Function
Public Function IID_IObjectWithProgID() As UUID
'{71e806fb-8dee-46fc-bf8c-7748a8a1ae13}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H71E806FB, CInt(&H8DEE), CInt(&H46FC), &HBF, &H8C, &H77, &H48, &HA8, &HA1, &HAE, &H13)
IID_IObjectWithProgID = iid
End Function
Public Function IID_IObjectWithCancelEvent() As UUID
'{F279B885-0AE9-4b85-AC06-DDECF9408941}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF279B885, CInt(&HAE9), CInt(&H4B85), &HAC, &H6, &HDD, &HEC, &HF9, &H40, &H89, &H41)
IID_IObjectWithCancelEvent = iid
End Function
Public Function IID_IObjectWithSelection() As UUID
'{1c9cd5bb-98e9-4491-a60f-31aacc72b83c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1C9CD5BB, CInt(&H98E9), CInt(&H4491), &HA6, &HF, &H31, &HAA, &HCC, &H72, &HB8, &H3C)
IID_IObjectWithSelection = iid
End Function
Public Function IID_IObjectWithBackReferences() As UUID
'{321a6a6a-d61f-4bf3-97ae-14be2986bb36}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H321A6A6A, CInt(&HD61F), CInt(&H4BF3), &H97, &HAE, &H14, &HBE, &H29, &H86, &HBB, &H36)
IID_IObjectWithBackReferences = iid
End Function
Public Function IID_IRemoteComputer() As UUID
'{000214FE-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214FE, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IRemoteComputer = iid
End Function
Public Function IID_IAccessibilityDockingServiceCallback() As UUID
'{157733FD-A592-42E5-B594-248468C5A81B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H157733FD, CInt(&HA592), CInt(&H42E5), &HB5, &H94, &H24, &H84, &H68, &HC5, &HA8, &H1B)
IID_IAccessibilityDockingServiceCallback = iid
End Function
Public Function IID_IAccessibilityDockingService() As UUID
'{8849DC22-CEDF-4C95-998D-051419DD3F76}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8849DC22, CInt(&HCEDF), CInt(&H4C95), &H99, &H8D, &H5, &H14, &H19, &HDD, &H3F, &H76)
IID_IAccessibilityDockingService = iid
End Function
Public Function IID_IHostDialogHelper() As UUID
'{53DEC138-A51E-11d2-861E-00C04FA35C89}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53DEC138, CInt(&HA51E), CInt(&H11D2), &H86, &H1E, &H0, &HC0, &H4F, &HA3, &H5C, &H89)
 IID_IHostDialogHelper = iid
End Function
Public Function IID_IFileSearchBand() As UUID
'{2D91EEA1-9932-11d2-BE86-00A0C9A83DA1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2D91EEA1, CInt(&H9932), CInt(&H11D2), &HBE, &H86, &H0, &HA0, &HC9, &HA8, &H3D, &HA1)
 IID_IFileSearchBand = iid
End Function
Public Function IID_IDispatchEx() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA6EF9860, &HC720, &H11D0, &H93, &H37, &H0, &HA0, &HC9, &HD, &HCA, &HA9)
IID_IDispatchEx = iid
End Function
Public Function IID_IDispError() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA6EF9861, &HC720, &H11D0, &H93, &H37, &H0, &HA0, &HC9, &HD, &HCA, &HA9)
IID_IDispError = iid
End Function
Public Function IID_IVariantChangeType() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA6EF9862, &HC720, &H11D0, &H93, &H37, &H0, &HA0, &HC9, &HD, &HCA, &HA9)
IID_IVariantChangeType = iid
End Function
Public Function IID_IProvideRuntimeContext() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H10E2414A, &HEC59, &H49D2, &HBC, &H51, &H5A, &HDD, &H2C, &H36, &HFE, &HBC)
IID_IProvideRuntimeContext = iid
End Function
Public Function IID_IObjectIdentity() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCA04B7E6, &HD21, &H11D1, &H8C, &HC5, &H0, &HC0, &H4F, &HC2, &HB0, &H85)
IID_IObjectIdentity = iid
End Function
Public Function IID_ICanHandleException() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC5598E60, &HB307, &H11D1, &HB2, &H7D, &H0, &H60, &H8, &HC3, &HFB, &HFB)
IID_ICanHandleException = iid
End Function
Public Function IID_IUIEventLogger() As UUID
'{ec3e1034-dbf4-41a1-95d5-03e0f1026e05}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEC3E1034, CInt(&HDBF4), CInt(&H41A1), &H95, &HD5, &H3, &HE0, &HF1, &H2, &H6E, &H5)
IID_IUIEventLogger = iid
End Function
Public Function IID_IUIEventingManager() As UUID
'{3BE6EA7F-9A9B-4198-9368-9B0F923BD534}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3BE6EA7F, CInt(&H9A9B), CInt(&H4198), &H93, &H68, &H9B, &HF, &H92, &H3B, &HD5, &H34)
IID_IUIEventingManager = iid
End Function
Public Function IID_IUISimplePropertySet() As UUID
'{c205bb48-5b1c-4219-a106-15bd0a5f24e2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC205BB48, CInt(&H5B1C), CInt(&H4219), &HA1, &H6, &H15, &HBD, &HA, &H5F, &H24, &HE2)
IID_IUISimplePropertySet = iid
End Function
Public Function IID_IUIRibbon() As UUID
'{803982ab-370a-4f7e-a9e7-8784036a6e26}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H803982AB, CInt(&H370A), CInt(&H4F7E), &HA9, &HE7, &H87, &H84, &H3, &H6A, &H6E, &H26)
IID_IUIRibbon = iid
End Function
Public Function IID_IUIFramework() As UUID
'{F4F0385D-6872-43a8-AD09-4C339CB3F5C5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF4F0385D, CInt(&H6872), CInt(&H43A8), &HAD, &H9, &H4C, &H33, &H9C, &HB3, &HF5, &HC5)
IID_IUIFramework = iid
End Function
Public Function IID_IUIContextualUI() As UUID
'{EEA11F37-7C46-437c-8E55-B52122B29293}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEEA11F37, CInt(&H7C46), CInt(&H437C), &H8E, &H55, &HB5, &H21, &H22, &HB2, &H92, &H93)
IID_IUIContextualUI = iid
End Function
Public Function IID_IUICollection() As UUID
'{DF4F45BF-6F9D-4dd7-9D68-D8F9CD18C4DB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDF4F45BF, CInt(&H6F9D), CInt(&H4DD7), &H9D, &H68, &HD8, &HF9, &HCD, &H18, &HC4, &HDB)
IID_IUICollection = iid
End Function
Public Function IID_IUICollectionChangedEvent() As UUID
'{6502AE91-A14D-44b5-BBD0-62AACC581D52}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6502AE91, CInt(&HA14D), CInt(&H44B5), &HBB, &HD0, &H62, &HAA, &HCC, &H58, &H1D, &H52)
IID_IUICollectionChangedEvent = iid
End Function
Public Function IID_IUICommandHandler() As UUID
'{75ae0a2d-dc03-4c9f-8883-069660d0beb6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H75AE0A2D, CInt(&HDC03), CInt(&H4C9F), &H88, &H83, &H6, &H96, &H60, &HD0, &HBE, &HB6)
IID_IUICommandHandler = iid
End Function
Public Function IID_IUIApplication() As UUID
'{D428903C-729A-491d-910D-682A08FF2522}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD428903C, CInt(&H729A), CInt(&H491D), &H91, &HD, &H68, &H2A, &H8, &HFF, &H25, &H22)
IID_IUIApplication = iid
End Function
Public Function IID_IUIImage() As UUID
'{23c8c838-4de6-436b-ab01-5554bb7c30dd}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H23C8C838, CInt(&H4DE6), CInt(&H436B), &HAB, &H1, &H55, &H54, &HBB, &H7C, &H30, &HDD)
IID_IUIImage = iid
End Function
Public Function IID_IUIImageFromBitmap() As UUID
'{18aba7f3-4c1c-4ba2-bf6c-f5c3326fa816}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H18ABA7F3, CInt(&H4C1C), CInt(&H4BA2), &HBF, &H6C, &HF5, &HC3, &H32, &H6F, &HA8, &H16)
IID_IUIImageFromBitmap = iid
End Function



Public Function CATID_ActiveScript() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF0B7A1A1, &H9847, &H11CF, &H8F, &H20, &H0, &H80, &H5F, &H2C, &HD0, &H64)
CATID_ActiveScript = iid
End Function
Public Function CATID_ActiveScriptParse() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF0B7A1A2, &H9847, &H11CF, &H8F, &H20, &H0, &H80, &H5F, &H2C, &HD0, &H64)
CATID_ActiveScriptParse = iid
End Function
Public Function CATID_ActiveScriptEncode() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF0B7A1A3, &H9847, &H11CF, &H8F, &H20, &H0, &H80, &H5F, &H2C, &HD0, &H64)
CATID_ActiveScriptEncode = iid
End Function
Public Function IID_IActiveScript() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBB1A2AE1, &HA4F9, &H11CF, &H8F, &H20, &H0, &H80, &H5F, &H2C, &HD0, &H64)
IID_IActiveScript = iid
End Function
Public Function IID_IActiveScriptEncode() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBB1A2AE3, &HA4F9, &H11CF, &H8F, &H20, &H0, &H80, &H5F, &H2C, &HD0, &H64)
IID_IActiveScriptEncode = iid
End Function
Public Function IID_IActiveScriptHostEncode() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBEE9B76E, &HCFE3, &H11D1, &HB7, &H47, &H0, &HC0, &H4F, &HC2, &HB0, &H85)
IID_IActiveScriptHostEncode = iid
End Function
#If Win64 Then
Public Function IID_IActiveScriptParse() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC7EF7658, &HE1EE, &H480E, &H97, &HEA, &HD5, &H2C, &HB4, &HD7, &H6D, &H17)
IID_IActiveScriptParse = iid
End Function
Public Function IID_IActiveScriptParseProcedureOld() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H21F57128, &H8C9, &H4638, &HBA, &H12, &H22, &HD1, &H5D, &H88, &HDC, &H5C)
IID_IActiveScriptParseProcedureOld = iid
End Function
Public Function IID_IActiveScriptParseProcedure() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC64713B6, &HE029, &H4CC5, &H92, &H0, &H43, &H8B, &H72, &H89, &HB, &H6A)
IID_IActiveScriptParseProcedure = iid
End Function
Public Function IID_IActiveScriptParseProcedure2() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFE7C4271, &H210C, &H448D, &H9F, &H54, &H76, &HDA, &HB7, &H4, &H7B, &H28)
IID_IActiveScriptParseProcedure2 = iid
End Function
#Else
Public Function IID_IActiveScriptParse() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBB1A2AE2, &HA4F9, &H11CF, &H8F, &H20, &H0, &H80, &H5F, &H2C, &HD0, &H64)
IID_IActiveScriptParse = iid
End Function
Public Function IID_IActiveScriptParseProcedureOld() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1CFF0050, &H6FDD, &H11D0, &H93, &H28, &H0, &HA0, &HC9, &HD, &HCA, &HA9)
IID_IActiveScriptParseProcedureOld = iid
End Function
Public Function IID_IActiveScriptParseProcedure() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAA5B6A80, &HB834, &H11D0, &H93, &H2F, &H0, &HA0, &HC9, &HD, &HCA, &HA9)
IID_IActiveScriptParseProcedure = iid
End Function
Public Function IID_IActiveScriptParseProcedure2() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H71EE5B20, &HFB04, &H11D1, &HB3, &HA8, &H0, &HA0, &HC9, &H11, &HE8, &HB2)
IID_IActiveScriptParseProcedure2 = iid
End Function
#End If
Public Function IID_IActiveScriptSite() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDB01A1E3, &HA42B, &H11CF, &H8F, &H20, &H0, &H80, &H5F, &H2C, &HD0, &H64)
IID_IActiveScriptSite = iid
End Function
Public Function IID_IActiveScriptSiteTraceInfo() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4B7272AE, &H1955, &H4BFE, &H98, &HB0, &H78, &H6, &H21, &H88, &H85, &H69)
IID_IActiveScriptSiteTraceInfo = iid
End Function
Public Function IID_IActiveScriptSiteWindow() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD10F6761, &H83E9, &H11CF, &H8F, &H20, &H0, &H80, &H5F, &H2C, &HD0, &H64)
IID_IActiveScriptSiteWindow = iid
End Function
Public Function IID_IActiveScriptSiteInterruptPoll() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H539698A0, &HCDCA, &H11CF, &HA5, &HEB, &H0, &HAA, &H0, &H47, &HA0, &H63)
IID_IActiveScriptSiteInterruptPoll = iid
End Function
Public Function IID_IActiveScriptSiteUIControl() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAEDAE97E, &HD7EE, &H4796, &HB9, &H60, &H7F, &H9, &H2A, &HE8, &H44, &HAB)
IID_IActiveScriptSiteUIControl = iid
End Function
Public Function IID_IActiveScriptError() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEAE1BA61, &HA4ED, &H11CF, &H8F, &H20, &H0, &H80, &H5F, &H2C, &HD0, &H64)
IID_IActiveScriptError = iid
End Function
Public Function IID_IActiveScriptError64() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB21FB2A1, &H5B8F, &H4963, &H8C, &H21, &H21, &H45, &HF, &H84, &HED, &H7F)
IID_IActiveScriptError64 = iid
End Function
Public Function IID_IBindEventHandler() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H63CDBCB0, &HC1B1, &H11D0, &H93, &H36, &H0, &HA0, &HC9, &HD, &HCA, &HA9)
IID_IBindEventHandler = iid
End Function
Public Function IID_IActiveScriptStats() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB8DA6310, &HE19B, &H11D0, &H93, &H3C, &H0, &HA0, &HC9, &HD, &HCA, &HA9)
IID_IActiveScriptStats = iid
End Function
Public Function IID_IActiveScriptProperty() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4954E0D0, &HFBC7, &H11D1, &H84, &H10, &H0, &H60, &H8, &HC3, &HFB, &HFC)
IID_IActiveScriptProperty = iid
End Function
Public Function IID_ITridentEventSink() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1DC9CA50, &H6EF, &H11D2, &H84, &H15, &H0, &H60, &H8, &HC3, &HFB, &HFC)
IID_ITridentEventSink = iid
End Function
Public Function IID_IActiveScriptGarbageCollector() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6AA2C4A0, &H2B53, &H11D4, &HA2, &HA0, &H0, &H10, &H4B, &HD3, &H50, &H90)
IID_IActiveScriptGarbageCollector = iid
End Function
Public Function IID_IActiveScriptSIPInfo() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H764651D0, &H38DE, &H11D4, &HA2, &HA3, &H0, &H10, &H4B, &HD3, &H50, &H90)
IID_IActiveScriptSIPInfo = iid
End Function
Public Function IID_IActiveScriptTraceInfo() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC35456E7, &HBEBF, &H4A1B, &H86, &HA9, &H24, &HD5, &H6B, &HE8, &HB3, &H69)
IID_IActiveScriptTraceInfo = iid
End Function
Public Function OID_VBSSIP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1629F04E, &H2799, &H4DB5, &H8F, &HE5, &HAC, &HE1, &HF, &H17, &HEB, &HAB)
OID_VBSSIP = iid
End Function
Public Function OID_JSSIP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C9E010, &H38CE, &H11D4, &HA2, &HA3, &H0, &H10, &H4B, &HD3, &H50, &H90)
OID_JSSIP = iid
End Function
Public Function OID_WSFSIP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1A610570, &H38CE, &H11D4, &HA2, &HA3, &H0, &H10, &H4B, &HD3, &H50, &H90)
OID_WSFSIP = iid
End Function
Public Function IID_IActiveScriptStringCompare() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H58562769, &HED52, &H42F7, &H84, &H3, &H49, &H63, &H51, &H4E, &H1F, &H11)
IID_IActiveScriptStringCompare = iid
End Function
Public Function IID_ISimpleFrameSite() As UUID
'{742B0E01-14E6-101B-914E-00AA00300CAB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H742B0E01, CInt(&H14E6), CInt(&H101B), &H91, &H4E, &H0, &HAA, &H0, &H30, &HC, &HAB)
IID_ISimpleFrameSite = iid
End Function
Public Function IID_IPropertyPage() As UUID
'{B196B28D-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B28D, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IPropertyPage = iid
End Function
Public Function IID_IPropertyPage2() As UUID
'{01E44665-24AC-101B-84ED-08002B2EC713}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1E44665, CInt(&H24AC), CInt(&H101B), &H84, &HED, &H8, &H0, &H2B, &H2E, &HC7, &H13)
IID_IPropertyPage2 = iid
End Function
Public Function IID_IOpcUri() As UUID
'{bc9c1b9b-d62c-49eb-aef0-3b4e0b28ebed}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBC9C1B9B, CInt(&HD62C), CInt(&H49EB), &HAE, &HF0, &H3B, &H4E, &HB, &H28, &HEB, &HED)
IID_IOpcUri = iid
End Function
Public Function IID_IOpcPartUri() As UUID
'{7d3babe7-88b2-46ba-85cb-4203cb016c87}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7D3BABE7, CInt(&H88B2), CInt(&H46BA), &H85, &HCB, &H42, &H3, &HCB, &H1, &H6C, &H87)
IID_IOpcPartUri = iid
End Function
Public Function IID_IOpcPackage() As UUID
'{42195949-3B79-4fc8-89C6-FC7FB979EE70}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H42195949, CInt(&H3B79), CInt(&H4FC8), &H89, &HC6, &HFC, &H7F, &HB9, &H79, &HEE, &H70)
IID_IOpcPackage = iid
End Function
Public Function IID_IOpcPart() As UUID
'{42195949-3B79-4fc8-89C6-FC7FB979EE71}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H42195949, CInt(&H3B79), CInt(&H4FC8), &H89, &HC6, &HFC, &H7F, &HB9, &H79, &HEE, &H71)
IID_IOpcPart = iid
End Function
Public Function IID_IOpcRelationship() As UUID
'{42195949-3B79-4fc8-89C6-FC7FB979EE72}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H42195949, CInt(&H3B79), CInt(&H4FC8), &H89, &HC6, &HFC, &H7F, &HB9, &H79, &HEE, &H72)
IID_IOpcRelationship = iid
End Function
Public Function IID_IOpcPartSet() As UUID
'{42195949-3B79-4fc8-89C6-FC7FB979EE73}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H42195949, CInt(&H3B79), CInt(&H4FC8), &H89, &HC6, &HFC, &H7F, &HB9, &H79, &HEE, &H73)
IID_IOpcPartSet = iid
End Function
Public Function IID_IOpcRelationshipSet() As UUID
'{42195949-3B79-4fc8-89C6-FC7FB979EE74}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H42195949, CInt(&H3B79), CInt(&H4FC8), &H89, &HC6, &HFC, &H7F, &HB9, &H79, &HEE, &H74)
IID_IOpcRelationshipSet = iid
End Function
Public Function IID_IOpcPartEnumerator() As UUID
'{42195949-3B79-4fc8-89C6-FC7FB979EE75}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H42195949, CInt(&H3B79), CInt(&H4FC8), &H89, &HC6, &HFC, &H7F, &HB9, &H79, &HEE, &H75)
IID_IOpcPartEnumerator = iid
End Function
Public Function IID_IOpcRelationshipEnumerator() As UUID
'{42195949-3B79-4fc8-89C6-FC7FB979EE76}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H42195949, CInt(&H3B79), CInt(&H4FC8), &H89, &HC6, &HFC, &H7F, &HB9, &H79, &HEE, &H76)
IID_IOpcRelationshipEnumerator = iid
End Function
Public Function IID_IOpcSignaturePartReference() As UUID
'{e24231ca-59f4-484e-b64b-36eeda36072c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE24231CA, CInt(&H59F4), CInt(&H484E), &HB6, &H4B, &H36, &HEE, &HDA, &H36, &H7, &H2C)
IID_IOpcSignaturePartReference = iid
End Function
Public Function IID_IOpcRelationshipSelector() As UUID
'{f8f26c7f-b28f-4899-84c8-5d5639ede75f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF8F26C7F, CInt(&HB28F), CInt(&H4899), &H84, &HC8, &H5D, &H56, &H39, &HED, &HE7, &H5F)
IID_IOpcRelationshipSelector = iid
End Function
Public Function IID_IOpcSignatureRelationshipReference() As UUID
'{57babac6-9d4a-4e50-8b86-e5d4051eae7c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H57BABAC6, CInt(&H9D4A), CInt(&H4E50), &H8B, &H86, &HE5, &HD4, &H5, &H1E, &HAE, &H7C)
IID_IOpcSignatureRelationshipReference = iid
End Function
Public Function IID_IOpcSignatureReference() As UUID
'{1b47005e-3011-4edc-be6f-0f65e5ab0342}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1B47005E, CInt(&H3011), CInt(&H4EDC), &HBE, &H6F, &HF, &H65, &HE5, &HAB, &H3, &H42)
IID_IOpcSignatureReference = iid
End Function
Public Function IID_IOpcSignatureCustomObject() As UUID
'{5d77a19e-62c1-44e7-becd-45da5ae51a56}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5D77A19E, CInt(&H62C1), CInt(&H44E7), &HBE, &HCD, &H45, &HDA, &H5A, &HE5, &H1A, &H56)
IID_IOpcSignatureCustomObject = iid
End Function
Public Function IID_IOpcSignaturePartReferenceEnumerator() As UUID
'{80eb1561-8c77-49cf-8266-459b356ee99a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H80EB1561, CInt(&H8C77), CInt(&H49CF), &H82, &H66, &H45, &H9B, &H35, &H6E, &HE9, &H9A)
IID_IOpcSignaturePartReferenceEnumerator = iid
End Function
Public Function IID_IOpcRelationshipSelectorEnumerator() As UUID
'{5e50a181-a91b-48ac-88d2-bca3d8f8c0b1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5E50A181, CInt(&HA91B), CInt(&H48AC), &H88, &HD2, &HBC, &HA3, &HD8, &HF8, &HC0, &HB1)
IID_IOpcRelationshipSelectorEnumerator = iid
End Function
Public Function IID_IOpcSignatureRelationshipReferenceEnumerator() As UUID
'{773ba3e4-f021-48e4-aa04-9816db5d3495}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H773BA3E4, CInt(&HF021), CInt(&H48E4), &HAA, &H4, &H98, &H16, &HDB, &H5D, &H34, &H95)
IID_IOpcSignatureRelationshipReferenceEnumerator = iid
End Function
Public Function IID_IOpcSignatureReferenceEnumerator() As UUID
'{cfa59a45-28b1-4868-969e-fa8097fdc12a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCFA59A45, CInt(&H28B1), CInt(&H4868), &H96, &H9E, &HFA, &H80, &H97, &HFD, &HC1, &H2A)
IID_IOpcSignatureReferenceEnumerator = iid
End Function
Public Function IID_IOpcSignatureCustomObjectEnumerator() As UUID
'{5ee4fe1d-e1b0-4683-8079-7ea0fcf80b4c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5EE4FE1D, CInt(&HE1B0), CInt(&H4683), &H80, &H79, &H7E, &HA0, &HFC, &HF8, &HB, &H4C)
IID_IOpcSignatureCustomObjectEnumerator = iid
End Function
Public Function IID_IOpcCertificateEnumerator() As UUID
'{85131937-8f24-421f-b439-59ab24d140b8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H85131937, CInt(&H8F24), CInt(&H421F), &HB4, &H39, &H59, &HAB, &H24, &HD1, &H40, &HB8)
IID_IOpcCertificateEnumerator = iid
End Function
Public Function IID_IOpcDigitalSignatureEnumerator() As UUID
'{967b6882-0ba3-4358-b9e7-b64c75063c5e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H967B6882, CInt(&HBA3), CInt(&H4358), &HB9, &HE7, &HB6, &H4C, &H75, &H6, &H3C, &H5E)
IID_IOpcDigitalSignatureEnumerator = iid
End Function
Public Function IID_IOpcSignaturePartReferenceSet() As UUID
'{6c9fe28c-ecd9-4b22-9d36-7fdde670fec0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6C9FE28C, CInt(&HECD9), CInt(&H4B22), &H9D, &H36, &H7F, &HDD, &HE6, &H70, &HFE, &HC0)
IID_IOpcSignaturePartReferenceSet = iid
End Function
Public Function IID_IOpcRelationshipSelectorSet() As UUID
'{6e34c269-a4d3-47c0-b5c4-87ff2b3b6136}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6E34C269, CInt(&HA4D3), CInt(&H47C0), &HB5, &HC4, &H87, &HFF, &H2B, &H3B, &H61, &H36)
IID_IOpcRelationshipSelectorSet = iid
End Function
Public Function IID_IOpcSignatureRelationshipReferenceSet() As UUID
'{9f863ca5-3631-404c-828d-807e0715069b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9F863CA5, CInt(&H3631), CInt(&H404C), &H82, &H8D, &H80, &H7E, &H7, &H15, &H6, &H9B)
IID_IOpcSignatureRelationshipReferenceSet = iid
End Function
Public Function IID_IOpcSignatureReferenceSet() As UUID
'{f3b02d31-ab12-42dd-9e2f-2b16761c3c1e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF3B02D31, CInt(&HAB12), CInt(&H42DD), &H9E, &H2F, &H2B, &H16, &H76, &H1C, &H3C, &H1E)
IID_IOpcSignatureReferenceSet = iid
End Function
Public Function IID_IOpcSignatureCustomObjectSet() As UUID
'{8f792ac5-7947-4e11-bc3d-2659ff046ae1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8F792AC5, CInt(&H7947), CInt(&H4E11), &HBC, &H3D, &H26, &H59, &HFF, &H4, &H6A, &HE1)
IID_IOpcSignatureCustomObjectSet = iid
End Function
Public Function IID_IOpcCertificateSet() As UUID
'{56ea4325-8e2d-4167-b1a4-e486d24c8fa7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56EA4325, CInt(&H8E2D), CInt(&H4167), &HB1, &HA4, &HE4, &H86, &HD2, &H4C, &H8F, &HA7)
IID_IOpcCertificateSet = iid
End Function
Public Function IID_IOpcDigitalSignature() As UUID
'{52ab21dd-1cd0-4949-bc80-0c1232d00cb4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H52AB21DD, CInt(&H1CD0), CInt(&H4949), &HBC, &H80, &HC, &H12, &H32, &HD0, &HC, &HB4)
IID_IOpcDigitalSignature = iid
End Function
Public Function IID_IOpcSigningOptions() As UUID
'{50d2d6a5-7aeb-46c0-b241-43ab0e9b407e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H50D2D6A5, CInt(&H7AEB), CInt(&H46C0), &HB2, &H41, &H43, &HAB, &HE, &H9B, &H40, &H7E)
IID_IOpcSigningOptions = iid
End Function
Public Function IID_IOpcDigitalSignatureManager() As UUID
'{d5e62a0b-696d-462f-94df-72e33cef2659}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD5E62A0B, CInt(&H696D), CInt(&H462F), &H94, &HDF, &H72, &HE3, &H3C, &HEF, &H26, &H59)
IID_IOpcDigitalSignatureManager = iid
End Function
Public Function IID_IOpcFactory() As UUID
'{6d0b4446-cd73-4ab3-94f4-8ccdf6116154}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6D0B4446, CInt(&HCD73), CInt(&H4AB3), &H94, &HF4, &H8C, &HCD, &HF6, &H11, &H61, &H54)
IID_IOpcFactory = iid
End Function
Public Function IID_IOleInPlaceObjectWindowless() As UUID
'{1C2056CC-5EF4-101B-8BC8-00AA003E3B29}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1C2056CC, CInt(&H5EF4), CInt(&H101B), &H8B, &HC8, &H0, &HAA, &H0, &H3E, &H3B, &H29)
IID_IOleInPlaceObjectWindowless = iid
End Function
Public Function IID_IOleInPlaceSiteEx() As UUID
'{9C2CAD80-3424-11CF-B670-00AA004CD6D8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9C2CAD80, CInt(&H3424), CInt(&H11CF), &HB6, &H70, &H0, &HAA, &H0, &H4C, &HD6, &HD8)
IID_IOleInPlaceSiteEx = iid
End Function
Public Function IID_IOleInPlaceSiteWindowless() As UUID
'{922EADA0-3424-11CF-B670-00AA004CD6D8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H922EADA0, CInt(&H3424), CInt(&H11CF), &HB6, &H70, &H0, &HAA, &H0, &H4C, &HD6, &HD8)
IID_IOleInPlaceSiteWindowless = iid
End Function
Public Function IID_IViewObjectEx() As UUID
'{3AF24292-0C96-11CE-A0CF-00AA00600AB8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3AF24292, CInt(&HC96), CInt(&H11CE), &HA0, &HCF, &H0, &HAA, &H0, &H60, &HA, &HB8)
IID_IViewObjectEx = iid
End Function
Public Function IID_IOleUndoUnit() As UUID
'{894AD3B0-EF97-11CE-9BC9-00AA00608E01}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H894AD3B0, CInt(&HEF97), CInt(&H11CE), &H9B, &HC9, &H0, &HAA, &H0, &H60, &H8E, &H1)
IID_IOleUndoUnit = iid
End Function
Public Function IID_IOleParentUndoUnit() As UUID
'{A1FAF330-EF97-11CE-9BC9-00AA00608E01}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA1FAF330, CInt(&HEF97), CInt(&H11CE), &H9B, &HC9, &H0, &HAA, &H0, &H60, &H8E, &H1)
IID_IOleParentUndoUnit = iid
End Function
Public Function IID_IEnumOleUndoUnits() As UUID
'{B3E7C340-EF97-11CE-9BC9-00AA00608E01}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB3E7C340, CInt(&HEF97), CInt(&H11CE), &H9B, &HC9, &H0, &HAA, &H0, &H60, &H8E, &H1)
IID_IEnumOleUndoUnits = iid
End Function
Public Function IID_IOleUndoManager() As UUID
'{D001F200-EF97-11CE-9BC9-00AA00608E01}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD001F200, CInt(&HEF97), CInt(&H11CE), &H9B, &HC9, &H0, &HAA, &H0, &H60, &H8E, &H1)
IID_IOleUndoManager = iid
End Function
Public Function IID_IPointerInactive() As UUID
'{55980BA0-35AA-11CF-B671-00AA004CD6D8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H55980BA0, CInt(&H35AA), CInt(&H11CF), &HB6, &H71, &H0, &HAA, &H0, &H4C, &HD6, &HD8)
IID_IPointerInactive = iid
End Function
Public Function IID_IAdviseSinkEx() As UUID
'{3AF24290-0C96-11CE-A0CF-00AA00600AB8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3AF24290, CInt(&HC96), CInt(&H11CE), &HA0, &HCF, &H0, &HAA, &H0, &H60, &HA, &HB8)
IID_IAdviseSinkEx = iid
End Function
Public Function IID_IQuickActivate() As UUID
'{CF51ED10-62FE-11CF-BF86-00A0C9034836}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCF51ED10, CInt(&H62FE), CInt(&H11CF), &HBF, &H86, &H0, &HA0, &HC9, &H3, &H48, &H36)
IID_IQuickActivate = iid
End Function
Public Function IID_IDataAdviseHolder() As UUID
'{00000110-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H110, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IDataAdviseHolder = iid
End Function
Public Function IID_IOleAdviseHolder() As UUID
'{00000111-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H111, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleAdviseHolder = iid
End Function
Public Function IID_IDropSourceNotify() As UUID
'{0000012B-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H12B, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IDropSourceNotify = iid
End Function
Public Function IID_IEnterpriseDropTarget() As UUID
'{390E3878-FD55-4E18-819D-4682081C0CFD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H390E3878, CInt(&HFD55), CInt(&H4E18), &H81, &H9D, &H46, &H82, &H8, &H1C, &HC, &HFD)
IID_IEnterpriseDropTarget = iid
End Function
Public Function IID_IContinue() As UUID
'{0000012a-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H12A, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IContinue = iid
End Function
Public Function IID_IDiskQuotaUser() As UUID
'{7988B574-EC89-11cf-9C00-00AA00A14F56}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7988B574, CInt(&HEC89), CInt(&H11CF), &H9C, &H0, &H0, &HAA, &H0, &HA1, &H4F, &H56)
IID_IDiskQuotaUser = iid
End Function
Public Function IID_IEnumDiskQuotaUsers() As UUID
'{7988B577-EC89-11cf-9C00-00AA00A14F56}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7988B577, CInt(&HEC89), CInt(&H11CF), &H9C, &H0, &H0, &HAA, &H0, &HA1, &H4F, &H56)
IID_IEnumDiskQuotaUsers = iid
End Function
Public Function IID_IDiskQuotaUserBatch() As UUID
'{7988B576-EC89-11cf-9C00-00AA00A14F56}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7988B576, CInt(&HEC89), CInt(&H11CF), &H9C, &H0, &H0, &HAA, &H0, &HA1, &H4F, &H56)
IID_IDiskQuotaUserBatch = iid
End Function
Public Function IID_IDiskQuotaControl() As UUID
'{7988B572-EC89-11cf-9C00-00AA00A14F56}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7988B572, CInt(&HEC89), CInt(&H11CF), &H9C, &H0, &H0, &HAA, &H0, &HA1, &H4F, &H56)
IID_IDiskQuotaControl = iid
End Function
Public Function IID_IDiskQuotaEvents() As UUID
'{7988B579-EC89-11cf-9C00-00AA00A14F56}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7988B579, CInt(&HEC89), CInt(&H11CF), &H9C, &H0, &H0, &HAA, &H0, &HA1, &H4F, &H56)
IID_IDiskQuotaEvents = iid
End Function
Public Function IID_IStorageProviderHandler() As UUID
'{301DFBE5-524C-4B0F-8B2D-21C40B3A2988}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H301DFBE5, CInt(&H524C), CInt(&H4B0F), &H8B, &H2D, &H21, &HC4, &HB, &H3A, &H29, &H88)
 IID_IStorageProviderHandler = iid
End Function
Public Function IID_IStorageProviderPropertyHandler() As UUID
'{301DFBE5-524C-4B0F-8B2D-21C40B3A2988}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H301DFBE5, CInt(&H524C), CInt(&H4B0F), &H8B, &H2D, &H21, &HC4, &HB, &H3A, &H29, &H88)
 IID_IStorageProviderPropertyHandler = iid
End Function
Public Function IID_ILocationReport() As UUID
'{C8B7F7EE-75D0-4db9-B62D-7A0F369CA456}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC8B7F7EE, CInt(&H75D0), CInt(&H4DB9), &HB6, &H2D, &H7A, &HF, &H36, &H9C, &HA4, &H56)
IID_ILocationReport = iid
End Function
Public Function IID_ILatLongReport() As UUID
'{7FED806D-0EF8-4f07-80AC-36A0BEAE3134}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7FED806D, CInt(&HEF8), CInt(&H4F07), &H80, &HAC, &H36, &HA0, &HBE, &HAE, &H31, &H34)
IID_ILatLongReport = iid
End Function
Public Function IID_ICivicAddressReport() As UUID
'{C0B19F70-4ADF-445d-87F2-CAD8FD711792}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC0B19F70, CInt(&H4ADF), CInt(&H445D), &H87, &HF2, &HCA, &HD8, &HFD, &H71, &H17, &H92)
IID_ICivicAddressReport = iid
End Function
Public Function IID_ILocation() As UUID
'{AB2ECE69-56D9-4F28-B525-DE1B0EE44237}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB2ECE69, CInt(&H56D9), CInt(&H4F28), &HB5, &H25, &HDE, &H1B, &HE, &HE4, &H42, &H37)
IID_ILocation = iid
End Function
Public Function IID_ILocationPower() As UUID
'{193E7729-AB6B-4b12-8617-7596E1BB191C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H193E7729, CInt(&HAB6B), CInt(&H4B12), &H86, &H17, &H75, &H96, &HE1, &HBB, &H19, &H1C)
IID_ILocationPower = iid
End Function
Public Function IID_IDefaultLocation() As UUID
'{A65AF77E-969A-4a2e-8ACA-33BB7CBB1235}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA65AF77E, CInt(&H969A), CInt(&H4A2E), &H8A, &HCA, &H33, &HBB, &H7C, &HBB, &H12, &H35)
IID_IDefaultLocation = iid
End Function
Public Function IID_ILocationEvents() As UUID
'{CAE02BBF-798B-4508-A207-35A7906DC73D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCAE02BBF, CInt(&H798B), CInt(&H4508), &HA2, &H7, &H35, &HA7, &H90, &H6D, &HC7, &H3D)
IID_ILocationEvents = iid
End Function
Public Function IID_IDispLatLongReport() As UUID
'{8AE32723-389B-4A11-9957-5BDD48FC9617}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8AE32723, CInt(&H389B), CInt(&H4A11), &H99, &H57, &H5B, &HDD, &H48, &HFC, &H96, &H17)
IID_IDispLatLongReport = iid
End Function
Public Function IID_IDispCivicAddressReport() As UUID
'{16FF1A34-9E30-42c3-B44D-E22513B5767A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H16FF1A34, CInt(&H9E30), CInt(&H42C3), &HB4, &H4D, &HE2, &H25, &H13, &HB5, &H76, &H7A)
IID_IDispCivicAddressReport = iid
End Function
Public Function IID_ILocationReportFactory() As UUID
'{2DAEC322-90B2-47e4-BB08-0DA841935A6B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2DAEC322, CInt(&H90B2), CInt(&H47E4), &HBB, &H8, &HD, &HA8, &H41, &H93, &H5A, &H6B)
IID_ILocationReportFactory = iid
End Function
Public Function IID_ILatLongReportFactory() As UUID
'{3F0804CB-B114-447D-83DD-390174EBB082}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3F0804CB, CInt(&HB114), CInt(&H447D), &H83, &HDD, &H39, &H1, &H74, &HEB, &HB0, &H82)
IID_ILatLongReportFactory = iid
End Function
Public Function IID_ICivicAddressReportFactory() As UUID
'{BF773B93-C64F-4bee-BEB2-67C0B8DF66E0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBF773B93, CInt(&HC64F), CInt(&H4BEE), &HBE, &HB2, &H67, &HC0, &HB8, &HDF, &H66, &HE0)
IID_ICivicAddressReportFactory = iid
End Function
Public Function IID_ISensorManager() As UUID
'{BD77DB67-45A8-42DC-8D00-6DCF15F8377A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBD77DB67, CInt(&H45A8), CInt(&H42DC), &H8D, &H0, &H6D, &HCF, &H15, &HF8, &H37, &H7A)
IID_ISensorManager = iid
End Function
Public Function IID_ILocationPermissions() As UUID
'{D5FB0A7F-E74E-44f5-8E02-4806863A274F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD5FB0A7F, CInt(&HE74E), CInt(&H44F5), &H8E, &H2, &H48, &H6, &H86, &H3A, &H27, &H4F)
IID_ILocationPermissions = iid
End Function
Public Function IID_ISensorCollection() As UUID
'{23571E11-E545-4DD8-A337-B89BF44B10DF}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H23571E11, CInt(&HE545), CInt(&H4DD8), &HA3, &H37, &HB8, &H9B, &HF4, &H4B, &H10, &HDF)
IID_ISensorCollection = iid
End Function
Public Function IID_ISensor() As UUID
'{5FA08F80-2657-458E-AF75-46F73FA6AC5C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5FA08F80, CInt(&H2657), CInt(&H458E), &HAF, &H75, &H46, &HF7, &H3F, &HA6, &HAC, &H5C)
IID_ISensor = iid
End Function
Public Function IID_ISensorDataReport() As UUID
'{0AB9DF9B-C4B5-4796-8898-0470706A2E1D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB9DF9B, CInt(&HC4B5), CInt(&H4796), &H88, &H98, &H4, &H70, &H70, &H6A, &H2E, &H1D)
IID_ISensorDataReport = iid
End Function
Public Function IID_ISensorManagerEvents() As UUID
'{9B3B0B86-266A-4AAD-B21F-FDE5501001B7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9B3B0B86, CInt(&H266A), CInt(&H4AAD), &HB2, &H1F, &HFD, &HE5, &H50, &H10, &H1, &HB7)
IID_ISensorManagerEvents = iid
End Function
Public Function IID_ISensorEvents() As UUID
'{5D8DCC91-4641-47E7-B7C3-B74F48A6C391}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5D8DCC91, CInt(&H4641), CInt(&H47E7), &HB7, &HC3, &HB7, &H4F, &H48, &HA6, &HC3, &H91)
IID_ISensorEvents = iid
End Function
Public Function IID_IAssemblyCache() As UUID
'{e707dcde-d1cd-11d2-bab9-00c04f8eceae}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE707DCDE, CInt(&HD1CD), CInt(&H11D2), &HBA, &HB9, &H0, &HC0, &H4F, &H8E, &HCE, &HAE)
 IID_IAssemblyCache = iid
End Function
Public Function IID_IAssemblyCacheItem() As UUID
'{9e3aaeb4-d1cd-11d2-bab9-00c04f8eceae}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9E3AAEB4, CInt(&HD1CD), CInt(&H11D2), &HBA, &HB9, &H0, &HC0, &H4F, &H8E, &HCE, &HAE)
 IID_IAssemblyCacheItem = iid
End Function
Public Function IID_IAssemblyName() As UUID
'{CD193BC0-B4BC-11d2-9833-00C04FC31D2E}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCD193BC0, CInt(&HB4BC), CInt(&H11D2), &H98, &H33, &H0, &HC0, &H4F, &HC3, &H1D, &H2E)
 IID_IAssemblyName = iid
End Function
Public Function IID_ITransaction() As UUID
'{0fb15084-af41-11ce-bd2b-204c4f4f5020}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFB15084, CInt(&HAF41), CInt(&H11CE), &HBD, &H2B, &H20, &H4C, &H4F, &H4F, &H50, &H20)
 IID_ITransaction = iid
End Function
Public Function IID_ITransactionCloner() As UUID
'{02656950-2152-11d0-944C-00A0C905416E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2656950, CInt(&H2152), CInt(&H11D0), &H94, &H4C, &H0, &HA0, &HC9, &H5, &H41, &H6E)
IID_ITransactionCloner = iid
End Function
Public Function IID_ITransaction2() As UUID
'{34021548-0065-11d3-bac1-00c04f797be2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H34021548, CInt(&H65), CInt(&H11D3), &HBA, &HC1, &H0, &HC0, &H4F, &H79, &H7B, &HE2)
IID_ITransaction2 = iid
End Function
Public Function IID_ITransactionDispenser() As UUID
'{3A6AD9E1-23B9-11cf-AD60-00AA00A74CCD}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3A6AD9E1, CInt(&H23B9), CInt(&H11CF), &HAD, &H60, &H0, &HAA, &H0, &HA7, &H4C, &HCD)
 IID_ITransactionDispenser = iid
End Function
Public Function IID_ITransactionOptions() As UUID
'{3A6AD9E0-23B9-11cf-AD60-00AA00A74CCD}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3A6AD9E0, CInt(&H23B9), CInt(&H11CF), &HAD, &H60, &H0, &HAA, &H0, &HA7, &H4C, &HCD)
 IID_ITransactionOptions = iid
End Function
Public Function IID_ITransactionOutcomeEvents() As UUID
'{3A6AD9E2-23B9-11cf-AD60-00AA00A74CCD}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3A6AD9E2, CInt(&H23B9), CInt(&H11CF), &HAD, &H60, &H0, &HAA, &H0, &HA7, &H4C, &HCD)
 IID_ITransactionOutcomeEvents = iid
End Function
Public Function IID_ITmNodeName() As UUID
'{30274F88-6EE4-474e-9B95-7807BC9EF8CF}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H30274F88, CInt(&H6EE4), CInt(&H474E), &H9B, &H95, &H78, &H7, &HBC, &H9E, &HF8, &HCF)
 IID_ITmNodeName = iid
End Function
Public Function IID_IKernelTransaction() As UUID
'{79427A2B-F895-40e0-BE79-B57DC82ED231}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H79427A2B, CInt(&HF895), CInt(&H40E0), &HBE, &H79, &HB5, &H7D, &HC8, &H2E, &HD2, &H31)
 IID_IKernelTransaction = iid
End Function
Public Function IID_INetworkListManager() As UUID
'{DCB00000-570F-4A9B-8D69-199FDBA5723B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDCB00000, CInt(&H570F), CInt(&H4A9B), &H8D, &H69, &H19, &H9F, &HDB, &HA5, &H72, &H3B)
IID_INetworkListManager = iid
End Function
Public Function IID_INetworkListManagerEvents() As UUID
'{DCB00001-570F-4A9B-8D69-199FDBA5723B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDCB00001, CInt(&H570F), CInt(&H4A9B), &H8D, &H69, &H19, &H9F, &HDB, &HA5, &H72, &H3B)
IID_INetworkListManagerEvents = iid
End Function
Public Function IID_INetwork() As UUID
'{DCB00002-570F-4A9B-8D69-199FDBA5723B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDCB00002, CInt(&H570F), CInt(&H4A9B), &H8D, &H69, &H19, &H9F, &HDB, &HA5, &H72, &H3B)
IID_INetwork = iid
End Function
Public Function IID_INetwork2() As UUID
'{B5550ABB-3391-4310-804F-25DCC325ED81}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB5550ABB, CInt(&H3391), CInt(&H4310), &H80, &H4F, &H25, &HDC, &HC3, &H25, &HED, &H81)
IID_INetwork2 = iid
End Function
Public Function IID_IEnumNetworks() As UUID
'{DCB00003-570F-4A9B-8D69-199FDBA5723B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDCB00003, CInt(&H570F), CInt(&H4A9B), &H8D, &H69, &H19, &H9F, &HDB, &HA5, &H72, &H3B)
IID_IEnumNetworks = iid
End Function
Public Function IID_INetworkEvents() As UUID
'{DCB00004-570F-4A9B-8D69-199FDBA5723B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDCB00004, CInt(&H570F), CInt(&H4A9B), &H8D, &H69, &H19, &H9F, &HDB, &HA5, &H72, &H3B)
IID_INetworkEvents = iid
End Function
Public Function IID_INetworkConnection() As UUID
'{DCB00005-570F-4A9B-8D69-199FDBA5723B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDCB00005, CInt(&H570F), CInt(&H4A9B), &H8D, &H69, &H19, &H9F, &HDB, &HA5, &H72, &H3B)
IID_INetworkConnection = iid
End Function
Public Function IID_INetworkConnection2() As UUID
'{00E676ED-5A35-4738-92EB-8581738D0F0A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE676ED, CInt(&H5A35), CInt(&H4738), &H92, &HEB, &H85, &H81, &H73, &H8D, &HF, &HA)
IID_INetworkConnection2 = iid
End Function
Public Function IID_IEnumNetworkConnections() As UUID
'{DCB00006-570F-4A9B-8D69-199FDBA5723B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDCB00006, CInt(&H570F), CInt(&H4A9B), &H8D, &H69, &H19, &H9F, &HDB, &HA5, &H72, &H3B)
IID_IEnumNetworkConnections = iid
End Function
Public Function IID_INetworkConnectionEvents() As UUID
'{DCB00007-570F-4A9B-8D69-199FDBA5723B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDCB00007, CInt(&H570F), CInt(&H4A9B), &H8D, &H69, &H19, &H9F, &HDB, &HA5, &H72, &H3B)
IID_INetworkConnectionEvents = iid
End Function
Public Function IID_INetworkCostManager() As UUID
'{DCB00008-570F-4A9B-8D69-199FDBA5723B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDCB00008, CInt(&H570F), CInt(&H4A9B), &H8D, &H69, &H19, &H9F, &HDB, &HA5, &H72, &H3B)
IID_INetworkCostManager = iid
End Function
Public Function IID_INetworkCostManagerEvents() As UUID
'{DCB00009-570F-4A9B-8D69-199FDBA5723B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDCB00009, CInt(&H570F), CInt(&H4A9B), &H8D, &H69, &H19, &H9F, &HDB, &HA5, &H72, &H3B)
IID_INetworkCostManagerEvents = iid
End Function
Public Function IID_INetworkConnectionCost() As UUID
'{DCB0000a-570F-4A9B-8D69-199FDBA5723B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDCB0000A, CInt(&H570F), CInt(&H4A9B), &H8D, &H69, &H19, &H9F, &HDB, &HA5, &H72, &H3B)
IID_INetworkConnectionCost = iid
End Function
Public Function IID_INetworkConnectionCostEvents() As UUID
'{DCB0000b-570F-4A9B-8D69-199FDBA5723B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDCB0000B, CInt(&H570F), CInt(&H4A9B), &H8D, &H69, &H19, &H9F, &HDB, &HA5, &H72, &H3B)
IID_INetworkConnectionCostEvents = iid
End Function
Public Function IID_ICredentialProviderCredential() As UUID
'{63913a93-40c1-481a-818d-4072ff8c70cc}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H63913A93, CInt(&H40C1), CInt(&H481A), &H81, &H8D, &H40, &H72, &HFF, &H8C, &H70, &HCC)
IID_ICredentialProviderCredential = iid
End Function
Public Function IID_IQueryContinueWithStatus() As UUID
'{9090be5b-502b-41fb-bccc-0049a6c7254b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9090BE5B, CInt(&H502B), CInt(&H41FB), &HBC, &HCC, &H0, &H49, &HA6, &HC7, &H25, &H4B)
IID_IQueryContinueWithStatus = iid
End Function
Public Function IID_IConnectableCredentialProviderCredential() As UUID
'{9387928b-ac75-4bf9-8ab2-2b93c4a55290}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9387928B, CInt(&HAC75), CInt(&H4BF9), &H8A, &HB2, &H2B, &H93, &HC4, &HA5, &H52, &H90)
IID_IConnectableCredentialProviderCredential = iid
End Function
Public Function IID_ICredentialProviderCredentialEvents() As UUID
'{fa6fa76b-66b7-4b11-95f1-86171118e816}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA6FA76B, CInt(&H66B7), CInt(&H4B11), &H95, &HF1, &H86, &H17, &H11, &H18, &HE8, &H16)
IID_ICredentialProviderCredentialEvents = iid
End Function
Public Function IID_ICredentialProvider() As UUID
'{d27c3481-5a1c-45b2-8aaa-c20ebbe8229e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD27C3481, CInt(&H5A1C), CInt(&H45B2), &H8A, &HAA, &HC2, &HE, &HBB, &HE8, &H22, &H9E)
IID_ICredentialProvider = iid
End Function
Public Function IID_ICredentialProviderEvents() As UUID
'{34201e5a-a787-41a3-a5a4-bd6dcf2a854e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H34201E5A, CInt(&HA787), CInt(&H41A3), &HA5, &HA4, &HBD, &H6D, &HCF, &H2A, &H85, &H4E)
IID_ICredentialProviderEvents = iid
End Function
Public Function IID_ICredentialProviderFilter() As UUID
'{a5da53f9-d475-4080-a120-910c4a739880}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA5DA53F9, CInt(&HD475), CInt(&H4080), &HA1, &H20, &H91, &HC, &H4A, &H73, &H98, &H80)
IID_ICredentialProviderFilter = iid
End Function
Public Function IID_ICredentialProviderCredential2() As UUID
'{fd672c54-40ea-4d6e-9b49-cfb1a7507bd7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFD672C54, CInt(&H40EA), CInt(&H4D6E), &H9B, &H49, &HCF, &HB1, &HA7, &H50, &H7B, &HD7)
IID_ICredentialProviderCredential2 = iid
End Function
Public Function IID_ICredentialProviderCredentialWithFieldOptions() As UUID
'{DBC6FB30-C843-49E3-A645-573E6F39446A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDBC6FB30, CInt(&HC843), CInt(&H49E3), &HA6, &H45, &H57, &H3E, &H6F, &H39, &H44, &H6A)
IID_ICredentialProviderCredentialWithFieldOptions = iid
End Function
Public Function IID_ICredentialProviderCredentialEvents2() As UUID
'{B53C00B6-9922-4B78-B1F4-DDFE774DC39B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB53C00B6, CInt(&H9922), CInt(&H4B78), &HB1, &HF4, &HDD, &HFE, &H77, &H4D, &HC3, &H9B)
IID_ICredentialProviderCredentialEvents2 = iid
End Function
Public Function IID_ICredentialProviderUser() As UUID
'{13793285-3ea6-40fd-b420-15f47da41fbb}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H13793285, CInt(&H3EA6), CInt(&H40FD), &HB4, &H20, &H15, &HF4, &H7D, &HA4, &H1F, &HBB)
IID_ICredentialProviderUser = iid
End Function
Public Function IID_ICredentialProviderUserArray() As UUID
'{90C119AE-0F18-4520-A1F1-114366A40FE8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H90C119AE, CInt(&HF18), CInt(&H4520), &HA1, &HF1, &H11, &H43, &H66, &HA4, &HF, &HE8)
IID_ICredentialProviderUserArray = iid
End Function
Public Function IID_ICredentialProviderSetUserArray() As UUID
'{095c1484-1c0c-4388-9c6d-500e61bf84bd}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H95C1484, CInt(&H1C0C), CInt(&H4388), &H9C, &H6D, &H50, &HE, &H61, &HBF, &H84, &HBD)
IID_ICredentialProviderSetUserArray = iid
End Function
Public Function IID_IThumbnailStreamCache() As UUID
'{90E11430-9569-41D8-AE75-6D4D2AE7CCA0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H90E11430, CInt(&H9569), CInt(&H41D8), &HAE, &H75, &H6D, &H4D, &H2A, &HE7, &HCC, &HA0)
 IID_IThumbnailStreamCache = iid
End Function
Public Function IID_IUIAnimationTimerUpdateHandler() As UUID
'{195509B7-5D5E-4e3e-B278-EE3759B367AD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H195509B7, CInt(&H5D5E), CInt(&H4E3E), &HB2, &H78, &HEE, &H37, &H59, &HB3, &H67, &HAD)
IID_IUIAnimationTimerUpdateHandler = iid
End Function
Public Function IID_IUIAnimationTimerClientEventHandler() As UUID
'{BEDB4DB6-94FA-4bfb-A47F-EF2D9E408C25}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBEDB4DB6, CInt(&H94FA), CInt(&H4BFB), &HA4, &H7F, &HEF, &H2D, &H9E, &H40, &H8C, &H25)
IID_IUIAnimationTimerClientEventHandler = iid
End Function
Public Function IID_IUIAnimationTimerEventHandler() As UUID
'{274A7DEA-D771-4095-ABBD-8DF7ABD23CE3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H274A7DEA, CInt(&HD771), CInt(&H4095), &HAB, &HBD, &H8D, &HF7, &HAB, &HD2, &H3C, &HE3)
IID_IUIAnimationTimerEventHandler = iid
End Function
Public Function IID_IUIAnimationManager2() As UUID
'{D8B6F7D4-4109-4d3f-ACEE-879926968CB1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD8B6F7D4, CInt(&H4109), CInt(&H4D3F), &HAC, &HEE, &H87, &H99, &H26, &H96, &H8C, &HB1)
IID_IUIAnimationManager2 = iid
End Function
Public Function IID_IUIAnimationVariable2() As UUID
'{4914B304-96AB-44d9-9E77-D5109B7E7466}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4914B304, CInt(&H96AB), CInt(&H44D9), &H9E, &H77, &HD5, &H10, &H9B, &H7E, &H74, &H66)
IID_IUIAnimationVariable2 = iid
End Function
Public Function IID_IUIAnimationTransition2() As UUID
'{62FF9123-A85A-4e9b-A218-435A93E268FD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H62FF9123, CInt(&HA85A), CInt(&H4E9B), &HA2, &H18, &H43, &H5A, &H93, &HE2, &H68, &HFD)
IID_IUIAnimationTransition2 = iid
End Function
Public Function IID_IUIAnimationManagerEventHandler2() As UUID
'{F6E022BA-BFF3-42EC-9033-E073F33E83C3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF6E022BA, CInt(&HBFF3), CInt(&H42EC), &H90, &H33, &HE0, &H73, &HF3, &H3E, &H83, &HC3)
IID_IUIAnimationManagerEventHandler2 = iid
End Function
Public Function IID_IUIAnimationVariableChangeHandler2() As UUID
'{63ACC8D2-6EAE-4bb0-B879-586DD8CFBE42}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H63ACC8D2, CInt(&H6EAE), CInt(&H4BB0), &HB8, &H79, &H58, &H6D, &HD8, &HCF, &HBE, &H42)
IID_IUIAnimationVariableChangeHandler2 = iid
End Function
Public Function IID_IUIAnimationVariableIntegerChangeHandler2() As UUID
'{829B6CF1-4F3A-4412-AE09-B243EB4C6B58}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H829B6CF1, CInt(&H4F3A), CInt(&H4412), &HAE, &H9, &HB2, &H43, &HEB, &H4C, &H6B, &H58)
IID_IUIAnimationVariableIntegerChangeHandler2 = iid
End Function
Public Function IID_IUIAnimationVariableCurveChangeHandler2() As UUID
'{72895E91-0145-4C21-9192-5AAB40EDDF80}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H72895E91, CInt(&H145), CInt(&H4C21), &H91, &H92, &H5A, &HAB, &H40, &HED, &HDF, &H80)
IID_IUIAnimationVariableCurveChangeHandler2 = iid
End Function
Public Function IID_IUIAnimationStoryboardEventHandler2() As UUID
'{BAC5F55A-BA7C-414C-B599-FBF850F553C6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBAC5F55A, CInt(&HBA7C), CInt(&H414C), &HB5, &H99, &HFB, &HF8, &H50, &HF5, &H53, &HC6)
IID_IUIAnimationStoryboardEventHandler2 = iid
End Function
Public Function IID_IUIAnimationLoopIterationChangeHandler2() As UUID
'{2D3B15A4-4762-47AB-A030-B23221DF3AE0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2D3B15A4, CInt(&H4762), CInt(&H47AB), &HA0, &H30, &HB2, &H32, &H21, &HDF, &H3A, &HE0)
IID_IUIAnimationLoopIterationChangeHandler2 = iid
End Function
Public Function IID_IUIAnimationPriorityComparison2() As UUID
'{5B6D7A37-4621-467C-8B05-70131DE62DDB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5B6D7A37, CInt(&H4621), CInt(&H467C), &H8B, &H5, &H70, &H13, &H1D, &HE6, &H2D, &HDB)
IID_IUIAnimationPriorityComparison2 = iid
End Function
Public Function IID_IUIAnimationTransitionLibrary2() As UUID
'{03CFAE53-9580-4ee3-B363-2ECE51B4AF6A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3CFAE53, CInt(&H9580), CInt(&H4EE3), &HB3, &H63, &H2E, &HCE, &H51, &HB4, &HAF, &H6A)
IID_IUIAnimationTransitionLibrary2 = iid
End Function
Public Function IID_IUIAnimationPrimitiveInterpolation() As UUID
'{BAB20D63-4361-45DA-A24F-AB8508846B5B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBAB20D63, CInt(&H4361), CInt(&H45DA), &HA2, &H4F, &HAB, &H85, &H8, &H84, &H6B, &H5B)
IID_IUIAnimationPrimitiveInterpolation = iid
End Function
Public Function IID_IUIAnimationInterpolator2() As UUID
'{EA76AFF8-EA22-4a23-A0EF-A6A966703518}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEA76AFF8, CInt(&HEA22), CInt(&H4A23), &HA0, &HEF, &HA6, &HA9, &H66, &H70, &H35, &H18)
IID_IUIAnimationInterpolator2 = iid
End Function
Public Function IID_IUIAnimationTransitionFactory2() As UUID
'{937D4916-C1A6-42d5-88D8-30344D6EFE31}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H937D4916, CInt(&HC1A6), CInt(&H42D5), &H88, &HD8, &H30, &H34, &H4D, &H6E, &HFE, &H31)
IID_IUIAnimationTransitionFactory2 = iid
End Function
Public Function IID_IUIAnimationStoryboard2() As UUID
'{AE289CD2-12D4-4945-9419-9E41BE034DF2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAE289CD2, CInt(&H12D4), CInt(&H4945), &H94, &H19, &H9E, &H41, &HBE, &H3, &H4D, &HF2)
IID_IUIAnimationStoryboard2 = iid
End Function
Public Function IID_IUIAnimationManager() As UUID
'{9169896C-AC8D-4e7d-94E5-67FA4DC2F2E8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9169896C, CInt(&HAC8D), CInt(&H4E7D), &H94, &HE5, &H67, &HFA, &H4D, &HC2, &HF2, &HE8)
IID_IUIAnimationManager = iid
End Function
Public Function IID_IUIAnimationVariable() As UUID
'{8CEEB155-2849-4ce5-9448-91FF70E1E4D9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CEEB155, CInt(&H2849), CInt(&H4CE5), &H94, &H48, &H91, &HFF, &H70, &HE1, &HE4, &HD9)
IID_IUIAnimationVariable = iid
End Function
Public Function IID_IUIAnimationStoryboard() As UUID
'{A8FF128F-9BF9-4af1-9E67-E5E410DEFB84}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA8FF128F, CInt(&H9BF9), CInt(&H4AF1), &H9E, &H67, &HE5, &HE4, &H10, &HDE, &HFB, &H84)
IID_IUIAnimationStoryboard = iid
End Function
Public Function IID_IUIAnimationTransition() As UUID
'{DC6CE252-F731-41cf-B610-614B6CA049AD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDC6CE252, CInt(&HF731), CInt(&H41CF), &HB6, &H10, &H61, &H4B, &H6C, &HA0, &H49, &HAD)
IID_IUIAnimationTransition = iid
End Function
Public Function IID_IUIAnimationManagerEventHandler() As UUID
'{783321ED-78A3-4366-B574-6AF607A64788}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H783321ED, CInt(&H78A3), CInt(&H4366), &HB5, &H74, &H6A, &HF6, &H7, &HA6, &H47, &H88)
IID_IUIAnimationManagerEventHandler = iid
End Function
Public Function IID_IUIAnimationVariableChangeHandler() As UUID
'{6358B7BA-87D2-42d5-BF71-82E919DD5862}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6358B7BA, CInt(&H87D2), CInt(&H42D5), &HBF, &H71, &H82, &HE9, &H19, &HDD, &H58, &H62)
IID_IUIAnimationVariableChangeHandler = iid
End Function
Public Function IID_IUIAnimationVariableIntegerChangeHandler() As UUID
'{BB3E1550-356E-44b0-99DA-85AC6017865E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBB3E1550, CInt(&H356E), CInt(&H44B0), &H99, &HDA, &H85, &HAC, &H60, &H17, &H86, &H5E)
IID_IUIAnimationVariableIntegerChangeHandler = iid
End Function
Public Function IID_IUIAnimationStoryboardEventHandler() As UUID
'{3D5C9008-EC7C-4364-9F8A-9AF3C58CBAE6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3D5C9008, CInt(&HEC7C), CInt(&H4364), &H9F, &H8A, &H9A, &HF3, &HC5, &H8C, &HBA, &HE6)
IID_IUIAnimationStoryboardEventHandler = iid
End Function
Public Function IID_IUIAnimationPriorityComparison() As UUID
'{83FA9B74-5F86-4618-BC6A-A2FAC19B3F44}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H83FA9B74, CInt(&H5F86), CInt(&H4618), &HBC, &H6A, &HA2, &HFA, &HC1, &H9B, &H3F, &H44)
IID_IUIAnimationPriorityComparison = iid
End Function
Public Function IID_IUIAnimationTransitionLibrary() As UUID
'{CA5A14B1-D24F-48b8-8FE4-C78169BA954E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCA5A14B1, CInt(&HD24F), CInt(&H48B8), &H8F, &HE4, &HC7, &H81, &H69, &HBA, &H95, &H4E)
IID_IUIAnimationTransitionLibrary = iid
End Function
Public Function IID_IUIAnimationInterpolator() As UUID
'{7815CBBA-DDF7-478c-A46C-7B6C738B7978}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7815CBBA, CInt(&HDDF7), CInt(&H478C), &HA4, &H6C, &H7B, &H6C, &H73, &H8B, &H79, &H78)
IID_IUIAnimationInterpolator = iid
End Function
Public Function IID_IUIAnimationTransitionFactory() As UUID
'{FCD91E03-3E3B-45ad-BBB1-6DFC8153743D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFCD91E03, CInt(&H3E3B), CInt(&H45AD), &HBB, &HB1, &H6D, &HFC, &H81, &H53, &H74, &H3D)
IID_IUIAnimationTransitionFactory = iid
End Function
Public Function IID_IUIAnimationTimer() As UUID
'{6B0EFAD1-A053-41d6-9085-33A689144665}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6B0EFAD1, CInt(&HA053), CInt(&H41D6), &H90, &H85, &H33, &HA6, &H89, &H14, &H46, &H65)
IID_IUIAnimationTimer = iid
End Function
Public Function IID_IThumbnailCachePrimer() As UUID
'{0f03f8fe-2b26-46f0-965a-212aa8d66b76}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF03F8FE, CInt(&H2B26), CInt(&H46F0), &H97, &H5A, &H21, &H2A, &HA8, &HD6, &H6B, &H76)
IID_IThumbnailCachePrimer = iid
End Function
Public Function IID_IMediaRadioManager() As UUID
'{6CFDCAB5-FC47-42A5-9241-074B58830E73}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6CFDCAB5, CInt(&HFC47), CInt(&H42A5), &H92, &H41, &H7, &H4B, &H58, &H83, &HE, &H73)
IID_IMediaRadioManager = iid
End Function
Public Function IID_IRadioInstanceCollection() As UUID
'{E5791FAE-5665-4E0C-95BE-5FDE31644185}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE5791FAE, CInt(&H5665), CInt(&H4E0C), &H95, &HBE, &H5F, &HDE, &H31, &H64, &H41, &H85)
IID_IRadioInstanceCollection = iid
End Function
Public Function IID_IRadioInstance() As UUID
'{70AA1C9E-F2B4-4C61-86D3-6B9FB75FD1A2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H70AA1C9E, CInt(&HF2B4), CInt(&H4C61), &H86, &HD3, &H6B, &H9F, &HB7, &H5F, &HD1, &HA2)
IID_IRadioInstance = iid
End Function
Public Function IID_IMediaRadioManagerNotifySink() As UUID
'{89D81F5F-C147-49ED-A11C-77B20C31E7C9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H89D81F5F, CInt(&HC147), CInt(&H49ED), &HA1, &H1C, &H77, &HB2, &HC, &H31, &HE7, &HC9)
IID_IMediaRadioManagerNotifySink = iid
End Function
Public Function IID_IPhotoAcquireItem() As UUID
'{00f21c97-28bf-4c02-b842-5e4e90139a30}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF21C97, CInt(&H28BF), CInt(&H4C02), &HB8, &H42, &H5E, &H4E, &H90, &H13, &H9A, &H30)
IID_IPhotoAcquireItem = iid
End Function
Public Function IID_IUserInputString() As UUID
'{00f243a1-205b-45ba-ae26-abbc53aa7a6f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF243A1, CInt(&H205B), CInt(&H45BA), &HAE, &H26, &HAB, &HBC, &H53, &HAA, &H7A, &H6F)
IID_IUserInputString = iid
End Function
Public Function IID_IPhotoAcquireProgressCB() As UUID
'{00f2ce1e-935e-4248-892c-130f32c45cb4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF2CE1E, CInt(&H935E), CInt(&H4248), &H89, &H2C, &H13, &HF, &H32, &HC4, &H5C, &HB4)
IID_IPhotoAcquireProgressCB = iid
End Function
Public Function IID_IPhotoProgressActionCB() As UUID
'{00f242d0-b206-4e7d-b4c1-4755bcbb9c9f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF242D0, CInt(&HB206), CInt(&H4E7D), &HB4, &HC1, &H47, &H55, &HBC, &HBB, &H9C, &H9F)
IID_IPhotoProgressActionCB = iid
End Function
Public Function IID_IPhotoProgressDialog() As UUID
'{00f246f9-0750-4f08-9381-2cd8e906a4ae}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF246F9, CInt(&H750), CInt(&H4F08), &H93, &H81, &H2C, &HD8, &HE9, &H6, &HA4, &HAE)
IID_IPhotoProgressDialog = iid
End Function
Public Function IID_IPhotoAcquireSource() As UUID
'{00f2c703-8613-4282-a53b-6ec59c5883ac}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF2C703, CInt(&H8613), CInt(&H4282), &HA5, &H3B, &H6E, &HC5, &H9C, &H58, &H83, &HAC)
IID_IPhotoAcquireSource = iid
End Function
Public Function IID_IPhotoAcquire() As UUID
'{00f23353-e31b-4955-a8ad-ca5ebf31e2ce}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF23353, CInt(&HE31B), CInt(&H4955), &HA8, &HAD, &HCA, &H5E, &HBF, &H31, &HE2, &HCE)
IID_IPhotoAcquire = iid
End Function
Public Function IID_IPhotoAcquireSettings() As UUID
'{00f2b868-dd67-487c-9553-049240767e91}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF2B868, CInt(&HDD67), CInt(&H487C), &H95, &H53, &H4, &H92, &H40, &H76, &H7E, &H91)
IID_IPhotoAcquireSettings = iid
End Function
Public Function IID_IPhotoAcquireOptionsDialog() As UUID
'{00f2b3ee-bf64-47ee-89f4-4dedd79643f2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF2B3EE, CInt(&HBF64), CInt(&H47EE), &H89, &HF4, &H4D, &HED, &HD7, &H96, &H43, &HF2)
IID_IPhotoAcquireOptionsDialog = iid
End Function
Public Function IID_IPhotoAcquireDeviceSelectionDialog() As UUID
'{00f28837-55dd-4f37-aaf5-6855a9640467}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF28837, CInt(&H55DD), CInt(&H4F37), &HAA, &HF5, &H68, &H55, &HA9, &H64, &H4, &H67)
IID_IPhotoAcquireDeviceSelectionDialog = iid
End Function
Public Function IID_IPhotoAcquirePlugin() As UUID
'{00f2dceb-ecb8-4f77-8e47-e7a987c83dd0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF2DCEB, CInt(&HECB8), CInt(&H4F77), &H8E, &H47, &HE7, &HA9, &H87, &HC8, &H3D, &HD0)
IID_IPhotoAcquirePlugin = iid
End Function
Public Function IID_ISyncMgrSynchronizeCallback() As UUID
'{6295DF41-35EE-11d1-8707-00C04FD93327}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6295DF41, CInt(&H35EE), CInt(&H11D1), &H87, &H7, &H0, &HC0, &H4F, &HD9, &H33, &H27)
IID_ISyncMgrSynchronizeCallback = iid
End Function
Public Function IID_ISyncMgrEnumItems() As UUID
'{6295DF2A-35EE-11d1-8707-00C04FD93327}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6295DF2A, CInt(&H35EE), CInt(&H11D1), &H87, &H7, &H0, &HC0, &H4F, &HD9, &H33, &H27)
 IID_ISyncMgrEnumItems = iid
End Function
Public Function IID_ISyncMgrSynchronize() As UUID
'{6295DF40-35EE-11d1-8707-00C04FD93327}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6295DF40, CInt(&H35EE), CInt(&H11D1), &H87, &H7, &H0, &HC0, &H4F, &HD9, &H33, &H27)
IID_ISyncMgrSynchronize = iid
End Function
Public Function IID_ISyncMgrSynchronizeInvoke() As UUID
'{6295DF2C-35EE-11d1-8707-00C04FD93327}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6295DF2C, CInt(&H35EE), CInt(&H11D1), &H87, &H7, &H0, &HC0, &H4F, &HD9, &H33, &H27)
IID_ISyncMgrSynchronizeInvoke = iid
End Function
Public Function IID_ISyncMgrRegister() As UUID
'{6295DF42-35EE-11d1-8707-00C04FD93327}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6295DF42, CInt(&H35EE), CInt(&H11D1), &H87, &H7, &H0, &HC0, &H4F, &HD9, &H33, &H27)
IID_ISyncMgrRegister = iid
End Function
Public Function IID_IScopeItem() As UUID
'{DD400FF4-A119-405F-970E-A9A5E7E828C0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDD400FF4, CInt(&HA119), CInt(&H405F), &H97, &HE, &HA9, &HA5, &HE7, &HE8, &H28, &HC0)
 IID_IScopeItem = iid
End Function
Public Function IID_IScope() As UUID
'{655D1685-2BFD-4F7F-AD22-5AB61C8D8798}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H655D1685, CInt(&H2BFD), CInt(&H4F7F), &HAD, &H22, &H5A, &HB6, &H1C, &H8D, &H87, &H98)
 IID_IScope = iid
End Function
Public Function IID_IBindScopeDialog() As UUID
'{655D1685-2BFD-4F7F-AD22-5AB61C8D8798}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H655D1685, CInt(&H2BFD), CInt(&H4F7F), &HAD, &H22, &H5A, &HB6, &H1C, &H8D, &H87, &H98)
 IID_IBindScopeDialog = iid
End Function
Public Function IID_IAttachmentExecute() As UUID
'{73db1241-1e85-4581-8e4f-a81e1d0f8c57}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H73DB1241, CInt(&H1E85), CInt(&H4581), &H8E, &H4F, &HA8, &H1E, &H1D, &HF, &H8C, &H57)
IID_IAttachmentExecute = iid
End Function
Public Function IID_IStorageProviderBanners() As UUID
'{5efb46d7-47C0-4b68-acda-ded47c90ec91}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5EFB46D7, CInt(&H47C0), CInt(&H4B68), &HAC, &HDA, &HDE, &HD4, &H7C, &H90, &HEC, &H91)
IID_IStorageProviderBanners = iid
End Function
Public Function IID_IVBGetControl() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H40A050A0, &H3C31, &H101B, &HA8, &H2E, &H8, &H0, &H2B, &H2B, &H23, &H37)
IID_IVBGetControl = iid
End Function
Public Function IID_IGetOleObject() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8A701DA0, &H4FEB, &H101B, &HA8, &H2E, &H8, &H0, &H2B, &H2B, &H23, &H37)
IID_IGetOleObject = iid
End Function
Public Function IID_IGetVBAObject() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H91733A60, &H3F4C, &H101B, &HA3, &HF6, &H0, &HAA, &H0, &H34, &HE4, &HE9)
IID_IGetVBAObject = iid
End Function
Public Function IID_IVBFormat() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9849FD60, &H3768, &H101B, &H8D, &H72, &HAE, &H61, &H64, &HFF, &HE3, &HCF)
IID_IVBFormat = iid
End Function
Public Function IID_IUPnPDeviceFinder() As UUID
'{ADDA3D55-6F72-4319-BFF9-18600A539B10}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HADDA3D55, CInt(&H6F72), CInt(&H4319), &HBF, &HF9, &H18, &H60, &HA, &H53, &H9B, &H10)
IID_IUPnPDeviceFinder = iid
End Function
Public Function IID_IUPnPAddressFamilyControl() As UUID
'{E3BF6178-694E-459F-A5A6-191EA0FFA1C7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE3BF6178, CInt(&H694E), CInt(&H459F), &HA5, &HA6, &H19, &H1E, &HA0, &HFF, &HA1, &HC7)
IID_IUPnPAddressFamilyControl = iid
End Function
Public Function IID_IUPnPHttpHeaderControl() As UUID
'{0405AF4F-8B5C-447C-80F2-B75984A31F3C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H405AF4F, CInt(&H8B5C), CInt(&H447C), &H80, &HF2, &HB7, &H59, &H84, &HA3, &H1F, &H3C)
IID_IUPnPHttpHeaderControl = iid
End Function
Public Function IID_IUPnPDeviceFinderCallback() As UUID
'{415A984A-88B3-49F3-92AF-0508BEDF0D6C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H415A984A, CInt(&H88B3), CInt(&H49F3), &H92, &HAF, &H5, &H8, &HBE, &HDF, &HD, &H6C)
IID_IUPnPDeviceFinderCallback = iid
End Function
Public Function IID_IUPnPServices() As UUID
'{3F8C8E9E-9A7A-4DC8-BC41-FF31FA374956}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3F8C8E9E, CInt(&H9A7A), CInt(&H4DC8), &HBC, &H41, &HFF, &H31, &HFA, &H37, &H49, &H56)
IID_IUPnPServices = iid
End Function
Public Function IID_IUPnPService() As UUID
'{A295019C-DC65-47DD-90DC-7FE918A1AB44}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA295019C, CInt(&HDC65), CInt(&H47DD), &H90, &HDC, &H7F, &HE9, &H18, &HA1, &HAB, &H44)
IID_IUPnPService = iid
End Function
Public Function IID_IUPnPAsyncResult() As UUID
'{4D65FD08-D13E-4274-9C8B-DD8D028C8644}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4D65FD08, CInt(&HD13E), CInt(&H4274), &H9C, &H8B, &HDD, &H8D, &H2, &H8C, &H86, &H44)
IID_IUPnPAsyncResult = iid
End Function
Public Function IID_IUPnPServiceAsync() As UUID
'{098BDAF5-5EC1-49e7-A260-B3A11DD8680C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H98BDAF5, CInt(&H5EC1), CInt(&H49E7), &HA2, &H60, &HB3, &HA1, &H1D, &HD8, &H68, &HC)
IID_IUPnPServiceAsync = iid
End Function
Public Function IID_IUPnPServiceCallback() As UUID
'{31fadca9-ab73-464b-b67d-5c1d0f83c8b8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H31FADCA9, CInt(&HAB73), CInt(&H464B), &HB6, &H7D, &H5C, &H1D, &HF, &H83, &HC8, &HB8)
IID_IUPnPServiceCallback = iid
End Function
Public Function IID_IUPnPServiceEnumProperty() As UUID
'{38873B37-91BB-49f4-B249-2E8EFBB8A816}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H38873B37, CInt(&H91BB), CInt(&H49F4), &HB2, &H49, &H2E, &H8E, &HFB, &HB8, &HA8, &H16)
IID_IUPnPServiceEnumProperty = iid
End Function
Public Function IID_IUPnPServiceDocumentAccess() As UUID
'{21905529-0A5E-4589-825D-7E6D87EA6998}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H21905529, CInt(&HA5E), CInt(&H4589), &H82, &H5D, &H7E, &H6D, &H87, &HEA, &H69, &H98)
IID_IUPnPServiceDocumentAccess = iid
End Function
Public Function IID_IUPnPDevices() As UUID
'{FDBC0C73-BDA3-4C66-AC4F-F2D96FDAD68C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFDBC0C73, CInt(&HBDA3), CInt(&H4C66), &HAC, &H4F, &HF2, &HD9, &H6F, &HDA, &HD6, &H8C)
IID_IUPnPDevices = iid
End Function
Public Function IID_IUPnPDevice() As UUID
'{3D44D0D1-98C9-4889-ACD1-F9D674BF2221}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3D44D0D1, CInt(&H98C9), CInt(&H4889), &HAC, &HD1, &HF9, &HD6, &H74, &HBF, &H22, &H21)
IID_IUPnPDevice = iid
End Function
Public Function IID_IUPnPDeviceDocumentAccess() As UUID
'{E7772804-3287-418e-9072-CF2B47238981}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE7772804, CInt(&H3287), CInt(&H418E), &H90, &H72, &HCF, &H2B, &H47, &H23, &H89, &H81)
IID_IUPnPDeviceDocumentAccess = iid
End Function
Public Function IID_IUPnPDeviceDocumentAccessEx() As UUID
'{C4BC4050-6178-4BD1-A4B8-6398321F3247}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC4BC4050, CInt(&H6178), CInt(&H4BD1), &HA4, &HB8, &H63, &H98, &H32, &H1F, &H32, &H47)
IID_IUPnPDeviceDocumentAccessEx = iid
End Function
Public Function IID_IUPnPDescriptionDocument() As UUID
'{11d1c1b2-7daa-4c9e-9595-7f82ed206d1e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H11D1C1B2, CInt(&H7DAA), CInt(&H4C9E), &H95, &H95, &H7F, &H82, &HED, &H20, &H6D, &H1E)
IID_IUPnPDescriptionDocument = iid
End Function
Public Function IID_IUPnPDeviceFinderAddCallbackWithInterface() As UUID
'{983dfc0b-1796-44df-8975-ca545b620ee5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H983DFC0B, CInt(&H1796), CInt(&H44DF), &H89, &H75, &HCA, &H54, &H5B, &H62, &HE, &HE5)
IID_IUPnPDeviceFinderAddCallbackWithInterface = iid
End Function
Public Function IID_IUPnPDescriptionDocumentCallback() As UUID
'{77394c69-5486-40d6-9bc3-4991983e02da}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H77394C69, CInt(&H5486), CInt(&H40D6), &H9B, &HC3, &H49, &H91, &H98, &H3E, &H2, &HDA)
IID_IUPnPDescriptionDocumentCallback = iid
End Function
Public Function IID_IUPnPEventSink() As UUID
'{204810b4-73b2-11d4-bf42-00b0d0118b56}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H204810B4, CInt(&H73B2), CInt(&H11D4), &HBF, &H42, &H0, &HB0, &HD0, &H11, &H8B, &H56)
IID_IUPnPEventSink = iid
End Function
Public Function IID_IUPnPEventSource() As UUID
'{204810b5-73b2-11d4-bf42-00b0d0118b56}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H204810B5, CInt(&H73B2), CInt(&H11D4), &HBF, &H42, &H0, &HB0, &HD0, &H11, &H8B, &H56)
IID_IUPnPEventSource = iid
End Function
Public Function IID_IUPnPRegistrar() As UUID
'{204810b6-73b2-11d4-bf42-00b0d0118b56}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H204810B6, CInt(&H73B2), CInt(&H11D4), &HBF, &H42, &H0, &HB0, &HD0, &H11, &H8B, &H56)
IID_IUPnPRegistrar = iid
End Function
Public Function IID_IUPnPReregistrar() As UUID
'{204810b7-73b2-11d4-bf42-00b0d0118b56}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H204810B7, CInt(&H73B2), CInt(&H11D4), &HBF, &H42, &H0, &HB0, &HD0, &H11, &H8B, &H56)
IID_IUPnPReregistrar = iid
End Function
Public Function IID_IUPnPDeviceControl() As UUID
'{204810ba-73b2-11d4-bf42-00b0d0118b56}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H204810BA, CInt(&H73B2), CInt(&H11D4), &HBF, &H42, &H0, &HB0, &HD0, &H11, &H8B, &H56)
IID_IUPnPDeviceControl = iid
End Function
Public Function IID_IUPnPDeviceControlHttpHeaders() As UUID
'{204810bb-73b2-11d4-bf42-00b0d0118b56}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H204810BB, CInt(&H73B2), CInt(&H11D4), &HBF, &H42, &H0, &HB0, &HD0, &H11, &H8B, &H56)
IID_IUPnPDeviceControlHttpHeaders = iid
End Function
Public Function IID_IUPnPDeviceProvider() As UUID
'{204810b8-73b2-11d4-bf42-00b0d0118b56}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H204810B8, CInt(&H73B2), CInt(&H11D4), &HBF, &H42, &H0, &HB0, &HD0, &H11, &H8B, &H56)
IID_IUPnPDeviceProvider = iid
End Function
Public Function IID_IUPnPRemoteEndpointInfo() As UUID
'{c92eb863-0269-4aff-9c72-75321bba2952}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC92EB863, CInt(&H269), CInt(&H4AFF), &H9C, &H72, &H75, &H32, &H1B, &HBA, &H29, &H52)
IID_IUPnPRemoteEndpointInfo = iid
End Function
Public Function IID_IObjectWithPackageFullName() As UUID
'{ED2AA515-602F-469C-A130-CE69FD0FA878}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HED2AA515, CInt(&H602F), CInt(&H469C), &HA1, &H30, &HCE, &H69, &HFD, &HF, &HA8, &H78)
IID_IObjectWithPackageFullName = iid
End Function
Public Function IID_ITipAutoCompleteProvider() As UUID
'{7C6CF46D-8404-46b9-AD33-F5B6036D4007}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7C6CF46D, CInt(&H8404), CInt(&H46B9), &HAD, &H33, &HF5, &HB6, &H3, &H6D, &H40, &H7)
IID_ITipAutoCompleteProvider = iid
End Function
Public Function IID_ITipAutoCompleteClient() As UUID
'{5E078E03-8265-4bbe-9487-D242EDBEF910}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5E078E03, CInt(&H8265), CInt(&H4BBE), &H94, &H87, &HD2, &H42, &HED, &HBE, &HF9, &H10)
IID_ITipAutoCompleteClient = iid
End Function
Public Function IID_IWinMLModel() As UUID
'{e2eeb6a9-f31f-4055-a521-e30b5b33664a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE2EEB6A9, CInt(&HF31F), CInt(&H4055), &HA5, &H21, &HE3, &HB, &H5B, &H33, &H66, &H4A)
IID_IWinMLModel = iid
End Function
Public Function IID_IWinMLEvaluationContext() As UUID
'{95848f9e-583d-4054-af12-916387cd8426}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H95848F9E, CInt(&H583D), CInt(&H4054), &HAF, &H12, &H91, &H63, &H87, &HCD, &H84, &H26)
IID_IWinMLEvaluationContext = iid
End Function
Public Function IID_IWinMLRuntime() As UUID
'{a0425329-40ae-48d9-bce3-829ef7b8a41a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA0425329, CInt(&H40AE), CInt(&H48D9), &HBC, &HE3, &H82, &H9E, &HF7, &HB8, &HA4, &H1A)
IID_IWinMLRuntime = iid
End Function
Public Function IID_IWinMLRuntimeFactory() As UUID
'{a807b84d-4ae5-4bc0-a76a-941aa246bd41}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA807B84D, CInt(&H4AE5), CInt(&H4BC0), &HA7, &H6A, &H94, &H1A, &HA2, &H46, &HBD, &H41)
IID_IWinMLRuntimeFactory = iid
End Function

Public Function GUID_DEVINTERFACE_SENSOR() As UUID
'{BA1BB692-9B7A-4833-9A1E-525ED134E7E2}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBA1BB692, CInt(&H9B7A), CInt(&H4833), &H9A, &H1E, &H52, &H5E, &HD1, &H34, &HE7, &HE2)
 GUID_DEVINTERFACE_SENSOR = iid
End Function
Public Function SENSOR_EVENT_STATE_CHANGED() As UUID
'{BFD96016-6BD7-4560-AD34-F2F6607E8F81}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBFD96016, CInt(&H6BD7), CInt(&H4560), &HAD, &H34, &HF2, &HF6, &H60, &H7E, &H8F, &H81)
 SENSOR_EVENT_STATE_CHANGED = iid
End Function
Public Function SENSOR_EVENT_DATA_UPDATED() As UUID
'{2ED0F2A4-0087-41D3-87DB-6773370B3C88}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2ED0F2A4, CInt(&H87), CInt(&H41D3), &H87, &HDB, &H67, &H73, &H37, &HB, &H3C, &H88)
 SENSOR_EVENT_DATA_UPDATED = iid
End Function
Public Function SENSOR_EVENT_PROPERTY_CHANGED() As UUID
'{2358F099-84C9-4D3D-90DF-C2421E2B2045}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2358F099, CInt(&H84C9), CInt(&H4D3D), &H90, &HDF, &HC2, &H42, &H1E, &H2B, &H20, &H45)
 SENSOR_EVENT_PROPERTY_CHANGED = iid
End Function
Public Function SENSOR_EVENT_ACCELEROMETER_SHAKE() As UUID
'{825F5A94-0F48-4396-9CA0-6ECB5C99D915}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H825F5A94, CInt(&HF48), CInt(&H4396), &H9C, &HA0, &H6E, &HCB, &H5C, &H99, &HD9, &H15)
 SENSOR_EVENT_ACCELEROMETER_SHAKE = iid
End Function
Public Function SENSOR_EVENT_PARAMETER_COMMON_GUID() As UUID
'{64346E30-8728-4B34-BDF6-4F52442C5C28}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H64346E30, CInt(&H8728), CInt(&H4B34), &HBD, &HF6, &H4F, &H52, &H44, &H2C, &H5C, &H28)
 SENSOR_EVENT_PARAMETER_COMMON_GUID = iid
End Function
Public Function SENSOR_PROPERTY_COMMON_GUID() As UUID
'{7F8383EC-D3EC-495C-A8CF-B8BBE85C2920}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7F8383EC, CInt(&HD3EC), CInt(&H495C), &HA8, &HCF, &HB8, &HBB, &HE8, &H5C, &H29, &H20)
 SENSOR_PROPERTY_COMMON_GUID = iid
End Function
Public Function SENSOR_ERROR_PARAMETER_COMMON_GUID() As UUID
'{77112BCD-FCE1-4f43-B8B8-A88256ADB4B3}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H77112BCD, CInt(&HFCE1), CInt(&H4F43), &HB8, &HB8, &HA8, &H82, &H56, &HAD, &HB4, &HB3)
 SENSOR_ERROR_PARAMETER_COMMON_GUID = iid
End Function
Public Function SENSOR_CATEGORY_ALL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC317C286, &HC468, &H4288, &H99, &H75, &HD4, &HC4, &H58, &H7C, &H44, &H2C)
SENSOR_CATEGORY_ALL = iid
End Function
Public Function SENSOR_CATEGORY_LOCATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBFA794E4, &HF964, &H4FDB, &H90, &HF6, &H51, &H5, &H6B, &HFE, &H4B, &H44)
SENSOR_CATEGORY_LOCATION = iid
End Function
Public Function SENSOR_CATEGORY_ENVIRONMENTAL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H323439AA, &H7F66, &H492B, &HBA, &HC, &H73, &HE9, &HAA, &HA, &H65, &HD5)
SENSOR_CATEGORY_ENVIRONMENTAL = iid
End Function
Public Function SENSOR_CATEGORY_MOTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCD09DAF1, &H3B2E, &H4C3D, &HB5, &H98, &HB5, &HE5, &HFF, &H93, &HFD, &H46)
SENSOR_CATEGORY_MOTION = iid
End Function
Public Function SENSOR_CATEGORY_ORIENTATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9E6C04B6, &H96FE, &H4954, &HB7, &H26, &H68, &H68, &H2A, &H47, &H3F, &H69)
SENSOR_CATEGORY_ORIENTATION = iid
End Function
Public Function SENSOR_CATEGORY_MECHANICAL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8D131D68, &H8EF7, &H4656, &H80, &HB5, &HCC, &HCB, &HD9, &H37, &H91, &HC5)
SENSOR_CATEGORY_MECHANICAL = iid
End Function
Public Function SENSOR_CATEGORY_ELECTRICAL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFB73FCD8, &HFC4A, &H483C, &HAC, &H58, &H27, &HB6, &H91, &HC6, &HBE, &HFF)
SENSOR_CATEGORY_ELECTRICAL = iid
End Function
Public Function SENSOR_CATEGORY_BIOMETRIC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCA19690F, &HA2C7, &H477D, &HA9, &H9E, &H99, &HEC, &H6E, &H2B, &H56, &H48)
SENSOR_CATEGORY_BIOMETRIC = iid
End Function
Public Function SENSOR_CATEGORY_LIGHT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H17A665C0, &H9063, &H4216, &HB2, &H2, &H5C, &H7A, &H25, &H5E, &H18, &HCE)
SENSOR_CATEGORY_LIGHT = iid
End Function
Public Function SENSOR_CATEGORY_SCANNER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB000E77E, &HF5B5, &H420F, &H81, &H5D, &H2, &H70, &HA7, &H26, &HF2, &H70)
SENSOR_CATEGORY_SCANNER = iid
End Function
Public Function SENSOR_CATEGORY_OTHER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2C90E7A9, &HF4C9, &H4FA2, &HAF, &H37, &H56, &HD4, &H71, &HFE, &H5A, &H3D)
SENSOR_CATEGORY_OTHER = iid
End Function
Public Function SENSOR_CATEGORY_UNSUPPORTED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2BEAE7FA, &H19B0, &H48C5, &HA1, &HF6, &HB5, &H48, &HD, &HC2, &H6, &HB0)
SENSOR_CATEGORY_UNSUPPORTED = iid
End Function
Public Function SENSOR_TYPE_LOCATION_GPS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HED4CA589, &H327A, &H4FF9, &HA5, &H60, &H91, &HDA, &H4B, &H48, &H27, &H5E)
SENSOR_TYPE_LOCATION_GPS = iid
End Function
Public Function SENSOR_TYPE_LOCATION_STATIC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H95F8184, &HFA9, &H4445, &H8E, &H6E, &HB7, &HF, &H32, &HB, &H6B, &H4C)
SENSOR_TYPE_LOCATION_STATIC = iid
End Function
Public Function SENSOR_TYPE_LOCATION_LOOKUP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3B2EAE4A, &H72CE, &H436D, &H96, &HD2, &H3C, &H5B, &H85, &H70, &HE9, &H87)
SENSOR_TYPE_LOCATION_LOOKUP = iid
End Function
Public Function SENSOR_TYPE_LOCATION_TRIANGULATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H691C341A, &H5406, &H4FE1, &H94, &H2F, &H22, &H46, &HCB, &HEB, &H39, &HE0)
SENSOR_TYPE_LOCATION_TRIANGULATION = iid
End Function
Public Function SENSOR_TYPE_LOCATION_OTHER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9B2D0566, &H368, &H4F71, &HB8, &H8D, &H53, &H3F, &H13, &H20, &H31, &HDE)
SENSOR_TYPE_LOCATION_OTHER = iid
End Function
Public Function SENSOR_TYPE_LOCATION_BROADCAST() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD26988CF, &H5162, &H4039, &HBB, &H17, &H4C, &H58, &HB6, &H98, &HE4, &H4A)
SENSOR_TYPE_LOCATION_BROADCAST = iid
End Function
Public Function SENSOR_TYPE_LOCATION_DEAD_RECKONING() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1A37D538, &HF28B, &H42DA, &H9F, &HCE, &HA9, &HD0, &HA2, &HA6, &HD8, &H29)
SENSOR_TYPE_LOCATION_DEAD_RECKONING = iid
End Function
Public Function SENSOR_TYPE_ENVIRONMENTAL_TEMPERATURE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4FD0EC4, &HD5DA, &H45FA, &H95, &HA9, &H5D, &HB3, &H8E, &HE1, &H93, &H6)
SENSOR_TYPE_ENVIRONMENTAL_TEMPERATURE = iid
End Function
Public Function SENSOR_TYPE_ENVIRONMENTAL_ATMOSPHERIC_PRESSURE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE903829, &HFF8A, &H4A93, &H97, &HDF, &H3D, &HCB, &HDE, &H40, &H22, &H88)
SENSOR_TYPE_ENVIRONMENTAL_ATMOSPHERIC_PRESSURE = iid
End Function
Public Function SENSOR_TYPE_ENVIRONMENTAL_HUMIDITY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5C72BF67, &HBD7E, &H4257, &H99, &HB, &H98, &HA3, &HBA, &H3B, &H40, &HA)
SENSOR_TYPE_ENVIRONMENTAL_HUMIDITY = iid
End Function
Public Function SENSOR_TYPE_ENVIRONMENTAL_WIND_SPEED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDD50607B, &HA45F, &H42CD, &H8E, &HFD, &HEC, &H61, &H76, &H1C, &H42, &H26)
SENSOR_TYPE_ENVIRONMENTAL_WIND_SPEED = iid
End Function
Public Function SENSOR_TYPE_ENVIRONMENTAL_WIND_DIRECTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9EF57A35, &H9306, &H434D, &HAF, &H9, &H37, &HFA, &H5A, &H9C, &H0, &HBD)
SENSOR_TYPE_ENVIRONMENTAL_WIND_DIRECTION = iid
End Function
Public Function SENSOR_TYPE_ACCELEROMETER_1D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC04D2387, &H7340, &H4CC2, &H99, &H1E, &H3B, &H18, &HCB, &H8E, &HF2, &HF4)
SENSOR_TYPE_ACCELEROMETER_1D = iid
End Function
Public Function SENSOR_TYPE_ACCELEROMETER_2D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB2C517A8, &HF6B5, &H4BA6, &HA4, &H23, &H5D, &HF5, &H60, &HB4, &HCC, &H7)
SENSOR_TYPE_ACCELEROMETER_2D = iid
End Function
Public Function SENSOR_TYPE_ACCELEROMETER_3D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC2FB0F5F, &HE2D2, &H4C78, &HBC, &HD0, &H35, &H2A, &H95, &H82, &H81, &H9D)
SENSOR_TYPE_ACCELEROMETER_3D = iid
End Function
Public Function SENSOR_TYPE_MOTION_DETECTOR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5C7C1A12, &H30A5, &H43B9, &HA4, &HB2, &HCF, &H9, &HEC, &H5B, &H7B, &HE8)
SENSOR_TYPE_MOTION_DETECTOR = iid
End Function
Public Function SENSOR_TYPE_GYROMETER_1D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFA088734, &HF552, &H4584, &H83, &H24, &HED, &HFA, &HF6, &H49, &H65, &H2C)
SENSOR_TYPE_GYROMETER_1D = iid
End Function
Public Function SENSOR_TYPE_GYROMETER_2D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H31EF4F83, &H919B, &H48BF, &H8D, &HE0, &H5D, &H7A, &H9D, &H24, &H5, &H56)
SENSOR_TYPE_GYROMETER_2D = iid
End Function
Public Function SENSOR_TYPE_GYROMETER_3D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9485F5A, &H759E, &H42C2, &HBD, &H4B, &HA3, &H49, &HB7, &H5C, &H86, &H43)
SENSOR_TYPE_GYROMETER_3D = iid
End Function
Public Function SENSOR_TYPE_SPEEDOMETER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6BD73C1F, &HBB4, &H4310, &H81, &HB2, &HDF, &HC1, &H8A, &H52, &HBF, &H94)
SENSOR_TYPE_SPEEDOMETER = iid
End Function
Public Function SENSOR_TYPE_COMPASS_1D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA415F6C5, &HCB50, &H49D0, &H8E, &H62, &HA8, &H27, &HB, &HD7, &HA2, &H6C)
SENSOR_TYPE_COMPASS_1D = iid
End Function
Public Function SENSOR_TYPE_COMPASS_2D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H15655CC0, &H997A, &H4D30, &H84, &HDB, &H57, &HCA, &HBA, &H36, &H48, &HBB)
SENSOR_TYPE_COMPASS_2D = iid
End Function
Public Function SENSOR_TYPE_COMPASS_3D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H76B5CE0D, &H17DD, &H414D, &H93, &HA1, &HE1, &H27, &HF4, &HB, &HDF, &H6E)
SENSOR_TYPE_COMPASS_3D = iid
End Function
Public Function SENSOR_TYPE_INCLINOMETER_1D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB96F98C5, &H7A75, &H4BA7, &H94, &HE9, &HAC, &H86, &H8C, &H96, &H6D, &HD8)
SENSOR_TYPE_INCLINOMETER_1D = iid
End Function
Public Function SENSOR_TYPE_INCLINOMETER_2D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAB140F6D, &H83EB, &H4264, &HB7, &HB, &HB1, &H6A, &H5B, &H25, &H6A, &H1)
SENSOR_TYPE_INCLINOMETER_2D = iid
End Function
Public Function SENSOR_TYPE_INCLINOMETER_3D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB84919FB, &HEA85, &H4976, &H84, &H44, &H6F, &H6F, &H5C, &H6D, &H31, &HDB)
SENSOR_TYPE_INCLINOMETER_3D = iid
End Function
Public Function SENSOR_TYPE_DISTANCE_1D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5F14AB2F, &H1407, &H4306, &HA9, &H3F, &HB1, &HDB, &HAB, &HE4, &HF9, &HC0)
SENSOR_TYPE_DISTANCE_1D = iid
End Function
Public Function SENSOR_TYPE_DISTANCE_2D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5CF9A46C, &HA9A2, &H4E55, &HB6, &HA1, &HA0, &H4A, &HAF, &HA9, &H5A, &H92)
SENSOR_TYPE_DISTANCE_2D = iid
End Function
Public Function SENSOR_TYPE_DISTANCE_3D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA20CAE31, &HE25, &H4772, &H9F, &HE5, &H96, &H60, &H8A, &H13, &H54, &HB2)
SENSOR_TYPE_DISTANCE_3D = iid
End Function
Public Function SENSOR_TYPE_AGGREGATED_QUADRANT_ORIENTATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9F81F1AF, &HC4AB, &H4307, &H99, &H4, &HC8, &H28, &HBF, &HB9, &H8, &H29)
SENSOR_TYPE_AGGREGATED_QUADRANT_ORIENTATION = iid
End Function
Public Function SENSOR_TYPE_AGGREGATED_DEVICE_ORIENTATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCDB5D8F7, &H3CFD, &H41C8, &H85, &H42, &HCC, &HE6, &H22, &HCF, &H5D, &H6E)
SENSOR_TYPE_AGGREGATED_DEVICE_ORIENTATION = iid
End Function
Public Function SENSOR_TYPE_AGGREGATED_SIMPLE_DEVICE_ORIENTATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H86A19291, &H482, &H402C, &HBF, &H4C, &HAD, &HDA, &HC5, &H2B, &H1C, &H39)
SENSOR_TYPE_AGGREGATED_SIMPLE_DEVICE_ORIENTATION = iid
End Function
Public Function SENSOR_TYPE_VOLTAGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC5484637, &H4FB7, &H4953, &H98, &HB8, &HA5, &H6D, &H8A, &HA1, &HFB, &H1E)
SENSOR_TYPE_VOLTAGE = iid
End Function
Public Function SENSOR_TYPE_CURRENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5ADC9FCE, &H15A0, &H4BBE, &HA1, &HAD, &H2D, &H38, &HA9, &HAE, &H83, &H1C)
SENSOR_TYPE_CURRENT = iid
End Function
Public Function SENSOR_TYPE_CAPACITANCE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCA2FFB1C, &H2317, &H49C0, &HA0, &HB4, &HB6, &H3C, &HE6, &H34, &H61, &HA0)
SENSOR_TYPE_CAPACITANCE = iid
End Function
Public Function SENSOR_TYPE_RESISTANCE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9993D2C8, &HC157, &H4A52, &HA7, &HB5, &H19, &H5C, &H76, &H3, &H72, &H31)
SENSOR_TYPE_RESISTANCE = iid
End Function
Public Function SENSOR_TYPE_INDUCTANCE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDC1D933F, &HC435, &H4C7D, &HA2, &HFE, &H60, &H71, &H92, &HA5, &H24, &HD3)
SENSOR_TYPE_INDUCTANCE = iid
End Function
Public Function SENSOR_TYPE_ELECTRICAL_POWER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H212F10F5, &H14AB, &H4376, &H9A, &H43, &HA7, &H79, &H40, &H98, &HC2, &HFE)
SENSOR_TYPE_ELECTRICAL_POWER = iid
End Function
Public Function SENSOR_TYPE_POTENTIOMETER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2B3681A9, &HCADC, &H45AA, &HA6, &HFF, &H54, &H95, &H7C, &H8B, &HB4, &H40)
SENSOR_TYPE_POTENTIOMETER = iid
End Function
Public Function SENSOR_TYPE_FREQUENCY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8CD2CBB6, &H73E6, &H4640, &HA7, &H9, &H72, &HAE, &H8F, &HB6, &HD, &H7F)
SENSOR_TYPE_FREQUENCY = iid
End Function
Public Function SENSOR_TYPE_BOOLEAN_SWITCH() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9C7E371F, &H1041, &H460B, &H8D, &H5C, &H71, &HE4, &H75, &H2E, &H35, &HC)
SENSOR_TYPE_BOOLEAN_SWITCH = iid
End Function
Public Function SENSOR_TYPE_MULTIVALUE_SWITCH() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB3EE4D76, &H37A4, &H4402, &HB2, &H5E, &H99, &HC6, &HA, &H77, &H5F, &HA1)
SENSOR_TYPE_MULTIVALUE_SWITCH = iid
End Function
Public Function SENSOR_TYPE_FORCE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC2AB2B02, &H1A1C, &H4778, &HA8, &H1B, &H95, &H4A, &H17, &H88, &HCC, &H75)
SENSOR_TYPE_FORCE = iid
End Function
Public Function SENSOR_TYPE_SCALE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC06DD92C, &H7FEB, &H438E, &H9B, &HF6, &H82, &H20, &H7F, &HFF, &H5B, &HB8)
SENSOR_TYPE_SCALE = iid
End Function
Public Function SENSOR_TYPE_PRESSURE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H26D31F34, &H6352, &H41CF, &HB7, &H93, &HEA, &H7, &H13, &HD5, &H3D, &H77)
SENSOR_TYPE_PRESSURE = iid
End Function
Public Function SENSOR_TYPE_STRAIN() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC6D1EC0E, &H6803, &H4361, &HAD, &H3D, &H85, &HBC, &HC5, &H8C, &H6D, &H29)
SENSOR_TYPE_STRAIN = iid
End Function
Public Function SENSOR_TYPE_BOOLEAN_SWITCH_ARRAY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H545C8BA5, &HB143, &H4545, &H86, &H8F, &HCA, &H7F, &HD9, &H86, &HB4, &HF6)
SENSOR_TYPE_BOOLEAN_SWITCH_ARRAY = iid
End Function
Public Function SENSOR_TYPE_HUMAN_PRESENCE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC138C12B, &HAD52, &H451C, &H93, &H75, &H87, &HF5, &H18, &HFF, &H10, &HC6)
SENSOR_TYPE_HUMAN_PRESENCE = iid
End Function
Public Function SENSOR_TYPE_HUMAN_PROXIMITY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5220DAE9, &H3179, &H4430, &H9F, &H90, &H6, &H26, &H6D, &H2A, &H34, &HDE)
SENSOR_TYPE_HUMAN_PROXIMITY = iid
End Function
Public Function SENSOR_TYPE_TOUCH() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H17DB3018, &H6C4, &H4F7D, &H81, &HAF, &H92, &H74, &HB7, &H59, &H9C, &H27)
SENSOR_TYPE_TOUCH = iid
End Function
Public Function SENSOR_TYPE_AMBIENT_LIGHT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H97F115C8, &H599A, &H4153, &H88, &H94, &HD2, &HD1, &H28, &H99, &H91, &H8A)
SENSOR_TYPE_AMBIENT_LIGHT = iid
End Function
Public Function SENSOR_TYPE_RFID_SCANNER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H44328EF5, &H2DD, &H4E8D, &HAD, &H5D, &H92, &H49, &H83, &H2B, &H2E, &HCA)
SENSOR_TYPE_RFID_SCANNER = iid
End Function
Public Function SENSOR_TYPE_BARCODE_SCANNER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H990B3D8F, &H85BB, &H45FF, &H91, &H4D, &H99, &H8C, &H4, &HF3, &H72, &HDF)
SENSOR_TYPE_BARCODE_SCANNER = iid
End Function
Public Function SENSOR_TYPE_CUSTOM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE83AF229, &H8640, &H4D18, &HA2, &H13, &HE2, &H26, &H75, &HEB, &HB2, &HC3)
SENSOR_TYPE_CUSTOM = iid
End Function
Public Function SENSOR_TYPE_UNKNOWN() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H10BA83E3, &HEF4F, &H41ED, &H98, &H85, &HA8, &H7D, &H64, &H35, &HA8, &HE1)
SENSOR_TYPE_UNKNOWN = iid
End Function
Public Function SENSOR_DATA_TYPE_COMMON_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDB5E0CF2, &HCF1F, &H4C18, &HB4, &H6C, &HD8, &H60, &H11, &HD6, &H21, &H50)
SENSOR_DATA_TYPE_COMMON_GUID = iid
End Function
Public Function SENSOR_DATA_TYPE_LOCATION_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H55C74D8, &HCA6F, &H47D6, &H95, &HC6, &H1E, &HD3, &H63, &H7A, &HF, &HF4)
SENSOR_DATA_TYPE_LOCATION_GUID = iid
End Function
Public Function SENSOR_DATA_TYPE_ENVIRONMENTAL_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8B0AA2F1, &H2D57, &H42EE, &H8C, &HC0, &H4D, &H27, &H62, &H2B, &H46, &HC4)
SENSOR_DATA_TYPE_ENVIRONMENTAL_GUID = iid
End Function
Public Function SENSOR_DATA_TYPE_MOTION_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3F8A69A2, &H7C5, &H4E48, &HA9, &H65, &HCD, &H79, &H7A, &HAB, &H56, &HD5)
SENSOR_DATA_TYPE_MOTION_GUID = iid
End Function
Public Function SENSOR_DATA_TYPE_ORIENTATION_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1637D8A2, &H4248, &H4275, &H86, &H5D, &H55, &H8D, &HE8, &H4A, &HED, &HFD)
SENSOR_DATA_TYPE_ORIENTATION_GUID = iid
End Function
Public Function SENSOR_DATA_TYPE_GUID_MECHANICAL_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H38564A7C, &HF2F2, &H49BB, &H9B, &H2B, &HBA, &H60, &HF6, &H6A, &H58, &HDF)
SENSOR_DATA_TYPE_GUID_MECHANICAL_GUID = iid
End Function
Public Function SENSOR_DATA_TYPE_BIOMETRIC_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2299288A, &H6D9E, &H4B0B, &HB7, &HEC, &H35, &H28, &HF8, &H9E, &H40, &HAF)
SENSOR_DATA_TYPE_BIOMETRIC_GUID = iid
End Function
Public Function SENSOR_DATA_TYPE_LIGHT_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE4C77CE2, &HDCB7, &H46E9, &H84, &H39, &H4F, &HEC, &H54, &H88, &H33, &HA6)
SENSOR_DATA_TYPE_LIGHT_GUID = iid
End Function
Public Function SENSOR_DATA_TYPE_SCANNER_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD7A59A3C, &H3421, &H44AB, &H8D, &H3A, &H9D, &HE8, &HAB, &H6C, &H4C, &HAE)
SENSOR_DATA_TYPE_SCANNER_GUID = iid
End Function
Public Function SENSOR_DATA_TYPE_ELECTRICAL_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBBB246D1, &HE242, &H4780, &HA2, &HD3, &HCD, &HED, &H84, &HF3, &H58, &H42)
SENSOR_DATA_TYPE_ELECTRICAL_GUID = iid
End Function
Public Function SENSOR_DATA_TYPE_CUSTOM_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB14C764F, &H7CF, &H41E8, &H9D, &H82, &HEB, &HE3, &HD0, &H77, &H6A, &H6F)
SENSOR_DATA_TYPE_CUSTOM_GUID = iid
End Function
Public Function SENSOR_PROPERTY_TEST_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE1E962F4, &H6E65, &H45F7, &H9C, &H36, &HD4, &H87, &HB7, &HB1, &HBD, &H34)
SENSOR_PROPERTY_TEST_GUID = iid
End Function



Public Function SID_STopLevelBrowser() As UUID
'{4C96BE40-915C-11CF-99D3-00AA004AE837}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4C96BE40, CInt(&H915C), CInt(&H11CF), &H99, &HD3, &H0, &HAA, &H0, &H4A, &HE8, &H37)
 SID_STopLevelBrowser = iid
End Function
Public Function SID_SExplorerBrowserFrame() As UUID
SID_SExplorerBrowserFrame = IID_ICommDlgBrowser
End Function
Public Function SID_SFolderView() As UUID
SID_SFolderView = IID_IFolderView
End Function
Public Function SID_SProfferService() As UUID
SID_SProfferService = IID_IProfferService
End Function
Public Function SID_WizardHost() As UUID
SID_WizardHost = IID_IWebWizardExtension
End Function
Public Function SID_CDWizardHost() As UUID
SID_CDWizardHost = IID_ICDBurnExt
End Function
Public Function SID_SBandSite() As UUID
SID_SBandSite = IID_IBandSite
End Function
Public Function SID_SNewMenuClient() As UUID
SID_SNewMenuClient = IID_INewMenuClient
End Function
Public Function SID_SNewWindowManager() As UUID
SID_SNewWindowManager = IID_INewWindowManager
End Function
Public Function SID_ExecuteCommandHost() As UUID
SID_ExecuteCommandHost = IID_IExecuteCommandHost
End Function
Public Function SID_SHandlerActivationHost() As UUID
SID_SHandlerActivationHost = IID_IHandlerActivationHost
End Function
Public Function SID_HandlerInfo() As UUID
SID_HandlerInfo = IID_IHandlerInfo
End Function
Public Function SID_LaunchSourceAppUserModelId() As UUID
'{2CE78010-74DB-48BC-9C6A-10F372495723}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2CE78010, CInt(&H74DB), CInt(&H48BC), &H9C, &H6A, &H10, &HF3, &H72, &H49, &H57, &H23)
 SID_LaunchSourceAppUserModelId = iid
End Function
Public Function SID_ShellExecuteNamedPropertyStore() As UUID
Static iid As UUID
If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEB84ADA2, CInt(&HFF), CInt(&H4992), &H83, &H24, &HED, &H5C, &HE0, &H61, &HCB, &H29)
SID_ShellExecuteNamedPropertyStore = iid
End Function
Public Function SID_LaunchSourceViewSizePreference() As UUID
'{80605492-67D9-414F-AF89-A1CDF1242BC1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H80605492, CInt(&H67D9), CInt(&H414F), &HAF, &H89, &HA1, &HCD, &HF1, &H24, &H2B, &HC1)
 SID_LaunchSourceViewSizePreference = iid
End Function
Public Function SID_LaunchTargetViewSizePreference() As UUID
'{26DB2472-B7B7-406B-9702-730A4E20D3BF}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H26DB2472, CInt(&HB7B7), CInt(&H406B), &H97, &H2, &H73, &HA, &H4E, &H20, &HD3, &HBF)
 SID_LaunchTargetViewSizePreference = iid
End Function
Public Function SID_LaunchTargetMonitor() As UUID
'{8D547FA1-CC45-40C8-B7C1-D48C183F13F3}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8D547FA1, CInt(&HCC45), CInt(&H40C8), &HB7, &HC1, &HD4, &H8C, &H18, &H3F, &H13, &HF3)
 SID_LaunchTargetMonitor = iid
End Function
Public Function SID_GetScriptSite() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDB01A1E3, &HA42B, &H11CF, &H8F, &H20, &H0, &H80, &H5F, &H2C, &HD0, &H64)
SID_GetScriptSite = iid
End Function
Public Function SID_VariantConversion() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1F101481, &HBCCD, &H11D0, &H93, &H36, &H0, &HA0, &HC9, &HD, &HCA, &HA9)
SID_VariantConversion = iid
End Function
Public Function SID_GetCaller() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4717CC40, &HBCB9, &H11D0, &H93, &H36, &H0, &HA0, &HC9, &HD, &HCA, &HA9)
SID_GetCaller = iid
End Function
Public Function SID_ProvideRuntimeContext() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H74A5040C, &HDD0C, &H48F0, &HAC, &H85, &H19, &H4C, &H32, &H59, &H18, &HA)
SID_ProvideRuntimeContext = iid
End Function
Public Function SID_EnumerableView() As UUID
SID_EnumerableView = IID_IEnumerableView
End Function
Public Function SID_SCommandBarState() As UUID
'{B99EAA5C-3850-4400-BC33-2CE534048BF8}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB99EAA5C, CInt(&H3850), CInt(&H4400), &HBC, &H33, &H2C, &HE5, &H34, &H4, &H8B, &HF8)
 SID_SCommandBarState = iid
End Function
Public Function SID_SBandHost() As UUID
SID_SBandHost = IID_IBandHost
End Function
Public Function SID_ExplorerPaneVisibility() As UUID
SID_ExplorerPaneVisibility = IID_IExplorerPaneVisibility
End Function
Public Function SID_SOleUndoManager() As UUID
SID_SOleUndoManager = IID_IOleUndoManager
End Function

Public Function FOLDERID_NetworkFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD20BEEC4, CInt(&H5CA8), CInt(&H4905), &HAE, &H3B, &HBF, &H25, &H1E, &HA0, &H9B, &H53)
 FOLDERID_NetworkFolder = iid
End Function

Public Function FOLDERID_ComputerFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAC0837C, CInt(&HBBF8), CInt(&H452A), &H85, &HD, &H79, &HD0, &H8E, &H66, &H7C, &HA7)
 FOLDERID_ComputerFolder = iid
End Function

Public Function FOLDERID_InternetFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D9F7874, CInt(&H4E0C), CInt(&H4904), &H96, &H7B, &H40, &HB0, &HD2, &HC, &H3E, &H4B)
 FOLDERID_InternetFolder = iid
End Function

Public Function FOLDERID_ControlPanelFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H82A74AEB, CInt(&HAEB4), CInt(&H465C), &HA0, &H14, &HD0, &H97, &HEE, &H34, &H6D, &H63)
 FOLDERID_ControlPanelFolder = iid
End Function

Public Function FOLDERID_PrintersFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H76FC4E2D, CInt(&HD6AD), CInt(&H4519), &HA6, &H63, &H37, &HBD, &H56, &H6, &H81, &H85)
 FOLDERID_PrintersFolder = iid
End Function

Public Function FOLDERID_SyncManagerFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43668BF8, CInt(&HC14E), CInt(&H49B2), &H97, &HC9, &H74, &H77, &H84, &HD7, &H84, &HB7)
 FOLDERID_SyncManagerFolder = iid
End Function

Public Function FOLDERID_SyncSetupFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF214138, CInt(&HB1D3), CInt(&H4A90), &HBB, &HA9, &H27, &HCB, &HC0, &HC5, &H38, &H9A)
 FOLDERID_SyncSetupFolder = iid
End Function

Public Function FOLDERID_ConflictFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4BFEFB45, CInt(&H347D), CInt(&H4006), &HA5, &HBE, &HAC, &HC, &HB0, &H56, &H71, &H92)
 FOLDERID_ConflictFolder = iid
End Function

Public Function FOLDERID_SyncResultsFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H289A9A43, CInt(&HBE44), CInt(&H4057), &HA4, &H1B, &H58, &H7A, &H76, &HD7, &HE7, &HF9)
 FOLDERID_SyncResultsFolder = iid
End Function

Public Function FOLDERID_RecycleBinFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB7534046, CInt(&H3ECB), CInt(&H4C18), &HBE, &H4E, &H64, &HCD, &H4C, &HB7, &HD6, &HAC)
 FOLDERID_RecycleBinFolder = iid
End Function

Public Function FOLDERID_ConnectionsFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6F0CD92B, CInt(&H2E97), CInt(&H45D1), &H88, &HFF, &HB0, &HD1, &H86, &HB8, &HDE, &HDD)
 FOLDERID_ConnectionsFolder = iid
End Function

Public Function FOLDERID_Fonts() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFD228CB7, CInt(&HAE11), CInt(&H4AE3), &H86, &H4C, &H16, &HF3, &H91, &HA, &HB8, &HFE)
 FOLDERID_Fonts = iid
End Function

Public Function FOLDERID_Desktop() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB4BFCC3A, CInt(&HDB2C), CInt(&H424C), &HB0, &H29, &H7F, &HE9, &H9A, &H87, &HC6, &H41)
 FOLDERID_Desktop = iid
End Function

Public Function FOLDERID_Startup() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB97D20BB, CInt(&HF46A), CInt(&H4C97), &HBA, &H10, &H5E, &H36, &H8, &H43, &H8, &H54)
 FOLDERID_Startup = iid
End Function

Public Function FOLDERID_Programs() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA77F5D77, CInt(&H2E2B), CInt(&H44C3), &HA6, &HA2, &HAB, &HA6, &H1, &H5, &H4A, &H51)
 FOLDERID_Programs = iid
End Function

Public Function FOLDERID_StartMenu() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H625B53C3, CInt(&HAB48), CInt(&H4EC1), &HBA, &H1F, &HA1, &HEF, &H41, &H46, &HFC, &H19)
 FOLDERID_StartMenu = iid
End Function

Public Function FOLDERID_Recent() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAE50C081, CInt(&HEBD2), CInt(&H438A), &H86, &H55, &H8A, &H9, &H2E, &H34, &H98, &H7A)
 FOLDERID_Recent = iid
End Function

Public Function FOLDERID_SendTo() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8983036C, CInt(&H27C0), CInt(&H404B), &H8F, &H8, &H10, &H2D, &H10, &HDC, &HFD, &H74)
 FOLDERID_SendTo = iid
End Function

Public Function FOLDERID_Documents() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFDD39AD0, CInt(&H238F), CInt(&H46AF), &HAD, &HB4, &H6C, &H85, &H48, &H3, &H69, &HC7)
 FOLDERID_Documents = iid
End Function

Public Function FOLDERID_Favorites() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1777F761, CInt(&H68AD), CInt(&H4D8A), &H87, &HBD, &H30, &HB7, &H59, &HFA, &H33, &HDD)
 FOLDERID_Favorites = iid
End Function

Public Function FOLDERID_NetHood() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC5ABBF53, CInt(&HE17F), CInt(&H4121), &H89, &H0, &H86, &H62, &H6F, &HC2, &HC9, &H73)
 FOLDERID_NetHood = iid
End Function

Public Function FOLDERID_PrintHood() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9274BD8D, CInt(&HCFD1), CInt(&H41C3), &HB3, &H5E, &HB1, &H3F, &H55, &HA7, &H58, &HF4)
 FOLDERID_PrintHood = iid
End Function

Public Function FOLDERID_Templates() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA63293E8, CInt(&H664E), CInt(&H48DB), &HA0, &H79, &HDF, &H75, &H9E, &H5, &H9, &HF7)
 FOLDERID_Templates = iid
End Function

Public Function FOLDERID_CommonStartup() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H82A5EA35, CInt(&HD9CD), CInt(&H47C5), &H96, &H29, &HE1, &H5D, &H2F, &H71, &H4E, &H6E)
 FOLDERID_CommonStartup = iid
End Function

Public Function FOLDERID_CommonPrograms() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H139D44E, CInt(&H6AFE), CInt(&H49F2), &H86, &H90, &H3D, &HAF, &HCA, &HE6, &HFF, &HB8)
 FOLDERID_CommonPrograms = iid
End Function

Public Function FOLDERID_CommonStartMenu() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA4115719, CInt(&HD62E), CInt(&H491D), &HAA, &H7C, &HE7, &H4B, &H8B, &HE3, &HB0, &H67)
 FOLDERID_CommonStartMenu = iid
End Function

Public Function FOLDERID_PublicDesktop() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC4AA340D, CInt(&HF20F), CInt(&H4863), &HAF, &HEF, &HF8, &H7E, &HF2, &HE6, &HBA, &H25)
 FOLDERID_PublicDesktop = iid
End Function

Public Function FOLDERID_ProgramData() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H62AB5D82, CInt(&HFDC1), CInt(&H4DC3), &HA9, &HDD, &H7, &HD, &H1D, &H49, &H5D, &H97)
 FOLDERID_ProgramData = iid
End Function

Public Function FOLDERID_CommonTemplates() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB94237E7, CInt(&H57AC), CInt(&H4347), &H91, &H51, &HB0, &H8C, &H6C, &H32, &HD1, &HF7)
 FOLDERID_CommonTemplates = iid
End Function

Public Function FOLDERID_PublicDocuments() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HED4824AF, CInt(&HDCE4), CInt(&H45A8), &H81, &HE2, &HFC, &H79, &H65, &H8, &H36, &H34)
 FOLDERID_PublicDocuments = iid
End Function

Public Function FOLDERID_RoamingAppData() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3EB685DB, CInt(&H65F9), CInt(&H4CF6), &HA0, &H3A, &HE3, &HEF, &H65, &H72, &H9F, &H3D)
 FOLDERID_RoamingAppData = iid
End Function

Public Function FOLDERID_LocalAppData() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF1B32785, CInt(&H6FBA), CInt(&H4FCF), &H9D, &H55, &H7B, &H8E, &H7F, &H15, &H70, &H91)
 FOLDERID_LocalAppData = iid
End Function

Public Function FOLDERID_LocalAppDataLow() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA520A1A4, CInt(&H1780), CInt(&H4FF6), &HBD, &H18, &H16, &H73, &H43, &HC5, &HAF, &H16)
 FOLDERID_LocalAppDataLow = iid
End Function

Public Function FOLDERID_InternetCache() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H352481E8, CInt(&H33BE), CInt(&H4251), &HBA, &H85, &H60, &H7, &HCA, &HED, &HCF, &H9D)
 FOLDERID_InternetCache = iid
End Function

Public Function FOLDERID_Cookies() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2B0F765D, CInt(&HC0E9), CInt(&H4171), &H90, &H8E, &H8, &HA6, &H11, &HB8, &H4F, &HF6)
 FOLDERID_Cookies = iid
End Function

Public Function FOLDERID_History() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD9DC8A3B, CInt(&HB784), CInt(&H432E), &HA7, &H81, &H5A, &H11, &H30, &HA7, &H59, &H63)
 FOLDERID_History = iid
End Function

Public Function FOLDERID_System() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1AC14E77, CInt(&H2E7), CInt(&H4E5D), &HB7, &H44, &H2E, &HB1, &HAE, &H51, &H98, &HB7)
 FOLDERID_System = iid
End Function

Public Function FOLDERID_SystemX86() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD65231B0, CInt(&HB2F1), CInt(&H4857), &HA4, &HCE, &HA8, &HE7, &HC6, &HEA, &H7D, &H27)
 FOLDERID_SystemX86 = iid
End Function

Public Function FOLDERID_Windows() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF38BF404, CInt(&H1D43), CInt(&H42F2), &H93, &H5, &H67, &HDE, &HB, &H28, &HFC, &H23)
 FOLDERID_Windows = iid
End Function

Public Function FOLDERID_Profile() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5E6C858F, CInt(&HE22), CInt(&H4760), &H9A, &HFE, &HEA, &H33, &H17, &HB6, &H71, &H73)
 FOLDERID_Profile = iid
End Function

Public Function FOLDERID_Pictures() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H33E28130, CInt(&H4E1E), CInt(&H4676), &H83, &H5A, &H98, &H39, &H5C, &H3B, &HC3, &HBB)
 FOLDERID_Pictures = iid
End Function

Public Function FOLDERID_ProgramFilesX86() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7C5A40EF, CInt(&HA0FB), CInt(&H4BFC), &H87, &H4A, &HC0, &HF2, &HE0, &HB9, &HFA, &H8E)
 FOLDERID_ProgramFilesX86 = iid
End Function

Public Function FOLDERID_ProgramFilesCommonX86() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDE974D24, CInt(&HD9C6), CInt(&H4D3E), &HBF, &H91, &HF4, &H45, &H51, &H20, &HB9, &H17)
 FOLDERID_ProgramFilesCommonX86 = iid
End Function

Public Function FOLDERID_ProgramFilesX64() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6D809377, CInt(&H6AF0), CInt(&H444B), &H89, &H57, &HA3, &H77, &H3F, &H2, &H20, &HE)
 FOLDERID_ProgramFilesX64 = iid
End Function

Public Function FOLDERID_ProgramFilesCommonX64() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6365D5A7, CInt(&HF0D), CInt(&H45E5), &H87, &HF6, &HD, &HA5, &H6B, &H6A, &H4F, &H7D)
 FOLDERID_ProgramFilesCommonX64 = iid
End Function

Public Function FOLDERID_ProgramFiles() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H905E63B6, CInt(&HC1BF), CInt(&H494E), &HB2, &H9C, &H65, &HB7, &H32, &HD3, &HD2, &H1A)
 FOLDERID_ProgramFiles = iid
End Function

Public Function FOLDERID_ProgramFilesCommon() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF7F1ED05, CInt(&H9F6D), CInt(&H47A2), &HAA, &HAE, &H29, &HD3, &H17, &HC6, &HF0, &H66)
 FOLDERID_ProgramFilesCommon = iid
End Function

Public Function FOLDERID_AdminTools() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H724EF170, CInt(&HA42D), CInt(&H4FEF), &H9F, &H26, &HB6, &HE, &H84, &H6F, &HBA, &H4F)
 FOLDERID_AdminTools = iid
End Function

Public Function FOLDERID_CommonAdminTools() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD0384E7D, CInt(&HBAC3), CInt(&H4797), &H8F, &H14, &HCB, &HA2, &H29, &HB3, &H92, &HB5)
 FOLDERID_CommonAdminTools = iid
End Function

Public Function FOLDERID_Music() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4BD8D571, CInt(&H6D19), CInt(&H48D3), &HBE, &H97, &H42, &H22, &H20, &H8, &HE, &H43)
 FOLDERID_Music = iid
End Function

Public Function FOLDERID_Videos() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H18989B1D, CInt(&H99B5), CInt(&H455B), &H84, &H1C, &HAB, &H7C, &H74, &HE4, &HDD, &HFC)
 FOLDERID_Videos = iid
End Function

Public Function FOLDERID_PublicPictures() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB6EBFB86, CInt(&H6907), CInt(&H413C), &H9A, &HF7, &H4F, &HC2, &HAB, &HF0, &H7C, &HC5)
 FOLDERID_PublicPictures = iid
End Function

Public Function FOLDERID_PublicMusic() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3214FAB5, CInt(&H9757), CInt(&H4298), &HBB, &H61, &H92, &HA9, &HDE, &HAA, &H44, &HFF)
 FOLDERID_PublicMusic = iid
End Function

Public Function FOLDERID_PublicVideos() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2400183A, CInt(&H6185), CInt(&H49FB), &HA2, &HD8, &H4A, &H39, &H2A, &H60, &H2B, &HA3)
 FOLDERID_PublicVideos = iid
End Function

Public Function FOLDERID_ResourceDir() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8AD10C31, CInt(&H2ADB), CInt(&H4296), &HA8, &HF7, &HE4, &H70, &H12, &H32, &HC9, &H72)
 FOLDERID_ResourceDir = iid
End Function

Public Function FOLDERID_LocalizedResourcesDir() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2A00375E, CInt(&H224C), CInt(&H49DE), &HB8, &HD1, &H44, &HD, &HF7, &HEF, &H3D, &HDC)
 FOLDERID_LocalizedResourcesDir = iid
End Function

Public Function FOLDERID_CommonOEMLinks() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC1BAE2D0, CInt(&H10DF), CInt(&H4334), &HBE, &HDD, &H7A, &HA2, &HB, &H22, &H7A, &H9D)
 FOLDERID_CommonOEMLinks = iid
End Function

Public Function FOLDERID_CDBurning() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9E52AB10, CInt(&HF80D), CInt(&H49DF), &HAC, &HB8, &H43, &H30, &HF5, &H68, &H78, &H55)
 FOLDERID_CDBurning = iid
End Function

Public Function FOLDERID_UserProfiles() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H762D272, CInt(&HC50A), CInt(&H4BB0), &HA3, &H82, &H69, &H7D, &HCD, &H72, &H9B, &H80)
 FOLDERID_UserProfiles = iid
End Function

Public Function FOLDERID_Playlists() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDE92C1C7, CInt(&H837F), CInt(&H4F69), &HA3, &HBB, &H86, &HE6, &H31, &H20, &H4A, &H23)
 FOLDERID_Playlists = iid
End Function

Public Function FOLDERID_SamplePlaylists() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H15CA69B3, CInt(&H30EE), CInt(&H49C1), &HAC, &HE1, &H6B, &H5E, &HC3, &H72, &HAF, &HB5)
 FOLDERID_SamplePlaylists = iid
End Function

Public Function FOLDERID_SampleMusic() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB250C668, CInt(&HF57D), CInt(&H4EE1), &HA6, &H3C, &H29, &HE, &HE7, &HD1, &HAA, &H1F)
 FOLDERID_SampleMusic = iid
End Function

Public Function FOLDERID_SamplePictures() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC4900540, CInt(&H2379), CInt(&H4C75), &H84, &H4B, &H64, &HE6, &HFA, &HF8, &H71, &H6B)
 FOLDERID_SamplePictures = iid
End Function

Public Function FOLDERID_SampleVideos() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H859EAD94, CInt(&H2E85), CInt(&H48AD), &HA7, &H1A, &H9, &H69, &HCB, &H56, &HA6, &HCD)
 FOLDERID_SampleVideos = iid
End Function

Public Function FOLDERID_PhotoAlbums() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H69D2CF90, CInt(&HFC33), CInt(&H4FB7), &H9A, &HC, &HEB, &HB0, &HF0, &HFC, &HB4, &H3C)
 FOLDERID_PhotoAlbums = iid
End Function

Public Function FOLDERID_Public() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDFDF76A2, CInt(&HC82A), CInt(&H4D63), &H90, &H6A, &H56, &H44, &HAC, &H45, &H73, &H85)
 FOLDERID_Public = iid
End Function

Public Function FOLDERID_ChangeRemovePrograms() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDF7266AC, CInt(&H9274), CInt(&H4867), &H8D, &H55, &H3B, &HD6, &H61, &HDE, &H87, &H2D)
 FOLDERID_ChangeRemovePrograms = iid
End Function

Public Function FOLDERID_AppUpdates() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA305CE99, CInt(&HF527), CInt(&H492B), &H8B, &H1A, &H7E, &H76, &HFA, &H98, &HD6, &HE4)
 FOLDERID_AppUpdates = iid
End Function

Public Function FOLDERID_AddNewPrograms() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDE61D971, CInt(&H5EBC), CInt(&H4F02), &HA3, &HA9, &H6C, &H82, &H89, &H5E, &H5C, &H4)
 FOLDERID_AddNewPrograms = iid
End Function

Public Function FOLDERID_Downloads() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H374DE290, CInt(&H123F), CInt(&H4565), &H91, &H64, &H39, &HC4, &H92, &H5E, &H46, &H7B)
 FOLDERID_Downloads = iid
End Function

Public Function FOLDERID_PublicDownloads() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3D644C9B, CInt(&H1FB8), CInt(&H4F30), &H9B, &H45, &HF6, &H70, &H23, &H5F, &H79, &HC0)
 FOLDERID_PublicDownloads = iid
End Function

Public Function FOLDERID_SavedSearches() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7D1D3A04, CInt(&HDEBB), CInt(&H4115), &H95, &HCF, &H2F, &H29, &HDA, &H29, &H20, &HDA)
 FOLDERID_SavedSearches = iid
End Function

Public Function FOLDERID_QuickLaunch() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H52A4F021, CInt(&H7B75), CInt(&H48A9), &H9F, &H6B, &H4B, &H87, &HA2, &H10, &HBC, &H8F)
 FOLDERID_QuickLaunch = iid
End Function

Public Function FOLDERID_Contacts() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H56784854, CInt(&HC6CB), CInt(&H462B), &H81, &H69, &H88, &HE3, &H50, &HAC, &HB8, &H82)
 FOLDERID_Contacts = iid
End Function

Public Function FOLDERID_SidebarParts() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA75D362E, CInt(&H50FC), CInt(&H4FB7), &HAC, &H2C, &HA8, &HBE, &HAA, &H31, &H44, &H93)
 FOLDERID_SidebarParts = iid
End Function

Public Function FOLDERID_SidebarDefaultParts() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7B396E54, CInt(&H9EC5), CInt(&H4300), &HBE, &HA, &H24, &H82, &HEB, &HAE, &H1A, &H26)
 FOLDERID_SidebarDefaultParts = iid
End Function

Public Function FOLDERID_TreeProperties() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5B3749AD, CInt(&HB49F), CInt(&H49C1), &H83, &HEB, &H15, &H37, &HF, &HBD, &H48, &H82)
 FOLDERID_TreeProperties = iid
End Function

Public Function FOLDERID_PublicGameTasks() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDEBF2536, CInt(&HE1A8), CInt(&H4C59), &HB6, &HA2, &H41, &H45, &H86, &H47, &H6A, &HEA)
 FOLDERID_PublicGameTasks = iid
End Function

Public Function FOLDERID_GameTasks() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H54FAE61, CInt(&H4DD8), CInt(&H4787), &H80, &HB6, &H9, &H2, &H20, &HC4, &HB7, &H0)
 FOLDERID_GameTasks = iid
End Function

Public Function FOLDERID_SavedGames() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4C5C32FF, CInt(&HBB9D), CInt(&H43B0), &HB5, &HB4, &H2D, &H72, &HE5, &H4E, &HAA, &HA4)
 FOLDERID_SavedGames = iid
End Function

Public Function FOLDERID_Games() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCAC52C1A, CInt(&HB53D), CInt(&H4EDC), &H92, &HD7, &H6B, &H2E, &H8A, &HC1, &H94, &H34)
 FOLDERID_Games = iid
End Function

Public Function FOLDERID_RecordedTV() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBD85E001, CInt(&H112E), CInt(&H431E), &H98, &H3B, &H7B, &H15, &HAC, &H9, &HFF, &HF1)
 FOLDERID_RecordedTV = iid
End Function

Public Function FOLDERID_SEARCH_MAPI() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H98EC0E18, CInt(&H2098), CInt(&H4D44), &H86, &H44, &H66, &H97, &H93, &H15, &HA2, &H81)
 FOLDERID_SEARCH_MAPI = iid
End Function

Public Function FOLDERID_SEARCH_CSC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEE32E446, CInt(&H31CA), CInt(&H4ABA), &H81, &H4F, &HA5, &HEB, &HD2, &HFD, &H6D, &H5E)
 FOLDERID_SEARCH_CSC = iid
End Function

Public Function FOLDERID_Links() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBFB9D5E0, CInt(&HC6A9), CInt(&H404C), &HB2, &HB2, &HAE, &H6D, &HB6, &HAF, &H49, &H68)
 FOLDERID_Links = iid
End Function

Public Function FOLDERID_UsersFiles() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF3CE0F7C, CInt(&H4901), CInt(&H4ACC), &H86, &H48, &HD5, &HD4, &H4B, &H4, &HEF, &H8F)
 FOLDERID_UsersFiles = iid
End Function

Public Function FOLDERID_SearchHome() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H190337D1, CInt(&HB8CA), CInt(&H4121), &HA6, &H39, &H6D, &H47, &H2D, &H16, &H97, &H2A)
 FOLDERID_SearchHome = iid
End Function

Public Function FOLDERID_OriginalImages() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2C36C0AA, CInt(&H5812), CInt(&H4B87), &HBF, &HD0, &H4C, &HD0, &HDF, &HB1, &H9B, &H39)
 FOLDERID_OriginalImages = iid
End Function

Public Function FOLDERID_HomeGroup() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H52528A6B, CInt(&HB9E3), CInt(&H4ADD), &HB6, &HD, &H58, &H8C, &H2D, &HBA, &H84, &H2D)
 FOLDERID_HomeGroup = iid
End Function
Public Function FOLDERID_AccountPictures() As UUID
'{008ca0b1-55b4-4c56-b8a8-4de4b299d3be}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CA0B1, CInt(&H55B4), CInt(&H4C56), &HB8, &HA8, &H4D, &HE4, &HB2, &H99, &HD3, &HBE)
FOLDERID_AccountPictures = iid
End Function
Public Function FOLDERID_AppDataDesktop() As UUID
'{B2C5E279-7ADD-439F-B28C-C41FE1BBF672}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB2C5E279, CInt(&H7ADD), CInt(&H439F), &HB2, &H8C, &HC4, &H1F, &HE1, &HBB, &HF6, &H72)
FOLDERID_AppDataDesktop = iid
End Function
Public Function FOLDERID_ApplicationShortcuts() As UUID
'{A3918781-E5F2-4890-B3D9-A7E54332328C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA3918781, CInt(&HE5F2), CInt(&H4890), &HB3, &HD9, &HA7, &HE5, &H43, &H32, &H32, &H8C)
FOLDERID_ApplicationShortcuts = iid
End Function
Public Function FOLDERID_AppsFolder() As UUID
'{1e87508d-89c2-42f0-8a7e-645a0f50ca58}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1E87508D, CInt(&H89C2), CInt(&H42F0), &H8A, &H7E, &H64, &H5A, &HF, &H50, &HCA, &H58)
FOLDERID_AppsFolder = iid
End Function
Public Function FOLDERID_CameraRoll() As UUID
'{AB5FB87B-7CE2-4F83-915D-550846C9537B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB5FB87B, CInt(&H7CE2), CInt(&H4F83), &H91, &H5D, &H55, &H8, &H46, &HC9, &H53, &H7B)
FOLDERID_CameraRoll = iid
End Function
Public Function FOLDERID_DeviceMetadataStore() As UUID
'{5CE4A5E9-E4EB-479D-B89F-130C02886155}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5CE4A5E9, CInt(&HE4EB), CInt(&H479D), &HB8, &H9F, &H13, &HC, &H2, &H88, &H61, &H55)
FOLDERID_DeviceMetadataStore = iid
End Function
Public Function FOLDERID_DocumentsLibrary() As UUID
'{7B0DB17D-9CD2-4A93-9733-46CC89022E7C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7B0DB17D, CInt(&H9CD2), CInt(&H4A93), &H97, &H33, &H46, &HCC, &H89, &H2, &H2E, &H7C)
FOLDERID_DocumentsLibrary = iid
End Function
Public Function FOLDERID_HomeGroupCurrentUser() As UUID
'{9B74B6A3-0DFD-4f11-9E78-5F7800F2E772}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9B74B6A3, CInt(&HDFD), CInt(&H4F11), &H9E, &H78, &H5F, &H78, &H0, &HF2, &HE7, &H72)
FOLDERID_HomeGroupCurrentUser = iid
End Function
Public Function FOLDERID_ImplicitAppShortcuts() As UUID
'{BCB5256F-79F6-4CEE-B725-DC34E402FD46}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBCB5256F, CInt(&H79F6), CInt(&H4CEE), &HB7, &H25, &HDC, &H34, &HE4, &H2, &HFD, &H46)
FOLDERID_ImplicitAppShortcuts = iid
End Function
Public Function FOLDERID_Libraries() As UUID
'{1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1B3EA5DC, CInt(&HB587), CInt(&H4786), &HB4, &HEF, &HBD, &H1D, &HC3, &H32, &HAE, &HAE)
FOLDERID_Libraries = iid
End Function
Public Function FOLDERID_MusicLibrary() As UUID
'{2112AB0A-C86A-4FFE-A368-0DE96E47012E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2112AB0A, CInt(&HC86A), CInt(&H4FFE), &HA3, &H68, &HD, &HE9, &H6E, &H47, &H1, &H2E)
FOLDERID_MusicLibrary = iid
End Function
Public Function FOLDERID_Objects3D() As UUID
'{31C0DD25-9439-4F12-BF41-7FF4EDA38722}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H31C0DD25, CInt(&H9439), CInt(&H4F12), &HBF, &H41, &H7F, &HF4, &HED, &HA3, &H87, &H22)
FOLDERID_Objects3D = iid
End Function
Public Function FOLDERID_PicturesLibrary() As UUID
'{A990AE9F-A03B-4E80-94BC-9912D7504104}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA990AE9F, CInt(&HA03B), CInt(&H4E80), &H94, &HBC, &H99, &H12, &HD7, &H50, &H41, &H4)
FOLDERID_PicturesLibrary = iid
End Function
Public Function FOLDERID_PublicLibraries() As UUID
'{48DAF80B-E6CF-4F4E-B800-0E69D84EE384}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H48DAF80B, CInt(&HE6CF), CInt(&H4F4E), &HB8, &H0, &HE, &H69, &HD8, &H4E, &HE3, &H84)
FOLDERID_PublicLibraries = iid
End Function
Public Function FOLDERID_PublicRingtones() As UUID
'{E555AB60-153B-4D17-9F04-A5FE99FC15EC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE555AB60, CInt(&H153B), CInt(&H4D17), &H9F, &H4, &HA5, &HFE, &H99, &HFC, &H15, &HEC)
FOLDERID_PublicRingtones = iid
End Function
Public Function FOLDERID_PublicUserTiles() As UUID
'{0482af6c-08f1-4c34-8c90-e17ec98b1e17}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H482AF6C, CInt(&H8F1), CInt(&H4C34), &H8C, &H90, &HE1, &H7E, &HC9, &H8B, &H1E, &H17)
FOLDERID_PublicUserTiles = iid
End Function
Public Function FOLDERID_RecordedTVLibrary() As UUID
'{1A6FDBA2-F42D-4358-A798-B74D745926C5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1A6FDBA2, CInt(&HF42D), CInt(&H4358), &HA7, &H98, &HB7, &H4D, &H74, &H59, &H26, &HC5)
FOLDERID_RecordedTVLibrary = iid
End Function
Public Function FOLDERID_Ringtones() As UUID
'{C870044B-F49E-4126-A9C3-B52A1FF411E8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC870044B, CInt(&HF49E), CInt(&H4126), &HA9, &HC3, &HB5, &H2A, &H1F, &HF4, &H11, &HE8)
FOLDERID_Ringtones = iid
End Function
Public Function FOLDERID_RoamedTileImages() As UUID
'{AAA8D5A5-F1D6-4259-BAA8-78E7EF60835E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAAA8D5A5, CInt(&HF1D6), CInt(&H4259), &HBA, &HA8, &H78, &HE7, &HEF, &H60, &H83, &H5E)
FOLDERID_RoamedTileImages = iid
End Function
Public Function FOLDERID_RoamingTiles() As UUID
'{00BCFC5A-ED94-4e48-96A1-3F6217F21990}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBCFC5A, CInt(&HED94), CInt(&H4E48), &H96, &HA1, &H3F, &H62, &H17, &HF2, &H19, &H90)
FOLDERID_RoamingTiles = iid
End Function
Public Function FOLDERID_SavedPictures() As UUID
'{3B193882-D3AD-4eab-965A-69829D1FB59F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3B193882, CInt(&HD3AD), CInt(&H4EAB), &H96, &H5A, &H69, &H82, &H9D, &H1F, &HB5, &H9F)
FOLDERID_SavedPictures = iid
End Function
Public Function FOLDERID_SavedPicturesLibrary() As UUID
'{E25B5812-BE88-4bd9-94B0-29233477B6C3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE25B5812, CInt(&HBE88), CInt(&H4BD9), &H94, &HB0, &H29, &H23, &H34, &H77, &HB6, &HC3)
FOLDERID_SavedPicturesLibrary = iid
End Function
Public Function FOLDERID_Screenshots() As UUID
'{b7bede81-df94-4682-a7d8-57a52620b86f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB7BEDE81, CInt(&HDF94), CInt(&H4682), &HA7, &HD8, &H57, &HA5, &H26, &H20, &HB8, &H6F)
FOLDERID_Screenshots = iid
End Function
Public Function FOLDERID_SearchHistory() As UUID
'{0D4C3DB6-03A3-462F-A0E6-08924C41B5D4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD4C3DB6, CInt(&H3A3), CInt(&H462F), &HA0, &HE6, &H8, &H92, &H4C, &H41, &HB5, &HD4)
FOLDERID_SearchHistory = iid
End Function
Public Function FOLDERID_SearchTemplates() As UUID
'{7E636BFE-DFA9-4D5E-B456-D7B39851D8A9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7E636BFE, CInt(&HDFA9), CInt(&H4D5E), &HB4, &H56, &HD7, &HB3, &H98, &H51, &HD8, &HA9)
FOLDERID_SearchTemplates = iid
End Function
Public Function FOLDERID_SkyDrive() As UUID
'{A52BBA46-E9E1-435f-B3D9-28DAA648C0F6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA52BBA46, CInt(&HE9E1), CInt(&H435F), &HB3, &HD9, &H28, &HDA, &HA6, &H48, &HC0, &HF6)
FOLDERID_SkyDrive = iid
End Function
Public Function FOLDERID_SkyDriveCameraRoll() As UUID
'{767E6811-49CB-4273-87C2-20F355E1085B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H767E6811, CInt(&H49CB), CInt(&H4273), &H87, &HC2, &H20, &HF3, &H55, &HE1, &H8, &H5B)
FOLDERID_SkyDriveCameraRoll = iid
End Function
Public Function FOLDERID_SkyDriveDocuments() As UUID
'{24D89E24-2F19-4534-9DDE-6A6671FBB8FE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H24D89E24, CInt(&H2F19), CInt(&H4534), &H9D, &HDE, &H6A, &H66, &H71, &HFB, &HB8, &HFE)
FOLDERID_SkyDriveDocuments = iid
End Function
Public Function FOLDERID_SkyDrivePictures() As UUID
'{339719B5-8C47-4894-94C2-D8F77ADD44A6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H339719B5, CInt(&H8C47), CInt(&H4894), &H94, &HC2, &HD8, &HF7, &H7A, &HDD, &H44, &HA6)
FOLDERID_SkyDrivePictures = iid
End Function
Public Function FOLDERID_UserPinned() As UUID
'{9E3995AB-1F9C-4F13-B827-48B24B6C7174}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9E3995AB, CInt(&H1F9C), CInt(&H4F13), &HB8, &H27, &H48, &HB2, &H4B, &H6C, &H71, &H74)
FOLDERID_UserPinned = iid
End Function
Public Function FOLDERID_UserProgramFiles() As UUID
'{5CD7AEE2-2219-4A67-B85D-6C9CE15660CB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5CD7AEE2, CInt(&H2219), CInt(&H4A67), &HB8, &H5D, &H6C, &H9C, &HE1, &H56, &H60, &HCB)
FOLDERID_UserProgramFiles = iid
End Function
Public Function FOLDERID_UserProgramFilesCommon() As UUID
'{BCBD3057-CA5C-4622-B42D-BC56DB0AE516}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBCBD3057, CInt(&HCA5C), CInt(&H4622), &HB4, &H2D, &HBC, &H56, &HDB, &HA, &HE5, &H16)
FOLDERID_UserProgramFilesCommon = iid
End Function
Public Function FOLDERID_UsersLibraries() As UUID
'{A302545D-DEFF-464b-ABE8-61C8648D939B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA302545D, CInt(&HDEFF), CInt(&H464B), &HAB, &HE8, &H61, &HC8, &H64, &H8D, &H93, &H9B)
FOLDERID_UsersLibraries = iid
End Function
Public Function FOLDERID_VideosLibrary() As UUID
'{491E922F-5643-4AF4-A7EB-4E7A138D8174 }
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H491E922F, CInt(&H5643), CInt(&H4AF4), &HA7, &HEB, &H4E, &H7A, &H13, &H8D, &H81, &H74)
FOLDERID_VideosLibrary = iid
End Function
Public Function FOLDERID_RetailDemo() As UUID
'{12D4C69E-24AD-4923-BE19-31321C43A767}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H12D4C69E, CInt(&H24AD), CInt(&H4923), &HBE, &H19, &H31, &H32, &H1C, &H43, &HA7, &H67)
FOLDERID_RetailDemo = iid
End Function
Public Function FOLDERID_Device() As UUID
'{1C2AC1DC-4358-4B6C-9733-AF21156576F0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1C2AC1DC, CInt(&H4358), CInt(&H4B6C), &H97, &H33, &HAF, &H21, &H15, &H65, &H76, &HF0)
FOLDERID_Device = iid
End Function
Public Function FOLDERID_DevelopmentFiles() As UUID
'{DBE8E08E-3053-4BBC-B183-2A7B2B191E59}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDBE8E08E, CInt(&H3053), CInt(&H4BBC), &HB1, &H83, &H2A, &H7B, &H2B, &H19, &H1E, &H59)
FOLDERID_DevelopmentFiles = iid
End Function
Public Function FOLDERID_AppCaptures() As UUID
'{EDC0FE71-98D8-4F4A-B920-C8DC133CB165}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEDC0FE71, CInt(&H98D8), CInt(&H4F4A), &HB9, &H20, &HC8, &HDC, &H13, &H3C, &HB1, &H65)
FOLDERID_AppCaptures = iid
End Function
Public Function FOLDERID_LocalDocuments() As UUID
'{f42ee2d3-909f-4907-8871-4c22fc0bf756}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF42EE2D3, CInt(&H909F), CInt(&H4907), &H88, &H71, &H4C, &H22, &HFC, &HB, &HF7, &H56)
FOLDERID_LocalDocuments = iid
End Function
Public Function FOLDERID_LocalPictures() As UUID
'{0ddd015d-b06c-45d5-8c4c-f59713854639}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDDD015D, CInt(&HB06C), CInt(&H45D5), &H8C, &H4C, &HF5, &H97, &H13, &H85, &H46, &H39)
FOLDERID_LocalPictures = iid
End Function
Public Function FOLDERID_LocalVideos() As UUID
'{35286a68-3c57-41a1-bbb1-0eae73d76c95}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H35286A68, CInt(&H3C57), CInt(&H41A1), &HBB, &HB1, &HE, &HAE, &H73, &HD7, &H6C, &H95)
FOLDERID_LocalVideos = iid
End Function
Public Function FOLDERID_LocalMusic() As UUID
'{a0c69a99-21c8-4671-8703-7934162fcf1d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA0C69A99, CInt(&H21C8), CInt(&H4671), &H87, &H3, &H79, &H34, &H16, &H2F, &HCF, &H1D)
FOLDERID_LocalMusic = iid
End Function
Public Function FOLDERID_LocalDownloads() As UUID
'{7d83ee9b-2244-4e70-b1f5-5393042af1e4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7D83EE9B, CInt(&H2244), CInt(&H4E70), &HB1, &HF5, &H53, &H93, &H4, &H2A, &HF1, &HE4)
FOLDERID_LocalDownloads = iid
End Function
Public Function FOLDERID_RecordedCalls() As UUID
'{2f8b40c2-83ed-48ee-b383-a1f157ec6f9a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2F8B40C2, CInt(&H83ED), CInt(&H48EE), &HB3, &H83, &HA1, &HF1, &H57, &HEC, &H6F, &H9A)
FOLDERID_RecordedCalls = iid
End Function
Public Function FOLDERID_AllAppMods() As UUID
'{7ad67899-66af-43ba-9156-6aad42e6c596}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7AD67899, CInt(&H66AF), CInt(&H43BA), &H91, &H56, &H6A, &HAD, &H42, &HE6, &HC5, &H96)
FOLDERID_AllAppMods = iid
End Function
Public Function FOLDERID_CurrentAppMods() As UUID
'{3db40b20-2a30-4dbe-917e-771dd21dd099}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3DB40B20, CInt(&H2A30), CInt(&H4DBE), &H91, &H7E, &H77, &H1D, &HD2, &H1D, &HD0, &H99)
FOLDERID_CurrentAppMods = iid
End Function
Public Function FOLDERID_AppDataDocuments() As UUID
'{7BE16610-1F7F-44AC-BFF0-83E15F2FFCA1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7BE16610, CInt(&H1F7F), CInt(&H44AC), &HBF, &HF0, &H83, &HE1, &H5F, &H2F, &HFC, &HA1)
FOLDERID_AppDataDocuments = iid
End Function
Public Function FOLDERID_AppDataFavorites() As UUID
'{7CFBEFBC-DE1F-45AA-B843-A542AC536CC9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7CFBEFBC, CInt(&HDE1F), CInt(&H45AA), &HB8, &H43, &HA5, &H42, &HAC, &H53, &H6C, &HC9)
FOLDERID_AppDataFavorites = iid
End Function
Public Function FOLDERID_AppDataProgramData() As UUID
'{559D40A3-A036-40FA-AF61-84CB430A4D34}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H559D40A3, CInt(&HA036), CInt(&H40FA), &HAF, &H61, &H84, &HCB, &H43, &HA, &H4D, &H34)
FOLDERID_AppDataProgramData = iid
End Function
Public Function FOLDERTYPEID_Invalid() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_Invalid, &H57807898, &H8C4F, &H4462, &HBB, &H63, &H71, &H4, &H23, &H80, &HB1, &H9)
End Function
Public Function FOLDERTYPEID_Generic() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_Generic, &H5C4F28B5, &HF869, &H4E84, &H8E, &H60, &HF1, &H1D, &HB9, &H7C, &H5C, &HC7)
End Function
Public Function FOLDERTYPEID_GenericSearchResults() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_GenericSearchResults, &H7FDE1A1E, &H8B31, &H49A5, &H93, &HB8, &H6B, &HE1, &H4C, &HFA, &H49, &H43)
End Function
Public Function FOLDERTYPEID_GenericLibrary() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_GenericLibrary, &H5F4EAB9A, &H6833, &H4F61, &H89, &H9D, &H31, &HCF, &H46, &H97, &H9D, &H49)
End Function
Public Function FOLDERTYPEID_Documents() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_Documents, &H7D49D726, &H3C21, &H4F05, &H99, &HAA, &HFD, &HC2, &HC9, &H47, &H46, &H56)
End Function
Public Function FOLDERTYPEID_Pictures() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_Pictures, &HB3690E58, &HE961, &H423B, &HB6, &H87, &H38, &H6E, &HBF, &HD8, &H32, &H39)
End Function
Public Function FOLDERTYPEID_Music() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_Music, &H94D6DDCC, &H4A68, &H4175, &HA3, &H74, &HBD, &H58, &H4A, &H51, &HB, &H78)
End Function
Public Function FOLDERTYPEID_Videos() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_Videos, &H5FA96407, &H7E77, &H483C, &HAC, &H93, &H69, &H1D, &H5, &H85, &HD, &HE8)
End Function
Public Function FOLDERTYPEID_UserFiles() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_UserFiles, &HCD0FC69B, &H71E2, &H46E5, &H96, &H90, &H5B, &HCD, &H9F, &H57, &HAA, &HB3)
End Function
Public Function FOLDERTYPEID_UsersLibraries() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_UsersLibraries, &HC4D98F09, &H6124, &H4FE0, &H99, &H42, &H82, &H64, &H16, &H8, &H2D, &HA9)
End Function
Public Function FOLDERTYPEID_OtherUsers() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_OtherUsers, &HB337FD00, &H9DD5, &H4635, &HA6, &HD4, &HDA, &H33, &HFD, &H10, &H2B, &H7A)
End Function
Public Function FOLDERTYPEID_PublishedItems() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_PublishedItems, &H7F2F5B96, &HFF74, &H41DA, &HAF, &HD8, &H1C, &H78, &HA5, &HF3, &HAE, &HA2)
End Function
Public Function FOLDERTYPEID_Communications() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_Communications, &H91475FE5, &H586B, &H4EBA, &H8D, &H75, &HD1, &H74, &H34, &HB8, &HCD, &HF6)
End Function
Public Function FOLDERTYPEID_Contacts() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_Contacts, &HDE2B70EC, &H9BF7, &H4A93, &HBD, &H3D, &H24, &H3F, &H78, &H81, &HD4, &H92)
End Function
Public Function FOLDERTYPEID_StartMenu() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_StartMenu, &HEF87B4CB, &HF2CE, &H4785, &H86, &H58, &H4C, &HA6, &HC6, &H3E, &H38, &HC6)
End Function
Public Function FOLDERTYPEID_RecordedTV() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_RecordedTV, &H5557A28F, &H5DA6, &H4F83, &H88, &H9, &HC2, &HC9, &H8A, &H11, &HA6, &HFA)
End Function
Public Function FOLDERTYPEID_SavedGames() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_SavedGames, &HD0363307, &H28CB, &H4106, &H9F, &H23, &H29, &H56, &HE3, &HE5, &HE0, &HE7)
End Function
Public Function FOLDERTYPEID_OpenSearch() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_OpenSearch, &H8FAF9629, &H1980, &H46FF, &H80, &H23, &H9D, &HCE, &HAB, &H9C, &H3E, &HE3)
End Function
Public Function FOLDERTYPEID_SearchConnector() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_SearchConnector, &H982725EE, &H6F47, &H479E, &HB4, &H47, &H81, &H2B, &HFA, &H7D, &H2E, &H8F)
End Function
Public Function FOLDERTYPEID_AccountPictures() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_AccountPictures, &HDB2A5D8F, &H6E6, &H4007, &HAB, &HA6, &HAF, &H87, &H7D, &H52, &H6E, &HA6)
End Function
Public Function FOLDERTYPEID_Games() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_Games, &HB689B0D0, &H76D3, &H4CBB, &H87, &HF7, &H58, &H5D, &HE, &HC, &HE0, &H70)
End Function
Public Function FOLDERTYPEID_ControlPanelCategory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_ControlPanelCategory, &HDE4F0660, &HFA10, &H4B8F, &HA4, &H94, &H6, &H8B, &H20, &HB2, &H23, &H7)
End Function
Public Function FOLDERTYPEID_ControlPanelClassic() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_ControlPanelClassic, &HC3794F3, &HB545, &H43AA, &HA3, &H29, &HC3, &H74, &H30, &HC5, &H8D, &H2A)
End Function
Public Function FOLDERTYPEID_Printers() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_Printers, &H2C7BBEC6, &HC844, &H4A0A, &H91, &HFA, &HCE, &HF6, &HF5, &H9C, &HFD, &HA1)
End Function
Public Function FOLDERTYPEID_RecycleBin() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_RecycleBin, &HD6D9E004, &HCD87, &H442B, &H9D, &H57, &H5E, &HA, &HEB, &H4F, &H6F, &H72)
End Function
Public Function FOLDERTYPEID_SoftwareExplorer() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_SoftwareExplorer, &HD674391B, &H52D9, &H4E07, &H83, &H4E, &H67, &HC9, &H86, &H10, &HF3, &H9D)
End Function
Public Function FOLDERTYPEID_CompressedFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_CompressedFolder, &H80213E82, &HBCFD, &H4C4F, &H88, &H17, &HBB, &H27, &H60, &H12, &H67, &HA9)
End Function
Public Function FOLDERTYPEID_NetworkExplorer() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_NetworkExplorer, &H25CC242B, &H9A7C, &H4F51, &H80, &HE0, &H7A, &H29, &H28, &HFE, &HBE, &H42)
End Function
Public Function FOLDERTYPEID_Searches() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_Searches, &HB0BA2E3, &H405F, &H415E, &HA6, &HEE, &HCA, &HD6, &H25, &H20, &H78, &H53)
End Function
Public Function FOLDERTYPEID_SearchHome() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_SearchHome, &H834D8A44, &H974, &H4ED6, &H86, &H6E, &HF2, &H3, &HD8, &HB, &H38, &H10)
End Function
Public Function FOLDERTYPEID_StorageProviderGeneric() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_StorageProviderGeneric, &H4F01EBC5, &H2385, &H41F2, &HA2, &H8E, &H2C, &H5C, &H91, &HFB, &H56, &HE0)
End Function
Public Function FOLDERTYPEID_StorageProviderDocuments() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_StorageProviderDocuments, &HDD61BD66, &H70E8, &H48DD, &H96, &H55, &H65, &HC5, &HE1, &HAA, &HC2, &HD1)
End Function
Public Function FOLDERTYPEID_StorageProviderPictures() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_StorageProviderPictures, &H71D642A9, &HF2B1, &H42CD, &HAD, &H92, &HEB, &H93, &H0, &HC7, &HCC, &HA)
End Function
Public Function FOLDERTYPEID_StorageProviderMusic() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_StorageProviderMusic, &H672ECD7E, &HAF04, &H4399, &H87, &H5C, &H2, &H90, &H84, &H5B, &H62, &H47)
End Function
Public Function FOLDERTYPEID_StorageProviderVideos() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(FOLDERTYPEID_StorageProviderVideos, &H51294DA1, &HD7B1, &H485B, &H9E, &H9A, &H17, &HCF, &HFE, &H33, &HE1, &H87)
End Function

Public Function VID_LargeIcons() As UUID
'{0057D0E0-3573-11CF-AE69-08002B2E1262}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57D0E0, CInt(&H3573), CInt(&H11CF), &HAE, &H69, &H8, &H0, &H2B, &H2E, &H12, &H62)
 VID_LargeIcons = iid
End Function
Public Function VID_SmallIcons() As UUID
'{089000C0-3573-11CF-AE69-08002B2E1262}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H89000C0, CInt(&H3573), CInt(&H11CF), &HAE, &H69, &H8, &H0, &H2B, &H2E, &H12, &H62)
 VID_SmallIcons = iid
End Function
Public Function VID_List() As UUID
'{0E1FA5E0-3573-11CF-AE69-08002B2E1262}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE1FA5E0, CInt(&H3573), CInt(&H11CF), &HAE, &H69, &H8, &H0, &H2B, &H2E, &H12, &H62)
 VID_List = iid
End Function
Public Function VID_Details() As UUID
'{137E7700-3573-11CF-AE69-08002B2E1262}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H137E7700, CInt(&H3573), CInt(&H11CF), &HAE, &H69, &H8, &H0, &H2B, &H2E, &H12, &H62)
 VID_Details = iid
End Function
Public Function VID_Tile() As UUID
'{65F125E5-7BE1-4810-BA9D-D271C8432CE3}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H65F125E5, CInt(&H7BE1), CInt(&H4810), &HBA, &H9D, &HD2, &H71, &HC8, &H43, &H2C, &HE3)
 VID_Tile = iid
End Function
Public Function EP_NavPane() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCB316B22, &H25F7, &H42B8, &H8A, &H9, &H54, &HD, &H23, &HA4, &H3C, &H2F)
EP_NavPane = iid
End Function
Public Function EP_Commands() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(EP_Commands, &HD9745868, &HCA5F, &H4A76, &H91, &HCD, &HF5, &HA1, &H29, &HFB, &HB0, &H76)
EP_Commands = iid
End Function
Public Function EP_Commands_Organize() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H72E81700, &HE3EC, &H4660, &HBF, &H24, &H3C, &H3B, &H7B, &H64, &H88, &H6)
EP_Commands_Organize = iid
End Function
Public Function EP_Commands_View() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H21F7C32D, &HEEAA, &H439B, &HBB, &H51, &H37, &HB9, &H6F, &HD6, &HA9, &H43)
EP_Commands_View = iid
End Function
Public Function EP_DetailsPane() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43ABF98B, &H89B8, &H472D, &HB9, &HCE, &HE6, &H9B, &H82, &H29, &HF0, &H19)
EP_DetailsPane = iid
End Function
Public Function EP_PreviewPane() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H893C63D1, &H45C8, &H4D17, &HBE, &H19, &H22, &H3B, &HE7, &H1B, &HE3, &H65)
EP_PreviewPane = iid
End Function
Public Function EP_QueryPane() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H65BCDE4F, &H4F07, &H4F27, &H83, &HA7, &H1A, &HFC, &HA4, &HDF, &H7D, &HDD)
EP_QueryPane = iid
End Function
Public Function EP_AdvQueryPane() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB4E9DB8B, &H34BA, &H4C39, &HB5, &HCC, &H16, &HA1, &HBD, &H2C, &H41, &H1C)
EP_AdvQueryPane = iid
End Function
Public Function EP_Ribbon() As UUID
'{d27524a8-c9f2-4834-a106-df8889fd4f37}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD27524A8, CInt(&HC9F2), CInt(&H4834), &HA1, &H6, &HDF, &H88, &H89, &HFD, &H4F, &H37)
 EP_Ribbon = iid
End Function
Public Function EP_StatusBar() As UUID
'{65fe56ce-5cfe-4bc4-ad8a-7ae3fe7e8f7c}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H65FE56CE, CInt(&H5CFE), CInt(&H4BC4), &HAD, &H8A, &H7A, &HE3, &HFE, &H7E, &H8F, &H7C)
 EP_StatusBar = iid
End Function


Public Function DOMAIN_JOIN_GUID() As UUID
'{1ce20aba-9851-4421-9430-1ddeb766e809}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1CE20ABA, CInt(&H9851), CInt(&H4421), &H94, &H30, &H1D, &HDE, &HB7, &H66, &HE8, &H9)
 DOMAIN_JOIN_GUID = iid
End Function
Public Function FIREWALL_PORT_CLOSE_GUID() As UUID
'{a144ed38-8e12-4de4-9d96-e64740b1a524}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA144ED38, CInt(&H8E12), CInt(&H4DE4), &H9D, &H96, &HE6, &H47, &H40, &HB1, &HA5, &H24)
 FIREWALL_PORT_CLOSE_GUID = iid
End Function
Public Function FIREWALL_PORT_OPEN_GUID() As UUID
'{b7569e07-8421-4ee0-ad10-86915afdad09}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB7569E07, CInt(&H8421), CInt(&H4EE0), &HAD, &H10, &H86, &H91, &H5A, &HFD, &HAD, &H9)
 FIREWALL_PORT_OPEN_GUID = iid
End Function
Public Function MACHINE_POLICY_PRESENT_GUID() As UUID
'{659FCAE6-5BDB-4DA9-B1FF-CA2A178D46E0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H659FCAE6, CInt(&H5BDB), CInt(&H4DA9), &HB1, &HFF, &HCA, &H2A, &H17, &H8D, &H46, &HE0)
 MACHINE_POLICY_PRESENT_GUID = iid
End Function
Public Function USER_POLICY_PRESENT_GUID() As UUID
'{54FB46C8-F089-464C-B1FD-59D1B62C3B50}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H54FB46C8, CInt(&HF089), CInt(&H464C), &HB1, &HFD, &H59, &HD1, &HB6, &H2C, &H3B, &H50)
 USER_POLICY_PRESENT_GUID = iid
End Function
Public Function RPC_INTERFACE_EVENT_GUID() As UUID
'{bc90d167-9470-4139-a9ba-be0bbbf5b74d}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBC90D167, CInt(&H9470), CInt(&H4139), &HA9, &HBA, &HBE, &HB, &HBB, &HF5, &HB7, &H4D)
 RPC_INTERFACE_EVENT_GUID = iid
End Function
Public Function NAMED_PIPE_EVENT_GUID() As UUID
'{1f81d131-3fac-4537-9e0c-7e7b0c2f4b55}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1F81D131, CInt(&H3FAC), CInt(&H4537), &H9E, &HC, &H7E, &H7B, &HC, &H2F, &H4B, &H55)
 NAMED_PIPE_EVENT_GUID = iid
End Function
Public Function CUSTOM_SYSTEM_STATE_CHANGE_EVENT_GUID() As UUID
'{2d7a2816-0c5e-45fc-9ce7-570e5ecde9c9}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2D7A2816, CInt(&HC5E), CInt(&H45FC), &H9C, &HE7, &H57, &HE, &H5E, &HCD, &HE9, &HC9)
 CUSTOM_SYSTEM_STATE_CHANGE_EVENT_GUID = iid
End Function

Public Function GUID_DEVCLASS_1394() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6BDD1FC1, &H810F, &H11D0, &HBE, &HC7, &H8, &H0, &H2B, &HE2, &H9, &H2F)
 GUID_DEVCLASS_1394 = iid
End Function
Public Function GUID_DEVCLASS_1394DEBUG() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H66F250D6, &H7801, &H4A64, &HB1, &H39, &HEE, &HA8, &HA, &H45, &HB, &H24)
 GUID_DEVCLASS_1394DEBUG = iid
End Function
Public Function GUID_DEVCLASS_61883() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7EBEFBC0, &H3200, &H11D2, &HB4, &HC2, &H0, &HA0, &HC9, &H69, &H7D, &H7)
 GUID_DEVCLASS_61883 = iid
End Function
Public Function GUID_DEVCLASS_ADAPTER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E964, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_ADAPTER = iid
End Function
Public Function GUID_DEVCLASS_APMSUPPORT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD45B1C18, &HC8FA, &H11D1, &H9F, &H77, &H0, &H0, &HF8, &H5, &HF5, &H30)
 GUID_DEVCLASS_APMSUPPORT = iid
End Function
Public Function GUID_DEVCLASS_AVC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC06FF265, &HAE09, &H48F0, &H81, &H2C, &H16, &H75, &H3D, &H7C, &HBA, &H83)
 GUID_DEVCLASS_AVC = iid
End Function
Public Function GUID_DEVCLASS_BATTERY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H72631E54, &H78A4, &H11D0, &HBC, &HF7, &H0, &HAA, &H0, &HB7, &HB3, &H2A)
 GUID_DEVCLASS_BATTERY = iid
End Function
Public Function GUID_DEVCLASS_BIOMETRIC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53D29EF7, &H377C, &H4D14, &H86, &H4B, &HEB, &H3A, &H85, &H76, &H93, &H59)
 GUID_DEVCLASS_BIOMETRIC = iid
End Function
Public Function GUID_DEVCLASS_BLUETOOTH() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE0CBF06C, &HCD8B, &H4647, &HBB, &H8A, &H26, &H3B, &H43, &HF0, &HF9, &H74)
 GUID_DEVCLASS_BLUETOOTH = iid
End Function
Public Function GUID_DEVCLASS_CAMERA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCA3E7AB9, &HB4C3, &H4AE6, &H82, &H51, &H57, &H9E, &HF9, &H33, &H89, &HF)
 GUID_DEVCLASS_CAMERA = iid
End Function
Public Function GUID_DEVCLASS_CDROM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E965, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_CDROM = iid
End Function
Public Function GUID_DEVCLASS_COMPUTEACCELERATOR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF01A9D53, &H3FF6, &H48D2, &H9F, &H97, &HC8, &HA7, &H0, &H4B, &HE1, &HC)
 GUID_DEVCLASS_COMPUTEACCELERATOR = iid
End Function
Public Function GUID_DEVCLASS_COMPUTER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E966, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_COMPUTER = iid
End Function
Public Function GUID_DEVCLASS_DECODER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6BDD1FC2, &H810F, &H11D0, &HBE, &HC7, &H8, &H0, &H2B, &HE2, &H9, &H2F)
 GUID_DEVCLASS_DECODER = iid
End Function
Public Function GUID_DEVCLASS_DISKDRIVE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E967, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_DISKDRIVE = iid
End Function
Public Function GUID_DEVCLASS_DISPLAY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E968, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_DISPLAY = iid
End Function
Public Function GUID_DEVCLASS_DOT4() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H48721B56, &H6795, &H11D2, &HB1, &HA8, &H0, &H80, &HC7, &H2E, &H74, &HA2)
 GUID_DEVCLASS_DOT4 = iid
End Function
Public Function GUID_DEVCLASS_DOT4PRINT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H49CE6AC8, &H6F86, &H11D2, &HB1, &HE5, &H0, &H80, &HC7, &H2E, &H74, &HA2)
 GUID_DEVCLASS_DOT4PRINT = iid
End Function
Public Function GUID_DEVCLASS_EHSTORAGESILO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9DA2B80F, &HF89F, &H4A49, &HA5, &HC2, &H51, &H1B, &H8, &H5B, &H9E, &H8A)
 GUID_DEVCLASS_EHSTORAGESILO = iid
End Function
Public Function GUID_DEVCLASS_ENUM1394() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC459DF55, &HDB08, &H11D1, &HB0, &H9, &H0, &HA0, &HC9, &H8, &H1F, &HF6)
 GUID_DEVCLASS_ENUM1394 = iid
End Function
Public Function GUID_DEVCLASS_EXTENSION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE2F84CE7, &H8EFA, &H411C, &HAA, &H69, &H97, &H45, &H4C, &HA4, &HCB, &H57)
 GUID_DEVCLASS_EXTENSION = iid
End Function
Public Function GUID_DEVCLASS_FDC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E969, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_FDC = iid
End Function
Public Function GUID_DEVCLASS_FIRMWARE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF2E7DD72, &H6468, &H4E36, &HB6, &HF1, &H64, &H88, &HF4, &H2C, &H1B, &H52)
 GUID_DEVCLASS_FIRMWARE = iid
End Function
Public Function GUID_DEVCLASS_FLOPPYDISK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E980, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_FLOPPYDISK = iid
End Function
Public Function GUID_DEVCLASS_GENERIC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFF494DF1, &HC4ED, &H4FAC, &H9B, &H3F, &H37, &H86, &HF6, &HE9, &H1E, &H7E)
 GUID_DEVCLASS_GENERIC = iid
End Function
Public Function GUID_DEVCLASS_GPS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6BDD1FC3, &H810F, &H11D0, &HBE, &HC7, &H8, &H0, &H2B, &HE2, &H9, &H2F)
 GUID_DEVCLASS_GPS = iid
End Function
Public Function GUID_DEVCLASS_HDC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E96A, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_HDC = iid
End Function
Public Function GUID_DEVCLASS_HIDCLASS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H745A17A0, &H74D3, &H11D0, &HB6, &HFE, &H0, &HA0, &HC9, &HF, &H57, &HDA)
 GUID_DEVCLASS_HIDCLASS = iid
End Function
Public Function GUID_DEVCLASS_HOLOGRAPHIC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD612553D, &H6B1, &H49CA, &H89, &H38, &HE3, &H9E, &HF8, &HE, &HB1, &H6F)
 GUID_DEVCLASS_HOLOGRAPHIC = iid
End Function
Public Function GUID_DEVCLASS_IMAGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6BDD1FC6, &H810F, &H11D0, &HBE, &HC7, &H8, &H0, &H2B, &HE2, &H9, &H2F)
 GUID_DEVCLASS_IMAGE = iid
End Function
Public Function GUID_DEVCLASS_INFINIBAND() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H30EF7132, &HD858, &H4A0C, &HAC, &H24, &HB9, &H2, &H8A, &H5C, &HCA, &H3F)
 GUID_DEVCLASS_INFINIBAND = iid
End Function
Public Function GUID_DEVCLASS_INFRARED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6BDD1FC5, &H810F, &H11D0, &HBE, &HC7, &H8, &H0, &H2B, &HE2, &H9, &H2F)
 GUID_DEVCLASS_INFRARED = iid
End Function
Public Function GUID_DEVCLASS_KEYBOARD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E96B, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_KEYBOARD = iid
End Function
Public Function GUID_DEVCLASS_LEGACYDRIVER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8ECC055D, &H47F, &H11D1, &HA5, &H37, &H0, &H0, &HF8, &H75, &H3E, &HD1)
 GUID_DEVCLASS_LEGACYDRIVER = iid
End Function
Public Function GUID_DEVCLASS_MEDIA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E96C, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_MEDIA = iid
End Function
Public Function GUID_DEVCLASS_MEDIUM_CHANGER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCE5939AE, &HEBDE, &H11D0, &HB1, &H81, &H0, &H0, &HF8, &H75, &H3E, &HC4)
 GUID_DEVCLASS_MEDIUM_CHANGER = iid
End Function
Public Function GUID_DEVCLASS_MEMORY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5099944A, &HF6B9, &H4057, &HA0, &H56, &H8C, &H55, &H2, &H28, &H54, &H4C)
 GUID_DEVCLASS_MEMORY = iid
End Function
Public Function GUID_DEVCLASS_MODEM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E96D, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_MODEM = iid
End Function
Public Function GUID_DEVCLASS_MONITOR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E96E, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_MONITOR = iid
End Function
Public Function GUID_DEVCLASS_MOUSE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E96F, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_MOUSE = iid
End Function
Public Function GUID_DEVCLASS_MTD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E970, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_MTD = iid
End Function
Public Function GUID_DEVCLASS_MULTIFUNCTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E971, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_MULTIFUNCTION = iid
End Function
Public Function GUID_DEVCLASS_MULTIPORTSERIAL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H50906CB8, &HBA12, &H11D1, &HBF, &H5D, &H0, &H0, &HF8, &H5, &HF5, &H30)
 GUID_DEVCLASS_MULTIPORTSERIAL = iid
End Function
Public Function GUID_DEVCLASS_NET() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E972, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_NET = iid
End Function
Public Function GUID_DEVCLASS_NETCLIENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E973, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_NETCLIENT = iid
End Function
Public Function GUID_DEVCLASS_NETDRIVER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H87EF9AD1, &H8F70, &H49EE, &HB2, &H15, &HAB, &H1F, &HCA, &HDC, &HBE, &H3C)
 GUID_DEVCLASS_NETDRIVER = iid
End Function
Public Function GUID_DEVCLASS_NETSERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E974, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_NETSERVICE = iid
End Function
Public Function GUID_DEVCLASS_NETTRANS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E975, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_NETTRANS = iid
End Function
Public Function GUID_DEVCLASS_NETUIO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H78912BC1, &HCB8E, &H4B28, &HA3, &H29, &HF3, &H22, &HEB, &HAD, &HBE, &HF)
 GUID_DEVCLASS_NETUIO = iid
End Function
Public Function GUID_DEVCLASS_NODRIVER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E976, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_NODRIVER = iid
End Function
Public Function GUID_DEVCLASS_PCMCIA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E977, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_PCMCIA = iid
End Function
Public Function GUID_DEVCLASS_PNPPRINTERS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4658EE7E, &HF050, &H11D1, &HB6, &HBD, &H0, &HC0, &H4F, &HA3, &H72, &HA7)
 GUID_DEVCLASS_PNPPRINTERS = iid
End Function
Public Function GUID_DEVCLASS_PORTS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E978, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_PORTS = iid
End Function
Public Function GUID_DEVCLASS_PRIMITIVE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H242681D1, &HEED3, &H41D2, &HA1, &HEF, &H14, &H68, &HFC, &H84, &H31, &H6)
 GUID_DEVCLASS_PRIMITIVE = iid
End Function
Public Function GUID_DEVCLASS_PRINTER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E979, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_PRINTER = iid
End Function
Public Function GUID_DEVCLASS_PRINTERUPGRADE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E97A, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_PRINTERUPGRADE = iid
End Function
Public Function GUID_DEVCLASS_PRINTQUEUE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1ED2BBF9, &H11F0, &H4084, &HB2, &H1F, &HAD, &H83, &HA8, &HE6, &HDC, &HDC)
 GUID_DEVCLASS_PRINTQUEUE = iid
End Function
Public Function GUID_DEVCLASS_PROCESSOR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H50127DC3, &HF36, &H415E, &HA6, &HCC, &H4C, &HB3, &HBE, &H91, &HB, &H65)
 GUID_DEVCLASS_PROCESSOR = iid
End Function
Public Function GUID_DEVCLASS_SBP2() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD48179BE, &HEC20, &H11D1, &HB6, &HB8, &H0, &HC0, &H4F, &HA3, &H72, &HA7)
 GUID_DEVCLASS_SBP2 = iid
End Function
Public Function GUID_DEVCLASS_SCMDISK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53966CB1, &H4D46, &H4166, &HBF, &H23, &HC5, &H22, &H40, &H3C, &HD4, &H95)
 GUID_DEVCLASS_SCMDISK = iid
End Function
Public Function GUID_DEVCLASS_SCMVOLUME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53CCB149, &HE543, &H4C84, &HB6, &HE0, &HBC, &HE4, &HF6, &HB7, &HE8, &H6)
 GUID_DEVCLASS_SCMVOLUME = iid
End Function
Public Function GUID_DEVCLASS_SCSIADAPTER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E97B, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_SCSIADAPTER = iid
End Function
Public Function GUID_DEVCLASS_SECURITYACCELERATOR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H268C95A1, &HEDFE, &H11D3, &H95, &HC3, &H0, &H10, &HDC, &H40, &H50, &HA5)
 GUID_DEVCLASS_SECURITYACCELERATOR = iid
End Function
Public Function GUID_DEVCLASS_SENSOR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5175D334, &HC371, &H4806, &HB3, &HBA, &H71, &HFD, &H53, &HC9, &H25, &H8D)
 GUID_DEVCLASS_SENSOR = iid
End Function
Public Function GUID_DEVCLASS_SIDESHOW() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H997B5D8D, &HC442, &H4F2E, &HBA, &HF3, &H9C, &H8E, &H67, &H1E, &H9E, &H21)
 GUID_DEVCLASS_SIDESHOW = iid
End Function
Public Function GUID_DEVCLASS_SMARTCARDREADER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H50DD5230, &HBA8A, &H11D1, &HBF, &H5D, &H0, &H0, &HF8, &H5, &HF5, &H30)
 GUID_DEVCLASS_SMARTCARDREADER = iid
End Function
Public Function GUID_DEVCLASS_SMRDISK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53487C23, &H680F, &H4585, &HAC, &HC3, &H1F, &H10, &HD6, &H77, &H7E, &H82)
 GUID_DEVCLASS_SMRDISK = iid
End Function
Public Function GUID_DEVCLASS_SMRVOLUME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53B3CF03, &H8F5A, &H4788, &H91, &HB6, &HD1, &H9E, &HD9, &HFC, &HCC, &HBF)
 GUID_DEVCLASS_SMRVOLUME = iid
End Function
Public Function GUID_DEVCLASS_SOFTWARECOMPONENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5C4C3332, &H344D, &H483C, &H87, &H39, &H25, &H9E, &H93, &H4C, &H9C, &HC8)
 GUID_DEVCLASS_SOFTWARECOMPONENT = iid
End Function
Public Function GUID_DEVCLASS_SOUND() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E97C, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_SOUND = iid
End Function
Public Function GUID_DEVCLASS_SYSTEM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E97D, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_SYSTEM = iid
End Function
Public Function GUID_DEVCLASS_TAPEDRIVE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6D807884, &H7D21, &H11CF, &H80, &H1C, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_TAPEDRIVE = iid
End Function
Public Function GUID_DEVCLASS_UNKNOWN() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E97E, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
 GUID_DEVCLASS_UNKNOWN = iid
End Function
Public Function GUID_DEVCLASS_UCM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE6F1AA1C, &H7F3B, &H4473, &HB2, &HE8, &HC9, &H7D, &H8A, &HC7, &H1D, &H53)
 GUID_DEVCLASS_UCM = iid
End Function
Public Function GUID_DEVCLASS_USB() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H36FC9E60, &HC465, &H11CF, &H80, &H56, &H44, &H45, &H53, &H54, &H0, &H0)
 GUID_DEVCLASS_USB = iid
End Function
Public Function GUID_DEVCLASS_VOLUME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H71A27CDD, &H812A, &H11D0, &HBE, &HC7, &H8, &H0, &H2B, &HE2, &H9, &H2F)
 GUID_DEVCLASS_VOLUME = iid
End Function
Public Function GUID_DEVCLASS_VOLUMESNAPSHOT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H533C5B84, &HEC70, &H11D2, &H95, &H5, &H0, &HC0, &H4F, &H79, &HDE, &HAF)
 GUID_DEVCLASS_VOLUMESNAPSHOT = iid
End Function
Public Function GUID_DEVCLASS_WCEUSBS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H25DBCE51, &H6C8F, &H4A72, &H8A, &H6D, &HB5, &H4C, &H2B, &H4F, &HC8, &H35)
 GUID_DEVCLASS_WCEUSBS = iid
End Function
Public Function GUID_DEVCLASS_WPD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEEC5AD98, &H8080, &H425F, &H92, &H2A, &HDA, &HBF, &H3D, &HE3, &HF6, &H9A)
 GUID_DEVCLASS_WPD = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_TOP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB369BAF4, &H5568, &H4E82, &HA8, &H7E, &HA9, &H3E, &HB1, &H6B, &HCA, &H87)
 GUID_DEVCLASS_FSFILTER_TOP = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_ACTIVITYMONITOR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB86DFF51, &HA31E, &H4BAC, &HB3, &HCF, &HE8, &HCF, &HE7, &H5C, &H9F, &HC2)
 GUID_DEVCLASS_FSFILTER_ACTIVITYMONITOR = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_UNDELETE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFE8F1572, &HC67A, &H48C0, &HBB, &HAC, &HB, &H5C, &H6D, &H66, &HCA, &HFB)
 GUID_DEVCLASS_FSFILTER_UNDELETE = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_ANTIVIRUS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB1D1A169, &HC54F, &H4379, &H81, &HDB, &HBE, &HE7, &HD8, &H8D, &H74, &H54)
 GUID_DEVCLASS_FSFILTER_ANTIVIRUS = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_REPLICATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H48D3EBC4, &H4CF8, &H48FF, &HB8, &H69, &H9C, &H68, &HAD, &H42, &HEB, &H9F)
 GUID_DEVCLASS_FSFILTER_REPLICATION = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_CONTINUOUSBACKUP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H71AA14F8, &H6FAD, &H4622, &HAD, &H77, &H92, &HBB, &H9D, &H7E, &H69, &H47)
 GUID_DEVCLASS_FSFILTER_CONTINUOUSBACKUP = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_CONTENTSCREENER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3E3F0674, &HC83C, &H4558, &HBB, &H26, &H98, &H20, &HE1, &HEB, &HA5, &HC5)
 GUID_DEVCLASS_FSFILTER_CONTENTSCREENER = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_QUOTAMANAGEMENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8503C911, &HA6C7, &H4919, &H8F, &H79, &H50, &H28, &HF5, &H86, &H6B, &HC)
 GUID_DEVCLASS_FSFILTER_QUOTAMANAGEMENT = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_SYSTEMRECOVERY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2DB15374, &H706E, &H4131, &HA0, &HC7, &HD7, &HC7, &H8E, &HB0, &H28, &H9A)
 GUID_DEVCLASS_FSFILTER_SYSTEMRECOVERY = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_CFSMETADATASERVER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCDCF0939, &HB75B, &H4630, &HBF, &H76, &H80, &HF7, &HBA, &H65, &H58, &H84)
 GUID_DEVCLASS_FSFILTER_CFSMETADATASERVER = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_HSM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD546500A, &H2AEB, &H45F6, &H94, &H82, &HF4, &HB1, &H79, &H9C, &H31, &H77)
 GUID_DEVCLASS_FSFILTER_HSM = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_COMPRESSION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF3586BAF, &HB5AA, &H49B5, &H8D, &H6C, &H5, &H69, &H28, &H4C, &H63, &H9F)
 GUID_DEVCLASS_FSFILTER_COMPRESSION = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_ENCRYPTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA0A701C0, &HA511, &H42FF, &HAA, &H6C, &H6, &HDC, &H3, &H95, &H57, &H6F)
 GUID_DEVCLASS_FSFILTER_ENCRYPTION = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_VIRTUALIZATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF75A86C0, &H10D8, &H4C3A, &HB2, &H33, &HED, &H60, &HE4, &HCD, &HFA, &HAC)
 GUID_DEVCLASS_FSFILTER_VIRTUALIZATION = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_PHYSICALQUOTAMANAGEMENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6A0A8E78, &HBBA6, &H4FC4, &HA7, &H9, &H1E, &H33, &HCD, &H9, &HD6, &H7E)
 GUID_DEVCLASS_FSFILTER_PHYSICALQUOTAMANAGEMENT = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_OPENFILEBACKUP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF8ECAFA6, &H66D1, &H41A5, &H89, &H9B, &H66, &H58, &H5D, &H72, &H16, &HB7)
 GUID_DEVCLASS_FSFILTER_OPENFILEBACKUP = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_SECURITYENHANCER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD02BC3DA, &HC8E, &H4945, &H9B, &HD5, &HF1, &H88, &H3C, &H22, &H6C, &H8C)
 GUID_DEVCLASS_FSFILTER_SECURITYENHANCER = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_COPYPROTECTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H89786FF1, &H9C12, &H402F, &H9C, &H9E, &H17, &H75, &H3C, &H7F, &H43, &H75)
 GUID_DEVCLASS_FSFILTER_COPYPROTECTION = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_BOTTOM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H37765EA0, &H5958, &H4FC9, &HB0, &H4B, &H2F, &HDF, &HEF, &H97, &HE5, &H9E)
 GUID_DEVCLASS_FSFILTER_BOTTOM = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_SYSTEM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5D1B9AAA, &H1E2, &H46AF, &H84, &H9F, &H27, &H2B, &H3F, &H32, &H4C, &H46)
 GUID_DEVCLASS_FSFILTER_SYSTEM = iid
End Function
Public Function GUID_DEVCLASS_FSFILTER_INFRASTRUCTURE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE55FA6F9, &H128C, &H4D04, &HAB, &HAB, &H63, &HC, &H74, &HB1, &H45, &H3A)
 GUID_DEVCLASS_FSFILTER_INFRASTRUCTURE = iid
End Function


Public Function GUID_DEVINTERFACE_DISPLAY_ADAPTER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5B45201D, &HF2F2, &H4F3B, &H85, &HBB, &H30, &HFF, &H1F, &H95, &H35, &H99)
GUID_DEVINTERFACE_DISPLAY_ADAPTER = iid
End Function
Public Function GUID_DEVINTERFACE_DISK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53F56307, &HB6BF, &H11D0, &H94, &HF2, &H0, &HA0, &HC9, &H1E, &HFB, &H8B)
GUID_DEVINTERFACE_DISK = iid
End Function
Public Function GUID_DEVINTERFACE_CDROM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53F56308, &HB6BF, &H11D0, &H94, &HF2, &H0, &HA0, &HC9, &H1E, &HFB, &H8B)
GUID_DEVINTERFACE_CDROM = iid
End Function
Public Function GUID_DEVINTERFACE_PARTITION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53F5630A, &HB6BF, &H11D0, &H94, &HF2, &H0, &HA0, &HC9, &H1E, &HFB, &H8B)
GUID_DEVINTERFACE_PARTITION = iid
End Function
Public Function GUID_DEVINTERFACE_TAPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53F5630B, &HB6BF, &H11D0, &H94, &HF2, &H0, &HA0, &HC9, &H1E, &HFB, &H8B)
GUID_DEVINTERFACE_TAPE = iid
End Function
Public Function GUID_DEVINTERFACE_WRITEONCEDISK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53F5630C, &HB6BF, &H11D0, &H94, &HF2, &H0, &HA0, &HC9, &H1E, &HFB, &H8B)
GUID_DEVINTERFACE_WRITEONCEDISK = iid
End Function
Public Function GUID_DEVINTERFACE_VOLUME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53F5630D, &HB6BF, &H11D0, &H94, &HF2, &H0, &HA0, &HC9, &H1E, &HFB, &H8B)
GUID_DEVINTERFACE_VOLUME = iid
End Function
Public Function GUID_DEVINTERFACE_MEDIUMCHANGER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53F56310, &HB6BF, &H11D0, &H94, &HF2, &H0, &HA0, &HC9, &H1E, &HFB, &H8B)
GUID_DEVINTERFACE_MEDIUMCHANGER = iid
End Function
Public Function GUID_DEVINTERFACE_FLOPPY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53F56311, &HB6BF, &H11D0, &H94, &HF2, &H0, &HA0, &HC9, &H1E, &HFB, &H8B)
GUID_DEVINTERFACE_FLOPPY = iid
End Function
Public Function GUID_DEVINTERFACE_CDCHANGER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53F56312, &HB6BF, &H11D0, &H94, &HF2, &H0, &HA0, &HC9, &H1E, &HFB, &H8B)
GUID_DEVINTERFACE_CDCHANGER = iid
End Function
Public Function GUID_DEVINTERFACE_STORAGEPORT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2ACCFE60, &HC130, &H11D2, &HB0, &H82, &H0, &HA0, &HC9, &H1E, &HFB, &H8B)
GUID_DEVINTERFACE_STORAGEPORT = iid
End Function
Public Function GUID_DEVINTERFACE_COMPORT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H86E0D1E0, &H8089, &H11D0, &H9C, &HE4, &H8, &H0, &H3E, &H30, &H1F, &H73)
GUID_DEVINTERFACE_COMPORT = iid
End Function
Public Function GUID_DEVINTERFACE_SERENUM_BUS_ENUMERATOR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D36E978, &HE325, &H11CE, &HBF, &HC1, &H8, &H0, &H2B, &HE1, &H3, &H18)
GUID_DEVINTERFACE_SERENUM_BUS_ENUMERATOR = iid
End Function
Public Function GUID_DEVINTERFACE_HID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D1E55B2, &HF16F, &H11CF, &H88, &HCB, &H0, &H11, &H11, &H0, &H0, &H30)
GUID_DEVINTERFACE_HID = iid
End Function
Public Function GUID_DEVINTERFACE_USB_HUB() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF18A0E88, &HC30C, &H11D0, &H88, &H15, &H0, &HA0, &HC9, &H6, &HBE, &HD8)
GUID_DEVINTERFACE_USB_HUB = iid
End Function
Public Function GUID_DEVINTERFACE_USB_DEVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA5DCBF10, &H6530, &H11D2, &H90, &H1F, &H0, &HC0, &H4F, &HB9, &H51, &HED)
GUID_DEVINTERFACE_USB_DEVICE = iid
End Function
Public Function GUID_DEVINTERFACE_USB_HOST_CONTROLLER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3ABF6F2D, &H71C4, &H462A, &H8A, &H92, &H1E, &H68, &H61, &HE6, &HAF, &H27)
GUID_DEVINTERFACE_USB_HOST_CONTROLLER = iid
End Function
Public Function GUID_USB_WMI_STD_DATA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4E623B20, &HCB14, &H11D1, &HB3, &H31, &H0, &HA0, &HC9, &H59, &HBB, &HD2)
GUID_USB_WMI_STD_DATA = iid
End Function
Public Function GUID_USB_WMI_STD_NOTIFICATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4E623B20, &HCB14, &H11D1, &HB3, &H31, &H0, &HA0, &HC9, &H59, &HBB, &HD2)
GUID_USB_WMI_STD_NOTIFICATION = iid
End Function
Public Function GUID_DEVINTERFACE_KEYBOARD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H884B96C3, &H56EF, &H11D1, &HBC, &H8C, &H0, &HA0, &HC9, &H14, &H5, &HDD)
GUID_DEVINTERFACE_KEYBOARD = iid
End Function
Public Function GUID_DEVINTERFACE_MOUSE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H378DE44C, &H56EF, &H11D1, &HBC, &H8C, &H0, &HA0, &HC9, &H14, &H5, &HDD)
GUID_DEVINTERFACE_MOUSE = iid
End Function
Public Function GUID_DEVINTERFACE_PARALLEL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H97F76EF0, &HF883, &H11D0, &HAF, &H1F, &H0, &H0, &HF8, &H0, &H84, &H5C)
GUID_DEVINTERFACE_PARALLEL = iid
End Function
Public Function GUID_DEVINTERFACE_PARCLASS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H811FC6A5, &HF728, &H11D0, &HA5, &H37, &H0, &H0, &HF8, &H75, &H3E, &HD1)
GUID_DEVINTERFACE_PARCLASS = iid
End Function


Public Sub FreeKnownFolderDefinitionFields(pKFD As KNOWNFOLDER_DEFINITION)
Call CoTaskMemFree(pKFD.pszName)
Call CoTaskMemFree(pKFD.pszDescription)
Call CoTaskMemFree(pKFD.pszRelativePath)
Call CoTaskMemFree(pKFD.pszParsingName)
Call CoTaskMemFree(pKFD.pszToolTip)
Call CoTaskMemFree(pKFD.pszLocalizedName)
Call CoTaskMemFree(pKFD.pszIcon)
Call CoTaskMemFree(pKFD.pszSecurity)
End Sub
