Attribute VB_Name = "mWIC"
Option Explicit

'-----------------------------------------------------------------------------------
'mWIC.bas - Part of oleexp
'
'This module contains IIDs, GUIDs, and CLSIDs for working with the Windows Imaging Component set of COM interfaces.
'
'Revision 1
' -Added CLSID_WICWebpDecoder
'
'-----------------------------------------------------------------------------------

Private Sub DEFINE_UUID(Name As UUID, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
  With Name
    .Data1 = L
    .Data2 = w1
    .Data3 = w2
    .Data4(0) = B0
    .Data4(1) = b1
    .Data4(2) = b2
    .Data4(3) = B3
    .Data4(4) = b4
    .Data4(5) = b5
    .Data4(6) = b6
    .Data4(7) = b7
  End With
End Sub
Public Function GUID_WICPixelFormatUndefined() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H0)
GUID_WICPixelFormatUndefined = iid
End Function
Public Function GUID_WICPixelFormatDontCare() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H0)
GUID_WICPixelFormatDontCare = iid
End Function
Public Function GUID_WICPixelFormat1bppIndexed() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H1)
GUID_WICPixelFormat1bppIndexed = iid
End Function
Public Function GUID_WICPixelFormat2bppIndexed() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H2)
GUID_WICPixelFormat2bppIndexed = iid
End Function
Public Function GUID_WICPixelFormat4bppIndexed() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H3)
GUID_WICPixelFormat4bppIndexed = iid
End Function
Public Function GUID_WICPixelFormat8bppIndexed() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H4)
GUID_WICPixelFormat8bppIndexed = iid
End Function
Public Function GUID_WICPixelFormatBlackWhite() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H5)
GUID_WICPixelFormatBlackWhite = iid
End Function
Public Function GUID_WICPixelFormat2bppGray() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H6)
GUID_WICPixelFormat2bppGray = iid
End Function
Public Function GUID_WICPixelFormat4bppGray() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H7)
GUID_WICPixelFormat4bppGray = iid
End Function
Public Function GUID_WICPixelFormat8bppGray() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H8)
GUID_WICPixelFormat8bppGray = iid
End Function
Public Function GUID_WICPixelFormat8bppAlpha() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE6CD0116, &HEEBA, &H4161, &HAA, &H85, &H27, &HDD, &H9F, &HB3, &HA8, &H95)
GUID_WICPixelFormat8bppAlpha = iid
End Function
Public Function GUID_WICPixelFormat16bppBGR555() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H9)
GUID_WICPixelFormat16bppBGR555 = iid
End Function
Public Function GUID_WICPixelFormat16bppBGR565() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &HA)
GUID_WICPixelFormat16bppBGR565 = iid
End Function
Public Function GUID_WICPixelFormat16bppBGRA5551() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5EC7C2B, &HF1E6, &H4961, &HAD, &H46, &HE1, &HCC, &H81, &HA, &H87, &HD2)
GUID_WICPixelFormat16bppBGRA5551 = iid
End Function
Public Function GUID_WICPixelFormat16bppGray() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &HB)
GUID_WICPixelFormat16bppGray = iid
End Function
Public Function GUID_WICPixelFormat24bppBGR() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &HC)
GUID_WICPixelFormat24bppBGR = iid
End Function
Public Function GUID_WICPixelFormat24bppRGB() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &HD)
GUID_WICPixelFormat24bppRGB = iid
End Function
Public Function GUID_WICPixelFormat32bppBGR() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &HE)
GUID_WICPixelFormat32bppBGR = iid
End Function
Public Function GUID_WICPixelFormat32bppBGRA() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &HF)
GUID_WICPixelFormat32bppBGRA = iid
End Function
Public Function GUID_WICPixelFormat32bppPBGRA() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H10)
GUID_WICPixelFormat32bppPBGRA = iid
End Function
Public Function GUID_WICPixelFormat32bppGrayFloat() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H11)
GUID_WICPixelFormat32bppGrayFloat = iid
End Function
Public Function GUID_WICPixelFormat32bppRGB() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD98C6B95, &H3EFE, &H47D6, &HBB, &H25, &HEB, &H17, &H48, &HAB, &HC, &HF1)
GUID_WICPixelFormat32bppRGB = iid
End Function
Public Function GUID_WICPixelFormat32bppRGBA() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF5C7AD2D, &H6A8D, &H43DD, &HA7, &HA8, &HA2, &H99, &H35, &H26, &H1A, &HE9)
GUID_WICPixelFormat32bppRGBA = iid
End Function
Public Function GUID_WICPixelFormat32bppPRGBA() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3CC4A650, &HA527, &H4D37, &HA9, &H16, &H31, &H42, &HC7, &HEB, &HED, &HBA)
GUID_WICPixelFormat32bppPRGBA = iid
End Function
Public Function GUID_WICPixelFormat48bppRGB() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H15)
GUID_WICPixelFormat48bppRGB = iid
End Function
Public Function GUID_WICPixelFormat48bppBGR() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE605A384, &HB468, &H46CE, &HBB, &H2E, &H36, &HF1, &H80, &HE6, &H43, &H13)
GUID_WICPixelFormat48bppBGR = iid
End Function
Public Function GUID_WICPixelFormat64bppRGB() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA1182111, &H186D, &H4D42, &HBC, &H6A, &H9C, &H83, &H3, &HA8, &HDF, &HF9)
GUID_WICPixelFormat64bppRGB = iid
End Function
Public Function GUID_WICPixelFormat64bppRGBA() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H16)
GUID_WICPixelFormat64bppRGBA = iid
End Function
Public Function GUID_WICPixelFormat64bppBGRA() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1562FF7C, &HD352, &H46F9, &H97, &H9E, &H42, &H97, &H6B, &H79, &H22, &H46)
GUID_WICPixelFormat64bppBGRA = iid
End Function
Public Function GUID_WICPixelFormat64bppPRGBA() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H17)
GUID_WICPixelFormat64bppPRGBA = iid
End Function
Public Function GUID_WICPixelFormat64bppPBGRA() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8C518E8E, &HA4EC, &H468B, &HAE, &H70, &HC9, &HA3, &H5A, &H9C, &H55, &H30)
GUID_WICPixelFormat64bppPBGRA = iid
End Function
Public Function GUID_WICPixelFormat16bppGrayFixedPoint() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H13)
GUID_WICPixelFormat16bppGrayFixedPoint = iid
End Function
Public Function GUID_WICPixelFormat32bppBGR101010() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H14)
GUID_WICPixelFormat32bppBGR101010 = iid
End Function
Public Function GUID_WICPixelFormat48bppRGBFixedPoint() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H12)
GUID_WICPixelFormat48bppRGBFixedPoint = iid
End Function
Public Function GUID_WICPixelFormat48bppBGRFixedPoint() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H49CA140E, &HCAB6, &H493B, &H9D, &HDF, &H60, &H18, &H7C, &H37, &H53, &H2A)
GUID_WICPixelFormat48bppBGRFixedPoint = iid
End Function
Public Function GUID_WICPixelFormat96bppRGBFixedPoint() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H18)
GUID_WICPixelFormat96bppRGBFixedPoint = iid
End Function
Public Function GUID_WICPixelFormat96bppRGBFloat() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE3FED78F, &HE8DB, &H4ACF, &H84, &HC1, &HE9, &H7F, &H61, &H36, &HB3, &H27)
GUID_WICPixelFormat96bppRGBFloat = iid
End Function
Public Function GUID_WICPixelFormat128bppRGBAFloat() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H19)
GUID_WICPixelFormat128bppRGBAFloat = iid
End Function
Public Function GUID_WICPixelFormat128bppPRGBAFloat() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H1A)
GUID_WICPixelFormat128bppPRGBAFloat = iid
End Function
Public Function GUID_WICPixelFormat128bppRGBFloat() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H1B)
GUID_WICPixelFormat128bppRGBFloat = iid
End Function
Public Function GUID_WICPixelFormat32bppCMYK() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H1C)
GUID_WICPixelFormat32bppCMYK = iid
End Function
Public Function GUID_WICPixelFormat64bppRGBAFixedPoint() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H1D)
GUID_WICPixelFormat64bppRGBAFixedPoint = iid
End Function
Public Function GUID_WICPixelFormat64bppBGRAFixedPoint() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H356DE33C, &H54D2, &H4A23, &HBB, &H4, &H9B, &H7B, &HF9, &HB1, &HD4, &H2D)
GUID_WICPixelFormat64bppBGRAFixedPoint = iid
End Function
Public Function GUID_WICPixelFormat64bppRGBFixedPoint() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H40)
GUID_WICPixelFormat64bppRGBFixedPoint = iid
End Function
Public Function GUID_WICPixelFormat128bppRGBAFixedPoint() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H1E)
GUID_WICPixelFormat128bppRGBAFixedPoint = iid
End Function
Public Function GUID_WICPixelFormat128bppRGBFixedPoint() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H41)
GUID_WICPixelFormat128bppRGBFixedPoint = iid
End Function
Public Function GUID_WICPixelFormat64bppRGBAHalf() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H3A)
GUID_WICPixelFormat64bppRGBAHalf = iid
End Function
Public Function GUID_WICPixelFormat64bppPRGBAHalf() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H58AD26C2, &HC623, &H4D9D, &HB3, &H20, &H38, &H7E, &H49, &HF8, &HC4, &H42)
GUID_WICPixelFormat64bppPRGBAHalf = iid
End Function
Public Function GUID_WICPixelFormat64bppRGBHalf() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H42)
GUID_WICPixelFormat64bppRGBHalf = iid
End Function
Public Function GUID_WICPixelFormat48bppRGBHalf() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H3B)
GUID_WICPixelFormat48bppRGBHalf = iid
End Function
Public Function GUID_WICPixelFormat32bppRGBE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H3D)
GUID_WICPixelFormat32bppRGBE = iid
End Function
Public Function GUID_WICPixelFormat16bppGrayHalf() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H3E)
GUID_WICPixelFormat16bppGrayHalf = iid
End Function
Public Function GUID_WICPixelFormat32bppGrayFixedPoint() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H3F)
GUID_WICPixelFormat32bppGrayFixedPoint = iid
End Function
Public Function GUID_WICPixelFormat32bppRGBA1010102() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H25238D72, &HFCF9, &H4522, &HB5, &H14, &H55, &H78, &HE5, &HAD, &H55, &HE0)
GUID_WICPixelFormat32bppRGBA1010102 = iid
End Function
Public Function GUID_WICPixelFormat32bppRGBA1010102XR() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDE6B9A, &HC101, &H434B, &HB5, &H2, &HD0, &H16, &H5E, &HE1, &H12, &H2C)
GUID_WICPixelFormat32bppRGBA1010102XR = iid
End Function
Public Function GUID_WICPixelFormat64bppCMYK() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H1F)
GUID_WICPixelFormat64bppCMYK = iid
End Function
Public Function GUID_WICPixelFormat24bpp3Channels() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H20)
GUID_WICPixelFormat24bpp3Channels = iid
End Function
Public Function GUID_WICPixelFormat32bpp4Channels() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H21)
GUID_WICPixelFormat32bpp4Channels = iid
End Function
Public Function GUID_WICPixelFormat40bpp5Channels() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H22)
GUID_WICPixelFormat40bpp5Channels = iid
End Function
Public Function GUID_WICPixelFormat48bpp6Channels() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H23)
GUID_WICPixelFormat48bpp6Channels = iid
End Function
Public Function GUID_WICPixelFormat56bpp7Channels() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H24)
GUID_WICPixelFormat56bpp7Channels = iid
End Function
Public Function GUID_WICPixelFormat64bpp8Channels() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H25)
GUID_WICPixelFormat64bpp8Channels = iid
End Function
Public Function GUID_WICPixelFormat48bpp3Channels() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H26)
GUID_WICPixelFormat48bpp3Channels = iid
End Function
Public Function GUID_WICPixelFormat64bpp4Channels() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H27)
GUID_WICPixelFormat64bpp4Channels = iid
End Function
Public Function GUID_WICPixelFormat80bpp5Channels() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H28)
GUID_WICPixelFormat80bpp5Channels = iid
End Function
Public Function GUID_WICPixelFormat96bpp6Channels() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H29)
GUID_WICPixelFormat96bpp6Channels = iid
End Function
Public Function GUID_WICPixelFormat112bpp7Channels() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H2A)
GUID_WICPixelFormat112bpp7Channels = iid
End Function
Public Function GUID_WICPixelFormat128bpp8Channels() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H2B)
GUID_WICPixelFormat128bpp8Channels = iid
End Function
Public Function GUID_WICPixelFormat40bppCMYKAlpha() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H2C)
GUID_WICPixelFormat40bppCMYKAlpha = iid
End Function
Public Function GUID_WICPixelFormat80bppCMYKAlpha() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H2D)
GUID_WICPixelFormat80bppCMYKAlpha = iid
End Function
Public Function GUID_WICPixelFormat32bpp3ChannelsAlpha() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H2E)
GUID_WICPixelFormat32bpp3ChannelsAlpha = iid
End Function
Public Function GUID_WICPixelFormat40bpp4ChannelsAlpha() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H2F)
GUID_WICPixelFormat40bpp4ChannelsAlpha = iid
End Function
Public Function GUID_WICPixelFormat48bpp5ChannelsAlpha() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H30)
GUID_WICPixelFormat48bpp5ChannelsAlpha = iid
End Function
Public Function GUID_WICPixelFormat56bpp6ChannelsAlpha() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H31)
GUID_WICPixelFormat56bpp6ChannelsAlpha = iid
End Function
Public Function GUID_WICPixelFormat64bpp7ChannelsAlpha() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H32)
GUID_WICPixelFormat64bpp7ChannelsAlpha = iid
End Function
Public Function GUID_WICPixelFormat72bpp8ChannelsAlpha() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H33)
GUID_WICPixelFormat72bpp8ChannelsAlpha = iid
End Function
Public Function GUID_WICPixelFormat64bpp3ChannelsAlpha() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H34)
GUID_WICPixelFormat64bpp3ChannelsAlpha = iid
End Function
Public Function GUID_WICPixelFormat80bpp4ChannelsAlpha() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H35)
GUID_WICPixelFormat80bpp4ChannelsAlpha = iid
End Function
Public Function GUID_WICPixelFormat96bpp5ChannelsAlpha() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H36)
GUID_WICPixelFormat96bpp5ChannelsAlpha = iid
End Function
Public Function GUID_WICPixelFormat112bpp6ChannelsAlpha() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H37)
GUID_WICPixelFormat112bpp6ChannelsAlpha = iid
End Function
Public Function GUID_WICPixelFormat128bpp7ChannelsAlpha() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H38)
GUID_WICPixelFormat128bpp7ChannelsAlpha = iid
End Function
Public Function GUID_WICPixelFormat144bpp8ChannelsAlpha() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H39)
GUID_WICPixelFormat144bpp8ChannelsAlpha = iid
End Function
Public Function GUID_WICPixelFormat8bppY() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H91B4DB54, &H2DF9, &H42F0, &HB4, &H49, &H29, &H9, &HBB, &H3D, &HF8, &H8E)
GUID_WICPixelFormat8bppY = iid
End Function
Public Function GUID_WICPixelFormat8bppCb() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1339F224, &H6BFE, &H4C3E, &H93, &H2, &HE4, &HF3, &HA6, &HD0, &HCA, &H2A)
GUID_WICPixelFormat8bppCb = iid
End Function
Public Function GUID_WICPixelFormat8bppCr() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB8145053, &H2116, &H49F0, &H88, &H35, &HED, &H84, &H4B, &H20, &H5C, &H51)
GUID_WICPixelFormat8bppCr = iid
End Function
Public Function GUID_WICPixelFormat16bppCbCr() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFF95BA6E, &H11E0, &H4263, &HBB, &H45, &H1, &H72, &H1F, &H34, &H60, &HA4)
GUID_WICPixelFormat16bppCbCr = iid
End Function
Public Function GUID_WICPixelFormat16bppYQuantizedDctCoefficients() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA355F433, &H48E8, &H4A42, &H84, &HD8, &HE2, &HAA, &H26, &HCA, &H80, &HA4)
GUID_WICPixelFormat16bppYQuantizedDctCoefficients = iid
End Function
Public Function GUID_WICPixelFormat16bppCbQuantizedDctCoefficients() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD2C4FF61, &H56A5, &H49C2, &H8B, &H5C, &H4C, &H19, &H25, &H96, &H48, &H37)
GUID_WICPixelFormat16bppCbQuantizedDctCoefficients = iid
End Function
Public Function GUID_WICPixelFormat16bppCrQuantizedDctCoefficients() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2FE354F0, &H1680, &H42D8, &H92, &H31, &HE7, &H3C, &H5, &H65, &HBF, &HC1)
GUID_WICPixelFormat16bppCrQuantizedDctCoefficients = iid
End Function

Public Function CLSID_WICImagingFactory() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCACAF262, &H9370, &H4615, &HA1, &H3B, &H9F, &H55, &H39, &HDA, &H4C, &HA)
CLSID_WICImagingFactory = iid
End Function
Public Function CLSID_WICImagingFactory1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCACAF262, &H9370, &H4615, &HA1, &H3B, &H9F, &H55, &H39, &HDA, &H4C, &HA)
CLSID_WICImagingFactory1 = iid
End Function
Public Function CLSID_WICImagingFactory2() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H317D06E8, &H5F24, &H433D, &HBD, &HF7, &H79, &HCE, &H68, &HD8, &HAB, &HC2)
CLSID_WICImagingFactory2 = iid
End Function
Public Function CLSID_WICPngDecoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H389EA17B, &H5078, &H4CDE, &HB6, &HEF, &H25, &HC1, &H51, &H75, &HC7, &H51)
CLSID_WICPngDecoder = iid
End Function
Public Function CLSID_WICPngDecoder1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H389EA17B, &H5078, &H4CDE, &HB6, &HEF, &H25, &HC1, &H51, &H75, &HC7, &H51)
CLSID_WICPngDecoder1 = iid
End Function
Public Function CLSID_WICPngDecoder2() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE018945B, &HAA86, &H4008, &H9B, &HD4, &H67, &H77, &HA1, &HE4, &HC, &H11)
CLSID_WICPngDecoder2 = iid
End Function
Public Function CLSID_WICBmpDecoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6B462062, &H7CBF, &H400D, &H9F, &HDB, &H81, &H3D, &HD1, &HF, &H27, &H78)
CLSID_WICBmpDecoder = iid
End Function
Public Function CLSID_WICIcoDecoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC61BFCDF, &H2E0F, &H4AAD, &HA8, &HD7, &HE0, &H6B, &HAF, &HEB, &HCD, &HFE)
CLSID_WICIcoDecoder = iid
End Function
Public Function CLSID_WICJpegDecoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9456A480, &HE88B, &H43EA, &H9E, &H73, &HB, &H2D, &H9B, &H71, &HB1, &HCA)
CLSID_WICJpegDecoder = iid
End Function
Public Function CLSID_WICGifDecoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H381DDA3C, &H9CE9, &H4834, &HA2, &H3E, &H1F, &H98, &HF8, &HFC, &H52, &HBE)
CLSID_WICGifDecoder = iid
End Function
Public Function CLSID_WICTiffDecoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB54E85D9, &HFE23, &H499F, &H8B, &H88, &H6A, &HCE, &HA7, &H13, &H75, &H2B)
CLSID_WICTiffDecoder = iid
End Function
Public Function CLSID_WICWebpDecoder() As UUID
'{7693e886-51c9-4070-8419-9f70738ec8fa}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7693E886, CInt(&H51C9), CInt(&H4070), &H84, &H19, &H9F, &H70, &H73, &H8E, &HC8, &HFA)
CLSID_WICWebpDecoder = iid
End Function
Public Function CLSID_WICWmpDecoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA26CEC36, &H234C, &H4950, &HAE, &H16, &HE3, &H4A, &HAC, &HE7, &H1D, &HD)
CLSID_WICWmpDecoder = iid
End Function
Public Function CLSID_WICDdsDecoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9053699F, &HA341, &H429D, &H9E, &H90, &HEE, &H43, &H7C, &HF8, &HC, &H73)
CLSID_WICDdsDecoder = iid
End Function
Public Function CLSID_WICBmpEncoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H69BE8BB4, &HD66D, &H47C8, &H86, &H5A, &HED, &H15, &H89, &H43, &H37, &H82)
CLSID_WICBmpEncoder = iid
End Function
Public Function CLSID_WICPngEncoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H27949969, &H876A, &H41D7, &H94, &H47, &H56, &H8F, &H6A, &H35, &HA4, &HDC)
CLSID_WICPngEncoder = iid
End Function
Public Function CLSID_WICJpegEncoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1A34F5C1, &H4A5A, &H46DC, &HB6, &H44, &H1F, &H45, &H67, &HE7, &HA6, &H76)
CLSID_WICJpegEncoder = iid
End Function
Public Function CLSID_WICGifEncoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H114F5598, &HB22, &H40A0, &H86, &HA1, &HC8, &H3E, &HA4, &H95, &HAD, &HBD)
CLSID_WICGifEncoder = iid
End Function
Public Function CLSID_WICTiffEncoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H131BE10, &H2001, &H4C5F, &HA9, &HB0, &HCC, &H88, &HFA, &HB6, &H4C, &HE8)
CLSID_WICTiffEncoder = iid
End Function
Public Function CLSID_WICWmpEncoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAC4CE3CB, &HE1C1, &H44CD, &H82, &H15, &H5A, &H16, &H65, &H50, &H9E, &HC2)
CLSID_WICWmpEncoder = iid
End Function
Public Function CLSID_WICDdsEncoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA61DDE94, &H66CE, &H4AC1, &H88, &H1B, &H71, &H68, &H5, &H88, &H89, &H5E)
CLSID_WICDdsEncoder = iid
End Function
Public Function CLSID_WICAdngDecoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H981D9411, &H909E, &H42A7, &H8F, &H5D, &HA7, &H47, &HFF, &H5, &H2E, &HDB)
CLSID_WICAdngDecoder = iid
End Function
Public Function CLSID_WICJpegQualcommPhoneEncoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H68ED5C62, &HF534, &H4979, &HB2, &HB3, &H68, &H6A, &H12, &HB2, &HB3, &H4C)
CLSID_WICJpegQualcommPhoneEncoder = iid
End Function
Public Function GUID_VendorMicrosoft() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF0E749CA, &HEDEF, &H4589, &HA7, &H3A, &HEE, &HE, &H62, &H6A, &H2A, &H2B)
GUID_VendorMicrosoft = iid
End Function
Public Function GUID_VendorMicrosoftBuiltIn() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H257A30FD, &H6B6, &H462B, &HAE, &HA4, &H63, &HF7, &HB, &H86, &HE5, &H33)
GUID_VendorMicrosoftBuiltIn = iid
End Function
Public Function GUID_ContainerFormatBmp() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAF1D87E, &HFCFE, &H4188, &HBD, &HEB, &HA7, &H90, &H64, &H71, &HCB, &HE3)
GUID_ContainerFormatBmp = iid
End Function
Public Function GUID_ContainerFormatPng() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1B7CFAF4, &H713F, &H473C, &HBB, &HCD, &H61, &H37, &H42, &H5F, &HAE, &HAF)
GUID_ContainerFormatPng = iid
End Function
Public Function GUID_ContainerFormatIco() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA3A860C4, &H338F, &H4C17, &H91, &H9A, &HFB, &HA4, &HB5, &H62, &H8F, &H21)
GUID_ContainerFormatIco = iid
End Function
Public Function GUID_ContainerFormatJpeg() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H19E4A5AA, &H5662, &H4FC5, &HA0, &HC0, &H17, &H58, &H2, &H8E, &H10, &H57)
GUID_ContainerFormatJpeg = iid
End Function
Public Function GUID_ContainerFormatTiff() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H163BCC30, &HE2E9, &H4F0B, &H96, &H1D, &HA3, &HE9, &HFD, &HB7, &H88, &HA3)
GUID_ContainerFormatTiff = iid
End Function
Public Function GUID_ContainerFormatGif() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1F8A5601, &H7D4D, &H4CBD, &H9C, &H82, &H1B, &HC8, &HD4, &HEE, &HB9, &HA5)
GUID_ContainerFormatGif = iid
End Function
Public Function GUID_ContainerFormatWmp() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H57A37CAA, &H367A, &H4540, &H91, &H6B, &HF1, &H83, &HC5, &H9, &H3A, &H4B)
GUID_ContainerFormatWmp = iid
End Function
Public Function GUID_ContainerFormatDds() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9967CB95, &H2E85, &H4AC8, &H8C, &HA2, &H83, &HD7, &HCC, &HD4, &H25, &HC9)
GUID_ContainerFormatDds = iid
End Function
Public Function GUID_ContainerFormatAdng() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF3FF6D0D, &H38C0, &H41C4, &HB1, &HFE, &H1F, &H38, &H24, &HF1, &H7B, &H84)
GUID_ContainerFormatAdng = iid
End Function
Public Function GUID_ContainerFormatWebp() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE094B0E2, &H67F2, &H45B3, &HB0, &HEA, &H11, &H53, &H37, &HCA, &H7C, &HF3)
GUID_ContainerFormatWebp = iid
End Function
Public Function GUID_ContainerFormatHeif() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE1E62521, &H6787, &H405B, &HA3, &H39, &H50, &H7, &H15, &HB5, &H76, &H3F)
GUID_ContainerFormatHeif = iid
End Function
Public Function CLSID_WICImagingCategories() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFAE3D380, &HFEA4, &H4623, &H8C, &H75, &HC6, &HB6, &H11, &H10, &HB6, &H81)
CLSID_WICImagingCategories = iid
End Function
Public Function CATID_WICBitmapDecoders() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7ED96837, &H96F0, &H4812, &HB2, &H11, &HF1, &H3C, &H24, &H11, &H7E, &HD3)
CATID_WICBitmapDecoders = iid
End Function
Public Function CATID_WICBitmapEncoders() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAC757296, &H3522, &H4E11, &H98, &H62, &HC1, &H7B, &HE5, &HA1, &H76, &H7E)
CATID_WICBitmapEncoders = iid
End Function
Public Function CATID_WICPixelFormats() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2B46E70F, &HCDA7, &H473E, &H89, &HF6, &HDC, &H96, &H30, &HA2, &H39, &HB)
CATID_WICPixelFormats = iid
End Function
Public Function CATID_WICFormatConverters() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7835EAE8, &HBF14, &H49D1, &H93, &HCE, &H53, &H3A, &H40, &H7B, &H22, &H48)
CATID_WICFormatConverters = iid
End Function
Public Function CATID_WICMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5AF94D8, &H7174, &H4CD2, &HBE, &H4A, &H41, &H24, &HB8, &HE, &HE4, &HB8)
CATID_WICMetadataReader = iid
End Function
Public Function CATID_WICMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HABE3B9A4, &H257D, &H4B97, &HBD, &H1A, &H29, &H4A, &HF4, &H96, &H22, &H2E)
CATID_WICMetadataWriter = iid
End Function
Public Function CLSID_WICDefaultFormatConverter() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1A3F11DC, &HB514, &H4B17, &H8C, &H5F, &H21, &H54, &H51, &H38, &H52, &HF1)
CLSID_WICDefaultFormatConverter = iid
End Function
Public Function CLSID_WICFormatConverterHighColor() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAC75D454, &H9F37, &H48F8, &HB9, &H72, &H4E, &H19, &HBC, &H85, &H60, &H11)
CLSID_WICFormatConverterHighColor = iid
End Function
Public Function CLSID_WICFormatConverterNChannel() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC17CABB2, &HD4A3, &H47D7, &HA5, &H57, &H33, &H9B, &H2E, &HFB, &HD4, &HF1)
CLSID_WICFormatConverterNChannel = iid
End Function
Public Function CLSID_WICFormatConverterWMPhoto() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9CB5172B, &HD600, &H46BA, &HAB, &H77, &H77, &HBB, &H7E, &H3A, &H0, &HD9)
CLSID_WICFormatConverterWMPhoto = iid
End Function
Public Function CLSID_WICPlanarFormatConverter() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H184132B8, &H32F8, &H4784, &H91, &H31, &HDD, &H72, &H24, &HB2, &H34, &H38)
CLSID_WICPlanarFormatConverter = iid
End Function


Public Function IID_IWICPalette() As UUID
'{00000040-a8f2-4877-ba0a-fd2b6645fb94}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H40, CInt(&HA8F2), CInt(&H4877), &HBA, &HA, &HFD, &H2B, &H66, &H45, &HFB, &H94)
IID_IWICPalette = iid
End Function
Public Function IID_IWICBitmapSource() As UUID
'{00000120-a8f2-4877-ba0a-fd2b6645fb94}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H120, CInt(&HA8F2), CInt(&H4877), &HBA, &HA, &HFD, &H2B, &H66, &H45, &HFB, &H94)
IID_IWICBitmapSource = iid
End Function
Public Function IID_IWICFormatConverter() As UUID
'{00000301-a8f2-4877-ba0a-fd2b6645fb94}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H301, CInt(&HA8F2), CInt(&H4877), &HBA, &HA, &HFD, &H2B, &H66, &H45, &HFB, &H94)
IID_IWICFormatConverter = iid
End Function
Public Function IID_IWICPlanarFormatConverter() As UUID
'{BEBEE9CB-83B0-4DCC-8132-B0AAA55EAC96}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBEBEE9CB, CInt(&H83B0), CInt(&H4DCC), &H81, &H32, &HB0, &HAA, &HA5, &H5E, &HAC, &H96)
IID_IWICPlanarFormatConverter = iid
End Function
Public Function IID_IWICBitmapScaler() As UUID
'{00000302-a8f2-4877-ba0a-fd2b6645fb94}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H302, CInt(&HA8F2), CInt(&H4877), &HBA, &HA, &HFD, &H2B, &H66, &H45, &HFB, &H94)
IID_IWICBitmapScaler = iid
End Function
Public Function IID_IWICBitmapClipper() As UUID
'{E4FBCF03-223D-4e81-9333-D635556DD1B5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE4FBCF03, CInt(&H223D), CInt(&H4E81), &H93, &H33, &HD6, &H35, &H55, &H6D, &HD1, &HB5)
IID_IWICBitmapClipper = iid
End Function
Public Function IID_IWICBitmapFlipRotator() As UUID
'{5009834F-2D6A-41ce-9E1B-17C5AFF7A782}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5009834F, CInt(&H2D6A), CInt(&H41CE), &H9E, &H1B, &H17, &HC5, &HAF, &HF7, &HA7, &H82)
IID_IWICBitmapFlipRotator = iid
End Function
Public Function IID_IWICBitmapLock() As UUID
'{00000123-a8f2-4877-ba0a-fd2b6645fb94}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H123, CInt(&HA8F2), CInt(&H4877), &HBA, &HA, &HFD, &H2B, &H66, &H45, &HFB, &H94)
IID_IWICBitmapLock = iid
End Function
Public Function IID_IWICBitmap() As UUID
'{00000121-a8f2-4877-ba0a-fd2b6645fb94}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H121, CInt(&HA8F2), CInt(&H4877), &HBA, &HA, &HFD, &H2B, &H66, &H45, &HFB, &H94)
IID_IWICBitmap = iid
End Function
Public Function IID_IWICColorContext() As UUID
'{3C613A02-34B2-44ea-9A7C-45AEA9C6FD6D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3C613A02, CInt(&H34B2), CInt(&H44EA), &H9A, &H7C, &H45, &HAE, &HA9, &HC6, &HFD, &H6D)
IID_IWICColorContext = iid
End Function
Public Function IID_IWICColorTransform() As UUID
'{B66F034F-D0E2-40ab-B436-6DE39E321A94}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB66F034F, CInt(&HD0E2), CInt(&H40AB), &HB4, &H36, &H6D, &HE3, &H9E, &H32, &H1A, &H94)
IID_IWICColorTransform = iid
End Function
Public Function IID_IWICFastMetadataEncoder() As UUID
'{B84E2C09-78C9-4AC4-8BD3-524AE1663A2F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB84E2C09, CInt(&H78C9), CInt(&H4AC4), &H8B, &HD3, &H52, &H4A, &HE1, &H66, &H3A, &H2F)
IID_IWICFastMetadataEncoder = iid
End Function
Public Function IID_IWICStream() As UUID
'{135FF860-22B7-4ddf-B0F6-218F4F299A43}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H135FF860, CInt(&H22B7), CInt(&H4DDF), &HB0, &HF6, &H21, &H8F, &H4F, &H29, &H9A, &H43)
IID_IWICStream = iid
End Function
Public Function IID_IWICEnumMetadataItem() As UUID
'{DC2BB46D-3F07-481E-8625-220C4AEDBB33}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDC2BB46D, CInt(&H3F07), CInt(&H481E), &H86, &H25, &H22, &HC, &H4A, &HED, &HBB, &H33)
IID_IWICEnumMetadataItem = iid
End Function
Public Function IID_IWICMetadataQueryReader() As UUID
'{30989668-E1C9-4597-B395-458EEDB808DF}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H30989668, CInt(&HE1C9), CInt(&H4597), &HB3, &H95, &H45, &H8E, &HED, &HB8, &H8, &HDF)
IID_IWICMetadataQueryReader = iid
End Function
Public Function IID_IWICMetadataQueryWriter() As UUID
'{A721791A-0DEF-4d06-BD91-2118BF1DB10B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA721791A, CInt(&HDEF), CInt(&H4D06), &HBD, &H91, &H21, &H18, &HBF, &H1D, &HB1, &HB)
IID_IWICMetadataQueryWriter = iid
End Function
Public Function IID_IWICBitmapEncoder() As UUID
'{00000103-a8f2-4877-ba0a-fd2b6645fb94}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H103, CInt(&HA8F2), CInt(&H4877), &HBA, &HA, &HFD, &H2B, &H66, &H45, &HFB, &H94)
IID_IWICBitmapEncoder = iid
End Function
Public Function IID_IWICBitmapFrameEncode() As UUID
'{00000105-a8f2-4877-ba0a-fd2b6645fb94}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H105, CInt(&HA8F2), CInt(&H4877), &HBA, &HA, &HFD, &H2B, &H66, &H45, &HFB, &H94)
IID_IWICBitmapFrameEncode = iid
End Function
Public Function IID_IWICPlanarBitmapFrameEncode() As UUID
'{F928B7B8-2221-40C1-B72E-7E82F1974D1A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF928B7B8, CInt(&H2221), CInt(&H40C1), &HB7, &H2E, &H7E, &H82, &HF1, &H97, &H4D, &H1A)
IID_IWICPlanarBitmapFrameEncode = iid
End Function
Public Function IID_IWICImageEncoder() As UUID
'{04C75BF8-3CE1-473B-ACC5-3CC4F5E94999}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4C75BF8, CInt(&H3CE1), CInt(&H473B), &HAC, &HC5, &H3C, &HC4, &HF5, &HE9, &H49, &H99)
IID_IWICImageEncoder = iid
End Function
Public Function IID_IWICBitmapDecoder() As UUID
'{9EDDE9E7-8DEE-47ea-99DF-E6FAF2ED44BF}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9EDDE9E7, CInt(&H8DEE), CInt(&H47EA), &H99, &HDF, &HE6, &HFA, &HF2, &HED, &H44, &HBF)
IID_IWICBitmapDecoder = iid
End Function
Public Function IID_IWICBitmapSourceTransform() As UUID
'{3B16811B-6A43-4ec9-B713-3D5A0C13B940}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3B16811B, CInt(&H6A43), CInt(&H4EC9), &HB7, &H13, &H3D, &H5A, &HC, &H13, &HB9, &H40)
IID_IWICBitmapSourceTransform = iid
End Function
Public Function IID_IWICPlanarBitmapSourceTransform() As UUID
'{3AFF9CCE-BE95-4303-B927-E7D16FF4A613}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3AFF9CCE, CInt(&HBE95), CInt(&H4303), &HB9, &H27, &HE7, &HD1, &H6F, &HF4, &HA6, &H13)
IID_IWICPlanarBitmapSourceTransform = iid
End Function
Public Function IID_IWICBitmapFrameDecode() As UUID
'{3B16811B-6A43-4ec9-A813-3D930C13B940}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3B16811B, CInt(&H6A43), CInt(&H4EC9), &HA8, &H13, &H3D, &H93, &HC, &H13, &HB9, &H40)
IID_IWICBitmapFrameDecode = iid
End Function
Public Function IID_IWICProgressiveLevelControl() As UUID
'{DAAC296F-7AA5-4dbf-8D15-225C5976F891}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDAAC296F, CInt(&H7AA5), CInt(&H4DBF), &H8D, &H15, &H22, &H5C, &H59, &H76, &HF8, &H91)
IID_IWICProgressiveLevelControl = iid
End Function
Public Function IID_IWICProgressCallback() As UUID
'{4776F9CD-9517-45FA-BF24-E89C5EC5C60C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4776F9CD, CInt(&H9517), CInt(&H45FA), &HBF, &H24, &HE8, &H9C, &H5E, &HC5, &HC6, &HC)
IID_IWICProgressCallback = iid
End Function
Public Function IID_IWICBitmapCodecProgressNotification() As UUID
'{64C1024E-C3CF-4462-8078-88C2B11C46D9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H64C1024E, CInt(&HC3CF), CInt(&H4462), &H80, &H78, &H88, &HC2, &HB1, &H1C, &H46, &HD9)
IID_IWICBitmapCodecProgressNotification = iid
End Function
Public Function IID_IWICComponentInfo() As UUID
'{23BC3F0A-698B-4357-886B-F24D50671334}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H23BC3F0A, CInt(&H698B), CInt(&H4357), &H88, &H6B, &HF2, &H4D, &H50, &H67, &H13, &H34)
IID_IWICComponentInfo = iid
End Function
Public Function IID_IWICFormatConverterInfo() As UUID
'{9F34FB65-13F4-4f15-BC57-3726B5E53D9F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9F34FB65, CInt(&H13F4), CInt(&H4F15), &HBC, &H57, &H37, &H26, &HB5, &HE5, &H3D, &H9F)
IID_IWICFormatConverterInfo = iid
End Function
Public Function IID_IWICBitmapCodecInfo() As UUID
'{E87A44C4-B76E-4c47-8B09-298EB12A2714}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE87A44C4, CInt(&HB76E), CInt(&H4C47), &H8B, &H9, &H29, &H8E, &HB1, &H2A, &H27, &H14)
IID_IWICBitmapCodecInfo = iid
End Function
Public Function IID_IWICBitmapEncoderInfo() As UUID
'{94C9B4EE-A09F-4f92-8A1E-4A9BCE7E76FB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H94C9B4EE, CInt(&HA09F), CInt(&H4F92), &H8A, &H1E, &H4A, &H9B, &HCE, &H7E, &H76, &HFB)
IID_IWICBitmapEncoderInfo = iid
End Function
Public Function IID_IWICBitmapDecoderInfo() As UUID
'{D8CD007F-D08F-4191-9BFC-236EA7F0E4B5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD8CD007F, CInt(&HD08F), CInt(&H4191), &H9B, &HFC, &H23, &H6E, &HA7, &HF0, &HE4, &HB5)
IID_IWICBitmapDecoderInfo = iid
End Function
Public Function IID_IWICPixelFormatInfo() As UUID
'{E8EDA601-3D48-431a-AB44-69059BE88BBE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE8EDA601, CInt(&H3D48), CInt(&H431A), &HAB, &H44, &H69, &H5, &H9B, &HE8, &H8B, &HBE)
IID_IWICPixelFormatInfo = iid
End Function
Public Function IID_IWICPixelFormatInfo2() As UUID
'{A9DB33A2-AF5F-43C7-B679-74F5984B5AA4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA9DB33A2, CInt(&HAF5F), CInt(&H43C7), &HB6, &H79, &H74, &HF5, &H98, &H4B, &H5A, &HA4)
IID_IWICPixelFormatInfo2 = iid
End Function
Public Function IID_IWICImagingFactory() As UUID
'{ec5ec8a9-c395-4314-9c77-54d7a935ff70}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEC5EC8A9, CInt(&HC395), CInt(&H4314), &H9C, &H77, &H54, &HD7, &HA9, &H35, &HFF, &H70)
IID_IWICImagingFactory = iid
End Function
Public Function IID_IWICImagingFactory2() As UUID
'{7B816B45-1996-4476-B132-DE9E247C8AF0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7B816B45, CInt(&H1996), CInt(&H4476), &HB1, &H32, &HDE, &H9E, &H24, &H7C, &H8A, &HF0)
IID_IWICImagingFactory2 = iid
End Function
Public Function IID_IWICDevelopRawNotificationCallback() As UUID
'{95c75a6e-3e8c-4ec2-85a8-aebcc551e59b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H95C75A6E, CInt(&H3E8C), CInt(&H4EC2), &H85, &HA8, &HAE, &HBC, &HC5, &H51, &HE5, &H9B)
IID_IWICDevelopRawNotificationCallback = iid
End Function
Public Function IID_IWICDevelopRaw() As UUID
'{fbec5e44-f7be-4b65-b7f8-c0c81fef026d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFBEC5E44, CInt(&HF7BE), CInt(&H4B65), &HB7, &HF8, &HC0, &HC8, &H1F, &HEF, &H2, &H6D)
IID_IWICDevelopRaw = iid
End Function
Public Function IID_IWICDdsDecoder() As UUID
'{409cd537-8532-40cb-9774-e2feb2df4e9c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H409CD537, CInt(&H8532), CInt(&H40CB), &H97, &H74, &HE2, &HFE, &HB2, &HDF, &H4E, &H9C)
IID_IWICDdsDecoder = iid
End Function
Public Function IID_IWICDdsEncoder() As UUID
'{5cacdb4c-407e-41b3-b936-d0f010cd6732}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5CACDB4C, CInt(&H407E), CInt(&H41B3), &HB9, &H36, &HD0, &HF0, &H10, &HCD, &H67, &H32)
IID_IWICDdsEncoder = iid
End Function
Public Function IID_IWICDdsFrameDecode() As UUID
'{3d4c0c61-18a4-41e4-bd80-481a4fc9f464}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3D4C0C61, CInt(&H18A4), CInt(&H41E4), &HBD, &H80, &H48, &H1A, &H4F, &HC9, &HF4, &H64)
IID_IWICDdsFrameDecode = iid
End Function
Public Function IID_IWICJpegFrameDecode() As UUID
'{8939F66E-C46A-4c21-A9D1-98B327CE1679}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8939F66E, CInt(&HC46A), CInt(&H4C21), &HA9, &HD1, &H98, &HB3, &H27, &HCE, &H16, &H79)
IID_IWICJpegFrameDecode = iid
End Function
Public Function IID_IWICJpegFrameEncode() As UUID
'{2F0C601F-D2C6-468C-ABFA-49495D983ED1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2F0C601F, CInt(&HD2C6), CInt(&H468C), &HAB, &HFA, &H49, &H49, &H5D, &H98, &H3E, &HD1)
IID_IWICJpegFrameEncode = iid
End Function
Public Function IID_IWICMetadataBlockReader() As UUID
'{FEAA2A8D-B3F3-43E4-B25C-D1DE990A1AE1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFEAA2A8D, CInt(&HB3F3), CInt(&H43E4), &HB2, &H5C, &HD1, &HDE, &H99, &HA, &H1A, &HE1)
IID_IWICMetadataBlockReader = iid
End Function
Public Function IID_IWICMetadataBlockWriter() As UUID
'{08FB9676-B444-41E8-8DBE-6A53A542BFF1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8FB9676, CInt(&HB444), CInt(&H41E8), &H8D, &HBE, &H6A, &H53, &HA5, &H42, &HBF, &HF1)
IID_IWICMetadataBlockWriter = iid
End Function
Public Function IID_IWICMetadataReader() As UUID
'{9204FE99-D8FC-4FD5-A001-9536B067A899}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9204FE99, CInt(&HD8FC), CInt(&H4FD5), &HA0, &H1, &H95, &H36, &HB0, &H67, &HA8, &H99)
IID_IWICMetadataReader = iid
End Function
Public Function IID_IWICMetadataWriter() As UUID
'{F7836E16-3BE0-470B-86BB-160D0AECD7DE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF7836E16, CInt(&H3BE0), CInt(&H470B), &H86, &HBB, &H16, &HD, &HA, &HEC, &HD7, &HDE)
IID_IWICMetadataWriter = iid
End Function
Public Function IID_IWICStreamProvider() As UUID
'{449494BC-B468-4927-96D7-BA90D31AB505}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H449494BC, CInt(&HB468), CInt(&H4927), &H96, &HD7, &HBA, &H90, &HD3, &H1A, &HB5, &H5)
IID_IWICStreamProvider = iid
End Function
Public Function IID_IWICPersistStream() As UUID
'{00675040-6908-45F8-86A3-49C7DFD6D9AD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H675040, CInt(&H6908), CInt(&H45F8), &H86, &HA3, &H49, &HC7, &HDF, &HD6, &HD9, &HAD)
IID_IWICPersistStream = iid
End Function
Public Function IID_IWICMetadataHandlerInfo() As UUID
'{ABA958BF-C672-44D1-8D61-CE6DF2E682C2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HABA958BF, CInt(&HC672), CInt(&H44D1), &H8D, &H61, &HCE, &H6D, &HF2, &HE6, &H82, &HC2)
IID_IWICMetadataHandlerInfo = iid
End Function
Public Function IID_IWICMetadataReaderInfo() As UUID
'{EEBF1F5B-07C1-4447-A3AB-22ACAF78A804}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEEBF1F5B, CInt(&H7C1), CInt(&H4447), &HA3, &HAB, &H22, &HAC, &HAF, &H78, &HA8, &H4)
IID_IWICMetadataReaderInfo = iid
End Function
Public Function IID_IWICMetadataWriterInfo() As UUID
'{B22E3FBA-3925-4323-B5C1-9EBFC430F236}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB22E3FBA, CInt(&H3925), CInt(&H4323), &HB5, &HC1, &H9E, &HBF, &HC4, &H30, &HF2, &H36)
IID_IWICMetadataWriterInfo = iid
End Function
Public Function IID_IWICComponentFactory() As UUID
'{412D0C3A-9650-44FA-AF5B-DD2A06C8E8FB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H412D0C3A, CInt(&H9650), CInt(&H44FA), &HAF, &H5B, &HDD, &H2A, &H6, &HC8, &HE8, &HFB)
IID_IWICComponentFactory = iid
End Function

Public Function GUID_MetadataFormatUnknown() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA45E592F, &H9078, &H4A7C, &HAD, &HB5, &H4E, &HDC, &H4F, &HD6, &H1B, &H1F)
GUID_MetadataFormatUnknown = iid
End Function
Public Function GUID_MetadataFormatIfd() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H537396C6, &H2D8A, &H4BB6, &H9B, &HF8, &H2F, &HA, &H8E, &H2A, &H3A, &HDF)
GUID_MetadataFormatIfd = iid
End Function
Public Function GUID_MetadataFormatSubIfd() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H58A2E128, &H2DB9, &H4E57, &HBB, &H14, &H51, &H77, &H89, &H1E, &HD3, &H31)
GUID_MetadataFormatSubIfd = iid
End Function
Public Function GUID_MetadataFormatExif() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1C3C4F9D, &HB84A, &H467D, &H94, &H93, &H36, &HCF, &HBD, &H59, &HEA, &H57)
GUID_MetadataFormatExif = iid
End Function
Public Function GUID_MetadataFormatGps() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7134AB8A, &H9351, &H44AD, &HAF, &H62, &H44, &H8D, &HB6, &HB5, &H2, &HEC)
GUID_MetadataFormatGps = iid
End Function
Public Function GUID_MetadataFormatInterop() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HED686F8E, &H681F, &H4C8B, &HBD, &H41, &HA8, &HAD, &HDB, &HF6, &HB3, &HFC)
GUID_MetadataFormatInterop = iid
End Function
Public Function GUID_MetadataFormatApp0() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H79007028, &H268D, &H45D6, &HA3, &HC2, &H35, &H4E, &H6A, &H50, &H4B, &HC9)
GUID_MetadataFormatApp0 = iid
End Function
Public Function GUID_MetadataFormatApp1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8FD3DFC3, &HF951, &H492B, &H81, &H7F, &H69, &HC2, &HE6, &HD9, &HA5, &HB0)
GUID_MetadataFormatApp1 = iid
End Function
Public Function GUID_MetadataFormatApp13() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H326556A2, &HF502, &H4354, &H9C, &HC0, &H8E, &H3F, &H48, &HEA, &HF6, &HB5)
GUID_MetadataFormatApp13 = iid
End Function
Public Function GUID_MetadataFormatIPTC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4FAB0914, &HE129, &H4087, &HA1, &HD1, &HBC, &H81, &H2D, &H45, &HA7, &HB5)
GUID_MetadataFormatIPTC = iid
End Function
Public Function GUID_MetadataFormatIRB() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H16100D66, &H8570, &H4BB9, &HB9, &H2D, &HFD, &HA4, &HB2, &H3E, &HCE, &H67)
GUID_MetadataFormatIRB = iid
End Function
Public Function GUID_MetadataFormat8BIMIPTC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H10568C, &H852, &H4E6A, &HB1, &H91, &H5C, &H33, &HAC, &H5B, &H4, &H30)
GUID_MetadataFormat8BIMIPTC = iid
End Function
Public Function GUID_MetadataFormat8BIMResolutionInfo() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H739F305D, &H81DB, &H43CB, &HAC, &H5E, &H55, &H1, &H3E, &HF9, &HF0, &H3)
GUID_MetadataFormat8BIMResolutionInfo = iid
End Function
Public Function GUID_MetadataFormat8BIMIPTCDigest() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1CA32285, &H9CCD, &H4786, &H8B, &HD8, &H79, &H53, &H9D, &HB6, &HA0, &H6)
GUID_MetadataFormat8BIMIPTCDigest = iid
End Function
Public Function GUID_MetadataFormatXMP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBB5ACC38, &HF216, &H4CEC, &HA6, &HC5, &H5F, &H6E, &H73, &H97, &H63, &HA9)
GUID_MetadataFormatXMP = iid
End Function
Public Function GUID_MetadataFormatThumbnail() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H243DCEE9, &H8703, &H40EE, &H8E, &HF0, &H22, &HA6, &H0, &HB8, &H5, &H8C)
GUID_MetadataFormatThumbnail = iid
End Function
Public Function GUID_MetadataFormatChunktEXt() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H568D8936, &HC0A9, &H4923, &H90, &H5D, &HDF, &H2B, &H38, &H23, &H8F, &HBC)
GUID_MetadataFormatChunktEXt = iid
End Function
Public Function GUID_MetadataFormatXMPStruct() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H22383CF1, &HED17, &H4E2E, &HAF, &H17, &HD8, &H5B, &H8F, &H6B, &H30, &HD0)
GUID_MetadataFormatXMPStruct = iid
End Function
Public Function GUID_MetadataFormatXMPBag() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H833CCA5F, &HDCB7, &H4516, &H80, &H6F, &H65, &H96, &HAB, &H26, &HDC, &HE4)
GUID_MetadataFormatXMPBag = iid
End Function
Public Function GUID_MetadataFormatXMPSeq() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H63E8DF02, &HEB6C, &H456C, &HA2, &H24, &HB2, &H5E, &H79, &H4F, &HD6, &H48)
GUID_MetadataFormatXMPSeq = iid
End Function
Public Function GUID_MetadataFormatXMPAlt() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7B08A675, &H91AA, &H481B, &HA7, &H98, &H4D, &HA9, &H49, &H8, &H61, &H3B)
GUID_MetadataFormatXMPAlt = iid
End Function
Public Function GUID_MetadataFormatLSD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE256031E, &H6299, &H4929, &HB9, &H8D, &H5A, &HC8, &H84, &HAF, &HBA, &H92)
GUID_MetadataFormatLSD = iid
End Function
Public Function GUID_MetadataFormatIMD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBD2BB086, &H4D52, &H48DD, &H96, &H77, &HDB, &H48, &H3E, &H85, &HAE, &H8F)
GUID_MetadataFormatIMD = iid
End Function
Public Function GUID_MetadataFormatGCE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2A25CAD8, &HDEEB, &H4C69, &HA7, &H88, &HE, &HC2, &H26, &H6D, &HCA, &HFD)
GUID_MetadataFormatGCE = iid
End Function
Public Function GUID_MetadataFormatAPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2E043DC2, &HC967, &H4E05, &H87, &H5E, &H61, &H8B, &HF6, &H7E, &H85, &HC3)
GUID_MetadataFormatAPE = iid
End Function
Public Function GUID_MetadataFormatJpegChrominance() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF73D0DCF, &HCEC6, &H4F85, &H9B, &HE, &H1C, &H39, &H56, &HB1, &HBE, &HF7)
GUID_MetadataFormatJpegChrominance = iid
End Function
Public Function GUID_MetadataFormatJpegLuminance() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H86908007, &HEDFC, &H4860, &H8D, &H4B, &H4E, &HE6, &HE8, &H3E, &H60, &H58)
GUID_MetadataFormatJpegLuminance = iid
End Function
Public Function GUID_MetadataFormatJpegComment() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H220E5F33, &HAFD3, &H474E, &H9D, &H31, &H7D, &H4F, &HE7, &H30, &HF5, &H57)
GUID_MetadataFormatJpegComment = iid
End Function
Public Function GUID_MetadataFormatGifComment() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC4B6E0E0, &HCFB4, &H4AD3, &HAB, &H33, &H9A, &HAD, &H23, &H55, &HA3, &H4A)
GUID_MetadataFormatGifComment = iid
End Function
Public Function GUID_MetadataFormatChunkgAMA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF00935A5, &H1D5D, &H4CD1, &H81, &HB2, &H93, &H24, &HD7, &HEC, &HA7, &H81)
GUID_MetadataFormatChunkgAMA = iid
End Function
Public Function GUID_MetadataFormatChunkbKGD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE14D3571, &H6B47, &H4DEA, &HB6, &HA, &H87, &HCE, &HA, &H78, &HDF, &HB7)
GUID_MetadataFormatChunkbKGD = iid
End Function
Public Function GUID_MetadataFormatChunkiTXt() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC2BEC729, &HB68, &H4B77, &HAA, &HE, &H62, &H95, &HA6, &HAC, &H18, &H14)
GUID_MetadataFormatChunkiTXt = iid
End Function
Public Function GUID_MetadataFormatChunkcHRM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9DB3655B, &H2842, &H44B3, &H80, &H67, &H12, &HE9, &HB3, &H75, &H55, &H6A)
GUID_MetadataFormatChunkcHRM = iid
End Function
Public Function GUID_MetadataFormatChunkhIST() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC59A82DA, &HDB74, &H48A4, &HBD, &H6A, &HB6, &H9C, &H49, &H31, &HEF, &H95)
GUID_MetadataFormatChunkhIST = iid
End Function
Public Function GUID_MetadataFormatChunkiCCP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEB4349AB, &HB685, &H450F, &H91, &HB5, &HE8, &H2, &HE8, &H92, &H53, &H6C)
GUID_MetadataFormatChunkiCCP = iid
End Function
Public Function GUID_MetadataFormatChunksRGB() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC115FD36, &HCC6F, &H4E3F, &H83, &H63, &H52, &H4B, &H87, &HC6, &HB0, &HD9)
GUID_MetadataFormatChunksRGB = iid
End Function
Public Function GUID_MetadataFormatChunktIME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6B00AE2D, &HE24B, &H460A, &H98, &HB6, &H87, &H8B, &HD0, &H30, &H72, &HFD)
GUID_MetadataFormatChunktIME = iid
End Function
Public Function GUID_MetadataFormatDds() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4A064603, &H8C33, &H4E60, &H9C, &H29, &H13, &H62, &H31, &H70, &H2D, &H8)
GUID_MetadataFormatDds = iid
End Function
Public Function GUID_MetadataFormatHeif() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H817EF3E1, &H1288, &H45F4, &HA8, &H52, &H26, &HD, &H9E, &H7C, &HCE, &H83)
GUID_MetadataFormatHeif = iid
End Function
Public Function GUID_MetadataFormatHeifHDR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(GUID_MetadataFormatHeifHDR, &H568B8D8A, &H1E65, &H438C, &H89, &H68, &HD6, &HE, &H10, &H12, &HBE, &HB9)
GUID_MetadataFormatHeifHDR = iid
End Function
Public Function GUID_MetadataFormatWebpANIM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6DC4FDA6, &H78E6, &H4102, &HAE, &H35, &HBC, &HFA, &H1E, &HDC, &HC7, &H8B)
GUID_MetadataFormatWebpANIM = iid
End Function
Public Function GUID_MetadataFormatWebpANMF() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43C105EE, &HB93B, &H4ABB, &HB0, &H3, &HA0, &H8C, &HD, &H87, &H4, &H71)
GUID_MetadataFormatWebpANMF = iid
End Function


Public Function CLSID_WICUnknownMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H699745C2, &H5066, &H4B82, &HA8, &HE3, &HD4, &H4, &H78, &HDB, &HEC, &H8C)
CLSID_WICUnknownMetadataReader = iid
End Function
Public Function CLSID_WICUnknownMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA09CCA86, &H27BA, &H4F39, &H90, &H53, &H12, &H1F, &HA4, &HDC, &H8, &HFC)
CLSID_WICUnknownMetadataWriter = iid
End Function
Public Function CLSID_WICApp0MetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF3C633A2, &H46C8, &H498E, &H8F, &HBB, &HCC, &H6F, &H72, &H1B, &HBC, &HDE)
CLSID_WICApp0MetadataWriter = iid
End Function
Public Function CLSID_WICApp0MetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43324B33, &HA78F, &H480F, &H91, &H11, &H96, &H38, &HAA, &HCC, &HC8, &H32)
CLSID_WICApp0MetadataReader = iid
End Function
Public Function CLSID_WICApp1MetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEE366069, &H1832, &H420F, &HB3, &H81, &H4, &H79, &HAD, &H6, &H6F, &H19)
CLSID_WICApp1MetadataWriter = iid
End Function
Public Function CLSID_WICApp1MetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDDE33513, &H774E, &H4BCD, &HAE, &H79, &H2, &HF4, &HAD, &HFE, &H62, &HFC)
CLSID_WICApp1MetadataReader = iid
End Function
Public Function CLSID_WICApp13MetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7B19A919, &HA9D6, &H49E5, &HBD, &H45, &H2, &HC3, &H4E, &H4E, &H4C, &HD5)
CLSID_WICApp13MetadataWriter = iid
End Function
Public Function CLSID_WICApp13MetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAA7E3C50, &H864C, &H4604, &HBC, &H4, &H8B, &HB, &H76, &HE6, &H37, &HF6)
CLSID_WICApp13MetadataReader = iid
End Function
Public Function CLSID_WICIfdMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8F914656, &H9D0A, &H4EB2, &H90, &H19, &HB, &HF9, &H6D, &H8A, &H9E, &HE6)
CLSID_WICIfdMetadataReader = iid
End Function
Public Function CLSID_WICIfdMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB1EBFC28, &HC9BD, &H47A2, &H8D, &H33, &HB9, &H48, &H76, &H97, &H77, &HA7)
CLSID_WICIfdMetadataWriter = iid
End Function
Public Function CLSID_WICSubIfdMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H50D42F09, &HECD1, &H4B41, &HB6, &H5D, &HDA, &H1F, &HDA, &HA7, &H56, &H63)
CLSID_WICSubIfdMetadataReader = iid
End Function
Public Function CLSID_WICSubIfdMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8ADE5386, &H8E9B, &H4F4C, &HAC, &HF2, &HF0, &H0, &H87, &H6, &HB2, &H38)
CLSID_WICSubIfdMetadataWriter = iid
End Function
Public Function CLSID_WICExifMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD9403860, &H297F, &H4A49, &HBF, &H9B, &H77, &H89, &H81, &H50, &HA4, &H42)
CLSID_WICExifMetadataReader = iid
End Function
Public Function CLSID_WICExifMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC9A14CDA, &HC339, &H460B, &H90, &H78, &HD4, &HDE, &HBC, &HFA, &HBE, &H91)
CLSID_WICExifMetadataWriter = iid
End Function
Public Function CLSID_WICGpsMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3697790B, &H223B, &H484E, &H99, &H25, &HC4, &H86, &H92, &H18, &HF1, &H7A)
CLSID_WICGpsMetadataReader = iid
End Function
Public Function CLSID_WICGpsMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCB8C13E4, &H62B5, &H4C96, &HA4, &H8B, &H6B, &HA6, &HAC, &HE3, &H9C, &H76)
CLSID_WICGpsMetadataWriter = iid
End Function
Public Function CLSID_WICInteropMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB5C8B898, &H74, &H459F, &HB7, &H0, &H86, &HD, &H46, &H51, &HEA, &H14)
CLSID_WICInteropMetadataReader = iid
End Function
Public Function CLSID_WICInteropMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H122EC645, &HCD7E, &H44D8, &HB1, &H86, &H2C, &H8C, &H20, &HC3, &HB5, &HF)
CLSID_WICInteropMetadataWriter = iid
End Function
Public Function CLSID_WICThumbnailMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFB012959, &HF4F6, &H44D7, &H9D, &H9, &HDA, &HA0, &H87, &HA9, &HDB, &H57)
CLSID_WICThumbnailMetadataReader = iid
End Function
Public Function CLSID_WICThumbnailMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD049B20C, &H5DD0, &H44FE, &HB0, &HB3, &H8F, &H92, &HC8, &HE6, &HD0, &H80)
CLSID_WICThumbnailMetadataWriter = iid
End Function
Public Function CLSID_WICIPTCMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3012959, &HF4F6, &H44D7, &H9D, &H9, &HDA, &HA0, &H87, &HA9, &HDB, &H57)
CLSID_WICIPTCMetadataReader = iid
End Function
Public Function CLSID_WICIPTCMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1249B20C, &H5DD0, &H44FE, &HB0, &HB3, &H8F, &H92, &HC8, &HE6, &HD0, &H80)
CLSID_WICIPTCMetadataWriter = iid
End Function
Public Function CLSID_WICIRBMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD4DCD3D7, &HB4C2, &H47D9, &HA6, &HBF, &HB8, &H9B, &HA3, &H96, &HA4, &HA3)
CLSID_WICIRBMetadataReader = iid
End Function
Public Function CLSID_WICIRBMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5C5C1935, &H235, &H4434, &H80, &HBC, &H25, &H1B, &HC1, &HEC, &H39, &HC6)
CLSID_WICIRBMetadataWriter = iid
End Function
Public Function CLSID_WIC8BIMIPTCMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H10668C, &H801, &H4DA6, &HA4, &HA4, &H82, &H65, &H22, &HB6, &HD2, &H8F)
CLSID_WIC8BIMIPTCMetadataReader = iid
End Function
Public Function CLSID_WIC8BIMIPTCMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H108226, &HEE41, &H44A2, &H9E, &H9C, &H4B, &HE4, &HD5, &HB1, &HD2, &HCD)
CLSID_WIC8BIMIPTCMetadataWriter = iid
End Function
Public Function CLSID_WIC8BIMResolutionInfoMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5805137A, &HE348, &H4F7C, &HB3, &HCC, &H6D, &HB9, &H96, &H5A, &H5, &H99)
CLSID_WIC8BIMResolutionInfoMetadataReader = iid
End Function
Public Function CLSID_WIC8BIMResolutionInfoMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4FF2FE0E, &HE74A, &H4B71, &H98, &HC4, &HAB, &H7D, &HC1, &H67, &H7, &HBA)
CLSID_WIC8BIMResolutionInfoMetadataWriter = iid
End Function
Public Function CLSID_WIC8BIMIPTCDigestMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2805F1E, &HD5AA, &H415B, &H82, &HC5, &H61, &HC0, &H33, &HA9, &H88, &HA6)
CLSID_WIC8BIMIPTCDigestMetadataReader = iid
End Function
Public Function CLSID_WIC8BIMIPTCDigestMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2DB5E62B, &HD67, &H495F, &H8F, &H9D, &HC2, &HF0, &H18, &H86, &H47, &HAC)
CLSID_WIC8BIMIPTCDigestMetadataWriter = iid
End Function
Public Function CLSID_WICPngTextMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4B59AFCC, &HB8C3, &H408A, &HB6, &H70, &H89, &HE5, &HFA, &HB6, &HFD, &HA7)
CLSID_WICPngTextMetadataReader = iid
End Function
Public Function CLSID_WICPngTextMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB5EBAFB9, &H253E, &H4A72, &HA7, &H44, &H7, &H62, &HD2, &H68, &H56, &H83)
CLSID_WICPngTextMetadataWriter = iid
End Function
Public Function CLSID_WICXMPMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H72B624DF, &HAE11, &H4948, &HA6, &H5C, &H35, &H1E, &HB0, &H82, &H94, &H19)
CLSID_WICXMPMetadataReader = iid
End Function
Public Function CLSID_WICXMPMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1765E14E, &H1BD4, &H462E, &HB6, &HB1, &H59, &HB, &HF1, &H26, &H2A, &HC6)
CLSID_WICXMPMetadataWriter = iid
End Function
Public Function CLSID_WICXMPStructMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B90D9A, &H8209, &H47F7, &H9C, &H52, &HE1, &H24, &H4B, &HF5, &HC, &HED)
CLSID_WICXMPStructMetadataReader = iid
End Function
Public Function CLSID_WICXMPStructMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H22C21F93, &H7DDB, &H411C, &H9B, &H17, &HC5, &HB7, &HBD, &H6, &H4A, &HBC)
CLSID_WICXMPStructMetadataWriter = iid
End Function
Public Function CLSID_WICXMPBagMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE7E79A30, &H4F2C, &H4FAB, &H8D, &H0, &H39, &H4F, &H2D, &H6B, &HBE, &HBE)
CLSID_WICXMPBagMetadataReader = iid
End Function
Public Function CLSID_WICXMPBagMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HED822C8C, &HD6BE, &H4301, &HA6, &H31, &HE, &H14, &H16, &HBA, &HD2, &H8F)
CLSID_WICXMPBagMetadataWriter = iid
End Function
Public Function CLSID_WICXMPSeqMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7F12E753, &HFC71, &H43D7, &HA5, &H1D, &H92, &HF3, &H59, &H77, &HAB, &HB5)
CLSID_WICXMPSeqMetadataReader = iid
End Function
Public Function CLSID_WICXMPSeqMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6D68D1DE, &HD432, &H4B0F, &H92, &H3A, &H9, &H11, &H83, &HA9, &HBD, &HA7)
CLSID_WICXMPSeqMetadataWriter = iid
End Function
Public Function CLSID_WICXMPAltMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAA94DCC2, &HB8B0, &H4898, &HB8, &H35, &H0, &HA, &HAB, &HD7, &H43, &H93)
CLSID_WICXMPAltMetadataReader = iid
End Function
Public Function CLSID_WICXMPAltMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H76C2A6C, &HF78F, &H4C46, &HA7, &H23, &H35, &H83, &HE7, &H8, &H76, &HEA)
CLSID_WICXMPAltMetadataWriter = iid
End Function
Public Function CLSID_WICLSDMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H41070793, &H59E4, &H479A, &HA1, &HF7, &H95, &H4A, &HDC, &H2E, &HF5, &HFC)
CLSID_WICLSDMetadataReader = iid
End Function
Public Function CLSID_WICLSDMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H73C037E7, &HE5D9, &H4954, &H87, &H6A, &H6D, &HA8, &H1D, &H6E, &H57, &H68)
CLSID_WICLSDMetadataWriter = iid
End Function
Public Function CLSID_WICGCEMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB92E345D, &HF52D, &H41F3, &HB5, &H62, &H8, &H1B, &HC7, &H72, &HE3, &HB9)
CLSID_WICGCEMetadataReader = iid
End Function
Public Function CLSID_WICGCEMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAF95DC76, &H16B2, &H47F4, &HB3, &HEA, &H3C, &H31, &H79, &H66, &H93, &HE7)
CLSID_WICGCEMetadataWriter = iid
End Function
Public Function CLSID_WICIMDMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7447A267, &H15, &H42C8, &HA8, &HF1, &HFB, &H3B, &H94, &HC6, &H83, &H61)
CLSID_WICIMDMetadataReader = iid
End Function
Public Function CLSID_WICIMDMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8C89071F, &H452E, &H4E95, &H96, &H82, &H9D, &H10, &H24, &H62, &H71, &H72)
CLSID_WICIMDMetadataWriter = iid
End Function
Public Function CLSID_WICAPEMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1767B93A, &HB021, &H44EA, &H92, &HF, &H86, &H3C, &H11, &HF4, &HF7, &H68)
CLSID_WICAPEMetadataReader = iid
End Function
Public Function CLSID_WICAPEMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBD6EDFCA, &H2890, &H482F, &HB2, &H33, &H8D, &H73, &H39, &HA1, &HCF, &H8D)
CLSID_WICAPEMetadataWriter = iid
End Function
Public Function CLSID_WICJpegChrominanceMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H50B1904B, &HF28F, &H4574, &H93, &HF4, &HB, &HAD, &HE8, &H2C, &H69, &HE9)
CLSID_WICJpegChrominanceMetadataReader = iid
End Function
Public Function CLSID_WICJpegChrominanceMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3FF566F0, &H6E6B, &H49D4, &H96, &HE6, &HB7, &H88, &H86, &H69, &H2C, &H62)
CLSID_WICJpegChrominanceMetadataWriter = iid
End Function
Public Function CLSID_WICJpegLuminanceMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H356F2F88, &H5A6, &H4728, &HB9, &HA4, &H1B, &HFB, &HCE, &H4, &HD8, &H38)
CLSID_WICJpegLuminanceMetadataReader = iid
End Function
Public Function CLSID_WICJpegLuminanceMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1D583ABC, &H8A0E, &H4657, &H99, &H82, &HA3, &H80, &HCA, &H58, &HFB, &H4B)
CLSID_WICJpegLuminanceMetadataWriter = iid
End Function
Public Function CLSID_WICJpegCommentMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9F66347C, &H60C4, &H4C4D, &HAB, &H58, &HD2, &H35, &H86, &H85, &HF6, &H7)
CLSID_WICJpegCommentMetadataReader = iid
End Function
Public Function CLSID_WICJpegCommentMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE573236F, &H55B1, &H4EDA, &H81, &HEA, &H9F, &H65, &HDB, &H2, &H90, &HD3)
CLSID_WICJpegCommentMetadataWriter = iid
End Function
Public Function CLSID_WICGifCommentMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H32557D3B, &H69DC, &H4F95, &H83, &H6E, &HF5, &H97, &H2B, &H2F, &H61, &H59)
CLSID_WICGifCommentMetadataReader = iid
End Function
Public Function CLSID_WICGifCommentMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA02797FC, &HC4AE, &H418C, &HAF, &H95, &HE6, &H37, &HC7, &HEA, &HD2, &HA1)
CLSID_WICGifCommentMetadataWriter = iid
End Function
Public Function CLSID_WICPngGamaMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3692CA39, &HE082, &H4350, &H9E, &H1F, &H37, &H4, &HCB, &H8, &H3C, &HD5)
CLSID_WICPngGamaMetadataReader = iid
End Function
Public Function CLSID_WICPngGamaMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFF036D13, &H5D4B, &H46DD, &HB1, &HF, &H10, &H66, &H93, &HD9, &HFE, &H4F)
CLSID_WICPngGamaMetadataWriter = iid
End Function
Public Function CLSID_WICPngBkgdMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCE7A4A6, &H3E8, &H4A60, &H9D, &H15, &H28, &H2E, &HF3, &H2E, &HE7, &HDA)
CLSID_WICPngBkgdMetadataReader = iid
End Function
Public Function CLSID_WICPngBkgdMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H68E3F2FD, &H31AE, &H4441, &HBB, &H6A, &HFD, &H70, &H47, &H52, &H5F, &H90)
CLSID_WICPngBkgdMetadataWriter = iid
End Function
Public Function CLSID_WICPngItxtMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAABFB2FA, &H3E1E, &H4A8F, &H89, &H77, &H55, &H56, &HFB, &H94, &HEA, &H23)
CLSID_WICPngItxtMetadataReader = iid
End Function
Public Function CLSID_WICPngItxtMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H31879719, &HE751, &H4DF8, &H98, &H1D, &H68, &HDF, &HF6, &H77, &H4, &HED)
CLSID_WICPngItxtMetadataWriter = iid
End Function
Public Function CLSID_WICPngChrmMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF90B5F36, &H367B, &H402A, &H9D, &HD1, &HBC, &HF, &HD5, &H9D, &H8F, &H62)
CLSID_WICPngChrmMetadataReader = iid
End Function
Public Function CLSID_WICPngChrmMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE23CE3EB, &H5608, &H4E83, &HBC, &HEF, &H27, &HB1, &H98, &H7E, &H51, &HD7)
CLSID_WICPngChrmMetadataWriter = iid
End Function
Public Function CLSID_WICPngHistMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H877A0BB7, &HA313, &H4491, &H87, &HB5, &H2E, &H6D, &H5, &H94, &HF5, &H20)
CLSID_WICPngHistMetadataReader = iid
End Function
Public Function CLSID_WICPngHistMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8A03E749, &H672E, &H446E, &HBF, &H1F, &H2C, &H11, &HD2, &H33, &HB6, &HFF)
CLSID_WICPngHistMetadataWriter = iid
End Function
Public Function CLSID_WICPngIccpMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF5D3E63B, &HCB0F, &H4628, &HA4, &H78, &H6D, &H82, &H44, &HBE, &H36, &HB1)
CLSID_WICPngIccpMetadataReader = iid
End Function
Public Function CLSID_WICPngIccpMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H16671E5F, &HCE6, &H4CC4, &H97, &H68, &HE8, &H9F, &HE5, &H1, &H8A, &HDE)
CLSID_WICPngIccpMetadataWriter = iid
End Function
Public Function CLSID_WICPngSrgbMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFB40360C, &H547E, &H4956, &HA3, &HB9, &HD4, &H41, &H88, &H59, &HBA, &H66)
CLSID_WICPngSrgbMetadataReader = iid
End Function
Public Function CLSID_WICPngSrgbMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA6EE35C6, &H87EC, &H47DF, &H9F, &H22, &H1D, &H5A, &HAD, &H84, &HC, &H82)
CLSID_WICPngSrgbMetadataWriter = iid
End Function
Public Function CLSID_WICPngTimeMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD94EDF02, &HEFE5, &H4F0D, &H85, &HC8, &HF5, &HA6, &H8B, &H30, &H0, &HB1)
CLSID_WICPngTimeMetadataReader = iid
End Function
Public Function CLSID_WICPngTimeMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1AB78400, &HB5A3, &H4D91, &H8A, &HCE, &H33, &HFC, &HD1, &H49, &H9B, &HE6)
CLSID_WICPngTimeMetadataWriter = iid
End Function
Public Function CLSID_WICDdsMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H276C88CA, &H7533, &H4A86, &HB6, &H76, &H66, &HB3, &H60, &H80, &HD4, &H84)
CLSID_WICDdsMetadataReader = iid
End Function
Public Function CLSID_WICDdsMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFD688BBD, &H31ED, &H4DB7, &HA7, &H23, &H93, &H49, &H27, &HD3, &H83, &H67)
CLSID_WICDdsMetadataWriter = iid
End Function
Public Function CLSID_WICHeifMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HACDDFC3F, &H85EC, &H41BC, &HBD, &HEF, &H1B, &HC2, &H62, &HE4, &HDB, &H5)
CLSID_WICHeifMetadataReader = iid
End Function
Public Function CLSID_WICHeifMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3AE45E79, &H40BC, &H4401, &HAC, &HE5, &HDD, &H3C, &HB1, &H6E, &H6A, &HFE)
CLSID_WICHeifMetadataWriter = iid
End Function
Public Function CLSID_WICHeifHDRMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2438DE3D, &H94D9, &H4BE8, &H84, &HA8, &H4D, &HE9, &H5A, &H57, &H5E, &H75)
CLSID_WICHeifHDRMetadataReader = iid
End Function
Public Function CLSID_WICHeifHDRMetadataWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB83135A2, &H8E7E, &H485E, &HA5, &H33, &HF9, &H36, &H21, &HDD, &H93, &HC8)
CLSID_WICHeifHDRMetadataWriter = iid
End Function
Public Function CLSID_WICWebpAnimMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H76F9911, &HA348, &H465C, &HA8, &H7, &HA2, &H52, &HF3, &HF2, &HD3, &HDE)
CLSID_WICWebpAnimMetadataReader = iid
End Function
Public Function CLSID_WICWebpAnmfMetadataReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H85A10B03, &HC9F6, &H439F, &HBE, &H5E, &HC0, &HFB, &HEF, &H67, &H80, &H7C)
CLSID_WICWebpAnmfMetadataReader = iid
End Function

