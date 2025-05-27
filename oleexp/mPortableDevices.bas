Attribute VB_Name = "mPortableDevices"
Option Explicit

'-----------------------------------------------------------------------
'mPortableDevices.bas - Part of oleexp
'
'This module contains UUIDs, PROPERTYKEYs, and other helpers for Portable Devices COM interfaces
'
'-----------------------------------------------------------------------


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
Private Sub DEFINE_PROPERTYKEY(Name As PROPERTYKEY, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte, pid As Long)
  With Name.fmtid
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
  Name.pid = pid
End Sub

Public Sub FreePortableDevicePnPIDs(lIDs() As Long)
Dim i As Long
For i = 0 To UBound(lIDs)
    CoTaskMemFree lIDs(i)
Next i
End Sub

Public Function CLSID_PortableDeviceManager() As UUID
'{0AF10CEC-2ECD-4B92-9581-34F6AE0637F3}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAF10CEC, CInt(&H2ECD), CInt(&H4B92), &H95, &H81, &H34, &HF6, &HAE, &H6, &H37, &HF3)
 CLSID_PortableDeviceManager = iid
End Function
Public Function CLSID_PortableDevice() As UUID
'{728A21C5-3D9E-48D7-9810-864848F0F404}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H728A21C5, CInt(&H3D9E), CInt(&H48D7), &H98, &H10, &H86, &H48, &H48, &HF0, &HF4, &H4)
 CLSID_PortableDevice = iid
End Function
Public Function CLSID_PortableDeviceFTM() As UUID
'{F7C0039A-4762-488A-B4B3-760EF9A1BA9B}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF7C0039A, CInt(&H4762), CInt(&H488A), &HB4, &HB3, &H76, &HE, &HF9, &HA1, &HBA, &H9B)
 CLSID_PortableDeviceFTM = iid
End Function
Public Function CLSID_WpdSerializer() As UUID
'{0B91A74B-AD7C-4A9D-B563-29EEF9167172}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB91A74B, CInt(&HAD7C), CInt(&H4A9D), &HB5, &H63, &H29, &HEE, &HF9, &H16, &H71, &H72)
 CLSID_WpdSerializer = iid
End Function
Public Function CLSID_PortableDeviceService() As UUID
'{EF5DB4C2-9312-422C-9152-411CD9C4DD84}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEF5DB4C2, CInt(&H9312), CInt(&H422C), &H91, &H52, &H41, &H1C, &HD9, &HC4, &HDD, &H84)
 CLSID_PortableDeviceService = iid
End Function
Public Function CLSID_PortableDeviceDispatchFactory() As UUID
'{43232233-8338-4658-AE01-0B4AE830B6B0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43232233, CInt(&H8338), CInt(&H4658), &HAE, &H1, &HB, &H4A, &HE8, &H30, &HB6, &HB0)
 CLSID_PortableDeviceDispatchFactory = iid
End Function
Public Function CLSID_PortableDeviceServiceFTM() As UUID
'{1649B154-C794-497A-9B03-F3F0121302F3}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1649B154, CInt(&HC794), CInt(&H497A), &H9B, &H3, &HF3, &HF0, &H12, &H13, &H2, &HF3)
 CLSID_PortableDeviceServiceFTM = iid
End Function
Public Function CLSID_PortableDeviceWebControl() As UUID
'{186dd02c-2dec-41b5-a7d4-b59056fade51}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H186DD02C, CInt(&H2DEC), CInt(&H41B5), &HA7, &HD4, &HB5, &H90, &H56, &HFA, &HDE, &H51)
 CLSID_PortableDeviceWebControl = iid
End Function
Public Function CLSID_PortableDeviceValues() As UUID
'{0c15d503-d017-47ce-9016-7b3f978721cc}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC15D503, CInt(&HD017), CInt(&H47CE), &H90, &H16, &H7B, &H3F, &H97, &H87, &H21, &HCC)
 CLSID_PortableDeviceValues = iid
End Function
Public Function CLSID_PortableDevicePropVariantCollection() As UUID
'{08a99e2f-6d6d-4b80-af5a-baf2bcbe4cb9}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8A99E2F, CInt(&H6D6D), CInt(&H4B80), &HAF, &H5A, &HBA, &HF2, &HBC, &HBE, &H4C, &HB9)
 CLSID_PortableDevicePropVariantCollection = iid
End Function
Public Function CLSID_PortableDeviceKeyCollection() As UUID
'{de2d022d-2480-43be-97f0-d1fa2cf98f4f}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDE2D022D, CInt(&H2480), CInt(&H43BE), &H97, &HF0, &HD1, &HFA, &H2C, &HF9, &H8F, &H4F)
 CLSID_PortableDeviceKeyCollection = iid
End Function
Public Function CLSID_EnumBthMtpConnectors() As UUID
'{A1570149-E645-4F43-8B0D-409B061DB2FC}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA1570149, CInt(&HE645), CInt(&H4F43), &H8B, &HD, &H40, &H9B, &H6, &H1D, &HB2, &HFC)
 CLSID_EnumBthMtpConnectors = iid
End Function
Public Function GUID_DEVINTERFACE_WPD_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9EF44F80, &H3D64, &H4246, &HA6, &HAA, &H20, &H6F, &H32, &H8D, &H1E, &HDC)
GUID_DEVINTERFACE_WPD_SERVICE = iid
End Function
Public Function GUID_DEVINTERFACE_WPD() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6AC27878, &HA6FA, &H4155, &HBA, &H85, &HF9, &H8F, &H49, &H1D, &H4F, &H33)
GUID_DEVINTERFACE_WPD = iid
End Function
Public Function GUID_DEVINTERFACE_WPD_PRIVATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBA0C718F, &H4DED, &H49B7, &HBD, &HD3, &HFA, &HBE, &H28, &H66, &H12, &H11)
GUID_DEVINTERFACE_WPD_PRIVATE = iid
End Function

Public Function WPD_EVENT_NOTIFICATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2BA2E40A, &H6B4C, &H4295, &HBB, &H43, &H26, &H32, &H2B, &H99, &HAE, &HB2)
WPD_EVENT_NOTIFICATION = iid
End Function
Public Function WPD_EVENT_OBJECT_ADDED() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA726DA95, &HE207, &H4B02, &H8D, &H44, &HBE, &HF2, &HE8, &H6C, &HBF, &HFC)
WPD_EVENT_OBJECT_ADDED = iid
End Function
Public Function WPD_EVENT_OBJECT_REMOVED() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBE82AB88, &HA52C, &H4823, &H96, &HE5, &HD0, &H27, &H26, &H71, &HFC, &H38)
WPD_EVENT_OBJECT_REMOVED = iid
End Function
Public Function WPD_EVENT_OBJECT_UPDATED() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1445A759, &H2E01, &H485D, &H9F, &H27, &HFF, &H7, &HDA, &HE6, &H97, &HAB)
WPD_EVENT_OBJECT_UPDATED = iid
End Function
Public Function WPD_EVENT_DEVICE_RESET() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7755CF53, &HC1ED, &H44F3, &HB5, &HA2, &H45, &H1E, &H2C, &H37, &H6B, &H27)
WPD_EVENT_DEVICE_RESET = iid
End Function
Public Function WPD_EVENT_DEVICE_CAPABILITIES_UPDATED() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H36885AA1, &HCD54, &H4DAA, &HB3, &HD0, &HAF, &HB3, &HE0, &H3F, &H59, &H99)
WPD_EVENT_DEVICE_CAPABILITIES_UPDATED = iid
End Function
Public Function WPD_EVENT_STORAGE_FORMAT() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3782616B, &H22BC, &H4474, &HA2, &H51, &H30, &H70, &HF8, &HD3, &H88, &H57)
WPD_EVENT_STORAGE_FORMAT = iid
End Function
Public Function WPD_EVENT_OBJECT_TRANSFER_REQUESTED() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8D16A0A1, &HF2C6, &H41DA, &H8F, &H19, &H5E, &H53, &H72, &H1A, &HDB, &HF2)
WPD_EVENT_OBJECT_TRANSFER_REQUESTED = iid
End Function
Public Function WPD_EVENT_DEVICE_REMOVED() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE4CBCA1B, &H6918, &H48B9, &H85, &HEE, &H2, &HBE, &H7C, &H85, &HA, &HF9)
WPD_EVENT_DEVICE_REMOVED = iid
End Function
Public Function WPD_EVENT_SERVICE_METHOD_COMPLETE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8A33F5F8, &HACC, &H4D9B, &H9C, &HC4, &H11, &H2D, &H35, &H3B, &H86, &HCA)
WPD_EVENT_SERVICE_METHOD_COMPLETE = iid
End Function
Public Function WPD_CONTENT_TYPE_FUNCTIONAL_OBJECT() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H99ED0160, &H17FF, &H4C44, &H9D, &H98, &H1D, &H7A, &H6F, &H94, &H19, &H21)
WPD_CONTENT_TYPE_FUNCTIONAL_OBJECT = iid
End Function
Public Function WPD_CONTENT_TYPE_FOLDER() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H27E2E392, &HA111, &H48E0, &HAB, &HC, &HE1, &H77, &H5, &HA0, &H5F, &H85)
WPD_CONTENT_TYPE_FOLDER = iid
End Function
Public Function WPD_CONTENT_TYPE_IMAGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEF2107D5, &HA52A, &H4243, &HA2, &H6B, &H62, &HD4, &H17, &H6D, &H76, &H3)
WPD_CONTENT_TYPE_IMAGE = iid
End Function
Public Function WPD_CONTENT_TYPE_DOCUMENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H680ADF52, &H950A, &H4041, &H9B, &H41, &H65, &HE3, &H93, &H64, &H81, &H55)
WPD_CONTENT_TYPE_DOCUMENT = iid
End Function
Public Function WPD_CONTENT_TYPE_CONTACT() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEABA8313, &H4525, &H4707, &H9F, &HE, &H87, &HC6, &H80, &H8E, &H94, &H35)
WPD_CONTENT_TYPE_CONTACT = iid
End Function
Public Function WPD_CONTENT_TYPE_CONTACT_GROUP() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H346B8932, &H4C36, &H40D8, &H94, &H15, &H18, &H28, &H29, &H1F, &H9D, &HE9)
WPD_CONTENT_TYPE_CONTACT_GROUP = iid
End Function
Public Function WPD_CONTENT_TYPE_AUDIO() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4AD2C85E, &H5E2D, &H45E5, &H88, &H64, &H4F, &H22, &H9E, &H3C, &H6C, &HF0)
WPD_CONTENT_TYPE_AUDIO = iid
End Function
Public Function WPD_CONTENT_TYPE_VIDEO() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9261B03C, &H3D78, &H4519, &H85, &HE3, &H2, &HC5, &HE1, &HF5, &HB, &HB9)
WPD_CONTENT_TYPE_VIDEO = iid
End Function
Public Function WPD_CONTENT_TYPE_TELEVISION() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H60A169CF, &HF2AE, &H4E21, &H93, &H75, &H96, &H77, &HF1, &H1C, &H1C, &H6E)
WPD_CONTENT_TYPE_TELEVISION = iid
End Function
Public Function WPD_CONTENT_TYPE_PLAYLIST() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1A33F7E4, &HAF13, &H48F5, &H99, &H4E, &H77, &H36, &H9D, &HFE, &H4, &HA3)
WPD_CONTENT_TYPE_PLAYLIST = iid
End Function
Public Function WPD_CONTENT_TYPE_MIXED_CONTENT_ALBUM() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF0C3AC, &HA593, &H49AC, &H92, &H19, &H24, &HAB, &HCA, &H5A, &H25, &H63)
WPD_CONTENT_TYPE_MIXED_CONTENT_ALBUM = iid
End Function
Public Function WPD_CONTENT_TYPE_AUDIO_ALBUM() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAA18737E, &H5009, &H48FA, &HAE, &H21, &H85, &HF2, &H43, &H83, &HB4, &HE6)
WPD_CONTENT_TYPE_AUDIO_ALBUM = iid
End Function
Public Function WPD_CONTENT_TYPE_IMAGE_ALBUM() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H75793148, &H15F5, &H4A30, &HA8, &H13, &H54, &HED, &H8A, &H37, &HE2, &H26)
WPD_CONTENT_TYPE_IMAGE_ALBUM = iid
End Function
Public Function WPD_CONTENT_TYPE_VIDEO_ALBUM() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H12B0DB7, &HD4C1, &H45D6, &HB0, &H81, &H94, &HB8, &H77, &H79, &H61, &H4F)
WPD_CONTENT_TYPE_VIDEO_ALBUM = iid
End Function
Public Function WPD_CONTENT_TYPE_MEMO() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9CD20ECF, &H3B50, &H414F, &HA6, &H41, &HE4, &H73, &HFF, &HE4, &H57, &H51)
WPD_CONTENT_TYPE_MEMO = iid
End Function
Public Function WPD_CONTENT_TYPE_EMAIL() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8038044A, &H7E51, &H4F8F, &H88, &H3D, &H1D, &H6, &H23, &HD1, &H45, &H33)
WPD_CONTENT_TYPE_EMAIL = iid
End Function
Public Function WPD_CONTENT_TYPE_APPOINTMENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFED060E, &H8793, &H4B1E, &H90, &HC9, &H48, &HAC, &H38, &H9A, &HC6, &H31)
WPD_CONTENT_TYPE_APPOINTMENT = iid
End Function
Public Function WPD_CONTENT_TYPE_TASK() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H63252F2C, &H887F, &H4CB6, &HB1, &HAC, &HD2, &H98, &H55, &HDC, &HEF, &H6C)
WPD_CONTENT_TYPE_TASK = iid
End Function
Public Function WPD_CONTENT_TYPE_PROGRAM() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD269F96A, &H247C, &H4BFF, &H98, &HFB, &H97, &HF3, &HC4, &H92, &H20, &HE6)
WPD_CONTENT_TYPE_PROGRAM = iid
End Function
Public Function WPD_CONTENT_TYPE_GENERIC_FILE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H85E0A6, &H8D34, &H45D7, &HBC, &H5C, &H44, &H7E, &H59, &HC7, &H3D, &H48)
WPD_CONTENT_TYPE_GENERIC_FILE = iid
End Function
Public Function WPD_CONTENT_TYPE_CALENDAR() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA1FD5967, &H6023, &H49A0, &H9D, &HF1, &HF8, &H6, &HB, &HE7, &H51, &HB0)
WPD_CONTENT_TYPE_CALENDAR = iid
End Function
Public Function WPD_CONTENT_TYPE_GENERIC_MESSAGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE80EAAF8, &HB2DB, &H4133, &HB6, &H7E, &H1B, &HEF, &H4B, &H4A, &H6E, &H5F)
WPD_CONTENT_TYPE_GENERIC_MESSAGE = iid
End Function
Public Function WPD_CONTENT_TYPE_NETWORK_ASSOCIATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H31DA7EE, &H18C8, &H4205, &H84, &H7E, &H89, &HA1, &H12, &H61, &HD0, &HF3)
WPD_CONTENT_TYPE_NETWORK_ASSOCIATION = iid
End Function
Public Function WPD_CONTENT_TYPE_CERTIFICATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDC3876E8, &HA948, &H4060, &H90, &H50, &HCB, &HD7, &H7E, &H8A, &H3D, &H87)
WPD_CONTENT_TYPE_CERTIFICATE = iid
End Function
Public Function WPD_CONTENT_TYPE_WIRELESS_PROFILE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBAC070A, &H9F5F, &H4DA4, &HA8, &HF6, &H3D, &HE4, &H4D, &H68, &HFD, &H6C)
WPD_CONTENT_TYPE_WIRELESS_PROFILE = iid
End Function
Public Function WPD_CONTENT_TYPE_MEDIA_CAST() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5E88B3CC, &H3E65, &H4E62, &HBF, &HFF, &H22, &H94, &H95, &H25, &H3A, &HB0)
WPD_CONTENT_TYPE_MEDIA_CAST = iid
End Function
Public Function WPD_CONTENT_TYPE_SECTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H821089F5, &H1D91, &H4DC9, &HBE, &H3C, &HBB, &HB1, &HB3, &H5B, &H18, &HCE)
WPD_CONTENT_TYPE_SECTION = iid
End Function
Public Function WPD_CONTENT_TYPE_UNSPECIFIED() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H28D8D31E, &H249C, &H454E, &HAA, &HBC, &H34, &H88, &H31, &H68, &HE6, &H34)
WPD_CONTENT_TYPE_UNSPECIFIED = iid
End Function
Public Function WPD_CONTENT_TYPE_ALL() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H80E170D2, &H1055, &H4A3E, &HB9, &H52, &H82, &HCC, &H4F, &H8A, &H86, &H89)
WPD_CONTENT_TYPE_ALL = iid
End Function
Public Function WPD_FUNCTIONAL_CATEGORY_DEVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8EA466B, &HE3A4, &H4336, &HA1, &HF3, &HA4, &H4D, &H2B, &H5C, &H43, &H8C)
WPD_FUNCTIONAL_CATEGORY_DEVICE = iid
End Function
Public Function WPD_FUNCTIONAL_CATEGORY_STORAGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H23F05BBC, &H15DE, &H4C2A, &HA5, &H5B, &HA9, &HAF, &H5C, &HE4, &H12, &HEF)
WPD_FUNCTIONAL_CATEGORY_STORAGE = iid
End Function
Public Function WPD_FUNCTIONAL_CATEGORY_STILL_IMAGE_CAPTURE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H613CA327, &HAB93, &H4900, &HB4, &HFA, &H89, &H5B, &HB5, &H87, &H4B, &H79)
WPD_FUNCTIONAL_CATEGORY_STILL_IMAGE_CAPTURE = iid
End Function
Public Function WPD_FUNCTIONAL_CATEGORY_AUDIO_CAPTURE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3F2A1919, &HC7C2, &H4A00, &H85, &H5D, &HF5, &H7C, &HF0, &H6D, &HEB, &HBB)
WPD_FUNCTIONAL_CATEGORY_AUDIO_CAPTURE = iid
End Function
Public Function WPD_FUNCTIONAL_CATEGORY_VIDEO_CAPTURE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE23E5F6B, &H7243, &H43AA, &H8D, &HF1, &HE, &HB3, &HD9, &H68, &HA9, &H18)
WPD_FUNCTIONAL_CATEGORY_VIDEO_CAPTURE = iid
End Function
Public Function WPD_FUNCTIONAL_CATEGORY_SMS() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H44A0B1, &HC1E9, &H4AFD, &HB3, &H58, &HA6, &H2C, &H61, &H17, &HC9, &HCF)
WPD_FUNCTIONAL_CATEGORY_SMS = iid
End Function
Public Function WPD_FUNCTIONAL_CATEGORY_RENDERING_INFORMATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8600BA4, &HA7BA, &H4A01, &HAB, &HE, &H0, &H65, &HD0, &HA3, &H56, &HD3)
WPD_FUNCTIONAL_CATEGORY_RENDERING_INFORMATION = iid
End Function
Public Function WPD_FUNCTIONAL_CATEGORY_NETWORK_CONFIGURATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H48F4DB72, &H7C6A, &H4AB0, &H9E, &H1A, &H47, &HE, &H3C, &HDB, &HF2, &H6A)
WPD_FUNCTIONAL_CATEGORY_NETWORK_CONFIGURATION = iid
End Function
Public Function WPD_FUNCTIONAL_CATEGORY_ALL() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2D8A6512, &HA74C, &H448E, &HBA, &H8A, &HF4, &HAC, &H7, &HC4, &H93, &H99)
WPD_FUNCTIONAL_CATEGORY_ALL = iid
End Function
Public Function WPD_OBJECT_FORMAT_ICON() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H77232ED, &H102C, &H4638, &H9C, &H22, &H83, &HF1, &H42, &HBF, &HC8, &H22)
WPD_OBJECT_FORMAT_ICON = iid
End Function
Public Function WPD_OBJECT_FORMAT_M4A() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H30ABA7AC, &H6FFD, &H4C23, &HA3, &H59, &H3E, &H9B, &H52, &HF3, &HF1, &HC8)
WPD_OBJECT_FORMAT_M4A = iid
End Function
Public Function WPD_OBJECT_FORMAT_NETWORK_ASSOCIATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB1020000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_NETWORK_ASSOCIATION = iid
End Function
Public Function WPD_OBJECT_FORMAT_X509V3CERTIFICATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB1030000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_X509V3CERTIFICATE = iid
End Function
Public Function WPD_OBJECT_FORMAT_MICROSOFT_WFC() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB1040000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_MICROSOFT_WFC = iid
End Function
Public Function WPD_OBJECT_FORMAT_3GPA() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE5172730, &HF971, &H41EF, &HA1, &HB, &H22, &H71, &HA0, &H1, &H9D, &H7A)
WPD_OBJECT_FORMAT_3GPA = iid
End Function
Public Function WPD_OBJECT_FORMAT_3G2A() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1A11202D, &H8759, &H4E34, &HBA, &H5E, &HB1, &H21, &H10, &H87, &HEE, &HE4)
WPD_OBJECT_FORMAT_3G2A = iid
End Function
Public Function WPD_OBJECT_FORMAT_ALL() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC1F62EB2, &H4BB3, &H479C, &H9C, &HFA, &H5, &HB5, &HF3, &HA5, &H7B, &H22)
WPD_OBJECT_FORMAT_ALL = iid
End Function
Public Function WPD_CATEGORY_NULL() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0)
 WPD_CATEGORY_NULL = iid
End Function
Public Function WPD_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C)
 WPD_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_OBJECT_PROPERTIES_V2() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H373CD3D, &H4A46, &H40D7, &HB4, &HD8, &H73, &HE8, &HDA, &H74, &HE7, &H75)
 WPD_OBJECT_PROPERTIES_V2 = iid
End Function
Public Function WPD_FUNCTIONAL_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8F052D93, &HABCA, &H4FC5, &HA5, &HAC, &HB0, &H1D, &HF4, &HDB, &HE5, &H98)
 WPD_FUNCTIONAL_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_STORAGE_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1A3057A, &H74D6, &H4E80, &HBE, &HA7, &HDC, &H4C, &H21, &H2C, &HE5, &HA)
 WPD_STORAGE_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_NETWORK_ASSOCIATION_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE4C93C1F, &HB203, &H43F1, &HA1, &H0, &H5A, &H7, &HD1, &H1B, &H2, &H74)
 WPD_NETWORK_ASSOCIATION_PROPERTIES_V1 = iid
End Function
Public Function WPD_STILL_IMAGE_CAPTURE_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60)
 WPD_STILL_IMAGE_CAPTURE_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_RENDERING_INFORMATION_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC53D039F, &HEE23, &H4A31, &H85, &H90, &H76, &H39, &H87, &H98, &H70, &HB4)
 WPD_RENDERING_INFORMATION_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_CLIENT_INFORMATION_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H204D9F0C, &H2292, &H4080, &H9F, &H42, &H40, &H66, &H4E, &H70, &HF8, &H59)
 WPD_CLIENT_INFORMATION_PROPERTIES_V1 = iid
End Function
Public Function WPD_PROPERTY_ATTRIBUTES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB7943D8, &H6332, &H445F, &HA0, &HD, &H8D, &H5E, &HF1, &HE9, &H6F, &H37)
 WPD_PROPERTY_ATTRIBUTES_V1 = iid
End Function
Public Function WPD_PROPERTY_ATTRIBUTES_V2() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5D9DA160, &H74AE, &H43CC, &H85, &HA9, &HFE, &H55, &H5A, &H80, &H79, &H8E)
 WPD_PROPERTY_ATTRIBUTES_V2 = iid
End Function
Public Function WPD_CLASS_EXTENSION_OPTIONS_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6309FFEF, &HA87C, &H4CA7, &H84, &H34, &H79, &H75, &H76, &HE4, &HA, &H96)
 WPD_CLASS_EXTENSION_OPTIONS_V1 = iid
End Function
Public Function WPD_CLASS_EXTENSION_OPTIONS_V2() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3E3595DA, &H4D71, &H49FE, &HA0, &HB4, &HD4, &H40, &H6C, &H3A, &HE9, &H3F)
 WPD_CLASS_EXTENSION_OPTIONS_V2 = iid
End Function
Public Function WPD_CLASS_EXTENSION_OPTIONS_V3() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H65C160F8, &H1367, &H4CE2, &H93, &H9D, &H83, &H10, &H83, &H9F, &HD, &H30)
 WPD_CLASS_EXTENSION_OPTIONS_V3 = iid
End Function
Public Function WPD_RESOURCE_ATTRIBUTES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1EB6F604, &H9278, &H429F, &H93, &HCC, &H5B, &HB8, &HC0, &H66, &H56, &HB6)
 WPD_RESOURCE_ATTRIBUTES_V1 = iid
End Function
Public Function WPD_DEVICE_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H26D4979A, &HE643, &H4626, &H9E, &H2B, &H73, &H6D, &HC0, &HC9, &H2F, &HDC)
 WPD_DEVICE_PROPERTIES_V1 = iid
End Function
Public Function WPD_DEVICE_PROPERTIES_V2() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H463DD662, &H7FC4, &H4291, &H91, &H1C, &H7F, &H4C, &H9C, &HCA, &H97, &H99)
 WPD_DEVICE_PROPERTIES_V2 = iid
End Function
Public Function WPD_DEVICE_PROPERTIES_V3() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6C2B878C, &HC2EC, &H490D, &HB4, &H25, &HD7, &HA7, &H5E, &H23, &HE5, &HED)
 WPD_DEVICE_PROPERTIES_V3 = iid
End Function
Public Function WPD_SERVICE_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7510698A, &HCB54, &H481C, &HB8, &HDB, &HD, &H75, &HC9, &H3F, &H1C, &H6)
 WPD_SERVICE_PROPERTIES_V1 = iid
End Function
Public Function WPD_EVENT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H15AB1953, &HF817, &H4FEF, &HA9, &H21, &H56, &H76, &HE8, &H38, &HF6, &HE0)
 WPD_EVENT_PROPERTIES_V1 = iid
End Function
Public Function WPD_EVENT_PROPERTIES_V2() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H52807B8A, &H4914, &H4323, &H9B, &H9A, &H74, &HF6, &H54, &HB2, &HB8, &H46)
 WPD_EVENT_PROPERTIES_V2 = iid
End Function
Public Function WPD_EVENT_OPTIONS_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB3D8DAD7, &HA361, &H4B83, &H8A, &H48, &H5B, &H2, &HCE, &H10, &H71, &H3B)
 WPD_EVENT_OPTIONS_V1 = iid
End Function
Public Function WPD_EVENT_ATTRIBUTES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10C96578, &H2E81, &H4111, &HAD, &HDE, &HE0, &H8C, &HA6, &H13, &H8F, &H6D)
 WPD_EVENT_ATTRIBUTES_V1 = iid
End Function
Public Function WPD_API_OPTIONS_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10E54A3E, &H52D, &H4777, &HA1, &H3C, &HDE, &H76, &H14, &HBE, &H2B, &HC4)
 WPD_API_OPTIONS_V1 = iid
End Function
Public Function WPD_FORMAT_ATTRIBUTES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA0A02000, &HBCAF, &H4BE8, &HB3, &HF5, &H23, &H3F, &H23, &H1C, &HF5, &H8F)
 WPD_FORMAT_ATTRIBUTES_V1 = iid
End Function
Public Function WPD_METHOD_ATTRIBUTES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF17A5071, &HF039, &H44AF, &H8E, &HFE, &H43, &H2C, &HF3, &H2E, &H43, &H2A)
 WPD_METHOD_ATTRIBUTES_V1 = iid
End Function
Public Function WPD_PARAMETER_ATTRIBUTES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE6864DD7, &HF325, &H45EA, &HA1, &HD5, &H97, &HCF, &H73, &HB6, &HCA, &H58)
 WPD_PARAMETER_ATTRIBUTES_V1 = iid
End Function
Public Function WPD_CATEGORY_COMMON() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF0422A9C, &H5DC8, &H4440, &HB5, &HBD, &H5D, &HF2, &H88, &H35, &H65, &H8A)
 WPD_CATEGORY_COMMON = iid
End Function
Public Function WPD_CATEGORY_OBJECT_ENUMERATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB7474E91, &HE7F8, &H4AD9, &HB4, &H0, &HAD, &H1A, &H4B, &H58, &HEE, &HEC)
 WPD_CATEGORY_OBJECT_ENUMERATION = iid
End Function
Public Function WPD_CATEGORY_OBJECT_PROPERTIES() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9E5582E4, &H814, &H44E6, &H98, &H1A, &HB2, &H99, &H8D, &H58, &H38, &H4)
 WPD_CATEGORY_OBJECT_PROPERTIES = iid
End Function
Public Function WPD_CATEGORY_OBJECT_PROPERTIES_BULK() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E)
 WPD_CATEGORY_OBJECT_PROPERTIES_BULK = iid
End Function
Public Function WPD_CATEGORY_OBJECT_RESOURCES() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A)
 WPD_CATEGORY_OBJECT_RESOURCES = iid
End Function
Public Function WPD_CATEGORY_OBJECT_MANAGEMENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89)
 WPD_CATEGORY_OBJECT_MANAGEMENT = iid
End Function
Public Function WPD_CATEGORY_CAPABILITIES() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56)
 WPD_CATEGORY_CAPABILITIES = iid
End Function
Public Function WPD_CATEGORY_STORAGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD8F907A6, &H34CC, &H45FA, &H97, &HFB, &HD0, &H7, &HFA, &H47, &HEC, &H94)
 WPD_CATEGORY_STORAGE = iid
End Function
Public Function WPD_CATEGORY_SMS() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAFC25D66, &HFE0D, &H4114, &H90, &H97, &H97, &HC, &H93, &HE9, &H20, &HD1)
 WPD_CATEGORY_SMS = iid
End Function
Public Function WPD_CATEGORY_STILL_IMAGE_CAPTURE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4FCD6982, &H22A2, &H4B05, &HA4, &H8B, &H62, &HD3, &H8B, &HF2, &H7B, &H32)
 WPD_CATEGORY_STILL_IMAGE_CAPTURE = iid
End Function
Public Function WPD_CATEGORY_MEDIA_CAPTURE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H59B433BA, &HFE44, &H4D8D, &H80, &H8C, &H6B, &HCB, &H9B, &HF, &H15, &HE8)
 WPD_CATEGORY_MEDIA_CAPTURE = iid
End Function
Public Function WPD_CATEGORY_DEVICE_HINTS() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD5FB92B, &HCB46, &H4C4F, &H83, &H43, &HB, &HC3, &HD3, &HF1, &H7C, &H84)
 WPD_CATEGORY_DEVICE_HINTS = iid
End Function
Public Function WPD_CLASS_EXTENSION_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H33FB0D11, &H64A3, &H4FAC, &HB4, &HC7, &H3D, &HFE, &HAA, &H99, &HB0, &H51)
 WPD_CLASS_EXTENSION_V1 = iid
End Function
Public Function WPD_CLASS_EXTENSION_V2() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7F0779B5, &HFA2B, &H4766, &H9C, &HB2, &HF7, &H3B, &HA3, &HB, &H67, &H58)
 WPD_CLASS_EXTENSION_V2 = iid
End Function
Public Function WPD_CATEGORY_NETWORK_CONFIGURATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H78F9C6FC, &H79B8, &H473C, &H90, &H60, &H6B, &HD2, &H3D, &HD0, &H72, &HC4)
 WPD_CATEGORY_NETWORK_CONFIGURATION = iid
End Function
Public Function WPD_CATEGORY_SERVICE_COMMON() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H322F071D, &H36EF, &H477F, &HB4, &HB5, &H6F, &H52, &HD7, &H34, &HBA, &HEE)
 WPD_CATEGORY_SERVICE_COMMON = iid
End Function
Public Function WPD_CATEGORY_SERVICE_CAPABILITIES() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89)
 WPD_CATEGORY_SERVICE_CAPABILITIES = iid
End Function
Public Function WPD_CATEGORY_SERVICE_METHODS() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2D521CA8, &HC1B0, &H4268, &HA3, &H42, &HCF, &H19, &H32, &H15, &H69, &HBC)
 WPD_CATEGORY_SERVICE_METHODS = iid
End Function
Public Function WPD_OBJECT_FORMAT_PROPERTIES_ONLY() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H30010000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_PROPERTIES_ONLY = iid
End Function
Public Function WPD_OBJECT_FORMAT_UNSPECIFIED() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H30000000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_UNSPECIFIED = iid
End Function
Public Function WPD_OBJECT_FORMAT_SCRIPT() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H30020000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_SCRIPT = iid
End Function
Public Function WPD_OBJECT_FORMAT_EXECUTABLE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H30030000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_EXECUTABLE = iid
End Function
Public Function WPD_OBJECT_FORMAT_TEXT() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H30040000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_TEXT = iid
End Function
Public Function WPD_OBJECT_FORMAT_HTML() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H30050000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_HTML = iid
End Function
Public Function WPD_OBJECT_FORMAT_DPOF() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H30060000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_DPOF = iid
End Function
Public Function WPD_OBJECT_FORMAT_AIFF() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H30070000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_AIFF = iid
End Function
Public Function WPD_OBJECT_FORMAT_WAVE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H30080000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_WAVE = iid
End Function
Public Function WPD_OBJECT_FORMAT_MP3() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H30090000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_MP3 = iid
End Function
Public Function WPD_OBJECT_FORMAT_AVI() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H300A0000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_AVI = iid
End Function
Public Function WPD_OBJECT_FORMAT_MPEG() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H300B0000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_MPEG = iid
End Function
Public Function WPD_OBJECT_FORMAT_ASF() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H300C0000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_ASF = iid
End Function
Public Function WPD_OBJECT_FORMAT_EXIF() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H38010000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_EXIF = iid
End Function
Public Function WPD_OBJECT_FORMAT_TIFFEP() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H38020000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_TIFFEP = iid
End Function
Public Function WPD_OBJECT_FORMAT_FLASHPIX() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H38030000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_FLASHPIX = iid
End Function
Public Function WPD_OBJECT_FORMAT_BMP() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H38040000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_BMP = iid
End Function
Public Function WPD_OBJECT_FORMAT_CIFF() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H38050000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_CIFF = iid
End Function
Public Function WPD_OBJECT_FORMAT_GIF() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H38070000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_GIF = iid
End Function
Public Function WPD_OBJECT_FORMAT_JFIF() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H38080000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_JFIF = iid
End Function
Public Function WPD_OBJECT_FORMAT_PCD() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H38090000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_PCD = iid
End Function
Public Function WPD_OBJECT_FORMAT_PICT() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H380A0000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_PICT = iid
End Function
Public Function WPD_OBJECT_FORMAT_PNG() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H380B0000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_PNG = iid
End Function
Public Function WPD_OBJECT_FORMAT_TIFF() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H380D0000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_TIFF = iid
End Function
Public Function WPD_OBJECT_FORMAT_TIFFIT() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H380E0000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_TIFFIT = iid
End Function
Public Function WPD_OBJECT_FORMAT_JP2() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H380F0000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_JP2 = iid
End Function
Public Function WPD_OBJECT_FORMAT_JPX() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H38100000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_JPX = iid
End Function
Public Function WPD_OBJECT_FORMAT_WBMP() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB8030000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_WBMP = iid
End Function
Public Function WPD_OBJECT_FORMAT_JPEGXR() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB8040000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_JPEGXR = iid
End Function
Public Function WPD_OBJECT_FORMAT_WINDOWSIMAGEFORMAT() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB8810000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_WINDOWSIMAGEFORMAT = iid
End Function
Public Function WPD_OBJECT_FORMAT_WMA() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9010000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_WMA = iid
End Function
Public Function WPD_OBJECT_FORMAT_WMV() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9810000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_WMV = iid
End Function
Public Function WPD_OBJECT_FORMAT_WPLPLAYLIST() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBA100000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_WPLPLAYLIST = iid
End Function
Public Function WPD_OBJECT_FORMAT_M3UPLAYLIST() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBA110000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_M3UPLAYLIST = iid
End Function
Public Function WPD_OBJECT_FORMAT_MPLPLAYLIST() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBA120000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_MPLPLAYLIST = iid
End Function
Public Function WPD_OBJECT_FORMAT_ASXPLAYLIST() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBA130000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_ASXPLAYLIST = iid
End Function
Public Function WPD_OBJECT_FORMAT_PLSPLAYLIST() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBA140000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_PLSPLAYLIST = iid
End Function
Public Function WPD_OBJECT_FORMAT_ABSTRACT_CONTACT_GROUP() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBA060000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_ABSTRACT_CONTACT_GROUP = iid
End Function
Public Function WPD_OBJECT_FORMAT_ABSTRACT_MEDIA_CAST() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBA0B0000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_ABSTRACT_MEDIA_CAST = iid
End Function
Public Function WPD_OBJECT_FORMAT_VCALENDAR1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBE020000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_VCALENDAR1 = iid
End Function
Public Function WPD_OBJECT_FORMAT_ICALENDAR() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBE030000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_ICALENDAR = iid
End Function
Public Function WPD_OBJECT_FORMAT_ABSTRACT_CONTACT() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBB810000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_ABSTRACT_CONTACT = iid
End Function
Public Function WPD_OBJECT_FORMAT_VCARD2() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBB820000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_VCARD2 = iid
End Function
Public Function WPD_OBJECT_FORMAT_VCARD3() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBB830000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_VCARD3 = iid
End Function
Public Function WPD_OBJECT_FORMAT_XML() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBA820000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_XML = iid
End Function
Public Function WPD_OBJECT_FORMAT_AAC() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9030000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_AAC = iid
End Function
Public Function WPD_OBJECT_FORMAT_AUDIBLE() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9040000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_AUDIBLE = iid
End Function
Public Function WPD_OBJECT_FORMAT_FLAC() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9060000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_FLAC = iid
End Function
Public Function WPD_OBJECT_FORMAT_QCELP() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9070000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_QCELP = iid
End Function
Public Function WPD_OBJECT_FORMAT_AMR() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9080000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_AMR = iid
End Function
Public Function WPD_OBJECT_FORMAT_OGG() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9020000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_OGG = iid
End Function
Public Function WPD_OBJECT_FORMAT_MP4() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9820000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_MP4 = iid
End Function
Public Function WPD_OBJECT_FORMAT_MP2() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9830000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_MP2 = iid
End Function
Public Function WPD_OBJECT_FORMAT_MICROSOFT_WORD() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBA830000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_MICROSOFT_WORD = iid
End Function
Public Function WPD_OBJECT_FORMAT_MHT_COMPILED_HTML() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBA840000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_MHT_COMPILED_HTML = iid
End Function
Public Function WPD_OBJECT_FORMAT_MICROSOFT_EXCEL() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBA850000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_MICROSOFT_EXCEL = iid
End Function
Public Function WPD_OBJECT_FORMAT_MICROSOFT_POWERPOINT() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBA860000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_MICROSOFT_POWERPOINT = iid
End Function
Public Function WPD_OBJECT_FORMAT_3GP() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9840000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_3GP = iid
End Function
Public Function WPD_OBJECT_FORMAT_3G2() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9850000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_3G2 = iid
End Function
Public Function WPD_OBJECT_FORMAT_AVCHD() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9860000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_AVCHD = iid
End Function
Public Function WPD_OBJECT_FORMAT_ATSCTS() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9870000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_ATSCTS = iid
End Function
Public Function WPD_OBJECT_FORMAT_DVBTS() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9880000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_DVBTS = iid
End Function
Public Function WPD_OBJECT_FORMAT_MKV() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9900000, &HAE6C, &H4804, &H98, &HBA, &HC5, &H7B, &H46, &H96, &H5F, &HE7)
WPD_OBJECT_FORMAT_MKV = iid
End Function
Public Function WPD_FOLDER_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7E9A7ABF, &HE568, &H4B34, &HAA, &H2F, &H13, &HBB, &H12, &HAB, &H17, &H7D)
 WPD_FOLDER_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_IMAGE_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H63D64908, &H9FA1, &H479F, &H85, &HBA, &H99, &H52, &H21, &H64, &H47, &HDB)
 WPD_IMAGE_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_DOCUMENT_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB110203, &HEB95, &H4F02, &H93, &HE0, &H97, &HC6, &H31, &H49, &H3A, &HD5)
 WPD_DOCUMENT_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_MEDIA_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8)
 WPD_MEDIA_PROPERTIES_V1 = iid
End Function
Public Function WPD_CONTACT_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B)
 WPD_CONTACT_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_MUSIC_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB324F56A, &HDC5D, &H46E5, &HB6, &HDF, &HD2, &HEA, &H41, &H48, &H88, &HC6)
 WPD_MUSIC_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_VIDEO_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H346F2163, &HF998, &H4146, &H8B, &H1, &HD1, &H9B, &H4C, &H0, &HDE, &H9A)
 WPD_VIDEO_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_COMMON_INFORMATION_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB28AE94B, &H5A4, &H4E8E, &HBE, &H1, &H72, &HCC, &H7E, &H9, &H9D, &H8F)
 WPD_COMMON_INFORMATION_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_MEMO_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5FFBFC7B, &H7483, &H41AD, &HAF, &HB9, &HDA, &H3F, &H4E, &H59, &H2B, &H8D)
 WPD_MEMO_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_EMAIL_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H41F8F65A, &H5484, &H4782, &HB1, &H3D, &H47, &H40, &HDD, &H7C, &H37, &HC5)
 WPD_EMAIL_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_APPOINTMENT_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF99EFD03, &H431D, &H40D8, &HA1, &HC9, &H4E, &H22, &HD, &H9C, &H88, &HD3)
 WPD_APPOINTMENT_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_TASK_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE354E95E, &HD8A0, &H4637, &HA0, &H3A, &HC, &HB2, &H68, &H38, &HDB, &HC7)
 WPD_TASK_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_SMS_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7E1074CC, &H50FF, &H4DD1, &HA7, &H42, &H53, &HBE, &H6F, &H9, &H3A, &HD)
 WPD_SMS_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_SECTION_OBJECT_PROPERTIES_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H516AFD2B, &HC64E, &H44F0, &H98, &HDC, &HBE, &HE1, &HC8, &H8F, &H7D, &H66)
 WPD_SECTION_OBJECT_PROPERTIES_V1 = iid
End Function
Public Function WPD_CONTACT_OTHER_FULL_POSTAL_ADDRESS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 24)
 WPD_CONTACT_OTHER_FULL_POSTAL_ADDRESS = pkk
End Function
Public Function WPD_CONTACT_OTHER_POSTAL_ADDRESS_LINE1() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 25)
 WPD_CONTACT_OTHER_POSTAL_ADDRESS_LINE1 = pkk
End Function
Public Function WPD_CONTACT_OTHER_POSTAL_ADDRESS_LINE2() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 26)
 WPD_CONTACT_OTHER_POSTAL_ADDRESS_LINE2 = pkk
End Function
Public Function WPD_CONTACT_OTHER_POSTAL_ADDRESS_CITY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 27)
 WPD_CONTACT_OTHER_POSTAL_ADDRESS_CITY = pkk
End Function
Public Function WPD_CONTACT_OTHER_POSTAL_ADDRESS_REGION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 28)
 WPD_CONTACT_OTHER_POSTAL_ADDRESS_REGION = pkk
End Function
Public Function WPD_CONTACT_OTHER_POSTAL_ADDRESS_POSTAL_CODE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 29)
 WPD_CONTACT_OTHER_POSTAL_ADDRESS_POSTAL_CODE = pkk
End Function
Public Function WPD_CONTACT_OTHER_POSTAL_ADDRESS_POSTAL_COUNTRY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 30)
 WPD_CONTACT_OTHER_POSTAL_ADDRESS_POSTAL_COUNTRY = pkk
End Function
Public Function WPD_CONTACT_PRIMARY_EMAIL_ADDRESS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 31)
 WPD_CONTACT_PRIMARY_EMAIL_ADDRESS = pkk
End Function
Public Function WPD_CONTACT_PERSONAL_EMAIL() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 32)
 WPD_CONTACT_PERSONAL_EMAIL = pkk
End Function
Public Function WPD_CONTACT_PERSONAL_EMAIL2() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 33)
 WPD_CONTACT_PERSONAL_EMAIL2 = pkk
End Function
Public Function WPD_CONTACT_BUSINESS_EMAIL() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 34)
 WPD_CONTACT_BUSINESS_EMAIL = pkk
End Function
Public Function WPD_CONTACT_BUSINESS_EMAIL2() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 35)
 WPD_CONTACT_BUSINESS_EMAIL2 = pkk
End Function
Public Function WPD_CONTACT_OTHER_EMAILS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 36)
 WPD_CONTACT_OTHER_EMAILS = pkk
End Function
Public Function WPD_CONTACT_PRIMARY_PHONE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 37)
 WPD_CONTACT_PRIMARY_PHONE = pkk
End Function
Public Function WPD_CONTACT_PERSONAL_PHONE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 38)
 WPD_CONTACT_PERSONAL_PHONE = pkk
End Function
Public Function WPD_CONTACT_PERSONAL_PHONE2() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 39)
 WPD_CONTACT_PERSONAL_PHONE2 = pkk
End Function
Public Function WPD_CONTACT_BUSINESS_PHONE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 40)
 WPD_CONTACT_BUSINESS_PHONE = pkk
End Function
Public Function WPD_CONTACT_BUSINESS_PHONE2() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 41)
 WPD_CONTACT_BUSINESS_PHONE2 = pkk
End Function
Public Function WPD_CONTACT_MOBILE_PHONE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 42)
 WPD_CONTACT_MOBILE_PHONE = pkk
End Function
Public Function WPD_CONTACT_MOBILE_PHONE2() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 43)
 WPD_CONTACT_MOBILE_PHONE2 = pkk
End Function
Public Function WPD_CONTACT_PERSONAL_FAX() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 44)
 WPD_CONTACT_PERSONAL_FAX = pkk
End Function
Public Function WPD_CONTACT_BUSINESS_FAX() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 45)
 WPD_CONTACT_BUSINESS_FAX = pkk
End Function
Public Function WPD_CONTACT_PAGER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 46)
 WPD_CONTACT_PAGER = pkk
End Function
Public Function WPD_CONTACT_OTHER_PHONES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 47)
 WPD_CONTACT_OTHER_PHONES = pkk
End Function
Public Function WPD_CONTACT_PRIMARY_WEB_ADDRESS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 48)
 WPD_CONTACT_PRIMARY_WEB_ADDRESS = pkk
End Function
Public Function WPD_CONTACT_PERSONAL_WEB_ADDRESS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 49)
 WPD_CONTACT_PERSONAL_WEB_ADDRESS = pkk
End Function
Public Function WPD_CONTACT_BUSINESS_WEB_ADDRESS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 50)
 WPD_CONTACT_BUSINESS_WEB_ADDRESS = pkk
End Function
Public Function WPD_CONTACT_INSTANT_MESSENGER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 51)
 WPD_CONTACT_INSTANT_MESSENGER = pkk
End Function
Public Function WPD_CONTACT_INSTANT_MESSENGER2() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 52)
 WPD_CONTACT_INSTANT_MESSENGER2 = pkk
End Function
Public Function WPD_CONTACT_INSTANT_MESSENGER3() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 53)
 WPD_CONTACT_INSTANT_MESSENGER3 = pkk
End Function
Public Function WPD_CONTACT_COMPANY_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 54)
 WPD_CONTACT_COMPANY_NAME = pkk
End Function
Public Function WPD_CONTACT_PHONETIC_COMPANY_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 55)
 WPD_CONTACT_PHONETIC_COMPANY_NAME = pkk
End Function
Public Function WPD_CONTACT_ROLE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 56)
 WPD_CONTACT_ROLE = pkk
End Function
Public Function WPD_CONTACT_BIRTHDATE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 57)
 WPD_CONTACT_BIRTHDATE = pkk
End Function
Public Function WPD_CONTACT_PRIMARY_FAX() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 58)
 WPD_CONTACT_PRIMARY_FAX = pkk
End Function
Public Function WPD_CONTACT_SPOUSE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 59)
 WPD_CONTACT_SPOUSE = pkk
End Function
Public Function WPD_CONTACT_CHILDREN() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 60)
 WPD_CONTACT_CHILDREN = pkk
End Function
Public Function WPD_CONTACT_ASSISTANT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 61)
 WPD_CONTACT_ASSISTANT = pkk
End Function
Public Function WPD_CONTACT_ANNIVERSARY_DATE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 62)
 WPD_CONTACT_ANNIVERSARY_DATE = pkk
End Function
Public Function WPD_CONTACT_RINGTONE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 63)
 WPD_CONTACT_RINGTONE = pkk
End Function
Public Function WPD_MUSIC_ALBUM() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB324F56A, &HDC5D, &H46E5, &HB6, &HDF, &HD2, &HEA, &H41, &H48, &H88, &HC6, 3)
 WPD_MUSIC_ALBUM = pkk
End Function
Public Function WPD_MUSIC_TRACK() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB324F56A, &HDC5D, &H46E5, &HB6, &HDF, &HD2, &HEA, &H41, &H48, &H88, &HC6, 4)
 WPD_MUSIC_TRACK = pkk
End Function
Public Function WPD_MUSIC_LYRICS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB324F56A, &HDC5D, &H46E5, &HB6, &HDF, &HD2, &HEA, &H41, &H48, &H88, &HC6, 6)
 WPD_MUSIC_LYRICS = pkk
End Function
Public Function WPD_MUSIC_MOOD() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB324F56A, &HDC5D, &H46E5, &HB6, &HDF, &HD2, &HEA, &H41, &H48, &H88, &HC6, 8)
 WPD_MUSIC_MOOD = pkk
End Function
Public Function WPD_AUDIO_BITRATE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB324F56A, &HDC5D, &H46E5, &HB6, &HDF, &HD2, &HEA, &H41, &H48, &H88, &HC6, 9)
 WPD_AUDIO_BITRATE = pkk
End Function
Public Function WPD_AUDIO_CHANNEL_COUNT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB324F56A, &HDC5D, &H46E5, &HB6, &HDF, &HD2, &HEA, &H41, &H48, &H88, &HC6, 10)
 WPD_AUDIO_CHANNEL_COUNT = pkk
End Function
Public Function WPD_AUDIO_FORMAT_CODE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB324F56A, &HDC5D, &H46E5, &HB6, &HDF, &HD2, &HEA, &H41, &H48, &H88, &HC6, 11)
 WPD_AUDIO_FORMAT_CODE = pkk
End Function
Public Function WPD_AUDIO_BIT_DEPTH() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB324F56A, &HDC5D, &H46E5, &HB6, &HDF, &HD2, &HEA, &H41, &H48, &H88, &HC6, 12)
 WPD_AUDIO_BIT_DEPTH = pkk
End Function
Public Function WPD_AUDIO_BLOCK_ALIGNMENT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB324F56A, &HDC5D, &H46E5, &HB6, &HDF, &HD2, &HEA, &H41, &H48, &H88, &HC6, 13)
 WPD_AUDIO_BLOCK_ALIGNMENT = pkk
End Function
Public Function WPD_VIDEO_AUTHOR() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H346F2163, &HF998, &H4146, &H8B, &H1, &HD1, &H9B, &H4C, &H0, &HDE, &H9A, 2)
 WPD_VIDEO_AUTHOR = pkk
End Function
Public Function WPD_VIDEO_RECORDEDTV_STATION_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H346F2163, &HF998, &H4146, &H8B, &H1, &HD1, &H9B, &H4C, &H0, &HDE, &H9A, 4)
 WPD_VIDEO_RECORDEDTV_STATION_NAME = pkk
End Function
Public Function WPD_VIDEO_RECORDEDTV_CHANNEL_NUMBER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H346F2163, &HF998, &H4146, &H8B, &H1, &HD1, &H9B, &H4C, &H0, &HDE, &H9A, 5)
 WPD_VIDEO_RECORDEDTV_CHANNEL_NUMBER = pkk
End Function
Public Function WPD_VIDEO_RECORDEDTV_REPEAT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H346F2163, &HF998, &H4146, &H8B, &H1, &HD1, &H9B, &H4C, &H0, &HDE, &H9A, 7)
 WPD_VIDEO_RECORDEDTV_REPEAT = pkk
End Function
Public Function WPD_VIDEO_BUFFER_SIZE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H346F2163, &HF998, &H4146, &H8B, &H1, &HD1, &H9B, &H4C, &H0, &HDE, &H9A, 8)
 WPD_VIDEO_BUFFER_SIZE = pkk
End Function
Public Function WPD_VIDEO_CREDITS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H346F2163, &HF998, &H4146, &H8B, &H1, &HD1, &H9B, &H4C, &H0, &HDE, &H9A, 9)
 WPD_VIDEO_CREDITS = pkk
End Function
Public Function WPD_VIDEO_KEY_FRAME_DISTANCE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H346F2163, &HF998, &H4146, &H8B, &H1, &HD1, &H9B, &H4C, &H0, &HDE, &H9A, 10)
 WPD_VIDEO_KEY_FRAME_DISTANCE = pkk
End Function
Public Function WPD_VIDEO_QUALITY_SETTING() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H346F2163, &HF998, &H4146, &H8B, &H1, &HD1, &H9B, &H4C, &H0, &HDE, &H9A, 11)
 WPD_VIDEO_QUALITY_SETTING = pkk
End Function
Public Function WPD_VIDEO_SCAN_TYPE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H346F2163, &HF998, &H4146, &H8B, &H1, &HD1, &H9B, &H4C, &H0, &HDE, &H9A, 12)
 WPD_VIDEO_SCAN_TYPE = pkk
End Function
Public Function WPD_VIDEO_BITRATE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H346F2163, &HF998, &H4146, &H8B, &H1, &HD1, &H9B, &H4C, &H0, &HDE, &H9A, 13)
 WPD_VIDEO_BITRATE = pkk
End Function
Public Function WPD_VIDEO_FOURCC_CODE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H346F2163, &HF998, &H4146, &H8B, &H1, &HD1, &H9B, &H4C, &H0, &HDE, &H9A, 14)
 WPD_VIDEO_FOURCC_CODE = pkk
End Function
Public Function WPD_VIDEO_FRAMERATE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H346F2163, &HF998, &H4146, &H8B, &H1, &HD1, &H9B, &H4C, &H0, &HDE, &H9A, 15)
 WPD_VIDEO_FRAMERATE = pkk
End Function
Public Function WPD_COMMON_INFORMATION_SUBJECT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB28AE94B, &H5A4, &H4E8E, &HBE, &H1, &H72, &HCC, &H7E, &H9, &H9D, &H8F, 2)
 WPD_COMMON_INFORMATION_SUBJECT = pkk
End Function
Public Function WPD_COMMON_INFORMATION_BODY_TEXT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB28AE94B, &H5A4, &H4E8E, &HBE, &H1, &H72, &HCC, &H7E, &H9, &H9D, &H8F, 3)
 WPD_COMMON_INFORMATION_BODY_TEXT = pkk
End Function
Public Function WPD_COMMON_INFORMATION_PRIORITY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB28AE94B, &H5A4, &H4E8E, &HBE, &H1, &H72, &HCC, &H7E, &H9, &H9D, &H8F, 4)
 WPD_COMMON_INFORMATION_PRIORITY = pkk
End Function
Public Function WPD_COMMON_INFORMATION_START_DATETIME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB28AE94B, &H5A4, &H4E8E, &HBE, &H1, &H72, &HCC, &H7E, &H9, &H9D, &H8F, 5)
 WPD_COMMON_INFORMATION_START_DATETIME = pkk
End Function
Public Function WPD_COMMON_INFORMATION_END_DATETIME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB28AE94B, &H5A4, &H4E8E, &HBE, &H1, &H72, &HCC, &H7E, &H9, &H9D, &H8F, 6)
 WPD_COMMON_INFORMATION_END_DATETIME = pkk
End Function
Public Function WPD_COMMON_INFORMATION_NOTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB28AE94B, &H5A4, &H4E8E, &HBE, &H1, &H72, &HCC, &H7E, &H9, &H9D, &H8F, 7)
 WPD_COMMON_INFORMATION_NOTES = pkk
End Function
Public Function WPD_EMAIL_TO_LINE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H41F8F65A, &H5484, &H4782, &HB1, &H3D, &H47, &H40, &HDD, &H7C, &H37, &HC5, 2)
 WPD_EMAIL_TO_LINE = pkk
End Function
Public Function WPD_EMAIL_CC_LINE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H41F8F65A, &H5484, &H4782, &HB1, &H3D, &H47, &H40, &HDD, &H7C, &H37, &HC5, 3)
 WPD_EMAIL_CC_LINE = pkk
End Function
Public Function WPD_EMAIL_BCC_LINE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H41F8F65A, &H5484, &H4782, &HB1, &H3D, &H47, &H40, &HDD, &H7C, &H37, &HC5, 4)
 WPD_EMAIL_BCC_LINE = pkk
End Function
Public Function WPD_EMAIL_HAS_BEEN_READ() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H41F8F65A, &H5484, &H4782, &HB1, &H3D, &H47, &H40, &HDD, &H7C, &H37, &HC5, 7)
 WPD_EMAIL_HAS_BEEN_READ = pkk
End Function
Public Function WPD_EMAIL_RECEIVED_TIME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H41F8F65A, &H5484, &H4782, &HB1, &H3D, &H47, &H40, &HDD, &H7C, &H37, &HC5, 8)
 WPD_EMAIL_RECEIVED_TIME = pkk
End Function
Public Function WPD_EMAIL_HAS_ATTACHMENTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H41F8F65A, &H5484, &H4782, &HB1, &H3D, &H47, &H40, &HDD, &H7C, &H37, &HC5, 9)
 WPD_EMAIL_HAS_ATTACHMENTS = pkk
End Function
Public Function WPD_EMAIL_SENDER_ADDRESS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H41F8F65A, &H5484, &H4782, &HB1, &H3D, &H47, &H40, &HDD, &H7C, &H37, &HC5, 10)
 WPD_EMAIL_SENDER_ADDRESS = pkk
End Function
Public Function WPD_APPOINTMENT_LOCATION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF99EFD03, &H431D, &H40D8, &HA1, &HC9, &H4E, &H22, &HD, &H9C, &H88, &HD3, 3)
 WPD_APPOINTMENT_LOCATION = pkk
End Function
Public Function WPD_APPOINTMENT_TYPE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF99EFD03, &H431D, &H40D8, &HA1, &HC9, &H4E, &H22, &HD, &H9C, &H88, &HD3, 7)
 WPD_APPOINTMENT_TYPE = pkk
End Function
Public Function WPD_APPOINTMENT_REQUIRED_ATTENDEES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF99EFD03, &H431D, &H40D8, &HA1, &HC9, &H4E, &H22, &HD, &H9C, &H88, &HD3, 8)
 WPD_APPOINTMENT_REQUIRED_ATTENDEES = pkk
End Function
Public Function WPD_APPOINTMENT_OPTIONAL_ATTENDEES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF99EFD03, &H431D, &H40D8, &HA1, &HC9, &H4E, &H22, &HD, &H9C, &H88, &HD3, 9)
 WPD_APPOINTMENT_OPTIONAL_ATTENDEES = pkk
End Function
Public Function WPD_APPOINTMENT_ACCEPTED_ATTENDEES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF99EFD03, &H431D, &H40D8, &HA1, &HC9, &H4E, &H22, &HD, &H9C, &H88, &HD3, 10)
 WPD_APPOINTMENT_ACCEPTED_ATTENDEES = pkk
End Function
Public Function WPD_APPOINTMENT_RESOURCES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF99EFD03, &H431D, &H40D8, &HA1, &HC9, &H4E, &H22, &HD, &H9C, &H88, &HD3, 11)
 WPD_APPOINTMENT_RESOURCES = pkk
End Function
Public Function WPD_APPOINTMENT_TENTATIVE_ATTENDEES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF99EFD03, &H431D, &H40D8, &HA1, &HC9, &H4E, &H22, &HD, &H9C, &H88, &HD3, 12)
 WPD_APPOINTMENT_TENTATIVE_ATTENDEES = pkk
End Function
Public Function WPD_APPOINTMENT_DECLINED_ATTENDEES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF99EFD03, &H431D, &H40D8, &HA1, &HC9, &H4E, &H22, &HD, &H9C, &H88, &HD3, 13)
 WPD_APPOINTMENT_DECLINED_ATTENDEES = pkk
End Function
Public Function WPD_TASK_STATUS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE354E95E, &HD8A0, &H4637, &HA0, &H3A, &HC, &HB2, &H68, &H38, &HDB, &HC7, 6)
 WPD_TASK_STATUS = pkk
End Function
Public Function WPD_TASK_PERCENT_COMPLETE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE354E95E, &HD8A0, &H4637, &HA0, &H3A, &HC, &HB2, &H68, &H38, &HDB, &HC7, 8)
 WPD_TASK_PERCENT_COMPLETE = pkk
End Function
Public Function WPD_TASK_REMINDER_DATE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE354E95E, &HD8A0, &H4637, &HA0, &H3A, &HC, &HB2, &H68, &H38, &HDB, &HC7, 10)
 WPD_TASK_REMINDER_DATE = pkk
End Function
Public Function WPD_TASK_OWNER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE354E95E, &HD8A0, &H4637, &HA0, &H3A, &HC, &HB2, &H68, &H38, &HDB, &HC7, 11)
 WPD_TASK_OWNER = pkk
End Function
Public Function WPD_SMS_PROVIDER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7E1074CC, &H50FF, &H4DD1, &HA7, &H42, &H53, &HBE, &H6F, &H9, &H3A, &HD, 2)
 WPD_SMS_PROVIDER = pkk
End Function
Public Function WPD_SMS_TIMEOUT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7E1074CC, &H50FF, &H4DD1, &HA7, &H42, &H53, &HBE, &H6F, &H9, &H3A, &HD, 3)
 WPD_SMS_TIMEOUT = pkk
End Function
Public Function WPD_SMS_MAX_PAYLOAD() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7E1074CC, &H50FF, &H4DD1, &HA7, &H42, &H53, &HBE, &H6F, &H9, &H3A, &HD, 4)
 WPD_SMS_MAX_PAYLOAD = pkk
End Function
Public Function WPD_SMS_ENCODING() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7E1074CC, &H50FF, &H4DD1, &HA7, &H42, &H53, &HBE, &H6F, &H9, &H3A, &HD, 5)
 WPD_SMS_ENCODING = pkk
End Function
Public Function WPD_SECTION_DATA_OFFSET() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H516AFD2B, &HC64E, &H44F0, &H98, &HDC, &HBE, &HE1, &HC8, &H8F, &H7D, &H66, 2)
 WPD_SECTION_DATA_OFFSET = pkk
End Function
Public Function WPD_SECTION_DATA_LENGTH() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H516AFD2B, &HC64E, &H44F0, &H98, &HDC, &HBE, &HE1, &HC8, &H8F, &H7D, &H66, 3)
 WPD_SECTION_DATA_LENGTH = pkk
End Function
Public Function WPD_SECTION_DATA_UNITS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H516AFD2B, &HC64E, &H44F0, &H98, &HDC, &HBE, &HE1, &HC8, &H8F, &H7D, &H66, 4)
 WPD_SECTION_DATA_UNITS = pkk
End Function
Public Function WPD_SECTION_DATA_REFERENCED_OBJECT_RESOURCE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H516AFD2B, &HC64E, &H44F0, &H98, &HDC, &HBE, &HE1, &HC8, &H8F, &H7D, &H66, 5)
 WPD_SECTION_DATA_REFERENCED_OBJECT_RESOURCE = pkk
End Function
Public Function WPD_OBJECT_ISHIDDEN() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 9)
 WPD_OBJECT_ISHIDDEN = pkk
End Function
Public Function WPD_OBJECT_ISSYSTEM() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 10)
 WPD_OBJECT_ISSYSTEM = pkk
End Function
Public Function WPD_OBJECT_SIZE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 11)
 WPD_OBJECT_SIZE = pkk
End Function
Public Function WPD_OBJECT_ORIGINAL_FILE_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 12)
 WPD_OBJECT_ORIGINAL_FILE_NAME = pkk
End Function
Public Function WPD_OBJECT_NON_CONSUMABLE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 13)
 WPD_OBJECT_NON_CONSUMABLE = pkk
End Function
Public Function WPD_OBJECT_KEYWORDS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 15)
 WPD_OBJECT_KEYWORDS = pkk
End Function
Public Function WPD_OBJECT_SYNC_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 16)
 WPD_OBJECT_SYNC_ID = pkk
End Function
Public Function WPD_OBJECT_IS_DRM_PROTECTED() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 17)
 WPD_OBJECT_IS_DRM_PROTECTED = pkk
End Function
Public Function WPD_OBJECT_DATE_CREATED() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 18)
 WPD_OBJECT_DATE_CREATED = pkk
End Function
Public Function WPD_OBJECT_DATE_MODIFIED() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 19)
 WPD_OBJECT_DATE_MODIFIED = pkk
End Function
Public Function WPD_OBJECT_DATE_AUTHORED() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 20)
 WPD_OBJECT_DATE_AUTHORED = pkk
End Function
Public Function WPD_OBJECT_BACK_REFERENCES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 21)
 WPD_OBJECT_BACK_REFERENCES = pkk
End Function
Public Function WPD_OBJECT_CAN_DELETE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 26)
 WPD_OBJECT_CAN_DELETE = pkk
End Function
Public Function WPD_OBJECT_LANGUAGE_LOCALE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 27)
 WPD_OBJECT_LANGUAGE_LOCALE = pkk
End Function
Public Function WPD_FOLDER_CONTENT_TYPES_ALLOWED() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7E9A7ABF, &HE568, &H4B34, &HAA, &H2F, &H13, &HBB, &H12, &HAB, &H17, &H7D, 2)
 WPD_FOLDER_CONTENT_TYPES_ALLOWED = pkk
End Function
Public Function WPD_IMAGE_BITDEPTH() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H63D64908, &H9FA1, &H479F, &H85, &HBA, &H99, &H52, &H21, &H64, &H47, &HDB, 3)
 WPD_IMAGE_BITDEPTH = pkk
End Function
Public Function WPD_IMAGE_CROPPED_STATUS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H63D64908, &H9FA1, &H479F, &H85, &HBA, &H99, &H52, &H21, &H64, &H47, &HDB, 4)
 WPD_IMAGE_CROPPED_STATUS = pkk
End Function
Public Function WPD_IMAGE_COLOR_CORRECTED_STATUS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H63D64908, &H9FA1, &H479F, &H85, &HBA, &H99, &H52, &H21, &H64, &H47, &HDB, 5)
 WPD_IMAGE_COLOR_CORRECTED_STATUS = pkk
End Function
Public Function WPD_IMAGE_FNUMBER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H63D64908, &H9FA1, &H479F, &H85, &HBA, &H99, &H52, &H21, &H64, &H47, &HDB, 6)
 WPD_IMAGE_FNUMBER = pkk
End Function
Public Function WPD_IMAGE_EXPOSURE_TIME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H63D64908, &H9FA1, &H479F, &H85, &HBA, &H99, &H52, &H21, &H64, &H47, &HDB, 7)
 WPD_IMAGE_EXPOSURE_TIME = pkk
End Function
Public Function WPD_IMAGE_EXPOSURE_INDEX() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H63D64908, &H9FA1, &H479F, &H85, &HBA, &H99, &H52, &H21, &H64, &H47, &HDB, 8)
 WPD_IMAGE_EXPOSURE_INDEX = pkk
End Function
Public Function WPD_IMAGE_HORIZONTAL_RESOLUTION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H63D64908, &H9FA1, &H479F, &H85, &HBA, &H99, &H52, &H21, &H64, &H47, &HDB, 9)
 WPD_IMAGE_HORIZONTAL_RESOLUTION = pkk
End Function
Public Function WPD_IMAGE_VERTICAL_RESOLUTION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H63D64908, &H9FA1, &H479F, &H85, &HBA, &H99, &H52, &H21, &H64, &H47, &HDB, 10)
 WPD_IMAGE_VERTICAL_RESOLUTION = pkk
End Function
Public Function WPD_MEDIA_TOTAL_BITRATE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 2)
 WPD_MEDIA_TOTAL_BITRATE = pkk
End Function
Public Function WPD_MEDIA_BITRATE_TYPE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 3)
 WPD_MEDIA_BITRATE_TYPE = pkk
End Function
Public Function WPD_MEDIA_COPYRIGHT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 4)
 WPD_MEDIA_COPYRIGHT = pkk
End Function
Public Function WPD_MEDIA_SUBSCRIPTION_CONTENT_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 5)
 WPD_MEDIA_SUBSCRIPTION_CONTENT_ID = pkk
End Function
Public Function WPD_MEDIA_USE_COUNT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 6)
 WPD_MEDIA_USE_COUNT = pkk
End Function
Public Function WPD_MEDIA_SKIP_COUNT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 7)
 WPD_MEDIA_SKIP_COUNT = pkk
End Function
Public Function WPD_MEDIA_LAST_ACCESSED_TIME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 8)
 WPD_MEDIA_LAST_ACCESSED_TIME = pkk
End Function
Public Function WPD_MEDIA_PARENTAL_RATING() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 9)
 WPD_MEDIA_PARENTAL_RATING = pkk
End Function
Public Function WPD_MEDIA_META_GENRE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 10)
 WPD_MEDIA_META_GENRE = pkk
End Function
Public Function WPD_MEDIA_COMPOSER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 11)
 WPD_MEDIA_COMPOSER = pkk
End Function
Public Function WPD_MEDIA_EFFECTIVE_RATING() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 12)
 WPD_MEDIA_EFFECTIVE_RATING = pkk
End Function
Public Function WPD_MEDIA_SUB_TITLE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 13)
 WPD_MEDIA_SUB_TITLE = pkk
End Function
Public Function WPD_MEDIA_RELEASE_DATE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 14)
 WPD_MEDIA_RELEASE_DATE = pkk
End Function
Public Function WPD_MEDIA_SAMPLE_RATE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 15)
 WPD_MEDIA_SAMPLE_RATE = pkk
End Function
Public Function WPD_MEDIA_STAR_RATING() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 16)
 WPD_MEDIA_STAR_RATING = pkk
End Function
Public Function WPD_MEDIA_USER_EFFECTIVE_RATING() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 17)
 WPD_MEDIA_USER_EFFECTIVE_RATING = pkk
End Function
Public Function WPD_MEDIA_TITLE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 18)
 WPD_MEDIA_TITLE = pkk
End Function
Public Function WPD_MEDIA_DURATION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 19)
 WPD_MEDIA_DURATION = pkk
End Function
Public Function WPD_MEDIA_BUY_NOW() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 20)
 WPD_MEDIA_BUY_NOW = pkk
End Function
Public Function WPD_MEDIA_ENCODING_PROFILE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 21)
 WPD_MEDIA_ENCODING_PROFILE = pkk
End Function
Public Function WPD_MEDIA_WIDTH() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 22)
 WPD_MEDIA_WIDTH = pkk
End Function
Public Function WPD_MEDIA_HEIGHT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 23)
 WPD_MEDIA_HEIGHT = pkk
End Function
Public Function WPD_MEDIA_ARTIST() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 24)
 WPD_MEDIA_ARTIST = pkk
End Function
Public Function WPD_MEDIA_ALBUM_ARTIST() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 25)
 WPD_MEDIA_ALBUM_ARTIST = pkk
End Function
Public Function WPD_MEDIA_OWNER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 26)
 WPD_MEDIA_OWNER = pkk
End Function
Public Function WPD_MEDIA_MANAGING_EDITOR() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 27)
 WPD_MEDIA_MANAGING_EDITOR = pkk
End Function
Public Function WPD_MEDIA_WEBMASTER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 28)
 WPD_MEDIA_WEBMASTER = pkk
End Function
Public Function WPD_MEDIA_SOURCE_URL() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 29)
 WPD_MEDIA_SOURCE_URL = pkk
End Function
Public Function WPD_MEDIA_DESTINATION_URL() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 30)
 WPD_MEDIA_DESTINATION_URL = pkk
End Function
Public Function WPD_MEDIA_DESCRIPTION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 31)
 WPD_MEDIA_DESCRIPTION = pkk
End Function
Public Function WPD_MEDIA_GENRE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 32)
 WPD_MEDIA_GENRE = pkk
End Function
Public Function WPD_MEDIA_TIME_BOOKMARK() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 33)
 WPD_MEDIA_TIME_BOOKMARK = pkk
End Function
Public Function WPD_MEDIA_OBJECT_BOOKMARK() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 34)
 WPD_MEDIA_OBJECT_BOOKMARK = pkk
End Function
Public Function WPD_MEDIA_LAST_BUILD_DATE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 35)
 WPD_MEDIA_LAST_BUILD_DATE = pkk
End Function
Public Function WPD_MEDIA_BYTE_BOOKMARK() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 36)
 WPD_MEDIA_BYTE_BOOKMARK = pkk
End Function
Public Function WPD_MEDIA_TIME_TO_LIVE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 37)
 WPD_MEDIA_TIME_TO_LIVE = pkk
End Function
Public Function WPD_MEDIA_GUID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 38)
 WPD_MEDIA_GUID = pkk
End Function
Public Function WPD_MEDIA_SUB_DESCRIPTION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 39)
 WPD_MEDIA_SUB_DESCRIPTION = pkk
End Function
Public Function WPD_MEDIA_AUDIO_ENCODING_PROFILE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2ED8BA05, &HAD3, &H42DC, &HB0, &HD0, &HBC, &H95, &HAC, &H39, &H6A, &HC8, 49)
 WPD_MEDIA_AUDIO_ENCODING_PROFILE = pkk
End Function
Public Function WPD_CONTACT_DISPLAY_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 2)
 WPD_CONTACT_DISPLAY_NAME = pkk
End Function
Public Function WPD_CONTACT_FIRST_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 3)
 WPD_CONTACT_FIRST_NAME = pkk
End Function
Public Function WPD_CONTACT_MIDDLE_NAMES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 4)
 WPD_CONTACT_MIDDLE_NAMES = pkk
End Function
Public Function WPD_CONTACT_LAST_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 5)
 WPD_CONTACT_LAST_NAME = pkk
End Function
Public Function WPD_CONTACT_PREFIX() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 6)
 WPD_CONTACT_PREFIX = pkk
End Function
Public Function WPD_CONTACT_SUFFIX() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 7)
 WPD_CONTACT_SUFFIX = pkk
End Function
Public Function WPD_CONTACT_PHONETIC_FIRST_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 8)
 WPD_CONTACT_PHONETIC_FIRST_NAME = pkk
End Function
Public Function WPD_CONTACT_PHONETIC_LAST_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 9)
 WPD_CONTACT_PHONETIC_LAST_NAME = pkk
End Function
Public Function WPD_CONTACT_PERSONAL_FULL_POSTAL_ADDRESS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 10)
 WPD_CONTACT_PERSONAL_FULL_POSTAL_ADDRESS = pkk
End Function
Public Function WPD_CONTACT_PERSONAL_POSTAL_ADDRESS_LINE1() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 11)
 WPD_CONTACT_PERSONAL_POSTAL_ADDRESS_LINE1 = pkk
End Function
Public Function WPD_CONTACT_PERSONAL_POSTAL_ADDRESS_LINE2() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 12)
 WPD_CONTACT_PERSONAL_POSTAL_ADDRESS_LINE2 = pkk
End Function
Public Function WPD_CONTACT_PERSONAL_POSTAL_ADDRESS_CITY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 13)
 WPD_CONTACT_PERSONAL_POSTAL_ADDRESS_CITY = pkk
End Function
Public Function WPD_CONTACT_PERSONAL_POSTAL_ADDRESS_REGION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 14)
 WPD_CONTACT_PERSONAL_POSTAL_ADDRESS_REGION = pkk
End Function
Public Function WPD_CONTACT_PERSONAL_POSTAL_ADDRESS_POSTAL_CODE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 15)
 WPD_CONTACT_PERSONAL_POSTAL_ADDRESS_POSTAL_CODE = pkk
End Function
Public Function WPD_CONTACT_PERSONAL_POSTAL_ADDRESS_COUNTRY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 16)
 WPD_CONTACT_PERSONAL_POSTAL_ADDRESS_COUNTRY = pkk
End Function
Public Function WPD_CONTACT_BUSINESS_FULL_POSTAL_ADDRESS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 17)
 WPD_CONTACT_BUSINESS_FULL_POSTAL_ADDRESS = pkk
End Function
Public Function WPD_CONTACT_BUSINESS_POSTAL_ADDRESS_LINE1() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 18)
 WPD_CONTACT_BUSINESS_POSTAL_ADDRESS_LINE1 = pkk
End Function
Public Function WPD_CONTACT_BUSINESS_POSTAL_ADDRESS_LINE2() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 19)
 WPD_CONTACT_BUSINESS_POSTAL_ADDRESS_LINE2 = pkk
End Function
Public Function WPD_CONTACT_BUSINESS_POSTAL_ADDRESS_CITY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 20)
 WPD_CONTACT_BUSINESS_POSTAL_ADDRESS_CITY = pkk
End Function
Public Function WPD_CONTACT_BUSINESS_POSTAL_ADDRESS_REGION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 21)
 WPD_CONTACT_BUSINESS_POSTAL_ADDRESS_REGION = pkk
End Function
Public Function WPD_CONTACT_BUSINESS_POSTAL_ADDRESS_POSTAL_CODE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 22)
 WPD_CONTACT_BUSINESS_POSTAL_ADDRESS_POSTAL_CODE = pkk
End Function
Public Function WPD_CONTACT_BUSINESS_POSTAL_ADDRESS_COUNTRY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFBD4FDAB, &H987D, &H4777, &HB3, &HF9, &H72, &H61, &H85, &HA9, &H31, &H2B, 23)
 WPD_CONTACT_BUSINESS_POSTAL_ADDRESS_COUNTRY = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_METHOD() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2D521CA8, &HC1B0, &H4268, &HA3, &H42, &HCF, &H19, &H32, &H15, &H69, &HBC, 1001)
 WPD_PROPERTY_SERVICE_METHOD = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_METHOD_PARAMETER_VALUES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2D521CA8, &HC1B0, &H4268, &HA3, &H42, &HCF, &H19, &H32, &H15, &H69, &HBC, 1002)
 WPD_PROPERTY_SERVICE_METHOD_PARAMETER_VALUES = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_METHOD_RESULT_VALUES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2D521CA8, &HC1B0, &H4268, &HA3, &H42, &HCF, &H19, &H32, &H15, &H69, &HBC, 1003)
 WPD_PROPERTY_SERVICE_METHOD_RESULT_VALUES = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_METHOD_CONTEXT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2D521CA8, &HC1B0, &H4268, &HA3, &H42, &HCF, &H19, &H32, &H15, &H69, &HBC, 1004)
 WPD_PROPERTY_SERVICE_METHOD_CONTEXT = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_METHOD_HRESULT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2D521CA8, &HC1B0, &H4268, &HA3, &H42, &HCF, &H19, &H32, &H15, &H69, &HBC, 1005)
 WPD_PROPERTY_SERVICE_METHOD_HRESULT = pkk
End Function
Public Function WPD_RESOURCE_DEFAULT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE81E79BE, &H34F0, &H41BF, &HB5, &H3F, &HF1, &HA0, &H6A, &HE8, &H78, &H42, 0)
 WPD_RESOURCE_DEFAULT = pkk
End Function
Public Function WPD_RESOURCE_CONTACT_PHOTO() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2C4D6803, &H80EA, &H4580, &HAF, &H9A, &H5B, &HE1, &HA2, &H3E, &HDD, &HCB, 0)
 WPD_RESOURCE_CONTACT_PHOTO = pkk
End Function
Public Function WPD_RESOURCE_THUMBNAIL() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC7C407BA, &H98FA, &H46B5, &H99, &H60, &H23, &HFE, &HC1, &H24, &HCF, &HDE, 0)
 WPD_RESOURCE_THUMBNAIL = pkk
End Function
Public Function WPD_RESOURCE_ICON() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF195FED8, &HAA28, &H4EE3, &HB1, &H53, &HE1, &H82, &HDD, &H5E, &HDC, &H39, 0)
 WPD_RESOURCE_ICON = pkk
End Function
Public Function WPD_RESOURCE_AUDIO_CLIP() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3BC13982, &H85B1, &H48E0, &H95, &HA6, &H8D, &H3A, &HD0, &H6B, &HE1, &H17, 0)
 WPD_RESOURCE_AUDIO_CLIP = pkk
End Function
Public Function WPD_RESOURCE_ALBUM_ART() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF02AA354, &H2300, &H4E2D, &HA1, &HB9, &H3B, &H67, &H30, &HF7, &HFA, &H21, 0)
 WPD_RESOURCE_ALBUM_ART = pkk
End Function
Public Function WPD_RESOURCE_GENERIC() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB9B9F515, &HBA70, &H4647, &H94, &HDC, &HFA, &H49, &H25, &HE9, &H5A, &H7, 0)
 WPD_RESOURCE_GENERIC = pkk
End Function
Public Function WPD_RESOURCE_VIDEO_CLIP() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB566EE42, &H6368, &H4290, &H86, &H62, &H70, &H18, &H2F, &HB7, &H9F, &H20, 0)
 WPD_RESOURCE_VIDEO_CLIP = pkk
End Function
Public Function WPD_RESOURCE_BRANDING_ART() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB633B1AE, &H6CAF, &H4A87, &H95, &H89, &H22, &HDE, &HD6, &HDD, &H58, &H99, 0)
 WPD_RESOURCE_BRANDING_ART = pkk
End Function
Public Function WPD_OBJECT_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 2)
 WPD_OBJECT_ID = pkk
End Function
Public Function WPD_OBJECT_PARENT_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 3)
 WPD_OBJECT_PARENT_ID = pkk
End Function
Public Function WPD_OBJECT_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 4)
 WPD_OBJECT_NAME = pkk
End Function
Public Function WPD_OBJECT_PERSISTENT_UNIQUE_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 5)
 WPD_OBJECT_PERSISTENT_UNIQUE_ID = pkk
End Function
Public Function WPD_OBJECT_FORMAT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 6)
 WPD_OBJECT_FORMAT = pkk
End Function
Public Function WPD_PROPERTY_CLASS_EXTENSION_DEVICE_INFORMATION_VALUES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H33FB0D11, &H64A3, &H4FAC, &HB4, &HC7, &H3D, &HFE, &HAA, &H99, &HB0, &H51, 1001)
 WPD_PROPERTY_CLASS_EXTENSION_DEVICE_INFORMATION_VALUES = pkk
End Function
Public Function WPD_PROPERTY_CLASS_EXTENSION_DEVICE_INFORMATION_WRITE_RESULTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H33FB0D11, &H64A3, &H4FAC, &HB4, &HC7, &H3D, &HFE, &HAA, &H99, &HB0, &H51, 1002)
 WPD_PROPERTY_CLASS_EXTENSION_DEVICE_INFORMATION_WRITE_RESULTS = pkk
End Function
Public Function WPD_COMMAND_CLASS_EXTENSION_REGISTER_SERVICE_INTERFACES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7F0779B5, &HFA2B, &H4766, &H9C, &HB2, &HF7, &H3B, &HA3, &HB, &H67, &H58, 2)
 WPD_COMMAND_CLASS_EXTENSION_REGISTER_SERVICE_INTERFACES = pkk
End Function
Public Function WPD_COMMAND_CLASS_EXTENSION_UNREGISTER_SERVICE_INTERFACES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7F0779B5, &HFA2B, &H4766, &H9C, &HB2, &HF7, &H3B, &HA3, &HB, &H67, &H58, 3)
 WPD_COMMAND_CLASS_EXTENSION_UNREGISTER_SERVICE_INTERFACES = pkk
End Function
Public Function WPD_PROPERTY_CLASS_EXTENSION_SERVICE_OBJECT_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7F0779B5, &HFA2B, &H4766, &H9C, &HB2, &HF7, &H3B, &HA3, &HB, &H67, &H58, 1001)
 WPD_PROPERTY_CLASS_EXTENSION_SERVICE_OBJECT_ID = pkk
End Function
Public Function WPD_PROPERTY_CLASS_EXTENSION_SERVICE_INTERFACES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7F0779B5, &HFA2B, &H4766, &H9C, &HB2, &HF7, &H3B, &HA3, &HB, &H67, &H58, 1002)
 WPD_PROPERTY_CLASS_EXTENSION_SERVICE_INTERFACES = pkk
End Function
Public Function WPD_PROPERTY_CLASS_EXTENSION_SERVICE_REGISTRATION_RESULTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7F0779B5, &HFA2B, &H4766, &H9C, &HB2, &HF7, &H3B, &HA3, &HB, &H67, &H58, 1003)
 WPD_PROPERTY_CLASS_EXTENSION_SERVICE_REGISTRATION_RESULTS = pkk
End Function
Public Function WPD_COMMAND_GENERATE_KEYPAIR() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78F9C6FC, &H79B8, &H473C, &H90, &H60, &H6B, &HD2, &H3D, &HD0, &H72, &HC4, 2)
 WPD_COMMAND_GENERATE_KEYPAIR = pkk
End Function
Public Function WPD_COMMAND_COMMIT_KEYPAIR() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78F9C6FC, &H79B8, &H473C, &H90, &H60, &H6B, &HD2, &H3D, &HD0, &H72, &HC4, 3)
 WPD_COMMAND_COMMIT_KEYPAIR = pkk
End Function
Public Function WPD_COMMAND_PROCESS_WIRELESS_PROFILE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78F9C6FC, &H79B8, &H473C, &H90, &H60, &H6B, &HD2, &H3D, &HD0, &H72, &HC4, 4)
 WPD_COMMAND_PROCESS_WIRELESS_PROFILE = pkk
End Function
Public Function WPD_PROPERTY_PUBLIC_KEY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78F9C6FC, &H79B8, &H473C, &H90, &H60, &H6B, &HD2, &H3D, &HD0, &H72, &HC4, 1001)
 WPD_PROPERTY_PUBLIC_KEY = pkk
End Function
Public Function WPD_COMMAND_SERVICE_COMMON_GET_SERVICE_OBJECT_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H322F071D, &H36EF, &H477F, &HB4, &HB5, &H6F, &H52, &HD7, &H34, &HBA, &HEE, 2)
 WPD_COMMAND_SERVICE_COMMON_GET_SERVICE_OBJECT_ID = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_OBJECT_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H322F071D, &H36EF, &H477F, &HB4, &HB5, &H6F, &H52, &HD7, &H34, &HBA, &HEE, 1001)
 WPD_PROPERTY_SERVICE_OBJECT_ID = pkk
End Function
Public Function WPD_COMMAND_SERVICE_CAPABILITIES_GET_SUPPORTED_METHODS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 2)
 WPD_COMMAND_SERVICE_CAPABILITIES_GET_SUPPORTED_METHODS = pkk
End Function
Public Function WPD_COMMAND_SERVICE_CAPABILITIES_GET_SUPPORTED_METHODS_BY_FORMAT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 3)
 WPD_COMMAND_SERVICE_CAPABILITIES_GET_SUPPORTED_METHODS_BY_FORMAT = pkk
End Function
Public Function WPD_COMMAND_SERVICE_CAPABILITIES_GET_METHOD_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 4)
 WPD_COMMAND_SERVICE_CAPABILITIES_GET_METHOD_ATTRIBUTES = pkk
End Function
Public Function WPD_COMMAND_SERVICE_CAPABILITIES_GET_METHOD_PARAMETER_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 5)
 WPD_COMMAND_SERVICE_CAPABILITIES_GET_METHOD_PARAMETER_ATTRIBUTES = pkk
End Function
Public Function WPD_COMMAND_SERVICE_CAPABILITIES_GET_SUPPORTED_FORMATS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 6)
 WPD_COMMAND_SERVICE_CAPABILITIES_GET_SUPPORTED_FORMATS = pkk
End Function
Public Function WPD_COMMAND_SERVICE_CAPABILITIES_GET_FORMAT_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 7)
 WPD_COMMAND_SERVICE_CAPABILITIES_GET_FORMAT_ATTRIBUTES = pkk
End Function
Public Function WPD_COMMAND_SERVICE_CAPABILITIES_GET_SUPPORTED_FORMAT_PROPERTIES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 8)
 WPD_COMMAND_SERVICE_CAPABILITIES_GET_SUPPORTED_FORMAT_PROPERTIES = pkk
End Function
Public Function WPD_COMMAND_SERVICE_CAPABILITIES_GET_FORMAT_PROPERTY_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 9)
 WPD_COMMAND_SERVICE_CAPABILITIES_GET_FORMAT_PROPERTY_ATTRIBUTES = pkk
End Function
Public Function WPD_COMMAND_SERVICE_CAPABILITIES_GET_SUPPORTED_EVENTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 10)
 WPD_COMMAND_SERVICE_CAPABILITIES_GET_SUPPORTED_EVENTS = pkk
End Function
Public Function WPD_COMMAND_SERVICE_CAPABILITIES_GET_EVENT_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 11)
 WPD_COMMAND_SERVICE_CAPABILITIES_GET_EVENT_ATTRIBUTES = pkk
End Function
Public Function WPD_COMMAND_SERVICE_CAPABILITIES_GET_EVENT_PARAMETER_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 12)
 WPD_COMMAND_SERVICE_CAPABILITIES_GET_EVENT_PARAMETER_ATTRIBUTES = pkk
End Function
Public Function WPD_COMMAND_SERVICE_CAPABILITIES_GET_INHERITED_SERVICES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 13)
 WPD_COMMAND_SERVICE_CAPABILITIES_GET_INHERITED_SERVICES = pkk
End Function
Public Function WPD_COMMAND_SERVICE_CAPABILITIES_GET_FORMAT_RENDERING_PROFILES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 14)
 WPD_COMMAND_SERVICE_CAPABILITIES_GET_FORMAT_RENDERING_PROFILES = pkk
End Function
Public Function WPD_COMMAND_SERVICE_CAPABILITIES_GET_SUPPORTED_COMMANDS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 15)
 WPD_COMMAND_SERVICE_CAPABILITIES_GET_SUPPORTED_COMMANDS = pkk
End Function
Public Function WPD_COMMAND_SERVICE_CAPABILITIES_GET_COMMAND_OPTIONS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 16)
 WPD_COMMAND_SERVICE_CAPABILITIES_GET_COMMAND_OPTIONS = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_SUPPORTED_METHODS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1001)
 WPD_PROPERTY_SERVICE_CAPABILITIES_SUPPORTED_METHODS = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_FORMAT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1002)
 WPD_PROPERTY_SERVICE_CAPABILITIES_FORMAT = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_METHOD() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1003)
 WPD_PROPERTY_SERVICE_CAPABILITIES_METHOD = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_METHOD_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1004)
 WPD_PROPERTY_SERVICE_CAPABILITIES_METHOD_ATTRIBUTES = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_PARAMETER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1005)
 WPD_PROPERTY_SERVICE_CAPABILITIES_PARAMETER = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_PARAMETER_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1006)
 WPD_PROPERTY_SERVICE_CAPABILITIES_PARAMETER_ATTRIBUTES = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_FORMATS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1007)
 WPD_PROPERTY_SERVICE_CAPABILITIES_FORMATS = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_FORMAT_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1008)
 WPD_PROPERTY_SERVICE_CAPABILITIES_FORMAT_ATTRIBUTES = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_PROPERTY_KEYS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1009)
 WPD_PROPERTY_SERVICE_CAPABILITIES_PROPERTY_KEYS = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_PROPERTY_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1010)
 WPD_PROPERTY_SERVICE_CAPABILITIES_PROPERTY_ATTRIBUTES = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_SUPPORTED_EVENTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1011)
 WPD_PROPERTY_SERVICE_CAPABILITIES_SUPPORTED_EVENTS = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_EVENT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1012)
 WPD_PROPERTY_SERVICE_CAPABILITIES_EVENT = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_EVENT_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1013)
 WPD_PROPERTY_SERVICE_CAPABILITIES_EVENT_ATTRIBUTES = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_INHERITANCE_TYPE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1014)
 WPD_PROPERTY_SERVICE_CAPABILITIES_INHERITANCE_TYPE = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_INHERITED_SERVICES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1015)
 WPD_PROPERTY_SERVICE_CAPABILITIES_INHERITED_SERVICES = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_RENDERING_PROFILES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1016)
 WPD_PROPERTY_SERVICE_CAPABILITIES_RENDERING_PROFILES = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_SUPPORTED_COMMANDS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1017)
 WPD_PROPERTY_SERVICE_CAPABILITIES_SUPPORTED_COMMANDS = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_COMMAND() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1018)
 WPD_PROPERTY_SERVICE_CAPABILITIES_COMMAND = pkk
End Function
Public Function WPD_PROPERTY_SERVICE_CAPABILITIES_COMMAND_OPTIONS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H24457E74, &H2E9F, &H44F9, &H8C, &H57, &H1D, &H1B, &HCB, &H17, &HB, &H89, 1019)
 WPD_PROPERTY_SERVICE_CAPABILITIES_COMMAND_OPTIONS = pkk
End Function
Public Function WPD_COMMAND_SERVICE_METHODS_START_INVOKE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2D521CA8, &HC1B0, &H4268, &HA3, &H42, &HCF, &H19, &H32, &H15, &H69, &HBC, 2)
 WPD_COMMAND_SERVICE_METHODS_START_INVOKE = pkk
End Function
Public Function WPD_COMMAND_SERVICE_METHODS_CANCEL_INVOKE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2D521CA8, &HC1B0, &H4268, &HA3, &H42, &HCF, &H19, &H32, &H15, &H69, &HBC, 3)
 WPD_COMMAND_SERVICE_METHODS_CANCEL_INVOKE = pkk
End Function
Public Function WPD_COMMAND_SERVICE_METHODS_END_INVOKE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2D521CA8, &HC1B0, &H4268, &HA3, &H42, &HCF, &H19, &H32, &H15, &H69, &HBC, 4)
 WPD_COMMAND_SERVICE_METHODS_END_INVOKE = pkk
End Function
Public Function WPD_COMMAND_OBJECT_RESOURCES_REVERT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 10)
 WPD_COMMAND_OBJECT_RESOURCES_REVERT = pkk
End Function
Public Function WPD_COMMAND_OBJECT_RESOURCES_SEEK() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 11)
 WPD_COMMAND_OBJECT_RESOURCES_SEEK = pkk
End Function
Public Function WPD_COMMAND_OBJECT_RESOURCES_COMMIT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 12)
 WPD_COMMAND_OBJECT_RESOURCES_COMMIT = pkk
End Function
Public Function WPD_COMMAND_OBJECT_RESOURCES_SEEK_IN_UNITS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 13)
 WPD_COMMAND_OBJECT_RESOURCES_SEEK_IN_UNITS = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_RESOURCES_OBJECT_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 1001)
 WPD_PROPERTY_OBJECT_RESOURCES_OBJECT_ID = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_RESOURCES_ACCESS_MODE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 1002)
 WPD_PROPERTY_OBJECT_RESOURCES_ACCESS_MODE = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_RESOURCES_RESOURCE_KEYS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 1003)
 WPD_PROPERTY_OBJECT_RESOURCES_RESOURCE_KEYS = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_RESOURCES_RESOURCE_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 1004)
 WPD_PROPERTY_OBJECT_RESOURCES_RESOURCE_ATTRIBUTES = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_RESOURCES_CONTEXT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 1005)
 WPD_PROPERTY_OBJECT_RESOURCES_CONTEXT = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_RESOURCES_NUM_BYTES_TO_READ() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 1006)
 WPD_PROPERTY_OBJECT_RESOURCES_NUM_BYTES_TO_READ = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_RESOURCES_NUM_BYTES_READ() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 1007)
 WPD_PROPERTY_OBJECT_RESOURCES_NUM_BYTES_READ = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_RESOURCES_NUM_BYTES_TO_WRITE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 1008)
 WPD_PROPERTY_OBJECT_RESOURCES_NUM_BYTES_TO_WRITE = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_RESOURCES_NUM_BYTES_WRITTEN() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 1009)
 WPD_PROPERTY_OBJECT_RESOURCES_NUM_BYTES_WRITTEN = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_RESOURCES_DATA() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 1010)
 WPD_PROPERTY_OBJECT_RESOURCES_DATA = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_RESOURCES_OPTIMAL_TRANSFER_BUFFER_SIZE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 1011)
 WPD_PROPERTY_OBJECT_RESOURCES_OPTIMAL_TRANSFER_BUFFER_SIZE = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_RESOURCES_SEEK_OFFSET() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 1012)
 WPD_PROPERTY_OBJECT_RESOURCES_SEEK_OFFSET = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_RESOURCES_SEEK_ORIGIN_FLAG() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 1013)
 WPD_PROPERTY_OBJECT_RESOURCES_SEEK_ORIGIN_FLAG = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_RESOURCES_POSITION_FROM_START() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 1014)
 WPD_PROPERTY_OBJECT_RESOURCES_POSITION_FROM_START = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_RESOURCES_SUPPORTS_UNITS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 1015)
 WPD_PROPERTY_OBJECT_RESOURCES_SUPPORTS_UNITS = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_RESOURCES_STREAM_UNITS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 1016)
 WPD_PROPERTY_OBJECT_RESOURCES_STREAM_UNITS = pkk
End Function
Public Function WPD_OPTION_OBJECT_RESOURCES_SEEK_ON_READ_SUPPORTED() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 5001)
 WPD_OPTION_OBJECT_RESOURCES_SEEK_ON_READ_SUPPORTED = pkk
End Function
Public Function WPD_OPTION_OBJECT_RESOURCES_SEEK_ON_WRITE_SUPPORTED() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 5002)
 WPD_OPTION_OBJECT_RESOURCES_SEEK_ON_WRITE_SUPPORTED = pkk
End Function
Public Function WPD_OPTION_OBJECT_RESOURCES_NO_INPUT_BUFFER_ON_READ() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 5003)
 WPD_OPTION_OBJECT_RESOURCES_NO_INPUT_BUFFER_ON_READ = pkk
End Function
Public Function WPD_COMMAND_OBJECT_MANAGEMENT_CREATE_OBJECT_WITH_PROPERTIES_ONLY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 2)
 WPD_COMMAND_OBJECT_MANAGEMENT_CREATE_OBJECT_WITH_PROPERTIES_ONLY = pkk
End Function
Public Function WPD_COMMAND_OBJECT_MANAGEMENT_CREATE_OBJECT_WITH_PROPERTIES_AND_DATA() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 3)
 WPD_COMMAND_OBJECT_MANAGEMENT_CREATE_OBJECT_WITH_PROPERTIES_AND_DATA = pkk
End Function
Public Function WPD_COMMAND_OBJECT_MANAGEMENT_WRITE_OBJECT_DATA() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 4)
 WPD_COMMAND_OBJECT_MANAGEMENT_WRITE_OBJECT_DATA = pkk
End Function
Public Function WPD_COMMAND_OBJECT_MANAGEMENT_COMMIT_OBJECT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 5)
 WPD_COMMAND_OBJECT_MANAGEMENT_COMMIT_OBJECT = pkk
End Function
Public Function WPD_COMMAND_OBJECT_MANAGEMENT_REVERT_OBJECT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 6)
 WPD_COMMAND_OBJECT_MANAGEMENT_REVERT_OBJECT = pkk
End Function
Public Function WPD_COMMAND_OBJECT_MANAGEMENT_DELETE_OBJECTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 7)
 WPD_COMMAND_OBJECT_MANAGEMENT_DELETE_OBJECTS = pkk
End Function
Public Function WPD_COMMAND_OBJECT_MANAGEMENT_MOVE_OBJECTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 8)
 WPD_COMMAND_OBJECT_MANAGEMENT_MOVE_OBJECTS = pkk
End Function
Public Function WPD_COMMAND_OBJECT_MANAGEMENT_COPY_OBJECTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 9)
 WPD_COMMAND_OBJECT_MANAGEMENT_COPY_OBJECTS = pkk
End Function
Public Function WPD_COMMAND_OBJECT_MANAGEMENT_UPDATE_OBJECT_WITH_PROPERTIES_AND_DATA() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 10)
 WPD_COMMAND_OBJECT_MANAGEMENT_UPDATE_OBJECT_WITH_PROPERTIES_AND_DATA = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_MANAGEMENT_CREATION_PROPERTIES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 1001)
 WPD_PROPERTY_OBJECT_MANAGEMENT_CREATION_PROPERTIES = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_MANAGEMENT_CONTEXT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 1002)
 WPD_PROPERTY_OBJECT_MANAGEMENT_CONTEXT = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_MANAGEMENT_NUM_BYTES_TO_WRITE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 1003)
 WPD_PROPERTY_OBJECT_MANAGEMENT_NUM_BYTES_TO_WRITE = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_MANAGEMENT_NUM_BYTES_WRITTEN() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 1004)
 WPD_PROPERTY_OBJECT_MANAGEMENT_NUM_BYTES_WRITTEN = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_MANAGEMENT_DATA() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 1005)
 WPD_PROPERTY_OBJECT_MANAGEMENT_DATA = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_MANAGEMENT_OBJECT_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 1006)
 WPD_PROPERTY_OBJECT_MANAGEMENT_OBJECT_ID = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_MANAGEMENT_DELETE_OPTIONS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 1007)
 WPD_PROPERTY_OBJECT_MANAGEMENT_DELETE_OPTIONS = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_MANAGEMENT_OPTIMAL_TRANSFER_BUFFER_SIZE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 1008)
 WPD_PROPERTY_OBJECT_MANAGEMENT_OPTIMAL_TRANSFER_BUFFER_SIZE = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_MANAGEMENT_OBJECT_IDS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 1009)
 WPD_PROPERTY_OBJECT_MANAGEMENT_OBJECT_IDS = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_MANAGEMENT_DELETE_RESULTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 1010)
 WPD_PROPERTY_OBJECT_MANAGEMENT_DELETE_RESULTS = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_MANAGEMENT_DESTINATION_FOLDER_OBJECT_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 1011)
 WPD_PROPERTY_OBJECT_MANAGEMENT_DESTINATION_FOLDER_OBJECT_ID = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_MANAGEMENT_MOVE_RESULTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 1012)
 WPD_PROPERTY_OBJECT_MANAGEMENT_MOVE_RESULTS = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_MANAGEMENT_COPY_RESULTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 1013)
 WPD_PROPERTY_OBJECT_MANAGEMENT_COPY_RESULTS = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_MANAGEMENT_UPDATE_PROPERTIES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 1014)
 WPD_PROPERTY_OBJECT_MANAGEMENT_UPDATE_PROPERTIES = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_MANAGEMENT_PROPERTY_KEYS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 1015)
 WPD_PROPERTY_OBJECT_MANAGEMENT_PROPERTY_KEYS = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_MANAGEMENT_OBJECT_FORMAT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 1016)
 WPD_PROPERTY_OBJECT_MANAGEMENT_OBJECT_FORMAT = pkk
End Function
Public Function WPD_OPTION_OBJECT_MANAGEMENT_RECURSIVE_DELETE_SUPPORTED() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1E43DD, &HA9ED, &H4341, &H8B, &HCC, &H18, &H61, &H92, &HAE, &HA0, &H89, 5001)
 WPD_OPTION_OBJECT_MANAGEMENT_RECURSIVE_DELETE_SUPPORTED = pkk
End Function
Public Function WPD_COMMAND_CAPABILITIES_GET_SUPPORTED_COMMANDS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 2)
 WPD_COMMAND_CAPABILITIES_GET_SUPPORTED_COMMANDS = pkk
End Function
Public Function WPD_COMMAND_CAPABILITIES_GET_COMMAND_OPTIONS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 3)
 WPD_COMMAND_CAPABILITIES_GET_COMMAND_OPTIONS = pkk
End Function
Public Function WPD_COMMAND_CAPABILITIES_GET_SUPPORTED_FUNCTIONAL_CATEGORIES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 4)
 WPD_COMMAND_CAPABILITIES_GET_SUPPORTED_FUNCTIONAL_CATEGORIES = pkk
End Function
Public Function WPD_COMMAND_CAPABILITIES_GET_FUNCTIONAL_OBJECTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 5)
 WPD_COMMAND_CAPABILITIES_GET_FUNCTIONAL_OBJECTS = pkk
End Function
Public Function WPD_COMMAND_CAPABILITIES_GET_SUPPORTED_CONTENT_TYPES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 6)
 WPD_COMMAND_CAPABILITIES_GET_SUPPORTED_CONTENT_TYPES = pkk
End Function
Public Function WPD_COMMAND_CAPABILITIES_GET_SUPPORTED_FORMATS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 7)
 WPD_COMMAND_CAPABILITIES_GET_SUPPORTED_FORMATS = pkk
End Function
Public Function WPD_COMMAND_CAPABILITIES_GET_SUPPORTED_FORMAT_PROPERTIES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 8)
 WPD_COMMAND_CAPABILITIES_GET_SUPPORTED_FORMAT_PROPERTIES = pkk
End Function
Public Function WPD_COMMAND_CAPABILITIES_GET_FIXED_PROPERTY_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 9)
 WPD_COMMAND_CAPABILITIES_GET_FIXED_PROPERTY_ATTRIBUTES = pkk
End Function
Public Function WPD_COMMAND_CAPABILITIES_GET_SUPPORTED_EVENTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 10)
 WPD_COMMAND_CAPABILITIES_GET_SUPPORTED_EVENTS = pkk
End Function
Public Function WPD_COMMAND_CAPABILITIES_GET_EVENT_OPTIONS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 11)
 WPD_COMMAND_CAPABILITIES_GET_EVENT_OPTIONS = pkk
End Function
Public Function WPD_PROPERTY_CAPABILITIES_SUPPORTED_COMMANDS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 1001)
 WPD_PROPERTY_CAPABILITIES_SUPPORTED_COMMANDS = pkk
End Function
Public Function WPD_PROPERTY_CAPABILITIES_COMMAND() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 1002)
 WPD_PROPERTY_CAPABILITIES_COMMAND = pkk
End Function
Public Function WPD_PROPERTY_CAPABILITIES_COMMAND_OPTIONS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 1003)
 WPD_PROPERTY_CAPABILITIES_COMMAND_OPTIONS = pkk
End Function
Public Function WPD_PROPERTY_CAPABILITIES_FUNCTIONAL_CATEGORIES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 1004)
 WPD_PROPERTY_CAPABILITIES_FUNCTIONAL_CATEGORIES = pkk
End Function
Public Function WPD_PROPERTY_CAPABILITIES_FUNCTIONAL_CATEGORY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 1005)
 WPD_PROPERTY_CAPABILITIES_FUNCTIONAL_CATEGORY = pkk
End Function
Public Function WPD_PROPERTY_CAPABILITIES_FUNCTIONAL_OBJECTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 1006)
 WPD_PROPERTY_CAPABILITIES_FUNCTIONAL_OBJECTS = pkk
End Function
Public Function WPD_PROPERTY_CAPABILITIES_CONTENT_TYPES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 1007)
 WPD_PROPERTY_CAPABILITIES_CONTENT_TYPES = pkk
End Function
Public Function WPD_PROPERTY_CAPABILITIES_CONTENT_TYPE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 1008)
 WPD_PROPERTY_CAPABILITIES_CONTENT_TYPE = pkk
End Function
Public Function WPD_PROPERTY_CAPABILITIES_FORMATS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 1009)
 WPD_PROPERTY_CAPABILITIES_FORMATS = pkk
End Function
Public Function WPD_PROPERTY_CAPABILITIES_FORMAT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 1010)
 WPD_PROPERTY_CAPABILITIES_FORMAT = pkk
End Function
Public Function WPD_PROPERTY_CAPABILITIES_PROPERTY_KEYS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 1011)
 WPD_PROPERTY_CAPABILITIES_PROPERTY_KEYS = pkk
End Function
Public Function WPD_PROPERTY_CAPABILITIES_PROPERTY_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 1012)
 WPD_PROPERTY_CAPABILITIES_PROPERTY_ATTRIBUTES = pkk
End Function
Public Function WPD_PROPERTY_CAPABILITIES_SUPPORTED_EVENTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 1013)
 WPD_PROPERTY_CAPABILITIES_SUPPORTED_EVENTS = pkk
End Function
Public Function WPD_PROPERTY_CAPABILITIES_EVENT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 1014)
 WPD_PROPERTY_CAPABILITIES_EVENT = pkk
End Function
Public Function WPD_PROPERTY_CAPABILITIES_EVENT_OPTIONS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCABEC78, &H6B74, &H41C6, &H92, &H16, &H26, &H39, &HD1, &HFC, &HE3, &H56, 1015)
 WPD_PROPERTY_CAPABILITIES_EVENT_OPTIONS = pkk
End Function
Public Function WPD_COMMAND_STORAGE_FORMAT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD8F907A6, &H34CC, &H45FA, &H97, &HFB, &HD0, &H7, &HFA, &H47, &HEC, &H94, 2)
 WPD_COMMAND_STORAGE_FORMAT = pkk
End Function
Public Function WPD_COMMAND_STORAGE_EJECT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD8F907A6, &H34CC, &H45FA, &H97, &HFB, &HD0, &H7, &HFA, &H47, &HEC, &H94, 4)
 WPD_COMMAND_STORAGE_EJECT = pkk
End Function
Public Function WPD_PROPERTY_STORAGE_OBJECT_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD8F907A6, &H34CC, &H45FA, &H97, &HFB, &HD0, &H7, &HFA, &H47, &HEC, &H94, 1001)
 WPD_PROPERTY_STORAGE_OBJECT_ID = pkk
End Function
Public Function WPD_PROPERTY_STORAGE_DESTINATION_OBJECT_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD8F907A6, &H34CC, &H45FA, &H97, &HFB, &HD0, &H7, &HFA, &H47, &HEC, &H94, 1002)
 WPD_PROPERTY_STORAGE_DESTINATION_OBJECT_ID = pkk
End Function
Public Function WPD_COMMAND_SMS_SEND() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAFC25D66, &HFE0D, &H4114, &H90, &H97, &H97, &HC, &H93, &HE9, &H20, &HD1, 2)
 WPD_COMMAND_SMS_SEND = pkk
End Function
Public Function WPD_PROPERTY_SMS_RECIPIENT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAFC25D66, &HFE0D, &H4114, &H90, &H97, &H97, &HC, &H93, &HE9, &H20, &HD1, 1001)
 WPD_PROPERTY_SMS_RECIPIENT = pkk
End Function
Public Function WPD_PROPERTY_SMS_MESSAGE_TYPE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAFC25D66, &HFE0D, &H4114, &H90, &H97, &H97, &HC, &H93, &HE9, &H20, &HD1, 1002)
 WPD_PROPERTY_SMS_MESSAGE_TYPE = pkk
End Function
Public Function WPD_PROPERTY_SMS_TEXT_MESSAGE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAFC25D66, &HFE0D, &H4114, &H90, &H97, &H97, &HC, &H93, &HE9, &H20, &HD1, 1003)
 WPD_PROPERTY_SMS_TEXT_MESSAGE = pkk
End Function
Public Function WPD_PROPERTY_SMS_BINARY_MESSAGE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAFC25D66, &HFE0D, &H4114, &H90, &H97, &H97, &HC, &H93, &HE9, &H20, &HD1, 1004)
 WPD_PROPERTY_SMS_BINARY_MESSAGE = pkk
End Function
Public Function WPD_OPTION_SMS_BINARY_MESSAGE_SUPPORTED() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAFC25D66, &HFE0D, &H4114, &H90, &H97, &H97, &HC, &H93, &HE9, &H20, &HD1, 5001)
 WPD_OPTION_SMS_BINARY_MESSAGE_SUPPORTED = pkk
End Function
Public Function WPD_COMMAND_STILL_IMAGE_CAPTURE_INITIATE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4FCD6982, &H22A2, &H4B05, &HA4, &H8B, &H62, &HD3, &H8B, &HF2, &H7B, &H32, 2)
 WPD_COMMAND_STILL_IMAGE_CAPTURE_INITIATE = pkk
End Function
Public Function WPD_COMMAND_MEDIA_CAPTURE_START() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H59B433BA, &HFE44, &H4D8D, &H80, &H8C, &H6B, &HCB, &H9B, &HF, &H15, &HE8, 2)
 WPD_COMMAND_MEDIA_CAPTURE_START = pkk
End Function
Public Function WPD_COMMAND_MEDIA_CAPTURE_STOP() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H59B433BA, &HFE44, &H4D8D, &H80, &H8C, &H6B, &HCB, &H9B, &HF, &H15, &HE8, 3)
 WPD_COMMAND_MEDIA_CAPTURE_STOP = pkk
End Function
Public Function WPD_COMMAND_MEDIA_CAPTURE_PAUSE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H59B433BA, &HFE44, &H4D8D, &H80, &H8C, &H6B, &HCB, &H9B, &HF, &H15, &HE8, 4)
 WPD_COMMAND_MEDIA_CAPTURE_PAUSE = pkk
End Function
Public Function WPD_COMMAND_DEVICE_HINTS_GET_CONTENT_LOCATION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5FB92B, &HCB46, &H4C4F, &H83, &H43, &HB, &HC3, &HD3, &HF1, &H7C, &H84, 2)
 WPD_COMMAND_DEVICE_HINTS_GET_CONTENT_LOCATION = pkk
End Function
Public Function WPD_PROPERTY_DEVICE_HINTS_CONTENT_TYPE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5FB92B, &HCB46, &H4C4F, &H83, &H43, &HB, &HC3, &HD3, &HF1, &H7C, &H84, 1001)
 WPD_PROPERTY_DEVICE_HINTS_CONTENT_TYPE = pkk
End Function
Public Function WPD_PROPERTY_DEVICE_HINTS_CONTENT_LOCATIONS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5FB92B, &HCB46, &H4C4F, &H83, &H43, &HB, &HC3, &HD3, &HF1, &H7C, &H84, 1002)
 WPD_PROPERTY_DEVICE_HINTS_CONTENT_LOCATIONS = pkk
End Function
Public Function WPD_COMMAND_CLASS_EXTENSION_WRITE_DEVICE_INFORMATION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H33FB0D11, &H64A3, &H4FAC, &HB4, &HC7, &H3D, &HFE, &HAA, &H99, &HB0, &H51, 2)
 WPD_COMMAND_CLASS_EXTENSION_WRITE_DEVICE_INFORMATION = pkk
End Function
Public Function WPD_PARAMETER_ATTRIBUTE_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE6864DD7, &HF325, &H45EA, &HA1, &HD5, &H97, &HCF, &H73, &HB6, &HCA, &H58, 13)
 WPD_PARAMETER_ATTRIBUTE_NAME = pkk
End Function
Public Function WPD_COMMAND_COMMON_RESET_DEVICE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF0422A9C, &H5DC8, &H4440, &HB5, &HBD, &H5D, &HF2, &H88, &H35, &H65, &H8A, 2)
 WPD_COMMAND_COMMON_RESET_DEVICE = pkk
End Function
Public Function WPD_COMMAND_COMMON_GET_OBJECT_IDS_FROM_PERSISTENT_UNIQUE_IDS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF0422A9C, &H5DC8, &H4440, &HB5, &HBD, &H5D, &HF2, &H88, &H35, &H65, &H8A, 3)
 WPD_COMMAND_COMMON_GET_OBJECT_IDS_FROM_PERSISTENT_UNIQUE_IDS = pkk
End Function
Public Function WPD_COMMAND_COMMON_SAVE_CLIENT_INFORMATION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF0422A9C, &H5DC8, &H4440, &HB5, &HBD, &H5D, &HF2, &H88, &H35, &H65, &H8A, 4)
 WPD_COMMAND_COMMON_SAVE_CLIENT_INFORMATION = pkk
End Function
Public Function WPD_PROPERTY_COMMON_COMMAND_CATEGORY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF0422A9C, &H5DC8, &H4440, &HB5, &HBD, &H5D, &HF2, &H88, &H35, &H65, &H8A, 1001)
 WPD_PROPERTY_COMMON_COMMAND_CATEGORY = pkk
End Function
Public Function WPD_PROPERTY_COMMON_COMMAND_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF0422A9C, &H5DC8, &H4440, &HB5, &HBD, &H5D, &HF2, &H88, &H35, &H65, &H8A, 1002)
 WPD_PROPERTY_COMMON_COMMAND_ID = pkk
End Function
Public Function WPD_PROPERTY_COMMON_HRESULT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF0422A9C, &H5DC8, &H4440, &HB5, &HBD, &H5D, &HF2, &H88, &H35, &H65, &H8A, 1003)
 WPD_PROPERTY_COMMON_HRESULT = pkk
End Function
Public Function WPD_PROPERTY_COMMON_DRIVER_ERROR_CODE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF0422A9C, &H5DC8, &H4440, &HB5, &HBD, &H5D, &HF2, &H88, &H35, &H65, &H8A, 1004)
 WPD_PROPERTY_COMMON_DRIVER_ERROR_CODE = pkk
End Function
Public Function WPD_PROPERTY_COMMON_COMMAND_TARGET() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF0422A9C, &H5DC8, &H4440, &HB5, &HBD, &H5D, &HF2, &H88, &H35, &H65, &H8A, 1006)
 WPD_PROPERTY_COMMON_COMMAND_TARGET = pkk
End Function
Public Function WPD_PROPERTY_COMMON_PERSISTENT_UNIQUE_IDS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF0422A9C, &H5DC8, &H4440, &HB5, &HBD, &H5D, &HF2, &H88, &H35, &H65, &H8A, 1007)
 WPD_PROPERTY_COMMON_PERSISTENT_UNIQUE_IDS = pkk
End Function
Public Function WPD_PROPERTY_COMMON_OBJECT_IDS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF0422A9C, &H5DC8, &H4440, &HB5, &HBD, &H5D, &HF2, &H88, &H35, &H65, &H8A, 1008)
 WPD_PROPERTY_COMMON_OBJECT_IDS = pkk
End Function
Public Function WPD_PROPERTY_COMMON_CLIENT_INFORMATION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF0422A9C, &H5DC8, &H4440, &HB5, &HBD, &H5D, &HF2, &H88, &H35, &H65, &H8A, 1009)
 WPD_PROPERTY_COMMON_CLIENT_INFORMATION = pkk
End Function
Public Function WPD_PROPERTY_COMMON_CLIENT_INFORMATION_CONTEXT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF0422A9C, &H5DC8, &H4440, &HB5, &HBD, &H5D, &HF2, &H88, &H35, &H65, &H8A, 1010)
 WPD_PROPERTY_COMMON_CLIENT_INFORMATION_CONTEXT = pkk
End Function
Public Function WPD_PROPERTY_COMMON_ACTIVITY_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF0422A9C, &H5DC8, &H4440, &HB5, &HBD, &H5D, &HF2, &H88, &H35, &H65, &H8A, 1011)
 WPD_PROPERTY_COMMON_ACTIVITY_ID = pkk
End Function
Public Function WPD_OPTION_VALID_OBJECT_IDS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF0422A9C, &H5DC8, &H4440, &HB5, &HBD, &H5D, &HF2, &H88, &H35, &H65, &H8A, 5001)
 WPD_OPTION_VALID_OBJECT_IDS = pkk
End Function
Public Function WPD_COMMAND_OBJECT_ENUMERATION_START_FIND() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB7474E91, &HE7F8, &H4AD9, &HB4, &H0, &HAD, &H1A, &H4B, &H58, &HEE, &HEC, 2)
 WPD_COMMAND_OBJECT_ENUMERATION_START_FIND = pkk
End Function
Public Function WPD_COMMAND_OBJECT_ENUMERATION_FIND_NEXT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB7474E91, &HE7F8, &H4AD9, &HB4, &H0, &HAD, &H1A, &H4B, &H58, &HEE, &HEC, 3)
 WPD_COMMAND_OBJECT_ENUMERATION_FIND_NEXT = pkk
End Function
Public Function WPD_COMMAND_OBJECT_ENUMERATION_END_FIND() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB7474E91, &HE7F8, &H4AD9, &HB4, &H0, &HAD, &H1A, &H4B, &H58, &HEE, &HEC, 4)
 WPD_COMMAND_OBJECT_ENUMERATION_END_FIND = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_ENUMERATION_PARENT_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB7474E91, &HE7F8, &H4AD9, &HB4, &H0, &HAD, &H1A, &H4B, &H58, &HEE, &HEC, 1001)
 WPD_PROPERTY_OBJECT_ENUMERATION_PARENT_ID = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_ENUMERATION_FILTER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB7474E91, &HE7F8, &H4AD9, &HB4, &H0, &HAD, &H1A, &H4B, &H58, &HEE, &HEC, 1002)
 WPD_PROPERTY_OBJECT_ENUMERATION_FILTER = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_ENUMERATION_OBJECT_IDS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB7474E91, &HE7F8, &H4AD9, &HB4, &H0, &HAD, &H1A, &H4B, &H58, &HEE, &HEC, 1003)
 WPD_PROPERTY_OBJECT_ENUMERATION_OBJECT_IDS = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_ENUMERATION_CONTEXT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB7474E91, &HE7F8, &H4AD9, &HB4, &H0, &HAD, &H1A, &H4B, &H58, &HEE, &HEC, 1004)
 WPD_PROPERTY_OBJECT_ENUMERATION_CONTEXT = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_ENUMERATION_NUM_OBJECTS_REQUESTED() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB7474E91, &HE7F8, &H4AD9, &HB4, &H0, &HAD, &H1A, &H4B, &H58, &HEE, &HEC, 1005)
 WPD_PROPERTY_OBJECT_ENUMERATION_NUM_OBJECTS_REQUESTED = pkk
End Function
Public Function WPD_COMMAND_OBJECT_PROPERTIES_GET_SUPPORTED() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9E5582E4, &H814, &H44E6, &H98, &H1A, &HB2, &H99, &H8D, &H58, &H38, &H4, 2)
 WPD_COMMAND_OBJECT_PROPERTIES_GET_SUPPORTED = pkk
End Function
Public Function WPD_COMMAND_OBJECT_PROPERTIES_GET_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9E5582E4, &H814, &H44E6, &H98, &H1A, &HB2, &H99, &H8D, &H58, &H38, &H4, 3)
 WPD_COMMAND_OBJECT_PROPERTIES_GET_ATTRIBUTES = pkk
End Function
Public Function WPD_COMMAND_OBJECT_PROPERTIES_GET() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9E5582E4, &H814, &H44E6, &H98, &H1A, &HB2, &H99, &H8D, &H58, &H38, &H4, 4)
 WPD_COMMAND_OBJECT_PROPERTIES_GET = pkk
End Function
Public Function WPD_COMMAND_OBJECT_PROPERTIES_SET() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9E5582E4, &H814, &H44E6, &H98, &H1A, &HB2, &H99, &H8D, &H58, &H38, &H4, 5)
 WPD_COMMAND_OBJECT_PROPERTIES_SET = pkk
End Function
Public Function WPD_COMMAND_OBJECT_PROPERTIES_GET_ALL() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9E5582E4, &H814, &H44E6, &H98, &H1A, &HB2, &H99, &H8D, &H58, &H38, &H4, 6)
 WPD_COMMAND_OBJECT_PROPERTIES_GET_ALL = pkk
End Function
Public Function WPD_COMMAND_OBJECT_PROPERTIES_DELETE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9E5582E4, &H814, &H44E6, &H98, &H1A, &HB2, &H99, &H8D, &H58, &H38, &H4, 7)
 WPD_COMMAND_OBJECT_PROPERTIES_DELETE = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_PROPERTIES_OBJECT_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9E5582E4, &H814, &H44E6, &H98, &H1A, &HB2, &H99, &H8D, &H58, &H38, &H4, 1001)
 WPD_PROPERTY_OBJECT_PROPERTIES_OBJECT_ID = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_PROPERTIES_PROPERTY_KEYS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9E5582E4, &H814, &H44E6, &H98, &H1A, &HB2, &H99, &H8D, &H58, &H38, &H4, 1002)
 WPD_PROPERTY_OBJECT_PROPERTIES_PROPERTY_KEYS = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_PROPERTIES_PROPERTY_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9E5582E4, &H814, &H44E6, &H98, &H1A, &HB2, &H99, &H8D, &H58, &H38, &H4, 1003)
 WPD_PROPERTY_OBJECT_PROPERTIES_PROPERTY_ATTRIBUTES = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_PROPERTIES_PROPERTY_VALUES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9E5582E4, &H814, &H44E6, &H98, &H1A, &HB2, &H99, &H8D, &H58, &H38, &H4, 1004)
 WPD_PROPERTY_OBJECT_PROPERTIES_PROPERTY_VALUES = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_PROPERTIES_PROPERTY_WRITE_RESULTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9E5582E4, &H814, &H44E6, &H98, &H1A, &HB2, &H99, &H8D, &H58, &H38, &H4, 1005)
 WPD_PROPERTY_OBJECT_PROPERTIES_PROPERTY_WRITE_RESULTS = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_PROPERTIES_PROPERTY_DELETE_RESULTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9E5582E4, &H814, &H44E6, &H98, &H1A, &HB2, &H99, &H8D, &H58, &H38, &H4, 1006)
 WPD_PROPERTY_OBJECT_PROPERTIES_PROPERTY_DELETE_RESULTS = pkk
End Function
Public Function WPD_COMMAND_OBJECT_PROPERTIES_BULK_GET_VALUES_BY_OBJECT_LIST_START() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 2)
 WPD_COMMAND_OBJECT_PROPERTIES_BULK_GET_VALUES_BY_OBJECT_LIST_START = pkk
End Function
Public Function WPD_COMMAND_OBJECT_PROPERTIES_BULK_GET_VALUES_BY_OBJECT_LIST_NEXT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 3)
 WPD_COMMAND_OBJECT_PROPERTIES_BULK_GET_VALUES_BY_OBJECT_LIST_NEXT = pkk
End Function
Public Function WPD_COMMAND_OBJECT_PROPERTIES_BULK_GET_VALUES_BY_OBJECT_LIST_END() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 4)
 WPD_COMMAND_OBJECT_PROPERTIES_BULK_GET_VALUES_BY_OBJECT_LIST_END = pkk
End Function
Public Function WPD_COMMAND_OBJECT_PROPERTIES_BULK_GET_VALUES_BY_OBJECT_FORMAT_START() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 5)
 WPD_COMMAND_OBJECT_PROPERTIES_BULK_GET_VALUES_BY_OBJECT_FORMAT_START = pkk
End Function
Public Function WPD_COMMAND_OBJECT_PROPERTIES_BULK_GET_VALUES_BY_OBJECT_FORMAT_NEXT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 6)
 WPD_COMMAND_OBJECT_PROPERTIES_BULK_GET_VALUES_BY_OBJECT_FORMAT_NEXT = pkk
End Function
Public Function WPD_COMMAND_OBJECT_PROPERTIES_BULK_GET_VALUES_BY_OBJECT_FORMAT_END() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 7)
 WPD_COMMAND_OBJECT_PROPERTIES_BULK_GET_VALUES_BY_OBJECT_FORMAT_END = pkk
End Function
Public Function WPD_COMMAND_OBJECT_PROPERTIES_BULK_SET_VALUES_BY_OBJECT_LIST_START() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 8)
 WPD_COMMAND_OBJECT_PROPERTIES_BULK_SET_VALUES_BY_OBJECT_LIST_START = pkk
End Function
Public Function WPD_COMMAND_OBJECT_PROPERTIES_BULK_SET_VALUES_BY_OBJECT_LIST_NEXT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 9)
 WPD_COMMAND_OBJECT_PROPERTIES_BULK_SET_VALUES_BY_OBJECT_LIST_NEXT = pkk
End Function
Public Function WPD_COMMAND_OBJECT_PROPERTIES_BULK_SET_VALUES_BY_OBJECT_LIST_END() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 10)
 WPD_COMMAND_OBJECT_PROPERTIES_BULK_SET_VALUES_BY_OBJECT_LIST_END = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_PROPERTIES_BULK_OBJECT_IDS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 1001)
 WPD_PROPERTY_OBJECT_PROPERTIES_BULK_OBJECT_IDS = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_PROPERTIES_BULK_CONTEXT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 1002)
 WPD_PROPERTY_OBJECT_PROPERTIES_BULK_CONTEXT = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_PROPERTIES_BULK_VALUES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 1003)
 WPD_PROPERTY_OBJECT_PROPERTIES_BULK_VALUES = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_PROPERTIES_BULK_PROPERTY_KEYS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 1004)
 WPD_PROPERTY_OBJECT_PROPERTIES_BULK_PROPERTY_KEYS = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_PROPERTIES_BULK_DEPTH() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 1005)
 WPD_PROPERTY_OBJECT_PROPERTIES_BULK_DEPTH = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_PROPERTIES_BULK_PARENT_OBJECT_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 1006)
 WPD_PROPERTY_OBJECT_PROPERTIES_BULK_PARENT_OBJECT_ID = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_PROPERTIES_BULK_OBJECT_FORMAT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 1007)
 WPD_PROPERTY_OBJECT_PROPERTIES_BULK_OBJECT_FORMAT = pkk
End Function
Public Function WPD_PROPERTY_OBJECT_PROPERTIES_BULK_WRITE_RESULTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11C824DD, &H4CD, &H4E4E, &H8C, &H7B, &HF6, &HEF, &HB7, &H94, &HD8, &H4E, 1008)
 WPD_PROPERTY_OBJECT_PROPERTIES_BULK_WRITE_RESULTS = pkk
End Function
Public Function WPD_COMMAND_OBJECT_RESOURCES_GET_SUPPORTED() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 2)
 WPD_COMMAND_OBJECT_RESOURCES_GET_SUPPORTED = pkk
End Function
Public Function WPD_COMMAND_OBJECT_RESOURCES_GET_ATTRIBUTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 3)
 WPD_COMMAND_OBJECT_RESOURCES_GET_ATTRIBUTES = pkk
End Function
Public Function WPD_COMMAND_OBJECT_RESOURCES_OPEN() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 4)
 WPD_COMMAND_OBJECT_RESOURCES_OPEN = pkk
End Function
Public Function WPD_COMMAND_OBJECT_RESOURCES_READ() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 5)
 WPD_COMMAND_OBJECT_RESOURCES_READ = pkk
End Function
Public Function WPD_COMMAND_OBJECT_RESOURCES_WRITE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 6)
 WPD_COMMAND_OBJECT_RESOURCES_WRITE = pkk
End Function
Public Function WPD_COMMAND_OBJECT_RESOURCES_CLOSE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 7)
 WPD_COMMAND_OBJECT_RESOURCES_CLOSE = pkk
End Function
Public Function WPD_COMMAND_OBJECT_RESOURCES_DELETE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 8)
 WPD_COMMAND_OBJECT_RESOURCES_DELETE = pkk
End Function
Public Function WPD_COMMAND_OBJECT_RESOURCES_CREATE_RESOURCE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3A2B22D, &HA595, &H4108, &HBE, &HA, &HFC, &H3C, &H96, &H5F, &H3D, &H4A, 9)
 WPD_COMMAND_OBJECT_RESOURCES_CREATE_RESOURCE = pkk
End Function
Public Function WPD_CLASS_EXTENSION_OPTIONS_SILENCE_AUTOPLAY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H65C160F8, &H1367, &H4CE2, &H93, &H9D, &H83, &H10, &H83, &H9F, &HD, &H30, 2)
 WPD_CLASS_EXTENSION_OPTIONS_SILENCE_AUTOPLAY = pkk
End Function
Public Function WPD_RESOURCE_ATTRIBUTE_TOTAL_SIZE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1EB6F604, &H9278, &H429F, &H93, &HCC, &H5B, &HB8, &HC0, &H66, &H56, &HB6, 2)
 WPD_RESOURCE_ATTRIBUTE_TOTAL_SIZE = pkk
End Function
Public Function WPD_RESOURCE_ATTRIBUTE_CAN_READ() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1EB6F604, &H9278, &H429F, &H93, &HCC, &H5B, &HB8, &HC0, &H66, &H56, &HB6, 3)
 WPD_RESOURCE_ATTRIBUTE_CAN_READ = pkk
End Function
Public Function WPD_RESOURCE_ATTRIBUTE_CAN_WRITE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1EB6F604, &H9278, &H429F, &H93, &HCC, &H5B, &HB8, &HC0, &H66, &H56, &HB6, 4)
 WPD_RESOURCE_ATTRIBUTE_CAN_WRITE = pkk
End Function
Public Function WPD_RESOURCE_ATTRIBUTE_CAN_DELETE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1EB6F604, &H9278, &H429F, &H93, &HCC, &H5B, &HB8, &HC0, &H66, &H56, &HB6, 5)
 WPD_RESOURCE_ATTRIBUTE_CAN_DELETE = pkk
End Function
Public Function WPD_RESOURCE_ATTRIBUTE_OPTIMAL_READ_BUFFER_SIZE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1EB6F604, &H9278, &H429F, &H93, &HCC, &H5B, &HB8, &HC0, &H66, &H56, &HB6, 6)
 WPD_RESOURCE_ATTRIBUTE_OPTIMAL_READ_BUFFER_SIZE = pkk
End Function
Public Function WPD_RESOURCE_ATTRIBUTE_OPTIMAL_WRITE_BUFFER_SIZE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1EB6F604, &H9278, &H429F, &H93, &HCC, &H5B, &HB8, &HC0, &H66, &H56, &HB6, 7)
 WPD_RESOURCE_ATTRIBUTE_OPTIMAL_WRITE_BUFFER_SIZE = pkk
End Function
Public Function WPD_RESOURCE_ATTRIBUTE_FORMAT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1EB6F604, &H9278, &H429F, &H93, &HCC, &H5B, &HB8, &HC0, &H66, &H56, &HB6, 8)
 WPD_RESOURCE_ATTRIBUTE_FORMAT = pkk
End Function
Public Function WPD_RESOURCE_ATTRIBUTE_RESOURCE_KEY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1EB6F604, &H9278, &H429F, &H93, &HCC, &H5B, &HB8, &HC0, &H66, &H56, &HB6, 9)
 WPD_RESOURCE_ATTRIBUTE_RESOURCE_KEY = pkk
End Function
Public Function WPD_DEVICE_SYNC_PARTNER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26D4979A, &HE643, &H4626, &H9E, &H2B, &H73, &H6D, &HC0, &HC9, &H2F, &HDC, 2)
 WPD_DEVICE_SYNC_PARTNER = pkk
End Function
Public Function WPD_DEVICE_FIRMWARE_VERSION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26D4979A, &HE643, &H4626, &H9E, &H2B, &H73, &H6D, &HC0, &HC9, &H2F, &HDC, 3)
 WPD_DEVICE_FIRMWARE_VERSION = pkk
End Function
Public Function WPD_DEVICE_POWER_LEVEL() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26D4979A, &HE643, &H4626, &H9E, &H2B, &H73, &H6D, &HC0, &HC9, &H2F, &HDC, 4)
 WPD_DEVICE_POWER_LEVEL = pkk
End Function
Public Function WPD_DEVICE_POWER_SOURCE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26D4979A, &HE643, &H4626, &H9E, &H2B, &H73, &H6D, &HC0, &HC9, &H2F, &HDC, 5)
 WPD_DEVICE_POWER_SOURCE = pkk
End Function
Public Function WPD_DEVICE_PROTOCOL() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26D4979A, &HE643, &H4626, &H9E, &H2B, &H73, &H6D, &HC0, &HC9, &H2F, &HDC, 6)
 WPD_DEVICE_PROTOCOL = pkk
End Function
Public Function WPD_DEVICE_MANUFACTURER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26D4979A, &HE643, &H4626, &H9E, &H2B, &H73, &H6D, &HC0, &HC9, &H2F, &HDC, 7)
 WPD_DEVICE_MANUFACTURER = pkk
End Function
Public Function WPD_DEVICE_MODEL() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26D4979A, &HE643, &H4626, &H9E, &H2B, &H73, &H6D, &HC0, &HC9, &H2F, &HDC, 8)
 WPD_DEVICE_MODEL = pkk
End Function
Public Function WPD_DEVICE_SERIAL_NUMBER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26D4979A, &HE643, &H4626, &H9E, &H2B, &H73, &H6D, &HC0, &HC9, &H2F, &HDC, 9)
 WPD_DEVICE_SERIAL_NUMBER = pkk
End Function
Public Function WPD_DEVICE_SUPPORTS_NON_CONSUMABLE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26D4979A, &HE643, &H4626, &H9E, &H2B, &H73, &H6D, &HC0, &HC9, &H2F, &HDC, 10)
 WPD_DEVICE_SUPPORTS_NON_CONSUMABLE = pkk
End Function
Public Function WPD_DEVICE_DATETIME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26D4979A, &HE643, &H4626, &H9E, &H2B, &H73, &H6D, &HC0, &HC9, &H2F, &HDC, 11)
 WPD_DEVICE_DATETIME = pkk
End Function
Public Function WPD_DEVICE_FRIENDLY_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26D4979A, &HE643, &H4626, &H9E, &H2B, &H73, &H6D, &HC0, &HC9, &H2F, &HDC, 12)
 WPD_DEVICE_FRIENDLY_NAME = pkk
End Function
Public Function WPD_DEVICE_SUPPORTED_DRM_SCHEMES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26D4979A, &HE643, &H4626, &H9E, &H2B, &H73, &H6D, &HC0, &HC9, &H2F, &HDC, 13)
 WPD_DEVICE_SUPPORTED_DRM_SCHEMES = pkk
End Function
Public Function WPD_DEVICE_SUPPORTED_FORMATS_ARE_ORDERED() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26D4979A, &HE643, &H4626, &H9E, &H2B, &H73, &H6D, &HC0, &HC9, &H2F, &HDC, 14)
 WPD_DEVICE_SUPPORTED_FORMATS_ARE_ORDERED = pkk
End Function
Public Function WPD_DEVICE_TYPE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26D4979A, &HE643, &H4626, &H9E, &H2B, &H73, &H6D, &HC0, &HC9, &H2F, &HDC, 15)
 WPD_DEVICE_TYPE = pkk
End Function
Public Function WPD_DEVICE_NETWORK_IDENTIFIER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26D4979A, &HE643, &H4626, &H9E, &H2B, &H73, &H6D, &HC0, &HC9, &H2F, &HDC, 16)
 WPD_DEVICE_NETWORK_IDENTIFIER = pkk
End Function
Public Function WPD_DEVICE_FUNCTIONAL_UNIQUE_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H463DD662, &H7FC4, &H4291, &H91, &H1C, &H7F, &H4C, &H9C, &HCA, &H97, &H99, 2)
 WPD_DEVICE_FUNCTIONAL_UNIQUE_ID = pkk
End Function
Public Function WPD_DEVICE_MODEL_UNIQUE_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H463DD662, &H7FC4, &H4291, &H91, &H1C, &H7F, &H4C, &H9C, &HCA, &H97, &H99, 3)
 WPD_DEVICE_MODEL_UNIQUE_ID = pkk
End Function
Public Function WPD_DEVICE_TRANSPORT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H463DD662, &H7FC4, &H4291, &H91, &H1C, &H7F, &H4C, &H9C, &HCA, &H97, &H99, 4)
 WPD_DEVICE_TRANSPORT = pkk
End Function
Public Function WPD_DEVICE_USE_DEVICE_STAGE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H463DD662, &H7FC4, &H4291, &H91, &H1C, &H7F, &H4C, &H9C, &HCA, &H97, &H99, 5)
 WPD_DEVICE_USE_DEVICE_STAGE = pkk
End Function
Public Function WPD_DEVICE_EDP_IDENTITY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6C2B878C, &HC2EC, &H490D, &HB4, &H25, &HD7, &HA7, &H5E, &H23, &HE5, &HED, 1)
 WPD_DEVICE_EDP_IDENTITY = pkk
End Function
Public Function WPD_SERVICE_VERSION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7510698A, &HCB54, &H481C, &HB8, &HDB, &HD, &H75, &HC9, &H3F, &H1C, &H6, 2)
 WPD_SERVICE_VERSION = pkk
End Function
Public Function WPD_EVENT_PARAMETER_PNP_DEVICE_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H15AB1953, &HF817, &H4FEF, &HA9, &H21, &H56, &H76, &HE8, &H38, &HF6, &HE0, 2)
 WPD_EVENT_PARAMETER_PNP_DEVICE_ID = pkk
End Function
Public Function WPD_EVENT_PARAMETER_EVENT_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H15AB1953, &HF817, &H4FEF, &HA9, &H21, &H56, &H76, &HE8, &H38, &HF6, &HE0, 3)
 WPD_EVENT_PARAMETER_EVENT_ID = pkk
End Function
Public Function WPD_EVENT_PARAMETER_OPERATION_STATE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H15AB1953, &HF817, &H4FEF, &HA9, &H21, &H56, &H76, &HE8, &H38, &HF6, &HE0, 4)
 WPD_EVENT_PARAMETER_OPERATION_STATE = pkk
End Function
Public Function WPD_EVENT_PARAMETER_OPERATION_PROGRESS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H15AB1953, &HF817, &H4FEF, &HA9, &H21, &H56, &H76, &HE8, &H38, &HF6, &HE0, 5)
 WPD_EVENT_PARAMETER_OPERATION_PROGRESS = pkk
End Function
Public Function WPD_EVENT_PARAMETER_OBJECT_PARENT_PERSISTENT_UNIQUE_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H15AB1953, &HF817, &H4FEF, &HA9, &H21, &H56, &H76, &HE8, &H38, &HF6, &HE0, 6)
 WPD_EVENT_PARAMETER_OBJECT_PARENT_PERSISTENT_UNIQUE_ID = pkk
End Function
Public Function WPD_EVENT_PARAMETER_OBJECT_CREATION_COOKIE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H15AB1953, &HF817, &H4FEF, &HA9, &H21, &H56, &H76, &HE8, &H38, &HF6, &HE0, 7)
 WPD_EVENT_PARAMETER_OBJECT_CREATION_COOKIE = pkk
End Function
Public Function WPD_EVENT_PARAMETER_CHILD_HIERARCHY_CHANGED() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H15AB1953, &HF817, &H4FEF, &HA9, &H21, &H56, &H76, &HE8, &H38, &HF6, &HE0, 8)
 WPD_EVENT_PARAMETER_CHILD_HIERARCHY_CHANGED = pkk
End Function
Public Function WPD_EVENT_PARAMETER_SERVICE_METHOD_CONTEXT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H52807B8A, &H4914, &H4323, &H9B, &H9A, &H74, &HF6, &H54, &HB2, &HB8, &H46, 2)
 WPD_EVENT_PARAMETER_SERVICE_METHOD_CONTEXT = pkk
End Function
Public Function WPD_EVENT_OPTION_IS_BROADCAST_EVENT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3D8DAD7, &HA361, &H4B83, &H8A, &H48, &H5B, &H2, &HCE, &H10, &H71, &H3B, 2)
 WPD_EVENT_OPTION_IS_BROADCAST_EVENT = pkk
End Function
Public Function WPD_EVENT_OPTION_IS_AUTOPLAY_EVENT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB3D8DAD7, &HA361, &H4B83, &H8A, &H48, &H5B, &H2, &HCE, &H10, &H71, &H3B, 3)
 WPD_EVENT_OPTION_IS_AUTOPLAY_EVENT = pkk
End Function
Public Function WPD_EVENT_ATTRIBUTE_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H10C96578, &H2E81, &H4111, &HAD, &HDE, &HE0, &H8C, &HA6, &H13, &H8F, &H6D, 2)
 WPD_EVENT_ATTRIBUTE_NAME = pkk
End Function
Public Function WPD_EVENT_ATTRIBUTE_PARAMETERS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H10C96578, &H2E81, &H4111, &HAD, &HDE, &HE0, &H8C, &HA6, &H13, &H8F, &H6D, 3)
 WPD_EVENT_ATTRIBUTE_PARAMETERS = pkk
End Function
Public Function WPD_EVENT_ATTRIBUTE_OPTIONS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H10C96578, &H2E81, &H4111, &HAD, &HDE, &HE0, &H8C, &HA6, &H13, &H8F, &H6D, 4)
 WPD_EVENT_ATTRIBUTE_OPTIONS = pkk
End Function
Public Function WPD_API_OPTION_USE_CLEAR_DATA_STREAM() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H10E54A3E, &H52D, &H4777, &HA1, &H3C, &HDE, &H76, &H14, &HBE, &H2B, &HC4, 2)
 WPD_API_OPTION_USE_CLEAR_DATA_STREAM = pkk
End Function
Public Function WPD_API_OPTION_IOCTL_ACCESS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H10E54A3E, &H52D, &H4777, &HA1, &H3C, &HDE, &H76, &H14, &HBE, &H2B, &HC4, 3)
 WPD_API_OPTION_IOCTL_ACCESS = pkk
End Function
Public Function WPD_FORMAT_ATTRIBUTE_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA0A02000, &HBCAF, &H4BE8, &HB3, &HF5, &H23, &H3F, &H23, &H1C, &HF5, &H8F, 2)
 WPD_FORMAT_ATTRIBUTE_NAME = pkk
End Function
Public Function WPD_FORMAT_ATTRIBUTE_MIMETYPE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA0A02000, &HBCAF, &H4BE8, &HB3, &HF5, &H23, &H3F, &H23, &H1C, &HF5, &H8F, 3)
 WPD_FORMAT_ATTRIBUTE_MIMETYPE = pkk
End Function
Public Function WPD_METHOD_ATTRIBUTE_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF17A5071, &HF039, &H44AF, &H8E, &HFE, &H43, &H2C, &HF3, &H2E, &H43, &H2A, 2)
 WPD_METHOD_ATTRIBUTE_NAME = pkk
End Function
Public Function WPD_METHOD_ATTRIBUTE_ASSOCIATED_FORMAT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF17A5071, &HF039, &H44AF, &H8E, &HFE, &H43, &H2C, &HF3, &H2E, &H43, &H2A, 3)
 WPD_METHOD_ATTRIBUTE_ASSOCIATED_FORMAT = pkk
End Function
Public Function WPD_METHOD_ATTRIBUTE_ACCESS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF17A5071, &HF039, &H44AF, &H8E, &HFE, &H43, &H2C, &HF3, &H2E, &H43, &H2A, 4)
 WPD_METHOD_ATTRIBUTE_ACCESS = pkk
End Function
Public Function WPD_METHOD_ATTRIBUTE_PARAMETERS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF17A5071, &HF039, &H44AF, &H8E, &HFE, &H43, &H2C, &HF3, &H2E, &H43, &H2A, 5)
 WPD_METHOD_ATTRIBUTE_PARAMETERS = pkk
End Function
Public Function WPD_PARAMETER_ATTRIBUTE_ORDER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE6864DD7, &HF325, &H45EA, &HA1, &HD5, &H97, &HCF, &H73, &HB6, &HCA, &H58, 2)
 WPD_PARAMETER_ATTRIBUTE_ORDER = pkk
End Function
Public Function WPD_PARAMETER_ATTRIBUTE_USAGE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE6864DD7, &HF325, &H45EA, &HA1, &HD5, &H97, &HCF, &H73, &HB6, &HCA, &H58, 3)
 WPD_PARAMETER_ATTRIBUTE_USAGE = pkk
End Function
Public Function WPD_PARAMETER_ATTRIBUTE_FORM() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE6864DD7, &HF325, &H45EA, &HA1, &HD5, &H97, &HCF, &H73, &HB6, &HCA, &H58, 4)
 WPD_PARAMETER_ATTRIBUTE_FORM = pkk
End Function
Public Function WPD_PARAMETER_ATTRIBUTE_DEFAULT_VALUE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE6864DD7, &HF325, &H45EA, &HA1, &HD5, &H97, &HCF, &H73, &HB6, &HCA, &H58, 5)
 WPD_PARAMETER_ATTRIBUTE_DEFAULT_VALUE = pkk
End Function
Public Function WPD_PARAMETER_ATTRIBUTE_RANGE_MIN() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE6864DD7, &HF325, &H45EA, &HA1, &HD5, &H97, &HCF, &H73, &HB6, &HCA, &H58, 6)
 WPD_PARAMETER_ATTRIBUTE_RANGE_MIN = pkk
End Function
Public Function WPD_PARAMETER_ATTRIBUTE_RANGE_MAX() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE6864DD7, &HF325, &H45EA, &HA1, &HD5, &H97, &HCF, &H73, &HB6, &HCA, &H58, 7)
 WPD_PARAMETER_ATTRIBUTE_RANGE_MAX = pkk
End Function
Public Function WPD_PARAMETER_ATTRIBUTE_RANGE_STEP() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE6864DD7, &HF325, &H45EA, &HA1, &HD5, &H97, &HCF, &H73, &HB6, &HCA, &H58, 8)
 WPD_PARAMETER_ATTRIBUTE_RANGE_STEP = pkk
End Function
Public Function WPD_PARAMETER_ATTRIBUTE_ENUMERATION_ELEMENTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE6864DD7, &HF325, &H45EA, &HA1, &HD5, &H97, &HCF, &H73, &HB6, &HCA, &H58, 9)
 WPD_PARAMETER_ATTRIBUTE_ENUMERATION_ELEMENTS = pkk
End Function
Public Function WPD_PARAMETER_ATTRIBUTE_REGULAR_EXPRESSION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE6864DD7, &HF325, &H45EA, &HA1, &HD5, &H97, &HCF, &H73, &HB6, &HCA, &H58, 10)
 WPD_PARAMETER_ATTRIBUTE_REGULAR_EXPRESSION = pkk
End Function
Public Function WPD_PARAMETER_ATTRIBUTE_MAX_SIZE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE6864DD7, &HF325, &H45EA, &HA1, &HD5, &H97, &HCF, &H73, &HB6, &HCA, &H58, 11)
 WPD_PARAMETER_ATTRIBUTE_MAX_SIZE = pkk
End Function
Public Function WPD_PARAMETER_ATTRIBUTE_VARTYPE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE6864DD7, &HF325, &H45EA, &HA1, &HD5, &H97, &HCF, &H73, &HB6, &HCA, &H58, 12)
 WPD_PARAMETER_ATTRIBUTE_VARTYPE = pkk
End Function
Public Function WPD_PROPERTY_NULL() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, 0)
 WPD_PROPERTY_NULL = pkk
End Function
Public Function WPD_OBJECT_CONTENT_TYPE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 7)
 WPD_OBJECT_CONTENT_TYPE = pkk
End Function
Public Function WPD_OBJECT_REFERENCES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 14)
 WPD_OBJECT_REFERENCES = pkk
End Function
Public Function WPD_OBJECT_CONTAINER_FUNCTIONAL_OBJECT_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 23)
 WPD_OBJECT_CONTAINER_FUNCTIONAL_OBJECT_ID = pkk
End Function
Public Function WPD_OBJECT_GENERATE_THUMBNAIL_FROM_RESOURCE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 24)
 WPD_OBJECT_GENERATE_THUMBNAIL_FROM_RESOURCE = pkk
End Function
Public Function WPD_OBJECT_HINT_LOCATION_DISPLAY_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF6B490D, &H5CD8, &H437A, &HAF, &HFC, &HDA, &H8B, &H60, &HEE, &H4A, &H3C, 25)
 WPD_OBJECT_HINT_LOCATION_DISPLAY_NAME = pkk
End Function
Public Function WPD_OBJECT_SUPPORTED_UNITS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H373CD3D, &H4A46, &H40D7, &HB4, &HD8, &H73, &HE8, &HDA, &H74, &HE7, &H75, 2)
 WPD_OBJECT_SUPPORTED_UNITS = pkk
End Function
Public Function WPD_FUNCTIONAL_OBJECT_CATEGORY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8F052D93, &HABCA, &H4FC5, &HA5, &HAC, &HB0, &H1D, &HF4, &HDB, &HE5, &H98, 2)
 WPD_FUNCTIONAL_OBJECT_CATEGORY = pkk
End Function
Public Function WPD_STORAGE_TYPE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1A3057A, &H74D6, &H4E80, &HBE, &HA7, &HDC, &H4C, &H21, &H2C, &HE5, &HA, 2)
 WPD_STORAGE_TYPE = pkk
End Function
Public Function WPD_STORAGE_FILE_SYSTEM_TYPE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1A3057A, &H74D6, &H4E80, &HBE, &HA7, &HDC, &H4C, &H21, &H2C, &HE5, &HA, 3)
 WPD_STORAGE_FILE_SYSTEM_TYPE = pkk
End Function
Public Function WPD_STORAGE_CAPACITY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1A3057A, &H74D6, &H4E80, &HBE, &HA7, &HDC, &H4C, &H21, &H2C, &HE5, &HA, 4)
 WPD_STORAGE_CAPACITY = pkk
End Function
Public Function WPD_STORAGE_FREE_SPACE_IN_BYTES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1A3057A, &H74D6, &H4E80, &HBE, &HA7, &HDC, &H4C, &H21, &H2C, &HE5, &HA, 5)
 WPD_STORAGE_FREE_SPACE_IN_BYTES = pkk
End Function
Public Function WPD_STORAGE_FREE_SPACE_IN_OBJECTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1A3057A, &H74D6, &H4E80, &HBE, &HA7, &HDC, &H4C, &H21, &H2C, &HE5, &HA, 6)
 WPD_STORAGE_FREE_SPACE_IN_OBJECTS = pkk
End Function
Public Function WPD_STORAGE_DESCRIPTION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1A3057A, &H74D6, &H4E80, &HBE, &HA7, &HDC, &H4C, &H21, &H2C, &HE5, &HA, 7)
 WPD_STORAGE_DESCRIPTION = pkk
End Function
Public Function WPD_STORAGE_SERIAL_NUMBER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1A3057A, &H74D6, &H4E80, &HBE, &HA7, &HDC, &H4C, &H21, &H2C, &HE5, &HA, 8)
 WPD_STORAGE_SERIAL_NUMBER = pkk
End Function
Public Function WPD_STORAGE_MAX_OBJECT_SIZE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1A3057A, &H74D6, &H4E80, &HBE, &HA7, &HDC, &H4C, &H21, &H2C, &HE5, &HA, 9)
 WPD_STORAGE_MAX_OBJECT_SIZE = pkk
End Function
Public Function WPD_STORAGE_CAPACITY_IN_OBJECTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1A3057A, &H74D6, &H4E80, &HBE, &HA7, &HDC, &H4C, &H21, &H2C, &HE5, &HA, 10)
 WPD_STORAGE_CAPACITY_IN_OBJECTS = pkk
End Function
Public Function WPD_STORAGE_ACCESS_CAPABILITY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1A3057A, &H74D6, &H4E80, &HBE, &HA7, &HDC, &H4C, &H21, &H2C, &HE5, &HA, 11)
 WPD_STORAGE_ACCESS_CAPABILITY = pkk
End Function
Public Function WPD_NETWORK_ASSOCIATION_HOST_NETWORK_IDENTIFIERS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE4C93C1F, &HB203, &H43F1, &HA1, &H0, &H5A, &H7, &HD1, &H1B, &H2, &H74, 2)
 WPD_NETWORK_ASSOCIATION_HOST_NETWORK_IDENTIFIERS = pkk
End Function
Public Function WPD_NETWORK_ASSOCIATION_X509V3SEQUENCE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE4C93C1F, &HB203, &H43F1, &HA1, &H0, &H5A, &H7, &HD1, &H1B, &H2, &H74, 3)
 WPD_NETWORK_ASSOCIATION_X509V3SEQUENCE = pkk
End Function
Public Function WPD_STILL_IMAGE_CAPTURE_RESOLUTION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 2)
 WPD_STILL_IMAGE_CAPTURE_RESOLUTION = pkk
End Function
Public Function WPD_STILL_IMAGE_CAPTURE_FORMAT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 3)
 WPD_STILL_IMAGE_CAPTURE_FORMAT = pkk
End Function
Public Function WPD_STILL_IMAGE_COMPRESSION_SETTING() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 4)
 WPD_STILL_IMAGE_COMPRESSION_SETTING = pkk
End Function
Public Function WPD_STILL_IMAGE_WHITE_BALANCE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 5)
 WPD_STILL_IMAGE_WHITE_BALANCE = pkk
End Function
Public Function WPD_STILL_IMAGE_RGB_GAIN() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 6)
 WPD_STILL_IMAGE_RGB_GAIN = pkk
End Function
Public Function WPD_STILL_IMAGE_FNUMBER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 7)
 WPD_STILL_IMAGE_FNUMBER = pkk
End Function
Public Function WPD_STILL_IMAGE_FOCAL_LENGTH() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 8)
 WPD_STILL_IMAGE_FOCAL_LENGTH = pkk
End Function
Public Function WPD_STILL_IMAGE_FOCUS_DISTANCE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 9)
 WPD_STILL_IMAGE_FOCUS_DISTANCE = pkk
End Function
Public Function WPD_STILL_IMAGE_FOCUS_MODE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 10)
 WPD_STILL_IMAGE_FOCUS_MODE = pkk
End Function
Public Function WPD_STILL_IMAGE_EXPOSURE_METERING_MODE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 11)
 WPD_STILL_IMAGE_EXPOSURE_METERING_MODE = pkk
End Function
Public Function WPD_STILL_IMAGE_FLASH_MODE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 12)
 WPD_STILL_IMAGE_FLASH_MODE = pkk
End Function
Public Function WPD_STILL_IMAGE_EXPOSURE_TIME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 13)
 WPD_STILL_IMAGE_EXPOSURE_TIME = pkk
End Function
Public Function WPD_STILL_IMAGE_EXPOSURE_PROGRAM_MODE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 14)
 WPD_STILL_IMAGE_EXPOSURE_PROGRAM_MODE = pkk
End Function
Public Function WPD_STILL_IMAGE_EXPOSURE_INDEX() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 15)
 WPD_STILL_IMAGE_EXPOSURE_INDEX = pkk
End Function
Public Function WPD_STILL_IMAGE_EXPOSURE_BIAS_COMPENSATION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 16)
 WPD_STILL_IMAGE_EXPOSURE_BIAS_COMPENSATION = pkk
End Function
Public Function WPD_STILL_IMAGE_CAPTURE_DELAY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 17)
 WPD_STILL_IMAGE_CAPTURE_DELAY = pkk
End Function
Public Function WPD_STILL_IMAGE_CAPTURE_MODE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 18)
 WPD_STILL_IMAGE_CAPTURE_MODE = pkk
End Function
Public Function WPD_STILL_IMAGE_CONTRAST() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 19)
 WPD_STILL_IMAGE_CONTRAST = pkk
End Function
Public Function WPD_STILL_IMAGE_SHARPNESS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 20)
 WPD_STILL_IMAGE_SHARPNESS = pkk
End Function
Public Function WPD_STILL_IMAGE_DIGITAL_ZOOM() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 21)
 WPD_STILL_IMAGE_DIGITAL_ZOOM = pkk
End Function
Public Function WPD_STILL_IMAGE_EFFECT_MODE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 22)
 WPD_STILL_IMAGE_EFFECT_MODE = pkk
End Function
Public Function WPD_STILL_IMAGE_BURST_NUMBER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 23)
 WPD_STILL_IMAGE_BURST_NUMBER = pkk
End Function
Public Function WPD_STILL_IMAGE_BURST_INTERVAL() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 24)
 WPD_STILL_IMAGE_BURST_INTERVAL = pkk
End Function
Public Function WPD_STILL_IMAGE_TIMELAPSE_NUMBER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 25)
 WPD_STILL_IMAGE_TIMELAPSE_NUMBER = pkk
End Function
Public Function WPD_STILL_IMAGE_TIMELAPSE_INTERVAL() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 26)
 WPD_STILL_IMAGE_TIMELAPSE_INTERVAL = pkk
End Function
Public Function WPD_STILL_IMAGE_FOCUS_METERING_MODE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 27)
 WPD_STILL_IMAGE_FOCUS_METERING_MODE = pkk
End Function
Public Function WPD_STILL_IMAGE_UPLOAD_URL() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 28)
 WPD_STILL_IMAGE_UPLOAD_URL = pkk
End Function
Public Function WPD_STILL_IMAGE_ARTIST() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 29)
 WPD_STILL_IMAGE_ARTIST = pkk
End Function
Public Function WPD_STILL_IMAGE_CAMERA_MODEL() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 30)
 WPD_STILL_IMAGE_CAMERA_MODEL = pkk
End Function
Public Function WPD_STILL_IMAGE_CAMERA_MANUFACTURER() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H58C571EC, &H1BCB, &H42A7, &H8A, &HC5, &HBB, &H29, &H15, &H73, &HA2, &H60, 31)
 WPD_STILL_IMAGE_CAMERA_MANUFACTURER = pkk
End Function
Public Function WPD_RENDERING_INFORMATION_PROFILES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC53D039F, &HEE23, &H4A31, &H85, &H90, &H76, &H39, &H87, &H98, &H70, &HB4, 2)
 WPD_RENDERING_INFORMATION_PROFILES = pkk
End Function
Public Function WPD_RENDERING_INFORMATION_PROFILE_ENTRY_TYPE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC53D039F, &HEE23, &H4A31, &H85, &H90, &H76, &H39, &H87, &H98, &H70, &HB4, 3)
 WPD_RENDERING_INFORMATION_PROFILE_ENTRY_TYPE = pkk
End Function
Public Function WPD_RENDERING_INFORMATION_PROFILE_ENTRY_CREATABLE_RESOURCES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC53D039F, &HEE23, &H4A31, &H85, &H90, &H76, &H39, &H87, &H98, &H70, &HB4, 4)
 WPD_RENDERING_INFORMATION_PROFILE_ENTRY_CREATABLE_RESOURCES = pkk
End Function
Public Function WPD_CLIENT_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H204D9F0C, &H2292, &H4080, &H9F, &H42, &H40, &H66, &H4E, &H70, &HF8, &H59, 2)
 WPD_CLIENT_NAME = pkk
End Function
Public Function WPD_CLIENT_MAJOR_VERSION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H204D9F0C, &H2292, &H4080, &H9F, &H42, &H40, &H66, &H4E, &H70, &HF8, &H59, 3)
 WPD_CLIENT_MAJOR_VERSION = pkk
End Function
Public Function WPD_CLIENT_MINOR_VERSION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H204D9F0C, &H2292, &H4080, &H9F, &H42, &H40, &H66, &H4E, &H70, &HF8, &H59, 4)
 WPD_CLIENT_MINOR_VERSION = pkk
End Function
Public Function WPD_CLIENT_REVISION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H204D9F0C, &H2292, &H4080, &H9F, &H42, &H40, &H66, &H4E, &H70, &HF8, &H59, 5)
 WPD_CLIENT_REVISION = pkk
End Function
Public Function WPD_CLIENT_WMDRM_APPLICATION_PRIVATE_KEY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H204D9F0C, &H2292, &H4080, &H9F, &H42, &H40, &H66, &H4E, &H70, &HF8, &H59, 6)
 WPD_CLIENT_WMDRM_APPLICATION_PRIVATE_KEY = pkk
End Function
Public Function WPD_CLIENT_WMDRM_APPLICATION_CERTIFICATE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H204D9F0C, &H2292, &H4080, &H9F, &H42, &H40, &H66, &H4E, &H70, &HF8, &H59, 7)
 WPD_CLIENT_WMDRM_APPLICATION_CERTIFICATE = pkk
End Function
Public Function WPD_CLIENT_SECURITY_QUALITY_OF_SERVICE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H204D9F0C, &H2292, &H4080, &H9F, &H42, &H40, &H66, &H4E, &H70, &HF8, &H59, 8)
 WPD_CLIENT_SECURITY_QUALITY_OF_SERVICE = pkk
End Function
Public Function WPD_CLIENT_DESIRED_ACCESS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H204D9F0C, &H2292, &H4080, &H9F, &H42, &H40, &H66, &H4E, &H70, &HF8, &H59, 9)
 WPD_CLIENT_DESIRED_ACCESS = pkk
End Function
Public Function WPD_CLIENT_SHARE_MODE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H204D9F0C, &H2292, &H4080, &H9F, &H42, &H40, &H66, &H4E, &H70, &HF8, &H59, 10)
 WPD_CLIENT_SHARE_MODE = pkk
End Function
Public Function WPD_CLIENT_EVENT_COOKIE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H204D9F0C, &H2292, &H4080, &H9F, &H42, &H40, &H66, &H4E, &H70, &HF8, &H59, 11)
 WPD_CLIENT_EVENT_COOKIE = pkk
End Function
Public Function WPD_CLIENT_MINIMUM_RESULTS_BUFFER_SIZE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H204D9F0C, &H2292, &H4080, &H9F, &H42, &H40, &H66, &H4E, &H70, &HF8, &H59, 12)
 WPD_CLIENT_MINIMUM_RESULTS_BUFFER_SIZE = pkk
End Function
Public Function WPD_CLIENT_MANUAL_CLOSE_ON_DISCONNECT() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H204D9F0C, &H2292, &H4080, &H9F, &H42, &H40, &H66, &H4E, &H70, &HF8, &H59, 13)
 WPD_CLIENT_MANUAL_CLOSE_ON_DISCONNECT = pkk
End Function
Public Function WPD_PROPERTY_ATTRIBUTE_FORM() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAB7943D8, &H6332, &H445F, &HA0, &HD, &H8D, &H5E, &HF1, &HE9, &H6F, &H37, 2)
 WPD_PROPERTY_ATTRIBUTE_FORM = pkk
End Function
Public Function WPD_PROPERTY_ATTRIBUTE_CAN_READ() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAB7943D8, &H6332, &H445F, &HA0, &HD, &H8D, &H5E, &HF1, &HE9, &H6F, &H37, 3)
 WPD_PROPERTY_ATTRIBUTE_CAN_READ = pkk
End Function
Public Function WPD_PROPERTY_ATTRIBUTE_CAN_WRITE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAB7943D8, &H6332, &H445F, &HA0, &HD, &H8D, &H5E, &HF1, &HE9, &H6F, &H37, 4)
 WPD_PROPERTY_ATTRIBUTE_CAN_WRITE = pkk
End Function
Public Function WPD_PROPERTY_ATTRIBUTE_CAN_DELETE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAB7943D8, &H6332, &H445F, &HA0, &HD, &H8D, &H5E, &HF1, &HE9, &H6F, &H37, 5)
 WPD_PROPERTY_ATTRIBUTE_CAN_DELETE = pkk
End Function
Public Function WPD_PROPERTY_ATTRIBUTE_DEFAULT_VALUE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAB7943D8, &H6332, &H445F, &HA0, &HD, &H8D, &H5E, &HF1, &HE9, &H6F, &H37, 6)
 WPD_PROPERTY_ATTRIBUTE_DEFAULT_VALUE = pkk
End Function
Public Function WPD_PROPERTY_ATTRIBUTE_FAST_PROPERTY() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAB7943D8, &H6332, &H445F, &HA0, &HD, &H8D, &H5E, &HF1, &HE9, &H6F, &H37, 7)
 WPD_PROPERTY_ATTRIBUTE_FAST_PROPERTY = pkk
End Function
Public Function WPD_PROPERTY_ATTRIBUTE_RANGE_MIN() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAB7943D8, &H6332, &H445F, &HA0, &HD, &H8D, &H5E, &HF1, &HE9, &H6F, &H37, 8)
 WPD_PROPERTY_ATTRIBUTE_RANGE_MIN = pkk
End Function
Public Function WPD_PROPERTY_ATTRIBUTE_RANGE_MAX() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAB7943D8, &H6332, &H445F, &HA0, &HD, &H8D, &H5E, &HF1, &HE9, &H6F, &H37, 9)
 WPD_PROPERTY_ATTRIBUTE_RANGE_MAX = pkk
End Function
Public Function WPD_PROPERTY_ATTRIBUTE_RANGE_STEP() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAB7943D8, &H6332, &H445F, &HA0, &HD, &H8D, &H5E, &HF1, &HE9, &H6F, &H37, 10)
 WPD_PROPERTY_ATTRIBUTE_RANGE_STEP = pkk
End Function
Public Function WPD_PROPERTY_ATTRIBUTE_ENUMERATION_ELEMENTS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAB7943D8, &H6332, &H445F, &HA0, &HD, &H8D, &H5E, &HF1, &HE9, &H6F, &H37, 11)
 WPD_PROPERTY_ATTRIBUTE_ENUMERATION_ELEMENTS = pkk
End Function
Public Function WPD_PROPERTY_ATTRIBUTE_REGULAR_EXPRESSION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAB7943D8, &H6332, &H445F, &HA0, &HD, &H8D, &H5E, &HF1, &HE9, &H6F, &H37, 12)
 WPD_PROPERTY_ATTRIBUTE_REGULAR_EXPRESSION = pkk
End Function
Public Function WPD_PROPERTY_ATTRIBUTE_MAX_SIZE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAB7943D8, &H6332, &H445F, &HA0, &HD, &H8D, &H5E, &HF1, &HE9, &H6F, &H37, 13)
 WPD_PROPERTY_ATTRIBUTE_MAX_SIZE = pkk
End Function
Public Function WPD_PROPERTY_ATTRIBUTE_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5D9DA160, &H74AE, &H43CC, &H85, &HA9, &HFE, &H55, &H5A, &H80, &H79, &H8E, 2)
 WPD_PROPERTY_ATTRIBUTE_NAME = pkk
End Function
Public Function WPD_PROPERTY_ATTRIBUTE_VARTYPE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5D9DA160, &H74AE, &H43CC, &H85, &HA9, &HFE, &H55, &H5A, &H80, &H79, &H8E, 3)
 WPD_PROPERTY_ATTRIBUTE_VARTYPE = pkk
End Function
Public Function WPD_CLASS_EXTENSION_OPTIONS_SUPPORTED_CONTENT_TYPES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6309FFEF, &HA87C, &H4CA7, &H84, &H34, &H79, &H75, &H76, &HE4, &HA, &H96, 2)
 WPD_CLASS_EXTENSION_OPTIONS_SUPPORTED_CONTENT_TYPES = pkk
End Function
Public Function WPD_CLASS_EXTENSION_OPTIONS_DONT_REGISTER_WPD_DEVICE_INTERFACE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6309FFEF, &HA87C, &H4CA7, &H84, &H34, &H79, &H75, &H76, &HE4, &HA, &H96, 3)
 WPD_CLASS_EXTENSION_OPTIONS_DONT_REGISTER_WPD_DEVICE_INTERFACE = pkk
End Function
Public Function WPD_CLASS_EXTENSION_OPTIONS_REGISTER_WPD_PRIVATE_DEVICE_INTERFACE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6309FFEF, &HA87C, &H4CA7, &H84, &H34, &H79, &H75, &H76, &HE4, &HA, &H96, 4)
 WPD_CLASS_EXTENSION_OPTIONS_REGISTER_WPD_PRIVATE_DEVICE_INTERFACE = pkk
End Function
Public Function WPD_CLASS_EXTENSION_OPTIONS_MULTITRANSPORT_MODE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3E3595DA, &H4D71, &H49FE, &HA0, &HB4, &HD4, &H40, &H6C, &H3A, &HE9, &H3F, 2)
 WPD_CLASS_EXTENSION_OPTIONS_MULTITRANSPORT_MODE = pkk
End Function
Public Function WPD_CLASS_EXTENSION_OPTIONS_DEVICE_IDENTIFICATION_VALUES() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3E3595DA, &H4D71, &H49FE, &HA0, &HB4, &HD4, &H40, &H6C, &H3A, &HE9, &H3F, 3)
 WPD_CLASS_EXTENSION_OPTIONS_DEVICE_IDENTIFICATION_VALUES = pkk
End Function
Public Function WPD_CLASS_EXTENSION_OPTIONS_TRANSPORT_BANDWIDTH() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3E3595DA, &H4D71, &H49FE, &HA0, &HB4, &HD4, &H40, &H6C, &H3A, &HE9, &H3F, 4)
 WPD_CLASS_EXTENSION_OPTIONS_TRANSPORT_BANDWIDTH = pkk
End Function


