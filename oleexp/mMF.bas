Attribute VB_Name = "mMF"
Option Explicit

Private Sub DEFINE_UUID(Name As UUID, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
  With Name
    .Data1 = L: .Data2 = w1: .Data3 = w2: .Data4(0) = B0: .Data4(1) = b1: .Data4(2) = b2: .Data4(3) = B3: .Data4(4) = b4: .Data4(5) = b5: .Data4(6) = b6: .Data4(7) = b7
  End With
End Sub
Private Sub DEFINE_PROPERTYKEY(Name As PROPERTYKEY, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte, pid As Long)
  With Name.fmtid
    .Data1 = L: .Data2 = w1: .Data3 = w2: .Data4(0) = B0: .Data4(1) = b1: .Data4(2) = b2: .Data4(3) = B3: .Data4(4) = b4: .Data4(5) = b5: .Data4(6) = b6: .Data4(7) = b7
  End With
  Name.pid = pid
End Sub
Public Function IID_IMFMediaSession() As UUID
'{90377834-21D0-4dee-8214-BA2E3E6C1127}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H90377834, CInt(&H21D0), CInt(&H4DEE), &H82, &H14, &HBA, &H2E, &H3E, &H6C, &H11, &H27)
IID_IMFMediaSession = iid
End Function
Public Function IID_IMFSourceResolver() As UUID
'{FBE5A32D-A497-4B61-BB85-97B1A848A6E3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFBE5A32D, CInt(&HA497), CInt(&H4B61), &HBB, &H85, &H97, &HB1, &HA8, &H48, &HA6, &HE3)
IID_IMFSourceResolver = iid
End Function
Public Function IID_IMFByteStream() As UUID
'{AD4C1B00-4BF7-422F-9175-756693D9130D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAD4C1B00, CInt(&H4BF7), CInt(&H422F), &H91, &H75, &H75, &H66, &H93, &HD9, &H13, &HD)
IID_IMFByteStream = iid
End Function
Public Function IID_IMFAsyncCallback() As UUID
'{A27003CF-2354-4F2A-8D6A-AB7CFF15437E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA27003CF, CInt(&H2354), CInt(&H4F2A), &H8D, &H6A, &HAB, &H7C, &HFF, &H15, &H43, &H7E)
IID_IMFAsyncCallback = iid
End Function
Public Function IID_IMFAsyncResult() As UUID
'{AC6B7889-0740-4D51-8619-905994A55CC6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAC6B7889, CInt(&H740), CInt(&H4D51), &H86, &H19, &H90, &H59, &H94, &HA5, &H5C, &HC6)
IID_IMFAsyncResult = iid
End Function
Public Function IID_IMFAttributes() As UUID
'{2CD2D921-C447-44A7-A13C-4ADABFC247E3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD2D921, CInt(&HC447), CInt(&H44A7), &HA1, &H3C, &H4A, &HDA, &HBF, &HC2, &H47, &HE3)
IID_IMFAttributes = iid
End Function
Public Function IID_IMFMediaEventGenerator() As UUID
'{2CD0BD52-BCD5-4B89-B62C-EADC0C031E7D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD0BD52, CInt(&HBCD5), CInt(&H4B89), &HB6, &H2C, &HEA, &HDC, &HC, &H3, &H1E, &H7D)
IID_IMFMediaEventGenerator = iid
End Function
Public Function IID_IMFMediaEvent() As UUID
'{2CD0BD52-BCD5-4B89-B62C-EADC0C031E7D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD0BD52, CInt(&HBCD5), CInt(&H4B89), &HB6, &H2C, &HEA, &HDC, &HC, &H3, &H1E, &H7D)
IID_IMFMediaEvent = iid
End Function
Public Function IID_IMFReadWriteClassFactory() As UUID
'{E7FE2E12-661C-40DA-92F9-4F002AB67627}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE7FE2E12, CInt(&H661C), CInt(&H40DA), &H92, &HF9, &H4F, &H0, &H2A, &HB6, &H76, &H27)
 IID_IMFReadWriteClassFactory = iid
End Function
Public Function IID_IMFMediaSource() As UUID
'{279A808D-AEC7-40C8-9C6B-A6B492C78A66}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H279A808D, CInt(&HAEC7), CInt(&H40C8), &H9C, &H6B, &HA6, &HB4, &H92, &HC7, &H8A, &H66)
IID_IMFMediaSource = iid
End Function
Public Function IID_IMFPresentationDescriptor() As UUID
'{03CB2711-24D7-4DB6-A17F-F3A7A479A536}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3CB2711, CInt(&H24D7), CInt(&H4DB6), &HA1, &H7F, &HF3, &HA7, &HA4, &H79, &HA5, &H36)
IID_IMFPresentationDescriptor = iid
End Function
Public Function IID_IMFStreamDescriptor() As UUID
'{56C03D9C-9DBB-45F5-AB4B-D80F47C05938}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56C03D9C, CInt(&H9DBB), CInt(&H45F5), &HAB, &H4B, &HD8, &HF, &H47, &HC0, &H59, &H38)
IID_IMFStreamDescriptor = iid
End Function
Public Function IID_IMFMediaTypeHandler() As UUID
'{E93DCF6C-4B07-4E1E-8123-AA16ED6EADF5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE93DCF6C, CInt(&H4B07), CInt(&H4E1E), &H81, &H23, &HAA, &H16, &HED, &H6E, &HAD, &HF5)
IID_IMFMediaTypeHandler = iid
End Function
Public Function IID_IMFMediaType() As UUID
'{44AE0FA8-EA31-4109-8D2E-4CAE4997C555}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H44AE0FA8, CInt(&HEA31), CInt(&H4109), &H8D, &H2E, &H4C, &HAE, &H49, &H97, &HC5, &H55)
IID_IMFMediaType = iid
End Function
Public Function IID_IMFSourceReader() As UUID
'{70AE66F2-C809-4E4F-8915-BDCB406B7993}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H70AE66F2, CInt(&HC809), CInt(&H4E4F), &H89, &H15, &HBD, &HCB, &H40, &H6B, &H79, &H93)
IID_IMFSourceReader = iid
End Function
Public Function IID_IMFSourceReaderEx() As UUID
'{7b981cf0-560e-4116-9875-b099895f23d7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7B981CF0, CInt(&H560E), CInt(&H4116), &H98, &H75, &HB0, &H99, &H89, &H5F, &H23, &HD7)
IID_IMFSourceReaderEx = iid
End Function
Public Function IID_IMFSourceReaderCallback() As UUID
'{deec8d99-fa1d-4d82-84c2-2c8969944867}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDEEC8D99, CInt(&HFA1D), CInt(&H4D82), &H84, &HC2, &H2C, &H89, &H69, &H94, &H48, &H67)
IID_IMFSourceReaderCallback = iid
End Function
Public Function IID_IMFSourceReaderCallback2() As UUID
'{CF839FE6-8C2A-4DD2-B6EA-C22D6961AF05}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCF839FE6, CInt(&H8C2A), CInt(&H4DD2), &HB6, &HEA, &HC2, &H2D, &H69, &H61, &HAF, &H5)
IID_IMFSourceReaderCallback2 = iid
End Function
Public Function IID_IMFSinkWriter() As UUID
'{3137f1cd-fe5e-4805-a5d8-fb477448cb3d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3137F1CD, CInt(&HFE5E), CInt(&H4805), &HA5, &HD8, &HFB, &H47, &H74, &H48, &HCB, &H3D)
IID_IMFSinkWriter = iid
End Function
Public Function IID_IMFSinkWriterEx() As UUID
'{588d72ab-5Bc1-496a-8714-b70617141b25}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H588D72AB, CInt(&H5BC1), CInt(&H496A), &H87, &H14, &HB7, &H6, &H17, &H14, &H1B, &H25)
IID_IMFSinkWriterEx = iid
End Function
Public Function IID_IMFSinkWriterEncoderConfig() As UUID
'{17C3779E-3CDE-4EDE-8C60-3899F5F53AD6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H17C3779E, CInt(&H3CDE), CInt(&H4EDE), &H8C, &H60, &H38, &H99, &HF5, &HF5, &H3A, &HD6)
IID_IMFSinkWriterEncoderConfig = iid
End Function
Public Function IID_IMFSinkWriterCallback() As UUID
'{666f76de-33d2-41b9-a458-29ed0a972c58}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H666F76DE, CInt(&H33D2), CInt(&H41B9), &HA4, &H58, &H29, &HED, &HA, &H97, &H2C, &H58)
IID_IMFSinkWriterCallback = iid
End Function
Public Function IID_IMFSinkWriterCallback2() As UUID
'{2456BD58-C067-4513-84FE-8D0C88FFDC61}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2456BD58, CInt(&HC067), CInt(&H4513), &H84, &HFE, &H8D, &HC, &H88, &HFF, &HDC, &H61)
IID_IMFSinkWriterCallback2 = iid
End Function
Public Function IID_IMFSample() As UUID
'{C40A00F2-B93A-4D80-AE8C-5A1C634F58E4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC40A00F2, CInt(&HB93A), CInt(&H4D80), &HAE, &H8C, &H5A, &H1C, &H63, &H4F, &H58, &HE4)
IID_IMFSample = iid
End Function
Public Function IID_IMFMediaBuffer() As UUID
'{045FA593-8799-42B8-BC8D-8968C6453507}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H45FA593, CInt(&H8799), CInt(&H42B8), &HBC, &H8D, &H89, &H68, &HC6, &H45, &H35, &H7)
IID_IMFMediaBuffer = iid
End Function
Public Function IID_IMFClock() As UUID
'{2eb1e945-18b8-4139-9b1a-d5d584818530}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2EB1E945, CInt(&H18B8), CInt(&H4139), &H9B, &H1A, &HD5, &HD5, &H84, &H81, &H85, &H30)
IID_IMFClock = iid
End Function
Public Function IID_IMFCollection() As UUID
'{5BC8A76B-869A-46a3-9B03-FA218A66AEBE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5BC8A76B, CInt(&H869A), CInt(&H46A3), &H9B, &H3, &HFA, &H21, &H8A, &H66, &HAE, &HBE)
IID_IMFCollection = iid
End Function
Public Function IID_IMF2DBuffer() As UUID
'{7dc9d5f9-9ed9-44ec-9bbf-0600bb589fbb}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7DC9D5F9, CInt(&H9ED9), CInt(&H44EC), &H9B, &HBF, &H6, &H0, &HBB, &H58, &H9F, &HBB)
IID_IMF2DBuffer = iid
End Function
Public Function IID_IMF2DBuffer2() As UUID
'{33ae5ea6-4316-436f-8ddd-d73d22f829ec}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H33AE5EA6, CInt(&H4316), CInt(&H436F), &H8D, &HDD, &HD7, &H3D, &H22, &HF8, &H29, &HEC)
IID_IMF2DBuffer2 = iid
End Function
Public Function IID_IMFDXGIBuffer() As UUID
'{e7174cfa-1c9e-48b1-8866-626226bfc258}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE7174CFA, CInt(&H1C9E), CInt(&H48B1), &H88, &H66, &H62, &H62, &H26, &HBF, &HC2, &H58)
IID_IMFDXGIBuffer = iid
End Function
Public Function IID_IMFTopologyNode() As UUID
'{83CF873A-F6DA-4bc8-823F-BACFD55DC430}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H83CF873A, CInt(&HF6DA), CInt(&H4BC8), &H82, &H3F, &HBA, &HCF, &HD5, &H5D, &HC4, &H30)
IID_IMFTopologyNode = iid
End Function
Public Function IID_IMFTopology() As UUID
'{83CF873A-F6DA-4bc8-823F-BACFD55DC433}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H83CF873A, CInt(&HF6DA), CInt(&H4BC8), &H82, &H3F, &HBA, &HCF, &HD5, &H5D, &HC4, &H33)
IID_IMFTopology = iid
End Function
Public Function IID_IMediaObject() As UUID
'{d8ad0f58-5494-4102-97c5-ec798e59bcf4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD8AD0F58, CInt(&H5494), CInt(&H4102), &H97, &HC5, &HEC, &H79, &H8E, &H59, &HBC, &HF4)
IID_IMediaObject = iid
End Function
Public Function IID_IEnumDMO() As UUID
'{2c3cd98a-2bfa-4a53-9c27-5249ba64ba0f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2C3CD98A, CInt(&H2BFA), CInt(&H4A53), &H9C, &H27, &H52, &H49, &HBA, &H64, &HBA, &HF)
IID_IEnumDMO = iid
End Function
Public Function IID_IMediaObjectInPlace() As UUID
'{651b9ad0-0fc7-4aa9-9538-d89931010741}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H651B9AD0, CInt(&HFC7), CInt(&H4AA9), &H95, &H38, &HD8, &H99, &H31, &H1, &H7, &H41)
IID_IMediaObjectInPlace = iid
End Function
Public Function IID_IDMOQualityControl() As UUID
'{65abea96-cf36-453f-af8a-705e98f16260}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H65ABEA96, CInt(&HCF36), CInt(&H453F), &HAF, &H8A, &H70, &H5E, &H98, &HF1, &H62, &H60)
IID_IDMOQualityControl = iid
End Function
Public Function IID_IDMOVideoOutputOptimizations() As UUID
'{be8f4f4e-5b16-4d29-b350-7f6b5d9298ac}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBE8F4F4E, CInt(&H5B16), CInt(&H4D29), &HB3, &H50, &H7F, &H6B, &H5D, &H92, &H98, &HAC)
IID_IDMOVideoOutputOptimizations = iid
End Function
Public Function IID_IMFAudioMediaType() As UUID
'{26a0adc3-ce26-4672-9304-69552edd3faf}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H26A0ADC3, CInt(&HCE26), CInt(&H4672), &H93, &H4, &H69, &H55, &H2E, &HDD, &H3F, &HAF)
IID_IMFAudioMediaType = iid
End Function
Public Function IID_IMFVideoMediaType() As UUID
'{b99f381f-a8f9-47a2-a5af-ca3a225a3890}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB99F381F, CInt(&HA8F9), CInt(&H47A2), &HA5, &HAF, &HCA, &H3A, &H22, &H5A, &H38, &H90)
IID_IMFVideoMediaType = iid
End Function
Public Function IID_IMFAsyncCallbackLogging() As UUID
'{c7a4dca1-f5f0-47b6-b92b-bf0106d25791}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC7A4DCA1, CInt(&HF5F0), CInt(&H47B6), &HB9, &H2B, &HBF, &H1, &H6, &HD2, &H57, &H91)
IID_IMFAsyncCallbackLogging = iid
End Function
Public Function IID_IMFByteStreamProxyClassFactory() As UUID
'{a6b43f84-5c0a-42e8-a44d-b1857a76992f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA6B43F84, CInt(&H5C0A), CInt(&H42E8), &HA4, &H4D, &HB1, &H85, &H7A, &H76, &H99, &H2F)
IID_IMFByteStreamProxyClassFactory = iid
End Function
Public Function IID_IMFSampleOutputStream() As UUID
'{8feed468-6f7e-440d-869a-49bdd283ad0d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8FEED468, CInt(&H6F7E), CInt(&H440D), &H86, &H9A, &H49, &HBD, &HD2, &H83, &HAD, &HD)
IID_IMFSampleOutputStream = iid
End Function
Public Function IID_IMFMediaEventQueue() As UUID
'{36f846fc-2256-48b6-b58e-e2b638316581}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H36F846FC, CInt(&H2256), CInt(&H48B6), &HB5, &H8E, &HE2, &HB6, &H38, &H31, &H65, &H81)
IID_IMFMediaEventQueue = iid
End Function
Public Function IID_IMFActivate() As UUID
'{7FEE9E9A-4A89-47a6-899C-B6A53A70FB67}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7FEE9E9A, CInt(&H4A89), CInt(&H47A6), &H89, &H9C, &HB6, &HA5, &H3A, &H70, &HFB, &H67)
IID_IMFActivate = iid
End Function
Public Function IID_IMFPluginControl() As UUID
'{5c6c44bf-1db6-435b-9249-e8cd10fdec96}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5C6C44BF, CInt(&H1DB6), CInt(&H435B), &H92, &H49, &HE8, &HCD, &H10, &HFD, &HEC, &H96)
IID_IMFPluginControl = iid
End Function
Public Function IID_IMFPluginControl2() As UUID
'{C6982083-3DDC-45CB-AF5E-0F7A8CE4DE77}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC6982083, CInt(&H3DDC), CInt(&H45CB), &HAF, &H5E, &HF, &H7A, &H8C, &HE4, &HDE, &H77)
IID_IMFPluginControl2 = iid
End Function
Public Function IID_IMFDXGIDeviceManager() As UUID
'{eb533d5d-2db6-40f8-97a9-494692014f07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEB533D5D, CInt(&H2DB6), CInt(&H40F8), &H97, &HA9, &H49, &H46, &H92, &H1, &H4F, &H7)
IID_IMFDXGIDeviceManager = iid
End Function
Public Function IID_IMFTransform() As UUID
'{bf94c121-5b05-4e6f-8000-ba598961414d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBF94C121, CInt(&H5B05), CInt(&H4E6F), &H80, &H0, &HBA, &H59, &H89, &H61, &H41, &H4D)
IID_IMFTransform = iid
End Function
Public Function IID_IMFDeviceTransform() As UUID
'{D818FBD8-FC46-42F2-87AC-1EA2D1F9BF32}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD818FBD8, CInt(&HFC46), CInt(&H42F2), &H87, &HAC, &H1E, &HA2, &HD1, &HF9, &HBF, &H32)
 IID_IMFDeviceTransform = iid
End Function
Public Function IID_IMFDeviceTransformCallback() As UUID
'{6D5CB646-29EC-41FB-8179-8C4C6D750811}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6D5CB646, CInt(&H29EC), CInt(&H41FB), &H81, &H79, &H8C, &H4C, &H6D, &H75, &H8, &H11)
 IID_IMFDeviceTransformCallback = iid
End Function
Public Function IID_IMFMediaSourceEx() As UUID
'{3C9B2EB9-86D5-4514-A394-F56664F9F0D8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3C9B2EB9, CInt(&H86D5), CInt(&H4514), &HA3, &H94, &HF5, &H66, &H64, &HF9, &HF0, &HD8)
IID_IMFMediaSourceEx = iid
End Function
Public Function IID_IMFClockConsumer() As UUID
'{6ef2a662-47c0-4666-b13d-cbb717f2fa2c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6EF2A662, CInt(&H47C0), CInt(&H4666), &HB1, &H3D, &HCB, &HB7, &H17, &HF2, &HFA, &H2C)
IID_IMFClockConsumer = iid
End Function
Public Function IID_IMFMediaStream() As UUID
'{D182108F-4EC6-443f-AA42-A71106EC825F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD182108F, CInt(&H4EC6), CInt(&H443F), &HAA, &H42, &HA7, &H11, &H6, &HEC, &H82, &H5F)
IID_IMFMediaStream = iid
End Function
Public Function IID_IMFMediaSink() As UUID
'{6ef2a660-47c0-4666-b13d-cbb717f2fa2c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6EF2A660, CInt(&H47C0), CInt(&H4666), &HB1, &H3D, &HCB, &HB7, &H17, &HF2, &HFA, &H2C)
IID_IMFMediaSink = iid
End Function
Public Function IID_IMFStreamSink() As UUID
'{0A97B3CF-8E7C-4a3d-8F8C-0C843DC247FB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA97B3CF, CInt(&H8E7C), CInt(&H4A3D), &H8F, &H8C, &HC, &H84, &H3D, &HC2, &H47, &HFB)
IID_IMFStreamSink = iid
End Function
Public Function IID_IMFVideoSampleAllocator() As UUID
'{86cbc910-e533-4751-8e3b-f19b5b806a03}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H86CBC910, CInt(&HE533), CInt(&H4751), &H8E, &H3B, &HF1, &H9B, &H5B, &H80, &H6A, &H3)
IID_IMFVideoSampleAllocator = iid
End Function
Public Function IID_IMFVideoSampleAllocatorNotify() As UUID
'{A792CDBE-C374-4e89-8335-278E7B9956A4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA792CDBE, CInt(&HC374), CInt(&H4E89), &H83, &H35, &H27, &H8E, &H7B, &H99, &H56, &HA4)
IID_IMFVideoSampleAllocatorNotify = iid
End Function
Public Function IID_IMFVideoSampleAllocatorNotifyEx() As UUID
'{3978AA1A-6D5B-4B7F-A340-90899189AE34}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3978AA1A, CInt(&H6D5B), CInt(&H4B7F), &HA3, &H40, &H90, &H89, &H91, &H89, &HAE, &H34)
IID_IMFVideoSampleAllocatorNotifyEx = iid
End Function
Public Function IID_IMFVideoSampleAllocatorCallback() As UUID
'{992388B4-3372-4f67-8B6F-C84C071F4751}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H992388B4, CInt(&H3372), CInt(&H4F67), &H8B, &H6F, &HC8, &H4C, &H7, &H1F, &H47, &H51)
IID_IMFVideoSampleAllocatorCallback = iid
End Function
Public Function IID_IMFVideoSampleAllocatorEx() As UUID
'{545b3a48-3283-4f62-866f-a62d8f598f9f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H545B3A48, CInt(&H3283), CInt(&H4F62), &H86, &H6F, &HA6, &H2D, &H8F, &H59, &H8F, &H9F)
IID_IMFVideoSampleAllocatorEx = iid
End Function
Public Function IID_IMFDXGIDeviceManagerSource() As UUID
'{20bc074b-7a8d-4609-8c3b-64a0a3b5d7ce}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20BC074B, CInt(&H7A8D), CInt(&H4609), &H8C, &H3B, &H64, &HA0, &HA3, &HB5, &HD7, &HCE)
IID_IMFDXGIDeviceManagerSource = iid
End Function
Public Function IID_IMFVideoProcessorControl() As UUID
'{A3F675D5-6119-4f7f-A100-1D8B280F0EFB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA3F675D5, CInt(&H6119), CInt(&H4F7F), &HA1, &H0, &H1D, &H8B, &H28, &HF, &HE, &HFB)
IID_IMFVideoProcessorControl = iid
End Function
Public Function IID_IMFVideoProcessorControl2() As UUID
'{BDE633D3-E1DC-4a7f-A693-BBAE399C4A20}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBDE633D3, CInt(&HE1DC), CInt(&H4A7F), &HA6, &H93, &HBB, &HAE, &H39, &H9C, &H4A, &H20)
IID_IMFVideoProcessorControl2 = iid
End Function
Public Function IID_IMFGetService() As UUID
'{fa993888-4383-415a-a930-dd472a8cf6f7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA993888, CInt(&H4383), CInt(&H415A), &HA9, &H30, &HDD, &H47, &H2A, &H8C, &HF6, &HF7)
IID_IMFGetService = iid
End Function
Public Function IID_IMFPresentationClock() As UUID
'{868CE85C-8EA9-4f55-AB82-B009A910A805}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H868CE85C, CInt(&H8EA9), CInt(&H4F55), &HAB, &H82, &HB0, &H9, &HA9, &H10, &HA8, &H5)
IID_IMFPresentationClock = iid
End Function
Public Function IID_IMFPresentationTimeSource() As UUID
'{7FF12CCE-F76F-41c2-863B-1666C8E5E139}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7FF12CCE, CInt(&HF76F), CInt(&H41C2), &H86, &H3B, &H16, &H66, &HC8, &HE5, &HE1, &H39)
IID_IMFPresentationTimeSource = iid
End Function
Public Function IID_IMFClockStateSink() As UUID
'{F6696E82-74F7-4f3d-A178-8A5E09C3659F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF6696E82, CInt(&H74F7), CInt(&H4F3D), &HA1, &H78, &H8A, &H5E, &H9, &HC3, &H65, &H9F)
IID_IMFClockStateSink = iid
End Function
Public Function IID_IMFTimer() As UUID
'{e56e4cbd-8f70-49d8-a0f8-edb3d6ab9bf2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE56E4CBD, CInt(&H8F70), CInt(&H49D8), &HA0, &HF8, &HED, &HB3, &HD6, &HAB, &H9B, &HF2)
IID_IMFTimer = iid
End Function
Public Function IID_IMFShutdown() As UUID
'{97ec2ea4-0e42-4937-97ac-9d6d328824e1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H97EC2EA4, CInt(&HE42), CInt(&H4937), &H97, &HAC, &H9D, &H6D, &H32, &H88, &H24, &HE1)
IID_IMFShutdown = iid
End Function
Public Function IID_IMFTopoLoader() As UUID
'{DE9A6157-F660-4643-B56A-DF9F7998C7CD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDE9A6157, CInt(&HF660), CInt(&H4643), &HB5, &H6A, &HDF, &H9F, &H79, &H98, &HC7, &HCD)
IID_IMFTopoLoader = iid
End Function
Public Function IID_IMFContentProtectionManager() As UUID
'{ACF92459-6A61-42bd-B57C-B43E51203CB0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HACF92459, CInt(&H6A61), CInt(&H42BD), &HB5, &H7C, &HB4, &H3E, &H51, &H20, &H3C, &HB0)
IID_IMFContentProtectionManager = iid
End Function
Public Function IID_IMFContentEnabler() As UUID
'{D3C4EF59-49CE-4381-9071-D5BCD044C770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD3C4EF59, CInt(&H49CE), CInt(&H4381), &H90, &H71, &HD5, &HBC, &HD0, &H44, &HC7, &H70)
IID_IMFContentEnabler = iid
End Function
Public Function IID_IMFMetadata() As UUID
'{F88CFB8C-EF16-4991-B450-CB8C69E51704}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF88CFB8C, CInt(&HEF16), CInt(&H4991), &HB4, &H50, &HCB, &H8C, &H69, &HE5, &H17, &H4)
IID_IMFMetadata = iid
End Function
Public Function IID_IMFMetadataProvider() As UUID
'{56181D2D-E221-4adb-B1C8-3CEE6A53F76F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56181D2D, CInt(&HE221), CInt(&H4ADB), &HB1, &HC8, &H3C, &HEE, &H6A, &H53, &HF7, &H6F)
IID_IMFMetadataProvider = iid
End Function
Public Function IID_IMFRateSupport() As UUID
'{0a9ccdbc-d797-4563-9667-94ec5d79292d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA9CCDBC, CInt(&HD797), CInt(&H4563), &H96, &H67, &H94, &HEC, &H5D, &H79, &H29, &H2D)
IID_IMFRateSupport = iid
End Function
Public Function IID_IMFRateControl() As UUID
'{88ddcd21-03c3-4275-91ed-55ee3929328f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H88DDCD21, CInt(&H3C3), CInt(&H4275), &H91, &HED, &H55, &HEE, &H39, &H29, &H32, &H8F)
IID_IMFRateControl = iid
End Function
Public Function IID_IMFTimecodeTranslate() As UUID
'{ab9d8661-f7e8-4ef4-9861-89f334f94e74}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB9D8661, CInt(&HF7E8), CInt(&H4EF4), &H98, &H61, &H89, &HF3, &H34, &HF9, &H4E, &H74)
IID_IMFTimecodeTranslate = iid
End Function
Public Function IID_IMFSeekInfo() As UUID
'{26AFEA53-D9ED-42B5-AB80-E64F9EE34779}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H26AFEA53, CInt(&HD9ED), CInt(&H42B5), &HAB, &H80, &HE6, &H4F, &H9E, &HE3, &H47, &H79)
IID_IMFSeekInfo = iid
End Function
Public Function IID_IMFSimpleAudioVolume() As UUID
'{089EDF13-CF71-4338-8D13-9E569DBDC319}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H89EDF13, CInt(&HCF71), CInt(&H4338), &H8D, &H13, &H9E, &H56, &H9D, &HBD, &HC3, &H19)
IID_IMFSimpleAudioVolume = iid
End Function
Public Function IID_IMFAudioStreamVolume() As UUID
'{76B1BBDB-4EC8-4f36-B106-70A9316DF593}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H76B1BBDB, CInt(&H4EC8), CInt(&H4F36), &HB1, &H6, &H70, &HA9, &H31, &H6D, &HF5, &H93)
IID_IMFAudioStreamVolume = iid
End Function
Public Function IID_IMFAudioPolicy() As UUID
'{a0638c2b-6465-4395-9ae7-a321a9fd2856}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA0638C2B, CInt(&H6465), CInt(&H4395), &H9A, &HE7, &HA3, &H21, &HA9, &HFD, &H28, &H56)
IID_IMFAudioPolicy = iid
End Function
Public Function IID_IMFSampleGrabberSinkCallback() As UUID
'{8C7B80BF-EE42-4b59-B1DF-55668E1BDCA8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8C7B80BF, CInt(&HEE42), CInt(&H4B59), &HB1, &HDF, &H55, &H66, &H8E, &H1B, &HDC, &HA8)
IID_IMFSampleGrabberSinkCallback = iid
End Function
Public Function IID_IMFSampleGrabberSinkCallback2() As UUID
'{ca86aa50-c46e-429e-ab27-16d6ac6844cb}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCA86AA50, CInt(&HC46E), CInt(&H429E), &HAB, &H27, &H16, &HD6, &HAC, &H68, &H44, &HCB)
IID_IMFSampleGrabberSinkCallback2 = iid
End Function
Public Function IID_IMFWorkQueueServices() As UUID
'{35FE1BB8-A3A9-40fe-BBEC-EB569C9CCCA3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H35FE1BB8, CInt(&HA3A9), CInt(&H40FE), &HBB, &HEC, &HEB, &H56, &H9C, &H9C, &HCC, &HA3)
IID_IMFWorkQueueServices = iid
End Function
Public Function IID_IMFWorkQueueServicesEx() As UUID
'{96bf961b-40fe-42f1-ba9d-320238b49700}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H96BF961B, CInt(&H40FE), CInt(&H42F1), &HBA, &H9D, &H32, &H2, &H38, &HB4, &H97, &H0)
IID_IMFWorkQueueServicesEx = iid
End Function
Public Function IID_IMFQualityManager() As UUID
'{8D009D86-5B9F-4115-B1FC-9F80D52AB8AB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8D009D86, CInt(&H5B9F), CInt(&H4115), &HB1, &HFC, &H9F, &H80, &HD5, &H2A, &HB8, &HAB)
IID_IMFQualityManager = iid
End Function
Public Function IID_IMFQualityAdvise() As UUID
'{EC15E2E9-E36B-4f7c-8758-77D452EF4CE7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEC15E2E9, CInt(&HE36B), CInt(&H4F7C), &H87, &H58, &H77, &HD4, &H52, &HEF, &H4C, &HE7)
IID_IMFQualityAdvise = iid
End Function
Public Function IID_IMFQualityAdvise2() As UUID
'{F3706F0D-8EA2-4886-8000-7155E9EC2EAE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF3706F0D, CInt(&H8EA2), CInt(&H4886), &H80, &H0, &H71, &H55, &HE9, &HEC, &H2E, &HAE)
IID_IMFQualityAdvise2 = iid
End Function
Public Function IID_IMFQualityAdviseLimits() As UUID
'{dfcd8e4d-30b5-4567-acaa-8eb5b7853dc9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDFCD8E4D, CInt(&H30B5), CInt(&H4567), &HAC, &HAA, &H8E, &HB5, &HB7, &H85, &H3D, &HC9)
IID_IMFQualityAdviseLimits = iid
End Function
Public Function IID_IMFRealTimeClient() As UUID
'{2347D60B-3FB5-480c-8803-8DF3ADCD3EF0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2347D60B, CInt(&H3FB5), CInt(&H480C), &H88, &H3, &H8D, &HF3, &HAD, &HCD, &H3E, &HF0)
IID_IMFRealTimeClient = iid
End Function
Public Function IID_IMFRealTimeClientEx() As UUID
'{03910848-AB16-4611-B100-17B88AE2F248}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3910848, CInt(&HAB16), CInt(&H4611), &HB1, &H0, &H17, &HB8, &H8A, &HE2, &HF2, &H48)
IID_IMFRealTimeClientEx = iid
End Function
Public Function IID_IMFSequencerSource() As UUID
'{197CD219-19CB-4de1-A64C-ACF2EDCBE59E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H197CD219, CInt(&H19CB), CInt(&H4DE1), &HA6, &H4C, &HAC, &HF2, &HED, &HCB, &HE5, &H9E)
IID_IMFSequencerSource = iid
End Function
Public Function IID_IMFMediaSourceTopologyProvider() As UUID
'{0E1D6009-C9F3-442d-8C51-A42D2D49452F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE1D6009, CInt(&HC9F3), CInt(&H442D), &H8C, &H51, &HA4, &H2D, &H2D, &H49, &H45, &H2F)
IID_IMFMediaSourceTopologyProvider = iid
End Function
Public Function IID_IMFMediaSourcePresentationProvider() As UUID
'{0E1D600a-C9F3-442d-8C51-A42D2D49452F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE1D600A, CInt(&HC9F3), CInt(&H442D), &H8C, &H51, &HA4, &H2D, &H2D, &H49, &H45, &H2F)
IID_IMFMediaSourcePresentationProvider = iid
End Function
Public Function IID_IMFTopologyNodeAttributeEditor() As UUID
'{676aa6dd-238a-410d-bb99-65668d01605a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H676AA6DD, CInt(&H238A), CInt(&H410D), &HBB, &H99, &H65, &H66, &H8D, &H1, &H60, &H5A)
IID_IMFTopologyNodeAttributeEditor = iid
End Function
Public Function IID_IMFByteStreamBuffering() As UUID
'{6d66d782-1d4f-4db7-8c63-cb8c77f1ef5e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6D66D782, CInt(&H1D4F), CInt(&H4DB7), &H8C, &H63, &HCB, &H8C, &H77, &HF1, &HEF, &H5E)
IID_IMFByteStreamBuffering = iid
End Function
Public Function IID_IMFByteStreamCacheControl() As UUID
'{F5042EA4-7A96-4a75-AA7B-2BE1EF7F88D5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF5042EA4, CInt(&H7A96), CInt(&H4A75), &HAA, &H7B, &H2B, &HE1, &HEF, &H7F, &H88, &HD5)
IID_IMFByteStreamCacheControl = iid
End Function
Public Function IID_IMFByteStreamTimeSeek() As UUID
'{64976BFA-FB61-4041-9069-8C9A5F659BEB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H64976BFA, CInt(&HFB61), CInt(&H4041), &H90, &H69, &H8C, &H9A, &H5F, &H65, &H9B, &HEB)
IID_IMFByteStreamTimeSeek = iid
End Function
Public Function IID_IMFByteStreamCacheControl2() As UUID
'{71CE469C-F34B-49EA-A56B-2D2A10E51149}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H71CE469C, CInt(&HF34B), CInt(&H49EA), &HA5, &H6B, &H2D, &H2A, &H10, &HE5, &H11, &H49)
IID_IMFByteStreamCacheControl2 = iid
End Function
Public Function IID_IMFNetCredential() As UUID
'{5b87ef6a-7ed8-434f-ba0e-184fac1628d1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5B87EF6A, CInt(&H7ED8), CInt(&H434F), &HBA, &HE, &H18, &H4F, &HAC, &H16, &H28, &HD1)
IID_IMFNetCredential = iid
End Function
Public Function IID_IMFNetCredentialManager() As UUID
'{5b87ef6b-7ed8-434f-ba0e-184fac1628d1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5B87EF6B, CInt(&H7ED8), CInt(&H434F), &HBA, &HE, &H18, &H4F, &HAC, &H16, &H28, &HD1)
IID_IMFNetCredentialManager = iid
End Function
Public Function IID_IMFNetCredentialCache() As UUID
'{5b87ef6c-7ed8-434f-ba0e-184fac1628d1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5B87EF6C, CInt(&H7ED8), CInt(&H434F), &HBA, &HE, &H18, &H4F, &HAC, &H16, &H28, &HD1)
IID_IMFNetCredentialCache = iid
End Function
Public Function IID_IMFSSLCertificateManager() As UUID
'{61f7d887-1230-4a8b-aeba-8ad434d1a64d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H61F7D887, CInt(&H1230), CInt(&H4A8B), &HAE, &HBA, &H8A, &HD4, &H34, &HD1, &HA6, &H4D)
IID_IMFSSLCertificateManager = iid
End Function
Public Function IID_IMFNetResourceFilter() As UUID
'{091878a3-bf11-4a5c-bc9f-33995b06ef2d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H91878A3, CInt(&HBF11), CInt(&H4A5C), &HBC, &H9F, &H33, &H99, &H5B, &H6, &HEF, &H2D)
IID_IMFNetResourceFilter = iid
End Function
Public Function IID_IMFSourceOpenMonitor() As UUID
'{059054B3-027C-494C-A27D-9113291CF87F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H59054B3, CInt(&H27C), CInt(&H494C), &HA2, &H7D, &H91, &H13, &H29, &H1C, &HF8, &H7F)
IID_IMFSourceOpenMonitor = iid
End Function
Public Function IID_IMFNetProxyLocator() As UUID
'{e9cd0383-a268-4bb4-82de-658d53574d41}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE9CD0383, CInt(&HA268), CInt(&H4BB4), &H82, &HDE, &H65, &H8D, &H53, &H57, &H4D, &H41)
IID_IMFNetProxyLocator = iid
End Function
Public Function IID_IMFNetProxyLocatorFactory() As UUID
'{e9cd0384-a268-4bb4-82de-658d53574d41}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE9CD0384, CInt(&HA268), CInt(&H4BB4), &H82, &HDE, &H65, &H8D, &H53, &H57, &H4D, &H41)
IID_IMFNetProxyLocatorFactory = iid
End Function
Public Function IID_IMFSaveJob() As UUID
'{e9931663-80bf-4c6e-98af-5dcf58747d1f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE9931663, CInt(&H80BF), CInt(&H4C6E), &H98, &HAF, &H5D, &HCF, &H58, &H74, &H7D, &H1F)
IID_IMFSaveJob = iid
End Function
Public Function IID_IMFNetSchemeHandlerConfig() As UUID
'{7BE19E73-C9BF-468a-AC5A-A5E8653BEC87}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7BE19E73, CInt(&HC9BF), CInt(&H468A), &HAC, &H5A, &HA5, &HE8, &H65, &H3B, &HEC, &H87)
IID_IMFNetSchemeHandlerConfig = iid
End Function
Public Function IID_IMFSchemeHandler() As UUID
'{6D4C7B74-52A0-4bb7-B0DB-55F29F47A668}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6D4C7B74, CInt(&H52A0), CInt(&H4BB7), &HB0, &HDB, &H55, &HF2, &H9F, &H47, &HA6, &H68)
IID_IMFSchemeHandler = iid
End Function
Public Function IID_IMFByteStreamHandler() As UUID
'{BB420AA4-765B-4a1f-91FE-D6A8A143924C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBB420AA4, CInt(&H765B), CInt(&H4A1F), &H91, &HFE, &HD6, &HA8, &HA1, &H43, &H92, &H4C)
IID_IMFByteStreamHandler = iid
End Function
Public Function IID_IMFTrustedInput() As UUID
'{542612C4-A1B8-4632-B521-DE11EA64A0B0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H542612C4, CInt(&HA1B8), CInt(&H4632), &HB5, &H21, &HDE, &H11, &HEA, &H64, &HA0, &HB0)
IID_IMFTrustedInput = iid
End Function
Public Function IID_IMFInputTrustAuthority() As UUID
'{D19F8E98-B126-4446-890C-5DCB7AD71453}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD19F8E98, CInt(&HB126), CInt(&H4446), &H89, &HC, &H5D, &HCB, &H7A, &HD7, &H14, &H53)
IID_IMFInputTrustAuthority = iid
End Function
Public Function IID_IMFTrustedOutput() As UUID
'{D19F8E95-B126-4446-890C-5DCB7AD71453}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD19F8E95, CInt(&HB126), CInt(&H4446), &H89, &HC, &H5D, &HCB, &H7A, &HD7, &H14, &H53)
IID_IMFTrustedOutput = iid
End Function
Public Function IID_IMFOutputTrustAuthority() As UUID
'{D19F8E94-B126-4446-890C-5DCB7AD71453}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD19F8E94, CInt(&HB126), CInt(&H4446), &H89, &HC, &H5D, &HCB, &H7A, &HD7, &H14, &H53)
IID_IMFOutputTrustAuthority = iid
End Function
Public Function IID_IMFOutputPolicy() As UUID
'{7F00F10A-DAED-41AF-AB26-5FDFA4DFBA3C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7F00F10A, CInt(&HDAED), CInt(&H41AF), &HAB, &H26, &H5F, &HDF, &HA4, &HDF, &HBA, &H3C)
IID_IMFOutputPolicy = iid
End Function
Public Function IID_IMFOutputSchema() As UUID
'{7BE0FC5B-ABD9-44FB-A5C8-F50136E71599}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7BE0FC5B, CInt(&HABD9), CInt(&H44FB), &HA5, &HC8, &HF5, &H1, &H36, &HE7, &H15, &H99)
IID_IMFOutputSchema = iid
End Function
Public Function IID_IMFSecureChannel() As UUID
'{d0ae555d-3b12-4d97-b060-0990bc5aeb67}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD0AE555D, CInt(&H3B12), CInt(&H4D97), &HB0, &H60, &H9, &H90, &HBC, &H5A, &HEB, &H67)
IID_IMFSecureChannel = iid
End Function
Public Function IID_IMFSampleProtection() As UUID
'{8e36395f-c7b9-43c4-a54d-512b4af63c95}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8E36395F, CInt(&HC7B9), CInt(&H43C4), &HA5, &H4D, &H51, &H2B, &H4A, &HF6, &H3C, &H95)
IID_IMFSampleProtection = iid
End Function
Public Function IID_IMFMediaSinkPreroll() As UUID
'{5dfd4b2a-7674-4110-a4e6-8a68fd5f3688}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5DFD4B2A, CInt(&H7674), CInt(&H4110), &HA4, &HE6, &H8A, &H68, &HFD, &H5F, &H36, &H88)
IID_IMFMediaSinkPreroll = iid
End Function
Public Function IID_IMFFinalizableMediaSink() As UUID
'{EAECB74A-9A50-42ce-9541-6A7F57AA4AD7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEAECB74A, CInt(&H9A50), CInt(&H42CE), &H95, &H41, &H6A, &H7F, &H57, &HAA, &H4A, &HD7)
IID_IMFFinalizableMediaSink = iid
End Function
Public Function IID_IMFStreamingSinkConfig() As UUID
'{9db7aa41-3cc5-40d4-8509-555804ad34cc}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9DB7AA41, CInt(&H3CC5), CInt(&H40D4), &H85, &H9, &H55, &H58, &H4, &HAD, &H34, &HCC)
IID_IMFStreamingSinkConfig = iid
End Function
Public Function IID_IMFRemoteProxy() As UUID
'{994e23ad-1cc2-493c-b9fa-46f1cb040fa4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H994E23AD, CInt(&H1CC2), CInt(&H493C), &HB9, &HFA, &H46, &HF1, &HCB, &H4, &HF, &HA4)
IID_IMFRemoteProxy = iid
End Function
Public Function IID_IMFObjectReferenceStream() As UUID
'{09EF5BE3-C8A7-469e-8B70-73BF25BB193F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9EF5BE3, CInt(&HC8A7), CInt(&H469E), &H8B, &H70, &H73, &HBF, &H25, &HBB, &H19, &H3F)
IID_IMFObjectReferenceStream = iid
End Function
Public Function IID_IMFPMPHost() As UUID
'{F70CA1A9-FDC7-4782-B994-ADFFB1C98606}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF70CA1A9, CInt(&HFDC7), CInt(&H4782), &HB9, &H94, &HAD, &HFF, &HB1, &HC9, &H86, &H6)
IID_IMFPMPHost = iid
End Function
Public Function IID_IMFPMPClient() As UUID
'{6C4E655D-EAD8-4421-B6B9-54DCDBBDF820}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6C4E655D, CInt(&HEAD8), CInt(&H4421), &HB6, &HB9, &H54, &HDC, &HDB, &HBD, &HF8, &H20)
IID_IMFPMPClient = iid
End Function
Public Function IID_IMFPMPServer() As UUID
'{994e23af-1cc2-493c-b9fa-46f1cb040fa4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H994E23AF, CInt(&H1CC2), CInt(&H493C), &HB9, &HFA, &H46, &HF1, &HCB, &H4, &HF, &HA4)
IID_IMFPMPServer = iid
End Function
Public Function IID_IMFRemoteDesktopPlugin() As UUID
'{1cde6309-cae0-4940-907e-c1ec9c3d1d4a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1CDE6309, CInt(&HCAE0), CInt(&H4940), &H90, &H7E, &HC1, &HEC, &H9C, &H3D, &H1D, &H4A)
IID_IMFRemoteDesktopPlugin = iid
End Function
Public Function IID_IMFSAMIStyle() As UUID
'{A7E025DD-5303-4a62-89D6-E747E1EFAC73}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA7E025DD, CInt(&H5303), CInt(&H4A62), &H89, &HD6, &HE7, &H47, &HE1, &HEF, &HAC, &H73)
IID_IMFSAMIStyle = iid
End Function
Public Function IID_IMFTranscodeProfile() As UUID
'{4ADFDBA3-7AB0-4953-A62B-461E7FF3DA1E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4ADFDBA3, CInt(&H7AB0), CInt(&H4953), &HA6, &H2B, &H46, &H1E, &H7F, &HF3, &HDA, &H1E)
IID_IMFTranscodeProfile = iid
End Function
Public Function IID_IMFTranscodeSinkInfoProvider() As UUID
'{8CFFCD2E-5A03-4a3a-AFF7-EDCD107C620E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CFFCD2E, CInt(&H5A03), CInt(&H4A3A), &HAF, &HF7, &HED, &HCD, &H10, &H7C, &H62, &HE)
IID_IMFTranscodeSinkInfoProvider = iid
End Function
Public Function IID_IMFFieldOfUseMFTUnlock() As UUID
'{508E71D3-EC66-4fc3-8775-B4B9ED6BA847}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H508E71D3, CInt(&HEC66), CInt(&H4FC3), &H87, &H75, &HB4, &HB9, &HED, &H6B, &HA8, &H47)
IID_IMFFieldOfUseMFTUnlock = iid
End Function
Public Function IID_IMFLocalMFTRegistration() As UUID
'{149c4d73-b4be-4f8d-8b87-079e926b6add}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H149C4D73, CInt(&HB4BE), CInt(&H4F8D), &H8B, &H87, &H7, &H9E, &H92, &H6B, &H6A, &HDD)
IID_IMFLocalMFTRegistration = iid
End Function
Public Function IID_IMFPMPHostApp() As UUID
'{84d2054a-3aa1-4728-a3b0-440a418cf49c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H84D2054A, CInt(&H3AA1), CInt(&H4728), &HA3, &HB0, &H44, &HA, &H41, &H8C, &HF4, &H9C)
IID_IMFPMPHostApp = iid
End Function
Public Function IID_IMFPMPClientApp() As UUID
'{c004f646-be2c-48f3-93a2-a0983eba1108}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC004F646, CInt(&HBE2C), CInt(&H48F3), &H93, &HA2, &HA0, &H98, &H3E, &HBA, &H11, &H8)
IID_IMFPMPClientApp = iid
End Function
Public Function IID_IMFMediaStreamSourceSampleRequest() As UUID
'{380b9af9-a85b-4e78-a2af-ea5ce645c6b4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H380B9AF9, CInt(&HA85B), CInt(&H4E78), &HA2, &HAF, &HEA, &H5C, &HE6, &H45, &HC6, &HB4)
IID_IMFMediaStreamSourceSampleRequest = iid
End Function
Public Function IID_IMFTrackedSample() As UUID
'{245BF8E9-0755-40f7-88A5-AE0F18D55E17}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H245BF8E9, CInt(&H755), CInt(&H40F7), &H88, &HA5, &HAE, &HF, &H18, &HD5, &H5E, &H17)
IID_IMFTrackedSample = iid
End Function
Public Function IID_IMFProtectedEnvironmentAccess() As UUID
'{ef5dc845-f0d9-4ec9-b00c-cb5183d38434}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEF5DC845, CInt(&HF0D9), CInt(&H4EC9), &HB0, &HC, &HCB, &H51, &H83, &HD3, &H84, &H34)
IID_IMFProtectedEnvironmentAccess = iid
End Function
Public Function IID_IMFSignedLibrary() As UUID
'{4a724bca-ff6a-4c07-8e0d-7a358421cf06}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4A724BCA, CInt(&HFF6A), CInt(&H4C07), &H8E, &HD, &H7A, &H35, &H84, &H21, &HCF, &H6)
IID_IMFSignedLibrary = iid
End Function
Public Function IID_IMFSystemId() As UUID
'{fff4af3a-1fc1-4ef9-a29b-d26c49e2f31a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFFF4AF3A, CInt(&H1FC1), CInt(&H4EF9), &HA2, &H9B, &HD2, &H6C, &H49, &HE2, &HF3, &H1A)
IID_IMFSystemId = iid
End Function
Public Function IID_IMFContentProtectionDevice() As UUID
'{E6257174-A060-4C9A-A088-3B1B471CAD28}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE6257174, CInt(&HA060), CInt(&H4C9A), &HA0, &H88, &H3B, &H1B, &H47, &H1C, &HAD, &H28)
IID_IMFContentProtectionDevice = iid
End Function
Public Function IID_IMFContentDecryptorContext() As UUID
'{7EC4B1BD-43FB-4763-85D2-64FCB5C5F4CB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7EC4B1BD, CInt(&H43FB), CInt(&H4763), &H85, &HD2, &H64, &HFC, &HB5, &HC5, &HF4, &HCB)
IID_IMFContentDecryptorContext = iid
End Function
Public Function IID_IMFVideoPositionMapper() As UUID
'{1F6A9F17-E70B-4e24-8AE4-0B2C3BA7A4AE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1F6A9F17, CInt(&HE70B), CInt(&H4E24), &H8A, &HE4, &HB, &H2C, &H3B, &HA7, &HA4, &HAE)
IID_IMFVideoPositionMapper = iid
End Function
Public Function IID_IMFVideoDeviceID() As UUID
'{A38D9567-5A9C-4f3c-B293-8EB415B279BA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA38D9567, CInt(&H5A9C), CInt(&H4F3C), &HB2, &H93, &H8E, &HB4, &H15, &HB2, &H79, &HBA)
IID_IMFVideoDeviceID = iid
End Function
Public Function IID_IMFVideoDisplayControl() As UUID
'{a490b1e4-ab84-4d31-a1b2-181e03b1077a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA490B1E4, CInt(&HAB84), CInt(&H4D31), &HA1, &HB2, &H18, &H1E, &H3, &HB1, &H7, &H7A)
IID_IMFVideoDisplayControl = iid
End Function
Public Function IID_IMFVideoPresenter() As UUID
'{29AFF080-182A-4a5d-AF3B-448F3A6346CB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H29AFF080, CInt(&H182A), CInt(&H4A5D), &HAF, &H3B, &H44, &H8F, &H3A, &H63, &H46, &HCB)
IID_IMFVideoPresenter = iid
End Function
Public Function IID_IMFDesiredSample() As UUID
'{56C294D0-753E-4260-8D61-A3D8820B1D54}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56C294D0, CInt(&H753E), CInt(&H4260), &H8D, &H61, &HA3, &HD8, &H82, &HB, &H1D, &H54)
IID_IMFDesiredSample = iid
End Function
Public Function IID_IMFVideoMixerControl() As UUID
'{A5C6C53F-C202-4aa5-9695-175BA8C508A5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA5C6C53F, CInt(&HC202), CInt(&H4AA5), &H96, &H95, &H17, &H5B, &HA8, &HC5, &H8, &HA5)
IID_IMFVideoMixerControl = iid
End Function
Public Function IID_IMFVideoMixerControl2() As UUID
'{8459616d-966e-4930-b658-54fa7e5a16d3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8459616D, CInt(&H966E), CInt(&H4930), &HB6, &H58, &H54, &HFA, &H7E, &H5A, &H16, &HD3)
IID_IMFVideoMixerControl2 = iid
End Function
Public Function IID_IMFVideoRenderer() As UUID
'{DFDFD197-A9CA-43d8-B341-6AF3503792CD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDFDFD197, CInt(&HA9CA), CInt(&H43D8), &HB3, &H41, &H6A, &HF3, &H50, &H37, &H92, &HCD)
IID_IMFVideoRenderer = iid
End Function
Public Function IID_IEVRFilterConfig() As UUID
'{83E91E85-82C1-4ea7-801D-85DC50B75086}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H83E91E85, CInt(&H82C1), CInt(&H4EA7), &H80, &H1D, &H85, &HDC, &H50, &HB7, &H50, &H86)
IID_IEVRFilterConfig = iid
End Function
Public Function IID_IEVRFilterConfigEx() As UUID
'{aea36028-796d-454f-beee-b48071e24304}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAEA36028, CInt(&H796D), CInt(&H454F), &HBE, &HEE, &HB4, &H80, &H71, &HE2, &H43, &H4)
IID_IEVRFilterConfigEx = iid
End Function
Public Function IID_IMFTopologyServiceLookup() As UUID
'{fa993889-4383-415a-a930-dd472a8cf6f7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA993889, CInt(&H4383), CInt(&H415A), &HA9, &H30, &HDD, &H47, &H2A, &H8C, &HF6, &HF7)
IID_IMFTopologyServiceLookup = iid
End Function
Public Function IID_IMFTopologyServiceLookupClient() As UUID
'{fa99388a-4383-415a-a930-dd472a8cf6f7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA99388A, CInt(&H4383), CInt(&H415A), &HA9, &H30, &HDD, &H47, &H2A, &H8C, &HF6, &HF7)
IID_IMFTopologyServiceLookupClient = iid
End Function
Public Function IID_IEVRTrustedVideoPlugin() As UUID
'{83A4CE40-7710-494b-A893-A472049AF630}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H83A4CE40, CInt(&H7710), CInt(&H494B), &HA8, &H93, &HA4, &H72, &H4, &H9A, &HF6, &H30)
IID_IEVRTrustedVideoPlugin = iid
End Function
Public Function IID_IMFPMediaPlayer() As UUID
'{A714590A-58AF-430a-85BF-44F5EC838D85}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA714590A, CInt(&H58AF), CInt(&H430A), &H85, &HBF, &H44, &HF5, &HEC, &H83, &H8D, &H85)
IID_IMFPMediaPlayer = iid
End Function
Public Function IID_IMFPMediaItem() As UUID
'{90EB3E6B-ECBF-45cc-B1DA-C6FE3EA70D57}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H90EB3E6B, CInt(&HECBF), CInt(&H45CC), &HB1, &HDA, &HC6, &HFE, &H3E, &HA7, &HD, &H57)
IID_IMFPMediaItem = iid
End Function
Public Function IID_IMFPMediaPlayerCallback() As UUID
'{766C8FFB-5FDB-4fea-A28D-B912996F51BD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H766C8FFB, CInt(&H5FDB), CInt(&H4FEA), &HA2, &H8D, &HB9, &H12, &H99, &H6F, &H51, &HBD)
IID_IMFPMediaPlayerCallback = iid
End Function
Public Function IID_IMFCaptureSource() As UUID
'{439a42a8-0d2c-4505-be83-f79b2a05d5c4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H439A42A8, CInt(&HD2C), CInt(&H4505), &HBE, &H83, &HF7, &H9B, &H2A, &H5, &HD5, &HC4)
IID_IMFCaptureSource = iid
End Function
Public Function IID_IMFCaptureEngine() As UUID
'{a6bba433-176b-48b2-b375-53aa03473207}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA6BBA433, CInt(&H176B), CInt(&H48B2), &HB3, &H75, &H53, &HAA, &H3, &H47, &H32, &H7)
IID_IMFCaptureEngine = iid
End Function
Public Function IID_IMFCaptureEngineClassFactory() As UUID
'{8f02d140-56fc-4302-a705-3a97c78be779}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8F02D140, CInt(&H56FC), CInt(&H4302), &HA7, &H5, &H3A, &H97, &HC7, &H8B, &HE7, &H79)
IID_IMFCaptureEngineClassFactory = iid
End Function
Public Function IID_IMFCaptureEngineOnSampleCallback2() As UUID
'{e37ceed7-340f-4514-9f4d-9c2ae026100b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE37CEED7, CInt(&H340F), CInt(&H4514), &H9F, &H4D, &H9C, &H2A, &HE0, &H26, &H10, &HB)
IID_IMFCaptureEngineOnSampleCallback2 = iid
End Function
Public Function IID_IMFCaptureSink2() As UUID
'{f9e4219e-6197-4b5e-b888-bee310ab2c59}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF9E4219E, CInt(&H6197), CInt(&H4B5E), &HB8, &H88, &HBE, &HE3, &H10, &HAB, &H2C, &H59)
IID_IMFCaptureSink2 = iid
End Function
Public Function IID_IMFCaptureRecordSink() As UUID
'{3323b55a-f92a-4fe2-8edc-e9bfc0634d77}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3323B55A, CInt(&HF92A), CInt(&H4FE2), &H8E, &HDC, &HE9, &HBF, &HC0, &H63, &H4D, &H77)
IID_IMFCaptureRecordSink = iid
End Function
Public Function IID_IMFCapturePreviewSink() As UUID
'{77346cfd-5b49-4d73-ace0-5b52a859f2e0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H77346CFD, CInt(&H5B49), CInt(&H4D73), &HAC, &HE0, &H5B, &H52, &HA8, &H59, &HF2, &HE0)
IID_IMFCapturePreviewSink = iid
End Function
Public Function IID_IMFCapturePhotoSink() As UUID
'{d2d43cc8-48bb-4aa7-95db-10c06977e777}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD2D43CC8, CInt(&H48BB), CInt(&H4AA7), &H95, &HDB, &H10, &HC0, &H69, &H77, &HE7, &H77)
IID_IMFCapturePhotoSink = iid
End Function
Public Function IID_IMFCaptureEngineOnEventCallback() As UUID
'{aeda51c0-9025-4983-9012-de597b88b089}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAEDA51C0, CInt(&H9025), CInt(&H4983), &H90, &H12, &HDE, &H59, &H7B, &H88, &HB0, &H89)
IID_IMFCaptureEngineOnEventCallback = iid
End Function
Public Function IID_IMFCaptureEngineOnSampleCallback() As UUID
'{52150b82-ab39-4467-980f-e48bf0822ecd}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H52150B82, CInt(&HAB39), CInt(&H4467), &H98, &HF, &HE4, &H8B, &HF0, &H82, &H2E, &HCD)
IID_IMFCaptureEngineOnSampleCallback = iid
End Function
Public Function IID_IMFCaptureSink() As UUID
'{72d6135b-35e9-412c-b926-fd5265f2a885}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H72D6135B, CInt(&H35E9), CInt(&H412C), &HB9, &H26, &HFD, &H52, &H65, &HF2, &HA8, &H85)
IID_IMFCaptureSink = iid
End Function
Public Function IID_IMFMediaError() As UUID
'{fc0e10d2-ab2a-4501-a951-06bb1075184c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFC0E10D2, CInt(&HAB2A), CInt(&H4501), &HA9, &H51, &H6, &HBB, &H10, &H75, &H18, &H4C)
IID_IMFMediaError = iid
End Function
Public Function IID_IMFMediaTimeRange() As UUID
'{db71a2fc-078a-414e-9df9-8c2531b0aa6c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDB71A2FC, CInt(&H78A), CInt(&H414E), &H9D, &HF9, &H8C, &H25, &H31, &HB0, &HAA, &H6C)
IID_IMFMediaTimeRange = iid
End Function
Public Function IID_IMFMediaEngineNotify() As UUID
'{fee7c112-e776-42b5-9bbf-0048524e2bd5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFEE7C112, CInt(&HE776), CInt(&H42B5), &H9B, &HBF, &H0, &H48, &H52, &H4E, &H2B, &HD5)
IID_IMFMediaEngineNotify = iid
End Function
Public Function IID_IMFMediaEngineSrcElements() As UUID
'{7a5e5354-b114-4c72-b991-3131d75032ea}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7A5E5354, CInt(&HB114), CInt(&H4C72), &HB9, &H91, &H31, &H31, &HD7, &H50, &H32, &HEA)
IID_IMFMediaEngineSrcElements = iid
End Function
Public Function IID_IMFMediaEngine() As UUID
'{98a1b0bb-03eb-4935-ae7c-93c1fa0e1c93}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H98A1B0BB, CInt(&H3EB), CInt(&H4935), &HAE, &H7C, &H93, &HC1, &HFA, &HE, &H1C, &H93)
IID_IMFMediaEngine = iid
End Function
Public Function IID_IMFMediaEngineEx() As UUID
'{83015ead-b1e6-40d0-a98a-37145ffe1ad1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H83015EAD, CInt(&HB1E6), CInt(&H40D0), &HA9, &H8A, &H37, &H14, &H5F, &HFE, &H1A, &HD1)
IID_IMFMediaEngineEx = iid
End Function
Public Function IID_IMFMediaEngineAudioEndpointId() As UUID
'{7a3bac98-0e76-49fb-8c20-8a86fd98eaf2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7A3BAC98, CInt(&HE76), CInt(&H49FB), &H8C, &H20, &H8A, &H86, &HFD, &H98, &HEA, &HF2)
IID_IMFMediaEngineAudioEndpointId = iid
End Function
Public Function IID_IMFMediaEngineExtension() As UUID
'{2f69d622-20b5-41e9-afdf-89ced1dda04e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2F69D622, CInt(&H20B5), CInt(&H41E9), &HAF, &HDF, &H89, &HCE, &HD1, &HDD, &HA0, &H4E)
IID_IMFMediaEngineExtension = iid
End Function
Public Function IID_IMFMediaEngineProtectedContent() As UUID
'{9f8021e8-9c8c-487e-bb5c-79aa4779938c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9F8021E8, CInt(&H9C8C), CInt(&H487E), &HBB, &H5C, &H79, &HAA, &H47, &H79, &H93, &H8C)
IID_IMFMediaEngineProtectedContent = iid
End Function
Public Function IID_IAudioSourceProvider() As UUID
'{EBBAF249-AFC2-4582-91C6-B60DF2E84954}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEBBAF249, CInt(&HAFC2), CInt(&H4582), &H91, &HC6, &HB6, &HD, &HF2, &HE8, &H49, &H54)
IID_IAudioSourceProvider = iid
End Function
Public Function IID_IMFMediaEngineWebSupport() As UUID
'{ba2743a1-07e0-48ef-84b6-9a2ed023ca6c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBA2743A1, CInt(&H7E0), CInt(&H48EF), &H84, &HB6, &H9A, &H2E, &HD0, &H23, &HCA, &H6C)
IID_IMFMediaEngineWebSupport = iid
End Function
Public Function IID_IMFMediaSourceExtensionNotify() As UUID
'{a7901327-05dd-4469-a7b7-0e01979e361d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA7901327, CInt(&H5DD), CInt(&H4469), &HA7, &HB7, &HE, &H1, &H97, &H9E, &H36, &H1D)
IID_IMFMediaSourceExtensionNotify = iid
End Function
Public Function IID_IMFBufferListNotify() As UUID
'{24cd47f7-81d8-4785-adb2-af697a963cd2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H24CD47F7, CInt(&H81D8), CInt(&H4785), &HAD, &HB2, &HAF, &H69, &H7A, &H96, &H3C, &HD2)
IID_IMFBufferListNotify = iid
End Function
Public Function IID_IMFSourceBufferNotify() As UUID
'{87e47623-2ceb-45d6-9b88-d8520c4dcbbc}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H87E47623, CInt(&H2CEB), CInt(&H45D6), &H9B, &H88, &HD8, &H52, &HC, &H4D, &HCB, &HBC)
IID_IMFSourceBufferNotify = iid
End Function
Public Function IID_IMFSourceBuffer() As UUID
'{e2cd3a4b-af25-4d3d-9110-da0e6f8ee877}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE2CD3A4B, CInt(&HAF25), CInt(&H4D3D), &H91, &H10, &HDA, &HE, &H6F, &H8E, &HE8, &H77)
IID_IMFSourceBuffer = iid
End Function
Public Function IID_IMFSourceBufferAppendMode() As UUID
'{19666fb4-babe-4c55-bc03-0a074da37e2a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H19666FB4, CInt(&HBABE), CInt(&H4C55), &HBC, &H3, &HA, &H7, &H4D, &HA3, &H7E, &H2A)
IID_IMFSourceBufferAppendMode = iid
End Function
Public Function IID_IMFSourceBufferList() As UUID
'{249981f8-8325-41f3-b80c-3b9e3aad0cbe}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H249981F8, CInt(&H8325), CInt(&H41F3), &HB8, &HC, &H3B, &H9E, &H3A, &HAD, &HC, &HBE)
IID_IMFSourceBufferList = iid
End Function
Public Function IID_IMFMediaSourceExtension() As UUID
'{e467b94e-a713-4562-a802-816a42e9008a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE467B94E, CInt(&HA713), CInt(&H4562), &HA8, &H2, &H81, &H6A, &H42, &HE9, &H0, &H8A)
IID_IMFMediaSourceExtension = iid
End Function
Public Function IID_IMFMediaSourceExtensionLiveSeekableRange() As UUID
'{5D1ABFD6-450A-4D92-9EFC-D6B6CBC1F4DA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5D1ABFD6, CInt(&H450A), CInt(&H4D92), &H9E, &HFC, &HD6, &HB6, &HCB, &HC1, &HF4, &HDA)
IID_IMFMediaSourceExtensionLiveSeekableRange = iid
End Function
Public Function IID_IMFMediaEngineEME() As UUID
'{50dc93e4-ba4f-4275-ae66-83e836e57469}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H50DC93E4, CInt(&HBA4F), CInt(&H4275), &HAE, &H66, &H83, &HE8, &H36, &HE5, &H74, &H69)
IID_IMFMediaEngineEME = iid
End Function
Public Function IID_IMFMediaEngineSrcElementsEx() As UUID
'{654a6bb3-e1a3-424a-9908-53a43a0dfda0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H654A6BB3, CInt(&HE1A3), CInt(&H424A), &H99, &H8, &H53, &HA4, &H3A, &HD, &HFD, &HA0)
IID_IMFMediaEngineSrcElementsEx = iid
End Function
Public Function IID_IMFMediaEngineNeedKeyNotify() As UUID
'{46a30204-a696-4b18-8804-246b8f031bb1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H46A30204, CInt(&HA696), CInt(&H4B18), &H88, &H4, &H24, &H6B, &H8F, &H3, &H1B, &HB1)
IID_IMFMediaEngineNeedKeyNotify = iid
End Function
Public Function IID_IMFMediaKeys() As UUID
'{5cb31c05-61ff-418f-afda-caaf41421a38}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5CB31C05, CInt(&H61FF), CInt(&H418F), &HAF, &HDA, &HCA, &HAF, &H41, &H42, &H1A, &H38)
IID_IMFMediaKeys = iid
End Function
Public Function IID_IMFMediaKeySession() As UUID
'{24fa67d5-d1d0-4dc5-995c-c0efdc191fb5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H24FA67D5, CInt(&HD1D0), CInt(&H4DC5), &H99, &H5C, &HC0, &HEF, &HDC, &H19, &H1F, &HB5)
IID_IMFMediaKeySession = iid
End Function
Public Function IID_IMFMediaKeySessionNotify() As UUID
'{6a0083f9-8947-4c1d-9ce0-cdee22b23135}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6A0083F9, CInt(&H8947), CInt(&H4C1D), &H9C, &HE0, &HCD, &HEE, &H22, &HB2, &H31, &H35)
IID_IMFMediaKeySessionNotify = iid
End Function
Public Function IID_IMFCdmSuspendNotify() As UUID
'{7a5645d2-43bd-47fd-87b7-dcd24cc7d692}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7A5645D2, CInt(&H43BD), CInt(&H47FD), &H87, &HB7, &HDC, &HD2, &H4C, &HC7, &HD6, &H92)
IID_IMFCdmSuspendNotify = iid
End Function
Public Function IID_IMFHDCPStatus() As UUID
'{DE400F54-5BF1-40CF-8964-0BEA136B1E3D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDE400F54, CInt(&H5BF1), CInt(&H40CF), &H89, &H64, &HB, &HEA, &H13, &H6B, &H1E, &H3D)
IID_IMFHDCPStatus = iid
End Function
Public Function IID_IMFMediaEngineOPMInfo() As UUID
'{765763e6-6c01-4b01-bb0f-b829f60ed28c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H765763E6, CInt(&H6C01), CInt(&H4B01), &HBB, &HF, &HB8, &H29, &HF6, &HE, &HD2, &H8C)
IID_IMFMediaEngineOPMInfo = iid
End Function
Public Function IID_IMFMediaEngineClassFactory() As UUID
'{4D645ACE-26AA-4688-9BE1-DF3516990B93}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4D645ACE, CInt(&H26AA), CInt(&H4688), &H9B, &HE1, &HDF, &H35, &H16, &H99, &HB, &H93)
IID_IMFMediaEngineClassFactory = iid
End Function
Public Function IID_IMFMediaEngineClassFactoryEx() As UUID
'{c56156c6-ea5b-48a5-9df8-fbe035d0929e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC56156C6, CInt(&HEA5B), CInt(&H48A5), &H9D, &HF8, &HFB, &HE0, &H35, &HD0, &H92, &H9E)
IID_IMFMediaEngineClassFactoryEx = iid
End Function
Public Function IID_IMFMediaEngineClassFactory2() As UUID
'{09083cef-867f-4bf6-8776-dee3a7b42fca}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9083CEF, CInt(&H867F), CInt(&H4BF6), &H87, &H76, &HDE, &HE3, &HA7, &HB4, &H2F, &HCA)
IID_IMFMediaEngineClassFactory2 = iid
End Function
Public Function IID_IMFExtendedDRMTypeSupport() As UUID
'{332EC562-3758-468D-A784-E38F23552128}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H332EC562, CInt(&H3758), CInt(&H468D), &HA7, &H84, &HE3, &H8F, &H23, &H55, &H21, &H28)
IID_IMFExtendedDRMTypeSupport = iid
End Function
Public Function IID_IMFMediaEngineSupportsSourceTransfer() As UUID
'{a724b056-1b2e-4642-a6f3-db9420c52908}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA724B056, CInt(&H1B2E), CInt(&H4642), &HA6, &HF3, &HDB, &H94, &H20, &HC5, &H29, &H8)
IID_IMFMediaEngineSupportsSourceTransfer = iid
End Function
Public Function IID_IMFMediaEngineTransferSource() As UUID
'{24230452-fe54-40cc-94f3-fcc394c340d6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H24230452, CInt(&HFE54), CInt(&H40CC), &H94, &HF3, &HFC, &HC3, &H94, &HC3, &H40, &HD6)
IID_IMFMediaEngineTransferSource = iid
End Function
Public Function IID_IMFTimedText() As UUID
'{1f2a94c9-a3df-430d-9d0f-acd85ddc29af}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1F2A94C9, CInt(&HA3DF), CInt(&H430D), &H9D, &HF, &HAC, &HD8, &H5D, &HDC, &H29, &HAF)
IID_IMFTimedText = iid
End Function
Public Function IID_IMFTimedTextNotify() As UUID
'{df6b87b6-ce12-45db-aba7-432fe054e57d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDF6B87B6, CInt(&HCE12), CInt(&H45DB), &HAB, &HA7, &H43, &H2F, &HE0, &H54, &HE5, &H7D)
IID_IMFTimedTextNotify = iid
End Function
Public Function IID_IMFTimedTextTrack() As UUID
'{8822c32d-654e-4233-bf21-d7f2e67d30d4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8822C32D, CInt(&H654E), CInt(&H4233), &HBF, &H21, &HD7, &HF2, &HE6, &H7D, &H30, &HD4)
IID_IMFTimedTextTrack = iid
End Function
Public Function IID_IMFTimedTextTrackList() As UUID
'{23ff334c-442c-445f-bccc-edc438aa11e2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H23FF334C, CInt(&H442C), CInt(&H445F), &HBC, &HCC, &HED, &HC4, &H38, &HAA, &H11, &HE2)
IID_IMFTimedTextTrackList = iid
End Function
Public Function IID_IMFTimedTextCue() As UUID
'{1e560447-9a2b-43e1-a94c-b0aaabfbfbc9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1E560447, CInt(&H9A2B), CInt(&H43E1), &HA9, &H4C, &HB0, &HAA, &HAB, &HFB, &HFB, &HC9)
IID_IMFTimedTextCue = iid
End Function
Public Function IID_IMFTimedTextFormattedText() As UUID
'{e13af3c1-4d47-4354-b1f5-e83ae0ecae60}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE13AF3C1, CInt(&H4D47), CInt(&H4354), &HB1, &HF5, &HE8, &H3A, &HE0, &HEC, &HAE, &H60)
IID_IMFTimedTextFormattedText = iid
End Function
Public Function IID_IMFTimedTextStyle() As UUID
'{09b2455d-b834-4f01-a347-9052e21c450e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9B2455D, CInt(&HB834), CInt(&H4F01), &HA3, &H47, &H90, &H52, &HE2, &H1C, &H45, &HE)
IID_IMFTimedTextStyle = iid
End Function
Public Function IID_IMFTimedTextRegion() As UUID
'{c8d22afc-bc47-4bdf-9b04-787e49ce3f58}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC8D22AFC, CInt(&HBC47), CInt(&H4BDF), &H9B, &H4, &H78, &H7E, &H49, &HCE, &H3F, &H58)
IID_IMFTimedTextRegion = iid
End Function
Public Function IID_IMFTimedTextBinary() As UUID
'{4ae3a412-0545-43c4-bf6f-6b97a5c6c432}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4AE3A412, CInt(&H545), CInt(&H43C4), &HBF, &H6F, &H6B, &H97, &HA5, &HC6, &HC4, &H32)
IID_IMFTimedTextBinary = iid
End Function
Public Function IID_IMFTimedTextCueList() As UUID
'{ad128745-211b-40a0-9981-fe65f166d0fd}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAD128745, CInt(&H211B), CInt(&H40A0), &H99, &H81, &HFE, &H65, &HF1, &H66, &HD0, &HFD)
IID_IMFTimedTextCueList = iid
End Function
Public Function IID_IMFTimedTextRuby() As UUID
'{76c6a6f5-4955-4de5-b27b-14b734cc14b4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H76C6A6F5, CInt(&H4955), CInt(&H4DE5), &HB2, &H7B, &H14, &HB7, &H34, &HCC, &H14, &HB4)
IID_IMFTimedTextRuby = iid
End Function
Public Function IID_IMFTimedTextBouten() As UUID
'{3c5f3e8a-90c0-464e-8136-898d2975f847}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3C5F3E8A, CInt(&H90C0), CInt(&H464E), &H81, &H36, &H89, &H8D, &H29, &H75, &HF8, &H47)
IID_IMFTimedTextBouten = iid
End Function
Public Function IID_IMFTimedTextStyle2() As UUID
'{db639199-c809-4c89-bfca-d0bbb9729d6e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDB639199, CInt(&HC809), CInt(&H4C89), &HBF, &HCA, &HD0, &HBB, &HB9, &H72, &H9D, &H6E)
IID_IMFTimedTextStyle2 = iid
End Function
Public Function IID_IMFMediaEngineEMENotify() As UUID
'{9e184d15-cdb7-4f86-b49e-566689f4a601}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9E184D15, CInt(&HCDB7), CInt(&H4F86), &HB4, &H9E, &H56, &H66, &H89, &HF4, &HA6, &H1)
IID_IMFMediaEngineEMENotify = iid
End Function
Public Function IID_IMFMediaKeySessionNotify2() As UUID
'{c3a9e92a-da88-46b0-a110-6cf953026cb9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC3A9E92A, CInt(&HDA88), CInt(&H46B0), &HA1, &H10, &H6C, &HF9, &H53, &H2, &H6C, &HB9)
IID_IMFMediaKeySessionNotify2 = iid
End Function
Public Function IID_IMFMediaKeySystemAccess() As UUID
'{aec63fda-7a97-4944-b35c-6c6df8085cc3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAEC63FDA, CInt(&H7A97), CInt(&H4944), &HB3, &H5C, &H6C, &H6D, &HF8, &H8, &H5C, &HC3)
IID_IMFMediaKeySystemAccess = iid
End Function
Public Function IID_IMFMediaEngineClassFactory3() As UUID
'{3787614f-65f7-4003-b673-ead8293a0e60}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3787614F, CInt(&H65F7), CInt(&H4003), &HB6, &H73, &HEA, &HD8, &H29, &H3A, &HE, &H60)
IID_IMFMediaEngineClassFactory3 = iid
End Function
Public Function IID_IMFMediaKeys2() As UUID
'{45892507-ad66-4de2-83a2-acbb13cd8d43}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H45892507, CInt(&HAD66), CInt(&H4DE2), &H83, &HA2, &HAC, &HBB, &H13, &HCD, &H8D, &H43)
IID_IMFMediaKeys2 = iid
End Function
Public Function IID_IMFMediaKeySession2() As UUID
'{e9707e05-6d55-4636-b185-3de21210bd75}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE9707E05, CInt(&H6D55), CInt(&H4636), &HB1, &H85, &H3D, &HE2, &H12, &H10, &HBD, &H75)
IID_IMFMediaKeySession2 = iid
End Function
Public Function IID_IMFMediaEngineClassFactory4() As UUID
'{fbe256c1-43cf-4a9b-8cb8-ce8632a34186}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFBE256C1, CInt(&H43CF), CInt(&H4A9B), &H8C, &HB8, &HCE, &H86, &H32, &HA3, &H41, &H86)
IID_IMFMediaEngineClassFactory4 = iid
End Function
Public Function IID_IMFContentDecryptionModuleSession() As UUID
'{4e233efd-1dd2-49e8-b577-d63eee4c0d33}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4E233EFD, CInt(&H1DD2), CInt(&H49E8), &HB5, &H77, &HD6, &H3E, &HEE, &H4C, &HD, &H33)
IID_IMFContentDecryptionModuleSession = iid
End Function
Public Function IID_IMFContentDecryptionModuleSessionCallbacks() As UUID
'{3f96ee40-ad81-4096-8470-59a4b770f89a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3F96EE40, CInt(&HAD81), CInt(&H4096), &H84, &H70, &H59, &HA4, &HB7, &H70, &HF8, &H9A)
IID_IMFContentDecryptionModuleSessionCallbacks = iid
End Function
Public Function IID_IMFContentDecryptionModule() As UUID
'{87be986c-10be-4943-bf48-4b54ce1983a2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H87BE986C, CInt(&H10BE), CInt(&H4943), &HBF, &H48, &H4B, &H54, &HCE, &H19, &H83, &HA2)
IID_IMFContentDecryptionModule = iid
End Function
Public Function IID_IMFContentDecryptionModuleAccess() As UUID
'{a853d1f4-e2a0-4303-9edc-f1a68ee43136}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA853D1F4, CInt(&HE2A0), CInt(&H4303), &H9E, &HDC, &HF1, &HA6, &H8E, &HE4, &H31, &H36)
IID_IMFContentDecryptionModuleAccess = iid
End Function
Public Function IID_IMFContentDecryptionModuleFactory() As UUID
'{7d5abf16-4cbb-4e08-b977-9ba59049943e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7D5ABF16, CInt(&H4CBB), CInt(&H4E08), &HB9, &H77, &H9B, &HA5, &H90, &H49, &H94, &H3E)
IID_IMFContentDecryptionModuleFactory = iid
End Function
Public Function IID_IMFDLNASinkInit() As UUID
'{0c012799-1b61-4c10-bda9-04445be5f561}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC012799, CInt(&H1B61), CInt(&H4C10), &HBD, &HA9, &H4, &H44, &H5B, &HE5, &HF5, &H61)
IID_IMFDLNASinkInit = iid
End Function
Public Function IID_IMFD3D12SynchronizationObjectCommands() As UUID
'{09D0F835-92FF-4E53-8EFA-40FAA551F233}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9D0F835, CInt(&H92FF), CInt(&H4E53), &H8E, &HFA, &H40, &HFA, &HA5, &H51, &HF2, &H33)
IID_IMFD3D12SynchronizationObjectCommands = iid
End Function
Public Function IID_IMFD3D12SynchronizationObject() As UUID
'{802302B0-82DE-45E1-B421-F19EE5BDAF23}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H802302B0, CInt(&H82DE), CInt(&H45E1), &HB4, &H21, &HF1, &H9E, &HE5, &HBD, &HAF, &H23)
IID_IMFD3D12SynchronizationObject = iid
End Function
Public Function IID_IAdvancedMediaCaptureInitializationSettings() As UUID
'{3DE21209-8BA6-4f2a-A577-2819B56FF14D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3DE21209, CInt(&H8BA6), CInt(&H4F2A), &HA5, &H77, &H28, &H19, &HB5, &H6F, &HF1, &H4D)
IID_IAdvancedMediaCaptureInitializationSettings = iid
End Function
Public Function IID_IAdvancedMediaCaptureSettings() As UUID
'{24E0485F-A33E-4aa1-B564-6019B1D14F65}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H24E0485F, CInt(&HA33E), CInt(&H4AA1), &HB5, &H64, &H60, &H19, &HB1, &HD1, &H4F, &H65)
IID_IAdvancedMediaCaptureSettings = iid
End Function
Public Function IID_IAdvancedMediaCapture() As UUID
'{D0751585-D216-4344-B5BF-463B68F977BB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD0751585, CInt(&HD216), CInt(&H4344), &HB5, &HBF, &H46, &H3B, &H68, &HF9, &H77, &HBB)
IID_IAdvancedMediaCapture = iid
End Function
Public Function IID_IMFSharingEngineClassFactory() As UUID
'{2BA61F92-8305-413B-9733-FAF15F259384}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2BA61F92, CInt(&H8305), CInt(&H413B), &H97, &H33, &HFA, &HF1, &H5F, &H25, &H93, &H84)
IID_IMFSharingEngineClassFactory = iid
End Function
Public Function IID_IMFMediaSharingEngine() As UUID
'{8D3CE1BF-2367-40E0-9EEE-40D377CC1B46}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8D3CE1BF, CInt(&H2367), CInt(&H40E0), &H9E, &HEE, &H40, &HD3, &H77, &HCC, &H1B, &H46)
IID_IMFMediaSharingEngine = iid
End Function
Public Function IID_IMFMediaSharingEngineClassFactory() As UUID
'{524D2BC4-B2B1-4FE5-8FAC-FA4E4512B4E0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H524D2BC4, CInt(&HB2B1), CInt(&H4FE5), &H8F, &HAC, &HFA, &H4E, &H45, &H12, &HB4, &HE0)
IID_IMFMediaSharingEngineClassFactory = iid
End Function
Public Function IID_IMFImageSharingEngine() As UUID
'{CFA0AE8E-7E1C-44D2-AE68-FC4C148A6354}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCFA0AE8E, CInt(&H7E1C), CInt(&H44D2), &HAE, &H68, &HFC, &H4C, &H14, &H8A, &H63, &H54)
IID_IMFImageSharingEngine = iid
End Function
Public Function IID_IMFImageSharingEngineClassFactory() As UUID
'{1FC55727-A7FB-4FC8-83AE-8AF024990AF1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1FC55727, CInt(&HA7FB), CInt(&H4FC8), &H83, &HAE, &H8A, &HF0, &H24, &H99, &HA, &HF1)
IID_IMFImageSharingEngineClassFactory = iid
End Function
Public Function IID_IPlayToControl() As UUID
'{607574EB-F4B6-45C1-B08C-CB715122901D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H607574EB, CInt(&HF4B6), CInt(&H45C1), &HB0, &H8C, &HCB, &H71, &H51, &H22, &H90, &H1D)
IID_IPlayToControl = iid
End Function
Public Function IID_IPlayToControlWithCapabilities() As UUID
'{AA9DD80F-C50A-4220-91C1-332287F82A34}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAA9DD80F, CInt(&HC50A), CInt(&H4220), &H91, &HC1, &H33, &H22, &H87, &HF8, &H2A, &H34)
IID_IPlayToControlWithCapabilities = iid
End Function
Public Function IID_IPlayToSourceClassFactory() As UUID
'{842B32A3-9B9B-4D1C-B3F3-49193248A554}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H842B32A3, CInt(&H9B9B), CInt(&H4D1C), &HB3, &HF3, &H49, &H19, &H32, &H48, &HA5, &H54)
IID_IPlayToSourceClassFactory = iid
End Function
Public Function IID_IMFSpatialAudioObjectBuffer() As UUID
'{d396ec8c-605e-4249-978d-72ad1c312872}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD396EC8C, CInt(&H605E), CInt(&H4249), &H97, &H8D, &H72, &HAD, &H1C, &H31, &H28, &H72)
IID_IMFSpatialAudioObjectBuffer = iid
End Function
Public Function IID_IMFSpatialAudioSample() As UUID
'{abf28a9B-3393-4290-ba79-5ffc46d986b2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HABF28A9B, CInt(&H3393), CInt(&H4290), &HBA, &H79, &H5F, &HFC, &H46, &HD9, &H86, &HB2)
IID_IMFSpatialAudioSample = iid
End Function
Public Function IID_IMFVirtualCamera() As UUID
'{1C08A864-EF6C-4C75-AF59-5F2D68DA9563}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1C08A864, CInt(&HEF6C), CInt(&H4C75), &HAF, &H59, &H5F, &H2D, &H68, &HDA, &H95, &H63)
IID_IMFVirtualCamera = iid
End Function
Public Function IID_IMFMuxStreamAttributesManager() As UUID
'{CE8BD576-E440-43B3-BE34-1E53F565F7E8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCE8BD576, CInt(&HE440), CInt(&H43B3), &HBE, &H34, &H1E, &H53, &HF5, &H65, &HF7, &HE8)
IID_IMFMuxStreamAttributesManager = iid
End Function
Public Function IID_IMFMuxStreamMediaTypeManager() As UUID
'{505A2C72-42F7-4690-AEAB-8F513D0FFDB8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H505A2C72, CInt(&H42F7), CInt(&H4690), &HAE, &HAB, &H8F, &H51, &H3D, &HF, &HFD, &HB8)
IID_IMFMuxStreamMediaTypeManager = iid
End Function
Public Function IID_IMFMuxStreamSampleManager() As UUID
'{74ABBC19-B1CC-4E41-BB8B-9D9B86A8F6CA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H74ABBC19, CInt(&HB1CC), CInt(&H4E41), &HBB, &H8B, &H9D, &H9B, &H86, &HA8, &HF6, &HCA)
IID_IMFMuxStreamSampleManager = iid
End Function
Public Function IID_IMFSecureBuffer() As UUID
'{C1209904-E584-4752-A2D6-7F21693F8B21}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC1209904, CInt(&HE584), CInt(&H4752), &HA2, &HD6, &H7F, &H21, &H69, &H3F, &H8B, &H21)
IID_IMFSecureBuffer = iid
End Function
Public Function IID_IMFNetCrossOriginSupport() As UUID
'{bc2b7d44-a72d-49d5-8376-1480dee58b22}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBC2B7D44, CInt(&HA72D), CInt(&H49D5), &H83, &H76, &H14, &H80, &HDE, &HE5, &H8B, &H22)
IID_IMFNetCrossOriginSupport = iid
End Function
Public Function IID_IMFHttpDownloadRequest() As UUID
'{F779FDDF-26E7-4270-8A8B-B983D1859DE0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF779FDDF, CInt(&H26E7), CInt(&H4270), &H8A, &H8B, &HB9, &H83, &HD1, &H85, &H9D, &HE0)
IID_IMFHttpDownloadRequest = iid
End Function
Public Function IID_IMFHttpDownloadSession() As UUID
'{71FA9A2C-53CE-4662-A132-1A7E8CBF62DB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H71FA9A2C, CInt(&H53CE), CInt(&H4662), &HA1, &H32, &H1A, &H7E, &H8C, &HBF, &H62, &HDB)
IID_IMFHttpDownloadSession = iid
End Function
Public Function IID_IMFHttpDownloadSessionProvider() As UUID
'{1B4CF4B9-3A16-4115-839D-03CC5C99DF01}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1B4CF4B9, CInt(&H3A16), CInt(&H4115), &H83, &H9D, &H3, &HCC, &H5C, &H99, &HDF, &H1)
IID_IMFHttpDownloadSessionProvider = iid
End Function
Public Function IID_IMFMediaSource2() As UUID
'{FBB03414-D13B-4786-8319-5AC51FC0A136}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFBB03414, CInt(&HD13B), CInt(&H4786), &H83, &H19, &H5A, &HC5, &H1F, &HC0, &HA1, &H36)
IID_IMFMediaSource2 = iid
End Function
Public Function IID_IMFMediaStream2() As UUID
'{C5BC37D6-75C7-46A1-A132-81B5F723C20F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC5BC37D6, CInt(&H75C7), CInt(&H46A1), &HA1, &H32, &H81, &HB5, &HF7, &H23, &HC2, &HF)
IID_IMFMediaStream2 = iid
End Function
Public Function IID_IMFSensorDevice() As UUID
'{FB9F48F2-2A18-4E28-9730-786F30F04DC4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFB9F48F2, CInt(&H2A18), CInt(&H4E28), &H97, &H30, &H78, &H6F, &H30, &HF0, &H4D, &HC4)
IID_IMFSensorDevice = iid
End Function
Public Function IID_IMFSensorGroup() As UUID
'{4110243A-9757-461F-89F1-F22345BCAB4E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4110243A, CInt(&H9757), CInt(&H461F), &H89, &HF1, &HF2, &H23, &H45, &HBC, &HAB, &H4E)
IID_IMFSensorGroup = iid
End Function
Public Function IID_IMFSensorStream() As UUID
'{E9A42171-C56E-498A-8B39-EDA5A070B7FC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE9A42171, CInt(&HC56E), CInt(&H498A), &H8B, &H39, &HED, &HA5, &HA0, &H70, &HB7, &HFC)
IID_IMFSensorStream = iid
End Function
Public Function IID_IMFSensorTransformFactory() As UUID
'{EED9C2EE-66B4-4F18-A697-AC7D3960215C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEED9C2EE, CInt(&H66B4), CInt(&H4F18), &HA6, &H97, &HAC, &H7D, &H39, &H60, &H21, &H5C)
IID_IMFSensorTransformFactory = iid
End Function
Public Function IID_IMFSensorProfile() As UUID
'{22F765D1-8DAB-4107-846D-56BAF72215E7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H22F765D1, CInt(&H8DAB), CInt(&H4107), &H84, &H6D, &H56, &HBA, &HF7, &H22, &H15, &HE7)
IID_IMFSensorProfile = iid
End Function
Public Function IID_IMFSensorProfileCollection() As UUID
'{C95EA55B-0187-48BE-9353-8D2507662351}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC95EA55B, CInt(&H187), CInt(&H48BE), &H93, &H53, &H8D, &H25, &H7, &H66, &H23, &H51)
IID_IMFSensorProfileCollection = iid
End Function
Public Function IID_IMFSensorProcessActivity() As UUID
'{39DC7F4A-B141-4719-813C-A7F46162A2B8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H39DC7F4A, CInt(&HB141), CInt(&H4719), &H81, &H3C, &HA7, &HF4, &H61, &H62, &HA2, &HB8)
IID_IMFSensorProcessActivity = iid
End Function
Public Function IID_IMFSensorActivityReport() As UUID
'{3E8C4BE1-A8C2-4528-90DE-2851BDE5FEAD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3E8C4BE1, CInt(&HA8C2), CInt(&H4528), &H90, &HDE, &H28, &H51, &HBD, &HE5, &HFE, &HAD)
IID_IMFSensorActivityReport = iid
End Function
Public Function IID_IMFSensorActivitiesReport() As UUID
'{683F7A5E-4A19-43CD-B1A9-DBF4AB3F7777}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H683F7A5E, CInt(&H4A19), CInt(&H43CD), &HB1, &HA9, &HDB, &HF4, &HAB, &H3F, &H77, &H77)
IID_IMFSensorActivitiesReport = iid
End Function
Public Function IID_IMFSensorActivitiesReportCallback() As UUID
'{DE5072EE-DBE3-46DC-8A87-B6F631194751}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDE5072EE, CInt(&HDBE3), CInt(&H46DC), &H8A, &H87, &HB6, &HF6, &H31, &H19, &H47, &H51)
IID_IMFSensorActivitiesReportCallback = iid
End Function
Public Function IID_IMFSensorActivityMonitor() As UUID
'{D0CEF145-B3F4-4340-A2E5-7A5080CA05CB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD0CEF145, CInt(&HB3F4), CInt(&H4340), &HA2, &HE5, &H7A, &H50, &H80, &HCA, &H5, &HCB)
IID_IMFSensorActivityMonitor = iid
End Function
Public Function IID_IMFExtendedCameraIntrinsicModel() As UUID
'{5C595E64-4630-4231-855A-12842F733245}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5C595E64, CInt(&H4630), CInt(&H4231), &H85, &H5A, &H12, &H84, &H2F, &H73, &H32, &H45)
IID_IMFExtendedCameraIntrinsicModel = iid
End Function
Public Function IID_IMFExtendedCameraIntrinsicsDistortionModel6KT() As UUID
'{74C2653B-5F55-4EB1-9F0F-18B8F68B7D3D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H74C2653B, CInt(&H5F55), CInt(&H4EB1), &H9F, &HF, &H18, &HB8, &HF6, &H8B, &H7D, &H3D)
IID_IMFExtendedCameraIntrinsicsDistortionModel6KT = iid
End Function
Public Function IID_IMFExtendedCameraIntrinsicsDistortionModelArcTan() As UUID
'{812D5F95-B572-45DC-BAFC-AE24199DDDA8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H812D5F95, CInt(&HB572), CInt(&H45DC), &HBA, &HFC, &HAE, &H24, &H19, &H9D, &HDD, &HA8)
IID_IMFExtendedCameraIntrinsicsDistortionModelArcTan = iid
End Function
Public Function IID_IMFExtendedCameraIntrinsics() As UUID
'{687F6DAC-6987-4750-A16A-734D1E7A10FE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H687F6DAC, CInt(&H6987), CInt(&H4750), &HA1, &H6A, &H73, &H4D, &H1E, &H7A, &H10, &HFE)
IID_IMFExtendedCameraIntrinsics = iid
End Function
Public Function IID_IMFExtendedCameraControl() As UUID
'{38E33520-FCA1-4845-A27A-68B7C6AB3789}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H38E33520, CInt(&HFCA1), CInt(&H4845), &HA2, &H7A, &H68, &HB7, &HC6, &HAB, &H37, &H89)
IID_IMFExtendedCameraControl = iid
End Function
Public Function IID_IMFExtendedCameraController() As UUID
'{B91EBFEE-CA03-4AF4-8A82-A31752F4A0FC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB91EBFEE, CInt(&HCA03), CInt(&H4AF4), &H8A, &H82, &HA3, &H17, &H52, &HF4, &HA0, &HFC)
IID_IMFExtendedCameraController = iid
End Function
Public Function IID_IMFRelativePanelReport() As UUID
'{F25362EA-2C0E-447F-81E2-755914CDC0C3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF25362EA, CInt(&H2C0E), CInt(&H447F), &H81, &HE2, &H75, &H59, &H14, &HCD, &HC0, &HC3)
IID_IMFRelativePanelReport = iid
End Function
Public Function IID_IMFRelativePanelWatcher() As UUID
'{421AF7F6-573E-4AD0-8FDA-2E57CEDB18C6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H421AF7F6, CInt(&H573E), CInt(&H4AD0), &H8F, &HDA, &H2E, &H57, &HCE, &HDB, &H18, &HC6)
IID_IMFRelativePanelWatcher = iid
End Function
Public Function IID_IMFVideoCaptureSampleAllocator() As UUID
'{725B77C7-CA9F-4FE5-9D72-9946BF9B3C70}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H725B77C7, CInt(&HCA9F), CInt(&H4FE5), &H9D, &H72, &H99, &H46, &HBF, &H9B, &H3C, &H70)
IID_IMFVideoCaptureSampleAllocator = iid
End Function
Public Function IID_IMFSampleAllocatorControl() As UUID
'{DA62B958-3A38-4A97-BD27-149C640C0771}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDA62B958, CInt(&H3A38), CInt(&H4A97), &HBD, &H27, &H14, &H9C, &H64, &HC, &H7, &H71)
IID_IMFSampleAllocatorControl = iid
End Function
Public Function IID_IMFCameraOcclusionStateReport() As UUID
'{1640B2CF-74DA-4462-A43B-B76D3BDC1434}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1640B2CF, CInt(&H74DA), CInt(&H4462), &HA4, &H3B, &HB7, &H6D, &H3B, &HDC, &H14, &H34)
IID_IMFCameraOcclusionStateReport = iid
End Function
Public Function IID_IMFCameraOcclusionStateReportCallback() As UUID
'{6E5841C7-3889-4019-9035-783FB19B5948}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6E5841C7, CInt(&H3889), CInt(&H4019), &H90, &H35, &H78, &H3F, &HB1, &H9B, &H59, &H48)
IID_IMFCameraOcclusionStateReportCallback = iid
End Function
Public Function IID_IMFCameraOcclusionStateMonitor() As UUID
'{CC692F46-C697-47E2-A72D-7B064617749B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCC692F46, CInt(&HC697), CInt(&H47E2), &HA7, &H2D, &H7B, &H6, &H46, &H17, &H74, &H9B)
IID_IMFCameraOcclusionStateMonitor = iid
End Function
Public Function IID_IMFCameraControlNotify() As UUID
'{E8F2540D-558A-4449-8B64-4863467A9FE8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE8F2540D, CInt(&H558A), CInt(&H4449), &H8B, &H64, &H48, &H63, &H46, &H7A, &H9F, &HE8)
IID_IMFCameraControlNotify = iid
End Function
Public Function IID_IMFCameraControlMonitor() As UUID
'{4D46F2C9-28BA-4970-8C7B-1F0C9D80AF69}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4D46F2C9, CInt(&H28BA), CInt(&H4970), &H8C, &H7B, &H1F, &HC, &H9D, &H80, &HAF, &H69)
IID_IMFCameraControlMonitor = iid
End Function
Public Function IID_IMFCameraControlDefaults() As UUID
'{75510662-B034-48F4-88A7-8DE61DAA4AF9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H75510662, CInt(&HB034), CInt(&H48F4), &H88, &HA7, &H8D, &HE6, &H1D, &HAA, &H4A, &HF9)
IID_IMFCameraControlDefaults = iid
End Function
Public Function IID_IMFCameraControlDefaultsCollection() As UUID
'{92D43D0F-54A8-4BAE-96DA-356D259A5C26}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H92D43D0F, CInt(&H54A8), CInt(&H4BAE), &H96, &HDA, &H35, &H6D, &H25, &H9A, &H5C, &H26)
IID_IMFCameraControlDefaultsCollection = iid
End Function
Public Function IID_IMFCameraConfigurationManager() As UUID
'{A624F617-4704-4206-8A6D-EBDA4A093985}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA624F617, CInt(&H4704), CInt(&H4206), &H8A, &H6D, &HEB, &HDA, &H4A, &H9, &H39, &H85)
IID_IMFCameraConfigurationManager = iid
End Function






Public Function MF_WVC1_PROG_SINGLE_SLICE_CONTENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H67EC2559, &HF2F, &H4420, &HA4, &HDD, &H2F, &H8E, &HE7, &HA5, &H73, &H8B)
MF_WVC1_PROG_SINGLE_SLICE_CONTENT = iid
End Function
Public Function MF_PROGRESSIVE_CODING_CONTENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8F020EEA, &H1508, &H471F, &H9D, &HA6, &H50, &H7D, &H7C, &HFA, &H40, &HDB)
MF_PROGRESSIVE_CODING_CONTENT = iid
End Function
Public Function MF_NALU_LENGTH_SET() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA7911D53, &H12A4, &H4965, &HAE, &H70, &H6E, &HAD, &HD6, &HFF, &H5, &H51)
MF_NALU_LENGTH_SET = iid
End Function
Public Function MF_NALU_LENGTH_INFORMATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H19124E7C, &HAD4B, &H465F, &HBB, &H18, &H20, &H18, &H62, &H87, &HB6, &HAF)
MF_NALU_LENGTH_INFORMATION = iid
End Function
Public Function MF_USER_DATA_PAYLOAD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD1D4985D, &HDC92, &H457A, &HB3, &HA0, &H65, &H1A, &H33, &HA3, &H10, &H47)
MF_USER_DATA_PAYLOAD = iid
End Function
Public Function MF_MPEG4SINK_SPSPPS_PASSTHROUGH() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5601A134, &H2005, &H4AD2, &HB3, &H7D, &H22, &HA6, &HC5, &H54, &HDE, &HB2)
MF_MPEG4SINK_SPSPPS_PASSTHROUGH = iid
End Function
Public Function MF_MPEG4SINK_MOOV_BEFORE_MDAT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF672E3AC, &HE1E6, &H4F10, &HB5, &HEC, &H5F, &H3B, &H30, &H82, &H88, &H16)
MF_MPEG4SINK_MOOV_BEFORE_MDAT = iid
End Function
Public Function MF_SESSION_TOPOLOADER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1E83D482, &H1F1C, &H4571, &H84, &H5, &H88, &HF4, &HB2, &H18, &H1F, &H71)
MF_SESSION_TOPOLOADER = iid
End Function
Public Function MF_SESSION_GLOBAL_TIME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1E83D482, &H1F1C, &H4571, &H84, &H5, &H88, &HF4, &HB2, &H18, &H1F, &H72)
MF_SESSION_GLOBAL_TIME = iid
End Function
Public Function MF_SESSION_QUALITY_MANAGER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1E83D482, &H1F1C, &H4571, &H84, &H5, &H88, &HF4, &HB2, &H18, &H1F, &H73)
MF_SESSION_QUALITY_MANAGER = iid
End Function
Public Function MF_SESSION_CONTENT_PROTECTION_MANAGER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1E83D482, &H1F1C, &H4571, &H84, &H5, &H88, &HF4, &HB2, &H18, &H1F, &H74)
MF_SESSION_CONTENT_PROTECTION_MANAGER = iid
End Function
Public Function MF_SESSION_SERVER_CONTEXT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAFE5B291, &H50FA, &H46E8, &HB9, &HBE, &HC, &HC, &H3C, &HE4, &HB3, &HA5)
MF_SESSION_SERVER_CONTEXT = iid
End Function
Public Function MF_SESSION_REMOTE_SOURCE_MODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF4033EF4, &H9BB3, &H4378, &H94, &H1F, &H85, &HA0, &H85, &H6B, &HC2, &H44)
MF_SESSION_REMOTE_SOURCE_MODE = iid
End Function
Public Function MF_SESSION_APPROX_EVENT_OCCURRENCE_TIME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H190E852F, &H6238, &H42D1, &HB5, &HAF, &H69, &HEA, &H33, &H8E, &HF8, &H50)
MF_SESSION_APPROX_EVENT_OCCURRENCE_TIME = iid
End Function
Public Function MF_PMP_SERVER_CONTEXT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2F00C910, &HD2CF, &H4278, &H8B, &H6A, &HD0, &H77, &HFA, &HC3, &HA2, &H5F)
MF_PMP_SERVER_CONTEXT = iid
End Function



Public Function MFPKEY_SourceOpenMonitor() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H74D4637, &HB5AE, &H465D, &HAF, &H17, &H1A, &H53, &H8D, &H28, &H59, &HDD, &H2)
End Function


' Type: VT_BOOL
' When this is set to VARIANT_TRUE, if an ASF Media Source is created,
' it will perform all seek operations approximately (and more quickly)
Public Function MFPKEY_ASFMediaSource_ApproxSeek() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &HB4CD270F, &H244D, &H4969, &HBB, &H92, &H3F, &HF, &HB8, &H31, &H6F, &H10, &H1)
MFPKEY_ASFMediaSource_ApproxSeek = pk
End Function

' Type: VT_BOOL
' When this is set to VARIANT_TRUE, if an ASF Media Source is created,
' it will perform iterative seek if there is  no index
Public Function MFPKEY_ASFMediaSource_IterativeSeekIfNoIndex() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H170B65DC, &H4A4E, &H407A, &HAC, &H22, &H57, &H7F, &H50, &HE4, &HA3, &H7C, &H1)
MFPKEY_ASFMediaSource_IterativeSeekIfNoIndex = pk
End Function
' Type: VT_UINT32
' Only valid when MFPKEY_ASFMediaSource_IterativeSeekIfNoIndex is set to TRUE
' The count is any integer [1, 10]
' If this value is not set, the default value 5 is used.
Public Function MFPKEY_ASFMediaSource_IterativeSeek_Max_Count() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H170B65DC, &H4A4E, &H407A, &HAC, &H22, &H57, &H7F, &H50, &HE4, &HA3, &H7C, &H2)
MFPKEY_ASFMediaSource_IterativeSeek_Max_Count = pk
End Function
' Type: VT_UINT32
' Only valid when MFPKEY_ASFMediaSource_IterativeSeekIfNoIndex is set to TRUE
' the tolerance zone is the difference that allowed between the real seek time and preferred seek time.
' Keyframe distance is recommended to use.
' If this value is not set, the default value 8000 millisecond is used.
Public Function MFPKEY_ASFMediaSource_IterativeSeek_Tolerance_In_MilliSecond() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H170B65DC, &H4A4E, &H407A, &HAC, &H22, &H57, &H7F, &H50, &HE4, &HA3, &H7C, &H3)
MFPKEY_ASFMediaSource_IterativeSeek_Tolerance_In_MilliSecond = pk
End Function
'
' DLNA Profile ID - needed for media sharing.
'
' {CFA31B45-525D-4998-BB44-3F7D81542FA4}
Public Function MFPKEY_Content_DLNA_Profile_ID() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &HCFA31B45, &H525D, &H4998, &HBB, &H44, &H3F, &H7D, &H81, &H54, &H2F, &HA4, &H1)
MFPKEY_Content_DLNA_Profile_ID = pk
End Function
' Type: VT_BOOL
' When this is set to VARIANT_TRUE, the media source is requested to disable any read-ahead.
' This can be a useful performance optimization to limit disk read when a media source will
' only be instantiated for limited tasks, such as reading video thumbnail data.
' Not all sources will support this feature.
' {26366C14-C5BF-4c76-887B-9F1754DB5F09}
Public Function MFPKEY_MediaSource_DisableReadAhead() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H26366C14, &HC5BF, &H4C76, &H88, &H7B, &H9F, &H17, &H54, &HDB, &H5F, &H9, &H1)
MFPKEY_MediaSource_DisableReadAhead = pk
End Function
' Type: VT_UINT32
' Sets the SBE mode.
' 0: default is to use the automatic stream mapping in the crossbar to the output
' 1: Crossbar output multiple streams mapped to the output
' 2: Crossbar mode where the application has to map the streams to the output (selection of the audio stream possible)
' {3FAE10BB-F859-4192-B562-1868D3DA3A02}
Public Function MFPKEY_SBESourceMode() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H3FAE10BB, &HF859, &H4192, &HB5, &H62, &H18, &H68, &HD3, &HDA, &H3A, &H2, &H1)
MFPKEY_SBESourceMode = pk
End Function
' Type: VT_UNKNOWN
' Defines an IMFAsyncCallback implementation that will create the a PMP session on behalf of the bytestream.
' {28bb4de2-26a2-4870-b720-d26bbeb14942}
Public Function MFPKEY_PMP_Creation_Callback() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H28BB4DE2, &H26A2, &H4870, &HB7, &H20, &HD2, &H6B, &HBE, &HB1, &H49, &H42, &H1)
MFPKEY_PMP_Creation_Callback = pk
End Function
' Type: VT_BOOL
' When set and TRUE, specifies that the HTTP caching bytestream should use URLMon to download
' content.  By default, WinHTTP will be used.
' {eda8afdf-c171-417f-8d17-2e0918303292}, 1
Public Function MFPKEY_HTTP_ByteStream_Enable_Urlmon() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &HEDA8AFDF, &HC171, &H417F, &H8D, &H17, &H2E, &H9, &H18, &H30, &H32, &H92, &H1)
MFPKEY_HTTP_ByteStream_Enable_Urlmon = pk
End Function
' Type: VT_UI4
' When MFPKEY_HTTP_ByteStream_Enable_Urlmon is turned on, this value specifies the urlmon
' bind flags as defined in the BINDF enumeration.  The default value is BINDF_ASYNCHRONOUS |
' BINDF_ASYNCSTORAGE | BINDF_NOWRITECACHE | BINDF_PULLDATA | BINDF_RESYNCHRONIZE
' {eda8afdf-c171-417f-8d17-2e0918303292}, 2
Public Function MFPKEY_HTTP_ByteStream_Urlmon_Bind_Flags() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &HEDA8AFDF, &HC171, &H417F, &H8D, &H17, &H2E, &H9, &H18, &H30, &H32, &H92, &H2)
MFPKEY_HTTP_ByteStream_Urlmon_Bind_Flags = pk
End Function
' Type: VT_VECTOR | VT_UI1
' When MFPKEY_HTTP_ByteStream_Enable_Urlmon is turned on, this value specifies the root security
' ID for urlmon.  By default, this value is null and no root security ID will be provided to
' urlmon.
' {eda8afdf-c171-417f-8d17-2e0918303292}, 3
Public Function MFPKEY_HTTP_ByteStream_Urlmon_Security_Id() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &HEDA8AFDF, &HC171, &H417F, &H8D, &H17, &H2E, &H9, &H18, &H30, &H32, &H92, &H3)
MFPKEY_HTTP_ByteStream_Urlmon_Security_Id = pk
End Function
' Type: VT_UNKNOWN
' When MFPKEY_HTTP_ByteStream_Enable_Urlmon is turned on, this value specifies an
' implementation of IWindowForBindingUI that can be used to obtain an HWND for urlmon
' UI.  By default, urlmon UI will be disabled.
' {eda8afdf-c171-417f-8d17-2e0918303292}, 4
Public Function MFPKEY_HTTP_ByteStream_Urlmon_Window() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &HEDA8AFDF, &HC171, &H417F, &H8D, &H17, &H2E, &H9, &H18, &H30, &H32, &H92, &H4)
MFPKEY_HTTP_ByteStream_Urlmon_Window = pk
End Function
' Type: VT_UNKNOWN
' When MFPKEY_HTTP_ByteStream_Enable_Urlmon is turned on, this value specifies an
' implementation of IServiceProvider that can be used to obtain services for the
' urlmon protocol handler.
' {eda8afdf-c171-417f-8d17-2e0918303292}, 5
Public Function MFPKEY_HTTP_ByteStream_Urlmon_Callback_QueryService() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &HEDA8AFDF, &HC171, &H417F, &H8D, &H17, &H2E, &H9, &H18, &H30, &H32, &H92, &H5)
MFPKEY_HTTP_ByteStream_Urlmon_Callback_QueryService = pk
End Function
' Type: VT_CLSID
' Set to the GUID that identifies the media protection system to use for the content.
' {636B271D-DDC7-49E9-A6C6-47385962E5BD}
Public Function MFPKEY_MediaProtectionSystemId() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H636B271D, &HDDC7, &H49E9, &HA6, &HC6, &H47, &H38, &H59, &H62, &HE5, &HBD, &H1)
MFPKEY_MediaProtectionSystemId = pk
End Function

' Type: VT_BLOB
' BLOB containing the context to use when initializing a media protection system's trusted input module.
' {636B271D-DDC7-49E9-A6C6-47385962E5BD}
Public Function MFPKEY_MediaProtectionSystemContext() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H636B271D, &HDDC7, &H49E9, &HA6, &HC6, &H47, &H38, &H59, &H62, &HE5, &HBD, &H2)
MFPKEY_MediaProtectionSystemContext = pk
End Function
' Type: VT_UNKNOWN
' Set to an IPropertySet that defines the mapping from Property system id to property system activation id.
' {636B271D-DDC7-49E9-A6C6-47385962E5BD}
Public Function MFPKEY_MediaProtectionSystemIdMapping() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H636B271D, &HDDC7, &H49E9, &HA6, &HC6, &H47, &H38, &H59, &H62, &HE5, &HBD, &H3)
MFPKEY_MediaProtectionSystemIdMapping = pk
End Function
' Type: VT_CLSID
' Set to the GUID that identifies the protection system in the container.
' {42AF3D7C-00CF-4a0f-81F0-ADF524A5A5B5}
Public Function MFPKEY_MediaProtectionContainerGuid() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H42AF3D7C, &HCF, &H4A0F, &H81, &HF0, &HAD, &HF5, &H24, &HA5, &HA5, &HB5, &H1)
MFPKEY_MediaProtectionContainerGuid = pk
End Function
' Type: VT_UNKNOWN
' Set to an IPropertySet that defines a mapping from track Type to IRandomAccessStream containing the DRM context
' {4454B092-D3DA-49b0-8452-6850C7DB764D}
Public Function MFPKEY_MediaProtectionSystemContextsPerTrack() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H4454B092, &HD3DA, &H49B0, &H84, &H52, &H68, &H50, &HC7, &HDB, &H76, &H4D, &H3)
MFPKEY_MediaProtectionSystemContextsPerTrack = pk
End Function
' Type: VT_BOOL
' When set and TRUE, specifies that the URL is being downloaded to disk instead of being played.
' {817f11b7-a982-46ec-a449-ef58aed53ca8}
Public Function MFPKEY_HTTP_ByteStream_Download_Mode() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H817F11B7, &HA982, &H46EC, &HA4, &H49, &HEF, &H58, &HAE, &HD5, &H3C, &HA8, &H1)
MFPKEY_HTTP_ByteStream_Download_Mode = pk
End Function
' TYPE: VT_UI4
' This property specifies how the HTTP Byte Stream should cache downloaded data.
' A value of 1 means that the downloaded data should be cached to disk.
' A value of 2 means that the downloaded data should be cached in memory.
' A value of 0 is the default, and means that the Byte Stream is free to choose the caching mode
' based on heuristics.
' {86a2403e-c78b-44d7-8bc8-ff7258117508}, 1
Public Function MFPKEY_HTTP_ByteStream_Caching_Mode() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H86A2403E, &HC78B, &H44D7, &H8B, &HC8, &HFF, &H72, &H58, &H11, &H75, &H8, &H1)
MFPKEY_HTTP_ByteStream_Caching_Mode = pk
End Function
' TYPE: VT_UI8
' This property specifies an upper limit on the amount of data, in bytes, that the
' HTTP Byte Stream caches on disk or in memory.
' The Byte Stream may choose a lower limit than the one specified.
' A value of 0 is the default, and means that the Byte Stream is free to limit the cache size
' based on heuristics.
' {86a2403e-c78b-44d7-8bc8-ff7258117508}, 2
Public Function MFPKEY_HTTP_ByteStream_Cache_Limit() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H86A2403E, &HC78B, &H44D7, &H8B, &HC8, &HFF, &H72, &H58, &H11, &H75, &H8, &H2)
MFPKEY_HTTP_ByteStream_Cache_Limit = pk
End Function

Public Function MFPKEY_CLSID() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &HC57A84C0, &H1A80, &H40A3, &H97, &HB5, &H92, &H72, &HA4, &H3, &HC8, &HAE, &H1)
 MFPKEY_CLSID = pk
End Function
Public Function MFPKEY_CATEGORY() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &HC57A84C0, &H1A80, &H40A3, &H97, &HB5, &H92, &H72, &HA4, &H3, &HC8, &HAE, &H2)
 MFPKEY_CATEGORY = pk
End Function
Public Function MFPKEY_EXATTRIBUTE_SUPPORTED() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H456FE843, &H3C87, &H40C0, &H94, &H9D, &H14, &H9, &HC9, &H7D, &HAB, &H2C, &H1)
 MFPKEY_EXATTRIBUTE_SUPPORTED = pk
End Function
Public Function MFPKEY_MULTICHANNEL_CHANNEL_MASK() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H58BDAF8C, &H3224, &H4692, &H86, &HD0, &H44, &HD6, &H5C, &H5B, &HF8, &H2B, &H1)
 MFPKEY_MULTICHANNEL_CHANNEL_MASK = pk
End Function
Public Function MF_EME_INITDATATYPES() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H497D231B, &H4EB9, &H4DF0, &HB4, &H74, &HB9, &HAF, &HEB, &HA, &HDF, &H38, PID_FIRST_USABLE + &H1)
 MF_EME_INITDATATYPES = pk
End Function
Public Function MF_EME_DISTINCTIVEID() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H7DC9C4A5, &H12BE, &H497E, &H8B, &HFF, &H9B, &H60, &HB2, &HDC, &H58, &H45, PID_FIRST_USABLE + &H2)
 MF_EME_DISTINCTIVEID = pk
End Function
Public Function MF_EME_PERSISTEDSTATE() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H5D4DF6AE, &H9AF1, &H4E3D, &H95, &H5B, &HE, &H4B, &HD2, &H2F, &HED, &HF0, PID_FIRST_USABLE + &H3)
 MF_EME_PERSISTEDSTATE = pk
End Function
Public Function MF_EME_AUDIOCAPABILITIES() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H980FBB84, &H297D, &H4EA7, &H89, &H5F, &HBC, &HF2, &H8A, &H46, &H28, &H81, PID_FIRST_USABLE + &H4)
 MF_EME_AUDIOCAPABILITIES = pk
End Function
Public Function MF_EME_VIDEOCAPABILITIES() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &HB172F83D, &H30DD, &H4C10, &H80, &H6, &HED, &H53, &HDA, &H4D, &H3B, &HDB, PID_FIRST_USABLE + &H5)
 MF_EME_VIDEOCAPABILITIES = pk
End Function
Public Function MF_EME_LABEL() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H9EAE270E, &HB2D7, &H4817, &HB8, &H8F, &H54, &H0, &H99, &HF2, &HEF, &H4E, PID_FIRST_USABLE + &H6)
 MF_EME_LABEL = pk
End Function
Public Function MF_EME_SESSIONTYPES() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H7623384F, &HF5, &H4376, &H86, &H98, &H34, &H58, &HDB, &H3, &HE, &HD5, PID_FIRST_USABLE + &H7)
 MF_EME_SESSIONTYPES = pk
End Function
Public Function MF_EME_ROBUSTNESS() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H9D3D2B9E, &H7023, &H4944, &HA8, &HF5, &HEC, &HCA, &H52, &HA4, &H69, &H90, PID_FIRST_USABLE + &H1)
 MF_EME_ROBUSTNESS = pk
End Function
Public Function MF_EME_CONTENTTYPE() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H289FB1FC, &HD9C4, &H4CC7, &HB2, &HBE, &H97, &H2B, &HE, &H9B, &H28, &H3A, PID_FIRST_USABLE + &H2)
 MF_EME_CONTENTTYPE = pk
End Function
Public Function MF_EME_CDM_INPRIVATESTOREPATH() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &HEC305FD9, &H39F, &H4AC8, &H98, &HDA, &HE7, &H92, &H1E, &H0, &H6A, &H90, PID_FIRST_USABLE + &H1)
 MF_EME_CDM_INPRIVATESTOREPATH = pk
End Function
Public Function MF_EME_CDM_STOREPATH() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &HF795841E, &H99F9, &H44D7, &HAF, &HC0, &HD3, &H9, &HC0, &H4C, &H94, &HAB, PID_FIRST_USABLE + &H2)
 MF_EME_CDM_STOREPATH = pk
End Function
Public Function MF_CONTENTDECRYPTIONMODULE_INPRIVATESTOREPATH() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H730CB3AC, &H51DC, &H49DA, &HA5, &H78, &HB9, &H53, &H86, &HB6, &H2A, &HFE, &H1)
 MF_CONTENTDECRYPTIONMODULE_INPRIVATESTOREPATH = pk
End Function
Public Function MF_CONTENTDECRYPTIONMODULE_STOREPATH() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H77D993B9, &HBA61, &H4BB7, &H92, &HC6, &H18, &HC8, &H6A, &H18, &H9C, &H6, &H2)
 MF_CONTENTDECRYPTIONMODULE_STOREPATH = pk
End Function
Public Function MF_CONTENTDECRYPTIONMODULE_PMPSTORECONTEXT() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H6D2A2835, &HC3A9, &H4681, &H97, &HF2, &HA, &HF5, &H6B, &HE9, &H34, &H46, &H3)
 MF_CONTENTDECRYPTIONMODULE_PMPSTORECONTEXT = pk
End Function
Public Function DEVPKEY_DeviceInterface_IsVirtualCamera() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H6EDC630D, &HC2E3, &H43B7, &HB2, &HD1, &H20, &H52, &H5A, &H1A, &HF1, &H20, 3)
 DEVPKEY_DeviceInterface_IsVirtualCamera = pk
End Function
Public Function DEVPKEY_DeviceInterface_IsWindowsCameraEffectAvailable() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H6EDC630D, &HC2E3, &H43B7, &HB2, &HD1, &H20, &H52, &H5A, &H1A, &HF1, &H20, 4)
 DEVPKEY_DeviceInterface_IsWindowsCameraEffectAvailable = pk
End Function
Public Function DEVPKEY_DeviceInterface_VirtualCameraAssociatedCameras() As PROPERTYKEY
Static pk As PROPERTYKEY
 If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &H6EDC630D, &HC2E3, &H43B7, &HB2, &HD1, &H20, &H52, &H5A, &H1A, &HF1, &H20, 5)
 DEVPKEY_DeviceInterface_VirtualCameraAssociatedCameras = pk
End Function


Public Function MF_TIME_FORMAT_ENTRY_RELATIVE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4399F178, &H46D3, &H4504, &HAF, &HDA, &H20, &HD3, &H2E, &H9B, &HA3, &H60)
MF_TIME_FORMAT_ENTRY_RELATIVE = iid
End Function
Public Function MF_SOURCE_STREAM_SUPPORTS_HW_CONNECTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA38253AA, &H6314, &H42FD, &HA3, &HCE, &HBB, &H27, &HB6, &H85, &H99, &H46)
MF_SOURCE_STREAM_SUPPORTS_HW_CONNECTION = iid
End Function
Public Function MF_STREAM_SINK_SUPPORTS_HW_CONNECTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9B465CBF, &H597, &H4F9E, &H9F, &H3C, &HB9, &H7E, &HEE, &HF9, &H3, &H59)
MF_STREAM_SINK_SUPPORTS_HW_CONNECTION = iid
End Function
Public Function MF_STREAM_SINK_SUPPORTS_ROTATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB3E96280, &HBD05, &H41A5, &H97, &HAD, &H8A, &H7F, &HEE, &H24, &HB9, &H12)
MF_STREAM_SINK_SUPPORTS_ROTATION = iid
End Function
Public Function MF_SINK_VIDEO_PTS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2162BDE7, &H421E, &H4B90, &H9B, &H33, &HE5, &H8F, &HBF, &H1D, &H58, &HB6)
MF_SINK_VIDEO_PTS = iid
End Function
Public Function MF_SINK_VIDEO_NATIVE_WIDTH() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE6D6A707, &H1505, &H4747, &H9B, &H10, &H72, &HD2, &HD1, &H58, &HCB, &H3A)
MF_SINK_VIDEO_NATIVE_WIDTH = iid
End Function
Public Function MF_SINK_VIDEO_NATIVE_HEIGHT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF0CA6705, &H490C, &H43E8, &H94, &H1C, &HC0, &HB3, &H20, &H6B, &H9A, &H65)
MF_SINK_VIDEO_NATIVE_HEIGHT = iid
End Function
Public Function MF_SINK_VIDEO_DISPLAY_ASPECT_RATIO_NUMERATOR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD0F33B22, &HB78A, &H4879, &HB4, &H55, &HF0, &H3E, &HF3, &HFA, &H82, &HCD)
MF_SINK_VIDEO_DISPLAY_ASPECT_RATIO_NUMERATOR = iid
End Function
Public Function MF_SINK_VIDEO_DISPLAY_ASPECT_RATIO_DENOMINATOR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6EA1EB97, &H1FE0, &H4F10, &HA6, &HE4, &H1F, &H4F, &H66, &H15, &H64, &HE0)
MF_SINK_VIDEO_DISPLAY_ASPECT_RATIO_DENOMINATOR = iid
End Function
Public Function MF_BD_MVC_PLANE_OFFSET_METADATA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H62A654E4, &HB76C, &H4901, &H98, &H23, &H2C, &HB6, &H15, &HD4, &H73, &H18)
MF_BD_MVC_PLANE_OFFSET_METADATA = iid
End Function
Public Function MF_LUMA_KEY_ENABLE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7369820F, &H76DE, &H43CA, &H92, &H84, &H47, &HB8, &HF3, &H7E, &H6, &H49)
MF_LUMA_KEY_ENABLE = iid
End Function
Public Function MF_LUMA_KEY_LOWER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H93D7B8D5, &HB81, &H4715, &HAE, &HA0, &H87, &H25, &H87, &H16, &H21, &HE9)
MF_LUMA_KEY_LOWER = iid
End Function
Public Function MF_LUMA_KEY_UPPER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD09F39BB, &H4602, &H4C31, &HA7, &H6, &HA1, &H21, &H71, &HA5, &H11, &HA)
MF_LUMA_KEY_UPPER = iid
End Function
Public Function MF_USER_EXTENDED_ATTRIBUTES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC02ABAC6, &HFEB2, &H4541, &H92, &H2F, &H92, &HB, &H43, &H70, &H27, &H22)
MF_USER_EXTENDED_ATTRIBUTES = iid
End Function
Public Function MF_INDEPENDENT_STILL_IMAGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEA12AF41, &H710, &H42C9, &HA1, &H27, &HDA, &HA3, &HE7, &H84, &H83, &HA5)
MF_INDEPENDENT_STILL_IMAGE = iid
End Function
Public Function MF_TOPOLOGY_PROJECTSTART() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7ED3F802, &H86BB, &H4B3F, &HB7, &HE4, &H7C, &HB4, &H3A, &HFD, &H4B, &H80)
MF_TOPOLOGY_PROJECTSTART = iid
End Function
Public Function MF_TOPOLOGY_PROJECTSTOP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7ED3F803, &H86BB, &H4B3F, &HB7, &HE4, &H7C, &HB4, &H3A, &HFD, &H4B, &H80)
MF_TOPOLOGY_PROJECTSTOP = iid
End Function
Public Function MF_TOPOLOGY_NO_MARKIN_MARKOUT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7ED3F804, &H86BB, &H4B3F, &HB7, &HE4, &H7C, &HB4, &H3A, &HFD, &H4B, &H80)
MF_TOPOLOGY_NO_MARKIN_MARKOUT = iid
End Function
Public Function MF_TOPOLOGY_DXVA_MODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1E8D34F6, &HF5AB, &H4E23, &HBB, &H88, &H87, &H4A, &HA3, &HA1, &HA7, &H4D)
MF_TOPOLOGY_DXVA_MODE = iid
End Function
Public Function MF_TOPOLOGY_ENABLE_XVP_FOR_PLAYBACK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1967731F, &HCD78, &H42FC, &HB0, &H26, &H9, &H92, &HA5, &H6E, &H56, &H93)
MF_TOPOLOGY_ENABLE_XVP_FOR_PLAYBACK = iid
End Function
Public Function MF_TOPOLOGY_STATIC_PLAYBACK_OPTIMIZATIONS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB86CAC42, &H41A6, &H4B79, &H89, &H7A, &H1A, &HB0, &HE5, &H2B, &H4A, &H1B)
MF_TOPOLOGY_STATIC_PLAYBACK_OPTIMIZATIONS = iid
End Function
Public Function MF_TOPOLOGY_PLAYBACK_MAX_DIMS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5715CF19, &H5768, &H44AA, &HAD, &H6E, &H87, &H21, &HF1, &HB0, &HF9, &HBB)
MF_TOPOLOGY_PLAYBACK_MAX_DIMS = iid
End Function
Public Function MF_TOPOLOGY_HARDWARE_MODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD2D362FD, &H4E4F, &H4191, &HA5, &H79, &HC6, &H18, &HB6, &H67, &H6, &HAF)
MF_TOPOLOGY_HARDWARE_MODE = iid
End Function
Public Function MF_TOPOLOGY_PLAYBACK_FRAMERATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC164737A, &HC2B1, &H4553, &H83, &HBB, &H5A, &H52, &H60, &H72, &H44, &H8F)
MF_TOPOLOGY_PLAYBACK_FRAMERATE = iid
End Function
Public Function MF_TOPOLOGY_DYNAMIC_CHANGE_NOT_ALLOWED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD529950B, &HD484, &H4527, &HA9, &HCD, &HB1, &H90, &H95, &H32, &HB5, &HB0)
MF_TOPOLOGY_DYNAMIC_CHANGE_NOT_ALLOWED = iid
End Function
Public Function MF_TOPOLOGY_ENUMERATE_SOURCE_TYPES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6248C36D, &H5D0B, &H4F40, &HA0, &HBB, &HB0, &HB3, &H5, &HF7, &H76, &H98)
MF_TOPOLOGY_ENUMERATE_SOURCE_TYPES = iid
End Function
Public Function MF_TOPOLOGY_START_TIME_ON_PRESENTATION_SWITCH() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC8CC113F, &H7951, &H4548, &HAA, &HD6, &H9E, &HD6, &H20, &H2E, &H62, &HB3)
MF_TOPOLOGY_START_TIME_ON_PRESENTATION_SWITCH = iid
End Function
Public Function MF_DISABLE_LOCALLY_REGISTERED_PLUGINS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H66B16DA9, &HADD4, &H47E0, &HA1, &H6B, &H5A, &HF1, &HFB, &H48, &H36, &H34)
MF_DISABLE_LOCALLY_REGISTERED_PLUGINS = iid
End Function
Public Function MF_LOCAL_PLUGIN_CONTROL_POLICY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD91B0085, &HC86D, &H4F81, &H88, &H22, &H8C, &H68, &HE1, &HD7, &HFA, &H4)
MF_LOCAL_PLUGIN_CONTROL_POLICY = iid
End Function
Public Function MF_TOPONODE_FLUSH() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBCE8, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPONODE_FLUSH = iid
End Function
Public Function MF_TOPONODE_DRAIN() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBCE9, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPONODE_DRAIN = iid
End Function
Public Function MF_TOPONODE_D3DAWARE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBCED, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPONODE_D3DAWARE = iid
End Function
Public Function MF_TOPOLOGY_RESOLUTION_STATUS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBCDE, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPOLOGY_RESOLUTION_STATUS = iid
End Function
Public Function MF_TOPONODE_ERRORCODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBCEE, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPONODE_ERRORCODE = iid
End Function
Public Function MF_TOPONODE_CONNECT_METHOD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBCF1, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPONODE_CONNECT_METHOD = iid
End Function
Public Function MF_TOPONODE_LOCKED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBCF7, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPONODE_LOCKED = iid
End Function
Public Function MF_TOPONODE_WORKQUEUE_ID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBCF8, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPONODE_WORKQUEUE_ID = iid
End Function
Public Function MF_TOPONODE_WORKQUEUE_MMCSS_CLASS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBCF9, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPONODE_WORKQUEUE_MMCSS_CLASS = iid
End Function
Public Function MF_TOPONODE_DECRYPTOR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBCFA, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPONODE_DECRYPTOR = iid
End Function
Public Function MF_TOPONODE_DISCARDABLE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBCFB, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPONODE_DISCARDABLE = iid
End Function
Public Function MF_TOPONODE_ERROR_MAJORTYPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBCFD, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPONODE_ERROR_MAJORTYPE = iid
End Function
Public Function MF_TOPONODE_ERROR_SUBTYPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBCFE, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPONODE_ERROR_SUBTYPE = iid
End Function
Public Function MF_TOPONODE_WORKQUEUE_MMCSS_TASKID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBCFF, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPONODE_WORKQUEUE_MMCSS_TASKID = iid
End Function
Public Function MF_TOPONODE_WORKQUEUE_MMCSS_PRIORITY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5001F840, &H2816, &H48F4, &H93, &H64, &HAD, &H1E, &HF6, &H61, &HA1, &H23)
MF_TOPONODE_WORKQUEUE_MMCSS_PRIORITY = iid
End Function
Public Function MF_TOPONODE_WORKQUEUE_ITEM_PRIORITY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA1FF99BE, &H5E97, &H4A53, &HB4, &H94, &H56, &H8C, &H64, &H2C, &HF, &HF3)
MF_TOPONODE_WORKQUEUE_ITEM_PRIORITY = iid
End Function
Public Function MF_TOPONODE_MARKIN_HERE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBD00, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPONODE_MARKIN_HERE = iid
End Function
Public Function MF_TOPONODE_MARKOUT_HERE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBD01, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPONODE_MARKOUT_HERE = iid
End Function
Public Function MF_TOPONODE_DECODER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494BBD02, &HB031, &H4E38, &H97, &HC4, &HD5, &H42, &H2D, &HD6, &H18, &HDC)
MF_TOPONODE_DECODER = iid
End Function
Public Function MF_TOPONODE_MEDIASTART() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H835C58EA, &HE075, &H4BC7, &HBC, &HBA, &H4D, &HE0, &H0, &HDF, &H9A, &HE6)
MF_TOPONODE_MEDIASTART = iid
End Function
Public Function MF_TOPONODE_MEDIASTOP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H835C58EB, &HE075, &H4BC7, &HBC, &HBA, &H4D, &HE0, &H0, &HDF, &H9A, &HE6)
MF_TOPONODE_MEDIASTOP = iid
End Function
Public Function MF_TOPONODE_SOURCE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H835C58EC, &HE075, &H4BC7, &HBC, &HBA, &H4D, &HE0, &H0, &HDF, &H9A, &HE6)
MF_TOPONODE_SOURCE = iid
End Function
Public Function MF_TOPONODE_PRESENTATION_DESCRIPTOR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H835C58ED, &HE075, &H4BC7, &HBC, &HBA, &H4D, &HE0, &H0, &HDF, &H9A, &HE6)
MF_TOPONODE_PRESENTATION_DESCRIPTOR = iid
End Function
Public Function MF_TOPONODE_STREAM_DESCRIPTOR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H835C58EE, &HE075, &H4BC7, &HBC, &HBA, &H4D, &HE0, &H0, &HDF, &H9A, &HE6)
MF_TOPONODE_STREAM_DESCRIPTOR = iid
End Function
Public Function MF_TOPONODE_SEQUENCE_ELEMENTID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H835C58EF, &HE075, &H4BC7, &HBC, &HBA, &H4D, &HE0, &H0, &HDF, &H9A, &HE6)
MF_TOPONODE_SEQUENCE_ELEMENTID = iid
End Function
Public Function MF_TOPONODE_TRANSFORM_OBJECTID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H88DCC0C9, &H293E, &H4E8B, &H9A, &HEB, &HA, &HD6, &H4C, &HC0, &H16, &HB0)
MF_TOPONODE_TRANSFORM_OBJECTID = iid
End Function
Public Function MF_TOPONODE_STREAMID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H14932F9B, &H9087, &H4BB4, &H84, &H12, &H51, &H67, &H14, &H5C, &HBE, &H4)
MF_TOPONODE_STREAMID = iid
End Function
Public Function MF_TOPONODE_NOSHUTDOWN_ON_REMOVE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H14932F9C, &H9087, &H4BB4, &H84, &H12, &H51, &H67, &H14, &H5C, &HBE, &H4)
MF_TOPONODE_NOSHUTDOWN_ON_REMOVE = iid
End Function
Public Function MF_TOPONODE_RATELESS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H14932F9D, &H9087, &H4BB4, &H84, &H12, &H51, &H67, &H14, &H5C, &HBE, &H4)
MF_TOPONODE_RATELESS = iid
End Function
Public Function MF_TOPONODE_DISABLE_PREROLL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H14932F9E, &H9087, &H4BB4, &H84, &H12, &H51, &H67, &H14, &H5C, &HBE, &H4)
MF_TOPONODE_DISABLE_PREROLL = iid
End Function
Public Function MF_TOPONODE_PRIMARYOUTPUT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6304EF99, &H16B2, &H4EBE, &H9D, &H67, &HE4, &HC5, &H39, &HB3, &HA2, &H59)
MF_TOPONODE_PRIMARYOUTPUT = iid
End Function
Public Function MF_PD_PMPHOST_CONTEXT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C990D31, &HBB8E, &H477A, &H85, &H98, &HD, &H5D, &H96, &HFC, &HD8, &H8A)
MF_PD_PMPHOST_CONTEXT = iid
End Function
Public Function MF_PD_APP_CONTEXT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C990D32, &HBB8E, &H477A, &H85, &H98, &HD, &H5D, &H96, &HFC, &HD8, &H8A)
MF_PD_APP_CONTEXT = iid
End Function
Public Function MF_PD_DURATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C990D33, &HBB8E, &H477A, &H85, &H98, &HD, &H5D, &H96, &HFC, &HD8, &H8A)
MF_PD_DURATION = iid
End Function
Public Function MF_PD_TOTAL_FILE_SIZE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C990D34, &HBB8E, &H477A, &H85, &H98, &HD, &H5D, &H96, &HFC, &HD8, &H8A)
MF_PD_TOTAL_FILE_SIZE = iid
End Function
Public Function MF_PD_AUDIO_ENCODING_BITRATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C990D35, &HBB8E, &H477A, &H85, &H98, &HD, &H5D, &H96, &HFC, &HD8, &H8A)
MF_PD_AUDIO_ENCODING_BITRATE = iid
End Function
Public Function MF_PD_VIDEO_ENCODING_BITRATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C990D36, &HBB8E, &H477A, &H85, &H98, &HD, &H5D, &H96, &HFC, &HD8, &H8A)
MF_PD_VIDEO_ENCODING_BITRATE = iid
End Function
Public Function MF_PD_MIME_TYPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C990D37, &HBB8E, &H477A, &H85, &H98, &HD, &H5D, &H96, &HFC, &HD8, &H8A)
MF_PD_MIME_TYPE = iid
End Function
Public Function MF_PD_LAST_MODIFIED_TIME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C990D38, &HBB8E, &H477A, &H85, &H98, &HD, &H5D, &H96, &HFC, &HD8, &H8A)
MF_PD_LAST_MODIFIED_TIME = iid
End Function
Public Function MF_PD_PLAYBACK_ELEMENT_ID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C990D39, &HBB8E, &H477A, &H85, &H98, &HD, &H5D, &H96, &HFC, &HD8, &H8A)
MF_PD_PLAYBACK_ELEMENT_ID = iid
End Function
Public Function MF_PD_PREFERRED_LANGUAGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C990D3A, &HBB8E, &H477A, &H85, &H98, &HD, &H5D, &H96, &HFC, &HD8, &H8A)
MF_PD_PREFERRED_LANGUAGE = iid
End Function
Public Function MF_PD_PLAYBACK_BOUNDARY_TIME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C990D3B, &HBB8E, &H477A, &H85, &H98, &HD, &H5D, &H96, &HFC, &HD8, &H8A)
MF_PD_PLAYBACK_BOUNDARY_TIME = iid
End Function
Public Function MF_PD_AUDIO_ISVARIABLEBITRATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H33026EE0, &HE387, &H4582, &HAE, &HA, &H34, &HA2, &HAD, &H3B, &HAA, &H18)
MF_PD_AUDIO_ISVARIABLEBITRATE = iid
End Function
Public Function MF_SD_LANGUAGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAF2180, &HBDC2, &H423C, &HAB, &HCA, &HF5, &H3, &H59, &H3B, &HC1, &H21)
MF_SD_LANGUAGE = iid
End Function
Public Function MF_SD_PROTECTED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAF2181, &HBDC2, &H423C, &HAB, &HCA, &HF5, &H3, &H59, &H3B, &HC1, &H21)
MF_SD_PROTECTED = iid
End Function
Public Function MF_SD_STREAM_NAME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4F1B099D, &HD314, &H41E5, &HA7, &H81, &H7F, &HEF, &HAA, &H4C, &H50, &H1F)
MF_SD_STREAM_NAME = iid
End Function
Public Function MF_SD_MUTUALLY_EXCLUSIVE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H23EF79C, &H388D, &H487F, &HAC, &H17, &H69, &H6C, &HD6, &HE3, &HC6, &HF5)
MF_SD_MUTUALLY_EXCLUSIVE = iid
End Function
Public Function MF_ACTIVATE_CUSTOM_VIDEO_MIXER_CLSID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBA491360, &HBE50, &H451E, &H95, &HAB, &H6D, &H4A, &HCC, &HC7, &HDA, &HD8)
MF_ACTIVATE_CUSTOM_VIDEO_MIXER_CLSID = iid
End Function
Public Function MF_ACTIVATE_CUSTOM_VIDEO_MIXER_ACTIVATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBA491361, &HBE50, &H451E, &H95, &HAB, &H6D, &H4A, &HCC, &HC7, &HDA, &HD8)
MF_ACTIVATE_CUSTOM_VIDEO_MIXER_ACTIVATE = iid
End Function
Public Function MF_ACTIVATE_CUSTOM_VIDEO_MIXER_FLAGS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBA491362, &HBE50, &H451E, &H95, &HAB, &H6D, &H4A, &HCC, &HC7, &HDA, &HD8)
MF_ACTIVATE_CUSTOM_VIDEO_MIXER_FLAGS = iid
End Function
Public Function MF_ACTIVATE_CUSTOM_VIDEO_PRESENTER_CLSID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBA491364, &HBE50, &H451E, &H95, &HAB, &H6D, &H4A, &HCC, &HC7, &HDA, &HD8)
MF_ACTIVATE_CUSTOM_VIDEO_PRESENTER_CLSID = iid
End Function
Public Function MF_ACTIVATE_CUSTOM_VIDEO_PRESENTER_ACTIVATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBA491365, &HBE50, &H451E, &H95, &HAB, &H6D, &H4A, &HCC, &HC7, &HDA, &HD8)
MF_ACTIVATE_CUSTOM_VIDEO_PRESENTER_ACTIVATE = iid
End Function
Public Function MF_ACTIVATE_CUSTOM_VIDEO_PRESENTER_FLAGS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBA491366, &HBE50, &H451E, &H95, &HAB, &H6D, &H4A, &HCC, &HC7, &HDA, &HD8)
MF_ACTIVATE_CUSTOM_VIDEO_PRESENTER_FLAGS = iid
End Function
Public Function MF_ACTIVATE_MFT_LOCKED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC1F6093C, &H7F65, &H4FBD, &H9E, &H39, &H5F, &HAE, &HC3, &HC4, &HFB, &HD7)
MF_ACTIVATE_MFT_LOCKED = iid
End Function
Public Function MF_ACTIVATE_VIDEO_WINDOW() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9A2DBBDD, &HF57E, &H4162, &H82, &HB9, &H68, &H31, &H37, &H76, &H82, &HD3)
MF_ACTIVATE_VIDEO_WINDOW = iid
End Function
Public Function MF_AUDIO_RENDERER_ATTRIBUTE_FLAGS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEDE4B5E0, &HF805, &H4D6C, &H99, &HB3, &HDB, &H1, &HBF, &H95, &HDF, &HAB)
MF_AUDIO_RENDERER_ATTRIBUTE_FLAGS = iid
End Function
Public Function MF_AUDIO_RENDERER_ATTRIBUTE_SESSION_ID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEDE4B5E3, &HF805, &H4D6C, &H99, &HB3, &HDB, &H1, &HBF, &H95, &HDF, &HAB)
MF_AUDIO_RENDERER_ATTRIBUTE_SESSION_ID = iid
End Function
Public Function MF_AUDIO_RENDERER_ATTRIBUTE_ENDPOINT_ID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB10AAEC3, &HEF71, &H4CC3, &HB8, &H73, &H5, &HA9, &HA0, &H8B, &H9F, &H8E)
MF_AUDIO_RENDERER_ATTRIBUTE_ENDPOINT_ID = iid
End Function
Public Function MF_AUDIO_RENDERER_ATTRIBUTE_ENDPOINT_ROLE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6BA644FF, &H27C5, &H4D02, &H98, &H87, &HC2, &H86, &H19, &HFD, &HB9, &H1B)
MF_AUDIO_RENDERER_ATTRIBUTE_ENDPOINT_ROLE = iid
End Function
Public Function MF_AUDIO_RENDERER_ATTRIBUTE_STREAM_CATEGORY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA9770471, &H92EC, &H4DF4, &H94, &HFE, &H81, &HC3, &H6F, &HC, &H3A, &H7A)
MF_AUDIO_RENDERER_ATTRIBUTE_STREAM_CATEGORY = iid
End Function
Public Function MFENABLETYPE_WMDRMV1_LicenseAcquisition() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4FF6EEAF, &HB43, &H4797, &H9B, &H85, &HAB, &HF3, &H18, &H15, &HE7, &HB0)
MFENABLETYPE_WMDRMV1_LicenseAcquisition = iid
End Function
Public Function MFENABLETYPE_WMDRMV7_LicenseAcquisition() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3306DF, &H4A06, &H4884, &HA0, &H97, &HEF, &H6D, &H22, &HEC, &H84, &HA3)
MFENABLETYPE_WMDRMV7_LicenseAcquisition = iid
End Function
Public Function MFENABLETYPE_WMDRMV7_Individualization() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HACD2C84A, &HB303, &H4F65, &HBC, &H2C, &H2C, &H84, &H8D, &H1, &HA9, &H89)
MFENABLETYPE_WMDRMV7_Individualization = iid
End Function
Public Function MFENABLETYPE_MF_UpdateRevocationInformation() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE558B0B5, &HB3C4, &H44A0, &H92, &H4C, &H50, &HD1, &H78, &H93, &H23, &H85)
MFENABLETYPE_MF_UpdateRevocationInformation = iid
End Function
Public Function MFENABLETYPE_MF_UpdateUntrustedComponent() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9879F3D6, &HCEE2, &H48E6, &HB5, &H73, &H97, &H67, &HAB, &H17, &H2F, &H16)
MFENABLETYPE_MF_UpdateUntrustedComponent = iid
End Function
Public Function MFENABLETYPE_MF_RebootRequired() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6D4D3D4B, &HECE, &H4652, &H8B, &H3A, &HF2, &HD2, &H42, &H60, &HD8, &H87)
MFENABLETYPE_MF_RebootRequired = iid
End Function
Public Function MF_METADATA_PROVIDER_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDB214084, &H58A4, &H4D2E, &HB8, &H4F, &H6F, &H75, &H5B, &H2F, &H7A, &HD)
MF_METADATA_PROVIDER_SERVICE = iid
End Function
Public Function MF_PROPERTY_HANDLER_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA3FACE02, &H32B8, &H41DD, &H90, &HE7, &H5F, &HEF, &H7C, &H89, &H91, &HB5)
MF_PROPERTY_HANDLER_SERVICE = iid
End Function
Public Function MF_RATE_CONTROL_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H866FA297, &HB802, &H4BF8, &H9D, &HC9, &H5E, &H3B, &H6A, &H9F, &H53, &HC9)
MF_RATE_CONTROL_SERVICE = iid
End Function
Public Function MF_TIMECODE_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA0D502A7, &HEB3, &H4885, &HB1, &HB9, &H9F, &HEB, &HD, &H8, &H34, &H54)
MF_TIMECODE_SERVICE = iid
End Function
Public Function MR_POLICY_VOLUME_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1ABAA2AC, &H9D3B, &H47C6, &HAB, &H48, &HC5, &H95, &H6, &HDE, &H78, &H4D)
MR_POLICY_VOLUME_SERVICE = iid
End Function
Public Function MR_CAPTURE_POLICY_VOLUME_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H24030ACD, &H107A, &H4265, &H97, &H5C, &H41, &H4E, &H33, &HE6, &H5F, &H2A)
MR_CAPTURE_POLICY_VOLUME_SERVICE = iid
End Function
Public Function MR_STREAM_VOLUME_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF8B5FA2F, &H32EF, &H46F5, &HB1, &H72, &H13, &H21, &H21, &H2F, &HB2, &HC4)
MR_STREAM_VOLUME_SERVICE = iid
End Function
Public Function MR_AUDIO_POLICY_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H911FD737, &H6775, &H4AB0, &HA6, &H14, &H29, &H78, &H62, &HFD, &HAC, &H88)
MR_AUDIO_POLICY_SERVICE = iid
End Function
Public Function MF_SAMPLEGRABBERSINK_SAMPLE_TIME_OFFSET() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H62E3D776, &H8100, &H4E03, &HA6, &HE8, &HBD, &H38, &H57, &HAC, &H9C, &H47)
MF_SAMPLEGRABBERSINK_SAMPLE_TIME_OFFSET = iid
End Function
Public Function MF_SAMPLEGRABBERSINK_IGNORE_CLOCK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEFDA2C0, &H2B69, &H4E2E, &HAB, &H8D, &H46, &HDC, &HBF, &HF7, &HD2, &H5D)
MF_SAMPLEGRABBERSINK_IGNORE_CLOCK = iid
End Function
Public Function MF_QUALITY_SERVICES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB7E2BE11, &H2F96, &H4640, &HB5, &H2C, &H28, &H23, &H65, &HBD, &HF1, &H6C)
MF_QUALITY_SERVICES = iid
End Function
Public Function MF_WORKQUEUE_SERVICES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8E37D489, &H41E0, &H413A, &H90, &H68, &H28, &H7C, &H88, &H6D, &H8D, &HDA)
MF_WORKQUEUE_SERVICES = iid
End Function
Public Function MF_QUALITY_NOTIFY_PROCESSING_LATENCY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF6B44AF8, &H604D, &H46FE, &HA9, &H5D, &H45, &H47, &H9B, &H10, &HC9, &HBC)
MF_QUALITY_NOTIFY_PROCESSING_LATENCY = iid
End Function
Public Function MF_QUALITY_NOTIFY_SAMPLE_LAG() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H30D15206, &HED2A, &H4760, &HBE, &H17, &HEB, &H4A, &H9F, &H12, &H29, &H5C)
MF_QUALITY_NOTIFY_SAMPLE_LAG = iid
End Function
Public Function MF_TIME_FORMAT_SEGMENT_OFFSET() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC8B8BE77, &H869C, &H431D, &H81, &H2E, &H16, &H96, &H93, &HF6, &H5A, &H39)
MF_TIME_FORMAT_SEGMENT_OFFSET = iid
End Function
Public Function MF_SOURCE_PRESENTATION_PROVIDER_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE002AADC, &HF4AF, &H4EE5, &H98, &H47, &H5, &H3E, &HDF, &H84, &H4, &H26)
MF_SOURCE_PRESENTATION_PROVIDER_SERVICE = iid
End Function
Public Function MF_TOPONODE_ATTRIBUTE_EDITOR_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H65656E1A, &H77F, &H4472, &H83, &HEF, &H31, &H6F, &H11, &HD5, &H8, &H7A)
MF_TOPONODE_ATTRIBUTE_EDITOR_SERVICE = iid
End Function
Public Function MFNETSOURCE_SSLCERTIFICATE_MANAGER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H55E6CB27, &HE69B, &H4267, &H94, &HC, &H2D, &H7E, &HC5, &HBB, &H8A, &HF)
MFNETSOURCE_SSLCERTIFICATE_MANAGER = iid
End Function
Public Function MFNETSOURCE_RESOURCE_FILTER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H815D0FF6, &H265A, &H4477, &H9E, &H46, &H7B, &H80, &HAD, &H80, &HB5, &HFB)
MFNETSOURCE_RESOURCE_FILTER = iid
End Function
Public Function MFNET_SAVEJOB_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB85A587F, &H3D02, &H4E52, &H95, &H65, &H55, &HD3, &HEC, &H1E, &H7F, &HF7)
MFNET_SAVEJOB_SERVICE = iid
End Function
Public Function MFNETSOURCE_STATISTICS_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F275, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_STATISTICS_SERVICE = iid
End Function
Public Function MFNETSOURCE_STATISTICS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F274, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_STATISTICS = iid
End Function
Public Function MFNETSOURCE_BUFFERINGTIME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F276, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_BUFFERINGTIME = iid
End Function
Public Function MFNETSOURCE_ACCELERATEDSTREAMINGDURATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F277, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_ACCELERATEDSTREAMINGDURATION = iid
End Function
Public Function MFNETSOURCE_MAXUDPACCELERATEDSTREAMINGDURATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4AAB2879, &HBBE1, &H4994, &H9F, &HF0, &H54, &H95, &HBD, &H25, &H1, &H29)
MFNETSOURCE_MAXUDPACCELERATEDSTREAMINGDURATION = iid
End Function
Public Function MFNETSOURCE_MAXBUFFERTIMEMS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H408B24E6, &H4038, &H4401, &HB5, &HB2, &HFE, &H70, &H1A, &H9E, &HBF, &H10)
MFNETSOURCE_MAXBUFFERTIMEMS = iid
End Function
Public Function MFNETSOURCE_CONNECTIONBANDWIDTH() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F278, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_CONNECTIONBANDWIDTH = iid
End Function
Public Function MFNETSOURCE_CACHEENABLED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F279, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_CACHEENABLED = iid
End Function
Public Function MFNETSOURCE_AUTORECONNECTLIMIT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F27A, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_AUTORECONNECTLIMIT = iid
End Function
Public Function MFNETSOURCE_RESENDSENABLED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F27B, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_RESENDSENABLED = iid
End Function
Public Function MFNETSOURCE_THINNINGENABLED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F27C, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_THINNINGENABLED = iid
End Function
Public Function MFNETSOURCE_PROTOCOL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F27D, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_PROTOCOL = iid
End Function
Public Function MFNETSOURCE_TRANSPORT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F27E, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_TRANSPORT = iid
End Function
Public Function MFNETSOURCE_PREVIEWMODEENABLED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F27F, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_PREVIEWMODEENABLED = iid
End Function
Public Function MFNETSOURCE_CREDENTIAL_MANAGER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F280, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_CREDENTIAL_MANAGER = iid
End Function
Public Function MFNETSOURCE_PPBANDWIDTH() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F281, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_PPBANDWIDTH = iid
End Function
Public Function MFNETSOURCE_AUTORECONNECTPROGRESS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F282, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_AUTORECONNECTPROGRESS = iid
End Function
Public Function MFNETSOURCE_PROXYLOCATORFACTORY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F283, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_PROXYLOCATORFACTORY = iid
End Function
Public Function MFNETSOURCE_BROWSERUSERAGENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F28B, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_BROWSERUSERAGENT = iid
End Function
Public Function MFNETSOURCE_BROWSERWEBPAGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F28C, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_BROWSERWEBPAGE = iid
End Function
Public Function MFNETSOURCE_PLAYERVERSION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F28D, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_PLAYERVERSION = iid
End Function
Public Function MFNETSOURCE_PLAYERID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F28E, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_PLAYERID = iid
End Function
Public Function MFNETSOURCE_HOSTEXE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F28F, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_HOSTEXE = iid
End Function
Public Function MFNETSOURCE_HOSTVERSION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F291, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_HOSTVERSION = iid
End Function
Public Function MFNETSOURCE_PLAYERUSERAGENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F292, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_PLAYERUSERAGENT = iid
End Function
Public Function MFNETSOURCE_CLIENTGUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H60A2C4A6, &HF197, &H4C14, &HA5, &HBF, &H88, &H83, &HD, &H24, &H58, &HAF)
MFNETSOURCE_CLIENTGUID = iid
End Function
Public Function MFNETSOURCE_LOGURL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F293, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_LOGURL = iid
End Function
Public Function MFNETSOURCE_ENABLE_UDP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F294, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_ENABLE_UDP = iid
End Function
Public Function MFNETSOURCE_ENABLE_TCP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F295, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_ENABLE_TCP = iid
End Function
Public Function MFNETSOURCE_ENABLE_MSB() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F296, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_ENABLE_MSB = iid
End Function
Public Function MFNETSOURCE_ENABLE_RTSP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F298, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_ENABLE_RTSP = iid
End Function
Public Function MFNETSOURCE_ENABLE_HTTP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F299, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_ENABLE_HTTP = iid
End Function
Public Function MFNETSOURCE_ENABLE_STREAMING() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F29C, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_ENABLE_STREAMING = iid
End Function
Public Function MFNETSOURCE_ENABLE_DOWNLOAD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F29D, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_ENABLE_DOWNLOAD = iid
End Function
Public Function MFNETSOURCE_ENABLE_PRIVATEMODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H824779D8, &HF18B, &H4405, &H8C, &HF1, &H46, &H4F, &HB5, &HAA, &H8F, &H71)
MFNETSOURCE_ENABLE_PRIVATEMODE = iid
End Function
Public Function MFNETSOURCE_UDP_PORT_RANGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F29A, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_UDP_PORT_RANGE = iid
End Function
Public Function MFNETSOURCE_PROXYINFO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F29B, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_PROXYINFO = iid
End Function
Public Function MFNETSOURCE_DRMNET_LICENSE_REPRESENTATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H47EAE1BD, &HBDFE, &H42E2, &H82, &HF3, &H54, &HA4, &H8C, &H17, &H96, &H2D)
MFNETSOURCE_DRMNET_LICENSE_REPRESENTATION = iid
End Function
Public Function MFNETSOURCE_PROXYSETTINGS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F287, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_PROXYSETTINGS = iid
End Function
Public Function MFNETSOURCE_PROXYHOSTNAME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F284, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_PROXYHOSTNAME = iid
End Function
Public Function MFNETSOURCE_PROXYPORT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F288, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_PROXYPORT = iid
End Function
Public Function MFNETSOURCE_PROXYEXCEPTIONLIST() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F285, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_PROXYEXCEPTIONLIST = iid
End Function
Public Function MFNETSOURCE_PROXYBYPASSFORLOCAL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F286, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_PROXYBYPASSFORLOCAL = iid
End Function
Public Function MFNETSOURCE_PROXYRERUNAUTODETECTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CB1F289, &H505, &H4C5D, &HAE, &H71, &HA, &H55, &H63, &H44, &HEF, &HA1)
MFNETSOURCE_PROXYRERUNAUTODETECTION = iid
End Function
Public Function MFNETSOURCE_STREAM_LANGUAGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9AB44318, &HF7CD, &H4F2D, &H8D, &H6D, &HFA, &H35, &HB4, &H92, &HCE, &HCB)
MFNETSOURCE_STREAM_LANGUAGE = iid
End Function
Public Function MFNETSOURCE_LOGPARAMS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H64936AE8, &H9418, &H453A, &H8C, &HDA, &H3E, &HA, &H66, &H8B, &H35, &H3B)
MFNETSOURCE_LOGPARAMS = iid
End Function
Public Function MFNETSOURCE_PEERMANAGER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H48B29ADB, &HFEBF, &H45EE, &HA9, &HBF, &HEF, &HB8, &H1C, &H49, &H2E, &HFC)
MFNETSOURCE_PEERMANAGER = iid
End Function
Public Function MFNETSOURCE_FRIENDLYNAME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5B2A7757, &HBC6B, &H447E, &HAA, &H6, &HD, &HDA, &H1C, &H64, &H6E, &H2F)
MFNETSOURCE_FRIENDLYNAME = iid
End Function
Public Function MF_BYTESTREAMHANDLER_ACCEPTS_SHARE_WRITE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA6E1F733, &H3001, &H4915, &H81, &H50, &H15, &H58, &HA2, &H18, &HE, &HC8)
MF_BYTESTREAMHANDLER_ACCEPTS_SHARE_WRITE = iid
End Function
Public Function MF_BYTESTREAM_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAB025E2B, &H16D9, &H4180, &HA1, &H27, &HBA, &H6C, &H70, &H15, &H61, &H61)
MF_BYTESTREAM_SERVICE = iid
End Function
Public Function MF_MEDIA_PROTECTION_MANAGER_PROPERTIES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H38BD81A9, &HACEA, &H4C73, &H89, &HB2, &H55, &H32, &HC0, &HAE, &HCA, &H79)
MF_MEDIA_PROTECTION_MANAGER_PROPERTIES = iid
End Function
Public Function MFCONNECTOR_UNKNOWN() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAC3AEF5C, &HCE43, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_UNKNOWN = iid
End Function
Public Function MFCONNECTOR_PCI() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAC3AEF5D, &HCE43, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_PCI = iid
End Function
Public Function MFCONNECTOR_PCIX() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAC3AEF5E, &HCE43, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_PCIX = iid
End Function
Public Function MFCONNECTOR_PCI_Express() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAC3AEF5F, &HCE43, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_PCI_Express = iid
End Function
Public Function MFCONNECTOR_AGP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAC3AEF60, &HCE43, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_AGP = iid
End Function
Public Function MFCONNECTOR_VGA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57CD5968, &HCE47, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_VGA = iid
End Function
Public Function MFCONNECTOR_SVIDEO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57CD5969, &HCE47, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_SVIDEO = iid
End Function
Public Function MFCONNECTOR_COMPOSITE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57CD596A, &HCE47, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_COMPOSITE = iid
End Function
Public Function MFCONNECTOR_COMPONENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57CD596B, &HCE47, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_COMPONENT = iid
End Function
Public Function MFCONNECTOR_DVI() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57CD596C, &HCE47, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_DVI = iid
End Function
Public Function MFCONNECTOR_HDMI() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57CD596D, &HCE47, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_HDMI = iid
End Function
Public Function MFCONNECTOR_LVDS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57CD596E, &HCE47, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_LVDS = iid
End Function
Public Function MFCONNECTOR_D_JPN() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57CD5970, &HCE47, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_D_JPN = iid
End Function
Public Function MFCONNECTOR_SDI() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57CD5971, &HCE47, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_SDI = iid
End Function
Public Function MFCONNECTOR_DISPLAYPORT_EXTERNAL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57CD5972, &HCE47, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_DISPLAYPORT_EXTERNAL = iid
End Function
Public Function MFCONNECTOR_DISPLAYPORT_EMBEDDED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57CD5973, &HCE47, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_DISPLAYPORT_EMBEDDED = iid
End Function
Public Function MFCONNECTOR_UDI_EXTERNAL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57CD5974, &HCE47, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_UDI_EXTERNAL = iid
End Function
Public Function MFCONNECTOR_UDI_EMBEDDED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57CD5975, &HCE47, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_UDI_EMBEDDED = iid
End Function
Public Function MFCONNECTOR_MIRACAST() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57CD5977, &HCE47, &H11D9, &H92, &HDB, &H0, &HB, &HDB, &H28, &HFF, &H98)
MFCONNECTOR_MIRACAST = iid
End Function
Public Function MFPROTECTION_DISABLE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8CC6D81B, &HFEC6, &H4D8F, &H96, &H4B, &HCF, &HBA, &HB, &HD, &HAD, &HD)
MFPROTECTION_DISABLE = iid
End Function
Public Function MFPROTECTION_CONSTRICTVIDEO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H193370CE, &HC5E4, &H4C3A, &H8A, &H66, &H69, &H59, &HB4, &HDA, &H44, &H42)
MFPROTECTION_CONSTRICTVIDEO = iid
End Function
Public Function MFPROTECTION_CONSTRICTVIDEO_NOOPM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA580E8CD, &HC247, &H4957, &HB9, &H83, &H3C, &H2E, &HEB, &HD1, &HFF, &H59)
MFPROTECTION_CONSTRICTVIDEO_NOOPM = iid
End Function
Public Function MFPROTECTION_CONSTRICTAUDIO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFFC99B44, &HDF48, &H4E16, &H8E, &H66, &H9, &H68, &H92, &HC1, &H57, &H8A)
MFPROTECTION_CONSTRICTAUDIO = iid
End Function
Public Function MFPROTECTION_TRUSTEDAUDIODRIVERS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H65BDF3D2, &H168, &H4816, &HA5, &H33, &H55, &HD4, &H7B, &H2, &H71, &H1)
MFPROTECTION_TRUSTEDAUDIODRIVERS = iid
End Function
Public Function MFPROTECTION_HDCP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAE7CC03D, &HC828, &H4021, &HAC, &HB7, &HD5, &H78, &HD2, &H7A, &HAF, &H13)
MFPROTECTION_HDCP = iid
End Function
Public Function MFPROTECTION_CGMSA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE57E69E9, &H226B, &H4D31, &HB4, &HE3, &HD3, &HDB, &H0, &H87, &H36, &HDD)
MFPROTECTION_CGMSA = iid
End Function
Public Function MFPROTECTION_ACP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC3FD11C6, &HF8B7, &H4D20, &HB0, &H8, &H1D, &HB1, &H7D, &H61, &HF2, &HDA)
MFPROTECTION_ACP = iid
End Function
Public Function MFPROTECTION_WMDRMOTA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA267A6A1, &H362E, &H47D0, &H88, &H5, &H46, &H28, &H59, &H8A, &H23, &HE4)
MFPROTECTION_WMDRMOTA = iid
End Function
Public Function MFPROTECTION_FFT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H462A56B2, &H2866, &H4BB6, &H98, &HD, &H6D, &H8D, &H9E, &HDB, &H1A, &H8C)
MFPROTECTION_FFT = iid
End Function
Public Function MFPROTECTION_PROTECTED_SURFACE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4F5D9566, &HE742, &H4A25, &H8D, &H1F, &HD2, &H87, &HB5, &HFA, &HA, &HDE)
MFPROTECTION_PROTECTED_SURFACE = iid
End Function
Public Function MFPROTECTION_DISABLE_SCREEN_SCRAPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA21179A4, &HB7CD, &H40D8, &H96, &H14, &H8E, &HF2, &H37, &H1B, &HA7, &H8D)
MFPROTECTION_DISABLE_SCREEN_SCRAPE = iid
End Function
Public Function MFPROTECTION_VIDEO_FRAMES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H36A59CBC, &H7401, &H4A8C, &HBC, &H20, &H46, &HA7, &HC9, &HE5, &H97, &HF0)
MFPROTECTION_VIDEO_FRAMES = iid
End Function
Public Function MFPROTECTION_HARDWARE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4EE7F0C1, &H9ED7, &H424F, &HB6, &HBE, &H99, &H6B, &H33, &H52, &H88, &H56)
MFPROTECTION_HARDWARE = iid
End Function
Public Function MFPROTECTION_HDCP_WITH_TYPE_ENFORCEMENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA4A585E8, &HED60, &H442D, &H81, &H4D, &HDB, &H4D, &H42, &H20, &HA0, &H6D)
MFPROTECTION_HDCP_WITH_TYPE_ENFORCEMENT = iid
End Function
Public Function MFPROTECTIONATTRIBUTE_BEST_EFFORT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC8E06331, &H75F0, &H4EC1, &H8E, &H77, &H17, &H57, &H8F, &H77, &H3B, &H46)
MFPROTECTIONATTRIBUTE_BEST_EFFORT = iid
End Function
Public Function MFPROTECTIONATTRIBUTE_FAIL_OVER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8536ABC5, &H38F1, &H4151, &H9C, &HCE, &HF5, &H5D, &H94, &H12, &H29, &HAC)
MFPROTECTIONATTRIBUTE_FAIL_OVER = iid
End Function
Public Function MFPROTECTION_GRAPHICS_TRANSFER_AES_ENCRYPTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC873DE64, &HD8A5, &H49E6, &H88, &HBB, &HFB, &H96, &H3F, &HD3, &HD4, &HCE)
MFPROTECTION_GRAPHICS_TRANSFER_AES_ENCRYPTION = iid
End Function
Public Function MFPROTECTIONATTRIBUTE_CONSTRICTVIDEO_IMAGESIZE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8476FC, &H4B58, &H4D80, &HA7, &H90, &HE7, &H29, &H76, &H73, &H16, &H1D)
MFPROTECTIONATTRIBUTE_CONSTRICTVIDEO_IMAGESIZE = iid
End Function
Public Function MFPROTECTIONATTRIBUTE_HDCP_SRM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6F302107, &H3477, &H4468, &H8A, &H8, &HEE, &HF9, &HDB, &H10, &HE2, &HF)
MFPROTECTIONATTRIBUTE_HDCP_SRM = iid
End Function
Public Function MF_SampleProtectionSalt() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5403DEEE, &HB9EE, &H438F, &HAA, &H83, &H38, &H4, &H99, &H7E, &H56, &H9D)
MF_SampleProtectionSalt = iid
End Function
Public Function MF_REMOTE_PROXY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2F00C90E, &HD2CF, &H4278, &H8B, &H6A, &HD0, &H77, &HFA, &HC3, &HA2, &H5F)
MF_REMOTE_PROXY = iid
End Function
Public Function CLSID_CreateMediaExtensionObject() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEF65A54D, &H788, &H45B8, &H8B, &H14, &HBC, &HF, &H6A, &H6B, &H51, &H37)
CLSID_CreateMediaExtensionObject = iid
End Function
Public Function MF_SAMI_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H49A89AE7, &HB4D9, &H4EF2, &HAA, &H5C, &HF6, &H5A, &H3E, &H5, &HAE, &H4E)
MF_SAMI_SERVICE = iid
End Function
Public Function MF_PD_SAMI_STYLELIST() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE0B73C7F, &H486D, &H484E, &H98, &H72, &H4D, &HE5, &H19, &H2A, &H7B, &HF8)
MF_PD_SAMI_STYLELIST = iid
End Function
Public Function MF_SD_SAMI_LANGUAGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H36FCB98A, &H6CD0, &H44CB, &HAC, &HB9, &HA8, &HF5, &H60, &HD, &HD0, &HBB)
MF_SD_SAMI_LANGUAGE = iid
End Function
Public Function MF_TRANSCODE_CONTAINERTYPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H150FF23F, &H4ABC, &H478B, &HAC, &H4F, &HE1, &H91, &H6F, &HBA, &H1C, &HCA)
MF_TRANSCODE_CONTAINERTYPE = iid
End Function
Public Function MFTranscodeContainerType_ASF() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H430F6F6E, &HB6BF, &H4FC1, &HA0, &HBD, &H9E, &HE4, &H6E, &HEE, &H2A, &HFB)
MFTranscodeContainerType_ASF = iid
End Function
Public Function MFTranscodeContainerType_MPEG4() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDC6CD05D, &HB9D0, &H40EF, &HBD, &H35, &HFA, &H62, &H2C, &H1A, &HB2, &H8A)
MFTranscodeContainerType_MPEG4 = iid
End Function
Public Function MFTranscodeContainerType_MP3() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE438B912, &H83F1, &H4DE6, &H9E, &H3A, &H9F, &HFB, &HC6, &HDD, &H24, &HD1)
MFTranscodeContainerType_MP3 = iid
End Function
Public Function MFTranscodeContainerType_FLAC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H31344AA3, &H5A9, &H42B5, &H90, &H1B, &H8E, &H9D, &H42, &H57, &HF7, &H5E)
MFTranscodeContainerType_FLAC = iid
End Function
Public Function MFTranscodeContainerType_3GP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H34C50167, &H4472, &H4F34, &H9E, &HA0, &HC4, &H9F, &HBA, &HCF, &H3, &H7D)
MFTranscodeContainerType_3GP = iid
End Function
Public Function MFTranscodeContainerType_AC3() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6D8D91C3, &H8C91, &H4ED1, &H87, &H42, &H8C, &H34, &H7D, &H5B, &H44, &HD0)
MFTranscodeContainerType_AC3 = iid
End Function
Public Function MFTranscodeContainerType_ADTS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H132FD27D, &HF02, &H43DE, &HA3, &H1, &H38, &HFB, &HBB, &HB3, &H83, &H4E)
MFTranscodeContainerType_ADTS = iid
End Function
Public Function MFTranscodeContainerType_MPEG2() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBFC2DBF9, &H7BB4, &H4F8F, &HAF, &HDE, &HE1, &H12, &HC4, &H4B, &HA8, &H82)
MFTranscodeContainerType_MPEG2 = iid
End Function
Public Function MFTranscodeContainerType_WAVE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H64C3453C, &HF26, &H4741, &HBE, &H63, &H87, &HBD, &HF8, &HBB, &H93, &H5B)
MFTranscodeContainerType_WAVE = iid
End Function
Public Function MFTranscodeContainerType_AVI() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7EDFE8AF, &H402F, &H4D76, &HA3, &H3C, &H61, &H9F, &HD1, &H57, &HD0, &HF1)
MFTranscodeContainerType_AVI = iid
End Function
Public Function MFTranscodeContainerType_FMPEG4() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9BA876F1, &H419F, &H4B77, &HA1, &HE0, &H35, &H95, &H9D, &H9D, &H40, &H4)
MFTranscodeContainerType_FMPEG4 = iid
End Function
Public Function MFTranscodeContainerType_AMR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H25D5AD3, &H621A, &H475B, &H96, &H4D, &H66, &HB1, &HC8, &H24, &HF0, &H79)
MFTranscodeContainerType_AMR = iid
End Function
Public Function MF_TRANSCODE_SKIP_METADATA_TRANSFER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4E4469EF, &HB571, &H4959, &H8F, &H83, &H3D, &HCF, &HBA, &H33, &HA3, &H93)
MF_TRANSCODE_SKIP_METADATA_TRANSFER = iid
End Function
Public Function MF_TRANSCODE_TOPOLOGYMODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3E3DF610, &H394A, &H40B2, &H9D, &HEA, &H3B, &HAB, &H65, &HB, &HEB, &HF2)
MF_TRANSCODE_TOPOLOGYMODE = iid
End Function
Public Function MF_TRANSCODE_ADJUST_PROFILE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9C37C21B, &H60F, &H487C, &HA6, &H90, &H80, &HD7, &HF5, &HD, &H1C, &H72)
MF_TRANSCODE_ADJUST_PROFILE = iid
End Function
Public Function MF_TRANSCODE_ENCODINGPROFILE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6947787C, &HF508, &H4EA9, &HB1, &HE9, &HA1, &HFE, &H3A, &H49, &HFB, &HC9)
MF_TRANSCODE_ENCODINGPROFILE = iid
End Function
Public Function MF_TRANSCODE_QUALITYVSSPEED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H98332DF8, &H3CD, &H476B, &H89, &HFA, &H3F, &H9E, &H44, &H2D, &HEC, &H9F)
MF_TRANSCODE_QUALITYVSSPEED = iid
End Function
Public Function MF_TRANSCODE_DONOT_INSERT_ENCODER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF45AA7CE, &HAB24, &H4012, &HA1, &H1B, &HDC, &H82, &H20, &H20, &H14, &H10)
MF_TRANSCODE_DONOT_INSERT_ENCODER = iid
End Function
Public Function MF_VIDEO_PROCESSOR_ALGORITHM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4A0A1E1F, &H272C, &H4FB6, &H9E, &HB1, &HDB, &H33, &HC, &HBC, &H97, &HCA)
MF_VIDEO_PROCESSOR_ALGORITHM = iid
End Function
Public Function MF_XVP_DISABLE_FRC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2C0AFA19, &H7A97, &H4D5A, &H9E, &HE8, &H16, &HD4, &HFC, &H51, &H8D, &H8C)
MF_XVP_DISABLE_FRC = iid
End Function
Public Function MF_XVP_CALLER_ALLOCATES_OUTPUT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4A2CABC, &HCAB, &H40B1, &HA1, &HB9, &H75, &HBC, &H36, &H58, &HF0, &H0)
MF_XVP_CALLER_ALLOCATES_OUTPUT = iid
End Function
Public Function CLSID_VideoProcessorMFT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H88753B26, &H5B24, &H49BD, &HB2, &HE7, &HC, &H44, &H5C, &H78, &HC9, &H82)
CLSID_VideoProcessorMFT = iid
End Function
Public Function MF_LOCAL_MFT_REGISTRATION_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDDF5CF9C, &H4506, &H45AA, &HAB, &HF0, &H6D, &H5D, &H94, &HDD, &H1B, &H4A)
MF_LOCAL_MFT_REGISTRATION_SERVICE = iid
End Function
Public Function MF_WRAPPED_SAMPLE_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H31F52BF2, &HD03E, &H4048, &H80, &HD0, &H9C, &H10, &H46, &HD8, &H7C, &H61)
MF_WRAPPED_SAMPLE_SERVICE = iid
End Function
Public Function MF_WRAPPED_OBJECT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2B182C4C, &HD6AC, &H49F4, &H89, &H15, &HF7, &H18, &H87, &HDB, &H70, &HCD)
MF_WRAPPED_OBJECT = iid
End Function
Public Function CLSID_HttpSchemePlugin() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H44CB442B, &H9DA9, &H49DF, &HB3, &HFD, &H2, &H37, &H77, &HB1, &H6E, &H50)
CLSID_HttpSchemePlugin = iid
End Function
Public Function CLSID_UrlmonSchemePlugin() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9EC4B4F9, &H3029, &H45AD, &H94, &H7B, &H34, &H4D, &HE2, &HA2, &H49, &HE2)
CLSID_UrlmonSchemePlugin = iid
End Function
Public Function CLSID_NetSchemePlugin() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE9F4EBAB, &HD97B, &H463E, &HA2, &HB1, &HC5, &H4E, &HE3, &HF9, &H41, &H4D)
CLSID_NetSchemePlugin = iid
End Function
Public Function MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC60AC5FE, &H252A, &H478F, &HA0, &HEF, &HBC, &H8F, &HA5, &HF7, &HCA, &HD3)
MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE = iid
End Function
Public Function MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_VIDCAP_HW_SOURCE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDE7046BA, &H54D6, &H4487, &HA2, &HA4, &HEC, &H7C, &HD, &H1B, &HD1, &H63)
MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_VIDCAP_HW_SOURCE = iid
End Function
Public Function MF_DEVSOURCE_ATTRIBUTE_FRIENDLY_NAME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H60D0E559, &H52F8, &H4FA2, &HBB, &HCE, &HAC, &HDB, &H34, &HA8, &HEC, &H1)
MF_DEVSOURCE_ATTRIBUTE_FRIENDLY_NAME = iid
End Function
Public Function MF_DEVSOURCE_ATTRIBUTE_MEDIA_TYPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H56A819CA, &HC78, &H4DE4, &HA0, &HA7, &H3D, &HDA, &HBA, &HF, &H24, &HD4)
MF_DEVSOURCE_ATTRIBUTE_MEDIA_TYPE = iid
End Function
Public Function MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_VIDCAP_CATEGORY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H77F0AE69, &HC3BD, &H4509, &H94, &H1D, &H46, &H7E, &H4D, &H24, &H89, &H9E)
MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_VIDCAP_CATEGORY = iid
End Function
Public Function MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_VIDCAP_SYMBOLIC_LINK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H58F0AAD8, &H22BF, &H4F8A, &HBB, &H3D, &HD2, &HC4, &H97, &H8C, &H6E, &H2F)
MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_VIDCAP_SYMBOLIC_LINK = iid
End Function
Public Function MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_AUDCAP_SYMBOLIC_LINK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H98D24B5E, &H5930, &H4614, &HB5, &HA1, &HF6, &H0, &HF9, &H35, &H5A, &H78)
MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_AUDCAP_SYMBOLIC_LINK = iid
End Function
Public Function MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_VIDCAP_MAX_BUFFERS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7DD9B730, &H4F2D, &H41D5, &H8F, &H95, &HC, &HC9, &HA9, &H12, &HBA, &H26)
MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_VIDCAP_MAX_BUFFERS = iid
End Function
Public Function MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_AUDCAP_ENDPOINT_ID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H30DA9258, &HFEB9, &H47A7, &HA4, &H53, &H76, &H3A, &H7A, &H8E, &H1C, &H5F)
MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_AUDCAP_ENDPOINT_ID = iid
End Function
Public Function MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_AUDCAP_ROLE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBC9D118E, &H8C67, &H4A18, &H85, &HD4, &H12, &HD3, &H0, &H40, &H5, &H52)
MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_AUDCAP_ROLE = iid
End Function
Public Function MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_AUDCAP_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H14DD9A1C, &H7CFF, &H41BE, &HB1, &HB9, &HBA, &H1A, &HC6, &HEC, &HB5, &H71)
MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_AUDCAP_GUID = iid
End Function
Public Function MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_VIDCAP_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8AC3587A, &H4AE7, &H42D8, &H99, &HE0, &HA, &H60, &H13, &HEE, &HF9, &HF)
MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_VIDCAP_GUID = iid
End Function
Public Function MF_DEVICESTREAM_IMAGE_STREAM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA7FFB865, &HE7B2, &H43B0, &H9F, &H6F, &H9A, &HF2, &HA0, &HE5, &HF, &HC0)
MF_DEVICESTREAM_IMAGE_STREAM = iid
End Function
Public Function MF_DEVICESTREAM_INDEPENDENT_IMAGE_STREAM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3EEEC7E, &HD605, &H4576, &H8B, &H29, &H65, &H80, &HB4, &H90, &HD7, &HD3)
MF_DEVICESTREAM_INDEPENDENT_IMAGE_STREAM = iid
End Function
Public Function MF_DEVICESTREAM_STREAM_ID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H11BD5120, &HD124, &H446B, &H88, &HE6, &H17, &H6, &H2, &H57, &HFF, &HF9)
MF_DEVICESTREAM_STREAM_ID = iid
End Function
Public Function MF_DEVICESTREAM_STREAM_CATEGORY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2939E7B8, &HA62E, &H4579, &HB6, &H74, &HD4, &H7, &H3D, &HFA, &HBB, &HBA)
MF_DEVICESTREAM_STREAM_CATEGORY = iid
End Function
Public Function MF_DEVICESTREAM_TRANSFORM_STREAM_ID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE63937B7, &HDAAF, &H4D49, &H81, &H5F, &HD8, &H26, &HF8, &HAD, &H31, &HE7)
MF_DEVICESTREAM_TRANSFORM_STREAM_ID = iid
End Function
Public Function MF_DEVICESTREAM_EXTENSION_PLUGIN_CLSID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H48E6558, &H60C4, &H4173, &HBD, &H5B, &H6A, &H3C, &HA2, &H89, &H6A, &HEE)
MF_DEVICESTREAM_EXTENSION_PLUGIN_CLSID = iid
End Function
Public Function MF_DEVICEMFT_EXTENSION_PLUGIN_CLSID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H844DBAE, &H34FA, &H48A0, &HA7, &H83, &H8E, &H69, &H6F, &HB1, &HC9, &HA8)
MF_DEVICEMFT_EXTENSION_PLUGIN_CLSID = iid
End Function
Public Function MF_DEVICESTREAM_EXTENSION_PLUGIN_CONNECTION_POINT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H37F9375C, &HE664, &H4EA4, &HAA, &HE4, &HCB, &H6D, &H1D, &HAC, &HA1, &HF4)
MF_DEVICESTREAM_EXTENSION_PLUGIN_CONNECTION_POINT = iid
End Function
Public Function MF_DEVICESTREAM_TAKEPHOTO_TRIGGER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1D180E34, &H538C, &H4FBB, &HA7, &H5A, &H85, &H9A, &HF7, &HD2, &H61, &HA6)
MF_DEVICESTREAM_TAKEPHOTO_TRIGGER = iid
End Function
Public Function MF_DEVICESTREAM_MAX_FRAME_BUFFERS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1684CEBE, &H3175, &H4985, &H88, &H2C, &HE, &HFD, &H3E, &H8A, &HC1, &H1E)
MF_DEVICESTREAM_MAX_FRAME_BUFFERS = iid
End Function
Public Function MF_DEVICEMFT_CONNECTED_FILTER_KSCONTROL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6A2C4FA6, &HD179, &H41CD, &H95, &H23, &H82, &H23, &H71, &HEA, &H40, &HE5)
MF_DEVICEMFT_CONNECTED_FILTER_KSCONTROL = iid
End Function
Public Function MF_DEVICEMFT_CONNECTED_PIN_KSCONTROL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE63310F7, &HB244, &H4EF8, &H9A, &H7D, &H24, &HC7, &H4E, &H32, &HEB, &HD0)
MF_DEVICEMFT_CONNECTED_PIN_KSCONTROL = iid
End Function
Public Function MF_DEVICE_THERMAL_STATE_CHANGED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H70CCD0AF, &HFC9F, &H4DEB, &HA8, &H75, &H9F, &HEC, &HD1, &H6C, &H5B, &HD4)
MF_DEVICE_THERMAL_STATE_CHANGED = iid
End Function
Public Function MFSampleExtension_DeviceTimestamp() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8F3E35E7, &H2DCD, &H4887, &H86, &H22, &H2A, &H58, &HBA, &HA6, &H52, &HB0)
MFSampleExtension_DeviceTimestamp = iid
End Function
Public Function MFSampleExtension_Spatial_CameraViewTransform() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4E251FA4, &H830F, &H4770, &H85, &H9A, &H4B, &H8D, &H99, &HAA, &H80, &H9B)
MFSampleExtension_Spatial_CameraViewTransform = iid
End Function
Public Function MFSampleExtension_Spatial_CameraCoordinateSystem() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9D13C82F, &H2199, &H4E67, &H91, &HCD, &HD1, &HA4, &H18, &H1F, &H25, &H34)
MFSampleExtension_Spatial_CameraCoordinateSystem = iid
End Function
Public Function MFSampleExtension_Spatial_CameraProjectionTransform() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H47F9FCB5, &H2A02, &H4F26, &HA4, &H77, &H79, &H2F, &HDF, &H95, &H88, &H6A)
MFSampleExtension_Spatial_CameraProjectionTransform = iid
End Function
Public Function CLSID_MPEG2ByteStreamPlugin() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H40871C59, &HAB40, &H471F, &H8D, &HC3, &H1F, &H25, &H9D, &H86, &H24, &H79)
CLSID_MPEG2ByteStreamPlugin = iid
End Function
Public Function MF_MEDIASOURCE_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF09992F7, &H9FBA, &H4C4A, &HA3, &H7F, &H8C, &H47, &HB4, &HE1, &HDF, &HE7)
MF_MEDIASOURCE_SERVICE = iid
End Function
Public Function MF_ACCESS_CONTROLLED_MEDIASOURCE_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H14A5031, &H2F05, &H4C6A, &H9F, &H9C, &H7D, &HD, &HC4, &HED, &HA5, &HF4)
MF_ACCESS_CONTROLLED_MEDIASOURCE_SERVICE = iid
End Function
Public Function MF_WRAPPED_BUFFER_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAB544072, &HC269, &H4EBC, &HA5, &H52, &H1C, &H3B, &H32, &HBE, &HD5, &HCA)
MF_WRAPPED_BUFFER_SERVICE = iid
End Function
Public Function MF_CONTENT_DECRYPTOR_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H68A72927, &HFC7B, &H44EE, &H85, &HF4, &H7C, &H51, &HBD, &H55, &HA6, &H59)
MF_CONTENT_DECRYPTOR_SERVICE = iid
End Function
Public Function MF_CONTENT_PROTECTION_DEVICE_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFF58436F, &H76A0, &H41FE, &HB5, &H66, &H10, &HCC, &H53, &H96, &H2E, &HDD)
MF_CONTENT_PROTECTION_DEVICE_SERVICE = iid
End Function
Public Function MF_SD_AUDIO_ENCODER_DELAY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8E85422C, &H73DE, &H403F, &H9A, &H35, &H55, &HA, &HD6, &HE8, &HB9, &H51)
MF_SD_AUDIO_ENCODER_DELAY = iid
End Function
Public Function MF_SD_AUDIO_ENCODER_PADDING() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H529C7F2C, &HAC4B, &H4E3F, &HBF, &HC3, &H9, &H2, &H19, &H49, &H82, &HCB)
MF_SD_AUDIO_ENCODER_PADDING = iid
End Function
Public Function CLSID_MSH264DecoderMFT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H62CE7E72, &H4C71, &H4D20, &HB1, &H5D, &H45, &H28, &H31, &HA8, &H7D, &H9D)
CLSID_MSH264DecoderMFT = iid
End Function
Public Function CLSID_MSH264EncoderMFT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6CA50344, &H51A, &H4DED, &H97, &H79, &HA4, &H33, &H5, &H16, &H5E, &H35)
CLSID_MSH264EncoderMFT = iid
End Function
Public Function CLSID_MSDDPlusDecMFT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H177C0AFE, &H900B, &H48D4, &H9E, &H4C, &H57, &HAD, &HD2, &H50, &HB3, &HD4)
CLSID_MSDDPlusDecMFT = iid
End Function
Public Function CLSID_MP3DecMediaObject() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBBEEA841, &HA63, &H4F52, &HA7, &HAB, &HA9, &HB3, &HA8, &H4E, &HD3, &H8A)
CLSID_MP3DecMediaObject = iid
End Function
Public Function CLSID_MSAACDecMFT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H32D186A7, &H218F, &H4C75, &H88, &H76, &HDD, &H77, &H27, &H3A, &H89, &H99)
CLSID_MSAACDecMFT = iid
End Function
Public Function CLSID_MSH265DecoderMFT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H420A51A3, &HD605, &H430C, &HB4, &HFC, &H45, &H27, &H4F, &HA6, &HC5, &H62)
CLSID_MSH265DecoderMFT = iid
End Function
Public Function CLSID_WMVDecoderMFT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H82D353DF, &H90BD, &H4382, &H8B, &HC2, &H3F, &H61, &H92, &HB7, &H6E, &H34)
CLSID_WMVDecoderMFT = iid
End Function
Public Function CLSID_WMADecMediaObject() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2EEB4ADF, &H4578, &H4D10, &HBC, &HA7, &HBB, &H95, &H5F, &H56, &H32, &HA)
CLSID_WMADecMediaObject = iid
End Function
Public Function CLSID_MSMPEGAudDecMFT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H70707B39, &HB2CA, &H4015, &HAB, &HEA, &HF8, &H44, &H7D, &H22, &HD8, &H8B)
CLSID_MSMPEGAudDecMFT = iid
End Function
Public Function CLSID_MSMPEGDecoderMFT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2D709E52, &H123F, &H49B5, &H9C, &HBC, &H9A, &HF5, &HCD, &HE2, &H8F, &HB9)
CLSID_MSMPEGDecoderMFT = iid
End Function
Public Function CLSID_AudioResamplerMediaObject() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF447B69E, &H1884, &H4A7E, &H80, &H55, &H34, &H6F, &H74, &HD6, &HED, &HB3)
CLSID_AudioResamplerMediaObject = iid
End Function
Public Function CLSID_MSVPxDecoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE3AAF548, &HC9A4, &H4C6E, &H23, &H4D, &H5A, &HDA, &H37, &H4B, &H0, &H0)
CLSID_MSVPxDecoder = iid
End Function
Public Function MF_D3D12_SYNCHRONIZATION_OBJECT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2A7C8D6A, &H85A6, &H494D, &HA0, &H46, &H6, &HEA, &H1A, &H13, &H8F, &H4B)
MF_D3D12_SYNCHRONIZATION_OBJECT = iid
End Function
Public Function MF_MT_D3D_RESOURCE_VERSION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H174F1E85, &HFE26, &H453D, &HB5, &H2E, &H5B, &HDD, &H4E, &H55, &HB9, &H44)
MF_MT_D3D_RESOURCE_VERSION = iid
End Function
Public Function MF_MT_D3D12_CPU_READBACK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H28EE9FE3, &HD481, &H46A6, &HB9, &H8A, &H7F, &H69, &HD5, &H28, &HE, &H82)
MF_MT_D3D12_CPU_READBACK = iid
End Function
Public Function MF_MT_D3D12_TEXTURE_LAYOUT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H97C85CAA, &HBEB, &H4EE1, &H97, &H15, &HF2, &H2F, &HAD, &H8C, &H10, &HF5)
MF_MT_D3D12_TEXTURE_LAYOUT = iid
End Function
Public Function MF_MT_D3D12_RESOURCE_FLAG_ALLOW_RENDER_TARGET() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEEAC2585, &H3430, &H498C, &H84, &HA2, &H77, &HB1, &HBB, &HA5, &H70, &HF6)
MF_MT_D3D12_RESOURCE_FLAG_ALLOW_RENDER_TARGET = iid
End Function
Public Function MF_MT_D3D12_RESOURCE_FLAG_ALLOW_DEPTH_STENCIL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB1138DC3, &H1D5, &H4C14, &H9B, &HDC, &HCD, &HC9, &H33, &H6F, &H55, &HB9)
MF_MT_D3D12_RESOURCE_FLAG_ALLOW_DEPTH_STENCIL = iid
End Function
Public Function MF_MT_D3D12_RESOURCE_FLAG_ALLOW_UNORDERED_ACCESS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H82C85647, &H5057, &H4960, &H95, &H59, &HF4, &H5B, &H8E, &H27, &H14, &H27)
MF_MT_D3D12_RESOURCE_FLAG_ALLOW_UNORDERED_ACCESS = iid
End Function
Public Function MF_MT_D3D12_RESOURCE_FLAG_DENY_SHADER_RESOURCE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBA06BFAC, &HFFE3, &H474A, &HAB, &H55, &H16, &H1E, &HE4, &H41, &H7A, &H2E)
MF_MT_D3D12_RESOURCE_FLAG_DENY_SHADER_RESOURCE = iid
End Function
Public Function MF_MT_D3D12_RESOURCE_FLAG_ALLOW_CROSS_ADAPTER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA6A1E439, &H2F96, &H4AB5, &H98, &HDC, &HAD, &HF7, &H49, &H73, &H50, &H5D)
MF_MT_D3D12_RESOURCE_FLAG_ALLOW_CROSS_ADAPTER = iid
End Function
Public Function MF_MT_D3D12_RESOURCE_FLAG_ALLOW_SIMULTANEOUS_ACCESS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA4940B2, &HCFD6, &H4738, &H9D, &H2, &H98, &H11, &H37, &H34, &H1, &H5A)
MF_MT_D3D12_RESOURCE_FLAG_ALLOW_SIMULTANEOUS_ACCESS = iid
End Function
Public Function MF_SA_D3D12_HEAP_FLAGS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H496B3266, &HD28F, &H4F8C, &H93, &HA7, &H4A, &H59, &H6B, &H1A, &H31, &HA1)
MF_SA_D3D12_HEAP_FLAGS = iid
End Function
Public Function MF_SA_D3D12_HEAP_TYPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H56F26A76, &HBBC1, &H4CE0, &HBB, &H11, &HE2, &H23, &H68, &HD8, &H74, &HED)
MF_SA_D3D12_HEAP_TYPE = iid
End Function
Public Function MF_SA_D3D12_CLEAR_VALUE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H86BA9A39, &H526, &H495D, &H9A, &HB5, &H54, &HEC, &H9F, &HAD, &H6F, &HC3)
MF_SA_D3D12_CLEAR_VALUE = iid
End Function
Public Function MF_CAPTURE_ENGINE_INITIALIZED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H219992BC, &HCF92, &H4531, &HA1, &HAE, &H96, &HE1, &HE8, &H86, &HC8, &HF1)
MF_CAPTURE_ENGINE_INITIALIZED = iid
End Function
Public Function MF_CAPTURE_ENGINE_PREVIEW_STARTED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA416DF21, &HF9D3, &H4A74, &H99, &H1B, &HB8, &H17, &H29, &H89, &H52, &HC4)
MF_CAPTURE_ENGINE_PREVIEW_STARTED = iid
End Function
Public Function MF_CAPTURE_ENGINE_PREVIEW_STOPPED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H13D5143C, &H1EDD, &H4E50, &HA2, &HEF, &H35, &HA, &H47, &H67, &H80, &H60)
MF_CAPTURE_ENGINE_PREVIEW_STOPPED = iid
End Function
Public Function MF_CAPTURE_ENGINE_RECORD_STARTED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAC2B027B, &HDDF9, &H48A0, &H89, &HBE, &H38, &HAB, &H35, &HEF, &H45, &HC0)
MF_CAPTURE_ENGINE_RECORD_STARTED = iid
End Function
Public Function MF_CAPTURE_ENGINE_RECORD_STOPPED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H55E5200A, &HF98F, &H4C0D, &HA9, &HEC, &H9E, &HB2, &H5E, &HD3, &HD7, &H73)
MF_CAPTURE_ENGINE_RECORD_STOPPED = iid
End Function
Public Function MF_CAPTURE_ENGINE_PHOTO_TAKEN() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3C50C445, &H7304, &H48EB, &H86, &H5D, &HBB, &HA1, &H9B, &HA3, &HAF, &H5C)
MF_CAPTURE_ENGINE_PHOTO_TAKEN = iid
End Function
Public Function MF_CAPTURE_SOURCE_CURRENT_DEVICE_MEDIA_TYPE_SET() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE7E75E4C, &H39C, &H4410, &H81, &H5B, &H87, &H41, &H30, &H7B, &H63, &HAA)
MF_CAPTURE_SOURCE_CURRENT_DEVICE_MEDIA_TYPE_SET = iid
End Function
Public Function MF_CAPTURE_ENGINE_ERROR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H46B89FC6, &H33CC, &H4399, &H9D, &HAD, &H78, &H4D, &HE7, &H7D, &H58, &H7C)
MF_CAPTURE_ENGINE_ERROR = iid
End Function
Public Function MF_CAPTURE_ENGINE_EFFECT_ADDED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAA8DC7B5, &HA048, &H4E13, &H8E, &HBE, &HF2, &H3C, &H46, &HC8, &H30, &HC1)
MF_CAPTURE_ENGINE_EFFECT_ADDED = iid
End Function
Public Function MF_CAPTURE_ENGINE_EFFECT_REMOVED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC6E8DB07, &HFB09, &H4A48, &H89, &HC6, &HBF, &H92, &HA0, &H42, &H22, &HC9)
MF_CAPTURE_ENGINE_EFFECT_REMOVED = iid
End Function
Public Function MF_CAPTURE_ENGINE_ALL_EFFECTS_REMOVED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFDED7521, &H8ED8, &H431A, &HA9, &H6B, &HF3, &HE2, &H56, &H5E, &H98, &H1C)
MF_CAPTURE_ENGINE_ALL_EFFECTS_REMOVED = iid
End Function
Public Function MF_CAPTURE_SINK_PREPARED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7BFCE257, &H12B1, &H4409, &H8C, &H34, &HD4, &H45, &HDA, &HAB, &H75, &H78)
MF_CAPTURE_SINK_PREPARED = iid
End Function
Public Function MF_CAPTURE_ENGINE_OUTPUT_MEDIA_TYPE_SET() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCAAAD994, &H83EC, &H45E9, &HA3, &HA, &H1F, &H20, &HAA, &HDB, &H98, &H31)
MF_CAPTURE_ENGINE_OUTPUT_MEDIA_TYPE_SET = iid
End Function
Public Function MF_CAPTURE_ENGINE_D3D_MANAGER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H76E25E7B, &HD595, &H4283, &H96, &H2C, &HC5, &H94, &HAF, &HD7, &H8D, &HDF)
MF_CAPTURE_ENGINE_D3D_MANAGER = iid
End Function
Public Function MF_CAPTURE_ENGINE_RECORD_SINK_VIDEO_MAX_UNPROCESSED_SAMPLES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB467F705, &H7913, &H4894, &H9D, &H42, &HA2, &H15, &HFE, &HA2, &H3D, &HA9)
MF_CAPTURE_ENGINE_RECORD_SINK_VIDEO_MAX_UNPROCESSED_SAMPLES = iid
End Function
Public Function MF_CAPTURE_ENGINE_RECORD_SINK_AUDIO_MAX_UNPROCESSED_SAMPLES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1CDDB141, &HA7F4, &H4D58, &H98, &H96, &H4D, &H15, &HA5, &H3C, &H4E, &HFE)
MF_CAPTURE_ENGINE_RECORD_SINK_AUDIO_MAX_UNPROCESSED_SAMPLES = iid
End Function
Public Function MF_CAPTURE_ENGINE_RECORD_SINK_VIDEO_MAX_PROCESSED_SAMPLES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE7B4A49E, &H382C, &H4AEF, &HA9, &H46, &HAE, &HD5, &H49, &HB, &H71, &H11)
MF_CAPTURE_ENGINE_RECORD_SINK_VIDEO_MAX_PROCESSED_SAMPLES = iid
End Function
Public Function MF_CAPTURE_ENGINE_RECORD_SINK_AUDIO_MAX_PROCESSED_SAMPLES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9896E12A, &HF707, &H4500, &HB6, &HBD, &HDB, &H8E, &HB8, &H10, &HB5, &HF)
MF_CAPTURE_ENGINE_RECORD_SINK_AUDIO_MAX_PROCESSED_SAMPLES = iid
End Function
Public Function MF_CAPTURE_ENGINE_USE_AUDIO_DEVICE_ONLY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1C8077DA, &H8466, &H4DC4, &H8B, &H8E, &H27, &H6B, &H3F, &H85, &H92, &H3B)
MF_CAPTURE_ENGINE_USE_AUDIO_DEVICE_ONLY = iid
End Function
Public Function MF_CAPTURE_ENGINE_USE_VIDEO_DEVICE_ONLY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7E025171, &HCF32, &H4F2E, &H8F, &H19, &H41, &H5, &H77, &HB7, &H3A, &H66)
MF_CAPTURE_ENGINE_USE_VIDEO_DEVICE_ONLY = iid
End Function
Public Function MF_CAPTURE_ENGINE_DISABLE_HARDWARE_TRANSFORMS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB7C42A6B, &H3207, &H4495, &HB4, &HE7, &H81, &HF9, &HC3, &H5D, &H59, &H91)
MF_CAPTURE_ENGINE_DISABLE_HARDWARE_TRANSFORMS = iid
End Function
Public Function MF_CAPTURE_ENGINE_DISABLE_DXVA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF9818862, &H179D, &H433F, &HA3, &H2F, &H74, &HCB, &HCF, &H74, &H46, &H6D)
MF_CAPTURE_ENGINE_DISABLE_DXVA = iid
End Function
Public Function MF_CAPTURE_ENGINE_MEDIASOURCE_CONFIG() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBC6989D2, &HFC1, &H46E1, &HA7, &H4F, &HEF, &HD3, &H6B, &HC7, &H88, &HDE)
MF_CAPTURE_ENGINE_MEDIASOURCE_CONFIG = iid
End Function
Public Function MF_CAPTURE_ENGINE_DECODER_MFT_FIELDOFUSE_UNLOCK_Attribute() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2B8AD2E8, &H7ACB, &H4321, &HA6, &H6, &H32, &H5C, &H42, &H49, &HF4, &HFC)
MF_CAPTURE_ENGINE_DECODER_MFT_FIELDOFUSE_UNLOCK_Attribute = iid
End Function
Public Function MF_CAPTURE_ENGINE_ENCODER_MFT_FIELDOFUSE_UNLOCK_Attribute() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H54C63A00, &H78D5, &H422F, &HAA, &H3E, &H5E, &H99, &HAC, &H64, &H92, &H69)
MF_CAPTURE_ENGINE_ENCODER_MFT_FIELDOFUSE_UNLOCK_Attribute = iid
End Function
Public Function MF_CAPTURE_ENGINE_EVENT_GENERATOR_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HABFA8AD5, &HFC6D, &H4911, &H87, &HE0, &H96, &H19, &H45, &HF8, &HF7, &HCE)
MF_CAPTURE_ENGINE_EVENT_GENERATOR_GUID = iid
End Function
Public Function MF_CAPTURE_ENGINE_EVENT_STREAM_INDEX() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H82697F44, &HB1CF, &H42EB, &H97, &H53, &HF8, &H6D, &H64, &H9C, &H88, &H65)
MF_CAPTURE_ENGINE_EVENT_STREAM_INDEX = iid
End Function
Public Function MF_CAPTURE_ENGINE_SELECTEDCAMERAPROFILE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3160B7E, &H1C6F, &H4DB2, &HAD, &H56, &HA7, &HC4, &H30, &HF8, &H23, &H92)
MF_CAPTURE_ENGINE_SELECTEDCAMERAPROFILE = iid
End Function
Public Function MF_CAPTURE_ENGINE_SELECTEDCAMERAPROFILE_INDEX() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CE88613, &H2214, &H46C3, &HB4, &H17, &H82, &HF8, &HA3, &H13, &HC9, &HC3)
MF_CAPTURE_ENGINE_SELECTEDCAMERAPROFILE_INDEX = iid
End Function
Public Function CLSID_MFCaptureEngine() As UUID
'{efce38d3-8914-4674-a7df-ae1b3d654b8a}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEFCE38D3, CInt(&H8914), CInt(&H4674), &HA7, &HDF, &HAE, &H1B, &H3D, &H65, &H4B, &H8A)
 CLSID_MFCaptureEngine = iid
End Function
Public Function CLSID_MFCaptureEngineClassFactory() As UUID
'{efce38d3-8914-4674-a7df-ae1b3d654b8a}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEFCE38D3, CInt(&H8914), CInt(&H4674), &HA7, &HDF, &HAE, &H1B, &H3D, &H65, &H4B, &H8A)
 CLSID_MFCaptureEngineClassFactory = iid
End Function
Public Function MFSampleExtension_DeviceReferenceSystemTime() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6523775A, &HBA2D, &H405F, &HB2, &HC5, &H1, &HFF, &H88, &HE2, &HE8, &HF6)
MFSampleExtension_DeviceReferenceSystemTime = iid
End Function
Public Function CLSID_MFReadWriteClassFactory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H48E2ED0F, &H98C2, &H4A37, &HBE, &HD5, &H16, &H63, &H12, &HDD, &HD8, &H3F)
CLSID_MFReadWriteClassFactory = iid
End Function
Public Function CLSID_MFSourceReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1777133C, &H881, &H411B, &HA5, &H77, &HAD, &H54, &H5F, &H7, &H14, &HC4)
CLSID_MFSourceReader = iid
End Function
Public Function MF_SOURCE_READER_ASYNC_CALLBACK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1E3DBEAC, &HBB43, &H4C35, &HB5, &H7, &HCD, &H64, &H44, &H64, &HC9, &H65)
 MF_SOURCE_READER_ASYNC_CALLBACK = iid
End Function
Public Function MF_SOURCE_READER_DISABLE_DXVA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAA456CFD, &H3943, &H4A1E, &HA7, &H7D, &H18, &H38, &HC0, &HEA, &H2E, &H35)
 MF_SOURCE_READER_DISABLE_DXVA = iid
End Function
Public Function MF_SOURCE_READER_MEDIASOURCE_CONFIG() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9085ABEB, &H354, &H48F9, &HAB, &HB5, &H20, &HD, &HF8, &H38, &HC6, &H8E)
 MF_SOURCE_READER_MEDIASOURCE_CONFIG = iid
End Function
Public Function MF_SOURCE_READER_MEDIASOURCE_CHARACTERISTICS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6D23F5C8, &HC5D7, &H4A9B, &H99, &H71, &H5D, &H11, &HF8, &HBC, &HA8, &H80)
 MF_SOURCE_READER_MEDIASOURCE_CHARACTERISTICS = iid
End Function
Public Function MF_SOURCE_READER_ENABLE_VIDEO_PROCESSING() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFB394F3D, &HCCF1, &H42EE, &HBB, &HB3, &HF9, &HB8, &H45, &HD5, &H68, &H1D)
 MF_SOURCE_READER_ENABLE_VIDEO_PROCESSING = iid
End Function
Public Function MF_SOURCE_READER_ENABLE_ADVANCED_VIDEO_PROCESSING() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF81DA2C, &HB537, &H4672, &HA8, &HB2, &HA6, &H81, &HB1, &H73, &H7, &HA3)
 MF_SOURCE_READER_ENABLE_ADVANCED_VIDEO_PROCESSING = iid
End Function
Public Function MF_SOURCE_READER_DISABLE_CAMERA_PLUGINS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9D3365DD, &H58F, &H4CFB, &H9F, &H97, &HB3, &H14, &HCC, &H99, &HC8, &HAD)
 MF_SOURCE_READER_DISABLE_CAMERA_PLUGINS = iid
End Function
Public Function MF_SOURCE_READER_DISCONNECT_MEDIASOURCE_ON_SHUTDOWN() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H56B67165, &H219E, &H456D, &HA2, &H2E, &H2D, &H30, &H4, &HC7, &HFE, &H56)
 MF_SOURCE_READER_DISCONNECT_MEDIASOURCE_ON_SHUTDOWN = iid
End Function
Public Function MF_SOURCE_READER_ENABLE_TRANSCODE_ONLY_TRANSFORMS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDFD4F008, &HB5FD, &H4E78, &HAE, &H44, &H62, &HA1, &HE6, &H7B, &HBE, &H27)
 MF_SOURCE_READER_ENABLE_TRANSCODE_ONLY_TRANSFORMS = iid
End Function
Public Function MF_SOURCE_READER_D3D11_BIND_FLAGS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H33F3197B, &HF73A, &H4E14, &H8D, &H85, &HE, &H4C, &H43, &H68, &H78, &H8D)
 MF_SOURCE_READER_D3D11_BIND_FLAGS = iid
End Function
Public Function CLSID_MFSinkWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA3BBFB17, &H8273, &H4E52, &H9E, &HE, &H97, &H39, &HDC, &H88, &H79, &H90)
CLSID_MFSinkWriter = iid
End Function
Public Function MF_SINK_WRITER_ASYNC_CALLBACK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H48CB183E, &H7B0B, &H46F4, &H82, &H2E, &H5E, &H1D, &H2D, &HDA, &H43, &H54)
 MF_SINK_WRITER_ASYNC_CALLBACK = iid
End Function
Public Function MF_SINK_WRITER_DISABLE_THROTTLING() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8B845D8, &H2B74, &H4AFE, &H9D, &H53, &HBE, &H16, &HD2, &HD5, &HAE, &H4F)
 MF_SINK_WRITER_DISABLE_THROTTLING = iid
End Function
Public Function MF_SINK_WRITER_D3D_MANAGER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEC822DA2, &HE1E9, &H4B29, &HA0, &HD8, &H56, &H3C, &H71, &H9F, &H52, &H69)
 MF_SINK_WRITER_D3D_MANAGER = iid
End Function
Public Function MF_SINK_WRITER_ENCODER_CONFIG() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAD91CD04, &HA7CC, &H4AC7, &H99, &HB6, &HA5, &H7B, &H9A, &H4A, &H7C, &H70)
 MF_SINK_WRITER_ENCODER_CONFIG = iid
End Function
Public Function MF_READWRITE_DISABLE_CONVERTERS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H98D5B065, &H1374, &H4847, &H8D, &H5D, &H31, &H52, &HF, &HEE, &H71, &H56)
 MF_READWRITE_DISABLE_CONVERTERS = iid
End Function
Public Function MF_READWRITE_ENABLE_HARDWARE_TRANSFORMS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA634A91C, &H822B, &H41B9, &HA4, &H94, &H4D, &HE4, &H64, &H36, &H12, &HB0)
 MF_READWRITE_ENABLE_HARDWARE_TRANSFORMS = iid
End Function
Public Function MF_READWRITE_MMCSS_CLASS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H39384300, &HD0EB, &H40B1, &H87, &HA0, &H33, &H18, &H87, &H1B, &H5A, &H53)
 MF_READWRITE_MMCSS_CLASS = iid
End Function
Public Function MF_READWRITE_MMCSS_PRIORITY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43AD19CE, &HF33F, &H4BA9, &HA5, &H80, &HE4, &HCD, &H12, &HF2, &HD1, &H44)
 MF_READWRITE_MMCSS_PRIORITY = iid
End Function
Public Function MF_READWRITE_MMCSS_CLASS_AUDIO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H430847DA, &H890, &H4B0E, &H93, &H8C, &H5, &H43, &H32, &HC5, &H47, &HE1)
 MF_READWRITE_MMCSS_CLASS_AUDIO = iid
End Function
Public Function MF_READWRITE_MMCSS_PRIORITY_AUDIO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H273DB885, &H2DE2, &H4DB2, &HA6, &HA7, &HFD, &HB6, &H6F, &HB4, &HB, &H61)
 MF_READWRITE_MMCSS_PRIORITY_AUDIO = iid
End Function
Public Function MF_READWRITE_D3D_OPTIONAL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H216479D9, &H3071, &H42CA, &HBB, &H6C, &H4C, &H22, &H10, &H2E, &H1D, &H18)
 MF_READWRITE_D3D_OPTIONAL = iid
End Function
Public Function MF_MEDIASINK_AUTOFINALIZE_SUPPORTED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H48C131BE, &H135A, &H41CB, &H82, &H90, &H3, &H65, &H25, &H9, &HC9, &H99)
 MF_MEDIASINK_AUTOFINALIZE_SUPPORTED = iid
End Function
Public Function MF_MEDIASINK_ENABLE_AUTOFINALIZE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H34014265, &HCB7E, &H4CDE, &HAC, &H7C, &HEF, &HFD, &H3B, &H3C, &H25, &H30)
 MF_MEDIASINK_ENABLE_AUTOFINALIZE = iid
End Function
Public Function MF_READWRITE_ENABLE_AUTOFINALIZE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDD7CA129, &H8CD1, &H4DC5, &H9D, &HDE, &HCE, &H16, &H86, &H75, &HDE, &H61)
 MF_READWRITE_ENABLE_AUTOFINALIZE = iid
End Function
Public Function MF_DMFT_FRAME_BUFFER_INFO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H396CE1C9, &H67A9, &H454C, &H87, &H97, &H95, &HA4, &H57, &H99, &HD8, &H4)
 MF_DMFT_FRAME_BUFFER_INFO = iid
End Function
Public Function MFT_AUDIO_DECODER_DEGRADATION_INFO_ATTRIBUTE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C3386AD, &HEC20, &H430D, &HB2, &HA5, &H50, &H5C, &H71, &H78, &HD9, &HC4)
 MFT_AUDIO_DECODER_DEGRADATION_INFO_ATTRIBUTE = iid
End Function
Public Function MF_MSE_CALLBACK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9063A7C0, &H42C5, &H4FFD, &HA8, &HA8, &H6F, &HCF, &H9E, &HA3, &HD0, &HC)
MF_MSE_CALLBACK = iid
End Function
Public Function MF_MSE_ACTIVELIST_CALLBACK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H949BDA0F, &H4549, &H46D5, &HAD, &H7F, &HB8, &H46, &HE1, &HAB, &H16, &H52)
MF_MSE_ACTIVELIST_CALLBACK = iid
End Function
Public Function MF_MSE_BUFFERLIST_CALLBACK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H42E669B0, &HD60E, &H4AFB, &HA8, &H5B, &HD8, &HE5, &HFE, &H6B, &HDA, &HB5)
MF_MSE_BUFFERLIST_CALLBACK = iid
End Function
Public Function MF_MSE_VP9_SUPPORT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H92D78429, &HD88B, &H4FF0, &H83, &H22, &H80, &H3E, &HFA, &H6E, &H96, &H26)
MF_MSE_VP9_SUPPORT = iid
End Function
Public Function MF_MSE_OPUS_SUPPORT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D224CC1, &H8CC4, &H48A3, &HA7, &HA7, &HE4, &HC1, &H6C, &HE6, &H38, &H8A)
MF_MSE_OPUS_SUPPORT = iid
End Function
Public Function MF_MEDIA_ENGINE_NEEDKEY_CALLBACK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7EA80843, &HB6E4, &H432C, &H8E, &HA4, &H78, &H48, &HFF, &HE4, &H22, &HE)
MF_MEDIA_ENGINE_NEEDKEY_CALLBACK = iid
End Function
Public Function MF_MEDIA_ENGINE_CALLBACK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC60381B8, &H83A4, &H41F8, &HA3, &HD0, &HDE, &H5, &H7, &H68, &H49, &HA9)
MF_MEDIA_ENGINE_CALLBACK = iid
End Function
Public Function MF_MEDIA_ENGINE_DXGI_MANAGER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H65702DA, &H1094, &H486D, &H86, &H17, &HEE, &H7C, &HC4, &HEE, &H46, &H48)
MF_MEDIA_ENGINE_DXGI_MANAGER = iid
End Function
Public Function MF_MEDIA_ENGINE_EXTENSION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3109FD46, &H60D, &H4B62, &H8D, &HCF, &HFA, &HFF, &H81, &H13, &H18, &HD2)
MF_MEDIA_ENGINE_EXTENSION = iid
End Function
Public Function MF_MEDIA_ENGINE_PLAYBACK_HWND() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD988879B, &H67C9, &H4D92, &HBA, &HA7, &H6E, &HAD, &HD4, &H46, &H3, &H9D)
MF_MEDIA_ENGINE_PLAYBACK_HWND = iid
End Function
Public Function MF_MEDIA_ENGINE_OPM_HWND() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA0BE8EE7, &H572, &H4F2C, &HA8, &H1, &H2A, &H15, &H1B, &HD3, &HE7, &H26)
MF_MEDIA_ENGINE_OPM_HWND = iid
End Function
Public Function MF_MEDIA_ENGINE_PLAYBACK_VISUAL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6DEBD26F, &H6AB9, &H4D7E, &HB0, &HEE, &HC6, &H1A, &H73, &HFF, &HAD, &H15)
MF_MEDIA_ENGINE_PLAYBACK_VISUAL = iid
End Function
Public Function MF_MEDIA_ENGINE_COREWINDOW() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFCCAE4DC, &HB7F, &H41C2, &H9F, &H96, &H46, &H59, &H94, &H8A, &HCD, &HDC)
MF_MEDIA_ENGINE_COREWINDOW = iid
End Function
Public Function MF_MEDIA_ENGINE_VIDEO_OUTPUT_FORMAT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5066893C, &H8CF9, &H42BC, &H8B, &H8A, &H47, &H22, &H12, &HE5, &H27, &H26)
MF_MEDIA_ENGINE_VIDEO_OUTPUT_FORMAT = iid
End Function
Public Function MF_MEDIA_ENGINE_CONTENT_PROTECTION_FLAGS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE0350223, &H5AAF, &H4D76, &HA7, &HC3, &H6, &HDE, &H70, &H89, &H4D, &HB4)
MF_MEDIA_ENGINE_CONTENT_PROTECTION_FLAGS = iid
End Function
Public Function MF_MEDIA_ENGINE_CONTENT_PROTECTION_MANAGER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFDD6DFAA, &HBD85, &H4AF3, &H9E, &HF, &HA0, &H1D, &H53, &H9D, &H87, &H6A)
MF_MEDIA_ENGINE_CONTENT_PROTECTION_MANAGER = iid
End Function
Public Function MF_MEDIA_ENGINE_AUDIO_ENDPOINT_ROLE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD2CB93D1, &H116A, &H44F2, &H93, &H85, &HF7, &HD0, &HFD, &HA2, &HFB, &H46)
MF_MEDIA_ENGINE_AUDIO_ENDPOINT_ROLE = iid
End Function
Public Function MF_MEDIA_ENGINE_AUDIO_CATEGORY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC8D4C51D, &H350E, &H41F2, &HBA, &H46, &HFA, &HEB, &HBB, &H8, &H57, &HF6)
MF_MEDIA_ENGINE_AUDIO_CATEGORY = iid
End Function
Public Function MF_MEDIA_ENGINE_STREAM_CONTAINS_ALPHA_CHANNEL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5CBFAF44, &HD2B2, &H4CFB, &H80, &HA7, &HD4, &H29, &HC7, &H4C, &H78, &H9D)
MF_MEDIA_ENGINE_STREAM_CONTAINS_ALPHA_CHANNEL = iid
End Function
Public Function MF_MEDIA_ENGINE_BROWSER_COMPATIBILITY_MODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4E0212E2, &HE18F, &H41E1, &H95, &HE5, &HC0, &HE7, &HE9, &H23, &H5B, &HC3)
MF_MEDIA_ENGINE_BROWSER_COMPATIBILITY_MODE = iid
End Function
Public Function MF_MEDIA_ENGINE_BROWSER_COMPATIBILITY_MODE_IE9() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H52C2D39, &H40C0, &H4188, &HAB, &H86, &HF8, &H28, &H27, &H3B, &H75, &H22)
MF_MEDIA_ENGINE_BROWSER_COMPATIBILITY_MODE_IE9 = iid
End Function
Public Function MF_MEDIA_ENGINE_BROWSER_COMPATIBILITY_MODE_IE10() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H11A47AFD, &H6589, &H4124, &HB3, &H12, &H61, &H58, &HEC, &H51, &H7F, &HC3)
MF_MEDIA_ENGINE_BROWSER_COMPATIBILITY_MODE_IE10 = iid
End Function
Public Function MF_MEDIA_ENGINE_BROWSER_COMPATIBILITY_MODE_IE11() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1CF1315F, &HCE3F, &H4035, &H93, &H91, &H16, &H14, &H2F, &H77, &H51, &H89)
MF_MEDIA_ENGINE_BROWSER_COMPATIBILITY_MODE_IE11 = iid
End Function
Public Function MF_MEDIA_ENGINE_BROWSER_COMPATIBILITY_MODE_IE_EDGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA6F3E465, &H3ACA, &H442C, &HA3, &HF0, &HAD, &H6D, &HDA, &HD8, &H39, &HAE)
MF_MEDIA_ENGINE_BROWSER_COMPATIBILITY_MODE_IE_EDGE = iid
End Function
Public Function MF_MEDIA_ENGINE_COMPATIBILITY_MODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3EF26AD4, &HDC54, &H45DE, &HB9, &HAF, &H76, &HC8, &HC6, &H6B, &HFA, &H8E)
MF_MEDIA_ENGINE_COMPATIBILITY_MODE = iid
End Function
Public Function MF_MEDIA_ENGINE_COMPATIBILITY_MODE_WWA_EDGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H15B29098, &H9F01, &H4E4D, &HB6, &H5A, &HC0, &H6C, &H6C, &H89, &HDA, &H2A)
MF_MEDIA_ENGINE_COMPATIBILITY_MODE_WWA_EDGE = iid
End Function
Public Function MF_MEDIA_ENGINE_COMPATIBILITY_MODE_WIN10() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5B25E089, &H6CA7, &H4139, &HA2, &HCB, &HFC, &HAA, &HB3, &H95, &H52, &HA3)
MF_MEDIA_ENGINE_COMPATIBILITY_MODE_WIN10 = iid
End Function
Public Function MF_MEDIA_ENGINE_SOURCE_RESOLVER_CONFIG_STORE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAC0C497, &HB3C4, &H48C9, &H9C, &HDE, &HBB, &H8C, &HA2, &H44, &H2C, &HA3)
MF_MEDIA_ENGINE_SOURCE_RESOLVER_CONFIG_STORE = iid
End Function
Public Function MF_MEDIA_ENGINE_TRACK_ID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H65BEA312, &H4043, &H4815, &H8E, &HAB, &H44, &HDC, &HE2, &HEF, &H8F, &H2A)
MF_MEDIA_ENGINE_TRACK_ID = iid
End Function
Public Function MF_MEDIA_ENGINE_TELEMETRY_APPLICATION_ID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1E7B273B, &HA7E4, &H402A, &H8F, &H51, &HC4, &H8E, &H88, &HA2, &HCA, &HBC)
MF_MEDIA_ENGINE_TELEMETRY_APPLICATION_ID = iid
End Function
Public Function MF_MEDIA_ENGINE_SYNCHRONOUS_CLOSE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC3C2E12F, &H7E0E, &H4E43, &HB9, &H1C, &HDC, &H99, &H2C, &HCD, &HFA, &H5E)
MF_MEDIA_ENGINE_SYNCHRONOUS_CLOSE = iid
End Function
Public Function MF_MEDIA_ENGINE_MEDIA_PLAYER_MODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3DDD8D45, &H5AA1, &H4112, &H82, &HE5, &H36, &HF6, &HA2, &H19, &H7E, &H6E)
MF_MEDIA_ENGINE_MEDIA_PLAYER_MODE = iid
End Function
Public Function CLSID_MFMediaEngineClassFactory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB44392DA, &H499B, &H446B, &HA4, &HCB, &H0, &H5F, &HEA, &HD0, &HE6, &HD5)
CLSID_MFMediaEngineClassFactory = iid
End Function
Public Function MF_MEDIA_ENGINE_TIMEDTEXT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H805EA411, &H92E0, &H4E59, &H9B, &H6E, &H5C, &H7D, &H79, &H15, &HE6, &H4F)
 MF_MEDIA_ENGINE_TIMEDTEXT = iid
End Function
Public Function MF_MEDIA_ENGINE_CONTINUE_ON_CODEC_ERROR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDBCDB7F9, &H48E4, &H4295, &HB7, &HD, &HD5, &H18, &H23, &H4E, &HEB, &H38)
MF_MEDIA_ENGINE_CONTINUE_ON_CODEC_ERROR = iid
End Function
Public Function MF_MEDIA_ENGINE_EME_CALLBACK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H494553A7, &HA481, &H4CB7, &HBE, &HC5, &H38, &H9, &H3, &H51, &H37, &H31)
MF_MEDIA_ENGINE_EME_CALLBACK = iid
End Function
Public Function MF_CONTENTDECRYPTIONMODULE_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H15320C45, &HFF80, &H484A, &H9D, &HCB, &HD, &HF8, &H94, &HE6, &H9A, &H1)
 MF_CONTENTDECRYPTIONMODULE_SERVICE = iid
End Function
Public Function CLSID_MPEG2DLNASink() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFA5FE7C5, &H6A1D, &H4B11, &HB4, &H1F, &HF9, &H59, &HD6, &HC7, &H65, &H0)
 CLSID_MPEG2DLNASink = iid
End Function
Public Function MF_MP2DLNA_USE_MMCSS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H54F3E2EE, &HA2A2, &H497D, &H98, &H34, &H97, &H3A, &HFD, &HE5, &H21, &HEB)
 MF_MP2DLNA_USE_MMCSS = iid
End Function
Public Function MF_MP2DLNA_VIDEO_BIT_RATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE88548DE, &H73B4, &H42D7, &H9C, &H75, &HAD, &HFA, &HA, &H2A, &H6E, &H4C)
 MF_MP2DLNA_VIDEO_BIT_RATE = iid
End Function
Public Function MF_MP2DLNA_AUDIO_BIT_RATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2D1C070E, &H2B5F, &H4AB3, &HA7, &HE6, &H8D, &H94, &H3B, &HA8, &HD0, &HA)
 MF_MP2DLNA_AUDIO_BIT_RATE = iid
End Function
Public Function MF_MP2DLNA_ENCODE_QUALITY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB52379D7, &H1D46, &H4FB6, &HA3, &H17, &HA4, &HA5, &HF6, &H9, &H59, &HF8)
 MF_MP2DLNA_ENCODE_QUALITY = iid
End Function
Public Function MF_MP2DLNA_STATISTICS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H75E488A3, &HD5AD, &H4898, &H85, &HE0, &HBC, &HCE, &H24, &HA7, &H22, &HD7)
 MF_MP2DLNA_STATISTICS = iid
End Function
Public Function MF_MEDIA_SHARING_ENGINE_DEVICE_NAME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H771E05D1, &H862F, &H4299, &H95, &HAC, &HAE, &H81, &HFD, &H14, &HF3, &HE7)
MF_MEDIA_SHARING_ENGINE_DEVICE_NAME = iid
End Function
Public Function MF_MEDIA_SHARING_ENGINE_DEVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB461C58A, &H7A08, &H4B98, &H99, &HA8, &H70, &HFD, &H5F, &H3B, &HAD, &HFD)
MF_MEDIA_SHARING_ENGINE_DEVICE = iid
End Function
Public Function MF_MEDIA_SHARING_ENGINE_INITIAL_SEEK_TIME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6F3497F5, &HD528, &H4A4F, &H8D, &HD7, &HDB, &H36, &H65, &H7E, &HC4, &HC9)
MF_MEDIA_SHARING_ENGINE_INITIAL_SEEK_TIME = iid
End Function
Public Function MF_SHUTDOWN_RENDERER_ON_ENGINE_SHUTDOWN() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC112D94D, &H6B9C, &H48F8, &HB6, &HF9, &H79, &H50, &HFF, &H9A, &HB7, &H1E)
MF_SHUTDOWN_RENDERER_ON_ENGINE_SHUTDOWN = iid
End Function
Public Function MF_PREFERRED_SOURCE_URI() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5FC85488, &H436A, &H4DB8, &H90, &HAF, &H4D, &HB4, &H2, &HAE, &H5C, &H57)
MF_PREFERRED_SOURCE_URI = iid
End Function
Public Function MF_SHARING_ENGINE_SHAREDRENDERER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEFA446A0, &H73E7, &H404E, &H8A, &HE2, &HFE, &HF6, &HA, &HF5, &HA3, &H2B)
MF_SHARING_ENGINE_SHAREDRENDERER = iid
End Function
Public Function MF_SHARING_ENGINE_CALLBACK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57DC1E95, &HD252, &H43FA, &H9B, &HBC, &H18, &H0, &H70, &HEE, &HFE, &H6D)
MF_SHARING_ENGINE_CALLBACK = iid
End Function
Public Function CLSID_MFMediaSharingEngineClassFactory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF8E307FB, &H6D45, &H4AD3, &H99, &H93, &H66, &HCD, &H5A, &H52, &H96, &H59)
CLSID_MFMediaSharingEngineClassFactory = iid
End Function
Public Function CLSID_MFImageSharingEngineClassFactory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB22C3339, &H87F3, &H4059, &HA0, &HC5, &H3, &H7A, &HA9, &H70, &H7E, &HAF)
CLSID_MFImageSharingEngineClassFactory = iid
End Function
Public Function CLSID_PlayToSourceClassFactory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDA17539A, &H3DC3, &H42C1, &HA7, &H49, &HA1, &H83, &HB5, &H1F, &H8, &H5E)
CLSID_PlayToSourceClassFactory = iid
End Function
Public Function GUID_PlayToService() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF6A8FF9D, &H9E14, &H41C9, &HBF, &HF, &H12, &HA, &H2B, &H3C, &HE1, &H20)
GUID_PlayToService = iid
End Function
Public Function GUID_NativeDeviceService() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEF71E53C, &H52F4, &H43C5, &HB8, &H6A, &HAD, &H6C, &HB2, &H16, &HA6, &H1E)
GUID_NativeDeviceService = iid
End Function
Public Function MF_DEVSOURCE_ATTRIBUTE_ENABLE_MS_CAMERA_EFFECTS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H28A5531A, &H57DD, &H4FD5, &HAA, &HA7, &H38, &H5A, &HBF, &H57, &HD7, &H85)
MF_DEVSOURCE_ATTRIBUTE_ENABLE_MS_CAMERA_EFFECTS = iid
End Function
Public Function MF_VIRTUALCAMERA_ASSOCIATED_CAMERA_SOURCES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1BB79E7C, &H5D83, &H438C, &H94, &HD8, &HE5, &HF0, &HDF, &H6D, &H32, &H79)
MF_VIRTUALCAMERA_ASSOCIATED_CAMERA_SOURCES = iid
End Function
Public Function MF_VIRTUALCAMERA_PROVIDE_ASSOCIATED_CAMERA_SOURCES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF0273718, &H4A4D, &H4AC5, &HA1, &H5D, &H30, &H5E, &HB5, &HE9, &H6, &H67)
MF_VIRTUALCAMERA_PROVIDE_ASSOCIATED_CAMERA_SOURCES = iid
End Function
Public Function MF_VIRTUALCAMERA_CONFIGURATION_APP_PACKAGE_FAMILY_NAME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H658ABE51, &H8044, &H462E, &H97, &HEA, &HE6, &H76, &HFD, &H72, &H5, &H5F)
MF_VIRTUALCAMERA_CONFIGURATION_APP_PACKAGE_FAMILY_NAME = iid
End Function
Public Function MF_FRAMESERVER_VCAMEVENT_EXTENDED_SOURCE_INITIALIZE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE52C4DFF, &HE46D, &H4D0B, &HBC, &H75, &HDD, &HD4, &HC8, &H72, &H3F, &H96)
MF_FRAMESERVER_VCAMEVENT_EXTENDED_SOURCE_INITIALIZE = iid
End Function
Public Function MF_FRAMESERVER_VCAMEVENT_EXTENDED_SOURCE_START() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB1EEB989, &HB456, &H4F4A, &HAE, &H40, &H7, &H9C, &H28, &HE2, &H4A, &HF8)
MF_FRAMESERVER_VCAMEVENT_EXTENDED_SOURCE_START = iid
End Function
Public Function MF_FRAMESERVER_VCAMEVENT_EXTENDED_SOURCE_STOP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB7FE7A61, &HFE91, &H415E, &H86, &H8, &HD3, &H7D, &HED, &HB1, &HA5, &H8B)
MF_FRAMESERVER_VCAMEVENT_EXTENDED_SOURCE_STOP = iid
End Function
Public Function MF_FRAMESERVER_VCAMEVENT_EXTENDED_SOURCE_UNINITIALIZE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA0EBABA7, &HA422, &H4E33, &H84, &H1, &HB3, &H7D, &H28, &H0, &HAA, &H67)
MF_FRAMESERVER_VCAMEVENT_EXTENDED_SOURCE_UNINITIALIZE = iid
End Function
Public Function MF_FRAMESERVER_VCAMEVENT_EXTENDED_PIPELINE_SHUTDOWN() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H45A81B31, &H43F8, &H4E5D, &H8C, &HE2, &H22, &HDC, &HE0, &H26, &H99, &H6D)
MF_FRAMESERVER_VCAMEVENT_EXTENDED_PIPELINE_SHUTDOWN = iid
End Function
Public Function MF_FRAMESERVER_VCAMEVENT_EXTENDED_CUSTOM_EVENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6E59489C, &H47D3, &H4467, &H83, &HEF, &H12, &HD3, &H4E, &H87, &H16, &H65)
MF_FRAMESERVER_VCAMEVENT_EXTENDED_CUSTOM_EVENT = iid
End Function
Public Function MFNETSOURCE_CROSS_ORIGIN_SUPPORT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9842207C, &HB02C, &H4271, &HA2, &HFC, &H72, &HE4, &H93, &H8, &HE5, &HC2)
MFNETSOURCE_CROSS_ORIGIN_SUPPORT = iid
End Function
Public Function MFNETSOURCE_HTTP_DOWNLOAD_SESSION_PROVIDER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7D55081E, &H307D, &H4D6D, &HA6, &H63, &HA9, &H3B, &HE9, &H7C, &H4B, &H5C)
MFNETSOURCE_HTTP_DOWNLOAD_SESSION_PROVIDER = iid
End Function
Public Function MF_SD_MEDIASOURCE_STATUS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1913678B, &HFC0F, &H44DA, &H8F, &H43, &H1B, &HA3, &HB5, &H26, &HF4, &HAE)
MF_SD_MEDIASOURCE_STATUS = iid
End Function
Public Function MF_SD_VIDEO_SPHERICAL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA51DA449, &H3FDC, &H478C, &HBC, &HB5, &H30, &HBE, &H76, &H59, &H5F, &H55)
MF_SD_VIDEO_SPHERICAL = iid
End Function
Public Function MF_SD_VIDEO_SPHERICAL_FORMAT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4A8FC407, &H6EA1, &H46C8, &HB5, &H67, &H69, &H71, &HD4, &HA1, &H39, &HC3)
MF_SD_VIDEO_SPHERICAL_FORMAT = iid
End Function
Public Function MF_SD_VIDEO_SPHERICAL_INITIAL_VIEWDIRECTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H11D25A49, &HBB62, &H467F, &H9D, &HB1, &HC1, &H71, &H65, &H71, &H6C, &H49)
MF_SD_VIDEO_SPHERICAL_INITIAL_VIEWDIRECTION = iid
End Function
Public Function MF_MEDIASOURCE_EXPOSE_ALL_STREAMS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE7F250B8, &H8FD9, &H4A09, &HB6, &HC1, &H6A, &H31, &H5C, &H7C, &H72, &HE)
MF_MEDIASOURCE_EXPOSE_ALL_STREAMS = iid
End Function
Public Function MF_ST_MEDIASOURCE_COLLECTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H616DE972, &H83AD, &H4950, &H81, &H70, &H63, &HD, &H19, &HCB, &HE3, &H7)
MF_ST_MEDIASOURCE_COLLECTION = iid
End Function
Public Function MF_DEVICESTREAM_FILTER_KSCONTROL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H46783CCA, &H3DF5, &H4923, &HA9, &HEF, &H36, &HB7, &H22, &H3E, &HDD, &HE0)
MF_DEVICESTREAM_FILTER_KSCONTROL = iid
End Function
Public Function MF_DEVICESTREAM_PIN_KSCONTROL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEF3EF9A7, &H87F2, &H48CA, &HBE, &H2, &H67, &H48, &H78, &H91, &H8E, &H98)
MF_DEVICESTREAM_PIN_KSCONTROL = iid
End Function
Public Function MF_DEVICESTREAM_SOURCE_ATTRIBUTES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2F8CB617, &H361B, &H434F, &H85, &HEA, &H99, &HA0, &H3E, &H1C, &HE4, &HE0)
MF_DEVICESTREAM_SOURCE_ATTRIBUTES = iid
End Function
Public Function MF_DEVICESTREAM_FRAMESERVER_HIDDEN() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF402567B, &H4D91, &H4179, &H96, &HD1, &H74, &HC8, &H48, &HC, &H20, &H34)
 MF_DEVICESTREAM_FRAMESERVER_HIDDEN = iid
End Function
Public Function MF_STF_VERSION_INFO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6770BD39, &HEF82, &H44EE, &HA4, &H9B, &H93, &H4B, &HEB, &H24, &HAE, &HF7)
 MF_STF_VERSION_INFO = iid
End Function
Public Function MF_STF_VERSION_DATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H31A165D5, &HDF67, &H4095, &H8E, &H44, &H88, &H68, &HFC, &H20, &HDB, &HFD)
 MF_STF_VERSION_DATE = iid
End Function
Public Function MF_DEVICESTREAM_REQUIRED_CAPABILITIES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6D8B957E, &H7CF6, &H43F4, &HAF, &H56, &H9C, &HE, &H1E, &H4F, &HCB, &HE1)
 MF_DEVICESTREAM_REQUIRED_CAPABILITIES = iid
End Function
Public Function MF_DEVICESTREAM_REQUIRED_SDDL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H331AE85D, &HC0D3, &H49BA, &H83, &HBA, &H82, &HA1, &H2D, &H63, &HCD, &HD6)
 MF_DEVICESTREAM_REQUIRED_SDDL = iid
End Function
Public Function MF_DEVICEMFT_SENSORPROFILE_COLLECTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H36EBDC44, &HB12C, &H441B, &H89, &HF4, &H8, &HB2, &HF4, &H1A, &H9C, &HFC)
MF_DEVICEMFT_SENSORPROFILE_COLLECTION = iid
End Function
Public Function MF_DEVICESTREAM_SENSORSTREAM_ID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE35B9FE4, &H659, &H4CAD, &HBB, &H51, &H33, &H16, &HB, &HE7, &HE4, &H13)
MF_DEVICESTREAM_SENSORSTREAM_ID = iid
End Function
Public Function CLSID_CameraConfigurationManager() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C92B540, &H5854, &H4A17, &H92, &HB6, &HAC, &H89, &HC9, &H6E, &H96, &H83)
CLSID_CameraConfigurationManager = iid
End Function
Public Function KSPROPERTYSETID_ANYCAMERACONTROL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H94DD0C30, &H28C7, &H4EFB, &H9D, &H6B, &H81, &H23, &H0, &HFB, &HC, &H7F)
KSPROPERTYSETID_ANYCAMERACONTROL = iid
End Function
Public Function MFStreamExtension_ExtendedCameraIntrinsics() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAA74B3DF, &H9A2C, &H48D6, &H83, &H93, &H5B, &HD1, &HC1, &HA8, &H1E, &H6E)
MFStreamExtension_ExtendedCameraIntrinsics = iid
End Function
Public Function MFSampleExtension_ExtendedCameraIntrinsics() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H560BC4A5, &H4DE0, &H4113, &H9C, &HDC, &H83, &H2D, &HB9, &H74, &HF, &H3D)
MFSampleExtension_ExtendedCameraIntrinsics = iid
End Function
Public Function MF_SA_D3D11_BINDFLAGS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEACF97AD, &H65C, &H4408, &HBE, &HE3, &HFD, &HCB, &HFD, &H12, &H8B, &HE2)
MF_SA_D3D11_BINDFLAGS = iid
End Function
Public Function MF_SA_D3D11_USAGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE85FE442, &H2CA3, &H486E, &HA9, &HC7, &H10, &H9D, &HDA, &H60, &H98, &H80)
MF_SA_D3D11_USAGE = iid
End Function
Public Function MF_SA_D3D11_AWARE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H206B4FC8, &HFCF9, &H4C51, &HAF, &HE3, &H97, &H64, &H36, &H9E, &H33, &HA0)
MF_SA_D3D11_AWARE = iid
End Function
Public Function MF_SA_D3D11_SHARED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7B8F32C3, &H6D96, &H4B89, &H92, &H3, &HDD, &H38, &HB6, &H14, &H14, &HF3)
MF_SA_D3D11_SHARED = iid
End Function
Public Function MF_SA_D3D11_SHARED_WITHOUT_MUTEX() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H39DBD44D, &H2E44, &H4931, &HA4, &HC8, &H35, &H2D, &H3D, &HC4, &H21, &H15)
MF_SA_D3D11_SHARED_WITHOUT_MUTEX = iid
End Function
Public Function MF_SA_D3D11_ALLOW_DYNAMIC_YUV_TEXTURE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCE06D49F, &H613, &H4B9D, &H86, &HA6, &HD8, &HC4, &HF9, &HC1, &H0, &H75)
MF_SA_D3D11_ALLOW_DYNAMIC_YUV_TEXTURE = iid
End Function
Public Function MF_SA_D3D11_HW_PROTECTED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3A8BA9D9, &H92CA, &H4307, &HA3, &H91, &H69, &H99, &HDB, &HF3, &HB6, &HCE)
MF_SA_D3D11_HW_PROTECTED = iid
End Function
Public Function MF_SA_D3D_AWARE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEAA35C29, &H775E, &H488E, &H9B, &H61, &HB3, &H28, &H3E, &H49, &H58, &H3B)
MF_SA_D3D_AWARE = iid
End Function
Public Function MFT_SUPPORT_3DVIDEO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H93F81B1, &H4F2E, &H4631, &H81, &H68, &H79, &H34, &H3, &H2A, &H1, &HD3)
MFT_SUPPORT_3DVIDEO = iid
End Function
Public Function MF_ENABLE_3DVIDEO_OUTPUT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBDAD7BCA, &HE5F, &H4B10, &HAB, &H16, &H26, &HDE, &H38, &H1B, &H62, &H93)
MF_ENABLE_3DVIDEO_OUTPUT = iid
End Function
Public Function MF_SA_BUFFERS_PER_SAMPLE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H873C5171, &H1E3D, &H4E25, &H98, &H8D, &HB4, &H33, &HCE, &H4, &H19, &H83)
MF_SA_BUFFERS_PER_SAMPLE = iid
End Function
Public Function MF_SA_D3D11_ALLOCATE_DISPLAYABLE_RESOURCES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEEFACE6D, &H2EA9, &H4ADF, &HBB, &HDF, &H7B, &HBC, &H48, &H2A, &H1B, &H6D)
MF_SA_D3D11_ALLOCATE_DISPLAYABLE_RESOURCES = iid
End Function
Public Function MFT_DECODER_EXPOSE_OUTPUT_TYPES_IN_NATIVE_ORDER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEF80833F, &HF8FA, &H44D9, &H80, &HD8, &H41, &HED, &H62, &H32, &H67, &HC)
MFT_DECODER_EXPOSE_OUTPUT_TYPES_IN_NATIVE_ORDER = iid
End Function
Public Function MFT_DECODER_QUALITY_MANAGEMENT_CUSTOM_CONTROL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA24E30D7, &HDE25, &H4558, &HBB, &HFB, &H71, &H7, &HA, &H2D, &H33, &H2E)
MFT_DECODER_QUALITY_MANAGEMENT_CUSTOM_CONTROL = iid
End Function
Public Function MFT_DECODER_QUALITY_MANAGEMENT_RECOVERY_WITHOUT_ARTIFACTS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD8980DEB, &HA48, &H425F, &H86, &H23, &H61, &H1D, &HB4, &H1D, &H38, &H10)
MFT_DECODER_QUALITY_MANAGEMENT_RECOVERY_WITHOUT_ARTIFACTS = iid
End Function
Public Function MFT_REMUX_MARK_I_PICTURE_AS_CLEAN_POINT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H364E8F85, &H3F2E, &H436C, &HB2, &HA2, &H44, &H40, &HA0, &H12, &HA9, &HE8)
MFT_REMUX_MARK_I_PICTURE_AS_CLEAN_POINT = iid
End Function
Public Function MFT_DECODER_FINAL_VIDEO_RESOLUTION_HINT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDC2F8496, &H15C4, &H407A, &HB6, &HF0, &H1B, &H66, &HAB, &H5F, &HBF, &H53)
MFT_DECODER_FINAL_VIDEO_RESOLUTION_HINT = iid
End Function
Public Function MFT_ENCODER_SUPPORTS_CONFIG_EVENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H86A355AE, &H3A77, &H4EC4, &H9F, &H31, &H1, &H14, &H9A, &H4E, &H92, &HDE)
MFT_ENCODER_SUPPORTS_CONFIG_EVENT = iid
End Function
Public Function MFT_ENUM_HARDWARE_VENDOR_ID_Attribute() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3AECB0CC, &H35B, &H4BCC, &H81, &H85, &H2B, &H8D, &H55, &H1E, &HF3, &HAF)
MFT_ENUM_HARDWARE_VENDOR_ID_Attribute = iid
End Function
Public Function MF_TRANSFORM_ASYNC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF81A699A, &H649A, &H497D, &H8C, &H73, &H29, &HF8, &HFE, &HD6, &HAD, &H7A)
MF_TRANSFORM_ASYNC = iid
End Function
Public Function MF_TRANSFORM_ASYNC_UNLOCK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE5666D6B, &H3422, &H4EB6, &HA4, &H21, &HDA, &H7D, &HB1, &HF8, &HE2, &H7)
MF_TRANSFORM_ASYNC_UNLOCK = iid
End Function
Public Function MF_TRANSFORM_FLAGS_Attribute() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9359BB7E, &H6275, &H46C4, &HA0, &H25, &H1C, &H1, &HE4, &H5F, &H1A, &H86)
MF_TRANSFORM_FLAGS_Attribute = iid
End Function
Public Function MF_TRANSFORM_CATEGORY_Attribute() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCEABBA49, &H506D, &H4757, &HA6, &HFF, &H66, &HC1, &H84, &H98, &H7E, &H4E)
MF_TRANSFORM_CATEGORY_Attribute = iid
End Function
Public Function MFT_TRANSFORM_CLSID_Attribute() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6821C42B, &H65A4, &H4E82, &H99, &HBC, &H9A, &H88, &H20, &H5E, &HCD, &HC)
MFT_TRANSFORM_CLSID_Attribute = iid
End Function
Public Function MFT_INPUT_TYPES_Attributes() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4276C9B1, &H759D, &H4BF3, &H9C, &HD0, &HD, &H72, &H3D, &H13, &H8F, &H96)
MFT_INPUT_TYPES_Attributes = iid
End Function
Public Function MFT_OUTPUT_TYPES_Attributes() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8EAE8CF3, &HA44F, &H4306, &HBA, &H5C, &HBF, &H5D, &HDA, &H24, &H28, &H18)
MFT_OUTPUT_TYPES_Attributes = iid
End Function
Public Function MFT_ENUM_HARDWARE_URL_Attribute() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2FB866AC, &HB078, &H4942, &HAB, &H6C, &H0, &H3D, &H5, &HCD, &HA6, &H74)
MFT_ENUM_HARDWARE_URL_Attribute = iid
End Function
Public Function MFT_FRIENDLY_NAME_Attribute() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H314FFBAE, &H5B41, &H4C95, &H9C, &H19, &H4E, &H7D, &H58, &H6F, &HAC, &HE3)
MFT_FRIENDLY_NAME_Attribute = iid
End Function
Public Function MFT_CONNECTED_STREAM_ATTRIBUTE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H71EEB820, &HA59F, &H4DE2, &HBC, &HEC, &H38, &HDB, &H1D, &HD6, &H11, &HA4)
MFT_CONNECTED_STREAM_ATTRIBUTE = iid
End Function
Public Function MFT_CONNECTED_TO_HW_STREAM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H34E6E728, &H6D6, &H4491, &HA5, &H53, &H47, &H95, &H65, &HD, &HB9, &H12)
MFT_CONNECTED_TO_HW_STREAM = iid
End Function
Public Function MFT_PREFERRED_OUTPUTTYPE_Attribute() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7E700499, &H396A, &H49EE, &HB1, &HB4, &HF6, &H28, &H2, &H1E, &H8C, &H9D)
MFT_PREFERRED_OUTPUTTYPE_Attribute = iid
End Function
Public Function MFT_PROCESS_LOCAL_Attribute() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H543186E4, &H4649, &H4E65, &HB5, &H88, &H4A, &HA3, &H52, &HAF, &HF3, &H79)
MFT_PROCESS_LOCAL_Attribute = iid
End Function
Public Function MFT_PREFERRED_ENCODER_PROFILE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53004909, &H1EF5, &H46D7, &HA1, &H8E, &H5A, &H75, &HF8, &HB5, &H90, &H5F)
MFT_PREFERRED_ENCODER_PROFILE = iid
End Function
Public Function MFT_HW_TIMESTAMP_WITH_QPC_Attribute() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8D030FB8, &HCC43, &H4258, &HA2, &H2E, &H92, &H10, &HBE, &HF8, &H9B, &HE4)
MFT_HW_TIMESTAMP_WITH_QPC_Attribute = iid
End Function
Public Function MFT_FIELDOFUSE_UNLOCK_Attribute() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8EC2E9FD, &H9148, &H410D, &H83, &H1E, &H70, &H24, &H39, &H46, &H1A, &H8E)
MFT_FIELDOFUSE_UNLOCK_Attribute = iid
End Function
Public Function MFT_CODEC_MERIT_Attribute() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H88A7CB15, &H7B07, &H4A34, &H91, &H28, &HE6, &H4C, &H67, &H3, &HC4, &HD3)
MFT_CODEC_MERIT_Attribute = iid
End Function
Public Function MFT_ENUM_TRANSCODE_ONLY_ATTRIBUTE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H111EA8CD, &HB62A, &H4BDB, &H89, &HF6, &H67, &HFF, &HCD, &HC2, &H45, &H8B)
MFT_ENUM_TRANSCODE_ONLY_ATTRIBUTE = iid
End Function
Public Function MFT_POLICY_SET_AWARE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5A633B19, &HCC39, &H4FA8, &H8C, &HA5, &H59, &H98, &H1B, &H7A, &H0, &H18)
MFT_POLICY_SET_AWARE = iid
End Function


Public Function MFP_PKEY_StreamRenderingResults() As PROPERTYKEY
    Static pk As PROPERTYKEY
    If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &HA7CF9740, &HE8D9, &H4A87, &HBD, &H8E, &H29, &H67, &H0, &H1F, &HD3, &HAD, &H1)
    MFP_PKEY_StreamRenderingResults = pk
End Function
Public Function MFP_PKEY_StreamIndex() As PROPERTYKEY
    Static pk As PROPERTYKEY
    If (pk.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(pk, &HA7CF9740, &HE8D9, &H4A87, &HBD, &H8E, &H29, &H67, &H0, &H1F, &HD3, &HAD, &H0)
    MFP_PKEY_StreamIndex = pk
End Function



Public Function MR_VIDEO_RENDER_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1092A86C, &HAB1A, &H459A, &HA3, &H36, &H83, &H1F, &HBC, &H4D, &H11, &HFF)
MR_VIDEO_RENDER_SERVICE = iid
End Function
Public Function MR_VIDEO_MIXER_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H73CD2FC, &H6CF4, &H40B7, &H88, &H59, &HE8, &H95, &H52, &HC8, &H41, &HF8)
MR_VIDEO_MIXER_SERVICE = iid
End Function
Public Function MR_VIDEO_ACCELERATION_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEFEF5175, &H5C7D, &H4CE2, &HBB, &HBD, &H34, &HFF, &H8B, &HCA, &H65, &H54)
MR_VIDEO_ACCELERATION_SERVICE = iid
End Function
Public Function MR_BUFFER_SERVICE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA562248C, &H9AC6, &H4FFC, &H9F, &HBA, &H3A, &HF8, &HF8, &HAD, &H1A, &H4D)
MR_BUFFER_SERVICE = iid
End Function
Public Function VIDEO_ZOOM_RECT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7AAA1638, &H1B7F, &H4C93, &HBD, &H89, &H5B, &H9C, &H9F, &HB6, &HFC, &HF0)
VIDEO_ZOOM_RECT = iid
End Function


Public Function MF_EVENT_SESSIONCAPS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7E5EBCD0, &H11B8, &H4ABE, &HAF, &HAD, &H10, &HF6, &H59, &H9A, &H7F, &H42)
MF_EVENT_SESSIONCAPS = iid
End Function
Public Function MF_EVENT_SESSIONCAPS_DELTA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7E5EBCD1, &H11B8, &H4ABE, &HAF, &HAD, &H10, &HF6, &H59, &H9A, &H7F, &H42)
MF_EVENT_SESSIONCAPS_DELTA = iid
End Function
Public Function MF_EVENT_TOPOLOGY_STATUS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H30C5018D, &H9A53, &H454B, &HAD, &H9E, &H6D, &H5F, &H8F, &HA7, &HC4, &H3B)
MF_EVENT_TOPOLOGY_STATUS = iid
End Function
Public Function MF_EVENT_START_PRESENTATION_TIME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5AD914D0, &H9B45, &H4A8D, &HA2, &HC0, &H81, &HD1, &HE5, &HB, &HFB, &H7)
MF_EVENT_START_PRESENTATION_TIME = iid
End Function
Public Function MF_EVENT_PRESENTATION_TIME_OFFSET() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5AD914D1, &H9B45, &H4A8D, &HA2, &HC0, &H81, &HD1, &HE5, &HB, &HFB, &H7)
MF_EVENT_PRESENTATION_TIME_OFFSET = iid
End Function
Public Function MF_EVENT_START_PRESENTATION_TIME_AT_OUTPUT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5AD914D2, &H9B45, &H4A8D, &HA2, &HC0, &H81, &HD1, &HE5, &HB, &HFB, &H7)
MF_EVENT_START_PRESENTATION_TIME_AT_OUTPUT = iid
End Function
Public Function MF_EVENT_SOURCE_FAKE_START() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA8CC55A7, &H6B31, &H419F, &H84, &H5D, &HFF, &HB3, &H51, &HA2, &H43, &H4B)
MF_EVENT_SOURCE_FAKE_START = iid
End Function
Public Function MF_EVENT_SOURCE_PROJECTSTART() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA8CC55A8, &H6B31, &H419F, &H84, &H5D, &HFF, &HB3, &H51, &HA2, &H43, &H4B)
MF_EVENT_SOURCE_PROJECTSTART = iid
End Function
Public Function MF_EVENT_SOURCE_ACTUAL_START() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA8CC55A9, &H6B31, &H419F, &H84, &H5D, &HFF, &HB3, &H51, &HA2, &H43, &H4B)
MF_EVENT_SOURCE_ACTUAL_START = iid
End Function
Public Function MF_EVENT_SOURCE_TOPOLOGY_CANCELED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDB62F650, &H9A5E, &H4704, &HAC, &HF3, &H56, &H3B, &HC6, &HA7, &H33, &H64)
MF_EVENT_SOURCE_TOPOLOGY_CANCELED = iid
End Function
Public Function MF_EVENT_SOURCE_CHARACTERISTICS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H47DB8490, &H8B22, &H4F52, &HAF, &HDA, &H9C, &HE1, &HB2, &HD3, &HCF, &HA8)
MF_EVENT_SOURCE_CHARACTERISTICS = iid
End Function
Public Function MF_EVENT_SOURCE_CHARACTERISTICS_OLD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H47DB8491, &H8B22, &H4F52, &HAF, &HDA, &H9C, &HE1, &HB2, &HD3, &HCF, &HA8)
MF_EVENT_SOURCE_CHARACTERISTICS_OLD = iid
End Function
Public Function MF_EVENT_DO_THINNING() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H321EA6FB, &HDAD9, &H46E4, &HB3, &H1D, &HD2, &HEA, &HE7, &H9, &HE, &H30)
MF_EVENT_DO_THINNING = iid
End Function
Public Function MF_EVENT_SCRUBSAMPLE_TIME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9AC712B3, &HDCB8, &H44D5, &H8D, &HC, &H37, &H45, &H5A, &H27, &H82, &HE3)
MF_EVENT_SCRUBSAMPLE_TIME = iid
End Function
Public Function MF_EVENT_OUTPUT_NODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H830F1A8B, &HC060, &H46DD, &HA8, &H1, &H1C, &H95, &HDE, &HC9, &HB1, &H7)
MF_EVENT_OUTPUT_NODE = iid
End Function
Public Function MF_EVENT_MFT_INPUT_STREAM_ID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF29C2CCA, &H7AE6, &H42D2, &HB2, &H84, &HBF, &H83, &H7C, &HC8, &H74, &HE2)
MF_EVENT_MFT_INPUT_STREAM_ID = iid
End Function
Public Function MF_EVENT_MFT_CONTEXT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB7CD31F1, &H899E, &H4B41, &H80, &HC9, &H26, &HA8, &H96, &HD3, &H29, &H77)
MF_EVENT_MFT_CONTEXT = iid
End Function
Public Function MF_EVENT_STREAM_METADATA_KEYDATA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCD59A4A1, &H4A3B, &H4BBD, &H86, &H65, &H72, &HA4, &HF, &HBE, &HA7, &H76)
MF_EVENT_STREAM_METADATA_KEYDATA = iid
End Function
Public Function MF_EVENT_STREAM_METADATA_CONTENT_KEYIDS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5063449D, &HCC29, &H4FC6, &HA7, &H5A, &HD2, &H47, &HB3, &H5A, &HF8, &H5C)
MF_EVENT_STREAM_METADATA_CONTENT_KEYIDS = iid
End Function
Public Function MF_EVENT_STREAM_METADATA_SYSTEMID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1EA2EF64, &HBA16, &H4A36, &H87, &H19, &HFE, &H75, &H60, &HBA, &H32, &HAD)
MF_EVENT_STREAM_METADATA_SYSTEMID = iid
End Function

Public Function MFSampleExtension_MaxDecodeFrameSize() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD3CC654F, &HF9F3, &H4A13, &H88, &H9F, &HF0, &H4E, &HB2, &HB5, &HB9, &H57)
MFSampleExtension_MaxDecodeFrameSize = iid
End Function
Public Function MFSampleExtension_AccumulatedNonRefPicPercent() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H79EA74DF, &HA740, &H445B, &HBC, &H98, &HC9, &HED, &H1F, &H26, &HE, &HEE)
MFSampleExtension_AccumulatedNonRefPicPercent = iid
End Function
Public Function MFSampleExtension_Encryption_ProtectionScheme() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD054D096, &H28BB, &H45DA, &H87, &HEC, &H74, &HF3, &H51, &H87, &H14, &H6)
MFSampleExtension_Encryption_ProtectionScheme = iid
End Function
Public Function MFSampleExtension_Encryption_CryptByteBlock() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9D84289B, &HC7F, &H4713, &HAB, &H95, &H10, &H8A, &HB4, &H2A, &HD8, &H1)
MFSampleExtension_Encryption_CryptByteBlock = iid
End Function
Public Function MFSampleExtension_Encryption_SkipByteBlock() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD550548, &H8317, &H4AB1, &H84, &H5F, &HD0, &H63, &H6, &HE2, &H93, &HE3)
MFSampleExtension_Encryption_SkipByteBlock = iid
End Function
Public Function MFSampleExtension_Encryption_SubSample_Mapping() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8444F27A, &H69A1, &H48DA, &HBD, &H8, &H11, &HCE, &HF3, &H68, &H30, &HD2)
MFSampleExtension_Encryption_SubSample_Mapping = iid
End Function
Public Function MFSampleExtension_Encryption_ClearSliceHeaderData() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5509A4F4, &H320D, &H4E6C, &H8D, &H1A, &H94, &HC6, &H6D, &HD2, &HC, &HB0)
MFSampleExtension_Encryption_ClearSliceHeaderData = iid
End Function
Public Function MFSampleExtension_Encryption_HardwareProtection_KeyInfoID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8CBFCCEB, &H94A5, &H4DE1, &H82, &H31, &HA8, &H5E, &H47, &HCF, &H81, &HE7)
MFSampleExtension_Encryption_HardwareProtection_KeyInfoID = iid
End Function
Public Function MFSampleExtension_Encryption_HardwareProtection_KeyInfo() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB2372080, &H455B, &H4DD7, &H99, &H89, &H1A, &H95, &H57, &H84, &HB7, &H54)
MFSampleExtension_Encryption_HardwareProtection_KeyInfo = iid
End Function
Public Function MFSampleExtension_Encryption_HardwareProtection_VideoDecryptorContext() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H693470C8, &HE837, &H47A0, &H88, &HCB, &H53, &H5B, &H90, &H5E, &H35, &H82)
MFSampleExtension_Encryption_HardwareProtection_VideoDecryptorContext = iid
End Function
Public Function MFSampleExtension_Encryption_Opaque_Data() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H224D77E5, &H1391, &H4FFB, &H9F, &H41, &HB4, &H32, &HF6, &H8C, &H61, &H1D)
MFSampleExtension_Encryption_Opaque_Data = iid
End Function
Public Function MFSampleExtension_NALULengthInfo() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H19124E7C, &HAD4B, &H465F, &HBB, &H18, &H20, &H18, &H62, &H87, &HB6, &HAF)
MFSampleExtension_NALULengthInfo = iid
End Function
Public Function MFSampleExtension_Encryption_ResumeVideoOutput() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA435ABA5, &HAFDE, &H4CF5, &HBC, &H1C, &HF6, &HAC, &HAF, &H13, &H94, &H9D)
MFSampleExtension_Encryption_ResumeVideoOutput = iid
End Function
Public Function MFSampleExtension_Encryption_NALUTypes() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB0F067C7, &H714C, &H416C, &H8D, &H59, &H5F, &H4D, &HDF, &H89, &H13, &HB6)
MFSampleExtension_Encryption_NALUTypes = iid
End Function
Public Function MFSampleExtension_Encryption_SPSPPSData() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAEDE0FA2, &HE0C, &H453C, &HB7, &HF3, &HDE, &H86, &H93, &H36, &H4D, &H11)
MFSampleExtension_Encryption_SPSPPSData = iid
End Function
Public Function MFSampleExtension_Encryption_SEIData() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CF0E972, &H4542, &H4687, &H99, &H99, &H58, &H5F, &H56, &H5F, &HBA, &H7D)
MFSampleExtension_Encryption_SEIData = iid
End Function
Public Function MFSampleExtension_Encryption_HardwareProtection() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9A2B2D2B, &H8270, &H43E3, &H84, &H48, &H99, &H4F, &H42, &H6E, &H88, &H86)
MFSampleExtension_Encryption_HardwareProtection = iid
End Function
Public Function MFSampleExtension_CleanPoint() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9CDF01D8, &HA0F0, &H43BA, &HB0, &H77, &HEA, &HA0, &H6C, &HBD, &H72, &H8A)
MFSampleExtension_CleanPoint = iid
End Function
Public Function MFSampleExtension_Discontinuity() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9CDF01D9, &HA0F0, &H43BA, &HB0, &H77, &HEA, &HA0, &H6C, &HBD, &H72, &H8A)
MFSampleExtension_Discontinuity = iid
End Function
Public Function MFSampleExtension_Token() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8294DA66, &HF328, &H4805, &HB5, &H51, &H0, &HDE, &HB4, &HC5, &H7A, &H61)
MFSampleExtension_Token = iid
End Function
Public Function MFSampleExtension_ClosedCaption_CEA708() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H26F09068, &HE744, &H47DC, &HAA, &H3, &HDB, &HF2, &H4, &H3, &HBD, &HE6)
MFSampleExtension_ClosedCaption_CEA708 = iid
End Function
Public Function MFSampleExtension_DecodeTimestamp() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H73A954D4, &H9E2, &H4861, &HBE, &HFC, &H94, &HBD, &H97, &HC0, &H8E, &H6E)
MFSampleExtension_DecodeTimestamp = iid
End Function
Public Function MFSampleExtension_VideoEncodeQP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB2EFE478, &HF979, &H4C66, &HB9, &H5E, &HEE, &H2B, &H82, &HC8, &H2F, &H36)
MFSampleExtension_VideoEncodeQP = iid
End Function
Public Function MFSampleExtension_VideoEncodePictureType() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H973704E6, &HCD14, &H483C, &H8F, &H20, &HC9, &HFC, &H9, &H28, &HBA, &HD5)
MFSampleExtension_VideoEncodePictureType = iid
End Function
Public Function MFSampleExtension_FrameCorruption() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB4DD4A8C, &HBEB, &H44C4, &H8B, &H75, &HB0, &H2B, &H91, &H3B, &H4, &HF0)
MFSampleExtension_FrameCorruption = iid
End Function
Public Function MFSampleExtension_DirtyRects() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9BA70225, &HB342, &H4E97, &H91, &H26, &HB, &H56, &H6A, &HB7, &HEA, &H7E)
MFSampleExtension_DirtyRects = iid
End Function
Public Function MFSampleExtension_MoveRegions() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE2A6C693, &H3A8B, &H4B8D, &H95, &HD0, &HF6, &H2, &H81, &HA1, &H2F, &HB7)
MFSampleExtension_MoveRegions = iid
End Function

Public Function MFSampleExtension_HDCP_OptionalHeader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9A2E7390, &H121F, &H455F, &H83, &H76, &HC9, &H74, &H28, &HE0, &HB5, &H40)
MFSampleExtension_HDCP_OptionalHeader = iid
End Function
Public Function MFSampleExtension_HDCP_FrameCounter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9D389C60, &HF507, &H4AA6, &HA4, &HA, &H71, &H2, &H7A, &H2, &HF3, &HDE)
MFSampleExtension_HDCP_FrameCounter = iid
End Function
Public Function MFSampleExtension_HDCP_StreamID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H177E5D74, &HC370, &H4A7A, &H95, &HA2, &H36, &H83, &H3C, &H1, &HD0, &HAF)
MFSampleExtension_HDCP_StreamID = iid
End Function
Public Function MFSampleExtension_Timestamp() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1E436999, &H69BE, &H4C7A, &H93, &H69, &H70, &H6, &H8C, &H2, &H60, &HCB)
MFSampleExtension_Timestamp = iid
End Function
Public Function MFSampleExtension_RepeatFrame() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H88BE738F, &H711, &H4F42, &HB4, &H58, &H34, &H4A, &HED, &H42, &HEC, &H2F)
MFSampleExtension_RepeatFrame = iid
End Function
Public Function MFT_ENCODER_ERROR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC8D1EDA4, &H98E4, &H41D5, &H92, &H97, &H44, &HF5, &H38, &H52, &HF9, &HE)
MFT_ENCODER_ERROR = iid
End Function
Public Function MFT_GFX_DRIVER_VERSION_ID_Attribute() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF34B9093, &H5E0, &H4B16, &H99, &H3D, &H3E, &H2A, &H2C, &HDE, &H6A, &HD3)
MFT_GFX_DRIVER_VERSION_ID_Attribute = iid
End Function
Public Function MFSampleExtension_DescrambleData() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43483BE6, &H4903, &H4314, &HB0, &H32, &H29, &H51, &H36, &H59, &H36, &HFC)
MFSampleExtension_DescrambleData = iid
End Function
Public Function MFSampleExtension_SampleKeyID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9ED713C8, &H9B87, &H4B26, &H82, &H97, &HA9, &H3B, &HC, &H5A, &H8A, &HCC)
MFSampleExtension_SampleKeyID = iid
End Function
Public Function MFSampleExtension_GenKeyFunc() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H441CA1EE, &H6B1F, &H4501, &H90, &H3A, &HDE, &H87, &HDF, &H42, &HF6, &HED)
MFSampleExtension_GenKeyFunc = iid
End Function
Public Function MFSampleExtension_GenKeyCtx() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H188120CB, &HD7DA, &H4B59, &H9B, &H3E, &H92, &H52, &HFD, &H37, &H30, &H1C)
MFSampleExtension_GenKeyCtx = iid
End Function
Public Function MFSampleExtension_PacketCrossOffsets() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2789671D, &H389F, &H40BB, &H90, &HD9, &HC2, &H82, &HF7, &H7F, &H9A, &HBD)
MFSampleExtension_PacketCrossOffsets = iid
End Function
Public Function MFSampleExtension_Encryption_SampleID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6698B84E, &HAFA, &H4330, &HAE, &HB2, &H1C, &HA, &H98, &HD7, &HA4, &H4D)
MFSampleExtension_Encryption_SampleID = iid
End Function
Public Function MFSampleExtension_Encryption_KeyID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H76376591, &H795F, &H4DA1, &H86, &HED, &H9D, &H46, &HEC, &HA1, &H9, &HA9)
MFSampleExtension_Encryption_KeyID = iid
End Function
Public Function MFSampleExtension_Content_KeyID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC6C7F5B0, &HACCA, &H415B, &H87, &HD9, &H10, &H44, &H14, &H69, &HEF, &HC6)
MFSampleExtension_Content_KeyID = iid
End Function
Public Function MFSampleExtension_Encryption_SubSampleMappingSplit() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFE0254B9, &H2AA5, &H4EDC, &H99, &HF7, &H17, &HE8, &H9D, &HBF, &H91, &H74)
MFSampleExtension_Encryption_SubSampleMappingSplit = iid
End Function
Public Function MFSampleExtension_Interlaced() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB1D5830A, &HDEB8, &H40E3, &H90, &HFA, &H38, &H99, &H43, &H71, &H64, &H61)
MFSampleExtension_Interlaced = iid
End Function
Public Function MFSampleExtension_BottomFieldFirst() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H941CE0A3, &H6AE3, &H4DDA, &H9A, &H8, &HA6, &H42, &H98, &H34, &H6, &H17)
MFSampleExtension_BottomFieldFirst = iid
End Function
Public Function MFSampleExtension_RepeatFirstField() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H304D257C, &H7493, &H4FBD, &HB1, &H49, &H92, &H28, &HDE, &H8D, &H9A, &H99)
MFSampleExtension_RepeatFirstField = iid
End Function
Public Function MFSampleExtension_SingleField() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9D85F816, &H658B, &H455A, &HBD, &HE0, &H9F, &HA7, &HE1, &H5A, &HB8, &HF9)
MFSampleExtension_SingleField = iid
End Function
Public Function MFSampleExtension_DerivedFromTopField() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6852465A, &HAE1C, &H4553, &H8E, &H9B, &HC3, &H42, &HF, &HCB, &H16, &H37)
MFSampleExtension_DerivedFromTopField = iid
End Function
Public Function MFSampleExtension_MeanAbsoluteDifference() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1CDBDE11, &H8B4, &H4311, &HA6, &HDD, &HF, &H9F, &H37, &H19, &H7, &HAA)
MFSampleExtension_MeanAbsoluteDifference = iid
End Function
Public Function MFSampleExtension_LongTermReferenceFrameInfo() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9154733F, &HE1BD, &H41BF, &H81, &HD3, &HFC, &HD9, &H18, &HF7, &H13, &H32)
MFSampleExtension_LongTermReferenceFrameInfo = iid
End Function
Public Function MFSampleExtension_ROIRectangle() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3414A438, &H4998, &H4D2C, &HBE, &H82, &HBE, &H3C, &HA0, &HB2, &H4D, &H43)
MFSampleExtension_ROIRectangle = iid
End Function
Public Function MFSampleExtension_LastSlice() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2B5D5457, &H5547, &H4F07, &HB8, &HC8, &HB4, &HA3, &HA9, &HA1, &HDA, &HAC)
MFSampleExtension_LastSlice = iid
End Function
Public Function MFSampleExtension_FeatureMap() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA032D165, &H46FC, &H400A, &HB4, &H49, &H49, &HDE, &H53, &HE6, &H2A, &H6E)
MFSampleExtension_FeatureMap = iid
End Function
Public Function MFSampleExtension_ChromaOnly() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1EB9179C, &HA01F, &H4845, &H8C, &H4, &HE, &H65, &HA2, &H6E, &HB0, &H4F)
MFSampleExtension_ChromaOnly = iid
End Function
Public Function MFSampleExtension_PhotoThumbnail() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H74BBC85C, &HC8BB, &H42DC, &HB5, &H86, &HDA, &H17, &HFF, &HD3, &H5D, &HCC)
MFSampleExtension_PhotoThumbnail = iid
End Function
Public Function MFSampleExtension_PhotoThumbnailMediaType() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H61AD5420, &HEBF8, &H4143, &H89, &HAF, &H6B, &HF2, &H5F, &H67, &H2D, &HEF)
MFSampleExtension_PhotoThumbnailMediaType = iid
End Function
Public Function MFSampleExtension_CaptureMetadata() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2EBE23A8, &HFAF5, &H444A, &HA6, &HA2, &HEB, &H81, &H8, &H80, &HAB, &H5D)
MFSampleExtension_CaptureMetadata = iid
End Function
Public Function MFSampleExtension_MDLCacheCookie() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5F002AF9, &HD8F9, &H41A3, &HB6, &HC3, &HA2, &HAD, &H43, &HF6, &H47, &HAD)
MFSampleExtension_MDLCacheCookie = iid
End Function
Public Function MF_CAPTURE_METADATA_PHOTO_FRAME_FLASH() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF9DD6C6, &H6003, &H45D8, &HBD, &H59, &HF1, &HF5, &H3E, &H3D, &H4, &HE8)
MF_CAPTURE_METADATA_PHOTO_FRAME_FLASH = iid
End Function
Public Function MF_CAPTURE_METADATA_FRAME_RAWSTREAM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9252077B, &H2680, &H49B9, &HAE, &H2, &HB1, &H90, &H75, &H97, &H3B, &H70)
MF_CAPTURE_METADATA_FRAME_RAWSTREAM = iid
End Function
Public Function MF_CAPTURE_METADATA_FOCUSSTATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA87EE154, &H997F, &H465D, &HB9, &H1F, &H29, &HD5, &H3B, &H98, &H2B, &H88)
MF_CAPTURE_METADATA_FOCUSSTATE = iid
End Function
Public Function MF_CAPTURE_METADATA_REQUESTED_FRAME_SETTING_ID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBB3716D9, &H8A61, &H47A4, &H81, &H97, &H45, &H9C, &H7F, &HF1, &H74, &HD5)
MF_CAPTURE_METADATA_REQUESTED_FRAME_SETTING_ID = iid
End Function
Public Function MF_CAPTURE_METADATA_EXPOSURE_TIME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H16B9AE99, &HCD84, &H4063, &H87, &H9D, &HA2, &H8C, &H76, &H33, &H72, &H9E)
MF_CAPTURE_METADATA_EXPOSURE_TIME = iid
End Function
Public Function MF_CAPTURE_METADATA_EXPOSURE_COMPENSATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD198AA75, &H4B62, &H4345, &HAB, &HF3, &H3C, &H31, &HFA, &H12, &HC2, &H99)
MF_CAPTURE_METADATA_EXPOSURE_COMPENSATION = iid
End Function
Public Function MF_CAPTURE_METADATA_ISO_SPEED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE528A68F, &HB2E3, &H44FE, &H8B, &H65, &H7, &HBF, &H4B, &H5A, &H13, &HFF)
MF_CAPTURE_METADATA_ISO_SPEED = iid
End Function
Public Function MF_CAPTURE_METADATA_LENS_POSITION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB5FC8E86, &H11D1, &H4E70, &H81, &H9B, &H72, &H3A, &H89, &HFA, &H45, &H20)
MF_CAPTURE_METADATA_LENS_POSITION = iid
End Function
Public Function MF_CAPTURE_METADATA_SCENE_MODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9CC3B54D, &H5ED3, &H4BAE, &HB3, &H88, &H76, &H70, &HAE, &HF5, &H9E, &H13)
MF_CAPTURE_METADATA_SCENE_MODE = iid
End Function
Public Function MF_CAPTURE_METADATA_FLASH() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4A51520B, &HFB36, &H446C, &H9D, &HF2, &H68, &H17, &H1B, &H9A, &H3, &H89)
MF_CAPTURE_METADATA_FLASH = iid
End Function
Public Function MF_CAPTURE_METADATA_FLASH_POWER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9C0E0D49, &H205, &H491A, &HBC, &H9D, &H2D, &H6E, &H1F, &H4D, &H56, &H84)
MF_CAPTURE_METADATA_FLASH_POWER = iid
End Function
Public Function MF_CAPTURE_METADATA_WHITEBALANCE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC736FD77, &HFB9, &H4E2E, &H97, &HA2, &HFC, &HD4, &H90, &H73, &H9E, &HE9)
MF_CAPTURE_METADATA_WHITEBALANCE = iid
End Function
Public Function MF_CAPTURE_METADATA_ZOOMFACTOR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE50B0B81, &HE501, &H42C2, &HAB, &HF2, &H85, &H7E, &HCB, &H13, &HFA, &H5C)
MF_CAPTURE_METADATA_ZOOMFACTOR = iid
End Function
Public Function MF_CAPTURE_METADATA_FACEROIS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H864F25A6, &H349F, &H46B1, &HA3, &HE, &H54, &HCC, &H22, &H92, &H8A, &H47)
MF_CAPTURE_METADATA_FACEROIS = iid
End Function
Public Function MF_CAPTURE_METADATA_FACEROITIMESTAMPS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE94D50CC, &H3DA0, &H44D4, &HBB, &H34, &H83, &H19, &H8A, &H74, &H18, &H68)
MF_CAPTURE_METADATA_FACEROITIMESTAMPS = iid
End Function
Public Function MF_CAPTURE_METADATA_FACEROICHARACTERIZATIONS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB927A1A8, &H18EF, &H46D3, &HB3, &HAF, &H69, &H37, &H2F, &H94, &HD9, &HB2)
MF_CAPTURE_METADATA_FACEROICHARACTERIZATIONS = iid
End Function
Public Function MF_CAPTURE_METADATA_ISO_GAINS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5802AC9, &HE1D, &H41C7, &HA8, &HC8, &H7E, &H73, &H69, &HF8, &H4E, &H1E)
MF_CAPTURE_METADATA_ISO_GAINS = iid
End Function
Public Function MF_CAPTURE_METADATA_SENSORFRAMERATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDB51357E, &H9D3D, &H4962, &HB0, &H6D, &H7, &HCE, &H65, &HD, &H9A, &HA)
MF_CAPTURE_METADATA_SENSORFRAMERATE = iid
End Function
Public Function MF_CAPTURE_METADATA_WHITEBALANCE_GAINS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE7570C8F, &H2DCB, &H4C7C, &HAA, &HCE, &H22, &HEC, &HE7, &HCC, &HE6, &H47)
MF_CAPTURE_METADATA_WHITEBALANCE_GAINS = iid
End Function
Public Function MF_CAPTURE_METADATA_HISTOGRAM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H85358432, &H2EF6, &H4BA9, &HA3, &HFB, &H6, &HD8, &H29, &H74, &HB8, &H95)
MF_CAPTURE_METADATA_HISTOGRAM = iid
End Function
Public Function MF_CAPTURE_METADATA_EXIF() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2E9575B8, &H8C31, &H4A02, &H85, &H75, &H42, &HB1, &H97, &HB7, &H15, &H92)
MF_CAPTURE_METADATA_EXIF = iid
End Function
Public Function MF_CAPTURE_METADATA_FRAME_ILLUMINATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6D688FFC, &H63D3, &H46FE, &HBA, &HDA, &H5B, &H94, &H7D, &HB0, &HD0, &H80)
MF_CAPTURE_METADATA_FRAME_ILLUMINATION = iid
End Function
Public Function MF_CAPTURE_METADATA_UVC_PAYLOADHEADER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF9F88A87, &HE1DD, &H441E, &H95, &HCB, &H42, &HE2, &H1A, &H64, &HF1, &HD9)
MF_CAPTURE_METADATA_UVC_PAYLOADHEADER = iid
End Function
Public Function MFSampleExtension_Depth_MinReliableDepth() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5F8582B2, &HE36B, &H47C8, &H9B, &H87, &HFE, &HE1, &HCA, &H72, &HC5, &HB0)
MFSampleExtension_Depth_MinReliableDepth = iid
End Function
Public Function MFSampleExtension_Depth_MaxReliableDepth() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE45545D1, &H1F0F, &H4A32, &HA8, &HA7, &H61, &H1, &HA2, &H4E, &HA8, &HBE)
MFSampleExtension_Depth_MaxReliableDepth = iid
End Function
Public Function MF_CAPTURE_METADATA_FIRST_SCANLINE_START_TIME_QPC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6A2C49F1, &HE052, &H46B6, &HB2, &HD9, &H73, &HC1, &H55, &H87, &H9, &HAF)
MF_CAPTURE_METADATA_FIRST_SCANLINE_START_TIME_QPC = iid
End Function
Public Function MF_CAPTURE_METADATA_LAST_SCANLINE_END_TIME_QPC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDCCADECB, &HC4D4, &H400D, &HB4, &H18, &H10, &HE8, &H85, &H25, &HE1, &HF6)
MF_CAPTURE_METADATA_LAST_SCANLINE_END_TIME_QPC = iid
End Function
Public Function MF_CAPTURE_METADATA_SCANLINE_TIME_QPC_ACCURACY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4CD79C51, &HF765, &H4B09, &HB1, &HE1, &H27, &HD1, &HF7, &HEB, &HEA, &H9)
MF_CAPTURE_METADATA_SCANLINE_TIME_QPC_ACCURACY = iid
End Function
Public Function MF_CAPTURE_METADATA_SCANLINE_DIRECTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6496A3BA, &H1907, &H49E6, &HB0, &HC3, &H12, &H37, &H95, &HF3, &H80, &HA9)
MF_CAPTURE_METADATA_SCANLINE_DIRECTION = iid
End Function
Public Function MF_CAPTURE_METADATA_DIGITALWINDOW() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H276F72A2, &H59C8, &H4F69, &H97, &HB4, &H6, &H8B, &H8C, &HE, &HC0, &H44)
MF_CAPTURE_METADATA_DIGITALWINDOW = iid
End Function
Public Function MF_CAPTURE_METADATA_FRAME_BACKGROUND_MASK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3F14DD3, &H75DD, &H433A, &HA8, &HE2, &H1E, &H3F, &H5F, &H2A, &H50, &HA0)
MF_CAPTURE_METADATA_FRAME_BACKGROUND_MASK = iid
End Function
Public Function MFT_CATEGORY_VIDEO_DECODER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD6C02D4B, &H6833, &H45B4, &H97, &H1A, &H5, &HA4, &HB0, &H4B, &HAB, &H91)
MFT_CATEGORY_VIDEO_DECODER = iid
End Function
Public Function MFT_CATEGORY_VIDEO_ENCODER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF79EAC7D, &HE545, &H4387, &HBD, &HEE, &HD6, &H47, &HD7, &HBD, &HE4, &H2A)
MFT_CATEGORY_VIDEO_ENCODER = iid
End Function
Public Function MFT_CATEGORY_VIDEO_EFFECT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H12E17C21, &H532C, &H4A6E, &H8A, &H1C, &H40, &H82, &H5A, &H73, &H63, &H97)
MFT_CATEGORY_VIDEO_EFFECT = iid
End Function
Public Function MFT_CATEGORY_MULTIPLEXER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H59C561E, &H5AE, &H4B61, &HB6, &H9D, &H55, &HB6, &H1E, &HE5, &H4A, &H7B)
MFT_CATEGORY_MULTIPLEXER = iid
End Function
Public Function MFT_CATEGORY_DEMULTIPLEXER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA8700A7A, &H939B, &H44C5, &H99, &HD7, &H76, &H22, &H6B, &H23, &HB3, &HF1)
MFT_CATEGORY_DEMULTIPLEXER = iid
End Function
Public Function MFT_CATEGORY_AUDIO_DECODER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9EA73FB4, &HEF7A, &H4559, &H8D, &H5D, &H71, &H9D, &H8F, &H4, &H26, &HC7)
MFT_CATEGORY_AUDIO_DECODER = iid
End Function
Public Function MFT_CATEGORY_AUDIO_ENCODER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H91C64BD0, &HF91E, &H4D8C, &H92, &H76, &HDB, &H24, &H82, &H79, &HD9, &H75)
MFT_CATEGORY_AUDIO_ENCODER = iid
End Function
Public Function MFT_CATEGORY_AUDIO_EFFECT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H11064C48, &H3648, &H4ED0, &H93, &H2E, &H5, &HCE, &H8A, &HC8, &H11, &HB7)
MFT_CATEGORY_AUDIO_EFFECT = iid
End Function
Public Function MFT_CATEGORY_VIDEO_PROCESSOR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H302EA3FC, &HAA5F, &H47F9, &H9F, &H7A, &HC2, &H18, &H8B, &HB1, &H63, &H2)
MFT_CATEGORY_VIDEO_PROCESSOR = iid
End Function
Public Function MFT_CATEGORY_OTHER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H90175D57, &HB7EA, &H4901, &HAE, &HB3, &H93, &H3A, &H87, &H47, &H75, &H6F)
MFT_CATEGORY_OTHER = iid
End Function
Public Function MFT_CATEGORY_ENCRYPTOR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB0C687BE, &H1CD, &H44B5, &HB8, &HB2, &H7C, &H1D, &H7E, &H5, &H8B, &H1F)
MFT_CATEGORY_ENCRYPTOR = iid
End Function
Public Function MFT_CATEGORY_VIDEO_RENDERER_EFFECT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H145CD8B4, &H92F4, &H4B23, &H8A, &HE7, &HE0, &HDF, &H6, &HC2, &HDA, &H95)
MFT_CATEGORY_VIDEO_RENDERER_EFFECT = iid
End Function
Public Function MFT_ENUM_VIDEO_RENDERER_EXTENSION_PROFILE() As UUID
'{62C56928-9A4E-443b-B9DC-CAC830C24100}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H62C56928, CInt(&H9A4E), CInt(&H443B), &HB9, &HDC, &HCA, &HC8, &H30, &HC2, &H41, &H0)
 MFT_ENUM_VIDEO_RENDERER_EXTENSION_PROFILE = iid
End Function
Public Function MFT_ENUM_ADAPTER_LUID() As UUID
'{1D39518C-E220-4DA8-A07F-BA172552D6B1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1D39518C, CInt(&HE220), CInt(&H4DA8), &HA0, &H7F, &HBA, &H17, &H25, &H52, &HD6, &HB1)
 MFT_ENUM_ADAPTER_LUID = iid
End Function
Public Function MFVideoFormat_H264_ES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3F40F4F0, &H5622, &H4FF8, &HB6, &HD8, &HA1, &H7A, &H58, &H4B, &HEE, &H5E)
MFVideoFormat_H264_ES = iid
End Function
Public Function MFVideoFormat_MPEG2() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE06D8026, &HDB46, &H11CF, &HB4, &HD1, &H0, &H80, &H5F, &H6C, &HBB, &HEA)
MFVideoFormat_MPEG2 = iid
End Function
Public Function MFVideoFormat_MPG2() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE06D8026, &HDB46, &H11CF, &HB4, &HD1, &H0, &H80, &H5F, &H6C, &HBB, &HEA)
MFVideoFormat_MPG2 = iid
End Function
Public Function MFAudioFormat_Dolby_AC3() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE06D802C, &HDB46, &H11CF, &HB4, &HD1, &H0, &H80, &H5F, &H6C, &HBB, &HEA)
MFAudioFormat_Dolby_AC3 = iid
End Function
Public Function MFAudioFormat_Dolby_DDPlus() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA7FB87AF, &H2D02, &H42FB, &HA4, &HD4, &H5, &HCD, &H93, &H84, &H3B, &HDD)
MFAudioFormat_Dolby_DDPlus = iid
End Function
Public Function MFAudioFormat_Dolby_AC4_V1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H36B7927C, &H3D87, &H4A2A, &H91, &H96, &HA2, &H1A, &HD9, &HE9, &H35, &HE6)
MFAudioFormat_Dolby_AC4_V1 = iid
End Function
Public Function MFAudioFormat_Dolby_AC4_V2() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7998B2A0, &H17DD, &H49B6, &H8D, &HFA, &H9B, &H27, &H85, &H52, &HA2, &HAC)
MFAudioFormat_Dolby_AC4_V2 = iid
End Function
Public Function MFAudioFormat_Dolby_AC4_V1_ES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9D8DCCC6, &HD156, &H4FB8, &H97, &H9C, &HA8, &H5B, &HE7, &HD2, &H1D, &HFA)
MFAudioFormat_Dolby_AC4_V1_ES = iid
End Function
Public Function MFAudioFormat_Dolby_AC4_V2_ES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7E58C9F9, &HB070, &H45F4, &H8C, &HCD, &HA9, &H9A, &H4, &H17, &HC1, &HAC)
MFAudioFormat_Dolby_AC4_V2_ES = iid
End Function
Public Function MFAudioFormat_MPEGH() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7C13C441, &HEBF8, &H4931, &HB6, &H78, &H80, &HB, &H19, &H24, &H22, &H36)
MFAudioFormat_MPEGH = iid
End Function
Public Function MFAudioFormat_MPEGH_ES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H19EE97FE, &H1BE0, &H4255, &HA8, &H76, &HE9, &H9F, &H53, &HA4, &H2A, &HE3)
MFAudioFormat_MPEGH_ES = iid
End Function
Public Function MFAudioFormat_Vorbis() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8D2FD10B, &H5841, &H4A6B, &H89, &H5, &H58, &H8F, &HEC, &H1A, &HDE, &HD9)
MFAudioFormat_Vorbis = iid
End Function
Public Function MFAudioFormat_DTS_RAW() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE06D8033, &HDB46, &H11CF, &HB4, &HD1, &H0, &H80, &H5F, &H6C, &HBB, &HEA)
MFAudioFormat_DTS_RAW = iid
End Function
Public Function MFAudioFormat_DTS_HD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA2E58EB7, &HFA9, &H48BB, &HA4, &HC, &HFA, &HE, &H15, &H6D, &H6, &H45)
MFAudioFormat_DTS_HD = iid
End Function
Public Function MFAudioFormat_DTS_XLL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H45B37C1B, &H8C70, &H4E59, &HA7, &HBE, &HA1, &HE4, &H2C, &H81, &HC8, &HD)
MFAudioFormat_DTS_XLL = iid
End Function
Public Function MFAudioFormat_DTS_LBR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC2FE6F0A, &H4E3C, &H4DF1, &H9B, &H60, &H50, &H86, &H30, &H91, &HE4, &HB9)
MFAudioFormat_DTS_LBR = iid
End Function
Public Function MFAudioFormat_DTS_UHD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H87020117, &HACE3, &H42DE, &HB7, &H3E, &HC6, &H56, &H70, &H62, &H63, &HF8)
MFAudioFormat_DTS_UHD = iid
End Function
Public Function MFAudioFormat_DTS_UHDY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9B9CCA00, &H91B9, &H4CCC, &H88, &H3A, &H8F, &H78, &H7A, &HC3, &HCC, &H86)
MFAudioFormat_DTS_UHDY = iid
End Function
Public Function MFAudioFormat_Float_SpatialObjects() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFA39CD94, &HBC64, &H4AB1, &H9B, &H71, &HDC, &HD0, &H9D, &H5A, &H7E, &H7A)
 MFAudioFormat_Float_SpatialObjects = iid
End Function
Public Function MFAudioFormat_LPCM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE06D8032, &HDB46, &H11CF, &HB4, &HD1, &H0, &H80, &H5F, &H6C, &HBB, &HEA)
MFAudioFormat_LPCM = iid
End Function
Public Function MFAudioFormat_PCM_HDCP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA5E7FF01, &H8411, &H4ACC, &HA8, &H65, &H5F, &H49, &H41, &H28, &H8D, &H80)
MFAudioFormat_PCM_HDCP = iid
End Function
Public Function MFAudioFormat_Dolby_AC3_HDCP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H97663A80, &H8FFB, &H4445, &HA6, &HBA, &H79, &H2D, &H90, &H8F, &H49, &H7F)
MFAudioFormat_Dolby_AC3_HDCP = iid
End Function
Public Function MFAudioFormat_AAC_HDCP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H419BCE76, &H8B72, &H400F, &HAD, &HEB, &H84, &HB5, &H7D, &H63, &H48, &H4D)
MFAudioFormat_AAC_HDCP = iid
End Function
Public Function MFAudioFormat_ADTS_HDCP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDA4963A3, &H14D8, &H4DCF, &H92, &HB7, &H19, &H3E, &HB8, &H43, &H63, &HDB)
MFAudioFormat_ADTS_HDCP = iid
End Function
Public Function MFAudioFormat_Base_HDCP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3884B5BC, &HE277, &H43FD, &H98, &H3D, &H3, &H8A, &HA8, &HD9, &HB6, &H5)
MFAudioFormat_Base_HDCP = iid
End Function
Public Function MFVideoFormat_H264_HDCP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5D0CE9DD, &H9817, &H49DA, &HBD, &HFD, &HF5, &HF5, &HB9, &H8F, &H18, &HA6)
MFVideoFormat_H264_HDCP = iid
End Function
Public Function MFVideoFormat_HEVC_HDCP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CFE0FE6, &H5C4, &H47DC, &H9D, &H70, &H4B, &HDB, &H29, &H59, &H72, &HF)
MFVideoFormat_HEVC_HDCP = iid
End Function
Public Function MFVideoFormat_Base_HDCP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEAC3B9D5, &HBD14, &H4237, &H8F, &H1F, &HBA, &HB4, &H28, &HE4, &H93, &H12)
MFVideoFormat_Base_HDCP = iid
End Function
Public Function MFMPEG4Format_Base() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H0, &H767A, &H494D, &HB4, &H78, &HF2, &H9D, &H25, &HDC, &H90, &H37)
MFMPEG4Format_Base = iid
End Function
Public Function MFSubtitleFormat_XML() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2006F94F, &H29CA, &H4195, &HB8, &HDB, &H0, &HDE, &HD8, &HFF, &HC, &H97)
MFSubtitleFormat_XML = iid
End Function
Public Function MFSubtitleFormat_TTML() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H73E73992, &H9A10, &H4356, &H95, &H57, &H71, &H94, &HE9, &H1E, &H3E, &H54)
MFSubtitleFormat_TTML = iid
End Function
Public Function MFSubtitleFormat_ATSC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7FA7FAA3, &HFEAE, &H4E16, &HAE, &HDF, &H36, &HB9, &HAC, &HFB, &HB0, &H99)
MFSubtitleFormat_ATSC = iid
End Function
Public Function MFSubtitleFormat_WebVTT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC886D215, &HF485, &H40BB, &H8D, &HB6, &HFA, &HDB, &HC6, &H19, &HA4, &H5D)
MFSubtitleFormat_WebVTT = iid
End Function
Public Function MFSubtitleFormat_SRT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5E467F2E, &H77CA, &H4CA5, &H83, &H91, &HD1, &H42, &HED, &H4B, &H76, &HC8)
MFSubtitleFormat_SRT = iid
End Function
Public Function MFSubtitleFormat_SSA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H57176A1B, &H1A9E, &H4EEA, &HAB, &HEF, &HC6, &H17, &H60, &H19, &H8A, &HC4)
MFSubtitleFormat_SSA = iid
End Function
Public Function MFSubtitleFormat_CustomUserData() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1BB3D849, &H6614, &H4D80, &H88, &H82, &HED, &H24, &HAA, &H82, &HDA, &H92)
MFSubtitleFormat_CustomUserData = iid
End Function
Public Function MFSubtitleFormat_PGS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H71F40E4A, &H1278, &H4442, &HB3, &HD, &H39, &HDD, &H1D, &H77, &H22, &HBC)
MFSubtitleFormat_PGS = iid
End Function
Public Function MFSubtitleFormat_VobSub() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6B8E40F4, &H8D2C, &H4CED, &HAD, &H91, &H59, &H60, &HE4, &H5B, &H44, &H33)
MFSubtitleFormat_VobSub = iid
End Function
Public Function MF_MT_MAJOR_TYPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H48EBA18E, &HF8C9, &H4687, &HBF, &H11, &HA, &H74, &HC9, &HF9, &H6A, &H8F)
MF_MT_MAJOR_TYPE = iid
End Function
Public Function MF_MT_SUBTYPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF7E34C9A, &H42E8, &H4714, &HB7, &H4B, &HCB, &H29, &HD7, &H2C, &H35, &HE5)
MF_MT_SUBTYPE = iid
End Function
Public Function MF_MT_ALL_SAMPLES_INDEPENDENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC9173739, &H5E56, &H461C, &HB7, &H13, &H46, &HFB, &H99, &H5C, &HB9, &H5F)
MF_MT_ALL_SAMPLES_INDEPENDENT = iid
End Function
Public Function MF_MT_FIXED_SIZE_SAMPLES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB8EBEFAF, &HB718, &H4E04, &HB0, &HA9, &H11, &H67, &H75, &HE3, &H32, &H1B)
MF_MT_FIXED_SIZE_SAMPLES = iid
End Function
Public Function MF_MT_COMPRESSED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3AFD0CEE, &H18F2, &H4BA5, &HA1, &H10, &H8B, &HEA, &H50, &H2E, &H1F, &H92)
MF_MT_COMPRESSED = iid
End Function
Public Function MF_MT_SAMPLE_SIZE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDAD3AB78, &H1990, &H408B, &HBC, &HE2, &HEB, &HA6, &H73, &HDA, &HCC, &H10)
MF_MT_SAMPLE_SIZE = iid
End Function
Public Function MF_MT_WRAPPED_TYPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D3F7B23, &HD02F, &H4E6C, &H9B, &HEE, &HE4, &HBF, &H2C, &H6C, &H69, &H5D)
MF_MT_WRAPPED_TYPE = iid
End Function
Public Function MF_MT_VIDEO_3D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCB5E88CF, &H7B5B, &H476B, &H85, &HAA, &H1C, &HA5, &HAE, &H18, &H75, &H55)
 MF_MT_VIDEO_3D = iid
End Function
Public Function MF_MT_VIDEO_3D_FORMAT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5315D8A0, &H87C5, &H4697, &HB7, &H93, &H66, &H6, &HC6, &H7C, &H4, &H9B)
MF_MT_VIDEO_3D_FORMAT = iid
End Function
Public Function MF_MT_VIDEO_3D_NUM_VIEWS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBB077E8A, &HDCBF, &H42EB, &HAF, &H60, &H41, &H8D, &HF9, &H8A, &HA4, &H95)
 MF_MT_VIDEO_3D_NUM_VIEWS = iid
End Function
Public Function MF_MT_VIDEO_3D_LEFT_IS_BASE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6D4B7BFF, &H5629, &H4404, &H94, &H8C, &HC6, &H34, &HF4, &HCE, &H26, &HD4)
 MF_MT_VIDEO_3D_LEFT_IS_BASE = iid
End Function
Public Function MF_MT_VIDEO_3D_FIRST_IS_LEFT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEC298493, &HADA, &H4EA1, &HA4, &HFE, &HCB, &HBD, &H36, &HCE, &H93, &H31)
 MF_MT_VIDEO_3D_FIRST_IS_LEFT = iid
End Function
Public Function MFSampleExtension_3DVideo() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF86F97A4, &HDD54, &H4E2E, &H9A, &H5E, &H55, &HFC, &H2D, &H74, &HA0, &H5)
 MFSampleExtension_3DVideo = iid
End Function
Public Function MFSampleExtension_3DVideo_SampleFormat() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8671772, &HE36F, &H4CFF, &H97, &HB3, &HD7, &H2E, &H20, &H98, &H7A, &H48)
 MFSampleExtension_3DVideo_SampleFormat = iid
End Function
Public Function MF_MT_VIDEO_ROTATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC380465D, &H2271, &H428C, &H9B, &H83, &HEC, &HEA, &H3B, &H4A, &H85, &HC1)
MF_MT_VIDEO_ROTATION = iid
End Function
Public Function MF_DEVICESTREAM_MULTIPLEXED_MANAGER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6EA542B0, &H281F, &H4231, &HA4, &H64, &HFE, &H2F, &H50, &H22, &H50, &H1C)
MF_DEVICESTREAM_MULTIPLEXED_MANAGER = iid
End Function
Public Function MF_MEDIATYPE_MULTIPLEXED_MANAGER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H13C78FB5, &HF275, &H4EA0, &HBB, &H5F, &H2, &H49, &H83, &H2B, &HD, &H6E)
MF_MEDIATYPE_MULTIPLEXED_MANAGER = iid
End Function
Public Function MFSampleExtension_MULTIPLEXED_MANAGER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8DCDEE79, &H6B5A, &H4C45, &H8D, &HB9, &H20, &HB3, &H95, &HF0, &H2F, &HCF)
MFSampleExtension_MULTIPLEXED_MANAGER = iid
End Function
Public Function MF_MT_SECURE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC5ACC4FD, &H304, &H4ECF, &H80, &H9F, &H47, &HBC, &H97, &HFF, &H63, &HBD)
MF_MT_SECURE = iid
End Function
Public Function MF_DEVICESTREAM_ATTRIBUTE_FRAMESOURCE_TYPES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H17145FD1, &H1B2B, &H423C, &H80, &H1, &H2B, &H68, &H33, &HED, &H35, &H88)
MF_DEVICESTREAM_ATTRIBUTE_FRAMESOURCE_TYPES = iid
End Function
Public Function MF_MT_ALPHA_MODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5D959B0D, &H4CBF, &H4D04, &H91, &H9F, &H3F, &H5F, &H7F, &H28, &H42, &H11)
MF_MT_ALPHA_MODE = iid
End Function
Public Function MF_MT_DEPTH_MEASUREMENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFD5AC489, &H917, &H4BB6, &H9D, &H54, &H31, &H22, &HBF, &H70, &H14, &H4B)
MF_MT_DEPTH_MEASUREMENT = iid
End Function
Public Function MF_MT_DEPTH_VALUE_UNIT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H21A800F5, &H3189, &H4797, &HBE, &HBA, &HF1, &H3C, &HD9, &HA3, &H1A, &H5E)
MF_MT_DEPTH_VALUE_UNIT = iid
End Function
Public Function MF_MT_VIDEO_NO_FRAME_ORDERING() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3F5B106F, &H6BC2, &H4EE3, &HB7, &HED, &H89, &H2, &HC1, &H8F, &H53, &H51)
MF_MT_VIDEO_NO_FRAME_ORDERING = iid
End Function
Public Function MF_MT_VIDEO_H264_NO_FMOASO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HED461CD6, &HEC9F, &H416A, &HA8, &HA3, &H26, &HD7, &HD3, &H10, &H18, &HD7)
MF_MT_VIDEO_H264_NO_FMOASO = iid
End Function
Public Function MFSampleExtension_ForwardedDecodeUnits() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H424C754C, &H97C8, &H48D6, &H87, &H77, &HFC, &H41, &HF7, &HB6, &H8, &H79)
MFSampleExtension_ForwardedDecodeUnits = iid
End Function
Public Function MFSampleExtension_TargetGlobalLuminance() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3F60EF36, &H31EF, &H4DAF, &H83, &H60, &H94, &H3, &H97, &HE4, &H1E, &HF3)
MFSampleExtension_TargetGlobalLuminance = iid
End Function
Public Function MFSampleExtension_ForwardedDecodeUnitType() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H89E57C7, &H47D3, &H4A26, &HBF, &H9C, &H4B, &H64, &HFA, &HFB, &H5D, &H1E)
MFSampleExtension_ForwardedDecodeUnitType = iid
End Function
Public Function MF_MT_FORWARD_CUSTOM_NALU() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HED336EFD, &H244F, &H428D, &H91, &H53, &H28, &HF3, &H99, &H45, &H88, &H90)
MF_MT_FORWARD_CUSTOM_NALU = iid
End Function
Public Function MF_MT_FORWARD_CUSTOM_SEI() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE27362F1, &HB136, &H41D1, &H95, &H94, &H3A, &H7E, &H4F, &HEB, &HF2, &HD1)
MF_MT_FORWARD_CUSTOM_SEI = iid
End Function
Public Function MF_MT_VIDEO_RENDERER_EXTENSION_PROFILE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8437D4B9, &HD448, &H4FCD, &H9B, &H6B, &H83, &H9B, &HF9, &H6C, &H77, &H98)
MF_MT_VIDEO_RENDERER_EXTENSION_PROFILE = iid
End Function
Public Function MF_DECODER_FWD_CUSTOM_SEI_DECODE_ORDER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF13BBE3C, &H36D4, &H410A, &HB9, &H85, &H7A, &H95, &H1A, &H1E, &H62, &H94)
MF_DECODER_FWD_CUSTOM_SEI_DECODE_ORDER = iid
End Function
Public Function MF_VIDEO_RENDERER_EFFECT_APP_SERVICE_NAME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC6052A80, &H6D9C, &H40A3, &H9D, &HB8, &HF0, &H27, &HA2, &H5C, &H9A, &HB9)
MF_VIDEO_RENDERER_EFFECT_APP_SERVICE_NAME = iid
End Function
Public Function MF_MT_AUDIO_NUM_CHANNELS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H37E48BF5, &H645E, &H4C5B, &H89, &HDE, &HAD, &HA9, &HE2, &H9B, &H69, &H6A)
MF_MT_AUDIO_NUM_CHANNELS = iid
End Function
Public Function MF_MT_AUDIO_SAMPLES_PER_SECOND() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5FAEEAE7, &H290, &H4C31, &H9E, &H8A, &HC5, &H34, &HF6, &H8D, &H9D, &HBA)
MF_MT_AUDIO_SAMPLES_PER_SECOND = iid
End Function
Public Function MF_MT_AUDIO_FLOAT_SAMPLES_PER_SECOND() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFB3B724A, &HCFB5, &H4319, &HAE, &HFE, &H6E, &H42, &HB2, &H40, &H61, &H32)
MF_MT_AUDIO_FLOAT_SAMPLES_PER_SECOND = iid
End Function
Public Function MF_MT_AUDIO_AVG_BYTES_PER_SECOND() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1AAB75C8, &HCFEF, &H451C, &HAB, &H95, &HAC, &H3, &H4B, &H8E, &H17, &H31)
MF_MT_AUDIO_AVG_BYTES_PER_SECOND = iid
End Function
Public Function MF_MT_AUDIO_BLOCK_ALIGNMENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H322DE230, &H9EEB, &H43BD, &HAB, &H7A, &HFF, &H41, &H22, &H51, &H54, &H1D)
MF_MT_AUDIO_BLOCK_ALIGNMENT = iid
End Function
Public Function MF_MT_AUDIO_BITS_PER_SAMPLE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF2DEB57F, &H40FA, &H4764, &HAA, &H33, &HED, &H4F, &H2D, &H1F, &HF6, &H69)
MF_MT_AUDIO_BITS_PER_SAMPLE = iid
End Function
Public Function MF_MT_AUDIO_VALID_BITS_PER_SAMPLE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD9BF8D6A, &H9530, &H4B7C, &H9D, &HDF, &HFF, &H6F, &HD5, &H8B, &HBD, &H6)
MF_MT_AUDIO_VALID_BITS_PER_SAMPLE = iid
End Function
Public Function MF_MT_AUDIO_SAMPLES_PER_BLOCK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAAB15AAC, &HE13A, &H4995, &H92, &H22, &H50, &H1E, &HA1, &H5C, &H68, &H77)
MF_MT_AUDIO_SAMPLES_PER_BLOCK = iid
End Function
Public Function MF_MT_AUDIO_CHANNEL_MASK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H55FB5765, &H644A, &H4CAF, &H84, &H79, &H93, &H89, &H83, &HBB, &H15, &H88)
MF_MT_AUDIO_CHANNEL_MASK = iid
End Function
Public Function MF_MT_AUDIO_FOLDDOWN_MATRIX() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9D62927C, &H36BE, &H4CF2, &HB5, &HC4, &HA3, &H92, &H6E, &H3E, &H87, &H11)
MF_MT_AUDIO_FOLDDOWN_MATRIX = iid
End Function
Public Function MF_MT_AUDIO_WMADRC_PEAKREF() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9D62927D, &H36BE, &H4CF2, &HB5, &HC4, &HA3, &H92, &H6E, &H3E, &H87, &H11)
MF_MT_AUDIO_WMADRC_PEAKREF = iid
End Function
Public Function MF_MT_AUDIO_WMADRC_PEAKTARGET() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9D62927E, &H36BE, &H4CF2, &HB5, &HC4, &HA3, &H92, &H6E, &H3E, &H87, &H11)
MF_MT_AUDIO_WMADRC_PEAKTARGET = iid
End Function
Public Function MF_MT_AUDIO_WMADRC_AVGREF() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9D62927F, &H36BE, &H4CF2, &HB5, &HC4, &HA3, &H92, &H6E, &H3E, &H87, &H11)
MF_MT_AUDIO_WMADRC_AVGREF = iid
End Function
Public Function MF_MT_AUDIO_WMADRC_AVGTARGET() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9D629280, &H36BE, &H4CF2, &HB5, &HC4, &HA3, &H92, &H6E, &H3E, &H87, &H11)
MF_MT_AUDIO_WMADRC_AVGTARGET = iid
End Function
Public Function MF_MT_AUDIO_PREFER_WAVEFORMATEX() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA901AABA, &HE037, &H458A, &HBD, &HF6, &H54, &H5B, &HE2, &H7, &H40, &H42)
MF_MT_AUDIO_PREFER_WAVEFORMATEX = iid
End Function
Public Function MF_MT_AAC_PAYLOAD_TYPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBFBABE79, &H7434, &H4D1C, &H94, &HF0, &H72, &HA3, &HB9, &HE1, &H71, &H88)
MF_MT_AAC_PAYLOAD_TYPE = iid
End Function
Public Function MF_MT_AAC_AUDIO_PROFILE_LEVEL_INDICATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7632F0E6, &H9538, &H4D61, &HAC, &HDA, &HEA, &H29, &HC8, &HC1, &H44, &H56)
MF_MT_AAC_AUDIO_PROFILE_LEVEL_INDICATION = iid
End Function
Public Function MF_MT_AUDIO_FLAC_MAX_BLOCK_SIZE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8B81ADAE, &H4B5A, &H4D40, &H80, &H22, &HF3, &H8D, &H9, &HCA, &H3C, &H5C)
MF_MT_AUDIO_FLAC_MAX_BLOCK_SIZE = iid
End Function
Public Function MF_MT_SPATIAL_AUDIO_MAX_DYNAMIC_OBJECTS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDCFBA24A, &H2609, &H4240, &HA7, &H21, &H3F, &HAE, &HA7, &H6A, &H4D, &HF9)
 MF_MT_SPATIAL_AUDIO_MAX_DYNAMIC_OBJECTS = iid
End Function
Public Function MF_MT_SPATIAL_AUDIO_OBJECT_METADATA_FORMAT_ID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2AB71BC0, &H6223, &H4BA7, &HAD, &H64, &H7B, &H94, &HB4, &H7A, &HE7, &H92)
 MF_MT_SPATIAL_AUDIO_OBJECT_METADATA_FORMAT_ID = iid
End Function
Public Function MF_MT_SPATIAL_AUDIO_OBJECT_METADATA_LENGTH() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H94BA8BE, &HD723, &H489F, &H92, &HFA, &H76, &H67, &H77, &HB3, &H47, &H26)
 MF_MT_SPATIAL_AUDIO_OBJECT_METADATA_LENGTH = iid
End Function
Public Function MF_MT_SPATIAL_AUDIO_MAX_METADATA_ITEMS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H11AA80B4, &HE0DA, &H47C6, &H80, &H60, &H96, &HC1, &H25, &H9A, &HE5, &HD)
 MF_MT_SPATIAL_AUDIO_MAX_METADATA_ITEMS = iid
End Function
Public Function MF_MT_SPATIAL_AUDIO_MIN_METADATA_ITEM_OFFSET_SPACING() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H83E96EC9, &H1184, &H417E, &H82, &H54, &H9F, &H26, &H91, &H58, &HFC, &H6)
 MF_MT_SPATIAL_AUDIO_MIN_METADATA_ITEM_OFFSET_SPACING = iid
End Function
Public Function MF_MT_SPATIAL_AUDIO_DATA_PRESENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6842F6E7, &HD43E, &H4EBB, &H9C, &H9C, &HC9, &H6F, &H41, &H78, &H48, &H63)
 MF_MT_SPATIAL_AUDIO_DATA_PRESENT = iid
End Function
Public Function MF_MT_SPATIAL_AUDIO_IS_PREVIRTUALIZED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4EACAB51, &HFFE5, &H421A, &HA2, &HA7, &H8B, &H74, &H9, &HA1, &HCA, &HC4)
MF_MT_SPATIAL_AUDIO_IS_PREVIRTUALIZED = iid
End Function
Public Function MF_MT_MPEGH_AUDIO_PROFILE_LEVEL_INDICATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H51267A39, &HDD0C, &H4BB9, &HA7, &HBD, &H91, &H73, &HAD, &H4B, &H13, &H1C)
MF_MT_MPEGH_AUDIO_PROFILE_LEVEL_INDICATION = iid
End Function
Public Function MF_MT_FRAME_SIZE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1652C33D, &HD6B2, &H4012, &HB8, &H34, &H72, &H3, &H8, &H49, &HA3, &H7D)
MF_MT_FRAME_SIZE = iid
End Function
Public Function MF_MT_FRAME_RATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC459A2E8, &H3D2C, &H4E44, &HB1, &H32, &HFE, &HE5, &H15, &H6C, &H7B, &HB0)
MF_MT_FRAME_RATE = iid
End Function
Public Function MF_MT_PIXEL_ASPECT_RATIO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC6376A1E, &H8D0A, &H4027, &HBE, &H45, &H6D, &H9A, &HA, &HD3, &H9B, &HB6)
MF_MT_PIXEL_ASPECT_RATIO = iid
End Function
Public Function MF_MT_DRM_FLAGS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8772F323, &H355A, &H4CC7, &HBB, &H78, &H6D, &H61, &HA0, &H48, &HAE, &H82)
MF_MT_DRM_FLAGS = iid
End Function
Public Function MF_MT_TIMESTAMP_CAN_BE_DTS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H24974215, &H1B7B, &H41E4, &H86, &H25, &HAC, &H46, &H9F, &H2D, &HED, &HAA)
MF_MT_TIMESTAMP_CAN_BE_DTS = iid
End Function
Public Function MF_MT_PAD_CONTROL_FLAGS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D0E73E5, &H80EA, &H4354, &HA9, &HD0, &H11, &H76, &HCE, &HB0, &H28, &HEA)
MF_MT_PAD_CONTROL_FLAGS = iid
End Function
Public Function MF_MT_SOURCE_CONTENT_HINT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H68ACA3CC, &H22D0, &H44E6, &H85, &HF8, &H28, &H16, &H71, &H97, &HFA, &H38)
MF_MT_SOURCE_CONTENT_HINT = iid
End Function
Public Function MF_MT_VIDEO_CHROMA_SITING() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H65DF2370, &HC773, &H4C33, &HAA, &H64, &H84, &H3E, &H6, &H8E, &HFB, &HC)
MF_MT_VIDEO_CHROMA_SITING = iid
End Function
Public Function MF_MT_INTERLACE_MODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE2724BB8, &HE676, &H4806, &HB4, &HB2, &HA8, &HD6, &HEF, &HB4, &H4C, &HCD)
MF_MT_INTERLACE_MODE = iid
End Function
Public Function MF_MT_TRANSFER_FUNCTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5FB0FCE9, &HBE5C, &H4935, &HA8, &H11, &HEC, &H83, &H8F, &H8E, &HED, &H93)
MF_MT_TRANSFER_FUNCTION = iid
End Function
Public Function MF_MT_VIDEO_PRIMARIES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDBFBE4D7, &H740, &H4EE0, &H81, &H92, &H85, &HA, &HB0, &HE2, &H19, &H35)
MF_MT_VIDEO_PRIMARIES = iid
End Function
Public Function MF_MT_MAX_LUMINANCE_LEVEL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H50253128, &HC110, &H4DE4, &H98, &HAE, &H46, &HA3, &H24, &HFA, &HE6, &HDA)
MF_MT_MAX_LUMINANCE_LEVEL = iid
End Function
Public Function MF_MT_MAX_FRAME_AVERAGE_LUMINANCE_LEVEL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H58D4BF57, &H6F52, &H4733, &HA1, &H95, &HA9, &HE2, &H9E, &HCF, &H9E, &H27)
MF_MT_MAX_FRAME_AVERAGE_LUMINANCE_LEVEL = iid
End Function
Public Function MF_MT_MAX_MASTERING_LUMINANCE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD6C6B997, &H272F, &H4CA1, &H8D, &H0, &H80, &H42, &H11, &H1A, &HF, &HF6)
MF_MT_MAX_MASTERING_LUMINANCE = iid
End Function
Public Function MF_MT_MIN_MASTERING_LUMINANCE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H839A4460, &H4E7E, &H4B4F, &HAE, &H79, &HCC, &H8, &H90, &H5C, &H7B, &H27)
MF_MT_MIN_MASTERING_LUMINANCE = iid
End Function
Public Function MF_MT_DECODER_USE_MAX_RESOLUTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4C547C24, &HAF9A, &H4F38, &H96, &HAD, &H97, &H87, &H73, &HCF, &H53, &HE7)
MF_MT_DECODER_USE_MAX_RESOLUTION = iid
End Function
Public Function MF_MT_DECODER_MAX_DPB_COUNT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H67BE144C, &H88B7, &H4CA9, &H96, &H28, &HC8, &H8, &HD5, &H26, &H22, &H17)
MF_MT_DECODER_MAX_DPB_COUNT = iid
End Function
Public Function MF_MT_CUSTOM_VIDEO_PRIMARIES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H47537213, &H8CFB, &H4722, &HAA, &H34, &HFB, &HC9, &HE2, &H4D, &H77, &HB8)
MF_MT_CUSTOM_VIDEO_PRIMARIES = iid
End Function
Public Function MF_MT_YUV_MATRIX() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3E23D450, &H2C75, &H4D25, &HA0, &HE, &HB9, &H16, &H70, &HD1, &H23, &H27)
MF_MT_YUV_MATRIX = iid
End Function
Public Function MF_MT_VIDEO_LIGHTING() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53A0529C, &H890B, &H4216, &H8B, &HF9, &H59, &H93, &H67, &HAD, &H6D, &H20)
MF_MT_VIDEO_LIGHTING = iid
End Function
Public Function MF_MT_VIDEO_NOMINAL_RANGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC21B8EE5, &HB956, &H4071, &H8D, &HAF, &H32, &H5E, &HDF, &H5C, &HAB, &H11)
MF_MT_VIDEO_NOMINAL_RANGE = iid
End Function
Public Function MF_MT_GEOMETRIC_APERTURE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H66758743, &H7E5F, &H400D, &H98, &HA, &HAA, &H85, &H96, &HC8, &H56, &H96)
MF_MT_GEOMETRIC_APERTURE = iid
End Function
Public Function MF_MT_MINIMUM_DISPLAY_APERTURE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD7388766, &H18FE, &H48C6, &HA1, &H77, &HEE, &H89, &H48, &H67, &HC8, &HC4)
MF_MT_MINIMUM_DISPLAY_APERTURE = iid
End Function
Public Function MF_MT_PAN_SCAN_APERTURE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H79614DDE, &H9187, &H48FB, &HB8, &HC7, &H4D, &H52, &H68, &H9D, &HE6, &H49)
MF_MT_PAN_SCAN_APERTURE = iid
End Function
Public Function MF_MT_PAN_SCAN_ENABLED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4B7F6BC3, &H8B13, &H40B2, &HA9, &H93, &HAB, &HF6, &H30, &HB8, &H20, &H4E)
MF_MT_PAN_SCAN_ENABLED = iid
End Function
Public Function MF_MT_AVG_BITRATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H20332624, &HFB0D, &H4D9E, &HBD, &HD, &HCB, &HF6, &H78, &H6C, &H10, &H2E)
MF_MT_AVG_BITRATE = iid
End Function
Public Function MF_MT_AVG_BIT_ERROR_RATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H799CABD6, &H3508, &H4DB4, &HA3, &HC7, &H56, &H9C, &HD5, &H33, &HDE, &HB1)
MF_MT_AVG_BIT_ERROR_RATE = iid
End Function
Public Function MF_MT_MAX_KEYFRAME_SPACING() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC16EB52B, &H73A1, &H476F, &H8D, &H62, &H83, &H9D, &H6A, &H2, &H6, &H52)
MF_MT_MAX_KEYFRAME_SPACING = iid
End Function
Public Function MF_MT_USER_DATA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB6BC765F, &H4C3B, &H40A4, &HBD, &H51, &H25, &H35, &HB6, &H6F, &HE0, &H9D)
MF_MT_USER_DATA = iid
End Function
Public Function MF_MT_OUTPUT_BUFFER_NUM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA505D3AC, &HF930, &H436E, &H8E, &HDE, &H93, &HA5, &H9, &HCE, &H23, &HB2)
MF_MT_OUTPUT_BUFFER_NUM = iid
End Function
Public Function MF_MT_REALTIME_CONTENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBB12D222, &H2BDB, &H425E, &H91, &HEC, &H23, &H8, &HE1, &H89, &HA5, &H8F)
MF_MT_REALTIME_CONTENT = iid
End Function
Public Function MF_MT_DEFAULT_STRIDE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H644B4E48, &H1E02, &H4516, &HB0, &HEB, &HC0, &H1C, &HA9, &HD4, &H9A, &HC6)
MF_MT_DEFAULT_STRIDE = iid
End Function
Public Function MF_MT_PALETTE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6D283F42, &H9846, &H4410, &HAF, &HD9, &H65, &H4D, &H50, &H3B, &H1A, &H54)
MF_MT_PALETTE = iid
End Function
Public Function MF_MT_AM_FORMAT_TYPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H73D1072D, &H1870, &H4174, &HA0, &H63, &H29, &HFF, &H4F, &HF6, &HC1, &H1E)
MF_MT_AM_FORMAT_TYPE = iid
End Function
Public Function MF_MT_VIDEO_PROFILE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAD76A80B, &H2D5C, &H4E0B, &HB3, &H75, &H64, &HE5, &H20, &H13, &H70, &H36)
MF_MT_VIDEO_PROFILE = iid
End Function
Public Function MF_MT_VIDEO_LEVEL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H96F66574, &H11C5, &H4015, &H86, &H66, &HBF, &HF5, &H16, &H43, &H6D, &HA7)
MF_MT_VIDEO_LEVEL = iid
End Function
Public Function MF_MT_MPEG_START_TIME_CODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H91F67885, &H4333, &H4280, &H97, &HCD, &HBD, &H5A, &H6C, &H3, &HA0, &H6E)
MF_MT_MPEG_START_TIME_CODE = iid
End Function
Public Function MF_MT_MPEG2_PROFILE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAD76A80B, &H2D5C, &H4E0B, &HB3, &H75, &H64, &HE5, &H20, &H13, &H70, &H36)
MF_MT_MPEG2_PROFILE = iid
End Function
Public Function MF_MT_MPEG2_LEVEL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H96F66574, &H11C5, &H4015, &H86, &H66, &HBF, &HF5, &H16, &H43, &H6D, &HA7)
MF_MT_MPEG2_LEVEL = iid
End Function
Public Function MF_MT_MPEG2_FLAGS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H31E3991D, &HF701, &H4B2F, &HB4, &H26, &H8A, &HE3, &HBD, &HA9, &HE0, &H4B)
MF_MT_MPEG2_FLAGS = iid
End Function
Public Function MF_MT_MPEG_SEQUENCE_HEADER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3C036DE7, &H3AD0, &H4C9E, &H92, &H16, &HEE, &H6D, &H6A, &HC2, &H1C, &HB3)
MF_MT_MPEG_SEQUENCE_HEADER = iid
End Function
Public Function MF_MT_MPEG2_STANDARD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA20AF9E8, &H928A, &H4B26, &HAA, &HA9, &HF0, &H5C, &H74, &HCA, &HC4, &H7C)
MF_MT_MPEG2_STANDARD = iid
End Function
Public Function MF_MT_MPEG2_TIMECODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5229BA10, &HE29D, &H4F80, &HA5, &H9C, &HDF, &H4F, &H18, &H2, &H7, &HD2)
MF_MT_MPEG2_TIMECODE = iid
End Function
Public Function MF_MT_MPEG2_CONTENT_PACKET() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H825D55E4, &H4F12, &H4197, &H9E, &HB3, &H59, &HB6, &HE4, &H71, &HF, &H6)
MF_MT_MPEG2_CONTENT_PACKET = iid
End Function
Public Function MF_MT_MPEG2_ONE_FRAME_PER_PACKET() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H91A49EB5, &H1D20, &H4B42, &HAC, &HE8, &H80, &H42, &H69, &HBF, &H95, &HED)
MF_MT_MPEG2_ONE_FRAME_PER_PACKET = iid
End Function
Public Function MF_MT_MPEG2_HDCP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H168F1B4A, &H3E91, &H450F, &HAE, &HA7, &HE4, &HBA, &HEA, &HDA, &HE5, &HBA)
MF_MT_MPEG2_HDCP = iid
End Function
Public Function MF_MT_H264_MAX_CODEC_CONFIG_DELAY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF5929986, &H4C45, &H4FBB, &HBB, &H49, &H6C, &HC5, &H34, &HD0, &H5B, &H9B)
MF_MT_H264_MAX_CODEC_CONFIG_DELAY = iid
End Function
Public Function MF_MT_H264_SUPPORTED_SLICE_MODES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC8BE1937, &H4D64, &H4549, &H83, &H43, &HA8, &H8, &H6C, &HB, &HFD, &HA5)
MF_MT_H264_SUPPORTED_SLICE_MODES = iid
End Function
Public Function MF_MT_H264_SUPPORTED_SYNC_FRAME_TYPES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H89A52C01, &HF282, &H48D2, &HB5, &H22, &H22, &HE6, &HAE, &H63, &H31, &H99)
MF_MT_H264_SUPPORTED_SYNC_FRAME_TYPES = iid
End Function
Public Function MF_MT_H264_RESOLUTION_SCALING() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE3854272, &HF715, &H4757, &HBA, &H90, &H1B, &H69, &H6C, &H77, &H34, &H57)
MF_MT_H264_RESOLUTION_SCALING = iid
End Function
Public Function MF_MT_H264_SIMULCAST_SUPPORT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9EA2D63D, &H53F0, &H4A34, &HB9, &H4E, &H9D, &HE4, &H9A, &H7, &H8C, &HB3)
MF_MT_H264_SIMULCAST_SUPPORT = iid
End Function
Public Function MF_MT_H264_SUPPORTED_RATE_CONTROL_MODES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6A8AC47E, &H519C, &H4F18, &H9B, &HB3, &H7E, &HEA, &HAE, &HA5, &H59, &H4D)
MF_MT_H264_SUPPORTED_RATE_CONTROL_MODES = iid
End Function
Public Function MF_MT_H264_MAX_MB_PER_SEC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H45256D30, &H7215, &H4576, &H93, &H36, &HB0, &HF1, &HBC, &HD5, &H9B, &HB2)
MF_MT_H264_MAX_MB_PER_SEC = iid
End Function
Public Function MF_MT_H264_SUPPORTED_USAGES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H60B1A998, &HDC01, &H40CE, &H97, &H36, &HAB, &HA8, &H45, &HA2, &HDB, &HDC)
MF_MT_H264_SUPPORTED_USAGES = iid
End Function
Public Function MF_MT_H264_CAPABILITIES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBB3BD508, &H490A, &H11E0, &H99, &HE4, &H13, &H16, &HDF, &HD7, &H20, &H85)
MF_MT_H264_CAPABILITIES = iid
End Function
Public Function MF_MT_H264_SVC_CAPABILITIES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF8993ABE, &HD937, &H4A8F, &HBB, &HCA, &H69, &H66, &HFE, &H9E, &H11, &H52)
MF_MT_H264_SVC_CAPABILITIES = iid
End Function
Public Function MF_MT_H264_USAGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H359CE3A5, &HAF00, &H49CA, &HA2, &HF4, &H2A, &HC9, &H4C, &HA8, &H2B, &H61)
MF_MT_H264_USAGE = iid
End Function
Public Function MF_MT_H264_RATE_CONTROL_MODES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H705177D8, &H45CB, &H11E0, &HAC, &H7D, &HB9, &H1C, &HE0, &HD7, &H20, &H85)
MF_MT_H264_RATE_CONTROL_MODES = iid
End Function
Public Function MF_MT_H264_LAYOUT_PER_STREAM() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H85E299B2, &H90E3, &H4FE8, &HB2, &HF5, &HC0, &H67, &HE0, &HBF, &HE5, &H7A)
MF_MT_H264_LAYOUT_PER_STREAM = iid
End Function
Public Function MF_MT_IN_BAND_PARAMETER_SET() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H75DA5090, &H910B, &H4A03, &H89, &H6C, &H7B, &H89, &H8F, &HEE, &HA5, &HAF)
MF_MT_IN_BAND_PARAMETER_SET = iid
End Function
Public Function MF_MT_MPEG4_TRACK_TYPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H54F486DD, &H9327, &H4F6D, &H80, &HAB, &H6F, &H70, &H9E, &HBB, &H4C, &HCE)
MF_MT_MPEG4_TRACK_TYPE = iid
End Function
Public Function MF_MT_CONTAINER_RATE_SCALING() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H83877F5E, &H444, &H4E28, &H84, &H79, &H6D, &HB0, &H98, &H9B, &H8C, &H9)
MF_MT_CONTAINER_RATE_SCALING = iid
End Function
Public Function MF_MT_DV_AAUX_SRC_PACK_0() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H84BD5D88, &HFB8, &H4AC8, &HBE, &H4B, &HA8, &H84, &H8B, &HEF, &H98, &HF3)
MF_MT_DV_AAUX_SRC_PACK_0 = iid
End Function
Public Function MF_MT_DV_AAUX_CTRL_PACK_0() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF731004E, &H1DD1, &H4515, &HAA, &HBE, &HF0, &HC0, &H6A, &HA5, &H36, &HAC)
MF_MT_DV_AAUX_CTRL_PACK_0 = iid
End Function
Public Function MF_MT_DV_AAUX_SRC_PACK_1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H720E6544, &H225, &H4003, &HA6, &H51, &H1, &H96, &H56, &H3A, &H95, &H8E)
MF_MT_DV_AAUX_SRC_PACK_1 = iid
End Function
Public Function MF_MT_DV_AAUX_CTRL_PACK_1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCD1F470D, &H1F04, &H4FE0, &HBF, &HB9, &HD0, &H7A, &HE0, &H38, &H6A, &HD8)
MF_MT_DV_AAUX_CTRL_PACK_1 = iid
End Function
Public Function MF_MT_DV_VAUX_SRC_PACK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H41402D9D, &H7B57, &H43C6, &HB1, &H29, &H2C, &HB9, &H97, &HF1, &H50, &H9)
MF_MT_DV_VAUX_SRC_PACK = iid
End Function
Public Function MF_MT_DV_VAUX_CTRL_PACK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2F84E1C4, &HDA1, &H4788, &H93, &H8E, &HD, &HFB, &HFB, &HB3, &H4B, &H48)
MF_MT_DV_VAUX_CTRL_PACK = iid
End Function
Public Function MF_MT_ARBITRARY_HEADER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9E6BD6F5, &H109, &H4F95, &H84, &HAC, &H93, &H9, &H15, &H3A, &H19, &HFC)
MF_MT_ARBITRARY_HEADER = iid
End Function
Public Function MF_MT_ARBITRARY_FORMAT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5A75B249, &HD7D, &H49A1, &HA1, &HC3, &HE0, &HD8, &H7F, &HC, &HAD, &HE5)
MF_MT_ARBITRARY_FORMAT = iid
End Function
Public Function MF_MT_IMAGE_LOSS_TOLERANT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HED062CF4, &HE34E, &H4922, &HBE, &H99, &H93, &H40, &H32, &H13, &H3D, &H7C)
MF_MT_IMAGE_LOSS_TOLERANT = iid
End Function
Public Function MF_MT_MPEG4_SAMPLE_DESCRIPTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H261E9D83, &H9529, &H4B8F, &HA1, &H11, &H8B, &H9C, &H95, &HA, &H81, &HA9)
MF_MT_MPEG4_SAMPLE_DESCRIPTION = iid
End Function
Public Function MF_MT_MPEG4_CURRENT_SAMPLE_ENTRY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9AA7E155, &HB64A, &H4C1D, &HA5, &H0, &H45, &H5D, &H60, &HB, &H65, &H60)
MF_MT_MPEG4_CURRENT_SAMPLE_ENTRY = iid
End Function
Public Function MF_SD_AMBISONICS_SAMPLE3D_DESCRIPTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF715CF3E, &HA964, &H4C3F, &H94, &HAE, &H9D, &H6B, &HA7, &H26, &H46, &H41)
MF_SD_AMBISONICS_SAMPLE3D_DESCRIPTION = iid
End Function
Public Function MF_MT_ORIGINAL_4CC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD7BE3FE0, &H2BC7, &H492D, &HB8, &H43, &H61, &HA1, &H91, &H9B, &H70, &HC3)
MF_MT_ORIGINAL_4CC = iid
End Function
Public Function MF_MT_ORIGINAL_WAVE_FORMAT_TAG() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8CBBC843, &H9FD9, &H49C2, &H88, &H2F, &HA7, &H25, &H86, &HC4, &H8, &HAD)
MF_MT_ORIGINAL_WAVE_FORMAT_TAG = iid
End Function
Public Function MF_MT_FRAME_RATE_RANGE_MIN() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD2E7558C, &HDC1F, &H403F, &H9A, &H72, &HD2, &H8B, &HB1, &HEB, &H3B, &H5E)
MF_MT_FRAME_RATE_RANGE_MIN = iid
End Function
Public Function MF_MT_FRAME_RATE_RANGE_MAX() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE3371D41, &HB4CF, &H4A05, &HBD, &H4E, &H20, &HB8, &H8B, &HB2, &HC4, &HD6)
MF_MT_FRAME_RATE_RANGE_MAX = iid
End Function
Public Function MF_LOW_LATENCY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9C27891A, &HED7A, &H40E1, &H88, &HE8, &HB2, &H27, &H27, &HA0, &H24, &HEE)
MF_LOW_LATENCY = iid
End Function
Public Function MF_VIDEO_MAX_MB_PER_SEC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE3F2E203, &HD445, &H4B8C, &H92, &H11, &HAE, &H39, &HD, &H3B, &HA0, &H17)
MF_VIDEO_MAX_MB_PER_SEC = iid
End Function
Public Function MF_DISABLE_FRAME_CORRUPTION_INFO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7086E16C, &H49C5, &H4201, &H88, &H2A, &H85, &H38, &HF3, &H8C, &HF1, &H3A)
MF_DISABLE_FRAME_CORRUPTION_INFO = iid
End Function
Public Function MFStreamExtension_CameraExtrinsics() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H686196D0, &H13E2, &H41D9, &H96, &H38, &HEF, &H3, &H2C, &H27, &H2A, &H52)
MFStreamExtension_CameraExtrinsics = iid
End Function
Public Function MFSampleExtension_CameraExtrinsics() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6B761658, &HB7EC, &H4C3B, &H82, &H25, &H86, &H23, &HCA, &HBE, &HC3, &H1D)
MFSampleExtension_CameraExtrinsics = iid
End Function
Public Function MFStreamExtension_PinholeCameraIntrinsics() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDBAC0455, &HEC8, &H4AEF, &H9C, &H32, &H7A, &H3E, &HE3, &H45, &H6F, &H53)
MFStreamExtension_PinholeCameraIntrinsics = iid
End Function
Public Function MFSampleExtension_PinholeCameraIntrinsics() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4EE3B6C5, &H6A15, &H4E72, &H97, &H61, &H70, &HC1, &HDB, &H8B, &H9F, &HE3)
MFSampleExtension_PinholeCameraIntrinsics = iid
End Function
Public Function MFMediaType_Default() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H81A412E6, &H8103, &H4B06, &H85, &H7F, &H18, &H62, &H78, &H10, &H24, &HAC)
MFMediaType_Default = iid
End Function
Public Function MFMediaType_Audio() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H73647561, &H0, &H10, &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
MFMediaType_Audio = iid
End Function
Public Function MFMediaType_Video() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H73646976, &H0, &H10, &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
MFMediaType_Video = iid
End Function
Public Function MFMediaType_Protected() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7B4B6FE6, &H9D04, &H4494, &HBE, &H14, &H7E, &HB, &HD0, &H76, &HC8, &HE4)
MFMediaType_Protected = iid
End Function
Public Function MFMediaType_SAMI() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE69669A0, &H3DCD, &H40CB, &H9E, &H2E, &H37, &H8, &H38, &H7C, &H6, &H16)
MFMediaType_SAMI = iid
End Function
Public Function MFMediaType_Script() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H72178C22, &HE45B, &H11D5, &HBC, &H2A, &H0, &HB0, &HD0, &HF3, &HF4, &HAB)
MFMediaType_Script = iid
End Function
Public Function MFMediaType_Image() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H72178C23, &HE45B, &H11D5, &HBC, &H2A, &H0, &HB0, &HD0, &HF3, &HF4, &HAB)
MFMediaType_Image = iid
End Function
Public Function MFMediaType_HTML() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H72178C24, &HE45B, &H11D5, &HBC, &H2A, &H0, &HB0, &HD0, &HF3, &HF4, &HAB)
MFMediaType_HTML = iid
End Function
Public Function MFMediaType_Binary() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H72178C25, &HE45B, &H11D5, &HBC, &H2A, &H0, &HB0, &HD0, &HF3, &HF4, &HAB)
MFMediaType_Binary = iid
End Function
Public Function MFMediaType_FileTransfer() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H72178C26, &HE45B, &H11D5, &HBC, &H2A, &H0, &HB0, &HD0, &HF3, &HF4, &HAB)
MFMediaType_FileTransfer = iid
End Function
Public Function MFMediaType_Stream() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE436EB83, &H524F, &H11CE, &H9F, &H53, &H0, &H20, &HAF, &HB, &HA7, &H70)
MFMediaType_Stream = iid
End Function
Public Function MFMediaType_MultiplexedFrames() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6EA542B0, &H281F, &H4231, &HA4, &H64, &HFE, &H2F, &H50, &H22, &H50, &H1C)
MFMediaType_MultiplexedFrames = iid
End Function
Public Function MFMediaType_Subtitle() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA6D13581, &HED50, &H4E65, &HAE, &H8, &H26, &H6, &H55, &H76, &HAA, &HCC)
MFMediaType_Subtitle = iid
End Function
Public Function MFMediaType_Perception() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H597FF6F9, &H6EA2, &H4670, &H85, &HB4, &HEA, &H84, &H7, &H3F, &HE9, &H40)
MFMediaType_Perception = iid
End Function
Public Function MFImageFormat_JPEG() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H19E4A5AA, &H5662, &H4FC5, &HA0, &HC0, &H17, &H58, &H2, &H8E, &H10, &H57)
MFImageFormat_JPEG = iid
End Function
Public Function MFImageFormat_RGB32() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H16, &H0, &H10, &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
MFImageFormat_RGB32 = iid
End Function
Public Function MFStreamFormat_MPEG2Transport() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE06D8023, &HDB46, &H11CF, &HB4, &HD1, &H0, &H80, &H5F, &H6C, &HBB, &HEA)
MFStreamFormat_MPEG2Transport = iid
End Function
Public Function MFStreamFormat_MPEG2Program() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H263067D1, &HD330, &H45DC, &HB6, &H69, &H34, &HD9, &H86, &HE4, &HE3, &HE1)
MFStreamFormat_MPEG2Program = iid
End Function
Public Function AM_MEDIA_TYPE_REPRESENTATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE2E42AD2, &H132C, &H491E, &HA2, &H68, &H3C, &H7C, &H2D, &HCA, &H18, &H1F)
AM_MEDIA_TYPE_REPRESENTATION = iid
End Function
Public Function FORMAT_MFVideoFormat() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAED4AB2D, &H7326, &H43CB, &H94, &H64, &HC8, &H79, &HCA, &HB9, &HC4, &H3D)
FORMAT_MFVideoFormat = iid
End Function
Public Function MFMediaType_Metadata() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2C8FA20C, &H82BB, &H4782, &H90, &HA0, &H98, &HA2, &HA5, &HBD, &H8E, &HF8)
MFMediaType_Metadata = iid
End Function
Public Function CLSID_MFSourceResolver() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H90EAB60F, &HE43A, &H4188, &HBC, &HC4, &HE4, &H7F, &HDF, &H4, &H86, &H8C)
CLSID_MFSourceResolver = iid
End Function
Public Function MF_DEVICESTREAM_ATTRIBUTE_FACEAUTH_CAPABILITY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCB6FD12A, &H2248, &H4E41, &HAD, &H46, &HE7, &H8B, &HB9, &HA, &HB9, &HFC)
MF_DEVICESTREAM_ATTRIBUTE_FACEAUTH_CAPABILITY = iid
End Function
Public Function MF_DEVICESTREAM_ATTRIBUTE_SECURE_CAPABILITY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H940FD626, &HEA6E, &H4684, &H98, &H40, &H36, &HBD, &H6E, &HC9, &HFB, &HEF)
MF_DEVICESTREAM_ATTRIBUTE_SECURE_CAPABILITY = iid
End Function



Public Function MFVideoFormat_Base() As UUID
'{00000000-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H0, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_Base = iid
End Function
Public Function MFVideoFormat_RGB32() As UUID
'{00000016-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H16, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_RGB32 = iid
End Function
Public Function MFVideoFormat_ARGB32() As UUID
'{00000015-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H15, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_ARGB32 = iid
End Function
Public Function MFVideoFormat_RGB24() As UUID
'{00000014-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H14, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_RGB24 = iid
End Function
Public Function MFVideoFormat_RGB555() As UUID
'{00000018-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H18, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_RGB555 = iid
End Function
Public Function MFVideoFormat_RGB565() As UUID
'{00000017-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H17, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_RGB565 = iid
End Function
Public Function MFVideoFormat_RGB8() As UUID
'{00000029-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H29, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_RGB8 = iid
End Function
Public Function MFVideoFormat_L8() As UUID
'{00000032-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H32, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_L8 = iid
End Function
Public Function MFVideoFormat_L16() As UUID
'{00000051-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H51, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_L16 = iid
End Function
Public Function MFVideoFormat_D16() As UUID
'{00000050-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H50, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_D16 = iid
End Function
Public Function MFVideoFormat_AI44() As UUID
'{34344941-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H34344941, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_AI44 = iid
End Function
Public Function MFVideoFormat_AYUV() As UUID
'{56555941-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H56555941, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_AYUV = iid
End Function
Public Function MFVideoFormat_YUY2() As UUID
'{32595559-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H32595559, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_YUY2 = iid
End Function
Public Function MFVideoFormat_YVYU() As UUID
'{55595659-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H55595659, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_YVYU = iid
End Function
Public Function MFVideoFormat_YVU9() As UUID
'{39555659-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H39555659, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_YVU9 = iid
End Function
Public Function MFVideoFormat_UYVY() As UUID
'{59565955-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H59565955, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_UYVY = iid
End Function
Public Function MFVideoFormat_NV11() As UUID
'{3131564E-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3131564E, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_NV11 = iid
End Function
Public Function MFVideoFormat_NV12() As UUID
'{3231564E-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3231564E, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_NV12 = iid
End Function
Public Function MFVideoFormat_YV12() As UUID
'{32315659-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H32315659, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_YV12 = iid
End Function
Public Function MFVideoFormat_I420() As UUID
'{30323449-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H30323449, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_I420 = iid
End Function
Public Function MFVideoFormat_IYUV() As UUID
'{56555949-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H56555949, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_IYUV = iid
End Function
Public Function MFVideoFormat_Y210() As UUID
'{30313259-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H30313259, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_Y210 = iid
End Function
Public Function MFVideoFormat_Y216() As UUID
'{36313259-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H36313259, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_Y216 = iid
End Function
Public Function MFVideoFormat_Y410() As UUID
'{30313459-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H30313459, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_Y410 = iid
End Function
Public Function MFVideoFormat_Y416() As UUID
'{36313459-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H36313459, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_Y416 = iid
End Function
Public Function MFVideoFormat_Y41P() As UUID
'{50313459-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H50313459, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_Y41P = iid
End Function
Public Function MFVideoFormat_Y41T() As UUID
'{54313459-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H54313459, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_Y41T = iid
End Function
Public Function MFVideoFormat_Y42T() As UUID
'{54323459-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H54323459, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_Y42T = iid
End Function
Public Function MFVideoFormat_P210() As UUID
'{30313250-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H30313250, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_P210 = iid
End Function
Public Function MFVideoFormat_P216() As UUID
'{36313250-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H36313250, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_P216 = iid
End Function
Public Function MFVideoFormat_P010() As UUID
'{30313050-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H30313050, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_P010 = iid
End Function
Public Function MFVideoFormat_P016() As UUID
'{36313050-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H36313050, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_P016 = iid
End Function
Public Function MFVideoFormat_v210() As UUID
'{30313276-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H30313276, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_v210 = iid
End Function
Public Function MFVideoFormat_v216() As UUID
'{36313276-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H36313276, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_v216 = iid
End Function
Public Function MFVideoFormat_v410() As UUID
'{30313476-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H30313476, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_v410 = iid
End Function
Public Function MFVideoFormat_MP43() As UUID
'{3334504D-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3334504D, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_MP43 = iid
End Function
Public Function MFVideoFormat_MP4S() As UUID
'{5334504D-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5334504D, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_MP4S = iid
End Function
Public Function MFVideoFormat_M4S2() As UUID
'{3253344D-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3253344D, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_M4S2 = iid
End Function
Public Function MFVideoFormat_MP4V() As UUID
'{5634504D-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5634504D, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_MP4V = iid
End Function
Public Function MFVideoFormat_WMV1() As UUID
'{31564D57-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H31564D57, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_WMV1 = iid
End Function
Public Function MFVideoFormat_WMV2() As UUID
'{32564D57-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H32564D57, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_WMV2 = iid
End Function
Public Function MFVideoFormat_WMV3() As UUID
'{33564D57-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H33564D57, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_WMV3 = iid
End Function
Public Function MFVideoFormat_WVC1() As UUID
'{31435657-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H31435657, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_WVC1 = iid
End Function
Public Function MFVideoFormat_MSS1() As UUID
'{3153534D-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3153534D, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_MSS1 = iid
End Function
Public Function MFVideoFormat_MSS2() As UUID
'{3253534D-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3253534D, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_MSS2 = iid
End Function
Public Function MFVideoFormat_MPG1() As UUID
'{3147504D-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3147504D, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_MPG1 = iid
End Function
Public Function MFVideoFormat_DVSL() As UUID
'{6C737664-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C737664, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_DVSL = iid
End Function
Public Function MFVideoFormat_DVSD() As UUID
'{64737664-0000-0010-8000-00AA00389B71}}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H64737664, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_DVSD = iid
End Function
Public Function MFVideoFormat_DVHD() As UUID
'{64687664-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H64687664, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_DVHD = iid
End Function
Public Function MFVideoFormat_DV25() As UUID
'{35327664-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H35327664, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_DV25 = iid
End Function
Public Function MFVideoFormat_DV50() As UUID
'{30357664-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H30357664, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_DV50 = iid
End Function
Public Function MFVideoFormat_DVH1() As UUID
'{31687664-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H31687664, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_DVH1 = iid
End Function
Public Function MFVideoFormat_DVC() As UUID
'{20637664-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H20637664, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_DVC = iid
End Function
Public Function MFVideoFormat_H264() As UUID
'{34363248-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H34363248, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_H264 = iid
End Function
Public Function MFVideoFormat_H265() As UUID
'{35363248-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H35363248, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_H265 = iid
End Function
Public Function MFVideoFormat_MJPG() As UUID
'{47504A4D-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H47504A4D, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_MJPG = iid
End Function
Public Function MFVideoFormat_420O() As UUID
'{4F303234-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4F303234, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_420O = iid
End Function
Public Function MFVideoFormat_HEVC() As UUID
'{43564548-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43564548, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_HEVC = iid
End Function
Public Function MFVideoFormat_HEVC_ES() As UUID
'{53564548-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H53564548, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_HEVC_ES = iid
End Function
Public Function MFVideoFormat_VP80() As UUID
'{30385056-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H30385056, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_VP80 = iid
End Function
Public Function MFVideoFormat_VP90() As UUID
'{30395056-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H30395056, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_VP90 = iid
End Function
Public Function MFVideoFormat_ORAW() As UUID
'{5741524F-0000-0010-8000-00AA00389B71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5741524F, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_ORAW = iid
End Function
Public Function MFVideoFormat_H263() As UUID
'{33363248-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H33363248, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_H263 = iid
End Function
Public Function MFVideoFormat_A2R10G10B10() As UUID
'{0000001f-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1F, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_A2R10G10B10 = iid
End Function
Public Function MFVideoFormat_A16B16G16R16F() As UUID
'{00000071-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H71, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_A16B16G16R16F = iid
End Function
Public Function MFVideoFormat_VP10() As UUID
'{30315056-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H30315056, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_VP10 = iid
End Function
Public Function MFVideoFormat_AV1() As UUID
'{31305641-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H31305641, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_AV1 = iid
End Function
Public Function MFVideoFormat_Theora() As UUID
'{6f656874-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6F656874, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFVideoFormat_Theora = iid
End Function

Public Function MFAudioFormat_Base() As UUID
'{00000000-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H0, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_Base = iid
End Function
Public Function MFAudioFormat_PCM() As UUID
'{00000001-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_PCM = iid
End Function
Public Function MFAudioFormat_Float() As UUID
'{00000003-0000-0010-8000-00aa00389b7}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H7)
 MFAudioFormat_Float = iid
End Function
Public Function MFAudioFormat_DTS() As UUID
'{00000008-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_DTS = iid
End Function
Public Function MFAudioFormat_Dolby_AC3_SPDIF() As UUID
'{00000092-0000-0010-8000-00aa00389b7}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H92, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H7)
 MFAudioFormat_Dolby_AC3_SPDIF = iid
End Function
Public Function MFAudioFormat_DRM() As UUID
'{00000009-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_DRM = iid
End Function
Public Function MFAudioFormat_WMAudioV8() As UUID
'{00000161-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H161, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_WMAudioV8 = iid
End Function
Public Function MFAudioFormat_WMAudioV9() As UUID
'{00000162-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H162, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_WMAudioV9 = iid
End Function
Public Function MFAudioFormat_WMAudio_Lossless() As UUID
'{00000163-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H163, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_WMAudio_Lossless = iid
End Function
Public Function MFAudioFormat_WMASPDIF() As UUID
'{00000164-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H164, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_WMASPDIF = iid
End Function
Public Function MFAudioFormat_MSP1() As UUID
'{0000000a-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_MSP1 = iid
End Function
Public Function MFAudioFormat_MP3() As UUID
'{00000055-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H55, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_MP3 = iid
End Function
Public Function MFAudioFormat_MPEG() As UUID
'{00000050-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H50, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_MPEG = iid
End Function
Public Function MFAudioFormat_AAC() As UUID
'{00001610-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1610, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_AAC = iid
End Function
Public Function MFAudioFormat_ADTS() As UUID
'{00001610-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1610, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_ADTS = iid
End Function
Public Function MFAudioFormat_AMR_NB() As UUID
'{00007361-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7361, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_AMR_NB = iid
End Function
Public Function MFAudioFormat_AMR_WB() As UUID
'{00007362-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7362, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_AMR_WB = iid
End Function
Public Function MFAudioFormat_AMR_WP() As UUID
'{00007363-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7363, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_AMR_WP = iid
End Function
Public Function MFAudioFormat_FLAC() As UUID
'{0000f1ac-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF1AC&, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_FLAC = iid
End Function
Public Function MFAudioFormat_ALAC() As UUID
'{00006c61-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C61, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_ALAC = iid
End Function
Public Function MFAudioFormat_Opus() As UUID
'{0000704f-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H704F, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_Opus = iid
End Function
Public Function MFAudioFormat_Dolby_AC4() As UUID
'{0000ac40-0000-0010-8000-00aa00389b71}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAC40&, CInt(&H0), CInt(&H10), &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
 MFAudioFormat_Dolby_AC4 = iid
End Function

Public Function CLSID_FaceDetectionMFT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC1E565E2, &HF2DE, &H4537, &H96, &H12, &H2F, &H30, &HA1, &H60, &HEB, &H5C)
CLSID_FaceDetectionMFT = iid
End Function
Public Function CLSID_FrameServerClassFactory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9A93092C, &H9CDC, &H49B8, &H83, &H49, &HCB, &HCF, &H31, &H45, &HFE, &HA)
CLSID_FrameServerClassFactory = iid
End Function
Public Function MF_CAMERASOURCE_PROVIDE_SELECTED_PROFILE_ON_START() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA9B46058, &H82F2, &H4E5C, &HBF, &H6E, &H25, &HB4, &HB0, &H9F, &H22, &HED)
MF_CAMERASOURCE_PROVIDE_SELECTED_PROFILE_ON_START = iid
End Function
Public Function MF_DEVSOURCE_ATTRIBUTE_FRAMESERVER_SHARE_MODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H44D1A9BC, &H2999, &H4238, &HAE, &H43, &H7, &H30, &HCE, &HB2, &HAB, &H1B)
MF_DEVSOURCE_ATTRIBUTE_FRAMESERVER_SHARE_MODE = iid
End Function
Public Function MFP_POSITIONTYPE_100NS() As UUID
    Static iid As UUID
    MFP_POSITIONTYPE_100NS = iid 'GUID_NULL
End Function
 
Public Function MF_PD_ASF_FILEPROPERTIES_FILE_ID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3DE649B4, &HD76D, &H4E66, &H9E, &HC9, &H78, &H12, &HF, &HB4, &HC7, &HE3)
MF_PD_ASF_FILEPROPERTIES_FILE_ID = iid
End Function
Public Function MF_PD_ASF_FILEPROPERTIES_CREATION_TIME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3DE649B6, &HD76D, &H4E66, &H9E, &HC9, &H78, &H12, &HF, &HB4, &HC7, &HE3)
MF_PD_ASF_FILEPROPERTIES_CREATION_TIME = iid
End Function
Public Function MF_PD_ASF_FILEPROPERTIES_PACKETS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3DE649B7, &HD76D, &H4E66, &H9E, &HC9, &H78, &H12, &HF, &HB4, &HC7, &HE3)
MF_PD_ASF_FILEPROPERTIES_PACKETS = iid
End Function
Public Function MF_PD_ASF_FILEPROPERTIES_PLAY_DURATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3DE649B8, &HD76D, &H4E66, &H9E, &HC9, &H78, &H12, &HF, &HB4, &HC7, &HE3)
MF_PD_ASF_FILEPROPERTIES_PLAY_DURATION = iid
End Function
Public Function MF_PD_ASF_FILEPROPERTIES_SEND_DURATION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3DE649B9, &HD76D, &H4E66, &H9E, &HC9, &H78, &H12, &HF, &HB4, &HC7, &HE3)
MF_PD_ASF_FILEPROPERTIES_SEND_DURATION = iid
End Function
Public Function MF_PD_ASF_FILEPROPERTIES_PREROLL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3DE649BA, &HD76D, &H4E66, &H9E, &HC9, &H78, &H12, &HF, &HB4, &HC7, &HE3)
MF_PD_ASF_FILEPROPERTIES_PREROLL = iid
End Function
Public Function MF_PD_ASF_FILEPROPERTIES_FLAGS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3DE649BB, &HD76D, &H4E66, &H9E, &HC9, &H78, &H12, &HF, &HB4, &HC7, &HE3)
MF_PD_ASF_FILEPROPERTIES_FLAGS = iid
End Function
Public Function MF_PD_ASF_FILEPROPERTIES_MIN_PACKET_SIZE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3DE649BC, &HD76D, &H4E66, &H9E, &HC9, &H78, &H12, &HF, &HB4, &HC7, &HE3)
MF_PD_ASF_FILEPROPERTIES_MIN_PACKET_SIZE = iid
End Function
Public Function MF_PD_ASF_FILEPROPERTIES_MAX_PACKET_SIZE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3DE649BD, &HD76D, &H4E66, &H9E, &HC9, &H78, &H12, &HF, &HB4, &HC7, &HE3)
MF_PD_ASF_FILEPROPERTIES_MAX_PACKET_SIZE = iid
End Function
Public Function MF_PD_ASF_FILEPROPERTIES_MAX_BITRATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3DE649BE, &HD76D, &H4E66, &H9E, &HC9, &H78, &H12, &HF, &HB4, &HC7, &HE3)
MF_PD_ASF_FILEPROPERTIES_MAX_BITRATE = iid
End Function
Public Function CLSID_WMDRMSystemID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8948BB22, &H11BD, &H4796, &H93, &HE3, &H97, &H4D, &H1B, &H57, &H56, &H78)
 CLSID_WMDRMSystemID = iid
End Function
Public Function MF_PD_ASF_CONTENTENCRYPTION_TYPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8520FE3D, &H277E, &H46EA, &H99, &HE4, &HE3, &HA, &H86, &HDB, &H12, &HBE)
MF_PD_ASF_CONTENTENCRYPTION_TYPE = iid
End Function
Public Function MF_PD_ASF_CONTENTENCRYPTION_KEYID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8520FE3E, &H277E, &H46EA, &H99, &HE4, &HE3, &HA, &H86, &HDB, &H12, &HBE)
MF_PD_ASF_CONTENTENCRYPTION_KEYID = iid
End Function
Public Function MF_PD_ASF_CONTENTENCRYPTION_SECRET_DATA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8520FE3F, &H277E, &H46EA, &H99, &HE4, &HE3, &HA, &H86, &HDB, &H12, &HBE)
MF_PD_ASF_CONTENTENCRYPTION_SECRET_DATA = iid
End Function
Public Function MF_PD_ASF_CONTENTENCRYPTION_LICENSE_URL() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8520FE40, &H277E, &H46EA, &H99, &HE4, &HE3, &HA, &H86, &HDB, &H12, &HBE)
MF_PD_ASF_CONTENTENCRYPTION_LICENSE_URL = iid
End Function
Public Function MF_PD_ASF_CONTENTENCRYPTIONEX_ENCRYPTION_DATA() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H62508BE5, &HECDF, &H4924, &HA3, &H59, &H72, &HBA, &HB3, &H39, &H7B, &H9D)
 MF_PD_ASF_CONTENTENCRYPTIONEX_ENCRYPTION_DATA = iid
End Function
Public Function MF_PD_ASF_LANGLIST() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF23DE43C, &H9977, &H460D, &HA6, &HEC, &H32, &H93, &H7F, &H16, &HF, &H7D)
 MF_PD_ASF_LANGLIST = iid
End Function
Public Function MF_PD_ASF_LANGLIST_LEGACYORDER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF23DE43D, &H9977, &H460D, &HA6, &HEC, &H32, &H93, &H7F, &H16, &HF, &H7D)
 MF_PD_ASF_LANGLIST_LEGACYORDER = iid
End Function
Public Function MF_PD_ASF_MARKER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5134330E, &H83A6, &H475E, &HA9, &HD5, &H4F, &HB8, &H75, &HFB, &H2E, &H31)
MF_PD_ASF_MARKER = iid
End Function
Public Function MF_PD_ASF_SCRIPT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE29CD0D7, &HD602, &H4923, &HA7, &HFE, &H73, &HFD, &H97, &HEC, &HC6, &H50)
 MF_PD_ASF_SCRIPT = iid
End Function
Public Function MF_PD_ASF_CODECLIST() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE4BB3509, &HC18D, &H4DF1, &HBB, &H99, &H7A, &H36, &HB3, &HCC, &H41, &H19)
MF_PD_ASF_CODECLIST = iid
End Function
Public Function MF_PD_ASF_METADATA_IS_VBR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5FC6947A, &HEF60, &H445D, &HB4, &H49, &H44, &H2E, &HCC, &H78, &HB4, &HC1)
 MF_PD_ASF_METADATA_IS_VBR = iid
End Function
Public Function MF_PD_ASF_METADATA_V8_VBRPEAK() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5FC6947B, &HEF60, &H445D, &HB4, &H49, &H44, &H2E, &HCC, &H78, &HB4, &HC1)
 MF_PD_ASF_METADATA_V8_VBRPEAK = iid
End Function
Public Function MF_PD_ASF_METADATA_V8_BUFFERAVERAGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5FC6947C, &HEF60, &H445D, &HB4, &H49, &H44, &H2E, &HCC, &H78, &HB4, &HC1)
 MF_PD_ASF_METADATA_V8_BUFFERAVERAGE = iid
End Function
Public Function MF_PD_ASF_METADATA_LEAKY_BUCKET_PAIRS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5FC6947D, &HEF60, &H445D, &HB4, &H49, &H44, &H2E, &HCC, &H78, &HB4, &HC1)
 MF_PD_ASF_METADATA_LEAKY_BUCKET_PAIRS = iid
End Function
Public Function MF_PD_ASF_DATA_START_OFFSET() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE7D5B3E7, &H1F29, &H45D3, &H88, &H22, &H3E, &H78, &HFA, &HE2, &H72, &HED)
MF_PD_ASF_DATA_START_OFFSET = iid
End Function
Public Function MF_PD_ASF_DATA_LENGTH() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE7D5B3E8, &H1F29, &H45D3, &H88, &H22, &H3E, &H78, &HFA, &HE2, &H72, &HED)
MF_PD_ASF_DATA_LENGTH = iid
End Function
Public Function MF_SD_ASF_EXTSTRMPROP_LANGUAGE_ID_INDEX() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H48F8A522, &H305D, &H422D, &H85, &H24, &H25, &H2, &HDD, &HA3, &H36, &H80)
MF_SD_ASF_EXTSTRMPROP_LANGUAGE_ID_INDEX = iid
End Function
Public Function MF_SD_ASF_EXTSTRMPROP_AVG_DATA_BITRATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H48F8A523, &H305D, &H422D, &H85, &H24, &H25, &H2, &HDD, &HA3, &H36, &H80)
MF_SD_ASF_EXTSTRMPROP_AVG_DATA_BITRATE = iid
End Function
Public Function MF_SD_ASF_EXTSTRMPROP_AVG_BUFFERSIZE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H48F8A524, &H305D, &H422D, &H85, &H24, &H25, &H2, &HDD, &HA3, &H36, &H80)
MF_SD_ASF_EXTSTRMPROP_AVG_BUFFERSIZE = iid
End Function
Public Function MF_SD_ASF_EXTSTRMPROP_MAX_DATA_BITRATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H48F8A525, &H305D, &H422D, &H85, &H24, &H25, &H2, &HDD, &HA3, &H36, &H80)
MF_SD_ASF_EXTSTRMPROP_MAX_DATA_BITRATE = iid
End Function
Public Function MF_SD_ASF_EXTSTRMPROP_MAX_BUFFERSIZE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H48F8A526, &H305D, &H422D, &H85, &H24, &H25, &H2, &HDD, &HA3, &H36, &H80)
MF_SD_ASF_EXTSTRMPROP_MAX_BUFFERSIZE = iid
End Function
Public Function MF_SD_ASF_STREAMBITRATES_BITRATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA8E182ED, &HAFC8, &H43D0, &HB0, &HD1, &HF6, &H5B, &HAD, &H9D, &HA5, &H58)
MF_SD_ASF_STREAMBITRATES_BITRATE = iid
End Function
Public Function MF_SD_ASF_METADATA_DEVICE_CONFORMANCE_TEMPLATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H245E929D, &HC44E, &H4F7E, &HBB, &H3C, &H77, &HD4, &HDF, &HD2, &H7F, &H8A)
MF_SD_ASF_METADATA_DEVICE_CONFORMANCE_TEMPLATE = iid
End Function
Public Function MF_PD_ASF_INFO_HAS_AUDIO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H80E62295, &H2296, &H4A44, &HB3, &H1C, &HD1, &H3, &HC6, &HFE, &HD2, &H3C)
MF_PD_ASF_INFO_HAS_AUDIO = iid
End Function
Public Function MF_PD_ASF_INFO_HAS_VIDEO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H80E62296, &H2296, &H4A44, &HB3, &H1C, &HD1, &H3, &HC6, &HFE, &HD2, &H3C)
MF_PD_ASF_INFO_HAS_VIDEO = iid
End Function
Public Function MF_PD_ASF_INFO_HAS_NON_AUDIO_VIDEO() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H80E62297, &H2296, &H4A44, &HB3, &H1C, &HD1, &H3, &HC6, &HFE, &HD2, &H3C)
MF_PD_ASF_INFO_HAS_NON_AUDIO_VIDEO = iid
End Function
Public Function MF_ASFPROFILE_MINPACKETSIZE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H22587626, &H47DE, &H4168, &H87, &HF5, &HB5, &HAA, &H9B, &H12, &HA8, &HF0)
MF_ASFPROFILE_MINPACKETSIZE = iid
End Function
Public Function MF_ASFPROFILE_MAXPACKETSIZE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H22587627, &H47DE, &H4168, &H87, &HF5, &HB5, &HAA, &H9B, &H12, &HA8, &HF0)
MF_ASFPROFILE_MAXPACKETSIZE = iid
End Function
Public Function MF_ASFSTREAMCONFIG_LEAKYBUCKET1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC69B5901, &HEA1A, &H4C9B, &HB6, &H92, &HE2, &HA0, &HD2, &H9A, &H8A, &HDD)
MF_ASFSTREAMCONFIG_LEAKYBUCKET1 = iid
End Function
Public Function MF_ASFSTREAMCONFIG_LEAKYBUCKET2() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC69B5902, &HEA1A, &H4C9B, &HB6, &H92, &HE2, &HA0, &HD2, &H9A, &H8A, &HDD)
MF_ASFSTREAMCONFIG_LEAKYBUCKET2 = iid
End Function
Public Function MFASFSampleExtension_SampleDuration() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC6BD9450, &H867F, &H4907, &H83, &HA3, &HC7, &H79, &H21, &HB7, &H33, &HAD)
MFASFSampleExtension_SampleDuration = iid
End Function
Public Function MFASFSampleExtension_OutputCleanPoint() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF72A3C6F, &H6EB4, &H4EBC, &HB1, &H92, &H9, &HAD, &H97, &H59, &HE8, &H28)
MFASFSampleExtension_OutputCleanPoint = iid
End Function
Public Function MFASFSampleExtension_SMPTE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H399595EC, &H8667, &H4E2D, &H8F, &HDB, &H98, &H81, &H4C, &HE7, &H6C, &H1E)
MFASFSampleExtension_SMPTE = iid
End Function
Public Function MFASFSampleExtension_FileName() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE165EC0E, &H19ED, &H45D7, &HB4, &HA7, &H25, &HCB, &HD1, &HE2, &H8E, &H9B)
MFASFSampleExtension_FileName = iid
End Function
Public Function MFASFSampleExtension_ContentType() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD590DC20, &H7BC, &H436C, &H9C, &HF7, &HF3, &HBB, &HFB, &HF1, &HA4, &HDC)
MFASFSampleExtension_ContentType = iid
End Function
Public Function MFASFSampleExtension_PixelAspectRatio() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B1EE554, &HF9EA, &H4BC8, &H82, &H1A, &H37, &H6B, &H74, &HE4, &HC4, &HB8)
MFASFSampleExtension_PixelAspectRatio = iid
End Function
Public Function MFASFSampleExtension_Encryption_SampleID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6698B84E, &HAFA, &H4330, &HAE, &HB2, &H1C, &HA, &H98, &HD7, &HA4, &H4D)
MFASFSampleExtension_Encryption_SampleID = iid
End Function
Public Function MFASFSampleExtension_Encryption_KeyID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H76376591, &H795F, &H4DA1, &H86, &HED, &H9D, &H46, &HEC, &HA1, &H9, &HA9)
MFASFSampleExtension_Encryption_KeyID = iid
End Function
Public Function MFASFMutexType_Language() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H72178C2B, &HE45B, &H11D5, &HBC, &H2A, &H0, &HB0, &HD0, &HF3, &HF4, &HAB)
MFASFMutexType_Language = iid
End Function
Public Function MFASFMutexType_Bitrate() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H72178C2C, &HE45B, &H11D5, &HBC, &H2A, &H0, &HB0, &HD0, &HF3, &HF4, &HAB)
MFASFMutexType_Bitrate = iid
End Function
Public Function MFASFMutexType_Presentation() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H72178C2D, &HE45B, &H11D5, &HBC, &H2A, &H0, &HB0, &HD0, &HF3, &HF4, &HAB)
MFASFMutexType_Presentation = iid
End Function
Public Function MFASFMutexType_Unknown() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H72178C2E, &HE45B, &H11D5, &HBC, &H2A, &H0, &HB0, &HD0, &HF3, &HF4, &HAB)
MFASFMutexType_Unknown = iid
End Function
Public Function MFASFINDEXER_TYPE_TIMECODE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H49815231, &H6BAD, &H44FD, &H81, &HA, &H3F, &H60, &H98, &H4E, &HC7, &HFD)
MFASFINDEXER_TYPE_TIMECODE = iid
End Function
Public Function MFASFSPLITTER_PACKET_BOUNDARY() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFE584A05, &HE8D6, &H42E3, &HB1, &H76, &HF1, &H21, &H17, &H5, &HFB, &H6F)
MFASFSPLITTER_PACKET_BOUNDARY = iid
End Function
Public Function MFPKEY_ASFMEDIASINK_BASE_SENDTIME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCDDCBC82, &H3411, &H4119, &H91, &H35, &H84, &H23, &HC4, &H1B, &H39, &H57, 3)
MFPKEY_ASFMEDIASINK_BASE_SENDTIME = pkk
End Function
Public Function MFPKEY_ASFMEDIASINK_AUTOADJUST_BITRATE() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCDDCBC82, &H3411, &H4119, &H91, &H35, &H84, &H23, &HC4, &H1B, &H39, &H57, 4)
MFPKEY_ASFMEDIASINK_AUTOADJUST_BITRATE = pkk
End Function
Public Function MFPKEY_ASFMEDIASINK_DRMACTION() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA1DB6F6C, &H1D0A, &H4CB6, &H82, &H54, &HCB, &H36, &HBE, &HED, &HBC, &H48, 5)
MFPKEY_ASFMEDIASINK_DRMACTION = pkk
End Function
Public Function MFPKEY_ASFSTREAMSINK_CORRECTED_LEAKYBUCKET() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA2F152FB, &H8AD9, &H4A11, &HB3, &H45, &H2C, &HE2, &HFA, &HD8, &H72, &H3D, 1)
MFPKEY_ASFSTREAMSINK_CORRECTED_LEAKYBUCKET = pkk
End Function
