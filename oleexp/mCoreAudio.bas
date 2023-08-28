Attribute VB_Name = "mCoreAudio"
Option Explicit
'Vars, PROPERTYKEY's and GUID's for Core Audio

'Update for oleexp 4.5: Additional PKEY_AudioEndpoint* values

'Upate for oleexp 6.2: Spatial audio IIDs.

' //,
' // Signatures for data structures.,
' //,
'Data structure signature: WARNING: CONVERT TO ANSI FIRST
Public Const APO_CONNECTION_DESCRIPTOR_SIGNATURE = "ACDS"
Public Const APO_CONNECTION_PROPERTY_SIGNATURE = "ACPS"
Public Const APO_CONNECTION_PROPERTY_V2_SIGNATURE = "ACP2"

' Min and max framerates for the engine,
Public Const AUDIO_MIN_FRAMERATE As Single = 10#      '// Minimum frame rate for APOs,
Public Const AUDIO_MAX_FRAMERATE As Single = 384000#  '// Maximum frame rate for APOs,

' Min and max # of channels (samples per frame) for the APOs,
Public Const AUDIO_MIN_CHANNELS = 1                       '// Current minimum number of channels for APOs,
Public Const AUDIO_MAX_CHANNELS = 4096                    '// Current maximum number of channels for APOs,
Public Function DEVINTERFACE_AUDIO_RENDER() As UUID
'{E6327CAD-DCEC-4949-AE8A-991E976A79D2}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE6327CAD, CInt(&HDCEC), CInt(&H4949), &HAE, &H8A, &H99, &H1E, &H97, &H6A, &H79, &HD2)
 DEVINTERFACE_AUDIO_RENDER = iid
End Function
Public Function DEVINTERFACE_AUDIO_CAPTURE() As UUID
'{2EEF81BE-33FA-4800-9670-1CD474972C3F}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2EEF81BE, CInt(&H33FA), CInt(&H4800), &H96, &H70, &H1C, &HD4, &H74, &H97, &H2C, &H3F)
 DEVINTERFACE_AUDIO_CAPTURE = iid
End Function
Public Function DEVINTERFACE_MIDI_OUTPUT() As UUID
'{6DC23320-AB33-4CE4-80D4-BBB3EBBF2814}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6DC23320, CInt(&HAB33), CInt(&H4CE4), &H80, &HD4, &HBB, &HB3, &HEB, &HBF, &H28, &H14)
 DEVINTERFACE_MIDI_OUTPUT = iid
End Function
Public Function DEVINTERFACE_MIDI_INPUT() As UUID
'{504BE32C-CCF6-4D2C-B73F-6F8B3747E22B}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H504BE32C, CInt(&HCCF6), CInt(&H4D2C), &HB7, &H3F, &H6F, &H8B, &H37, &H47, &HE2, &H2B)
 DEVINTERFACE_MIDI_INPUT = iid
End Function
Public Function EVENTCONTEXT_VOLUMESLIDER() As UUID
'{E2C2E9DE-09B1-4B04-84E5-07931225EE04}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE2C2E9DE, CInt(&H9B1), CInt(&H4B04), &H84, &HE5, &H7, &H93, &H12, &H25, &HEE, &H4)
 EVENTCONTEXT_VOLUMESLIDER = iid
End Function

Public Function IID_IMMNotificationClient() As UUID
'{7991EEC9-7E89-4D85-8390-6C703CEC60C0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7991EEC9, CInt(&H7E89), CInt(&H4D85), &H83, &H90, &H6C, &H70, &H3C, &HEC, &H60, &HC0)
IID_IMMNotificationClient = iid
End Function
Public Function IID_IMMDevice() As UUID
'{D666063F-1587-4E43-81F1-B948E807363F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD666063F, CInt(&H1587), CInt(&H4E43), &H81, &HF1, &HB9, &H48, &HE8, &H7, &H36, &H3F)
IID_IMMDevice = iid
End Function
Public Function IID_IMMDeviceCollection() As UUID
'{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBD7A1BE, CInt(&H7A1A), CInt(&H44DB), &H83, &H97, &HCC, &H53, &H92, &H38, &H7B, &H5E)
IID_IMMDeviceCollection = iid
End Function
Public Function IID_IMMEndpoint() As UUID
'{1BE09788-6894-4089-8586-9A2A6C265AC5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1BE09788, CInt(&H6894), CInt(&H4089), &H85, &H86, &H9A, &H2A, &H6C, &H26, &H5A, &HC5)
IID_IMMEndpoint = iid
End Function
Public Function IID_IMMDeviceEnumerator() As UUID
'{A95664D2-9614-4F35-A746-DE8DB63617E6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA95664D2, CInt(&H9614), CInt(&H4F35), &HA7, &H46, &HDE, &H8D, &HB6, &H36, &H17, &HE6)
IID_IMMDeviceEnumerator = iid
End Function
Public Function IID_IMMDeviceActivator() As UUID
'{3B0D0EA4-D0A9-4B0E-935B-09516746FAC0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3B0D0EA4, CInt(&HD0A9), CInt(&H4B0E), &H93, &H5B, &H9, &H51, &H67, &H46, &HFA, &HC0)
IID_IMMDeviceActivator = iid
End Function
Public Function IID_IActivateAudioInterfaceCompletionHandler() As UUID
'{41D949AB-9862-444A-80F6-C261334DA5EB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H41D949AB, CInt(&H9862), CInt(&H444A), &H80, &HF6, &HC2, &H61, &H33, &H4D, &HA5, &HEB)
IID_IActivateAudioInterfaceCompletionHandler = iid
End Function
Public Function IID_IActivateAudioInterfaceAsyncOperation() As UUID
'{72A22D78-CDE4-431D-B8CC-843A71199B6D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H72A22D78, CInt(&HCDE4), CInt(&H431D), &HB8, &HCC, &H84, &H3A, &H71, &H19, &H9B, &H6D)
IID_IActivateAudioInterfaceAsyncOperation = iid
End Function
Public Function IID_IAudioEndpointVolumeCallback() As UUID
'{657804FA-D6AD-4496-8A60-352752AF4F89}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H657804FA, CInt(&HD6AD), CInt(&H4496), &H8A, &H60, &H35, &H27, &H52, &HAF, &H4F, &H89)
IID_IAudioEndpointVolumeCallback = iid
End Function
Public Function IID_IAudioEndpointVolume() As UUID
'{5CDF2C82-841E-4546-9722-0CF74078229A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5CDF2C82, CInt(&H841E), CInt(&H4546), &H97, &H22, &HC, &HF7, &H40, &H78, &H22, &H9A)
IID_IAudioEndpointVolume = iid
End Function
Public Function IID_IAudioEndpointVolumeEx() As UUID
'{66E11784-F695-4F28-A505-A7080081A78F}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H66E11784, CInt(&HF695), CInt(&H4F28), &HA5, &H5, &HA7, &H8, &H0, &H81, &HA7, &H8F)
 IID_IAudioEndpointVolumeEx = iid
End Function
Public Function IID_IAudioMeterInformation() As UUID
'{C02216F6-8C67-4B5B-9D00-D008E73E0064}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC02216F6, CInt(&H8C67), CInt(&H4B5B), &H9D, &H0, &HD0, &H8, &HE7, &H3E, &H0, &H64)
IID_IAudioMeterInformation = iid
End Function
Public Function IID_IAudioEndpointFormatControl() As UUID
'{784CFD40-9F89-456E-A1A6-873B006A664E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H784CFD40, CInt(&H9F89), CInt(&H456E), &HA1, &HA6, &H87, &H3B, &H0, &H6A, &H66, &H4E)
IID_IAudioEndpointFormatControl = iid
End Function
Public Function IID_IKsControl() As UUID
'{28F54685-06FD-11D2-B27A-00A0C9223196}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H28F54685, CInt(&H6FD), CInt(&H11D2), &HB2, &H7A, &H0, &HA0, &HC9, &H22, &H31, &H96)
IID_IKsControl = iid
End Function
Public Function IID_IAudioVolumeLevel() As UUID
'{7FB7B48F-531D-44A2-BCB3-5AD5A134B3DC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7FB7B48F, CInt(&H531D), CInt(&H44A2), &HBC, &HB3, &H5A, &HD5, &HA1, &H34, &HB3, &HDC)
IID_IAudioVolumeLevel = iid
End Function
Public Function IID_IAudioChannelConfig() As UUID
'{BB11C46F-EC28-493C-B88A-5DB88062CE98}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBB11C46F, CInt(&HEC28), CInt(&H493C), &HB8, &H8A, &H5D, &HB8, &H80, &H62, &HCE, &H98)
IID_IAudioChannelConfig = iid
End Function
Public Function IID_IAudioLoudness() As UUID
'{7D8B1437-DD53-4350-9C1B-1EE2890BD938}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7D8B1437, CInt(&HDD53), CInt(&H4350), &H9C, &H1B, &H1E, &HE2, &H89, &HB, &HD9, &H38)
IID_IAudioLoudness = iid
End Function
Public Function IID_IAudioInputSelector() As UUID
'{4F03DC02-5E6E-4653-8F72-A030C123D598}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4F03DC02, CInt(&H5E6E), CInt(&H4653), &H8F, &H72, &HA0, &H30, &HC1, &H23, &HD5, &H98)
IID_IAudioInputSelector = iid
End Function
Public Function IID_IAudioOutputSelector() As UUID
'{BB515F69-94A7-429e-8B9C-271B3F11A3AB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBB515F69, CInt(&H94A7), CInt(&H429E), &H8B, &H9C, &H27, &H1B, &H3F, &H11, &HA3, &HAB)
IID_IAudioOutputSelector = iid
End Function
Public Function IID_IAudioMute() As UUID
'{DF45AEEA-B74A-4B6B-AFAD-2366B6AA012E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDF45AEEA, CInt(&HB74A), CInt(&H4B6B), &HAF, &HAD, &H23, &H66, &HB6, &HAA, &H1, &H2E)
IID_IAudioMute = iid
End Function
Public Function IID_IPerChannelDbLevel() As UUID
'{C2F8E001-F205-4BC9-99BC-C13B1E048CCB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC2F8E001, CInt(&HF205), CInt(&H4BC9), &H99, &HBC, &HC1, &H3B, &H1E, &H4, &H8C, &HCB)
IID_IPerChannelDbLevel = iid
End Function
Public Function IID_IAudioBass() As UUID
'{A2B1A1D9-4DB3-425D-A2B2-BD335CB3E2E5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA2B1A1D9, CInt(&H4DB3), CInt(&H425D), &HA2, &HB2, &HBD, &H33, &H5C, &HB3, &HE2, &HE5)
IID_IAudioBass = iid
End Function
Public Function IID_IAudioMidrange() As UUID
'{5E54B6D7-B44B-40D9-9A9E-E691D9CE6EDF}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5E54B6D7, CInt(&HB44B), CInt(&H40D9), &H9A, &H9E, &HE6, &H91, &HD9, &HCE, &H6E, &HDF)
IID_IAudioMidrange = iid
End Function
Public Function IID_IAudioTreble() As UUID
'{0A717812-694E-4907-B74B-BAFA5CFDCA7B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA717812, CInt(&H694E), CInt(&H4907), &HB7, &H4B, &HBA, &HFA, &H5C, &HFD, &HCA, &H7B)
IID_IAudioTreble = iid
End Function
Public Function IID_IAudioAutoGainControl() As UUID
'{85401FD4-6DE4-4b9d-9869-2D6753A82F3C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H85401FD4, CInt(&H6DE4), CInt(&H4B9D), &H98, &H69, &H2D, &H67, &H53, &HA8, &H2F, &H3C)
IID_IAudioAutoGainControl = iid
End Function
Public Function IID_IAudioPeakMeter() As UUID
'{DD79923C-0599-45e0-B8B6-C8DF7DB6E796}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDD79923C, CInt(&H599), CInt(&H45E0), &HB8, &HB6, &HC8, &HDF, &H7D, &HB6, &HE7, &H96)
IID_IAudioPeakMeter = iid
End Function
Public Function IID_IDeviceSpecificProperty() As UUID
'{3B22BCBF-2586-4af0-8583-205D391B807C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3B22BCBF, CInt(&H2586), CInt(&H4AF0), &H85, &H83, &H20, &H5D, &H39, &H1B, &H80, &H7C)
IID_IDeviceSpecificProperty = iid
End Function
Public Function IID_IKsFormatSupport() As UUID
'{3CB4A69D-BB6F-4D2B-95B7-452D2C155DB5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3CB4A69D, CInt(&HBB6F), CInt(&H4D2B), &H95, &HB7, &H45, &H2D, &H2C, &H15, &H5D, &HB5)
IID_IKsFormatSupport = iid
End Function
Public Function IID_IKsJackDescription() As UUID
'{4509F757-2D46-4637-8E62-CE7DB944F57B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4509F757, CInt(&H2D46), CInt(&H4637), &H8E, &H62, &HCE, &H7D, &HB9, &H44, &HF5, &H7B)
IID_IKsJackDescription = iid
End Function
Public Function IID_IKsJackDescription2() As UUID
'{478F3A9B-E0C9-4827-9228-6F5505FFE76A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H478F3A9B, CInt(&HE0C9), CInt(&H4827), &H92, &H28, &H6F, &H55, &H5, &HFF, &HE7, &H6A)
IID_IKsJackDescription2 = iid
End Function
Public Function IID_IKsJackSinkInformation() As UUID
'{D9BD72ED-290F-4581-9FF3-61027A8FE532}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD9BD72ED, CInt(&H290F), CInt(&H4581), &H9F, &HF3, &H61, &H2, &H7A, &H8F, &HE5, &H32)
IID_IKsJackSinkInformation = iid
End Function
Public Function IID_IKsJackContainerId() As UUID
'{C99AF463-D629-4EC4-8C00-E54D68154248}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC99AF463, CInt(&HD629), CInt(&H4EC4), &H8C, &H0, &HE5, &H4D, &H68, &H15, &H42, &H48)
IID_IKsJackContainerId = iid
End Function
Public Function IID_IPartsList() As UUID
'{6DAA848C-5EB0-45CC-AEA5-998A2CDA1FFB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6DAA848C, CInt(&H5EB0), CInt(&H45CC), &HAE, &HA5, &H99, &H8A, &H2C, &HDA, &H1F, &HFB)
IID_IPartsList = iid
End Function
Public Function IID_IPart() As UUID
'{AE2DE0E4-5BCA-4F2D-AA46-5D13F8FDB3A9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAE2DE0E4, CInt(&H5BCA), CInt(&H4F2D), &HAA, &H46, &H5D, &H13, &HF8, &HFD, &HB3, &HA9)
IID_IPart = iid
End Function
Public Function IID_IConnector() As UUID
'{9c2c4058-23f5-41de-877a-df3af236a09e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9C2C4058, CInt(&H23F5), CInt(&H41DE), &H87, &H7A, &HDF, &H3A, &HF2, &H36, &HA0, &H9E)
IID_IConnector = iid
End Function
Public Function IID_ISubunit() As UUID
'{82149A85-DBA6-4487-86BB-EA8F7FEFCC71}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H82149A85, CInt(&HDBA6), CInt(&H4487), &H86, &HBB, &HEA, &H8F, &H7F, &HEF, &HCC, &H71)
IID_ISubunit = iid
End Function
Public Function IID_IControlInterface() As UUID
'{45d37c3f-5140-444a-ae24-400789f3cbf3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H45D37C3F, CInt(&H5140), CInt(&H444A), &HAE, &H24, &H40, &H7, &H89, &HF3, &HCB, &HF3)
IID_IControlInterface = iid
End Function
Public Function IID_IControlChangeNotify() As UUID
'{A09513ED-C709-4d21-BD7B-5F34C47F3947}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA09513ED, CInt(&HC709), CInt(&H4D21), &HBD, &H7B, &H5F, &H34, &HC4, &H7F, &H39, &H47)
IID_IControlChangeNotify = iid
End Function
Public Function IID_IDeviceTopology() As UUID
'{2A07407E-6497-4A18-9787-32F79BD0D98F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2A07407E, CInt(&H6497), CInt(&H4A18), &H97, &H87, &H32, &HF7, &H9B, &HD0, &HD9, &H8F)
IID_IDeviceTopology = iid
End Function
Public Function IID_IAudioClient() As UUID
'{1CB9AD4C-DBFA-4c32-B178-C2F568A703B2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1CB9AD4C, CInt(&HDBFA), CInt(&H4C32), &HB1, &H78, &HC2, &HF5, &H68, &HA7, &H3, &HB2)
IID_IAudioClient = iid
End Function
Public Function IID_IAudioClient2() As UUID
'{726778CD-F60A-4eda-82DE-E47610CD78AA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H726778CD, CInt(&HF60A), CInt(&H4EDA), &H82, &HDE, &HE4, &H76, &H10, &HCD, &H78, &HAA)
IID_IAudioClient2 = iid
End Function
Public Function IID_IAudioClient3() As UUID
'{7ED4EE07-8E67-4CD4-8C1A-2B7A5987AD42}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7ED4EE07, CInt(&H8E67), CInt(&H4CD4), &H8C, &H1A, &H2B, &H7A, &H59, &H87, &HAD, &H42)
IID_IAudioClient3 = iid
End Function
Public Function IID_IAudioRenderClient() As UUID
'{F294ACFC-3146-4483-A7BF-ADDCA7C260E2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF294ACFC, CInt(&H3146), CInt(&H4483), &HA7, &HBF, &HAD, &HDC, &HA7, &HC2, &H60, &HE2)
IID_IAudioRenderClient = iid
End Function
Public Function IID_IAudioCaptureClient() As UUID
'{C8ADBD64-E71E-48a0-A4DE-185C395CD317}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC8ADBD64, CInt(&HE71E), CInt(&H48A0), &HA4, &HDE, &H18, &H5C, &H39, &H5C, &HD3, &H17)
IID_IAudioCaptureClient = iid
End Function
Public Function IID_IAudioClock() As UUID
'{CD63314F-3FBA-4a1b-812C-EF96358728E7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCD63314F, CInt(&H3FBA), CInt(&H4A1B), &H81, &H2C, &HEF, &H96, &H35, &H87, &H28, &HE7)
IID_IAudioClock = iid
End Function
Public Function IID_IAudioClock2() As UUID
'{6f49ff73-6727-49ac-a008-d98cf5e70048}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6F49FF73, CInt(&H6727), CInt(&H49AC), &HA0, &H8, &HD9, &H8C, &HF5, &HE7, &H0, &H48)
IID_IAudioClock2 = iid
End Function
Public Function IID_IAudioClockAdjustment() As UUID
'{f6e4c0a0-46d9-4fb8-be21-57a3ef2b626c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF6E4C0A0, CInt(&H46D9), CInt(&H4FB8), &HBE, &H21, &H57, &HA3, &HEF, &H2B, &H62, &H6C)
IID_IAudioClockAdjustment = iid
End Function
Public Function IID_ISimpleAudioVolume() As UUID
'{87CE5498-68D6-44E5-9215-6DA47EF883D8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H87CE5498, CInt(&H68D6), CInt(&H44E5), &H92, &H15, &H6D, &HA4, &H7E, &HF8, &H83, &HD8)
IID_ISimpleAudioVolume = iid
End Function
Public Function IID_IAudioStreamVolume() As UUID
'{93014887-242D-4068-8A15-CF5E93B90FE3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H93014887, CInt(&H242D), CInt(&H4068), &H8A, &H15, &HCF, &H5E, &H93, &HB9, &HF, &HE3)
IID_IAudioStreamVolume = iid
End Function
Public Function IID_IChannelAudioVolume() As UUID
'{1C158861-B533-4B30-B1CF-E853E51C59B8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1C158861, CInt(&HB533), CInt(&H4B30), &HB1, &HCF, &HE8, &H53, &HE5, &H1C, &H59, &HB8)
IID_IChannelAudioVolume = iid
End Function
Public Function IID_IAudioSessionEvents() As UUID
'{24918ACC-64B3-37C1-8CA9-74A66E9957A8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H24918ACC, CInt(&H64B3), CInt(&H37C1), &H8C, &HA9, &H74, &HA6, &H6E, &H99, &H57, &HA8)
IID_IAudioSessionEvents = iid
End Function
Public Function IID_IAudioSessionControl() As UUID
'{F4B1A599-7266-4319-A8CA-E70ACB11E8CD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF4B1A599, CInt(&H7266), CInt(&H4319), &HA8, &HCA, &HE7, &HA, &HCB, &H11, &HE8, &HCD)
IID_IAudioSessionControl = iid
End Function
Public Function IID_IAudioSessionControl2() As UUID
'{bfb7ff88-7239-4fc9-8fa2-07c950be9c6d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBFB7FF88, CInt(&H7239), CInt(&H4FC9), &H8F, &HA2, &H7, &HC9, &H50, &HBE, &H9C, &H6D)
IID_IAudioSessionControl2 = iid
End Function
Public Function IID_IAudioSessionManager() As UUID
'{BFA971F1-4D5E-40BB-935E-967039BFBEE4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBFA971F1, CInt(&H4D5E), CInt(&H40BB), &H93, &H5E, &H96, &H70, &H39, &HBF, &HBE, &HE4)
IID_IAudioSessionManager = iid
End Function
Public Function IID_IAudioVolumeDuckNotification() As UUID
'{C3B284D4-6D39-4359-B3CF-B56DDB3BB39C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC3B284D4, CInt(&H6D39), CInt(&H4359), &HB3, &HCF, &HB5, &H6D, &HDB, &H3B, &HB3, &H9C)
IID_IAudioVolumeDuckNotification = iid
End Function
Public Function IID_IAudioSessionNotification() As UUID
'{641DD20B-4D41-49CC-ABA3-174B9477BB08}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H641DD20B, CInt(&H4D41), CInt(&H49CC), &HAB, &HA3, &H17, &H4B, &H94, &H77, &HBB, &H8)
IID_IAudioSessionNotification = iid
End Function
Public Function IID_IAudioSessionEnumerator() As UUID
'{E2F5BB11-0570-40CA-ACDD-3AA01277DEE8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE2F5BB11, CInt(&H570), CInt(&H40CA), &HAC, &HDD, &H3A, &HA0, &H12, &H77, &HDE, &HE8)
IID_IAudioSessionEnumerator = iid
End Function
Public Function IID_IAudioSessionManager2() As UUID
'{77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H77AA99A0, CInt(&H1BD6), CInt(&H484F), &H8B, &HC7, &H2C, &H65, &H4C, &H9A, &H9B, &H6F)
IID_IAudioSessionManager2 = iid
End Function
Public Function IID_IAudioLfxControl() As UUID
'{076A6922-D802-4F83-BAF6-409D9CA11BFE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H76A6922, CInt(&HD802), CInt(&H4F83), &HBA, &HF6, &H40, &H9D, &H9C, &HA1, &H1B, &HFE)
IID_IAudioLfxControl = iid
End Function
Public Function IID_IAudioEndpointLastBufferControl() As UUID
'{F8520DD3-8F9D-4437-9861-62F584C33DD6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF8520DD3, CInt(&H8F9D), CInt(&H4437), &H98, &H61, &H62, &HF5, &H84, &HC3, &H3D, &HD6)
IID_IAudioEndpointLastBufferControl = iid
End Function
Public Function IID_IAudioSystemEffects() As UUID
'{5FA00F27-ADD6-499a-8A9D-6B98521FA75B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5FA00F27, CInt(&HADD6), CInt(&H499A), &H8A, &H9D, &H6B, &H98, &H52, &H1F, &HA7, &H5B)
IID_IAudioSystemEffects = iid
End Function
Public Function IID_IAudioSystemEffects2() As UUID
'{BAFE99D2-7436-44CE-9E0E-4D89AFBFFF56}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBAFE99D2, CInt(&H7436), CInt(&H44CE), &H9E, &HE, &H4D, &H89, &HAF, &HBF, &HFF, &H56)
IID_IAudioSystemEffects2 = iid
End Function
Public Function IID_IAudioEndpointOffloadStreamVolume() As UUID
'{64F1DD49-71CA-4281-8672-3A9EDDD1D0B6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H64F1DD49, CInt(&H71CA), CInt(&H4281), &H86, &H72, &H3A, &H9E, &HDD, &HD1, &HD0, &HB6)
IID_IAudioEndpointOffloadStreamVolume = iid
End Function
Public Function IID_IAudioEndpointOffloadStreamMute() As UUID
'{DFE21355-5EC2-40E0-8D6B-710AC3C00249}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDFE21355, CInt(&H5EC2), CInt(&H40E0), &H8D, &H6B, &H71, &HA, &HC3, &HC0, &H2, &H49)
IID_IAudioEndpointOffloadStreamMute = iid
End Function
Public Function IID_IAudioEndpointOffloadStreamMeter() As UUID
'{E1546DCE-9DD1-418B-9AB2-348CED161C86}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE1546DCE, CInt(&H9DD1), CInt(&H418B), &H9A, &HB2, &H34, &H8C, &HED, &H16, &H1C, &H86)
IID_IAudioEndpointOffloadStreamMeter = iid
End Function
Public Function IID_IHardwareAudioEngineBase() As UUID
'{EDDCE3E4-F3C1-453a-B461-223563CBD886}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEDDCE3E4, CInt(&HF3C1), CInt(&H453A), &HB4, &H61, &H22, &H35, &H63, &HCB, &HD8, &H86)
IID_IHardwareAudioEngineBase = iid
End Function
Public Function IID_ISpatialAudioMetadataWriter() As UUID
'{1B17CA01-2955-444D-A430-537DC589A844}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1B17CA01, CInt(&H2955), CInt(&H444D), &HA4, &H30, &H53, &H7D, &HC5, &H89, &HA8, &H44)
IID_ISpatialAudioMetadataWriter = iid
End Function
Public Function IID_ISpatialAudioMetadataReader() As UUID
'{B78E86A2-31D9-4C32-94D2-7DF40FC7EBEC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB78E86A2, CInt(&H31D9), CInt(&H4C32), &H94, &HD2, &H7D, &HF4, &HF, &HC7, &HEB, &HEC)
IID_ISpatialAudioMetadataReader = iid
End Function
Public Function IID_ISpatialAudioMetadataCopier() As UUID
'{D224B233-E251-4FD0-9CA2-D5ECF9A68404}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD224B233, CInt(&HE251), CInt(&H4FD0), &H9C, &HA2, &HD5, &HEC, &HF9, &HA6, &H84, &H4)
IID_ISpatialAudioMetadataCopier = iid
End Function
Public Function IID_ISpatialAudioMetadataItemsBuffer() As UUID
'{42640A16-E1BD-42D9-9FF6-031AB71A2DBA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H42640A16, CInt(&HE1BD), CInt(&H42D9), &H9F, &HF6, &H3, &H1A, &HB7, &H1A, &H2D, &HBA)
IID_ISpatialAudioMetadataItemsBuffer = iid
End Function
Public Function IID_ISpatialAudioMetadataClient() As UUID
'{777D4A3B-F6FF-4A26-85DC-68D7CDEDA1D4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H777D4A3B, CInt(&HF6FF), CInt(&H4A26), &H85, &HDC, &H68, &HD7, &HCD, &HED, &HA1, &HD4)
IID_ISpatialAudioMetadataClient = iid
End Function
Public Function IID_ISpatialAudioObjectForMetadataCommands() As UUID
'{0DF2C94B-F5F9-472D-AF6B-C46E0AC9CD05}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDF2C94B, CInt(&HF5F9), CInt(&H472D), &HAF, &H6B, &HC4, &H6E, &HA, &HC9, &HCD, &H5)
IID_ISpatialAudioObjectForMetadataCommands = iid
End Function
Public Function IID_ISpatialAudioObjectForMetadataItems() As UUID
'{DDEA49FF-3BC0-4377-8AAD-9FBCFD808566}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDDEA49FF, CInt(&H3BC0), CInt(&H4377), &H8A, &HAD, &H9F, &HBC, &HFD, &H80, &H85, &H66)
IID_ISpatialAudioObjectForMetadataItems = iid
End Function
Public Function IID_ISpatialAudioObjectRenderStreamForMetadata() As UUID
'{BBC9C907-48D5-4A2E-A0C7-F7F0D67C1FB1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBBC9C907, CInt(&H48D5), CInt(&H4A2E), &HA0, &HC7, &HF7, &HF0, &HD6, &H7C, &H1F, &HB1)
IID_ISpatialAudioObjectRenderStreamForMetadata = iid
End Function
Public Function IID_IAudioFormatEnumerator() As UUID
'{DCDAA858-895A-4A22-A5EB-67BDA506096D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDCDAA858, CInt(&H895A), CInt(&H4A22), &HA5, &HEB, &H67, &HBD, &HA5, &H6, &H9, &H6D)
IID_IAudioFormatEnumerator = iid
End Function
Public Function IID_ISpatialAudioObjectBase() As UUID
'{CCE0B8F2-8D4D-4EFB-A8CF-3D6ECF1C30E0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCCE0B8F2, CInt(&H8D4D), CInt(&H4EFB), &HA8, &HCF, &H3D, &H6E, &HCF, &H1C, &H30, &HE0)
IID_ISpatialAudioObjectBase = iid
End Function
Public Function IID_ISpatialAudioObject() As UUID
'{DDE28967-521B-46E5-8F00-BD6F2BC8AB1D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDDE28967, CInt(&H521B), CInt(&H46E5), &H8F, &H0, &HBD, &H6F, &H2B, &HC8, &HAB, &H1D)
IID_ISpatialAudioObject = iid
End Function
Public Function IID_ISpatialAudioObjectRenderStream() As UUID
'{BAB5F473-B423-477B-85F5-B5A332A04153}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBAB5F473, CInt(&HB423), CInt(&H477B), &H85, &HF5, &HB5, &HA3, &H32, &HA0, &H41, &H53)
IID_ISpatialAudioObjectRenderStream = iid
End Function
Public Function IID_ISpatialAudioObjectRenderStreamNotify() As UUID
'{DDDF83E6-68D7-4C70-883F-A1836AFB4A50}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDDDF83E6, CInt(&H68D7), CInt(&H4C70), &H88, &H3F, &HA1, &H83, &H6A, &HFB, &H4A, &H50)
IID_ISpatialAudioObjectRenderStreamNotify = iid
End Function
Public Function IID_ISpatialAudioClient() As UUID
'{BBF8E066-AAAA-49BE-9A4D-FD2A858EA27F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBBF8E066, CInt(&HAAAA), CInt(&H49BE), &H9A, &H4D, &HFD, &H2A, &H85, &H8E, &HA2, &H7F)
IID_ISpatialAudioClient = iid
End Function
Public Function IID_ISpatialAudioClient2() As UUID
'{caabe452-a66a-4bee-a93e-e320463f6a53}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCAABE452, CInt(&HA66A), CInt(&H4BEE), &HA9, &H3E, &HE3, &H20, &H46, &H3F, &H6A, &H53)
IID_ISpatialAudioClient2 = iid
End Function
Public Function IID_ISpatialAudioObjectForHrtf() As UUID
'{D7436ADE-1978-4E14-ABA0-555BD8EB83B4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD7436ADE, CInt(&H1978), CInt(&H4E14), &HAB, &HA0, &H55, &H5B, &HD8, &HEB, &H83, &HB4)
IID_ISpatialAudioObjectForHrtf = iid
End Function
Public Function IID_ISpatialAudioObjectRenderStreamForHrtf() As UUID
'{E08DEEF9-5363-406E-9FDC-080EE247BBE0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE08DEEF9, CInt(&H5363), CInt(&H406E), &H9F, &HDC, &H8, &HE, &HE2, &H47, &HBB, &HE0)
IID_ISpatialAudioObjectRenderStreamForHrtf = iid
End Function
Public Function IID_ISpatialAudioMetadataItems() As UUID
'{BCD7C78F-3098-4F22-B547-A2F25A381269}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBCD7C78F, CInt(&H3098), CInt(&H4F22), &HB5, &H47, &HA2, &HF2, &H5A, &H38, &H12, &H69)
IID_ISpatialAudioMetadataItems = iid
End Function
Public Function IID_IAcousticEchoCancellationControl() As UUID
'{f4ae25b5-aaa3-437d-b6b3-dbbe2d0e9549}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF4AE25B5, CInt(&HAAA3), CInt(&H437D), &HB6, &HB3, &HDB, &HBE, &H2D, &HE, &H95, &H49)
IID_IAcousticEchoCancellationControl = iid
End Function
Public Function IID_IAudioClientDuckingControl() As UUID
'{C789D381-A28C-4168-B28F-D3A837924DC3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC789D381, CInt(&HA28C), CInt(&H4168), &HB2, &H8F, &HD3, &HA8, &H37, &H92, &H4D, &HC3)
IID_IAudioClientDuckingControl = iid
End Function
Public Function IID_IAudioViewManagerService() As UUID
'{A7A7EF10-1F49-45E0-AD35-612057CC8F74}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA7A7EF10, CInt(&H1F49), CInt(&H45E0), &HAD, &H35, &H61, &H20, &H57, &HCC, &H8F, &H74)
IID_IAudioViewManagerService = iid
End Function
Public Function IID_IAudioEffectsChangedNotificationClient() As UUID
'{A5DED44F-3C5D-4B2B-BD1E-5DC1EE20BBF6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA5DED44F, CInt(&H3C5D), CInt(&H4B2B), &HBD, &H1E, &H5D, &HC1, &HEE, &H20, &HBB, &HF6)
IID_IAudioEffectsChangedNotificationClient = iid
End Function
Public Function IID_IAudioEffectsManager() As UUID
'{4460B3AE-4B44-4527-8676-7548A8ACD260}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4460B3AE, CInt(&H4B44), CInt(&H4527), &H86, &H76, &H75, &H48, &HA8, &HAC, &HD2, &H60)
IID_IAudioEffectsManager = iid
End Function
Public Function IID_IAudioMediaType() As UUID
'{4E997F73-B71F-4798-873B-ED7DFCF15B4D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4E997F73, CInt(&HB71F), CInt(&H4798), &H87, &H3B, &HED, &H7D, &HFC, &HF1, &H5B, &H4D)
IID_IAudioMediaType = iid
End Function
Public Function IID_IAudioProcessingObjectRT() As UUID
'{9E1D6A6D-DDBC-4E95-A4C7-AD64BA37846C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9E1D6A6D, CInt(&HDDBC), CInt(&H4E95), &HA4, &HC7, &HAD, &H64, &HBA, &H37, &H84, &H6C)
IID_IAudioProcessingObjectRT = iid
End Function
Public Function IID_IAudioProcessingObjectVBR() As UUID
'{7ba1db8f-78ad-49cd-9591-f79d80a17c81}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7BA1DB8F, CInt(&H78AD), CInt(&H49CD), &H95, &H91, &HF7, &H9D, &H80, &HA1, &H7C, &H81)
IID_IAudioProcessingObjectVBR = iid
End Function
Public Function IID_IAudioProcessingObjectConfiguration() As UUID
'{0E5ED805-ABA6-49c3-8F9A-2B8C889C4FA8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE5ED805, CInt(&HABA6), CInt(&H49C3), &H8F, &H9A, &H2B, &H8C, &H88, &H9C, &H4F, &HA8)
IID_IAudioProcessingObjectConfiguration = iid
End Function
Public Function IID_IAudioProcessingObject() As UUID
'{FD7F2B29-24D0-4b5c-B177-592C39F9CA10}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFD7F2B29, CInt(&H24D0), CInt(&H4B5C), &HB1, &H77, &H59, &H2C, &H39, &HF9, &HCA, &H10)
IID_IAudioProcessingObject = iid
End Function
Public Function IID_IAudioDeviceModulesClient() As UUID
'{98F37DAC-D0B6-49F5-896A-AA4D169A4C48}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H98F37DAC, CInt(&HD0B6), CInt(&H49F5), &H89, &H6A, &HAA, &H4D, &H16, &H9A, &H4C, &H48)
IID_IAudioDeviceModulesClient = iid
End Function
Public Function IID_IAudioSystemEffectsCustomFormats() As UUID
'{B1176E34-BB7F-4f05-BEBD-1B18A534E097}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB1176E34, CInt(&HBB7F), CInt(&H4F05), &HBE, &HBD, &H1B, &H18, &HA5, &H34, &HE0, &H97)
IID_IAudioSystemEffectsCustomFormats = iid
End Function
Public Function IID_IApoAuxiliaryInputConfiguration() As UUID
'{4CEB0AAB-FA19-48ED-A857-87771AE1B768}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4CEB0AAB, CInt(&HFA19), CInt(&H48ED), &HA8, &H57, &H87, &H77, &H1A, &HE1, &HB7, &H68)
IID_IApoAuxiliaryInputConfiguration = iid
End Function
Public Function IID_IApoAuxiliaryInputRT() As UUID
'{F851809C-C177-49A0-B1B2-B66F017943AB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF851809C, CInt(&HC177), CInt(&H49A0), &HB1, &HB2, &HB6, &H6F, &H1, &H79, &H43, &HAB)
IID_IApoAuxiliaryInputRT = iid
End Function
Public Function IID_IApoAcousticEchoCancellation() As UUID
'{25385759-3236-4101-A943-25693DFB5D2D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H25385759, CInt(&H3236), CInt(&H4101), &HA9, &H43, &H25, &H69, &H3D, &HFB, &H5D, &H2D)
IID_IApoAcousticEchoCancellation = iid
End Function
Public Function IID_IAudioSystemEffectsPropertyChangeNotificationClient() As UUID
'{20049D40-56D5-400E-A2EF-385599FEED49}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20049D40, CInt(&H56D5), CInt(&H400E), &HA2, &HEF, &H38, &H55, &H99, &HFE, &HED, &H49)
IID_IAudioSystemEffectsPropertyChangeNotificationClient = iid
End Function
Public Function IID_IAudioSystemEffectsPropertyStore() As UUID
'{302AE7F9-D7E0-43E4-971B-1F8293613D2A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H302AE7F9, CInt(&HD7E0), CInt(&H43E4), &H97, &H1B, &H1F, &H82, &H93, &H61, &H3D, &H2A)
IID_IAudioSystemEffectsPropertyStore = iid
End Function



Public Function PKEY_AudioEndpoint_FormFactor() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE, 0)
PKEY_AudioEndpoint_FormFactor = pkk
End Function
Public Function PKEY_AudioEndpoint_ControlPanelPageProvider() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE, 1)
PKEY_AudioEndpoint_ControlPanelPageProvider = pkk
End Function
Public Function PKEY_AudioEndpoint_Association() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE, 2)
PKEY_AudioEndpoint_Association = pkk
End Function
Public Function PKEY_AudioEndpoint_PhysicalSpeakers() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE, 3)
PKEY_AudioEndpoint_PhysicalSpeakers = pkk
End Function
Public Function PKEY_AudioEndpoint_GUID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE, 4)
PKEY_AudioEndpoint_GUID = pkk
End Function
Public Function PKEY_AudioEndpoint_Disable_SysFx() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE, 5)
PKEY_AudioEndpoint_Disable_SysFx = pkk
End Function
Public Function PKEY_AudioEndpoint_FullRangeSpeakers() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE, 6)
PKEY_AudioEndpoint_FullRangeSpeakers = pkk
End Function
Public Function PKEY_AudioEndpoint_Supports_EventDriven_Mode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE, 7)
PKEY_AudioEndpoint_Supports_EventDriven_Mode = pkk
End Function
Public Function PKEY_AudioEndpoint_JackSubType() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE, 8)
PKEY_AudioEndpoint_JackSubType = pkk
End Function
Public Function PKEY_AudioEndpoint_Default_VolumeInDb() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE, 9)
PKEY_AudioEndpoint_Default_VolumeInDb = pkk
End Function
Public Function PKEY_AudioEngine_DeviceFormat() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF19F064D, &H82C, &H4E27, &HBC, &H73, &H68, &H82, &HA1, &HBB, &H8E, &H4C, 0)
PKEY_AudioEngine_DeviceFormat = pkk
End Function
Public Function PKEY_AudioEngine_OEMFormat() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE4870E26, &H3CC5, &H4CD2, &HBA, &H46, &HCA, &HA, &H9A, &H70, &HED, &H4, 3)
PKEY_AudioEngine_OEMFormat = pkk
End Function
Public Function PKEY_AudioEndpointLogo_IconEffects() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF1AB780D, &H2010, &H4ED3, &HA3, &HA6, &H8B, &H87, &HF0, &HF0, &HC4, &H76, 0)
PKEY_AudioEndpointLogo_IconEffects = pkk
End Function
Public Function PKEY_AudioEndpointLogo_IconPath() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF1AB780D, &H2010, &H4ED3, &HA3, &HA6, &H8B, &H87, &HF0, &HF0, &HC4, &H76, 1)
PKEY_AudioEndpointLogo_IconPath = pkk
End Function
Public Function PKEY_AudioEndpointSettings_MenuText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14242002, &H320, &H4DE4, &H95, &H55, &HA7, &HD8, &H2B, &H73, &HC2, &H86, 0)
PKEY_AudioEndpointSettings_MenuText = pkk
End Function
Public Function PKEY_AudioEndpointSettings_LaunchContract() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14242002, &H320, &H4DE4, &H95, &H55, &HA7, &HD8, &H2B, &H73, &HC2, &H86, 1)
PKEY_AudioEndpointSettings_LaunchContract = pkk
End Function

'[ Description ("Not a mistake.") ]
Public Function PKEY_FX_EffectPack_Schema_V1() As UUID
'{7abf23d9-727e-4d0b-86a3-dd501d260001}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7ABF23D9, CInt(&H727E), CInt(&H4D0B), &H86, &HA3, &HDD, &H50, &H1D, &H26, &H0, &H1)
 PKEY_FX_EffectPack_Schema_V1 = iid
End Function

Public Function PKEY_FX_Association() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 0)
PKEY_FX_Association = pkk
End Function
Public Function PKEY_FX_PreMixEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 1)
PKEY_FX_PreMixEffectClsid = pkk
End Function
Public Function PKEY_FX_PostMixEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 2)
PKEY_FX_PostMixEffectClsid = pkk
End Function
Public Function PKEY_FX_UserInterfaceClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 3)
PKEY_FX_UserInterfaceClsid = pkk
End Function
Public Function PKEY_FX_FriendlyName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 4)
PKEY_FX_FriendlyName = pkk
End Function
Public Function PKEY_FX_StreamEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 5)
PKEY_FX_StreamEffectClsid = pkk
End Function
Public Function PKEY_FX_ModeEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 6)
PKEY_FX_ModeEffectClsid = pkk
End Function
Public Function PKEY_FX_EndpointEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 7)
PKEY_FX_EndpointEffectClsid = pkk
End Function
Public Function PKEY_FX_KeywordDetector_StreamEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 8)
PKEY_FX_KeywordDetector_StreamEffectClsid = pkk
End Function
Public Function PKEY_FX_KeywordDetector_ModeEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 9)
PKEY_FX_KeywordDetector_ModeEffectClsid = pkk
End Function
Public Function PKEY_FX_KeywordDetector_EndpointEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 10)
PKEY_FX_KeywordDetector_EndpointEffectClsid = pkk
End Function
Public Function PKEY_FX_Offload_StreamEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 11)
PKEY_FX_Offload_StreamEffectClsid = pkk
End Function
Public Function PKEY_FX_Offload_ModeEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 12)
PKEY_FX_Offload_ModeEffectClsid = pkk
End Function
Public Function PKEY_CompositeFX_StreamEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 13)
PKEY_CompositeFX_StreamEffectClsid = pkk
End Function
Public Function PKEY_CompositeFX_ModeEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 14)
PKEY_CompositeFX_ModeEffectClsid = pkk
End Function
Public Function PKEY_CompositeFX_EndpointEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 15)
PKEY_CompositeFX_EndpointEffectClsid = pkk
End Function
Public Function PKEY_CompositeFX_KeywordDetector_StreamEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 16)
PKEY_CompositeFX_KeywordDetector_StreamEffectClsid = pkk
End Function
Public Function PKEY_CompositeFX_KeywordDetector_ModeEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 17)
PKEY_CompositeFX_KeywordDetector_ModeEffectClsid = pkk
End Function
Public Function PKEY_CompositeFX_KeywordDetector_EndpointEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 18)
PKEY_CompositeFX_KeywordDetector_EndpointEffectClsid = pkk
End Function
Public Function PKEY_CompositeFX_Offload_StreamEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 19)
PKEY_CompositeFX_Offload_StreamEffectClsid = pkk
End Function
Public Function PKEY_CompositeFX_Offload_ModeEffectClsid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 20)
PKEY_CompositeFX_Offload_ModeEffectClsid = pkk
End Function
Public Function PKEY_FX_SupportAppLauncher() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 21)
PKEY_FX_SupportAppLauncher = pkk
End Function
Public Function PKEY_FX_SupportedFormats() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 22)
PKEY_FX_SupportedFormats = pkk
End Function
Public Function PKEY_FX_Enumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 23)
PKEY_FX_Enumerator = pkk
End Function
Public Function PKEY_FX_VersionMajor() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 24)
PKEY_FX_VersionMajor = pkk
End Function
Public Function PKEY_FX_VersionMinor() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 25)
PKEY_FX_VersionMinor = pkk
End Function
Public Function PKEY_FX_Author() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 26)
PKEY_FX_Author = pkk
End Function
Public Function PKEY_FX_ObjectId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 27)
PKEY_FX_ObjectId = pkk
End Function
Public Function PKEY_FX_State() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 28)
PKEY_FX_State = pkk
End Function
Public Function PKEY_FX_EffectPackSchema_Version() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 29)
PKEY_FX_EffectPackSchema_Version = pkk
End Function
Public Function PKEY_FX_ApplyToBluetooth() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 30)
PKEY_FX_ApplyToBluetooth = pkk
End Function
Public Function PKEY_FX_ApplyToUsb() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 31)
PKEY_FX_ApplyToUsb = pkk
End Function
Public Function PKEY_FX_ApplyToRender() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 32)
PKEY_FX_ApplyToRender = pkk
End Function
Public Function PKEY_FX_ApplyToCapture() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 33)
PKEY_FX_ApplyToCapture = pkk
End Function
Public Function PKEY_FX_RequestSetAsDefault() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 34)
PKEY_FX_RequestSetAsDefault = pkk
End Function
Public Function PKEY_FX_RequestSetAsDefaultPriority() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD04E05A6, &H594B, &H4FB6, &HA8, &HD, &H1, &HAF, &H5E, &HED, &H7D, &H1D, 35)
PKEY_FX_RequestSetAsDefaultPriority = pkk
End Function
Public Function PKEY_SFX_ProcessingModes_Supported_For_Streaming() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD3993A3F, &H99C2, &H4402, &HB5, &HEC, &HA9, &H2A, &H3, &H67, &H66, &H4B, 5)
PKEY_SFX_ProcessingModes_Supported_For_Streaming = pkk
End Function
Public Function PKEY_MFX_ProcessingModes_Supported_For_Streaming() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD3993A3F, &H99C2, &H4402, &HB5, &HEC, &HA9, &H2A, &H3, &H67, &H66, &H4B, 6)
PKEY_MFX_ProcessingModes_Supported_For_Streaming = pkk
End Function
Public Function PKEY_EFX_ProcessingModes_Supported_For_Streaming() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD3993A3F, &H99C2, &H4402, &HB5, &HEC, &HA9, &H2A, &H3, &H67, &H66, &H4B, 7)
PKEY_EFX_ProcessingModes_Supported_For_Streaming = pkk
End Function
Public Function PKEY_SFX_KeywordDetector_ProcessingModes_Supported_For_Streaming() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD3993A3F, &H99C2, &H4402, &HB5, &HEC, &HA9, &H2A, &H3, &H67, &H66, &H4B, 8)
PKEY_SFX_KeywordDetector_ProcessingModes_Supported_For_Streaming = pkk
End Function
Public Function PKEY_MFX_KeywordDetector_ProcessingModes_Supported_For_Streaming() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD3993A3F, &H99C2, &H4402, &HB5, &HEC, &HA9, &H2A, &H3, &H67, &H66, &H4B, 9)
PKEY_MFX_KeywordDetector_ProcessingModes_Supported_For_Streaming = pkk
End Function
Public Function PKEY_EFX_KeywordDetector_ProcessingModes_Supported_For_Streaming() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD3993A3F, &H99C2, &H4402, &HB5, &HEC, &HA9, &H2A, &H3, &H67, &H66, &H4B, 10)
PKEY_EFX_KeywordDetector_ProcessingModes_Supported_For_Streaming = pkk
End Function
Public Function PKEY_SFX_Offload_ProcessingModes_Supported_For_Streaming() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD3993A3F, &H99C2, &H4402, &HB5, &HEC, &HA9, &H2A, &H3, &H67, &H66, &H4B, 11)
PKEY_SFX_Offload_ProcessingModes_Supported_For_Streaming = pkk
End Function
Public Function PKEY_MFX_Offload_ProcessingModes_Supported_For_Streaming() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD3993A3F, &H99C2, &H4402, &HB5, &HEC, &HA9, &H2A, &H3, &H67, &H66, &H4B, 12)
PKEY_MFX_Offload_ProcessingModes_Supported_For_Streaming = pkk
End Function
Public Function PKEY_APO_SWFallback_ProcessingModes() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD3993A3F, &H99C2, &H4402, &HB5, &HEC, &HA9, &H2A, &H3, &H67, &H66, &H4B, 13)
PKEY_APO_SWFallback_ProcessingModes = pkk
End Function


Public Function GetCAStatusStr(lStatus As DEVICE_STATE) As String
Select Case lStatus
    Case DEVICE_STATE_ACTIVE: GetCAStatusStr = "Active"
    Case DEVICE_STATE_DISABLED: GetCAStatusStr = "Disabled"
    Case DEVICE_STATE_NOTPRESENT: GetCAStatusStr = "Not present"
    Case DEVICE_STATE_UNPLUGGED: GetCAStatusStr = "Unplugged"
    Case Else: GetCAStatusStr = "<invalid>"
End Select
End Function
Public Function GetCARoleStr(rl As ERole) As String
Select Case rl
    Case eCommunications: GetCARoleStr = "Communications"
    Case eConsole: GetCARoleStr = "Console"
    Case eMultimedia: GetCARoleStr = "Multimedia"
End Select
End Function
Public Function GetCAFlowStr(fl As EDataFlow) As String
Select Case fl
    Case eCapture: GetCAFlowStr = "Capture"
    Case eRender: GetCAFlowStr = "Render"
    Case eAll: GetCAFlowStr = "All"
End Select
End Function
