Attribute VB_Name = "mDirectShow"
Option Explicit

'----------------------------------------------------------------------
'mDirectShow.bas - Part of oleexp
'
'This module contains IIDs for working with DirectShow COM interfaces.
'
'----------------------------------------------------------------------

Public Const MPBOOL_TRUE As Single = 1!
Public Const MPBOOL_FALSE As Single = 0!

Public Const DWORD_ALLPARAMS = -1
Public Function IID_IMediaParamInfo() As UUID
'{6d6cbb60-a223-44aa-842f-a2f06750be6d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6D6CBB60, CInt(&HA223), CInt(&H44AA), &H84, &H2F, &HA2, &HF0, &H67, &H50, &HBE, &H6D)
IID_IMediaParamInfo = iid
End Function
Public Function IID_IMediaParams() As UUID
'{6d6cbb61-a223-44aa-842f-a2f06750be6e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6D6CBB61, CInt(&HA223), CInt(&H44AA), &H84, &H2F, &HA2, &HF0, &H67, &H50, &HBE, &H6E)
IID_IMediaParams = iid
End Function
Public Function GUID_TIME_REFERENCE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H93AD712B, &HDAA0, &H4FFE, &HBC, &H81, &HB0, &HCE, &H50, &HF, &HCD, &HD9)
GUID_TIME_REFERENCE = iid
End Function
Public Function GUID_TIME_MUSIC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H574C49D, &H5B04, &H4B15, &HA5, &H42, &HAE, &H28, &H20, &H30, &H11, &H7B)
GUID_TIME_MUSIC = iid
End Function
Public Function GUID_TIME_SAMPLES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA8593D05, &HC43, &H4984, &H9A, &H63, &H97, &HAF, &H9E, &H2, &HC4, &HC0)
GUID_TIME_SAMPLES = iid
End Function

Public Function LIBID_QuartzNetTypeLib() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B1, &HAD4, &H11CE, &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
LIBID_QuartzNetTypeLib = iid
End Function

Public Function CLSID_MP3DecoderDMO() As UUID
'{BBEEA841-0A63-4F52-A7AB-A9B3A84ED38A}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBBEEA841, CInt(&HA63), CInt(&H4F52), &HA7, &HAB, &HA9, &HB3, &HA8, &H4E, &HD3, &H8A)
 CLSID_MP3DecoderDMO = iid
End Function

Public Function CLSID_CaptureGraphBuilder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBF87B6E0, CInt(&H8C27), CInt(&H11D0), &HB3, &HF0, &H0, &HAA, &H0, &H37, &H61, &HC5)
 CLSID_CaptureGraphBuilder = iid
End Function

Public Function CLSID_CaptureGraphBuilder2() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBF87B6E1, CInt(&H8C27), CInt(&H11D0), &HB3, &HF0, &H0, &HAA, &H0, &H37, &H61, &HC5)
 CLSID_CaptureGraphBuilder2 = iid
End Function

Public Function CLSID_ProtoFilterGraph() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE436EBB0, CInt(&H524F), CInt(&H11CE), &H9F, &H53, &H0, &H20, &HAF, &HB, &HA7, &H70)
 CLSID_ProtoFilterGraph = iid
End Function

Public Function CLSID_SystemClock() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE436EBB1, CInt(&H524F), CInt(&H11CE), &H9F, &H53, &H0, &H20, &HAF, &HB, &HA7, &H70)
 CLSID_SystemClock = iid
End Function

Public Function CLSID_FilterMapper() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE436EBB2, CInt(&H524F), CInt(&H11CE), &H9F, &H53, &H0, &H20, &HAF, &HB, &HA7, &H70)
 CLSID_FilterMapper = iid
End Function

Public Function CLSID_FilterGraph() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE436EBB3, CInt(&H524F), CInt(&H11CE), &H9F, &H53, &H0, &H20, &HAF, &HB, &HA7, &H70)
 CLSID_FilterGraph = iid
End Function

Public Function CLSID_FilterGraphNoThread() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE436EBB8, CInt(&H524F), CInt(&H11CE), &H9F, &H53, &H0, &H20, &HAF, &HB, &HA7, &H70)
 CLSID_FilterGraphNoThread = iid
End Function

Public Function CLSID_FilterGraphPrivateThread() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA3ECBC41, CInt(&H581A), CInt(&H4476), &HB6, &H93, &HA6, &H33, &H40, &H46, &H2D, &H8B)
 CLSID_FilterGraphPrivateThread = iid
End Function

Public Function CLSID_MPEG1Doc() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE4BBD160, CInt(&H4269), CInt(&H11CE), &H83, &H8D, &H0, &HAA, &H0, &H55, &H59, &H5A)
 CLSID_MPEG1Doc = iid
End Function

Public Function CLSID_FileSource() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H701722E0, CInt(&H8AE3), CInt(&H11CE), &HA8, &H5C, &H0, &HAA, &H0, &H2F, &HEA, &HB5)
 CLSID_FileSource = iid
End Function

Public Function CLSID_MPEG1PacketPlayer() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H26C25940, CInt(&H4CA9), CInt(&H11CE), &HA8, &H28, &H0, &HAA, &H0, &H2F, &HEA, &HB5)
 CLSID_MPEG1PacketPlayer = iid
End Function

Public Function CLSID_MPEG1Splitter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H336475D0, CInt(&H942A), CInt(&H11CE), &HA8, &H70, &H0, &HAA, &H0, &H2F, &HEA, &HB5)
 CLSID_MPEG1Splitter = iid
End Function

Public Function CLSID_CMpegVideoCodec() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFEB50740, CInt(&H7BEF), CInt(&H11CE), &H9B, &HD9, &H0, &H0, &HE2, &H2, &H59, &H9C)
 CLSID_CMpegVideoCodec = iid
End Function

Public Function CLSID_CMpegAudioCodec() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4A2286E0, CInt(&H7BEF), CInt(&H11CE), &H9B, &HD9, &H0, &H0, &HE2, &H2, &H59, &H9C)
 CLSID_CMpegAudioCodec = iid
End Function

Public Function CLSID_TextRender() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE30629D3, CInt(&H27E5), CInt(&H11CE), &H87, &H5D, &H0, &H60, &H8C, &HB7, &H80, &H66)
 CLSID_TextRender = iid
End Function

Public Function CLSID_InfTee() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF8388A40, CInt(&HD5BB), CInt(&H11D0), &HBE, &H5A, &H0, &H80, &HC7, &H6, &H56, &H8E)
 CLSID_InfTee = iid
End Function

Public Function CLSID_AviSplitter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B544C20, CInt(&HFD0B), CInt(&H11CE), &H8C, &H63, &H0, &HAA, &H0, &H44, &HB5, &H1E)
 CLSID_AviSplitter = iid
End Function

Public Function CLSID_AviReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B544C21, CInt(&HFD0B), CInt(&H11CE), &H8C, &H63, &H0, &HAA, &H0, &H44, &HB5, &H1E)
 CLSID_AviReader = iid
End Function

Public Function CLSID_VfwCapture() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B544C22, CInt(&HFD0B), CInt(&H11CE), &H8C, &H63, &H0, &HAA, &H0, &H44, &HB5, &H1E)
 CLSID_VfwCapture = iid
End Function

Public Function CLSID_CaptureProperties() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B544C22, CInt(&HFD0B), CInt(&H11CE), &H8C, &H63, &H0, &HAA, &H0, &H44, &HB5, &H1F)
 CLSID_CaptureProperties = iid
End Function

Public Function CLSID_FGControl() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE436EBB4, CInt(&H524F), CInt(&H11CE), &H9F, &H53, &H0, &H20, &HAF, &HB, &HA7, &H70)
 CLSID_FGControl = iid
End Function

Public Function CLSID_MOVReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H44584800, CInt(&HF8EE), CInt(&H11CE), &HB2, &HD4, &H0, &HDD, &H1, &H10, &H1B, &H85)
 CLSID_MOVReader = iid
End Function

Public Function CLSID_QuickTimeParser() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD51BD5A0, CInt(&H7548), CInt(&H11CF), &HA5, &H20, &H0, &H80, &HC7, &H7E, &HF5, &H8A)
 CLSID_QuickTimeParser = iid
End Function

Public Function CLSID_QTDec() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFDFE9681, CInt(&H74A3), CInt(&H11D0), &HAF, &HA7, &H0, &HAA, &H0, &HB6, &H7A, &H42)
 CLSID_QTDec = iid
End Function

Public Function CLSID_AVIDoc() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD3588AB0, CInt(&H781), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
 CLSID_AVIDoc = iid
End Function

Public Function CLSID_VideoRenderer() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H70E102B0, CInt(&H5556), CInt(&H11CE), &H97, &HC0, &H0, &HAA, &H0, &H55, &H59, &H5A)
 CLSID_VideoRenderer = iid
End Function

Public Function CLSID_Colour() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1643E180, CInt(&H90F5), CInt(&H11CE), &H97, &HD5, &H0, &HAA, &H0, &H55, &H59, &H5A)
 CLSID_Colour = iid
End Function

Public Function CLSID_Dither() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1DA08500, CInt(&H9EDC), CInt(&H11CF), &HBC, &H10, &H0, &HAA, &H0, &HAC, &H74, &HF6)
 CLSID_Dither = iid
End Function

Public Function CLSID_ModexRenderer() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7167665, CInt(&H5011), CInt(&H11CF), &HBF, &H33, &H0, &HAA, &H0, &H55, &H59, &H5A)
 CLSID_ModexRenderer = iid
End Function

Public Function CLSID_AudioRender() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE30629D1, CInt(&H27E5), CInt(&H11CE), &H87, &H5D, &H0, &H60, &H8C, &HB7, &H80, &H66)
 CLSID_AudioRender = iid
End Function

Public Function CLSID_AudioProperties() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5589FAF, CInt(&HC356), CInt(&H11CE), &HBF, &H1, &H0, &HAA, &H0, &H55, &H59, &H5A)
 CLSID_AudioProperties = iid
End Function

Public Function CLSID_DSoundRender() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H79376820, CInt(&H7D0), CInt(&H11CF), &HA2, &H4D, &H0, &H20, &HAF, &HD7, &H97, &H67)
 CLSID_DSoundRender = iid
End Function

Public Function CLSID_AudioRecord() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE30629D2, CInt(&H27E5), CInt(&H11CE), &H87, &H5D, &H0, &H60, &H8C, &HB7, &H80, &H66)
 CLSID_AudioRecord = iid
End Function

Public Function CLSID_AudioInputMixerProperties() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2CA8CA52, CInt(&H3C3F), CInt(&H11D2), &HB7, &H3D, &H0, &HC0, &H4F, &HB6, &HBD, &H3D)
 CLSID_AudioInputMixerProperties = iid
End Function

Public Function CLSID_AVIDec() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCF49D4E0, CInt(&H1115), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
 CLSID_AVIDec = iid
End Function

Public Function CLSID_AVIDraw() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA888DF60, CInt(&H1E90), CInt(&H11CF), &HAC, &H98, &H0, &HAA, &H0, &H4C, &HF, &HA9)
 CLSID_AVIDraw = iid
End Function

Public Function CLSID_ACMWrapper() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6A08CF80, CInt(&HE18), CInt(&H11CF), &HA2, &H4D, &H0, &H20, &HAF, &HD7, &H97, &H67)
 CLSID_ACMWrapper = iid
End Function

Public Function CLSID_AsyncReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE436EBB5, CInt(&H524F), CInt(&H11CE), &H9F, &H53, &H0, &H20, &HAF, &HB, &HA7, &H70)
 CLSID_AsyncReader = iid
End Function

Public Function CLSID_URLReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE436EBB6, CInt(&H524F), CInt(&H11CE), &H9F, &H53, &H0, &H20, &HAF, &HB, &HA7, &H70)
 CLSID_URLReader = iid
End Function

Public Function CLSID_PersistMonikerPID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE436EBB7, CInt(&H524F), CInt(&H11CE), &H9F, &H53, &H0, &H20, &HAF, &HB, &HA7, &H70)
 CLSID_PersistMonikerPID = iid
End Function

Public Function CLSID_AVICo() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD76E2820, CInt(&H1563), CInt(&H11CF), &HAC, &H98, &H0, &HAA, &H0, &H4C, &HF, &HA9)
 CLSID_AVICo = iid
End Function

Public Function CLSID_FileWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8596E5F0, CInt(&HDA5), CInt(&H11D0), &HBD, &H21, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_FileWriter = iid
End Function

Public Function CLSID_AviDest() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE2510970, CInt(&HF137), CInt(&H11CE), &H8B, &H67, &H0, &HAA, &H0, &HA3, &HF1, &HA6)
 CLSID_AviDest = iid
End Function

Public Function CLSID_AviMuxProptyPage() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC647B5C0, CInt(&H157C), CInt(&H11D0), &HBD, &H23, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_AviMuxProptyPage = iid
End Function

Public Function CLSID_AviMuxProptyPage1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA9AE910, CInt(&H85C0), CInt(&H11D0), &HBD, &H42, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_AviMuxProptyPage1 = iid
End Function

Public Function CLSID_AVIMIDIRender() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7B65360, CInt(&HC445), CInt(&H11CE), &HAF, &HDE, &H0, &HAA, &H0, &H6C, &H14, &HF4)
 CLSID_AVIMIDIRender = iid
End Function

Public Function CLSID_WMAsfReader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H187463A0, CInt(&H5BB7), CInt(&H11D3), &HAC, &HBE, &H0, &H80, &HC7, &H5E, &H24, &H6E)
 CLSID_WMAsfReader = iid
End Function

Public Function CLSID_WMAsfWriter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7C23220E, CInt(&H55BB), CInt(&H11D3), &H8B, &H16, &H0, &HC0, &H4F, &HB6, &HBD, &H3D)
 CLSID_WMAsfWriter = iid
End Function

Public Function CLSID_MPEG2Demultiplexer() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAFB6C280, CInt(&H2C41), CInt(&H11D3), &H8A, &H60, &H0, &H0, &HF8, &H1E, &HE, &H4A)
 CLSID_MPEG2Demultiplexer = iid
End Function

Public Function CLSID_MPEG2Demultiplexer_NoClock() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H687D3367, CInt(&H3644), CInt(&H467A), &HAD, &HFE, &H6C, &HD7, &HA8, &H5C, &H4A, &H2C)
 CLSID_MPEG2Demultiplexer_NoClock = iid
End Function

Public Function CLSID_MMSPLITTER() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3AE86B20, CInt(&H7BE8), CInt(&H11D1), &HAB, &HE6, &H0, &HA0, &HC9, &H5, &HF3, &H75)
 CLSID_MMSPLITTER = iid
End Function

Public Function CLSID_StreamBufferSink() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2DB47AE5, CInt(&HCF39), CInt(&H43C2), &HB4, &HD6, &HC, &HD8, &HD9, &H9, &H46, &HF4)
 CLSID_StreamBufferSink = iid
End Function

Public Function CLSID_SBE2Sink() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE2448508, CInt(&H95DA), CInt(&H4205), &H9A, &H27, &H7E, &HC8, &H1E, &H72, &H3B, &H1A)
 CLSID_SBE2Sink = iid
End Function

Public Function CLSID_StreamBufferSource() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC9F5FE02, CInt(&HF851), CInt(&H4EB5), &H99, &HEE, &HAD, &H60, &H2A, &HF1, &HE6, &H19)
 CLSID_StreamBufferSource = iid
End Function

Public Function CLSID_StreamBufferConfig() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFA8A68B2, CInt(&HC864), CInt(&H4BA2), &HAD, &H53, &HD3, &H87, &H6A, &H87, &H49, &H4B)
 CLSID_StreamBufferConfig = iid
End Function

Public Function CLSID_StreamBufferPropertyHandler() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE37A73F8, CInt(&HFB01), CInt(&H43DC), &H91, &H4E, &HAA, &HEE, &H76, &H9, &H5A, &HB9)
 CLSID_StreamBufferPropertyHandler = iid
End Function

Public Function CLSID_StreamBufferThumbnailHandler() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H713790EE, CInt(&H5EE1), CInt(&H45BA), &H80, &H70, &HA1, &H33, &H7D, &H27, &H62, &HFA)
 CLSID_StreamBufferThumbnailHandler = iid
End Function

Public Function CLSID_Mpeg2VideoStreamAnalyzer() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6CFAD761, CInt(&H735D), CInt(&H4AA5), &H8A, &HFC, &HAF, &H91, &HA7, &HD6, &H1E, &HBA)
 CLSID_Mpeg2VideoStreamAnalyzer = iid
End Function

Public Function CLSID_StreamBufferRecordingAttributes() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCCAA63AC, CInt(&H1057), CInt(&H4778), &HAE, &H92, &H12, &H6, &HAB, &H9A, &HCE, &HE6)
 CLSID_StreamBufferRecordingAttributes = iid
End Function

Public Function CLSID_StreamBufferComposeRecording() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD682C4BA, CInt(&HA90A), CInt(&H42FE), &HB9, &HE1, &H3, &H10, &H98, &H49, &HC4, &H23)
 CLSID_StreamBufferComposeRecording = iid
End Function

Public Function CLSID_SBE2File() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H93A094D7, CInt(&H51E8), CInt(&H485B), &H90, &H4A, &H8D, &H6B, &H97, &HDC, &H6B, &H39)
 CLSID_SBE2File = iid
End Function

Public Function CLSID_DVVideoCodec() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB1B77C00, CInt(&HC3E4), CInt(&H11CF), &HAF, &H79, &H0, &HAA, &H0, &HB6, &H7A, &H42)
 CLSID_DVVideoCodec = iid
End Function

Public Function CLSID_DVVideoEnc() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H13AA3650, CInt(&HBB6F), CInt(&H11D0), &HAF, &HB9, &H0, &HAA, &H0, &HB6, &H7A, &H42)
 CLSID_DVVideoEnc = iid
End Function

Public Function CLSID_DVSplitter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4EB31670, CInt(&H9FC6), CInt(&H11CF), &HAF, &H6E, &H0, &HAA, &H0, &HB6, &H7A, &H42)
 CLSID_DVSplitter = iid
End Function

Public Function CLSID_DVMux() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H129D7E40, CInt(&HC10D), CInt(&H11D0), &HAF, &HB9, &H0, &HAA, &H0, &HB6, &H7A, &H42)
 CLSID_DVMux = iid
End Function

Public Function CLSID_SeekingPassThru() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H60AF76C, CInt(&H68DD), CInt(&H11D0), &H8F, &HC1, &H0, &HC0, &H4F, &HD9, &H18, &H9D)
 CLSID_SeekingPassThru = iid
End Function

Public Function CLSID_Line21Decoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6E8D4A20, CInt(&H310C), CInt(&H11D0), &HB7, &H9A, &H0, &HAA, &H0, &H37, &H67, &HA7)
 CLSID_Line21Decoder = iid
End Function

Public Function CLSID_Line21Decoder2() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE4206432, CInt(&H1A1), CInt(&H4BEE), &HB3, &HE1, &H37, &H2, &HC8, &HED, &HC5, &H74)
 CLSID_Line21Decoder2 = iid
End Function

Public Function CLSID_CCAFilter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3D07A539, CInt(&H35CA), CInt(&H447C), &H9B, &H5, &H8D, &H85, &HCE, &H92, &H4F, &H9E)
 CLSID_CCAFilter = iid
End Function

Public Function CLSID_OverlayMixer() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCD8743A1, CInt(&H3736), CInt(&H11D0), &H9E, &H69, &H0, &HC0, &H4F, &HD7, &HC1, &H5B)
 CLSID_OverlayMixer = iid
End Function

Public Function CLSID_VBISurfaces() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H814B9800, CInt(&H1C88), CInt(&H11D1), &HBA, &HD9, &H0, &H60, &H97, &H44, &H11, &H1A)
 CLSID_VBISurfaces = iid
End Function

Public Function CLSID_WSTDecoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H70BC06E0, CInt(&H5666), CInt(&H11D3), &HA1, &H84, &H0, &H10, &H5A, &HEF, &H9F, &H33)
 CLSID_WSTDecoder = iid
End Function

Public Function CLSID_MjpegDec() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H301056D0, CInt(&H6DFF), CInt(&H11D2), &H9E, &HEB, &H0, &H60, &H8, &H3, &H9E, &H37)
 CLSID_MjpegDec = iid
End Function

Public Function CLSID_MJPGEnc() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB80AB0A0, CInt(&H7416), CInt(&H11D2), &H9E, &HEB, &H0, &H60, &H8, &H3, &H9E, &H37)
 CLSID_MJPGEnc = iid
End Function

Public Function CLSID_SystemDeviceEnum() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H62BE5D10, CInt(&H60EB), CInt(&H11D0), &HBD, &H3B, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_SystemDeviceEnum = iid
End Function

Public Function CLSID_CDeviceMoniker() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4315D437, CInt(&H5B8C), CInt(&H11D0), &HBD, &H3B, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_CDeviceMoniker = iid
End Function

Public Function CLSID_VideoInputDeviceCategory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H860BB310, CInt(&H5D01), CInt(&H11D0), &HBD, &H3B, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_VideoInputDeviceCategory = iid
End Function

Public Function CLSID_CVidCapClassManager() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H860BB310, CInt(&H5D01), CInt(&H11D0), &HBD, &H3B, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_CVidCapClassManager = iid
End Function

Public Function CLSID_LegacyAmFilterCategory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H83863F1, CInt(&H70DE), CInt(&H11D0), &HBD, &H40, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_LegacyAmFilterCategory = iid
End Function

Public Function CLSID_CQzFilterClassManager() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H83863F1, CInt(&H70DE), CInt(&H11D0), &HBD, &H40, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_CQzFilterClassManager = iid
End Function

Public Function CLSID_VideoCompressorCategory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H33D9A760, CInt(&H90C8), CInt(&H11D0), &HBD, &H43, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_VideoCompressorCategory = iid
End Function

Public Function CLSID_CIcmCoClassManager() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H33D9A760, CInt(&H90C8), CInt(&H11D0), &HBD, &H43, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_CIcmCoClassManager = iid
End Function

Public Function CLSID_AudioCompressorCategory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H33D9A761, CInt(&H90C8), CInt(&H11D0), &HBD, &H43, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_AudioCompressorCategory = iid
End Function

Public Function CLSID_CAcmCoClassManager() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H33D9A761, CInt(&H90C8), CInt(&H11D0), &HBD, &H43, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_CAcmCoClassManager = iid
End Function

Public Function CLSID_AudioInputDeviceCategory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H33D9A762, CInt(&H90C8), CInt(&H11D0), &HBD, &H43, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_AudioInputDeviceCategory = iid
End Function

Public Function CLSID_CWaveinClassManager() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H33D9A762, CInt(&H90C8), CInt(&H11D0), &HBD, &H43, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_CWaveinClassManager = iid
End Function

Public Function CLSID_AudioRendererCategory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE0F158E1, CInt(&HCB04), CInt(&H11D0), &HBD, &H4E, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_AudioRendererCategory = iid
End Function

Public Function CLSID_CWaveOutClassManager() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE0F158E1, CInt(&HCB04), CInt(&H11D0), &HBD, &H4E, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_CWaveOutClassManager = iid
End Function

Public Function CLSID_MidiRendererCategory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4EFE2452, CInt(&H168A), CInt(&H11D1), &HBC, &H76, &H0, &HC0, &H4F, &HB9, &H45, &H3B)
 CLSID_MidiRendererCategory = iid
End Function

Public Function CLSID_CMidiOutClassManager() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4EFE2452, CInt(&H168A), CInt(&H11D1), &HBC, &H76, &H0, &HC0, &H4F, &HB9, &H45, &H3B)
 CLSID_CMidiOutClassManager = iid
End Function

Public Function CLSID_TransmitCategory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCC7BFB41, CInt(&HF175), CInt(&H11D1), &HA3, &H92, &H0, &HE0, &H29, &H1F, &H39, &H59)
 CLSID_TransmitCategory = iid
End Function

Public Function CLSID_DeviceControlCategory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCC7BFB46, CInt(&HF175), CInt(&H11D1), &HA3, &H92, &H0, &HE0, &H29, &H1F, &H39, &H59)
 CLSID_DeviceControlCategory = iid
End Function

Public Function CLSID_ActiveMovieCategories() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDA4E3DA0, CInt(&HD07D), CInt(&H11D0), &HBD, &H50, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_ActiveMovieCategories = iid
End Function

Public Function CLSID_DVDHWDecodersCategory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2721AE20, CInt(&H7E70), CInt(&H11D0), &HA5, &HD6, &H28, &HDB, &H4, &HC1, &H0, &H0)
 CLSID_DVDHWDecodersCategory = iid
End Function

Public Function CLSID_MediaEncoderCategory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7D22E920, CInt(&H5CA9), CInt(&H4787), &H8C, &H2B, &HA6, &H77, &H9B, &HD1, &H17, &H81)
 CLSID_MediaEncoderCategory = iid
End Function

Public Function CLSID_MediaMultiplexerCategory() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H236C9559, CInt(&HADCE), CInt(&H4736), &HBF, &H72, &HBA, &HB3, &H4E, &H39, &H21, &H96)
 CLSID_MediaMultiplexerCategory = iid
End Function

Public Function CLSID_FilterMapper2() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCDA42200, CInt(&HBD88), CInt(&H11D0), &HBD, &H4E, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_FilterMapper2 = iid
End Function

Public Function CLSID_MemoryAllocator() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1E651CC0, CInt(&HB199), CInt(&H11D0), &H82, &H12, &H0, &HC0, &H4F, &HC3, &H2C, &H45)
 CLSID_MemoryAllocator = iid
End Function

Public Function CLSID_MediaPropertyBag() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCDBD8D00, CInt(&HC193), CInt(&H11D0), &HBD, &H4E, &H0, &HA0, &HC9, &H11, &HCE, &H86)
 CLSID_MediaPropertyBag = iid
End Function

Public Function CLSID_DvdGraphBuilder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFCC152B7, CInt(&HF372), CInt(&H11D0), &H8E, &H0, &H0, &HC0, &H4F, &HD7, &HC0, &H8B)
 CLSID_DvdGraphBuilder = iid
End Function

Public Function CLSID_DVDNavigator() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9B8C4620, CInt(&H2C1A), CInt(&H11D0), &H84, &H93, &H0, &HA0, &H24, &H38, &HAD, &H48)
 CLSID_DVDNavigator = iid
End Function

Public Function CLSID_DVDState() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF963C5CF, CInt(&HA659), CInt(&H4A93), &H96, &H38, &HCA, &HF3, &HCD, &H27, &H7D, &H13)
 CLSID_DVDState = iid
End Function

Public Function CLSID_SmartTee() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCC58E280, CInt(&H8AA1), CInt(&H11D1), &HB3, &HF1, &H0, &HAA, &H0, &H37, &H61, &HC5)
 CLSID_SmartTee = iid
End Function

Public Function CLSID_DtvCcFilter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFB056BA0, CInt(&H2502), CInt(&H45B9), &H8E, &H86, &H2B, &H40, &HDE, &H84, &HAD, &H29)
 CLSID_DtvCcFilter = iid
End Function

Public Function CLSID_CaptionsFilter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2F7EE4B6, CInt(&H6FF5), CInt(&H4EB4), &HB2, &H4A, &H2B, &HFC, &H41, &H11, &H71, &H71)
 CLSID_CaptionsFilter = iid
End Function

Public Function CLSID_SubtitlesFilter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9F22CFEA, CInt(&HCE07), CInt(&H41AB), &H8B, &HA0, &HC7, &H36, &H4A, &HF9, &HA, &HF9)
 CLSID_SubtitlesFilter = iid
End Function

Public Function CLSID_DirectShowPluginControl() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8670C736, CInt(&HF614), CInt(&H427B), &H8A, &HDA, &HBB, &HAD, &HC5, &H87, &H19, &H4B)
 CLSID_DirectShowPluginControl = iid
End Function



Public Function IID_IAMCollection() As UUID
'{56A868B9-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B9, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IAMCollection = iid
End Function
Public Function IID_IMediaControl() As UUID
'{56A868B1-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B1, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IMediaControl = iid
End Function
Public Function IID_IMediaEvent() As UUID
'{56A868B6-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B6, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IMediaEvent = iid
End Function
Public Function IID_IMediaEventEx() As UUID
'{56A868C0-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868C0, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IMediaEventEx = iid
End Function
Public Function IID_IMediaPosition() As UUID
'{56A868B2-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B2, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IMediaPosition = iid
End Function
Public Function IID_IBasicAudio() As UUID
'{56A868B3-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B3, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IBasicAudio = iid
End Function
Public Function IID_IVideoWindow() As UUID
'{56A868B4-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B4, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IVideoWindow = iid
End Function
Public Function IID_IBasicVideo() As UUID
'{56A868B5-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B5, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IBasicVideo = iid
End Function
Public Function IID_IBasicVideo2() As UUID
'{329BB360-F6EA-11D1-9038-00A0C9697298}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H329BB360, CInt(&HF6EA), CInt(&H11D1), &H90, &H38, &H0, &HA0, &HC9, &H69, &H72, &H98)
IID_IBasicVideo2 = iid
End Function
Public Function IID_IDeferredCommand() As UUID
'{56A868B8-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B8, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IDeferredCommand = iid
End Function
Public Function IID_IQueueCommand() As UUID
'{56A868B7-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B7, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IQueueCommand = iid
End Function
Public Function IID_IFilterInfo() As UUID
'{E436EBB3-524F-11CE-9F53-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE436EBB3, CInt(&H524F), CInt(&H11CE), &H9F, &H53, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IFilterInfo = iid
End Function
Public Function IID_IRegFilterInfo() As UUID
'{56A868BB-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868BB, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IRegFilterInfo = iid
End Function
Public Function IID_IMediaTypeInfo() As UUID
'{56A868BC-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868BC, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IMediaTypeInfo = iid
End Function
Public Function IID_IPinInfo() As UUID
'{56A868BD-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868BD, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IPinInfo = iid
End Function
Public Function IID_IAMStats() As UUID
'{BC9BCF80-DCD2-11D2-ABF6-00A0C905F375}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBC9BCF80, CInt(&HDCD2), CInt(&H11D2), &HAB, &HF6, &H0, &HA0, &HC9, &H5, &HF3, &H75)
IID_IAMStats = iid
End Function
Public Function IID_IEnumMediaTypes() As UUID
'{89c31040-846b-11ce-97d3-00aa0055595a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H89C31040, CInt(&H846B), CInt(&H11CE), &H97, &HD3, &H0, &HAA, &H0, &H55, &H59, &H5A)
IID_IEnumMediaTypes = iid
End Function
Public Function IID_IPin() As UUID
'{56a86891-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A86891, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IPin = iid
End Function
Public Function IID_IEnumPins() As UUID
'{56a86892-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A86892, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IEnumPins = iid
End Function
Public Function IID_IReferenceClock() As UUID
'{56a86897-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A86897, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IReferenceClock = iid
End Function
Public Function IID_IMediaFilter() As UUID
'{56a86899-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A86899, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IMediaFilter = iid
End Function
Public Function IID_IBaseFilter() As UUID
'{56a86895-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A86895, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IBaseFilter = iid
End Function
Public Function IID_IEnumFilters() As UUID
'{56a86893-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A86893, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IEnumFilters = iid
End Function
Public Function IID_IFilterGraph() As UUID
'{56a8689f-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A8689F, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IFilterGraph = iid
End Function
Public Function IID_IFilterGraph2() As UUID
'{36b73882-c2c8-11cf-8b46-00805f6cef60}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H36B73882, CInt(&HC2C8), CInt(&H11CF), &H8B, &H46, &H0, &H80, &H5F, &H6C, &HEF, &H60)
 IID_IFilterGraph2 = iid
End Function
Public Function IID_IFilterGraph3() As UUID
'{aaf38154-b80b-422f-91e6-b66467509a07}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAAF38154, CInt(&HB80B), CInt(&H422F), &H91, &HE6, &HB6, &H64, &H67, &H50, &H9A, &H7)
 IID_IFilterGraph3 = iid
End Function
Public Function IID_IFileSinkFilter() As UUID
'{a2104830-7c70-11cf-8bce-00aa00a3f1a6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA2104830, CInt(&H7C70), CInt(&H11CF), &H8B, &HCE, &H0, &HAA, &H0, &HA3, &HF1, &HA6)
IID_IFileSinkFilter = iid
End Function
Public Function IID_IAMCopyCaptureFileProgress() As UUID
'{670d1d20-a068-11d0-b3f0-00aa003761c5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H670D1D20, CInt(&HA068), CInt(&H11D0), &HB3, &HF0, &H0, &HAA, &H0, &H37, &H61, &HC5)
IID_IAMCopyCaptureFileProgress = iid
End Function
Public Function IID_IGraphBuilder() As UUID
'{56a868a9-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868A9, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IGraphBuilder = iid
End Function
Public Function IID_ICaptureGraphBuilder() As UUID
'{bf87b6e0-8c27-11d0-b3f0-00aa003761c5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBF87B6E0, CInt(&H8C27), CInt(&H11D0), &HB3, &HF0, &H0, &HAA, &H0, &H37, &H61, &HC5)
IID_ICaptureGraphBuilder = iid
End Function
Public Function IID_ICaptureGraphBuilder2() As UUID
'{93E5A4E0-2D50-11d2-ABFA-00A0C9C6E38D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H93E5A4E0, CInt(&H2D50), CInt(&H11D2), &HAB, &HFA, &H0, &HA0, &HC9, &HC6, &HE3, &H8D)
IID_ICaptureGraphBuilder2 = iid
End Function
Public Function IID_IAMChannelInfo() As UUID
'{FA2AA8F1-8B62-11D0-A520-000000000000}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA2AA8F1, CInt(&H8B62), CInt(&H11D0), &HA5, &H20, &H0, &H0, &H0, &H0, &H0, &H0)
IID_IAMChannelInfo = iid
End Function
Public Function IID_IAMNetworkStatus() As UUID
'{FA2AA8F3-8B62-11D0-A520-000000000000}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA2AA8F3, CInt(&H8B62), CInt(&H11D0), &HA5, &H20, &H0, &H0, &H0, &H0, &H0, &H0)
IID_IAMNetworkStatus = iid
End Function
Public Function IID_IAMNetShowExProps() As UUID
'{FA2AA8F5-8B62-11D0-A520-000000000000}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA2AA8F5, CInt(&H8B62), CInt(&H11D0), &HA5, &H20, &H0, &H0, &H0, &H0, &H0, &H0)
IID_IAMNetShowExProps = iid
End Function
Public Function IID_IAMExtendedErrorInfo() As UUID
'{FA2AA8F6-8B62-11D0-A520-000000000000}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA2AA8F6, CInt(&H8B62), CInt(&H11D0), &HA5, &H20, &H0, &H0, &H0, &H0, &H0, &H0)
IID_IAMExtendedErrorInfo = iid
End Function
Public Function IID_IAMNetShowPreroll() As UUID
'{AAE7E4E2-6388-11D1-8D93-006097C9A2B2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAAE7E4E2, CInt(&H6388), CInt(&H11D1), &H8D, &H93, &H0, &H60, &H97, &HC9, &HA2, &HB2)
IID_IAMNetShowPreroll = iid
End Function
Public Function IID_IAMMediaContent() As UUID
'{FA2AA8F4-8B62-11D0-A520-000000000000}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA2AA8F4, CInt(&H8B62), CInt(&H11D0), &HA5, &H20, &H0, &H0, &H0, &H0, &H0, &H0)
IID_IAMMediaContent = iid
End Function
Public Function IID_IAMExtendedSeeking() As UUID
'{FA2AA8F9-8B62-11D0-A520-000000000000}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA2AA8F9, CInt(&H8B62), CInt(&H11D0), &HA5, &H20, &H0, &H0, &H0, &H0, &H0, &H0)
IID_IAMExtendedSeeking = iid
End Function
Public Function IID_IAMMediaContent2() As UUID
'{CE8F78C1-74D9-11D2-B09D-00A0C9A81117}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCE8F78C1, CInt(&H74D9), CInt(&H11D2), &HB0, &H9D, &H0, &HA0, &HC9, &HA8, &H11, &H17)
IID_IAMMediaContent2 = iid
End Function
Public Function IID_IAMAnalogVideoDecoder() As UUID
'{C6E13350-30AC-11d0-A18C-00A0C9118956}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC6E13350, CInt(&H30AC), CInt(&H11D0), &HA1, &H8C, &H0, &HA0, &HC9, &H11, &H89, &H56)
IID_IAMAnalogVideoDecoder = iid
End Function
Public Function IID_IAMAsyncReaderTimestampScaling() As UUID
'{cf7b26fc-9a00-485b-8147-3e789d5e8f67}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCF7B26FC, CInt(&H9A00), CInt(&H485B), &H81, &H47, &H3E, &H78, &H9D, &H5E, &H8F, &H67)
IID_IAMAsyncReaderTimestampScaling = iid
End Function
Public Function IID_IAMAudioInputMixer() As UUID
'{54C39221-8380-11d0-B3F0-00AA003761C5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H54C39221, CInt(&H8380), CInt(&H11D0), &HB3, &HF0, &H0, &HAA, &H0, &H37, &H61, &HC5)
IID_IAMAudioInputMixer = iid
End Function
Public Function IID_IAMAudioRendererStats() As UUID
'{22320CB2-D41A-11d2-BF7C-D7CB9DF0BF93}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H22320CB2, CInt(&HD41A), CInt(&H11D2), &HBF, &H7C, &HD7, &HCB, &H9D, &HF0, &HBF, &H93)
IID_IAMAudioRendererStats = iid
End Function
Public Function IID_IAMBufferNegotiation() As UUID
'{56ED71A0-AF5F-11D0-B3F0-00AA003761C5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56ED71A0, CInt(&HAF5F), CInt(&H11D0), &HB3, &HF0, &H0, &HAA, &H0, &H37, &H61, &HC5)
IID_IAMBufferNegotiation = iid
End Function
Public Function IID_IAMCameraControl() As UUID
'{C6E13370-30AC-11d0-A18C-00A0C9118956}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC6E13370, CInt(&H30AC), CInt(&H11D0), &HA1, &H8C, &H0, &HA0, &HC9, &H11, &H89, &H56)
IID_IAMCameraControl = iid
End Function
Public Function IID_IAMCertifiedOutputProtection() As UUID
'{6feded3e-0ff1-4901-a2f1-43f7012c8515}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FEDED3E, CInt(&HFF1), CInt(&H4901), &HA2, &HF1, &H43, &HF7, &H1, &H2C, &H85, &H15)
IID_IAMCertifiedOutputProtection = iid
End Function
Public Function IID_IAMClockAdjust() As UUID
'{4d5466b0-a49c-11d1-abe8-00a0c905f375}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4D5466B0, CInt(&HA49C), CInt(&H11D1), &HAB, &HE8, &H0, &HA0, &HC9, &H5, &HF3, &H75)
IID_IAMClockAdjust = iid
End Function
Public Function IID_IAMClockSlave() As UUID
'{9FD52741-176D-4b36-8F51-CA8F933223BE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9FD52741, CInt(&H176D), CInt(&H4B36), &H8F, &H51, &HCA, &H8F, &H93, &H32, &H23, &HBE)
IID_IAMClockSlave = iid
End Function
Public Function IID_IAMCrossbar() As UUID
'{C6E13380-30AC-11d0-A18C-00A0C9118956}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC6E13380, CInt(&H30AC), CInt(&H11D0), &HA1, &H8C, &H0, &HA0, &HC9, &H11, &H89, &H56)
IID_IAMCrossbar = iid
End Function
Public Function IID_IAMDecoderCaps() As UUID
'{c0dff467-d499-4986-972b-e1d9090fa941}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC0DFF467, CInt(&HD499), CInt(&H4986), &H97, &H2B, &HE1, &HD9, &H9, &HF, &HA9, &H41)
IID_IAMDecoderCaps = iid
End Function
Public Function IID_IMediaSample() As UUID
'{56a8689a-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H56A8689A, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
 IID_IMediaSample = iid
End Function
Public Function IID_IMediaSample2() As UUID
'{36b73884-c2c8-11cf-8b46-00805f6cef60}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H36B73884, CInt(&HC2C8), CInt(&H11CF), &H8B, &H46, &H0, &H80, &H5F, &H6C, &HEF, &H60)
 IID_IMediaSample2 = iid
End Function
Public Function IID_ISampleGrabber() As UUID
'{6B652FFF-11FE-4fce-92AD-0266B5D7C78F}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6B652FFF, CInt(&H11FE), CInt(&H4FCE), &H92, &HAD, &H2, &H66, &HB5, &HD7, &HC7, &H8F)
 IID_ISampleGrabber = iid
End Function
Public Function IID_ISampleGrabberCB() As UUID
'{0579154a-2b53-4994-b0d0-e773148eff85}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H579154A, CInt(&H2B53), CInt(&H4994), &HB0, &HD0, &HE7, &H73, &H14, &H8E, &HFF, &H85)
 IID_ISampleGrabberCB = iid
End Function
