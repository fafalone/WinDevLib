Attribute VB_Name = "mSpeech"
Option Explicit

'IIDs and Constants for Microsoft Speech API Interfaces

    Public Const Speech_Default_Weight As Single = 1 '// DEFAULT_WEIGHT;
    Public Const Speech_Max_Word_Length = 128 '// SP_MAX_WORD_LENGTH;
    Public Const Speech_Max_Pron_Length = 384 ' / SP_MAX_PRON_LENGTH;
    Public Const Speech_StreamPos_Asap = 0 '// SP_STREAMPOS_ASAP;
    Public Const Speech_StreamPos_RealTime = -1 ', //SP_STREAMPOS_REALTIME;
    Public Const SpeechAllElements = -1 ', //SPPR_ALL_ELEMENTS;

    '//--- Root of registry entries for speech use
Public Const SpeechRegistryUserRoot = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Speech"
Public Const SpeechRegistryLocalMachineRoot = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech"

    '//--- Object Token Categories for speech resource management
Public Const SpeechCategoryAudioOut = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\AudioOutput"
Public Const SpeechCategoryAudioIn = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\AudioInput"
Public Const SpeechCategoryVoices = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices"
Public Const SpeechCategoryRecognizers = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Recognizers"
Public Const SpeechCategoryAppLexicons = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\AppLexicons"
Public Const SpeechCategoryPhoneConverters = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\PhoneConverters"
Public Const SpeechCategoryRecoProfiles = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Speech\RecoProfiles"

    '//--- User Lexicon Token Id
Public Const SpeechTokenIdUserLexicon = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Speech\CurrentUserLexicon"

    '//--- Standard token values
Public Const SpeechTokenValueCLSID = "CLSID"
Public Const SpeechTokenKeyFiles = "Files"
Public Const SpeechTokenKeyUI = "UI"
Public Const SpeechTokenKeyAttributes = "Attributes"

    '//--- Standard voice category values
Public Const SpeechVoiceCategoryTTSRate = "DefaultTTSRate"

    '//--- Standard SR Engine properties
Public Const SpeechPropertyResourceUsage = "ResourceUsage"
Public Const SpeechPropertyHighConfidenceThreshold = "HighConfidenceThreshold"
Public Const SpeechPropertyNormalConfidenceThreshold = "NormalConfidenceThreshold"
Public Const SpeechPropertyLowConfidenceThreshold = "LowConfidenceThreshold"
Public Const SpeechPropertyResponseSpeed = "ResponseSpeed"
Public Const SpeechPropertyComplexResponseSpeed = "ComplexResponseSpeed"
Public Const SpeechPropertyAdaptationOn = "AdaptationOn"

    '//--- Standard SAPI Recognition Topics
Public Const SpeechDictationTopicSpelling = "Spelling"

    '//--- Special Tags used in SR grammars
Public Const SpeechGrammarTagWildcard = "..."
Public Const SpeechGrammarTagDictation = "*"
Public Const SpeechGrammarTagUnlimitedDictation = "*+"

    '//--- TokenUI constants
Public Const SpeechEngineProperties = "EngineProperties"
Public Const SpeechAddRemoveWord = "AddRemoveWord"
Public Const SpeechUserTraining = "UserTraining"
Public Const SpeechMicTraining = "MicTraining"
Public Const SpeechRecoProfileProperties = "RecoProfileProperties"
Public Const SpeechAudioProperties = "AudioProperties"
Public Const SpeechAudioVolume = "AudioVolume"

    '//--- ISpeechVoice::Skip constant
Public Const SpeechVoiceSkipTypeSentence = "Sentence"

    ' // The SpeechAudioFormat object includes a guid that can be used to set the format to
    ' //  a custom value.
Public Const SpeechAudioFormatGUIDWave = "{C31ADBAE-527F-4ff5-A230-F62BB61FF70C}"
Public Const SpeechAudioFormatGUIDText = "{7CEEF9F9-3D13-11d2-9EE7-00C04F797396}"

Public Const SP_LOW_CONFIDENCE As Byte = 255
Public Const SP_NORMAL_CONFIDENCE As Byte = 0
Public Const SP_HIGH_CONFIDENCE As Byte = 1

' // CFG default weight
' // MIDL does not support floating point in the RHS.
' // Thus, using 1.0 instead of 1 resulted in unexpected behavior in the resulting Type library.
Public Const DEFAULT_WEIGHT As Single = 1

' // Lexicon word and pronunciation limits
Public Const SP_MAX_WORD_LENGTH As Long = 128
Public Const SP_MAX_PRON_LENGTH As Long = 384

' //#If _SAPI_BUILD_VER >= 0x053
' // Flag used in EmulateRecognitionEx to indicate re-sending an existing result
Public Const SP_EMULATE_RESULT As Long = &H40000000

'//--- TokenUI constants
Public Const SPDUI_EngineProperties = "EngineProperties"
Public Const SPDUI_AddRemoveWord = "AddRemoveWord"
Public Const SPDUI_UserTraining = "UserTraining"
Public Const SPDUI_MicTraining = "MicTraining"
Public Const SPDUI_RecoProfileProperties = "RecoProfileProperties"
Public Const SPDUI_AudioProperties = "AudioProperties"
Public Const SPDUI_AudioVolume = "AudioVolume"
Public Const SPDUI_UserEnrollment = "UserEnrollment"
Public Const SPDUI_ShareData = "ShareData"

'// new for Vista.  Nothing prevents use downlevel if an engine exposes them
Public Const SPDUI_Tutorial = "Tutorial"

'//--- Root of registry entries for speech use
Public Const SPREG_USER_ROOT = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Speech"
Public Const SPREG_LOCAL_MACHINE_ROOT = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech"

'//--- Categories for speech resource management
Public Const SPCAT_AUDIOOUT = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\AudioOutput"
Public Const SPCAT_AUDIOIN = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\AudioInput"
Public Const SPCAT_VOICES = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices"
Public Const SPCAT_RECOGNIZERS = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Recognizers"
Public Const SPCAT_APPLEXICONS = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\AppLexicons"
Public Const SPCAT_PHONECONVERTERS = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\PhoneConverters"
Public Const SPCAT_TEXTNORMALIZERS = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\TextNormalizers"
Public Const SPCAT_RECOPROFILES = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Speech\RecoProfiles"

'//--- Specific token ids of interest
Public Const SPMMSYS_AUDIO_IN_TOKEN_ID = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\AudioInput\TokenEnums\MMAudioIn\"
Public Const SPMMSYS_AUDIO_OUT_TOKEN_ID = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\AudioOutput\TokenEnums\MMAudioOut\"
Public Const SPCURRENT_USER_LEXICON_TOKEN_ID = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Speech\CurrentUserLexicon"
'// Shortcuts only supported on Vista and above
Public Const SPCURRENT_USER_SHORTCUT_TOKEN_ID = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Speech\CurrentUserShortcut"

'//--- Standard token values
Public Const SPTOKENVALUE_CLSID = "CLSID"
Public Const SPTOKENKEY_FILES = "Files"
Public Const SPTOKENKEY_UI = "UI"
Public Const SPTOKENKEY_ATTRIBUTES = "Attributes"

Public Const SPTOKENKEY_RETAINEDAUDIO = "SecondsPerRetainedAudioEvent"
Public Const SPTOKENKEY_AUDIO_LATENCY_WARNING = "LatencyWarningThreshold"
Public Const SPTOKENKEY_AUDIO_LATENCY_TRUNCATE = "LatencyTruncateThreshold"
Public Const SPTOKENKEY_AUDIO_LATENCY_UPDATE_INTERVAL = "LatencyUpdateInterval"


'//--- Standard voice category values
Public Const SPVOICECATEGORY_TTSRATE = "DefaultTTSRate"

'//--- Standard SR Engine properties
Public Const SPPROP_RESOURCE_USAGE = "ResourceUsage"
Public Const SPPROP_HIGH_CONFIDENCE_THRESHOLD = "HighConfidenceThreshold"
Public Const SPPROP_NORMAL_CONFIDENCE_THRESHOLD = "NormalConfidenceThreshold"
Public Const SPPROP_LOW_CONFIDENCE_THRESHOLD = "LowConfidenceThreshold"
Public Const SPPROP_RESPONSE_SPEED = "ResponseSpeed"
Public Const SPPROP_COMPLEX_RESPONSE_SPEED = "ComplexResponseSpeed"
Public Const SPPROP_ADAPTATION_ON = "AdaptationOn"

'// new for Vista, but nothing prevents engines that run downlevel from supporting these
Public Const SPPROP_PERSISTED_BACKGROUND_ADAPTATION = "PersistedBackgroundAdaptation"
Public Const SPPROP_PERSISTED_LANGUAGE_MODEL_ADAPTATION = "PersistedLanguageModelAdaptation"
Public Const SPPROP_UX_IS_LISTENING = "UXIsListening"

'//--- Standard SAPI Recognition Topics
Public Const SPTOPIC_SPELLING = "Spelling"

'// CFG Wildcard token
Public Const SPWILDCARD = "..."

'// CFG Dication token
Public Const SPDICTATION = "*"
Public Const SPINFDICTATION = "*+"


'// Registry key that stores a list of object token CLSIDs marked as safe to instantiate from HKCU
Public Const SPREG_SAFE_USER_TOKENS = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\UserTokens"

' Public Const SPFEI_FLAGCHECK As LongLong = ((1 << SPEI_RESERVED1) Or (1 << SPEI_RESERVED2))
Public Const SPFEI_FLAGCHECK As Currency = 966367.6416@ '((1 << 30) Or (1 << 33))
Public Const SPFEI_ALL_TTS_EVENTS As Currency = 966374.195@ ' (&H000000000000FFFE& Or SPFEI_FLAGCHECK)
Public Const SPFEI_ALL_SR_EVENTS As Currency = 112589883310.08@   '(&H003FFFFC00000000 Or SPFEI_FLAGCHECK)
Public Const SPFEI_ALL_EVENTS As Currency = -115292150460684.6977@ ' &HEFFFFFFFFFFFFFFF

Public Const SP_MAX_LANGIDS = 20 '        // Engine can support up to 20 languages at once

Public Const SP_STREAMPOS_ASAPS As Currency = 0
Public Const SP_STREAMPOS_REALTIMES As Currency = -1

Public Const SPRULETRANS_TEXTBUFFER As Long = (-1)
Public Const SPRULETRANS_WILDCARD As Long = (-2)
Public Const SPRULETRANS_DICTATION As Long = (-3)


Public Function IID_ISpeechDataKey() As UUID
'{CE17C09B-4EFA-44d5-A4C9-59D9585AB0CD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCE17C09B, CInt(&H4EFA), CInt(&H44D5), &HA4, &HC9, &H59, &HD9, &H58, &H5A, &HB0, &HCD)
IID_ISpeechDataKey = iid
End Function
Public Function IID_ISpeechObjectToken() As UUID
'{C74A3ADC-B727-4500-A84A-B526721C8B8C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC74A3ADC, CInt(&HB727), CInt(&H4500), &HA8, &H4A, &HB5, &H26, &H72, &H1C, &H8B, &H8C)
IID_ISpeechObjectToken = iid
End Function
Public Function IID_ISpeechObjectTokens() As UUID
'{9285B776-2E7B-4bc0-B53E-580EB6FA967F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9285B776, CInt(&H2E7B), CInt(&H4BC0), &HB5, &H3E, &H58, &HE, &HB6, &HFA, &H96, &H7F)
IID_ISpeechObjectTokens = iid
End Function
Public Function IID_ISpeechObjectTokenCategory() As UUID
'{CA7EAC50-2D01-4145-86D4-5AE7D70F4469}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCA7EAC50, CInt(&H2D01), CInt(&H4145), &H86, &HD4, &H5A, &HE7, &HD7, &HF, &H44, &H69)
IID_ISpeechObjectTokenCategory = iid
End Function
Public Function IID_ISpeechAudioFormat() As UUID
'{E6E9C590-3E18-40e3-8299-061F98BDE7C7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE6E9C590, CInt(&H3E18), CInt(&H40E3), &H82, &H99, &H6, &H1F, &H98, &HBD, &HE7, &HC7)
IID_ISpeechAudioFormat = iid
End Function
Public Function IID_ISpeechBaseStream() As UUID
'{6450336F-7D49-4ced-8097-49D6DEE37294}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6450336F, CInt(&H7D49), CInt(&H4CED), &H80, &H97, &H49, &HD6, &HDE, &HE3, &H72, &H94)
IID_ISpeechBaseStream = iid
End Function
Public Function IID_ISpeechAudio() As UUID
'{CFF8E175-019E-11d3-A08E-00C04F8EF9B5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCFF8E175, CInt(&H19E), CInt(&H11D3), &HA0, &H8E, &H0, &HC0, &H4F, &H8E, &HF9, &HB5)
IID_ISpeechAudio = iid
End Function
Public Function IID_ISpeechMMSysAudio() As UUID
'{3C76AF6D-1FD7-4831-81D1-3B71D5A13C44}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3C76AF6D, CInt(&H1FD7), CInt(&H4831), &H81, &HD1, &H3B, &H71, &HD5, &HA1, &H3C, &H44)
IID_ISpeechMMSysAudio = iid
End Function
Public Function IID_ISpeechFileStream() As UUID
'{AF67F125-AB39-4e93-B4A2-CC2E66E182A7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAF67F125, CInt(&HAB39), CInt(&H4E93), &HB4, &HA2, &HCC, &H2E, &H66, &HE1, &H82, &HA7)
IID_ISpeechFileStream = iid
End Function
Public Function IID_ISpeechCustomStream() As UUID
'{1A9E9F4F-104F-4db8-A115-EFD7FD0C97AE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1A9E9F4F, CInt(&H104F), CInt(&H4DB8), &HA1, &H15, &HEF, &HD7, &HFD, &HC, &H97, &HAE)
IID_ISpeechCustomStream = iid
End Function
Public Function IID_ISpeechMemoryStream() As UUID
'{EEB14B68-808B-4abe-A5EA-B51DA7588008}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEEB14B68, CInt(&H808B), CInt(&H4ABE), &HA5, &HEA, &HB5, &H1D, &HA7, &H58, &H80, &H8)
IID_ISpeechMemoryStream = iid
End Function
Public Function IID_ISpeechAudioStatus() As UUID
'{C62D9C91-7458-47f6-862D-1EF86FB0B278}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC62D9C91, CInt(&H7458), CInt(&H47F6), &H86, &H2D, &H1E, &HF8, &H6F, &HB0, &HB2, &H78)
IID_ISpeechAudioStatus = iid
End Function
Public Function IID_ISpeechAudioBufferInfo() As UUID
'{11B103D8-1142-4edf-A093-82FB3915F8CC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H11B103D8, CInt(&H1142), CInt(&H4EDF), &HA0, &H93, &H82, &HFB, &H39, &H15, &HF8, &HCC)
IID_ISpeechAudioBufferInfo = iid
End Function
Public Function IID_ISpeechWaveFormatEx() As UUID
'{7A1EF0D5-1581-4741-88E4-209A49F11A10}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7A1EF0D5, CInt(&H1581), CInt(&H4741), &H88, &HE4, &H20, &H9A, &H49, &HF1, &H1A, &H10)
IID_ISpeechWaveFormatEx = iid
End Function
Public Function IID_ISpeechVoice() As UUID
'{269316D8-57BD-11D2-9EEE-00C04F797396}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H269316D8, CInt(&H57BD), CInt(&H11D2), &H9E, &HEE, &H0, &HC0, &H4F, &H79, &H73, &H96)
IID_ISpeechVoice = iid
End Function
Public Function IID_ISpeechVoiceStatus() As UUID
'{8BE47B07-57F6-11d2-9EEE-00C04F797396}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8BE47B07, CInt(&H57F6), CInt(&H11D2), &H9E, &HEE, &H0, &HC0, &H4F, &H79, &H73, &H96)
IID_ISpeechVoiceStatus = iid
End Function
Public Function IID_ISpeechVoiceEvents() As UUID
'{A372ACD1-3BEF-4bbd-8FFB-CB3E2B416AF8}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA372ACD1, CInt(&H3BEF), CInt(&H4BBD), &H8F, &HFB, &HCB, &H3E, &H2B, &H41, &H6A, &HF8)
 IID_ISpeechVoiceEvents = iid
End Function
Public Function IID_ISpeechRecoContextEvents() As UUID
'{7B8FCB42-0E9D-4f00-A048-7B04D6179D3D}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7B8FCB42, CInt(&HE9D), CInt(&H4F00), &HA0, &H48, &H7B, &H4, &HD6, &H17, &H9D, &H3D)
 IID_ISpeechRecoContextEvents = iid
End Function
Public Function IID_ISpeechRecognizer() As UUID
'{2D5F1C0C-BD75-4b08-9478-3B11FEA2586C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2D5F1C0C, CInt(&HBD75), CInt(&H4B08), &H94, &H78, &H3B, &H11, &HFE, &HA2, &H58, &H6C)
IID_ISpeechRecognizer = iid
End Function
Public Function IID_ISpeechRecognizerStatus() As UUID
'{BFF9E781-53EC-484e-BB8A-0E1B5551E35C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBFF9E781, CInt(&H53EC), CInt(&H484E), &HBB, &H8A, &HE, &H1B, &H55, &H51, &HE3, &H5C)
IID_ISpeechRecognizerStatus = iid
End Function
Public Function IID_ISpeechRecoContext() As UUID
'{580AA49D-7E1E-4809-B8E2-57DA806104B8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H580AA49D, CInt(&H7E1E), CInt(&H4809), &HB8, &HE2, &H57, &HDA, &H80, &H61, &H4, &HB8)
IID_ISpeechRecoContext = iid
End Function
Public Function IID_ISpeechRecoGrammar() As UUID
'{B6D6F79F-2158-4e50-B5BC-9A9CCD852A09}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB6D6F79F, CInt(&H2158), CInt(&H4E50), &HB5, &HBC, &H9A, &H9C, &HCD, &H85, &H2A, &H9)
IID_ISpeechRecoGrammar = iid
End Function
Public Function IID_ISpeechGrammarRule() As UUID
'{AFE719CF-5DD1-44f2-999C-7A399F1CFCCC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAFE719CF, CInt(&H5DD1), CInt(&H44F2), &H99, &H9C, &H7A, &H39, &H9F, &H1C, &HFC, &HCC)
IID_ISpeechGrammarRule = iid
End Function
Public Function IID_ISpeechGrammarRules() As UUID
'{6FFA3B44-FC2D-40d1-8AFC-32911C7F1AD1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FFA3B44, CInt(&HFC2D), CInt(&H40D1), &H8A, &HFC, &H32, &H91, &H1C, &H7F, &H1A, &HD1)
IID_ISpeechGrammarRules = iid
End Function
Public Function IID_ISpeechGrammarRuleState() As UUID
'{D4286F2C-EE67-45ae-B928-28D695362EDA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD4286F2C, CInt(&HEE67), CInt(&H45AE), &HB9, &H28, &H28, &HD6, &H95, &H36, &H2E, &HDA)
IID_ISpeechGrammarRuleState = iid
End Function
Public Function IID_ISpeechGrammarRuleStateTransitions() As UUID
'{EABCE657-75BC-44a2-AA7F-C56476742963}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEABCE657, CInt(&H75BC), CInt(&H44A2), &HAA, &H7F, &HC5, &H64, &H76, &H74, &H29, &H63)
IID_ISpeechGrammarRuleStateTransitions = iid
End Function
Public Function IID_ISpeechGrammarRuleStateTransition() As UUID
'{CAFD1DB1-41D1-4a06-9863-E2E81DA17A9A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCAFD1DB1, CInt(&H41D1), CInt(&H4A06), &H98, &H63, &HE2, &HE8, &H1D, &HA1, &H7A, &H9A)
IID_ISpeechGrammarRuleStateTransition = iid
End Function
Public Function IID_ISpeechTextSelectionInformation() As UUID
'{3B9C7E7A-6EEE-4DED-9092-11657279ADBE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3B9C7E7A, CInt(&H6EEE), CInt(&H4DED), &H90, &H92, &H11, &H65, &H72, &H79, &HAD, &HBE)
IID_ISpeechTextSelectionInformation = iid
End Function
Public Function IID_ISpeechRecoResult() As UUID
'{ED2879CF-CED9-4ee6-A534-DE0191D5468D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HED2879CF, CInt(&HCED9), CInt(&H4EE6), &HA5, &H34, &HDE, &H1, &H91, &HD5, &H46, &H8D)
IID_ISpeechRecoResult = iid
End Function
Public Function IID_ISpeechXMLRecoResult() As UUID
'{AAEC54AF-8F85-4924-944D-B79D39D72E19}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAAEC54AF, CInt(&H8F85), CInt(&H4924), &H94, &H4D, &HB7, &H9D, &H39, &HD7, &H2E, &H19)
IID_ISpeechXMLRecoResult = iid
End Function
Public Function IID_ISpeechRecoResult2() As UUID
'{8E0A246D-D3C8-45de-8657-04290C458C3C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8E0A246D, CInt(&HD3C8), CInt(&H45DE), &H86, &H57, &H4, &H29, &HC, &H45, &H8C, &H3C)
IID_ISpeechRecoResult2 = iid
End Function
Public Function IID_ISpeechRecoResultDispatch() As UUID
'{6D60EB64-ACED-40a6-BBF3-4E557F71DEE2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6D60EB64, CInt(&HACED), CInt(&H40A6), &HBB, &HF3, &H4E, &H55, &H7F, &H71, &HDE, &HE2)
IID_ISpeechRecoResultDispatch = iid
End Function
Public Function IID_ISpeechPhraseInfoBuilder() As UUID
'{3B151836-DF3A-4E0A-846C-D2ADC9334333}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3B151836, CInt(&HDF3A), CInt(&H4E0A), &H84, &H6C, &HD2, &HAD, &HC9, &H33, &H43, &H33)
IID_ISpeechPhraseInfoBuilder = iid
End Function
Public Function IID_ISpeechRecoResultTimes() As UUID
'{62B3B8FB-F6E7-41be-BDCB-056B1C29EFC0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H62B3B8FB, CInt(&HF6E7), CInt(&H41BE), &HBD, &HCB, &H5, &H6B, &H1C, &H29, &HEF, &HC0)
IID_ISpeechRecoResultTimes = iid
End Function
Public Function IID_ISpeechPhraseAlternate() As UUID
'{27864A2A-2B9F-4cb8-92D3-0D2722FD1E73}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H27864A2A, CInt(&H2B9F), CInt(&H4CB8), &H92, &HD3, &HD, &H27, &H22, &HFD, &H1E, &H73)
IID_ISpeechPhraseAlternate = iid
End Function
Public Function IID_ISpeechPhraseAlternates() As UUID
'{B238B6D5-F276-4c3d-A6C1-2974801C3CC2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB238B6D5, CInt(&HF276), CInt(&H4C3D), &HA6, &HC1, &H29, &H74, &H80, &H1C, &H3C, &HC2)
IID_ISpeechPhraseAlternates = iid
End Function
Public Function IID_ISpeechPhraseInfo() As UUID
'{961559CF-4E67-4662-8BF0-D93F1FCD61B3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H961559CF, CInt(&H4E67), CInt(&H4662), &H8B, &HF0, &HD9, &H3F, &H1F, &HCD, &H61, &HB3)
IID_ISpeechPhraseInfo = iid
End Function
Public Function IID_ISpeechPhraseElement() As UUID
'{E6176F96-E373-4801-B223-3B62C068C0B4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE6176F96, CInt(&HE373), CInt(&H4801), &HB2, &H23, &H3B, &H62, &HC0, &H68, &HC0, &HB4)
IID_ISpeechPhraseElement = iid
End Function
Public Function IID_ISpeechPhraseElements() As UUID
'{0626B328-3478-467d-A0B3-D0853B93DDA3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H626B328, CInt(&H3478), CInt(&H467D), &HA0, &HB3, &HD0, &H85, &H3B, &H93, &HDD, &HA3)
IID_ISpeechPhraseElements = iid
End Function
Public Function IID_ISpeechPhraseReplacement() As UUID
'{2890A410-53A7-4fb5-94EC-06D4998E3D02}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2890A410, CInt(&H53A7), CInt(&H4FB5), &H94, &HEC, &H6, &HD4, &H99, &H8E, &H3D, &H2)
IID_ISpeechPhraseReplacement = iid
End Function
Public Function IID_ISpeechPhraseReplacements() As UUID
'{38BC662F-2257-4525-959E-2069D2596C05}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H38BC662F, CInt(&H2257), CInt(&H4525), &H95, &H9E, &H20, &H69, &HD2, &H59, &H6C, &H5)
IID_ISpeechPhraseReplacements = iid
End Function
Public Function IID_ISpeechPhraseProperty() As UUID
'{CE563D48-961E-4732-A2E1-378A42B430BE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCE563D48, CInt(&H961E), CInt(&H4732), &HA2, &HE1, &H37, &H8A, &H42, &HB4, &H30, &HBE)
IID_ISpeechPhraseProperty = iid
End Function
Public Function IID_ISpeechPhraseProperties() As UUID
'{08166B47-102E-4b23-A599-BDB98DBFD1F4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8166B47, CInt(&H102E), CInt(&H4B23), &HA5, &H99, &HBD, &HB9, &H8D, &HBF, &HD1, &HF4)
IID_ISpeechPhraseProperties = iid
End Function
Public Function IID_ISpeechPhraseRule() As UUID
'{A7BFE112-A4A0-48d9-B602-C313843F6964}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA7BFE112, CInt(&HA4A0), CInt(&H48D9), &HB6, &H2, &HC3, &H13, &H84, &H3F, &H69, &H64)
IID_ISpeechPhraseRule = iid
End Function
Public Function IID_ISpeechPhraseRules() As UUID
'{9047D593-01DD-4b72-81A3-E4A0CA69F407}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9047D593, CInt(&H1DD), CInt(&H4B72), &H81, &HA3, &HE4, &HA0, &HCA, &H69, &HF4, &H7)
IID_ISpeechPhraseRules = iid
End Function
Public Function IID_ISpeechLexicon() As UUID
'{3DA7627A-C7AE-4b23-8708-638C50362C25}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3DA7627A, CInt(&HC7AE), CInt(&H4B23), &H87, &H8, &H63, &H8C, &H50, &H36, &H2C, &H25)
IID_ISpeechLexicon = iid
End Function
Public Function IID_ISpeechLexiconWords() As UUID
'{8D199862-415E-47d5-AC4F-FAA608B424E6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8D199862, CInt(&H415E), CInt(&H47D5), &HAC, &H4F, &HFA, &HA6, &H8, &HB4, &H24, &HE6)
IID_ISpeechLexiconWords = iid
End Function
Public Function IID_ISpeechLexiconWord() As UUID
'{4E5B933C-C9BE-48ed-8842-1EE51BB1D4FF}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4E5B933C, CInt(&HC9BE), CInt(&H48ED), &H88, &H42, &H1E, &HE5, &H1B, &HB1, &HD4, &HFF)
IID_ISpeechLexiconWord = iid
End Function
Public Function IID_ISpeechLexiconPronunciations() As UUID
'{72829128-5682-4704-A0D4-3E2BB6F2EAD3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H72829128, CInt(&H5682), CInt(&H4704), &HA0, &HD4, &H3E, &H2B, &HB6, &HF2, &HEA, &HD3)
IID_ISpeechLexiconPronunciations = iid
End Function
Public Function IID_ISpeechLexiconPronunciation() As UUID
'{95252C5D-9E43-4f4a-9899-48EE73352F9F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H95252C5D, CInt(&H9E43), CInt(&H4F4A), &H98, &H99, &H48, &HEE, &H73, &H35, &H2F, &H9F)
IID_ISpeechLexiconPronunciation = iid
End Function
Public Function IID_ISpeechPhoneConverter() As UUID
'{C3E4F353-433F-43d6-89A1-6A62A7054C3D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC3E4F353, CInt(&H433F), CInt(&H43D6), &H89, &HA1, &H6A, &H62, &HA7, &H5, &H4C, &H3D)
IID_ISpeechPhoneConverter = iid
End Function
Public Function IID_ISpNotifySource() As UUID
'{5EFF4AEF-8487-11D2-961C-00C04F8EE628}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5EFF4AEF, CInt(&H8487), CInt(&H11D2), &H96, &H1C, &H0, &HC0, &H4F, &H8E, &HE6, &H28)
IID_ISpNotifySource = iid
End Function
Public Function IID_ISpNotifySink() As UUID
'{259684DC-37C3-11D2-9603-00C04F8EE628}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H259684DC, CInt(&H37C3), CInt(&H11D2), &H96, &H3, &H0, &HC0, &H4F, &H8E, &HE6, &H28)
IID_ISpNotifySink = iid
End Function
Public Function IID_ISpNotifyTranslator() As UUID
'{ACA16614-5D3D-11D2-960E-00C04F8EE628}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HACA16614, CInt(&H5D3D), CInt(&H11D2), &H96, &HE, &H0, &HC0, &H4F, &H8E, &HE6, &H28)
IID_ISpNotifyTranslator = iid
End Function
Public Function IID_ISpDataKey() As UUID
'{14056581-E16C-11D2-BB90-00C04F8EE6C0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H14056581, CInt(&HE16C), CInt(&H11D2), &HBB, &H90, &H0, &HC0, &H4F, &H8E, &HE6, &HC0)
IID_ISpDataKey = iid
End Function
Public Function IID_ISpRegDataKey() As UUID
'{92A66E2B-C830-4149-83DF-6FC2BA1E7A5B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H92A66E2B, CInt(&HC830), CInt(&H4149), &H83, &HDF, &H6F, &HC2, &HBA, &H1E, &H7A, &H5B)
IID_ISpRegDataKey = iid
End Function
Public Function IID_ISpObjectTokenCategory() As UUID
'{2D3D3845-39AF-4850-BBF9-40B49780011D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2D3D3845, CInt(&H39AF), CInt(&H4850), &HBB, &HF9, &H40, &HB4, &H97, &H80, &H1, &H1D)
IID_ISpObjectTokenCategory = iid
End Function
Public Function IID_ISpObjectToken() As UUID
'{14056589-E16C-11D2-BB90-00C04F8EE6C0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H14056589, CInt(&HE16C), CInt(&H11D2), &HBB, &H90, &H0, &HC0, &H4F, &H8E, &HE6, &HC0)
IID_ISpObjectToken = iid
End Function
Public Function IID_ISpObjectTokenInit() As UUID
'{B8AAB0CF-346F-49D8-9499-C8B03F161D51}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB8AAB0CF, CInt(&H346F), CInt(&H49D8), &H94, &H99, &HC8, &HB0, &H3F, &H16, &H1D, &H51)
IID_ISpObjectTokenInit = iid
End Function
Public Function IID_IEnumSpObjectTokens() As UUID
'{06B64F9E-7FDA-11D2-B4F2-00C04F797396}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6B64F9E, CInt(&H7FDA), CInt(&H11D2), &HB4, &HF2, &H0, &HC0, &H4F, &H79, &H73, &H96)
IID_IEnumSpObjectTokens = iid
End Function
Public Function IID_ISpObjectWithToken() As UUID
'{5B559F40-E952-11D2-BB91-00C04F8EE6C0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5B559F40, CInt(&HE952), CInt(&H11D2), &HBB, &H91, &H0, &HC0, &H4F, &H8E, &HE6, &HC0)
IID_ISpObjectWithToken = iid
End Function
Public Function IID_ISpResourceManager() As UUID
'{93384E18-5014-43D5-ADBB-A78E055926BD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H93384E18, CInt(&H5014), CInt(&H43D5), &HAD, &HBB, &HA7, &H8E, &H5, &H59, &H26, &HBD)
IID_ISpResourceManager = iid
End Function
Public Function IID_ISpEventSource() As UUID
'{BE7A9CCE-5F9E-11D2-960F-00C04F8EE628}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBE7A9CCE, CInt(&H5F9E), CInt(&H11D2), &H96, &HF, &H0, &HC0, &H4F, &H8E, &HE6, &H28)
IID_ISpEventSource = iid
End Function
Public Function IID_ISpEventSource2() As UUID
'{2373A435-6A4B-429e-A6AC-D4231A61975B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2373A435, CInt(&H6A4B), CInt(&H429E), &HA6, &HAC, &HD4, &H23, &H1A, &H61, &H97, &H5B)
IID_ISpEventSource2 = iid
End Function
Public Function IID_ISpEventSink() As UUID
'{BE7A9CC9-5F9E-11D2-960F-00C04F8EE628}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBE7A9CC9, CInt(&H5F9E), CInt(&H11D2), &H96, &HF, &H0, &HC0, &H4F, &H8E, &HE6, &H28)
IID_ISpEventSink = iid
End Function
Public Function IID_ISpStreamFormat() As UUID
'{BED530BE-2606-4F4D-A1C0-54C5CDA5566F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBED530BE, CInt(&H2606), CInt(&H4F4D), &HA1, &HC0, &H54, &HC5, &HCD, &HA5, &H56, &H6F)
IID_ISpStreamFormat = iid
End Function
Public Function IID_ISpStream() As UUID
'{12E3CCA9-7518-44C5-A5E7-BA5A79CB929E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H12E3CCA9, CInt(&H7518), CInt(&H44C5), &HA5, &HE7, &HBA, &H5A, &H79, &HCB, &H92, &H9E)
IID_ISpStream = iid
End Function
Public Function IID_ISpStreamFormatConverter() As UUID
'{678A932C-EA71-4446-9B41-78FDA6280A29}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H678A932C, CInt(&HEA71), CInt(&H4446), &H9B, &H41, &H78, &HFD, &HA6, &H28, &HA, &H29)
IID_ISpStreamFormatConverter = iid
End Function
Public Function IID_ISpAudio() As UUID
'{C05C768F-FAE8-4EC2-8E07-338321C12452}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC05C768F, CInt(&HFAE8), CInt(&H4EC2), &H8E, &H7, &H33, &H83, &H21, &HC1, &H24, &H52)
IID_ISpAudio = iid
End Function
Public Function IID_ISpMMSysAudio() As UUID
'{15806F6E-1D70-4B48-98E6-3B1A007509AB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H15806F6E, CInt(&H1D70), CInt(&H4B48), &H98, &HE6, &H3B, &H1A, &H0, &H75, &H9, &HAB)
IID_ISpMMSysAudio = iid
End Function
Public Function IID_ISpTranscript() As UUID
'{10F63BCE-201A-11D3-AC70-00C04F8EE6C0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10F63BCE, CInt(&H201A), CInt(&H11D3), &HAC, &H70, &H0, &HC0, &H4F, &H8E, &HE6, &HC0)
IID_ISpTranscript = iid
End Function
Public Function IID_ISpLexicon() As UUID
'{DA41A7C2-5383-4DB2-916B-6C1719E3DB58}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDA41A7C2, CInt(&H5383), CInt(&H4DB2), &H91, &H6B, &H6C, &H17, &H19, &HE3, &HDB, &H58)
IID_ISpLexicon = iid
End Function
Public Function IID_ISpContainerLexicon() As UUID
'{8565572F-C094-41CC-B56E-10BD9C3FF044}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8565572F, CInt(&HC094), CInt(&H41CC), &HB5, &H6E, &H10, &HBD, &H9C, &H3F, &HF0, &H44)
IID_ISpContainerLexicon = iid
End Function
Public Function IID_ISpShortcut() As UUID
'{3DF681E2-EA56-11D9-8BDE-F66BAD1E3F3A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3DF681E2, CInt(&HEA56), CInt(&H11D9), &H8B, &HDE, &HF6, &H6B, &HAD, &H1E, &H3F, &H3A)
IID_ISpShortcut = iid
End Function
Public Function IID_ISpPhoneConverter() As UUID
'{8445C581-0CAC-4A38-ABFE-9B2CE2826455}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8445C581, CInt(&HCAC), CInt(&H4A38), &HAB, &HFE, &H9B, &H2C, &HE2, &H82, &H64, &H55)
IID_ISpPhoneConverter = iid
End Function
Public Function IID_ISpPhoneticAlphabetConverter() As UUID
'{133ADCD4-19B4-4020-9FDC-842E78253B17}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H133ADCD4, CInt(&H19B4), CInt(&H4020), &H9F, &HDC, &H84, &H2E, &H78, &H25, &H3B, &H17)
IID_ISpPhoneticAlphabetConverter = iid
End Function
Public Function IID_ISpPhoneticAlphabetSelection() As UUID
'{B2745EFD-42CE-48ca-81F1-A96E02538A90}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB2745EFD, CInt(&H42CE), CInt(&H48CA), &H81, &HF1, &HA9, &H6E, &H2, &H53, &H8A, &H90)
IID_ISpPhoneticAlphabetSelection = iid
End Function
Public Function IID_ISpVoice() As UUID
'{6C44DF74-72B9-4992-A1EC-EF996E0422D4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6C44DF74, CInt(&H72B9), CInt(&H4992), &HA1, &HEC, &HEF, &H99, &H6E, &H4, &H22, &HD4)
IID_ISpVoice = iid
End Function
Public Function IID_ISpPhrase() As UUID
'{1A5C0354-B621-4b5a-8791-D306ED379E53}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1A5C0354, CInt(&HB621), CInt(&H4B5A), &H87, &H91, &HD3, &H6, &HED, &H37, &H9E, &H53)
IID_ISpPhrase = iid
End Function
Public Function IID_ISpPhraseAlt() As UUID
'{8FCEBC98-4E49-4067-9C6C-D86A0E092E3D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8FCEBC98, CInt(&H4E49), CInt(&H4067), &H9C, &H6C, &HD8, &H6A, &HE, &H9, &H2E, &H3D)
IID_ISpPhraseAlt = iid
End Function
Public Function IID_ISpPhrase2() As UUID
'{F264DA52-E457-4696-B856-A737B717AF79}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF264DA52, CInt(&HE457), CInt(&H4696), &HB8, &H56, &HA7, &H37, &HB7, &H17, &HAF, &H79)
IID_ISpPhrase2 = iid
End Function
Public Function IID_ISpRecoResult() As UUID
'{20B053BE-E235-43cd-9A2A-8D17A48B7842}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20B053BE, CInt(&HE235), CInt(&H43CD), &H9A, &H2A, &H8D, &H17, &HA4, &H8B, &H78, &H42)
IID_ISpRecoResult = iid
End Function
Public Function IID_ISpRecoResult2() As UUID
'{27CAC6C4-88F2-41f2-8817-0C95E59F1E6E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H27CAC6C4, CInt(&H88F2), CInt(&H41F2), &H88, &H17, &HC, &H95, &HE5, &H9F, &H1E, &H6E)
IID_ISpRecoResult2 = iid
End Function
Public Function IID_ISpXMLRecoResult() As UUID
'{AE39362B-45A8-4074-9B9E-CCF49AA2D0B6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAE39362B, CInt(&H45A8), CInt(&H4074), &H9B, &H9E, &HCC, &HF4, &H9A, &HA2, &HD0, &HB6)
IID_ISpXMLRecoResult = iid
End Function
Public Function IID_ISpGrammarBuilder() As UUID
'{8137828F-591A-4A42-BE58-49EA7EBAAC68}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8137828F, CInt(&H591A), CInt(&H4A42), &HBE, &H58, &H49, &HEA, &H7E, &HBA, &HAC, &H68)
IID_ISpGrammarBuilder = iid
End Function
Public Function IID_ISpRecoGrammar() As UUID
'{2177DB29-7F45-47D0-8554-067E91C80502}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2177DB29, CInt(&H7F45), CInt(&H47D0), &H85, &H54, &H6, &H7E, &H91, &HC8, &H5, &H2)
IID_ISpRecoGrammar = iid
End Function
Public Function IID_ISpGrammarBuilder2() As UUID
'{8AB10026-20CC-4b20-8C22-A49C9BA78F60}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8AB10026, CInt(&H20CC), CInt(&H4B20), &H8C, &H22, &HA4, &H9C, &H9B, &HA7, &H8F, &H60)
IID_ISpGrammarBuilder2 = iid
End Function
Public Function IID_ISpRecoGrammar2() As UUID
'{4B37BC9E-9ED6-44a3-93D3-18F022B79EC3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4B37BC9E, CInt(&H9ED6), CInt(&H44A3), &H93, &HD3, &H18, &HF0, &H22, &HB7, &H9E, &HC3)
IID_ISpRecoGrammar2 = iid
End Function
Public Function IID_ISpeechResourceLoader() As UUID
'{B9AC5783-FCD0-4b21-B119-B4F8DA8FD2C3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9AC5783, CInt(&HFCD0), CInt(&H4B21), &HB1, &H19, &HB4, &HF8, &HDA, &H8F, &HD2, &HC3)
IID_ISpeechResourceLoader = iid
End Function
Public Function IID_ISpRecoContext() As UUID
'{F740A62F-7C15-489E-8234-940A33D9272D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF740A62F, CInt(&H7C15), CInt(&H489E), &H82, &H34, &H94, &HA, &H33, &HD9, &H27, &H2D)
IID_ISpRecoContext = iid
End Function
Public Function IID_ISpRecoContext2() As UUID
'{BEAD311C-52FF-437f-9464-6B21054CA73D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBEAD311C, CInt(&H52FF), CInt(&H437F), &H94, &H64, &H6B, &H21, &H5, &H4C, &HA7, &H3D)
IID_ISpRecoContext2 = iid
End Function
Public Function IID_ISpProperties() As UUID
'{5B4FB971-B115-4DE1-AD97-E482E3BF6EE4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5B4FB971, CInt(&HB115), CInt(&H4DE1), &HAD, &H97, &HE4, &H82, &HE3, &HBF, &H6E, &HE4)
IID_ISpProperties = iid
End Function
Public Function IID_ISpRecognizer() As UUID
'{C2B5F241-DAA0-4507-9E16-5A1EAA2B7A5C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC2B5F241, CInt(&HDAA0), CInt(&H4507), &H9E, &H16, &H5A, &H1E, &HAA, &H2B, &H7A, &H5C)
IID_ISpRecognizer = iid
End Function
Public Function IID_ISpSerializeState() As UUID
'{21B501A0-0EC7-46c9-92C3-A2BC784C54B9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H21B501A0, CInt(&HEC7), CInt(&H46C9), &H92, &HC3, &HA2, &HBC, &H78, &H4C, &H54, &HB9)
IID_ISpSerializeState = iid
End Function
Public Function IID_ISpRecognizer2() As UUID
'{8FC6D974-C81E-4098-93C5-0147F61ED4D3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8FC6D974, CInt(&HC81E), CInt(&H4098), &H93, &HC5, &H1, &H47, &HF6, &H1E, &HD4, &HD3)
IID_ISpRecognizer2 = iid
End Function
Public Function IID_ISpRecoCategory() As UUID
'{DA0CD0F9-14A2-4f09-8C2A-85CC48979345}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDA0CD0F9, CInt(&H14A2), CInt(&H4F09), &H8C, &H2A, &H85, &HCC, &H48, &H97, &H93, &H45)
IID_ISpRecoCategory = iid
End Function
Public Function IID_ISpRecognizer3() As UUID
'{DF1B943C-5838-4AA2-8706-D7CD5B333499}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDF1B943C, CInt(&H5838), CInt(&H4AA2), &H87, &H6, &HD7, &HCD, &H5B, &H33, &H34, &H99)
IID_ISpRecognizer3 = iid
End Function
Public Function IID_ISpEnginePronunciation() As UUID
'{C360CE4B-76D1-4214-AD68-52657D5083DA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC360CE4B, CInt(&H76D1), CInt(&H4214), &HAD, &H68, &H52, &H65, &H7D, &H50, &H83, &HDA)
IID_ISpEnginePronunciation = iid
End Function
Public Function IID_ISpDisplayAlternates() As UUID
'{C8D7C7E2-0DDE-44b7-AFE3-B0C991FBEB5E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC8D7C7E2, CInt(&HDDE), CInt(&H44B7), &HAF, &HE3, &HB0, &HC9, &H91, &HFB, &HEB, &H5E)
IID_ISpDisplayAlternates = iid
End Function

