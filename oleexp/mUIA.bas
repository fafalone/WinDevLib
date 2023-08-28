Attribute VB_Name = "mUIA"
Option Explicit

'GUIDs for UIAutomation

Public Const UIA_ScrollPatternNoScroll As Double = -1 'For some reason, if you put double consts in a typelib, VB6 crashes when the object browser displays it

Public Function IID_IRawElementProviderSimple() As UUID
'{d6dd68d1-86fd-4332-8666-9abedea2d24c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD6DD68D1, CInt(&H86FD), CInt(&H4332), &H86, &H66, &H9A, &HBE, &HDE, &HA2, &HD2, &H4C)
IID_IRawElementProviderSimple = iid
End Function
Public Function IID_IAccessibleEx() As UUID
'{f8b80ada-2c44-48d0-89be-5ff23c9cd875}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF8B80ADA, CInt(&H2C44), CInt(&H48D0), &H89, &HBE, &H5F, &HF2, &H3C, &H9C, &HD8, &H75)
IID_IAccessibleEx = iid
End Function
Public Function IID_IRawElementProviderSimple2() As UUID
'{A0A839A9-8DA1-4A82-806A-8E0D44E79F56}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA0A839A9, CInt(&H8DA1), CInt(&H4A82), &H80, &H6A, &H8E, &HD, &H44, &HE7, &H9F, &H56)
IID_IRawElementProviderSimple2 = iid
End Function
Public Function IID_IRawElementProviderSimple3() As UUID
'{fcf5d820-d7ec-4613-bdf6-42a84ce7daaf}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFCF5D820, CInt(&HD7EC), CInt(&H4613), &HBD, &HF6, &H42, &HA8, &H4C, &HE7, &HDA, &HAF)
IID_IRawElementProviderSimple3 = iid
End Function
Public Function IID_IRawElementProviderFragment() As UUID
'{f7063da8-8359-439c-9297-bbc5299a7d87}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF7063DA8, CInt(&H8359), CInt(&H439C), &H92, &H97, &HBB, &HC5, &H29, &H9A, &H7D, &H87)
IID_IRawElementProviderFragment = iid
End Function
Public Function IID_IRawElementProviderFragmentRoot() As UUID
'{620ce2a5-ab8f-40a9-86cb-de3c75599b58}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H620CE2A5, CInt(&HAB8F), CInt(&H40A9), &H86, &HCB, &HDE, &H3C, &H75, &H59, &H9B, &H58)
IID_IRawElementProviderFragmentRoot = iid
End Function
Public Function IID_IRawElementProviderAdviseEvents() As UUID
'{a407b27b-0f6d-4427-9292-473c7bf93258}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA407B27B, CInt(&HF6D), CInt(&H4427), &H92, &H92, &H47, &H3C, &H7B, &HF9, &H32, &H58)
IID_IRawElementProviderAdviseEvents = iid
End Function
Public Function IID_IRawElementProviderHwndOverride() As UUID
'{1d5df27c-8947-4425-b8d9-79787bb460b8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1D5DF27C, CInt(&H8947), CInt(&H4425), &HB8, &HD9, &H79, &H78, &H7B, &HB4, &H60, &HB8)
IID_IRawElementProviderHwndOverride = iid
End Function
Public Function IID_IProxyProviderWinEventSink() As UUID
'{4fd82b78-a43e-46ac-9803-0a6969c7c183}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4FD82B78, CInt(&HA43E), CInt(&H46AC), &H98, &H3, &HA, &H69, &H69, &HC7, &HC1, &H83)
IID_IProxyProviderWinEventSink = iid
End Function
Public Function IID_IProxyProviderWinEventHandler() As UUID
'{89592ad4-f4e0-43d5-a3b6-bad7e111b435}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H89592AD4, CInt(&HF4E0), CInt(&H43D5), &HA3, &HB6, &HBA, &HD7, &HE1, &H11, &HB4, &H35)
IID_IProxyProviderWinEventHandler = iid
End Function
Public Function IID_IRawElementProviderWindowlessSite() As UUID
'{0a2a93cc-bfad-42ac-9b2e-0991fb0d3ea0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA2A93CC, CInt(&HBFAD), CInt(&H42AC), &H9B, &H2E, &H9, &H91, &HFB, &HD, &H3E, &HA0)
IID_IRawElementProviderWindowlessSite = iid
End Function
Public Function IID_IAccessibleHostingElementProviders() As UUID
'{33AC331B-943E-4020-B295-DB37784974A3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H33AC331B, CInt(&H943E), CInt(&H4020), &HB2, &H95, &HDB, &H37, &H78, &H49, &H74, &HA3)
IID_IAccessibleHostingElementProviders = iid
End Function
Public Function IID_IRawElementProviderHostingAccessibles() As UUID
'{24BE0B07-D37D-487A-98CF-A13ED465E9B3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H24BE0B07, CInt(&HD37D), CInt(&H487A), &H98, &HCF, &HA1, &H3E, &HD4, &H65, &HE9, &HB3)
IID_IRawElementProviderHostingAccessibles = iid
End Function
Public Function IID_IDockProvider() As UUID
'{159bc72c-4ad3-485e-9637-d7052edf0146}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H159BC72C, CInt(&H4AD3), CInt(&H485E), &H96, &H37, &HD7, &H5, &H2E, &HDF, &H1, &H46)
IID_IDockProvider = iid
End Function
Public Function IID_IExpandCollapseProvider() As UUID
'{d847d3a5-cab0-4a98-8c32-ecb45c59ad24}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD847D3A5, CInt(&HCAB0), CInt(&H4A98), &H8C, &H32, &HEC, &HB4, &H5C, &H59, &HAD, &H24)
IID_IExpandCollapseProvider = iid
End Function
Public Function IID_IGridProvider() As UUID
'{b17d6187-0907-464b-a168-0ef17a1572b1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB17D6187, CInt(&H907), CInt(&H464B), &HA1, &H68, &HE, &HF1, &H7A, &H15, &H72, &HB1)
IID_IGridProvider = iid
End Function
Public Function IID_IGridItemProvider() As UUID
'{d02541f1-fb81-4d64-ae32-f520f8a6dbd1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD02541F1, CInt(&HFB81), CInt(&H4D64), &HAE, &H32, &HF5, &H20, &HF8, &HA6, &HDB, &HD1)
IID_IGridItemProvider = iid
End Function
Public Function IID_IInvokeProvider() As UUID
'{54fcb24b-e18e-47a2-b4d3-eccbe77599a2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H54FCB24B, CInt(&HE18E), CInt(&H47A2), &HB4, &HD3, &HEC, &HCB, &HE7, &H75, &H99, &HA2)
IID_IInvokeProvider = iid
End Function
Public Function IID_IMultipleViewProvider() As UUID
'{6278cab1-b556-4a1a-b4e0-418acc523201}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6278CAB1, CInt(&HB556), CInt(&H4A1A), &HB4, &HE0, &H41, &H8A, &HCC, &H52, &H32, &H1)
IID_IMultipleViewProvider = iid
End Function
Public Function IID_IRangeValueProvider() As UUID
'{36dc7aef-33e6-4691-afe1-2be7274b3d33}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H36DC7AEF, CInt(&H33E6), CInt(&H4691), &HAF, &HE1, &H2B, &HE7, &H27, &H4B, &H3D, &H33)
IID_IRangeValueProvider = iid
End Function
Public Function IID_IScrollItemProvider() As UUID
'{2360c714-4bf1-4b26-ba65-9b21316127eb}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2360C714, CInt(&H4BF1), CInt(&H4B26), &HBA, &H65, &H9B, &H21, &H31, &H61, &H27, &HEB)
IID_IScrollItemProvider = iid
End Function
Public Function IID_ISelectionProvider() As UUID
'{fb8b03af-3bdf-48d4-bd36-1a65793be168}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFB8B03AF, CInt(&H3BDF), CInt(&H48D4), &HBD, &H36, &H1A, &H65, &H79, &H3B, &HE1, &H68)
IID_ISelectionProvider = iid
End Function
Public Function IID_ISelectionProvider2() As UUID
'{14f68475-ee1c-44f6-a869-d239381f0fe7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H14F68475, CInt(&HEE1C), CInt(&H44F6), &HA8, &H69, &HD2, &H39, &H38, &H1F, &HF, &HE7)
IID_ISelectionProvider2 = iid
End Function
Public Function IID_IScrollProvider() As UUID
'{b38b8077-1fc3-42a5-8cae-d40c2215055a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB38B8077, CInt(&H1FC3), CInt(&H42A5), &H8C, &HAE, &HD4, &HC, &H22, &H15, &H5, &H5A)
IID_IScrollProvider = iid
End Function
Public Function IID_ISelectionItemProvider() As UUID
'{2acad808-b2d4-452d-a407-91ff1ad167b2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2ACAD808, CInt(&HB2D4), CInt(&H452D), &HA4, &H7, &H91, &HFF, &H1A, &HD1, &H67, &HB2)
IID_ISelectionItemProvider = iid
End Function
Public Function IID_ISynchronizedInputProvider() As UUID
'{29db1a06-02ce-4cf7-9b42-565d4fab20ee}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H29DB1A06, CInt(&H2CE), CInt(&H4CF7), &H9B, &H42, &H56, &H5D, &H4F, &HAB, &H20, &HEE)
IID_ISynchronizedInputProvider = iid
End Function
Public Function IID_ITableProvider() As UUID
'{9c860395-97b3-490a-b52a-858cc22af166}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9C860395, CInt(&H97B3), CInt(&H490A), &HB5, &H2A, &H85, &H8C, &HC2, &H2A, &HF1, &H66)
IID_ITableProvider = iid
End Function
Public Function IID_ITableItemProvider() As UUID
'{b9734fa6-771f-4d78-9c90-2517999349cd}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB9734FA6, CInt(&H771F), CInt(&H4D78), &H9C, &H90, &H25, &H17, &H99, &H93, &H49, &HCD)
IID_ITableItemProvider = iid
End Function
Public Function IID_IToggleProvider() As UUID
'{56d00bd0-c4f4-433c-a836-1a52a57e0892}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56D00BD0, CInt(&HC4F4), CInt(&H433C), &HA8, &H36, &H1A, &H52, &HA5, &H7E, &H8, &H92)
IID_IToggleProvider = iid
End Function
Public Function IID_ITransformProvider() As UUID
'{6829ddc4-4f91-4ffa-b86f-bd3e2987cb4c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6829DDC4, CInt(&H4F91), CInt(&H4FFA), &HB8, &H6F, &HBD, &H3E, &H29, &H87, &HCB, &H4C)
IID_ITransformProvider = iid
End Function
Public Function IID_IValueProvider() As UUID
'{c7935180-6fb3-4201-b174-7df73adbf64a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC7935180, CInt(&H6FB3), CInt(&H4201), &HB1, &H74, &H7D, &HF7, &H3A, &HDB, &HF6, &H4A)
IID_IValueProvider = iid
End Function
Public Function IID_IWindowProvider() As UUID
'{987df77b-db06-4d77-8f8a-86a9c3bb90b9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H987DF77B, CInt(&HDB06), CInt(&H4D77), &H8F, &H8A, &H86, &HA9, &HC3, &HBB, &H90, &HB9)
IID_IWindowProvider = iid
End Function
Public Function IID_ILegacyIAccessibleProvider() As UUID
'{e44c3566-915d-4070-99c6-047bff5a08f5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE44C3566, CInt(&H915D), CInt(&H4070), &H99, &HC6, &H4, &H7B, &HFF, &H5A, &H8, &HF5)
IID_ILegacyIAccessibleProvider = iid
End Function
Public Function IID_IItemContainerProvider() As UUID
'{e747770b-39ce-4382-ab30-d8fb3f336f24}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE747770B, CInt(&H39CE), CInt(&H4382), &HAB, &H30, &HD8, &HFB, &H3F, &H33, &H6F, &H24)
IID_IItemContainerProvider = iid
End Function
Public Function IID_IVirtualizedItemProvider() As UUID
'{cb98b665-2d35-4fac-ad35-f3c60d0c0b8b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCB98B665, CInt(&H2D35), CInt(&H4FAC), &HAD, &H35, &HF3, &HC6, &HD, &HC, &HB, &H8B)
IID_IVirtualizedItemProvider = iid
End Function
Public Function IID_IObjectModelProvider() As UUID
'{3ad86ebd-f5ef-483d-bb18-b1042a475d64}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3AD86EBD, CInt(&HF5EF), CInt(&H483D), &HBB, &H18, &HB1, &H4, &H2A, &H47, &H5D, &H64)
IID_IObjectModelProvider = iid
End Function
Public Function IID_IAnnotationProvider() As UUID
'{f95c7e80-bd63-4601-9782-445ebff011fc}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF95C7E80, CInt(&HBD63), CInt(&H4601), &H97, &H82, &H44, &H5E, &HBF, &HF0, &H11, &HFC)
IID_IAnnotationProvider = iid
End Function
Public Function IID_IStylesProvider() As UUID
'{19b6b649-f5d7-4a6d-bdcb-129252be588a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H19B6B649, CInt(&HF5D7), CInt(&H4A6D), &HBD, &HCB, &H12, &H92, &H52, &HBE, &H58, &H8A)
IID_IStylesProvider = iid
End Function
Public Function IID_ISpreadsheetProvider() As UUID
'{6f6b5d35-5525-4f80-b758-85473832ffc7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6F6B5D35, CInt(&H5525), CInt(&H4F80), &HB7, &H58, &H85, &H47, &H38, &H32, &HFF, &HC7)
IID_ISpreadsheetProvider = iid
End Function
Public Function IID_ISpreadsheetItemProvider() As UUID
'{eaed4660-7b3d-4879-a2e6-365ce603f3d0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEAED4660, CInt(&H7B3D), CInt(&H4879), &HA2, &HE6, &H36, &H5C, &HE6, &H3, &HF3, &HD0)
IID_ISpreadsheetItemProvider = iid
End Function
Public Function IID_ITransformProvider2() As UUID
'{4758742f-7ac2-460c-bc48-09fc09308a93}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4758742F, CInt(&H7AC2), CInt(&H460C), &HBC, &H48, &H9, &HFC, &H9, &H30, &H8A, &H93)
IID_ITransformProvider2 = iid
End Function
Public Function IID_IDragProvider() As UUID
'{6aa7bbbb-7ff9-497d-904f-d20b897929d8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6AA7BBBB, CInt(&H7FF9), CInt(&H497D), &H90, &H4F, &HD2, &HB, &H89, &H79, &H29, &HD8)
IID_IDragProvider = iid
End Function
Public Function IID_IDropTargetProvider() As UUID
'{bae82bfd-358a-481c-85a0-d8b4d90a5d61}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBAE82BFD, CInt(&H358A), CInt(&H481C), &H85, &HA0, &HD8, &HB4, &HD9, &HA, &H5D, &H61)
IID_IDropTargetProvider = iid
End Function
Public Function IID_ITextProvider() As UUID
'{3589c92c-63f3-4367-99bb-ada653b77cf2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3589C92C, CInt(&H63F3), CInt(&H4367), &H99, &HBB, &HAD, &HA6, &H53, &HB7, &H7C, &HF2)
IID_ITextProvider = iid
End Function
Public Function IID_ITextProvider2() As UUID
'{0dc5e6ed-3e16-4bf1-8f9a-a979878bc195}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDC5E6ED, CInt(&H3E16), CInt(&H4BF1), &H8F, &H9A, &HA9, &H79, &H87, &H8B, &HC1, &H95)
IID_ITextProvider2 = iid
End Function
Public Function IID_ITextEditProvider() As UUID
'{EA3605B4-3A05-400E-B5F9-4E91B40F6176}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEA3605B4, CInt(&H3A05), CInt(&H400E), &HB5, &HF9, &H4E, &H91, &HB4, &HF, &H61, &H76)
IID_ITextEditProvider = iid
End Function
Public Function IID_ITextRangeProvider() As UUID
'{5347ad7b-c355-46f8-aff5-909033582f63}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5347AD7B, CInt(&HC355), CInt(&H46F8), &HAF, &HF5, &H90, &H90, &H33, &H58, &H2F, &H63)
IID_ITextRangeProvider = iid
End Function
Public Function IID_ITextRangeProvider2() As UUID
'{9BBCE42C-1921-4F18-89CA-DBA1910A0386}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9BBCE42C, CInt(&H1921), CInt(&H4F18), &H89, &HCA, &HDB, &HA1, &H91, &HA, &H3, &H86)
IID_ITextRangeProvider2 = iid
End Function
Public Function IID_ITextChildProvider() As UUID
'{4c2de2b9-c88f-4f88-a111-f1d336b7d1a9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4C2DE2B9, CInt(&HC88F), CInt(&H4F88), &HA1, &H11, &HF1, &HD3, &H36, &HB7, &HD1, &HA9)
IID_ITextChildProvider = iid
End Function
Public Function IID_ICustomNavigationProvider() As UUID
'{2062A28A-8C07-4B94-8E12-7037C622AEB8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2062A28A, CInt(&H8C07), CInt(&H4B94), &H8E, &H12, &H70, &H37, &HC6, &H22, &HAE, &HB8)
IID_ICustomNavigationProvider = iid
End Function
Public Function IID_IUIAutomationPatternInstance() As UUID
'{c03a7fe4-9431-409f-bed8-ae7c2299bc8d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC03A7FE4, CInt(&H9431), CInt(&H409F), &HBE, &HD8, &HAE, &H7C, &H22, &H99, &HBC, &H8D)
IID_IUIAutomationPatternInstance = iid
End Function
Public Function IID_IUIAutomationPatternHandler() As UUID
'{d97022f3-a947-465e-8b2a-ac4315fa54e8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD97022F3, CInt(&HA947), CInt(&H465E), &H8B, &H2A, &HAC, &H43, &H15, &HFA, &H54, &HE8)
IID_IUIAutomationPatternHandler = iid
End Function
Public Function IID_IUIAutomationRegistrar() As UUID
'{8609c4ec-4a1a-4d88-a357-5a66e060e1cf}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8609C4EC, CInt(&H4A1A), CInt(&H4D88), &HA3, &H57, &H5A, &H66, &HE0, &H60, &HE1, &HCF)
IID_IUIAutomationRegistrar = iid
End Function
Public Function IID_IUIAutomationCondition() As UUID
'{352ffba8-0973-437c-a61f-f64cafd81df9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H352FFBA8, CInt(&H973), CInt(&H437C), &HA6, &H1F, &HF6, &H4C, &HAF, &HD8, &H1D, &HF9)
IID_IUIAutomationCondition = iid
End Function
Public Function IID_IUIAutomationBoolCondition() As UUID
'{1b4e1f2e-75eb-4d0b-8952-5a69988e2307}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1B4E1F2E, CInt(&H75EB), CInt(&H4D0B), &H89, &H52, &H5A, &H69, &H98, &H8E, &H23, &H7)
IID_IUIAutomationBoolCondition = iid
End Function
Public Function IID_IUIAutomationPropertyCondition() As UUID
'{99ebf2cb-5578-4267-9ad4-afd6ea77e94b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H99EBF2CB, CInt(&H5578), CInt(&H4267), &H9A, &HD4, &HAF, &HD6, &HEA, &H77, &HE9, &H4B)
IID_IUIAutomationPropertyCondition = iid
End Function
Public Function IID_IUIAutomationAndCondition() As UUID
'{a7d0af36-b912-45fe-9855-091ddc174aec}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA7D0AF36, CInt(&HB912), CInt(&H45FE), &H98, &H55, &H9, &H1D, &HDC, &H17, &H4A, &HEC)
IID_IUIAutomationAndCondition = iid
End Function
Public Function IID_IUIAutomationOrCondition() As UUID
'{8753f032-3db1-47b5-a1fc-6e34a266c712}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8753F032, CInt(&H3DB1), CInt(&H47B5), &HA1, &HFC, &H6E, &H34, &HA2, &H66, &HC7, &H12)
IID_IUIAutomationOrCondition = iid
End Function
Public Function IID_IUIAutomationNotCondition() As UUID
'{f528b657-847b-498c-8896-d52b565407a1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF528B657, CInt(&H847B), CInt(&H498C), &H88, &H96, &HD5, &H2B, &H56, &H54, &H7, &HA1)
IID_IUIAutomationNotCondition = iid
End Function
Public Function IID_IUIAutomationCacheRequest() As UUID
'{b32a92b5-bc25-4078-9c08-d7ee95c48e03}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB32A92B5, CInt(&HBC25), CInt(&H4078), &H9C, &H8, &HD7, &HEE, &H95, &HC4, &H8E, &H3)
IID_IUIAutomationCacheRequest = iid
End Function
Public Function IID_IUIAutomationTreeWalker() As UUID
'{4042c624-389c-4afc-a630-9df854a541fc}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4042C624, CInt(&H389C), CInt(&H4AFC), &HA6, &H30, &H9D, &HF8, &H54, &HA5, &H41, &HFC)
IID_IUIAutomationTreeWalker = iid
End Function
Public Function IID_IUIAutomationEventHandler() As UUID
'{146c3c17-f12e-4e22-8c27-f894b9b79c69}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H146C3C17, CInt(&HF12E), CInt(&H4E22), &H8C, &H27, &HF8, &H94, &HB9, &HB7, &H9C, &H69)
IID_IUIAutomationEventHandler = iid
End Function
Public Function IID_IUIAutomationPropertyChangedEventHandler() As UUID
'{40cd37d4-c756-4b0c-8c6f-bddfeeb13b50}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H40CD37D4, CInt(&HC756), CInt(&H4B0C), &H8C, &H6F, &HBD, &HDF, &HEE, &HB1, &H3B, &H50)
IID_IUIAutomationPropertyChangedEventHandler = iid
End Function
Public Function IID_IUIAutomationStructureChangedEventHandler() As UUID
'{e81d1b4e-11c5-42f8-9754-e7036c79f054}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE81D1B4E, CInt(&H11C5), CInt(&H42F8), &H97, &H54, &HE7, &H3, &H6C, &H79, &HF0, &H54)
IID_IUIAutomationStructureChangedEventHandler = iid
End Function
Public Function IID_IUIAutomationFocusChangedEventHandler() As UUID
'{c270f6b5-5c69-4290-9745-7a7f97169468}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC270F6B5, CInt(&H5C69), CInt(&H4290), &H97, &H45, &H7A, &H7F, &H97, &H16, &H94, &H68)
IID_IUIAutomationFocusChangedEventHandler = iid
End Function
Public Function IID_IUIAutomationTextEditTextChangedEventHandler() As UUID
'{92FAA680-E704-4156-931A-E32D5BB38F3F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H92FAA680, CInt(&HE704), CInt(&H4156), &H93, &H1A, &HE3, &H2D, &H5B, &HB3, &H8F, &H3F)
IID_IUIAutomationTextEditTextChangedEventHandler = iid
End Function
Public Function IID_IUIAutomationChangesEventHandler() As UUID
'{58EDCA55-2C3E-4980-B1B9-56C17F27A2A0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H58EDCA55, CInt(&H2C3E), CInt(&H4980), &HB1, &HB9, &H56, &HC1, &H7F, &H27, &HA2, &HA0)
IID_IUIAutomationChangesEventHandler = iid
End Function
Public Function IID_IUIAutomationNotificationEventHandler() As UUID
'{C7CB2637-E6C2-4D0C-85DE-4948C02175C7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC7CB2637, CInt(&HE6C2), CInt(&H4D0C), &H85, &HDE, &H49, &H48, &HC0, &H21, &H75, &HC7)
IID_IUIAutomationNotificationEventHandler = iid
End Function
Public Function IID_IUIAutomationInvokePattern() As UUID
'{fb377fbe-8ea6-46d5-9c73-6499642d3059}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFB377FBE, CInt(&H8EA6), CInt(&H46D5), &H9C, &H73, &H64, &H99, &H64, &H2D, &H30, &H59)
IID_IUIAutomationInvokePattern = iid
End Function
Public Function IID_IUIAutomationDockPattern() As UUID
'{fde5ef97-1464-48f6-90bf-43d0948e86ec}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFDE5EF97, CInt(&H1464), CInt(&H48F6), &H90, &HBF, &H43, &HD0, &H94, &H8E, &H86, &HEC)
IID_IUIAutomationDockPattern = iid
End Function
Public Function IID_IUIAutomationExpandCollapsePattern() As UUID
'{619be086-1f4e-4ee4-bafa-210128738730}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H619BE086, CInt(&H1F4E), CInt(&H4EE4), &HBA, &HFA, &H21, &H1, &H28, &H73, &H87, &H30)
IID_IUIAutomationExpandCollapsePattern = iid
End Function
Public Function IID_IUIAutomationGridPattern() As UUID
'{414c3cdc-856b-4f5b-8538-3131c6302550}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H414C3CDC, CInt(&H856B), CInt(&H4F5B), &H85, &H38, &H31, &H31, &HC6, &H30, &H25, &H50)
IID_IUIAutomationGridPattern = iid
End Function
Public Function IID_IUIAutomationGridItemPattern() As UUID
'{78f8ef57-66c3-4e09-bd7c-e79b2004894d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H78F8EF57, CInt(&H66C3), CInt(&H4E09), &HBD, &H7C, &HE7, &H9B, &H20, &H4, &H89, &H4D)
IID_IUIAutomationGridItemPattern = iid
End Function
Public Function IID_IUIAutomationMultipleViewPattern() As UUID
'{8d253c91-1dc5-4bb5-b18f-ade16fa495e8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8D253C91, CInt(&H1DC5), CInt(&H4BB5), &HB1, &H8F, &HAD, &HE1, &H6F, &HA4, &H95, &HE8)
IID_IUIAutomationMultipleViewPattern = iid
End Function
Public Function IID_IUIAutomationObjectModelPattern() As UUID
'{71c284b3-c14d-4d14-981e-19751b0d756d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H71C284B3, CInt(&HC14D), CInt(&H4D14), &H98, &H1E, &H19, &H75, &H1B, &HD, &H75, &H6D)
IID_IUIAutomationObjectModelPattern = iid
End Function
Public Function IID_IUIAutomationRangeValuePattern() As UUID
'{59213f4f-7346-49e5-b120-80555987a148}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H59213F4F, CInt(&H7346), CInt(&H49E5), &HB1, &H20, &H80, &H55, &H59, &H87, &HA1, &H48)
IID_IUIAutomationRangeValuePattern = iid
End Function
Public Function IID_IUIAutomationScrollPattern() As UUID
'{88f4d42a-e881-459d-a77c-73bbbb7e02dc}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H88F4D42A, CInt(&HE881), CInt(&H459D), &HA7, &H7C, &H73, &HBB, &HBB, &H7E, &H2, &HDC)
IID_IUIAutomationScrollPattern = iid
End Function
Public Function IID_IUIAutomationScrollItemPattern() As UUID
'{b488300f-d015-4f19-9c29-bb595e3645ef}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB488300F, CInt(&HD015), CInt(&H4F19), &H9C, &H29, &HBB, &H59, &H5E, &H36, &H45, &HEF)
IID_IUIAutomationScrollItemPattern = iid
End Function
Public Function IID_IUIAutomationSelectionPattern() As UUID
'{5ed5202e-b2ac-47a6-b638-4b0bf140d78e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5ED5202E, CInt(&HB2AC), CInt(&H47A6), &HB6, &H38, &H4B, &HB, &HF1, &H40, &HD7, &H8E)
IID_IUIAutomationSelectionPattern = iid
End Function
Public Function IID_IUIAutomationSelectionPattern2() As UUID
'{0532bfae-c011-4e32-a343-6d642d798555}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H532BFAE, CInt(&HC011), CInt(&H4E32), &HA3, &H43, &H6D, &H64, &H2D, &H79, &H85, &H55)
IID_IUIAutomationSelectionPattern2 = iid
End Function
Public Function IID_IUIAutomationSelectionItemPattern() As UUID
'{a8efa66a-0fda-421a-9194-38021f3578ea}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA8EFA66A, CInt(&HFDA), CInt(&H421A), &H91, &H94, &H38, &H2, &H1F, &H35, &H78, &HEA)
IID_IUIAutomationSelectionItemPattern = iid
End Function
Public Function IID_IUIAutomationSynchronizedInputPattern() As UUID
'{2233be0b-afb7-448b-9fda-3b378aa5eae1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2233BE0B, CInt(&HAFB7), CInt(&H448B), &H9F, &HDA, &H3B, &H37, &H8A, &HA5, &HEA, &HE1)
IID_IUIAutomationSynchronizedInputPattern = iid
End Function
Public Function IID_IUIAutomationTablePattern() As UUID
'{620e691c-ea96-4710-a850-754b24ce2417}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H620E691C, CInt(&HEA96), CInt(&H4710), &HA8, &H50, &H75, &H4B, &H24, &HCE, &H24, &H17)
IID_IUIAutomationTablePattern = iid
End Function
Public Function IID_IUIAutomationTableItemPattern() As UUID
'{0b964eb3-ef2e-4464-9c79-61d61737a27e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB964EB3, CInt(&HEF2E), CInt(&H4464), &H9C, &H79, &H61, &HD6, &H17, &H37, &HA2, &H7E)
IID_IUIAutomationTableItemPattern = iid
End Function
Public Function IID_IUIAutomationTogglePattern() As UUID
'{94cf8058-9b8d-4ab9-8bfd-4cd0a33c8c70}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H94CF8058, CInt(&H9B8D), CInt(&H4AB9), &H8B, &HFD, &H4C, &HD0, &HA3, &H3C, &H8C, &H70)
IID_IUIAutomationTogglePattern = iid
End Function
Public Function IID_IUIAutomationTransformPattern() As UUID
'{a9b55844-a55d-4ef0-926d-569c16ff89bb}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA9B55844, CInt(&HA55D), CInt(&H4EF0), &H92, &H6D, &H56, &H9C, &H16, &HFF, &H89, &HBB)
IID_IUIAutomationTransformPattern = iid
End Function
Public Function IID_IUIAutomationValuePattern() As UUID
'{a94cd8b1-0844-4cd6-9d2d-640537ab39e9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA94CD8B1, CInt(&H844), CInt(&H4CD6), &H9D, &H2D, &H64, &H5, &H37, &HAB, &H39, &HE9)
IID_IUIAutomationValuePattern = iid
End Function
Public Function IID_IUIAutomationWindowPattern() As UUID
'{0faef453-9208-43ef-bbb2-3b485177864f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFAEF453, CInt(&H9208), CInt(&H43EF), &HBB, &HB2, &H3B, &H48, &H51, &H77, &H86, &H4F)
IID_IUIAutomationWindowPattern = iid
End Function
Public Function IID_IUIAutomationTextRange() As UUID
'{a543cc6a-f4ae-494b-8239-c814481187a8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA543CC6A, CInt(&HF4AE), CInt(&H494B), &H82, &H39, &HC8, &H14, &H48, &H11, &H87, &HA8)
IID_IUIAutomationTextRange = iid
End Function
Public Function IID_IUIAutomationTextRange2() As UUID
'{BB9B40E0-5E04-46BD-9BE0-4B601B9AFAD4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBB9B40E0, CInt(&H5E04), CInt(&H46BD), &H9B, &HE0, &H4B, &H60, &H1B, &H9A, &HFA, &HD4)
IID_IUIAutomationTextRange2 = iid
End Function
Public Function IID_IUIAutomationTextRange3() As UUID
'{6A315D69-5512-4C2E-85F0-53FCE6DD4BC2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6A315D69, CInt(&H5512), CInt(&H4C2E), &H85, &HF0, &H53, &HFC, &HE6, &HDD, &H4B, &HC2)
IID_IUIAutomationTextRange3 = iid
End Function
Public Function IID_IUIAutomationTextRangeArray() As UUID
'{ce4ae76a-e717-4c98-81ea-47371d028eb6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCE4AE76A, CInt(&HE717), CInt(&H4C98), &H81, &HEA, &H47, &H37, &H1D, &H2, &H8E, &HB6)
IID_IUIAutomationTextRangeArray = iid
End Function
Public Function IID_IUIAutomationTextPattern() As UUID
'{32eba289-3583-42c9-9c59-3b6d9a1e9b6a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H32EBA289, CInt(&H3583), CInt(&H42C9), &H9C, &H59, &H3B, &H6D, &H9A, &H1E, &H9B, &H6A)
IID_IUIAutomationTextPattern = iid
End Function
Public Function IID_IUIAutomationTextPattern2() As UUID
'{506a921a-fcc9-409f-b23b-37eb74106872}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H506A921A, CInt(&HFCC9), CInt(&H409F), &HB2, &H3B, &H37, &HEB, &H74, &H10, &H68, &H72)
IID_IUIAutomationTextPattern2 = iid
End Function
Public Function IID_IUIAutomationTextEditPattern() As UUID
'{17E21576-996C-4870-99D9-BFF323380C06}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H17E21576, CInt(&H996C), CInt(&H4870), &H99, &HD9, &HBF, &HF3, &H23, &H38, &HC, &H6)
IID_IUIAutomationTextEditPattern = iid
End Function
Public Function IID_IUIAutomationCustomNavigationPattern() As UUID
'{01EA217A-1766-47ED-A6CC-ACF492854B1F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1EA217A, CInt(&H1766), CInt(&H47ED), &HA6, &HCC, &HAC, &HF4, &H92, &H85, &H4B, &H1F)
IID_IUIAutomationCustomNavigationPattern = iid
End Function
Public Function IID_IUIAutomationActiveTextPositionChangedEventHandler() As UUID
'{F97933B0-8DAE-4496-8997-5BA015FE0D82}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF97933B0, CInt(&H8DAE), CInt(&H4496), &H89, &H97, &H5B, &HA0, &H15, &HFE, &HD, &H82)
IID_IUIAutomationActiveTextPositionChangedEventHandler = iid
End Function
Public Function IID_IUIAutomationLegacyIAccessiblePattern() As UUID
'{828055ad-355b-4435-86d5-3b51c14a9b1b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H828055AD, CInt(&H355B), CInt(&H4435), &H86, &HD5, &H3B, &H51, &HC1, &H4A, &H9B, &H1B)
IID_IUIAutomationLegacyIAccessiblePattern = iid
End Function
Public Function IID_IUIAutomationItemContainerPattern() As UUID
'{c690fdb2-27a8-423c-812d-429773c9084e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC690FDB2, CInt(&H27A8), CInt(&H423C), &H81, &H2D, &H42, &H97, &H73, &HC9, &H8, &H4E)
IID_IUIAutomationItemContainerPattern = iid
End Function
Public Function IID_IUIAutomationVirtualizedItemPattern() As UUID
'{6ba3d7a6-04cf-4f11-8793-a8d1cde9969f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6BA3D7A6, CInt(&H4CF), CInt(&H4F11), &H87, &H93, &HA8, &HD1, &HCD, &HE9, &H96, &H9F)
IID_IUIAutomationVirtualizedItemPattern = iid
End Function
Public Function IID_IUIAutomationAnnotationPattern() As UUID
'{9a175b21-339e-41b1-8e8b-623f6b681098}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9A175B21, CInt(&H339E), CInt(&H41B1), &H8E, &H8B, &H62, &H3F, &H6B, &H68, &H10, &H98)
IID_IUIAutomationAnnotationPattern = iid
End Function
Public Function IID_IUIAutomationStylesPattern() As UUID
'{85b5f0a2-bd79-484a-ad2b-388c9838d5fb}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H85B5F0A2, CInt(&HBD79), CInt(&H484A), &HAD, &H2B, &H38, &H8C, &H98, &H38, &HD5, &HFB)
IID_IUIAutomationStylesPattern = iid
End Function
Public Function IID_IUIAutomationSpreadsheetPattern() As UUID
'{7517a7c8-faae-4de9-9f08-29b91e8595c1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7517A7C8, CInt(&HFAAE), CInt(&H4DE9), &H9F, &H8, &H29, &HB9, &H1E, &H85, &H95, &HC1)
IID_IUIAutomationSpreadsheetPattern = iid
End Function
Public Function IID_IUIAutomationSpreadsheetItemPattern() As UUID
'{7d4fb86c-8d34-40e1-8e83-62c15204e335}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7D4FB86C, CInt(&H8D34), CInt(&H40E1), &H8E, &H83, &H62, &HC1, &H52, &H4, &HE3, &H35)
IID_IUIAutomationSpreadsheetItemPattern = iid
End Function
Public Function IID_IUIAutomationTransformPattern2() As UUID
'{6d74d017-6ecb-4381-b38b-3c17a48ff1c2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6D74D017, CInt(&H6ECB), CInt(&H4381), &HB3, &H8B, &H3C, &H17, &HA4, &H8F, &HF1, &HC2)
IID_IUIAutomationTransformPattern2 = iid
End Function
Public Function IID_IUIAutomationTextChildPattern() As UUID
'{6552b038-ae05-40c8-abfd-aa08352aab86}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6552B038, CInt(&HAE05), CInt(&H40C8), &HAB, &HFD, &HAA, &H8, &H35, &H2A, &HAB, &H86)
IID_IUIAutomationTextChildPattern = iid
End Function
Public Function IID_IUIAutomationDragPattern() As UUID
'{1dc7b570-1f54-4bad-bcda-d36a722fb7bd}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1DC7B570, CInt(&H1F54), CInt(&H4BAD), &HBC, &HDA, &HD3, &H6A, &H72, &H2F, &HB7, &HBD)
IID_IUIAutomationDragPattern = iid
End Function
Public Function IID_IUIAutomationDropTargetPattern() As UUID
'{69a095f7-eee4-430e-a46b-fb73b1ae39a5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H69A095F7, CInt(&HEEE4), CInt(&H430E), &HA4, &H6B, &HFB, &H73, &HB1, &HAE, &H39, &HA5)
IID_IUIAutomationDropTargetPattern = iid
End Function
Public Function IID_IUIAutomationElement() As UUID
'{d22108aa-8ac5-49a5-837b-37bbb3d7591e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD22108AA, CInt(&H8AC5), CInt(&H49A5), &H83, &H7B, &H37, &HBB, &HB3, &HD7, &H59, &H1E)
IID_IUIAutomationElement = iid
End Function
Public Function IID_IUIAutomationElement2() As UUID
'{6749c683-f70d-4487-a698-5f79d55290d6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6749C683, CInt(&HF70D), CInt(&H4487), &HA6, &H98, &H5F, &H79, &HD5, &H52, &H90, &HD6)
IID_IUIAutomationElement2 = iid
End Function
Public Function IID_IUIAutomationElement3() As UUID
'{8471DF34-AEE0-4A01-A7DE-7DB9AF12C296}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8471DF34, CInt(&HAEE0), CInt(&H4A01), &HA7, &HDE, &H7D, &HB9, &HAF, &H12, &HC2, &H96)
IID_IUIAutomationElement3 = iid
End Function
Public Function IID_IUIAutomationElement4() As UUID
'{3B6E233C-52FB-4063-A4C9-77C075C2A06B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3B6E233C, CInt(&H52FB), CInt(&H4063), &HA4, &HC9, &H77, &HC0, &H75, &HC2, &HA0, &H6B)
IID_IUIAutomationElement4 = iid
End Function
Public Function IID_IUIAutomationElement5() As UUID
'{98141C1D-0D0E-4175-BBE2-6BFF455842A7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H98141C1D, CInt(&HD0E), CInt(&H4175), &HBB, &HE2, &H6B, &HFF, &H45, &H58, &H42, &HA7)
IID_IUIAutomationElement5 = iid
End Function
Public Function IID_IUIAutomationElement6() As UUID
'{4780d450-8bca-4977-afa5-a4a517f555e3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4780D450, CInt(&H8BCA), CInt(&H4977), &HAF, &HA5, &HA4, &HA5, &H17, &HF5, &H55, &HE3)
IID_IUIAutomationElement6 = iid
End Function
Public Function IID_IUIAutomationElement7() As UUID
'{204e8572-cfc3-4c11-b0c8-7da7420750b7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H204E8572, CInt(&HCFC3), CInt(&H4C11), &HB0, &HC8, &H7D, &HA7, &H42, &H7, &H50, &HB7)
IID_IUIAutomationElement7 = iid
End Function
Public Function IID_IUIAutomationElement8() As UUID
'{8C60217D-5411-4CDE-BCC0-1CEDA223830C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8C60217D, CInt(&H5411), CInt(&H4CDE), &HBC, &HC0, &H1C, &HED, &HA2, &H23, &H83, &HC)
IID_IUIAutomationElement8 = iid
End Function
Public Function IID_IUIAutomationElement9() As UUID
'{39325fac-039d-440e-a3a3-5eb81a5cecc3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H39325FAC, CInt(&H39D), CInt(&H440E), &HA3, &HA3, &H5E, &HB8, &H1A, &H5C, &HEC, &HC3)
IID_IUIAutomationElement9 = iid
End Function
Public Function IID_IUIAutomationElementArray() As UUID
'{14314595-b4bc-4055-95f2-58f2e42c9855}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H14314595, CInt(&HB4BC), CInt(&H4055), &H95, &HF2, &H58, &HF2, &HE4, &H2C, &H98, &H55)
IID_IUIAutomationElementArray = iid
End Function
Public Function IID_IUIAutomationProxyFactory() As UUID
'{85b94ecd-849d-42b6-b94d-d6db23fdf5a4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H85B94ECD, CInt(&H849D), CInt(&H42B6), &HB9, &H4D, &HD6, &HDB, &H23, &HFD, &HF5, &HA4)
IID_IUIAutomationProxyFactory = iid
End Function
Public Function IID_IUIAutomationProxyFactoryEntry() As UUID
'{d50e472e-b64b-490c-bca1-d30696f9f289}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD50E472E, CInt(&HB64B), CInt(&H490C), &HBC, &HA1, &HD3, &H6, &H96, &HF9, &HF2, &H89)
IID_IUIAutomationProxyFactoryEntry = iid
End Function
Public Function IID_IUIAutomationProxyFactoryMapping() As UUID
'{09e31e18-872d-4873-93d1-1e541ec133fd}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9E31E18, CInt(&H872D), CInt(&H4873), &H93, &HD1, &H1E, &H54, &H1E, &HC1, &H33, &HFD)
IID_IUIAutomationProxyFactoryMapping = iid
End Function
Public Function IID_IUIAutomationEventHandlerGroup() As UUID
'{C9EE12F2-C13B-4408-997C-639914377F4E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC9EE12F2, CInt(&HC13B), CInt(&H4408), &H99, &H7C, &H63, &H99, &H14, &H37, &H7F, &H4E)
IID_IUIAutomationEventHandlerGroup = iid
End Function
Public Function IID_IUIAutomation() As UUID
'{30cbe57d-d9d0-452a-ab13-7ac5ac4825ee}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H30CBE57D, CInt(&HD9D0), CInt(&H452A), &HAB, &H13, &H7A, &HC5, &HAC, &H48, &H25, &HEE)
IID_IUIAutomation = iid
End Function
Public Function IID_IUIAutomation2() As UUID
'{34723aff-0c9d-49d0-9896-7ab52df8cd8a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H34723AFF, CInt(&HC9D), CInt(&H49D0), &H98, &H96, &H7A, &HB5, &H2D, &HF8, &HCD, &H8A)
IID_IUIAutomation2 = iid
End Function
Public Function IID_IUIAutomation3() As UUID
'{73D768DA-9B51-4B89-936E-C209290973E7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H73D768DA, CInt(&H9B51), CInt(&H4B89), &H93, &H6E, &HC2, &H9, &H29, &H9, &H73, &HE7)
IID_IUIAutomation3 = iid
End Function
Public Function IID_IUIAutomation4() As UUID
'{1189C02A-05F8-4319-8E21-E817E3DB2860}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1189C02A, CInt(&H5F8), CInt(&H4319), &H8E, &H21, &HE8, &H17, &HE3, &HDB, &H28, &H60)
IID_IUIAutomation4 = iid
End Function
Public Function IID_IUIAutomation5() As UUID
'{25F700C8-D816-4057-A9DC-3CBDEE77E256}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H25F700C8, CInt(&HD816), CInt(&H4057), &HA9, &HDC, &H3C, &HBD, &HEE, &H77, &HE2, &H56)
IID_IUIAutomation5 = iid
End Function
Public Function IID_IUIAutomation6() As UUID
'{AAE072DA-29E3-413D-87A7-192DBF81ED10}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAAE072DA, CInt(&H29E3), CInt(&H413D), &H87, &HA7, &H19, &H2D, &HBF, &H81, &HED, &H10)
IID_IUIAutomation6 = iid
End Function




Public Function RuntimeId_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA39EEBFA, &H7FBA, &H4C89, &HB4, &HD4, &HB9, &H9E, &H2D, &HE7, &HD1, &H60)
RuntimeId_Property_GUID = iid
End Function
Public Function BoundingRectangle_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7BBFE8B2, &H3BFC, &H48DD, &HB7, &H29, &HC7, &H94, &HB8, &H46, &HE9, &HA1)
BoundingRectangle_Property_GUID = iid
End Function
Public Function ProcessId_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H40499998, &H9C31, &H4245, &HA4, &H3, &H87, &H32, &HE, &H59, &HEA, &HF6)
ProcessId_Property_GUID = iid
End Function
Public Function ControlType_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCA774FEA, &H28AC, &H4BC2, &H94, &HCA, &HAC, &HEC, &H6D, &H6C, &H10, &HA3)
ControlType_Property_GUID = iid
End Function
Public Function LocalizedControlType_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8763404F, &HA1BD, &H452A, &H89, &HC4, &H3F, &H1, &HD3, &H83, &H38, &H6)
LocalizedControlType_Property_GUID = iid
End Function
Public Function Name_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC3A6921B, &H4A99, &H44F1, &HBC, &HA6, &H61, &H18, &H70, &H52, &HC4, &H31)
Name_Property_GUID = iid
End Function
Public Function AcceleratorKey_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H514865DF, &H2557, &H4CB9, &HAE, &HED, &H6C, &HED, &H8, &H4C, &HE5, &H2C)
AcceleratorKey_Property_GUID = iid
End Function
Public Function AccessKey_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6827B12, &HA7F9, &H4A15, &H91, &H7C, &HFF, &HA5, &HAD, &H3E, &HB0, &HA7)
AccessKey_Property_GUID = iid
End Function
Public Function HasKeyboardFocus_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCF8AFD39, &H3F46, &H4800, &H96, &H56, &HB2, &HBF, &H12, &H52, &H99, &H5)
HasKeyboardFocus_Property_GUID = iid
End Function
Public Function IsKeyboardFocusable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF7B8552A, &H859, &H4B37, &HB9, &HCB, &H51, &HE7, &H20, &H92, &HF2, &H9F)
IsKeyboardFocusable_Property_GUID = iid
End Function
Public Function IsEnabled_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2109427F, &HDA60, &H4FED, &HBF, &H1B, &H26, &H4B, &HDC, &HE6, &HEB, &H3A)
IsEnabled_Property_GUID = iid
End Function
Public Function AutomationId_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC82C0500, &HB60E, &H4310, &HA2, &H67, &H30, &H3C, &H53, &H1F, &H8E, &HE5)
AutomationId_Property_GUID = iid
End Function
Public Function ClassName_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H157B7215, &H894F, &H4B65, &H84, &HE2, &HAA, &HC0, &HDA, &H8, &HB1, &H6B)
ClassName_Property_GUID = iid
End Function
Public Function HelpText_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8555685, &H977, &H45C7, &HA7, &HA6, &HAB, &HAF, &H56, &H84, &H12, &H1A)
HelpText_Property_GUID = iid
End Function
Public Function ClickablePoint_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H196903B, &HB203, &H4818, &HA9, &HF3, &HF0, &H8E, &H67, &H5F, &H23, &H41)
ClickablePoint_Property_GUID = iid
End Function
Public Function Culture_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE2D74F27, &H3D79, &H4DC2, &HB8, &H8B, &H30, &H44, &H96, &H3A, &H8A, &HFB)
Culture_Property_GUID = iid
End Function
Public Function IsControlElement_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H95F35085, &HABCC, &H4AFD, &HA5, &HF4, &HDB, &HB4, &H6C, &H23, &HF, &HDB)
IsControlElement_Property_GUID = iid
End Function
Public Function IsContentElement_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4BDA64A8, &HF5D8, &H480B, &H81, &H55, &HEF, &H2E, &H89, &HAD, &HB6, &H72)
IsContentElement_Property_GUID = iid
End Function
Public Function LabeledBy_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE5B8924B, &HFC8A, &H4A35, &H80, &H31, &HCF, &H78, &HAC, &H43, &HE5, &H5E)
LabeledBy_Property_GUID = iid
End Function
Public Function IsPassword_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE8482EB1, &H687C, &H497B, &HBE, &HBC, &H3, &HBE, &H53, &HEC, &H14, &H54)
IsPassword_Property_GUID = iid
End Function
Public Function NewNativeWindowHandle_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5196B33B, &H380A, &H4982, &H95, &HE1, &H91, &HF3, &HEF, &H60, &HE0, &H24)
NewNativeWindowHandle_Property_GUID = iid
End Function
Public Function ItemType_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCDDA434D, &H6222, &H413B, &HA6, &H8A, &H32, &H5D, &HD1, &HD4, &HF, &H39)
ItemType_Property_GUID = iid
End Function
Public Function IsOffscreen_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3C3D160, &HDB79, &H42DB, &HA2, &HEF, &H1C, &H23, &H1E, &HED, &HE5, &H7)
IsOffscreen_Property_GUID = iid
End Function
Public Function Orientation_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA01EEE62, &H3884, &H4415, &H88, &H7E, &H67, &H8E, &HC2, &H1E, &H39, &HBA)
Orientation_Property_GUID = iid
End Function
Public Function FrameworkId_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDBFD9900, &H7E1A, &H4F58, &HB6, &H1B, &H70, &H63, &H12, &HF, &H77, &H3B)
FrameworkId_Property_GUID = iid
End Function
Public Function IsRequiredForForm_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4F5F43CF, &H59FB, &H4BDE, &HA2, &H70, &H60, &H2E, &H5E, &H11, &H41, &HE9)
IsRequiredForForm_Property_GUID = iid
End Function
Public Function ItemStatus_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H51DE0321, &H3973, &H43E7, &H89, &H13, &HB, &H8, &HE8, &H13, &HC3, &H7F)
ItemStatus_Property_GUID = iid
End Function
Public Function AriaRole_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDD207B95, &HBE4A, &H4E0D, &HB7, &H27, &H63, &HAC, &HE9, &H4B, &H69, &H16)
AriaRole_Property_GUID = iid
End Function
Public Function AriaProperties_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4213678C, &HE025, &H4922, &HBE, &HB5, &HE4, &H3B, &HA0, &H8E, &H62, &H21)
AriaProperties_Property_GUID = iid
End Function
Public Function IsDataValidForForm_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H445AC684, &HC3FC, &H4DD9, &HAC, &HF8, &H84, &H5A, &H57, &H92, &H96, &HBA)
IsDataValidForForm_Property_GUID = iid
End Function
Public Function ControllerFor_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H51124C8A, &HA5D2, &H4F13, &H9B, &HE6, &H7F, &HA8, &HBA, &H9D, &H3A, &H90)
ControllerFor_Property_GUID = iid
End Function
Public Function DescribedBy_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7C5865B8, &H9992, &H40FD, &H8D, &HB0, &H6B, &HF1, &HD3, &H17, &HF9, &H98)
DescribedBy_Property_GUID = iid
End Function
Public Function FlowsTo_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE4F33D20, &H559A, &H47FB, &HA8, &H30, &HF9, &HCB, &H4F, &HF1, &HA7, &HA)
FlowsTo_Property_GUID = iid
End Function
Public Function ProviderDescription_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDCA5708A, &HC16B, &H4CD9, &HB8, &H89, &HBE, &HB1, &H6A, &H80, &H49, &H4)
ProviderDescription_Property_GUID = iid
End Function
Public Function OptimizeForVisualContent_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6A852250, &HC75A, &H4E5D, &HB8, &H58, &HE3, &H81, &HB0, &HF7, &H88, &H61)
OptimizeForVisualContent_Property_GUID = iid
End Function
Public Function IsDockPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2600A4C4, &H2FF8, &H4C96, &HAE, &H31, &H8F, &HE6, &H19, &HA1, &H3C, &H6C)
IsDockPatternAvailable_Property_GUID = iid
End Function
Public Function IsExpandCollapsePatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H929D3806, &H5287, &H4725, &HAA, &H16, &H22, &H2A, &HFC, &H63, &HD5, &H95)
IsExpandCollapsePatternAvailable_Property_GUID = iid
End Function
Public Function IsGridItemPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5A43E524, &HF9A2, &H4B12, &H84, &HC8, &HB4, &H8A, &H3E, &HFE, &HDD, &H34)
IsGridItemPatternAvailable_Property_GUID = iid
End Function
Public Function IsGridPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5622C26C, &HF0EF, &H4F3B, &H97, &HCB, &H71, &H4C, &H8, &H68, &H58, &H8B)
IsGridPatternAvailable_Property_GUID = iid
End Function
Public Function IsInvokePatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4E725738, &H8364, &H4679, &HAA, &H6C, &HF3, &HF4, &H19, &H31, &HF7, &H50)
IsInvokePatternAvailable_Property_GUID = iid
End Function
Public Function IsMultipleViewPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFF0A31EB, &H8E25, &H469D, &H8D, &H6E, &HE7, &H71, &HA2, &H7C, &H1B, &H90)
IsMultipleViewPatternAvailable_Property_GUID = iid
End Function
Public Function IsRangeValuePatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFDA4244A, &HEB4D, &H43FF, &HB5, &HAD, &HED, &H36, &HD3, &H73, &HEC, &H4C)
IsRangeValuePatternAvailable_Property_GUID = iid
End Function
Public Function IsScrollPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3EBB7B4A, &H828A, &H4B57, &H9D, &H22, &H2F, &HEA, &H16, &H32, &HED, &HD)
IsScrollPatternAvailable_Property_GUID = iid
End Function
Public Function IsScrollItemPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1CAD1A05, &H927, &H4B76, &H97, &HE1, &HF, &HCD, &HB2, &H9, &HB9, &H8A)
IsScrollItemPatternAvailable_Property_GUID = iid
End Function
Public Function IsSelectionItemPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8BECD62D, &HBC3, &H4109, &HBE, &HE2, &H8E, &H67, &H15, &H29, &HE, &H68)
IsSelectionItemPatternAvailable_Property_GUID = iid
End Function
Public Function IsSelectionPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF588ACBE, &HC769, &H4838, &H9A, &H60, &H26, &H86, &HDC, &H11, &H88, &HC4)
IsSelectionPatternAvailable_Property_GUID = iid
End Function
Public Function IsTablePatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCB83575F, &H45C2, &H4048, &H9C, &H76, &H15, &H97, &H15, &HA1, &H39, &HDF)
IsTablePatternAvailable_Property_GUID = iid
End Function
Public Function IsTableItemPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEB36B40D, &H8EA4, &H489B, &HA0, &H13, &HE6, &HD, &H59, &H51, &HFE, &H34)
IsTableItemPatternAvailable_Property_GUID = iid
End Function
Public Function IsTextPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFBE2D69D, &HAFF6, &H4A45, &H82, &HE2, &HFC, &H92, &HA8, &H2F, &H59, &H17)
IsTextPatternAvailable_Property_GUID = iid
End Function
Public Function IsTogglePatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H78686D53, &HFCD0, &H4B83, &H9B, &H78, &H58, &H32, &HCE, &H63, &HBB, &H5B)
IsTogglePatternAvailable_Property_GUID = iid
End Function
Public Function IsTransformPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA7F78804, &HD68B, &H4077, &HA5, &HC6, &H7A, &H5E, &HA1, &HAC, &H31, &HC5)
IsTransformPatternAvailable_Property_GUID = iid
End Function
Public Function IsValuePatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB5020A7, &H2119, &H473B, &HBE, &H37, &H5C, &HEB, &H98, &HBB, &HFB, &H22)
IsValuePatternAvailable_Property_GUID = iid
End Function
Public Function IsWindowPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE7A57BB1, &H5888, &H4155, &H98, &HDC, &HB4, &H22, &HFD, &H57, &HF2, &HBC)
IsWindowPatternAvailable_Property_GUID = iid
End Function
Public Function IsLegacyIAccessiblePatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD8EBD0C7, &H929A, &H4EE7, &H8D, &H3A, &HD3, &HD9, &H44, &H13, &H2, &H7B)
IsLegacyIAccessiblePatternAvailable_Property_GUID = iid
End Function
Public Function IsItemContainerPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H624B5CA7, &HFE40, &H4957, &HA0, &H19, &H20, &HC4, &HCF, &H11, &H92, &HF)
IsItemContainerPatternAvailable_Property_GUID = iid
End Function
Public Function IsVirtualizedItemPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H302CB151, &H2AC8, &H45D6, &H97, &H7B, &HD2, &HB3, &HA5, &HA5, &H3F, &H20)
IsVirtualizedItemPatternAvailable_Property_GUID = iid
End Function
Public Function IsSynchronizedInputPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H75D69CC5, &HD2BF, &H4943, &H87, &H6E, &HB4, &H5B, &H62, &HA6, &HCC, &H66)
IsSynchronizedInputPatternAvailable_Property_GUID = iid
End Function
Public Function IsObjectModelPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6B21D89B, &H2841, &H412F, &H8E, &HF2, &H15, &HCA, &H95, &H23, &H18, &HBA)
IsObjectModelPatternAvailable_Property_GUID = iid
End Function
Public Function IsAnnotationPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB5B3238, &H6D5C, &H41B6, &HBC, &HC4, &H5E, &H80, &H7F, &H65, &H51, &HC4)
IsAnnotationPatternAvailable_Property_GUID = iid
End Function
Public Function IsTextPattern2Available_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H41CF921D, &HE3F1, &H4B22, &H9C, &H81, &HE1, &HC3, &HED, &H33, &H1C, &H22)
IsTextPattern2Available_Property_GUID = iid
End Function
Public Function IsTextEditPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7843425C, &H8B32, &H484C, &H9A, &HB5, &HE3, &H20, &H5, &H71, &HFF, &HDA)
IsTextEditPatternAvailable_Property_GUID = iid
End Function
Public Function IsCustomNavigationPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8F8E80D4, &H2351, &H48E0, &H87, &H4A, &H54, &HAA, &H73, &H13, &H88, &H9A)
IsCustomNavigationPatternAvailable_Property_GUID = iid
End Function
Public Function IsStylesPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H27F353D3, &H459C, &H4B59, &HA4, &H90, &H50, &H61, &H1D, &HAC, &HAF, &HB5)
IsStylesPatternAvailable_Property_GUID = iid
End Function
Public Function IsSpreadsheetPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6FF43732, &HE4B4, &H4555, &H97, &HBC, &HEC, &HDB, &HBC, &H4D, &H18, &H88)
IsSpreadsheetPatternAvailable_Property_GUID = iid
End Function
Public Function IsSpreadsheetItemPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9FE79B2A, &H2F94, &H43FD, &H99, &H6B, &H54, &H9E, &H31, &H6F, &H4A, &HCD)
IsSpreadsheetItemPatternAvailable_Property_GUID = iid
End Function
Public Function IsTransformPattern2Available_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H25980B4B, &HBE04, &H4710, &HAB, &H4A, &HFD, &HA3, &H1D, &HBD, &H28, &H95)
IsTransformPattern2Available_Property_GUID = iid
End Function
Public Function IsTextChildPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H559E65DF, &H30FF, &H43B5, &HB5, &HED, &H5B, &H28, &H3B, &H80, &HC7, &HE9)
IsTextChildPatternAvailable_Property_GUID = iid
End Function
Public Function IsDragPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE997A7B7, &H1D39, &H4CA7, &HBE, &HF, &H27, &H7F, &HCF, &H56, &H5, &HCC)
IsDragPatternAvailable_Property_GUID = iid
End Function
Public Function IsDropTargetPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H686B62E, &H8E19, &H4AAF, &H87, &H3D, &H38, &H4F, &H6D, &H3B, &H92, &HBE)
IsDropTargetPatternAvailable_Property_GUID = iid
End Function
Public Function IsStructuredMarkupPatternAvailable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB0D4C196, &H2C0B, &H489C, &HB1, &H65, &HA4, &H5, &H92, &H8C, &H6F, &H3D)
IsStructuredMarkupPatternAvailable_Property_GUID = iid
End Function
Public Function IsPeripheral_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDA758276, &H7ED5, &H49D4, &H8E, &H68, &HEC, &HC9, &HA2, &HD3, &H0, &HDD)
IsPeripheral_Property_GUID = iid
End Function
Public Function PositionInSet_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H33D1DC54, &H641E, &H4D76, &HA6, &HB1, &H13, &HF3, &H41, &HC1, &HF8, &H96)
PositionInSet_Property_GUID = iid
End Function
Public Function SizeOfSet_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1600D33C, &H3B9F, &H4369, &H94, &H31, &HAA, &H29, &H3F, &H34, &H4C, &HF1)
SizeOfSet_Property_GUID = iid
End Function
Public Function Level_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H242AC529, &HCD36, &H400F, &HAA, &HD9, &H78, &H76, &HEF, &H3A, &HF6, &H27)
Level_Property_GUID = iid
End Function
Public Function AnnotationTypes_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H64B71F76, &H53C4, &H4696, &HA2, &H19, &H20, &HE9, &H40, &HC9, &HA1, &H76)
AnnotationTypes_Property_GUID = iid
End Function
Public Function AnnotationObjects_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H310910C8, &H7C6E, &H4F20, &HBE, &HCD, &H4A, &HAF, &H6D, &H19, &H11, &H56)
AnnotationObjects_Property_GUID = iid
End Function
Public Function LandmarkType_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H454045F2, &H6F61, &H49F7, &HA4, &HF8, &HB5, &HF0, &HCF, &H82, &HDA, &H1E)
LandmarkType_Property_GUID = iid
End Function
Public Function LocalizedLandmarkType_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7AC81980, &HEAFB, &H4FB2, &HBF, &H91, &HF4, &H85, &HBE, &HF5, &HE8, &HE1)
LocalizedLandmarkType_Property_GUID = iid
End Function
Public Function FullDescription_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD4450FF, &H6AEF, &H4F33, &H95, &HDD, &H7B, &HEF, &HA7, &H2A, &H43, &H91)
FullDescription_Property_GUID = iid
End Function
Public Function Value_Value_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE95F5E64, &H269F, &H4A85, &HBA, &H99, &H40, &H92, &HC3, &HEA, &H29, &H86)
Value_Value_Property_GUID = iid
End Function
Public Function Value_IsReadOnly_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEB090F30, &HE24C, &H4799, &HA7, &H5, &HD, &H24, &H7B, &HC0, &H37, &HF8)
Value_IsReadOnly_Property_GUID = iid
End Function
Public Function RangeValue_Value_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H131F5D98, &HC50C, &H489D, &HAB, &HE5, &HAE, &H22, &H8, &H98, &HC5, &HF7)
RangeValue_Value_Property_GUID = iid
End Function
Public Function RangeValue_IsReadOnly_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H25FA1055, &HDEBF, &H4373, &HA7, &H9E, &H1F, &H1A, &H19, &H8, &HD3, &HC4)
RangeValue_IsReadOnly_Property_GUID = iid
End Function
Public Function RangeValue_Minimum_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H78CBD3B2, &H684D, &H4860, &HAF, &H93, &HD1, &HF9, &H5C, &HB0, &H22, &HFD)
RangeValue_Minimum_Property_GUID = iid
End Function
Public Function RangeValue_Maximum_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H19319914, &HF979, &H4B35, &HA1, &HA6, &HD3, &H7E, &H5, &H43, &H34, &H73)
RangeValue_Maximum_Property_GUID = iid
End Function
Public Function RangeValue_LargeChange_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA1F96325, &H3A3D, &H4B44, &H8E, &H1F, &H4A, &H46, &HD9, &H84, &H40, &H19)
RangeValue_LargeChange_Property_GUID = iid
End Function
Public Function RangeValue_SmallChange_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H81C2C457, &H3941, &H4107, &H99, &H75, &H13, &H97, &H60, &HF7, &HC0, &H72)
RangeValue_SmallChange_Property_GUID = iid
End Function
Public Function Scroll_HorizontalScrollPercent_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC7C13C0E, &HEB21, &H47FF, &HAC, &HC4, &HB5, &HA3, &H35, &HF, &H51, &H91)
Scroll_HorizontalScrollPercent_Property_GUID = iid
End Function
Public Function Scroll_HorizontalViewSize_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H70C2E5D4, &HFCB0, &H4713, &HA9, &HAA, &HAF, &H92, &HFF, &H79, &HE4, &HCD)
Scroll_HorizontalViewSize_Property_GUID = iid
End Function
Public Function Scroll_VerticalScrollPercent_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C8D7099, &HB2A8, &H4948, &HBF, &HF7, &H3C, &HF9, &H5, &H8B, &HFE, &HFB)
Scroll_VerticalScrollPercent_Property_GUID = iid
End Function
Public Function Scroll_VerticalViewSize_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDE6A2E22, &HD8C7, &H40C5, &H83, &HBA, &HE5, &HF6, &H81, &HD5, &H31, &H8)
Scroll_VerticalViewSize_Property_GUID = iid
End Function
Public Function Scroll_HorizontallyScrollable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8B925147, &H28CD, &H49AE, &HBD, &H63, &HF4, &H41, &H18, &HD2, &HE7, &H19)
Scroll_HorizontallyScrollable_Property_GUID = iid
End Function
Public Function Scroll_VerticallyScrollable_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H89164798, &H68, &H4315, &HB8, &H9A, &H1E, &H7C, &HFB, &HBC, &H3D, &HFC)
Scroll_VerticallyScrollable_Property_GUID = iid
End Function
Public Function Selection_Selection_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAA6DC2A2, &HE2B, &H4D38, &H96, &HD5, &H34, &HE4, &H70, &HB8, &H18, &H53)
Selection_Selection_Property_GUID = iid
End Function
Public Function Selection_CanSelectMultiple_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H49D73DA5, &HC883, &H4500, &H88, &H3D, &H8F, &HCF, &H8D, &HAF, &H6C, &HBE)
Selection_CanSelectMultiple_Property_GUID = iid
End Function
Public Function Selection_IsSelectionRequired_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB1AE4422, &H63FE, &H44E7, &HA5, &HA5, &HA7, &H38, &HC8, &H29, &HB1, &H9A)
Selection_IsSelectionRequired_Property_GUID = iid
End Function
Public Function Grid_RowCount_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2A9505BF, &HC2EB, &H4FB6, &HB3, &H56, &H82, &H45, &HAE, &H53, &H70, &H3E)
Grid_RowCount_Property_GUID = iid
End Function
Public Function Grid_ColumnCount_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFE96F375, &H44AA, &H4536, &HAC, &H7A, &H2A, &H75, &HD7, &H1A, &H3E, &HFC)
Grid_ColumnCount_Property_GUID = iid
End Function
Public Function GridItem_Row_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6223972A, &HC945, &H4563, &H93, &H29, &HFD, &HC9, &H74, &HAF, &H25, &H53)
GridItem_Row_Property_GUID = iid
End Function
Public Function GridItem_Column_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC774C15C, &H62C0, &H4519, &H8B, &HDC, &H47, &HBE, &H57, &H3C, &H8A, &HD5)
GridItem_Column_Property_GUID = iid
End Function
Public Function GridItem_RowSpan_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4582291C, &H466B, &H4E93, &H8E, &H83, &H3D, &H17, &H15, &HEC, &HC, &H5E)
GridItem_RowSpan_Property_GUID = iid
End Function
Public Function GridItem_ColumnSpan_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H583EA3F5, &H86D0, &H4B08, &HA6, &HEC, &H2C, &H54, &H63, &HFF, &HC1, &H9)
GridItem_ColumnSpan_Property_GUID = iid
End Function
Public Function GridItem_Parent_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9D912252, &HB97F, &H4ECC, &H85, &H10, &HEA, &HE, &H33, &H42, &H7C, &H72)
GridItem_Parent_Property_GUID = iid
End Function
Public Function Dock_DockPosition_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6D67F02E, &HC0B0, &H4B10, &HB5, &HB9, &H18, &HD6, &HEC, &HF9, &H87, &H60)
Dock_DockPosition_Property_GUID = iid
End Function
Public Function ExpandCollapse_ExpandCollapseState_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H275A4C48, &H85A7, &H4F69, &HAB, &HA0, &HAF, &H15, &H76, &H10, &H0, &H2B)
ExpandCollapse_ExpandCollapseState_Property_GUID = iid
End Function
Public Function MultipleView_CurrentView_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7A81A67A, &HB94F, &H4875, &H91, &H8B, &H65, &HC8, &HD2, &HF9, &H98, &HE5)
MultipleView_CurrentView_Property_GUID = iid
End Function
Public Function MultipleView_SupportedViews_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8D5DB9FD, &HCE3C, &H4AE7, &HB7, &H88, &H40, &HA, &H3C, &H64, &H55, &H47)
MultipleView_SupportedViews_Property_GUID = iid
End Function
Public Function Window_CanMaximize_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H64FFF53F, &H635D, &H41C1, &H95, &HC, &HCB, &H5A, &HDF, &HBE, &H28, &HE3)
Window_CanMaximize_Property_GUID = iid
End Function
Public Function Window_CanMinimize_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB73B4625, &H5988, &H4B97, &HB4, &HC2, &HA6, &HFE, &H6E, &H78, &HC8, &HC6)
Window_CanMinimize_Property_GUID = iid
End Function
Public Function Window_WindowVisualState_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4AB7905F, &HE860, &H453E, &HA3, &HA, &HF6, &H43, &H1E, &H5D, &HAA, &HD5)
Window_WindowVisualState_Property_GUID = iid
End Function
Public Function Window_WindowInteractionState_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4FED26A4, &H455, &H4FA2, &HB2, &H1C, &HC4, &HDA, &H2D, &HB1, &HFF, &H9C)
Window_WindowInteractionState_Property_GUID = iid
End Function
Public Function Window_IsModal_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFF4E6892, &H37B9, &H4FCA, &H85, &H32, &HFF, &HE6, &H74, &HEC, &HFE, &HED)
Window_IsModal_Property_GUID = iid
End Function
Public Function Window_IsTopmost_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEF7D85D3, &H937, &H4962, &H92, &H41, &HB6, &H23, &H45, &HF2, &H40, &H41)
Window_IsTopmost_Property_GUID = iid
End Function
Public Function SelectionItem_IsSelected_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF122835F, &HCD5F, &H43DF, &HB7, &H9D, &H4B, &H84, &H9E, &H9E, &H60, &H20)
SelectionItem_IsSelected_Property_GUID = iid
End Function
Public Function SelectionItem_SelectionContainer_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA4365B6E, &H9C1E, &H4B63, &H8B, &H53, &HC2, &H42, &H1D, &HD1, &HE8, &HFB)
SelectionItem_SelectionContainer_Property_GUID = iid
End Function
Public Function Table_RowHeaders_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD9E35B87, &H6EB8, &H4562, &HAA, &HC6, &HA8, &HA9, &H7, &H52, &H36, &HA8)
Table_RowHeaders_Property_GUID = iid
End Function
Public Function Table_ColumnHeaders_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAFF1D72B, &H968D, &H42B1, &HB4, &H59, &H15, &HB, &H29, &H9D, &HA6, &H64)
Table_ColumnHeaders_Property_GUID = iid
End Function
Public Function Table_RowOrColumnMajor_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H83BE75C3, &H29FE, &H4A30, &H85, &HE1, &H2A, &H62, &H77, &HFD, &H10, &H6E)
Table_RowOrColumnMajor_Property_GUID = iid
End Function
Public Function TableItem_RowHeaderItems_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB3F853A0, &H574, &H4CD8, &HBC, &HD7, &HED, &H59, &H23, &H57, &H2D, &H97)
TableItem_RowHeaderItems_Property_GUID = iid
End Function
Public Function TableItem_ColumnHeaderItems_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H967A56A3, &H74B6, &H431E, &H8D, &HE6, &H99, &HC4, &H11, &H3, &H1C, &H58)
TableItem_ColumnHeaderItems_Property_GUID = iid
End Function
Public Function Toggle_ToggleState_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB23CDC52, &H22C2, &H4C6C, &H9D, &HED, &HF5, &HC4, &H22, &H47, &H9E, &HDE)
Toggle_ToggleState_Property_GUID = iid
End Function
Public Function Transform_CanMove_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B75824D, &H208B, &H4FDF, &HBC, &HCD, &HF1, &HF4, &HE5, &H74, &H1F, &H4F)
Transform_CanMove_Property_GUID = iid
End Function
Public Function Transform_CanResize_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBB98DCA5, &H4C1A, &H41D4, &HA4, &HF6, &HEB, &HC1, &H28, &H64, &H41, &H80)
Transform_CanResize_Property_GUID = iid
End Function
Public Function Transform_CanRotate_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H10079B48, &H3849, &H476F, &HAC, &H96, &H44, &HA9, &H5C, &H84, &H40, &HD9)
Transform_CanRotate_Property_GUID = iid
End Function
Public Function LegacyIAccessible_ChildId_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9A191B5D, &H9EF2, &H4787, &HA4, &H59, &HDC, &HDE, &H88, &H5D, &HD4, &HE8)
LegacyIAccessible_ChildId_Property_GUID = iid
End Function
Public Function LegacyIAccessible_Name_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCAEB063D, &H40AE, &H4869, &HAA, &H5A, &H1B, &H8E, &H5D, &H66, &H67, &H39)
LegacyIAccessible_Name_Property_GUID = iid
End Function
Public Function LegacyIAccessible_Value_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB5C5B0B6, &H8217, &H4A77, &H97, &HA5, &H19, &HA, &H85, &HED, &H1, &H56)
LegacyIAccessible_Value_Property_GUID = iid
End Function
Public Function LegacyIAccessible_Description_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H46448418, &H7D70, &H4EA9, &H9D, &H27, &HB7, &HE7, &H75, &HCF, &H2A, &HD7)
LegacyIAccessible_Description_Property_GUID = iid
End Function
Public Function LegacyIAccessible_Role_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6856E59F, &HCBAF, &H4E31, &H93, &HE8, &HBC, &HBF, &H6F, &H7E, &H49, &H1C)
LegacyIAccessible_Role_Property_GUID = iid
End Function
Public Function LegacyIAccessible_State_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDF985854, &H2281, &H4340, &HAB, &H9C, &HC6, &HE, &H2C, &H58, &H3, &HF6)
LegacyIAccessible_State_Property_GUID = iid
End Function
Public Function LegacyIAccessible_Help_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H94402352, &H161C, &H4B77, &HA9, &H8D, &HA8, &H72, &HCC, &H33, &H94, &H7A)
LegacyIAccessible_Help_Property_GUID = iid
End Function
Public Function LegacyIAccessible_KeyboardShortcut_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8F6909AC, &HB8, &H4259, &HA4, &H1C, &H96, &H62, &H66, &HD4, &H3A, &H8A)
LegacyIAccessible_KeyboardShortcut_Property_GUID = iid
End Function
Public Function LegacyIAccessible_Selection_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8AA8B1E0, &H891, &H40CC, &H8B, &H6, &H90, &HD7, &HD4, &H16, &H62, &H19)
LegacyIAccessible_Selection_Property_GUID = iid
End Function
Public Function LegacyIAccessible_DefaultAction_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3B331729, &HEAAD, &H4502, &HB8, &H5F, &H92, &H61, &H56, &H22, &H91, &H3C)
LegacyIAccessible_DefaultAction_Property_GUID = iid
End Function
Public Function Annotation_AnnotationTypeId_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H20AE484F, &H69EF, &H4C48, &H8F, &H5B, &HC4, &H93, &H8B, &H20, &H6A, &HC7)
Annotation_AnnotationTypeId_Property_GUID = iid
End Function
Public Function Annotation_AnnotationTypeName_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9B818892, &H5AC9, &H4AF9, &HAA, &H96, &HF5, &H8A, &H77, &HB0, &H58, &HE3)
Annotation_AnnotationTypeName_Property_GUID = iid
End Function
Public Function Annotation_Author_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7A528462, &H9C5C, &H4A03, &HA9, &H74, &H8B, &H30, &H7A, &H99, &H37, &HF2)
Annotation_Author_Property_GUID = iid
End Function
Public Function Annotation_DateTime_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H99B5CA5D, &H1ACF, &H414B, &HA4, &HD0, &H6B, &H35, &HB, &H4, &H75, &H78)
Annotation_DateTime_Property_GUID = iid
End Function
Public Function Annotation_Target_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB71B302D, &H2104, &H44AD, &H9C, &H5C, &H9, &H2B, &H49, &H7, &HD7, &HF)
Annotation_Target_Property_GUID = iid
End Function
Public Function Styles_StyleId_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDA82852F, &H3817, &H4233, &H82, &HAF, &H2, &H27, &H9E, &H72, &HCC, &H77)
Styles_StyleId_Property_GUID = iid
End Function
Public Function Styles_StyleName_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1C12B035, &H5D1, &H4F55, &H9E, &H8E, &H14, &H89, &HF3, &HFF, &H55, &HD)
Styles_StyleName_Property_GUID = iid
End Function
Public Function Styles_FillColor_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H63EFF97A, &HA1C5, &H4B1D, &H84, &HEB, &HB7, &H65, &HF2, &HED, &HD6, &H32)
Styles_FillColor_Property_GUID = iid
End Function
Public Function Styles_FillPatternStyle_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H81CF651F, &H482B, &H4451, &HA3, &HA, &HE1, &H54, &H5E, &H55, &H4F, &HB8)
Styles_FillPatternStyle_Property_GUID = iid
End Function
Public Function Styles_Shape_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC71A23F8, &H778C, &H400D, &H84, &H58, &H3B, &H54, &H3E, &H52, &H69, &H84)
Styles_Shape_Property_GUID = iid
End Function
Public Function Styles_FillPatternColor_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H939A59FE, &H8FBD, &H4E75, &HA2, &H71, &HAC, &H45, &H95, &H19, &H51, &H63)
Styles_FillPatternColor_Property_GUID = iid
End Function
Public Function Styles_ExtendedProperties_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF451CDA0, &HBA0A, &H4681, &HB0, &HB0, &HD, &HBD, &HB5, &H3E, &H58, &HF3)
Styles_ExtendedProperties_Property_GUID = iid
End Function
Public Function SpreadsheetItem_Formula_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE602E47D, &H1B47, &H4BEA, &H87, &HCF, &H3B, &HB, &HB, &H5C, &H15, &HB6)
SpreadsheetItem_Formula_Property_GUID = iid
End Function
Public Function SpreadsheetItem_AnnotationObjects_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA3194C38, &HC9BC, &H4604, &H93, &H96, &HAE, &H3F, &H9F, &H45, &H7F, &H7B)
SpreadsheetItem_AnnotationObjects_Property_GUID = iid
End Function
Public Function SpreadsheetItem_AnnotationTypes_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC70C51D0, &HD602, &H4B45, &HAF, &HBC, &HB4, &H71, &H2B, &H96, &HD7, &H2B)
SpreadsheetItem_AnnotationTypes_Property_GUID = iid
End Function
Public Function Transform2_CanZoom_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF357E890, &HA756, &H4359, &H9C, &HA6, &H86, &H70, &H2B, &HF8, &HF3, &H81)
Transform2_CanZoom_Property_GUID = iid
End Function
Public Function LiveSetting_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC12BCD8E, &H2A8E, &H4950, &H8A, &HE7, &H36, &H25, &H11, &H1D, &H58, &HEB)
LiveSetting_Property_GUID = iid
End Function
Public Function Drag_IsGrabbed_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H45F206F3, &H75CC, &H4CCA, &HA9, &HB9, &HFC, &HDF, &HB9, &H82, &HD8, &HA2)
Drag_IsGrabbed_Property_GUID = iid
End Function
Public Function Drag_GrabbedItems_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H77C1562C, &H7B86, &H4B21, &H9E, &HD7, &H3C, &HEF, &HDA, &H6F, &H4C, &H43)
Drag_GrabbedItems_Property_GUID = iid
End Function
Public Function Drag_DropEffect_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H646F2779, &H48D3, &H4B23, &H89, &H2, &H4B, &HF1, &H0, &H0, &H5D, &HF3)
Drag_DropEffect_Property_GUID = iid
End Function
Public Function Drag_DropEffects_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF5D61156, &H7CE6, &H49BE, &HA8, &H36, &H92, &H69, &HDC, &HEC, &H92, &HF)
Drag_DropEffects_Property_GUID = iid
End Function
Public Function DropTarget_DropTargetEffect_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8BB75975, &HA0CA, &H4981, &HB8, &H18, &H87, &HFC, &H66, &HE9, &H50, &H9D)
DropTarget_DropTargetEffect_Property_GUID = iid
End Function
Public Function DropTarget_DropTargetEffects_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBC1DD4ED, &HCB89, &H45F1, &HA5, &H92, &HE0, &H3B, &H8, &HAE, &H79, &HF)
DropTarget_DropTargetEffects_Property_GUID = iid
End Function
Public Function Transform2_ZoomLevel_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEEE29F1A, &HF4A2, &H4B5B, &HAC, &H65, &H95, &HCF, &H93, &H28, &H33, &H87)
Transform2_ZoomLevel_Property_GUID = iid
End Function
Public Function Transform2_ZoomMinimum_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H742CCC16, &H4AD1, &H4E07, &H96, &HFE, &HB1, &H22, &HC6, &HE6, &HB2, &H2B)
Transform2_ZoomMinimum_Property_GUID = iid
End Function
Public Function Transform2_ZoomMaximum_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H42AB6B77, &HCEB0, &H4ECA, &HB8, &H2A, &H6C, &HFA, &H5F, &HA1, &HFC, &H8)
Transform2_ZoomMaximum_Property_GUID = iid
End Function
Public Function FlowsFrom_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5C6844F, &H19DE, &H48F8, &H95, &HFA, &H88, &HD, &H5B, &HF, &HD6, &H15)
FlowsFrom_Property_GUID = iid
End Function
Public Function FillColor_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6E0EC4D0, &HE2A8, &H4A56, &H9D, &HE7, &H95, &H33, &H89, &H93, &H3B, &H39)
FillColor_Property_GUID = iid
End Function
Public Function OutlineColor_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC395D6C0, &H4B55, &H4762, &HA0, &H73, &HFD, &H30, &H3A, &H63, &H4F, &H52)
OutlineColor_Property_GUID = iid
End Function
Public Function FillType_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC6FC74E4, &H8CB9, &H429C, &HA9, &HE1, &H9B, &HC4, &HAC, &H37, &H2B, &H62)
FillType_Property_GUID = iid
End Function
Public Function VisualEffects_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE61A8565, &HAAD9, &H46D7, &H9E, &H70, &H4E, &H8A, &H84, &H20, &HD4, &H20)
VisualEffects_Property_GUID = iid
End Function
Public Function OutlineThickness_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H13E67CC7, &HDAC2, &H4888, &HBD, &HD3, &H37, &H5C, &H62, &HFA, &H96, &H18)
OutlineThickness_Property_GUID = iid
End Function
Public Function CenterPoint_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCB00C08, &H540C, &H4EDB, &H94, &H45, &H26, &H35, &H9E, &HA6, &H97, &H85)
CenterPoint_Property_GUID = iid
End Function
Public Function Rotation_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H767CDC7D, &HAEC0, &H4110, &HAD, &H32, &H30, &HED, &HD4, &H3, &H49, &H2E)
Rotation_Property_GUID = iid
End Function
Public Function Size_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2B5F761D, &HF885, &H4404, &H97, &H3F, &H9B, &H1D, &H98, &HE3, &H6D, &H8F)
Size_Property_GUID = iid
End Function
Public Function ToolTipOpened_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3F4B97FF, &H2EDC, &H451D, &HBC, &HA4, &H95, &HA3, &H18, &H8D, &H5B, &H3)
ToolTipOpened_Event_GUID = iid
End Function
Public Function ToolTipClosed_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H276D71EF, &H24A9, &H49B6, &H8E, &H97, &HDA, &H98, &HB4, &H1, &HBB, &HCD)
ToolTipClosed_Event_GUID = iid
End Function
Public Function StructureChanged_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H59977961, &H3EDD, &H4B11, &HB1, &H3B, &H67, &H6B, &H2A, &H2A, &H6C, &HA9)
StructureChanged_Event_GUID = iid
End Function
Public Function MenuOpened_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEBE2E945, &H66CA, &H4ED1, &H9F, &HF8, &H2A, &HD7, &HDF, &HA, &H1B, &H8)
MenuOpened_Event_GUID = iid
End Function
Public Function AutomationPropertyChanged_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2527FBA1, &H8D7A, &H4630, &HA4, &HCC, &HE6, &H63, &H15, &H94, &H2F, &H52)
AutomationPropertyChanged_Event_GUID = iid
End Function
Public Function AutomationFocusChanged_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB68A1F17, &HF60D, &H41A7, &HA3, &HCC, &HB0, &H52, &H92, &H15, &H5F, &HE0)
AutomationFocusChanged_Event_GUID = iid
End Function
Public Function ActiveTextPositionChanged_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA5C09E9C, &HC77D, &H4F25, &HB4, &H91, &HE5, &HBB, &H70, &H17, &HCB, &HD4)
ActiveTextPositionChanged_Event_GUID = iid
End Function
Public Function AsyncContentLoaded_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5FDEE11C, &HD2FA, &H4FB9, &H90, &H4E, &H5C, &HBE, &HE8, &H94, &HD5, &HEF)
AsyncContentLoaded_Event_GUID = iid
End Function
Public Function MenuClosed_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CF1266E, &H1582, &H4041, &HAC, &HD7, &H88, &HA3, &H5A, &H96, &H52, &H97)
MenuClosed_Event_GUID = iid
End Function
Public Function LayoutInvalidated_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HED7D6544, &HA6BD, &H4595, &H9B, &HAE, &H3D, &H28, &H94, &H6C, &HC7, &H15)
LayoutInvalidated_Event_GUID = iid
End Function
Public Function Invoke_Invoked_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDFD699F0, &HC915, &H49DD, &HB4, &H22, &HDD, &HE7, &H85, &HC3, &HD2, &H4B)
Invoke_Invoked_Event_GUID = iid
End Function
Public Function SelectionItem_ElementAddedToSelectionEvent_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3C822DD1, &HC407, &H4DBA, &H91, &HDD, &H79, &HD4, &HAE, &HD0, &HAE, &HC6)
SelectionItem_ElementAddedToSelectionEvent_Event_GUID = iid
End Function
Public Function SelectionItem_ElementRemovedFromSelectionEvent_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H97FA8A9, &H7079, &H41AF, &H8B, &H9C, &H9, &H34, &HD8, &H30, &H5E, &H5C)
SelectionItem_ElementRemovedFromSelectionEvent_Event_GUID = iid
End Function
Public Function SelectionItem_ElementSelectedEvent_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB9C7DBFB, &H4EBE, &H4532, &HAA, &HF4, &H0, &H8C, &HF6, &H47, &H23, &H3C)
SelectionItem_ElementSelectedEvent_Event_GUID = iid
End Function
Public Function Selection_InvalidatedEvent_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCAC14904, &H16B4, &H4B53, &H8E, &H47, &H4C, &HB1, &HDF, &H26, &H7B, &HB7)
Selection_InvalidatedEvent_Event_GUID = iid
End Function
Public Function Text_TextSelectionChangedEvent_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H918EDAA1, &H71B3, &H49AE, &H97, &H41, &H79, &HBE, &HB8, &HD3, &H58, &HF3)
Text_TextSelectionChangedEvent_Event_GUID = iid
End Function
Public Function Text_TextChangedEvent_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4A342082, &HF483, &H48C4, &HAC, &H11, &HA8, &H4B, &H43, &H5E, &H2A, &H84)
Text_TextChangedEvent_Event_GUID = iid
End Function
Public Function Window_WindowOpened_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD3E81D06, &HDE45, &H4F2F, &H96, &H33, &HDE, &H9E, &H2, &HFB, &H65, &HAF)
Window_WindowOpened_Event_GUID = iid
End Function
Public Function Window_WindowClosed_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEDF141F8, &HFA67, &H4E22, &HBB, &HF7, &H94, &H4E, &H5, &H73, &H5E, &HE2)
Window_WindowClosed_Event_GUID = iid
End Function
Public Function MenuModeStart_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H18D7C631, &H166A, &H4AC9, &HAE, &H3B, &HEF, &H4B, &H54, &H20, &HE6, &H81)
MenuModeStart_Event_GUID = iid
End Function
Public Function MenuModeEnd_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9ECD4C9F, &H80DD, &H47B8, &H82, &H67, &H5A, &HEC, &H6, &HBB, &H2C, &HFF)
MenuModeEnd_Event_GUID = iid
End Function
Public Function InputReachedTarget_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H93ED549A, &H549, &H40F0, &HBE, &HDB, &H28, &HE4, &H4F, &H7D, &HE2, &HA3)
InputReachedTarget_Event_GUID = iid
End Function
Public Function InputReachedOtherElement_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HED201D8A, &H4E6C, &H415E, &HA8, &H74, &H24, &H60, &HC9, &HB6, &H6B, &HA8)
InputReachedOtherElement_Event_GUID = iid
End Function
Public Function InputDiscarded_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7F36C367, &H7B18, &H417C, &H97, &HE3, &H9D, &H58, &HDD, &HC9, &H44, &HAB)
InputDiscarded_Event_GUID = iid
End Function
Public Function SystemAlert_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD271545D, &H7A3A, &H47A7, &H84, &H74, &H81, &HD2, &H9A, &H24, &H51, &HC9)
SystemAlert_Event_GUID = iid
End Function
Public Function LiveRegionChanged_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H102D5E90, &HE6A9, &H41B6, &HB1, &HC5, &HA9, &HB1, &H92, &H9D, &H95, &H10)
LiveRegionChanged_Event_GUID = iid
End Function
Public Function HostedFragmentRootsInvalidated_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE6BDB03E, &H921, &H4EC5, &H8D, &HCF, &HEA, &HE8, &H77, &HB0, &H42, &H6B)
HostedFragmentRootsInvalidated_Event_GUID = iid
End Function
Public Function Drag_DragStart_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H883A480B, &H3AA9, &H429D, &H95, &HE4, &HD9, &HC8, &HD0, &H11, &HF0, &HDD)
Drag_DragStart_Event_GUID = iid
End Function
Public Function Drag_DragCancel_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC3EDE6FA, &H3451, &H4E0F, &H9E, &H71, &HDF, &H9C, &H28, &HA, &H46, &H57)
Drag_DragCancel_Event_GUID = iid
End Function
Public Function Drag_DragComplete_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H38E96188, &HEF1F, &H463E, &H91, &HCA, &H3A, &H77, &H92, &HC2, &H9C, &HAF)
Drag_DragComplete_Event_GUID = iid
End Function
Public Function DropTarget_DragEnter_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAAD9319B, &H32C, &H4A88, &H96, &H1D, &H1C, &HF5, &H79, &H58, &H1E, &H34)
DropTarget_DragEnter_Event_GUID = iid
End Function
Public Function DropTarget_DragLeave_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF82EB15, &H24A2, &H4988, &H92, &H17, &HDE, &H16, &H2A, &HEE, &H27, &H2B)
DropTarget_DragLeave_Event_GUID = iid
End Function
Public Function DropTarget_Dropped_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H622CEAD8, &H1EDB, &H4A3D, &HAB, &HBC, &HBE, &H22, &H11, &HFF, &H68, &HB5)
DropTarget_Dropped_Event_GUID = iid
End Function
Public Function StructuredMarkup_CompositionComplete_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC48A3C17, &H677A, &H4047, &HA6, &H8D, &HFC, &H12, &H57, &H52, &H8A, &HEF)
StructuredMarkup_CompositionComplete_Event_GUID = iid
End Function
Public Function StructuredMarkup_Deleted_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF9D0A020, &HE1C1, &H4ECF, &HB9, &HAA, &H52, &HEF, &HDE, &H7E, &H41, &HE1)
StructuredMarkup_Deleted_Event_GUID = iid
End Function
Public Function StructuredMarkup_SelectionChanged_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA7C815F7, &HFF9F, &H41C7, &HA3, &HA7, &HAB, &H6C, &HBF, &HDB, &H49, &H3)
StructuredMarkup_SelectionChanged_Event_GUID = iid
End Function
Public Function Invoke_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD976C2FC, &H66EA, &H4A6E, &HB2, &H8F, &HC2, &H4C, &H75, &H46, &HAD, &H37)
Invoke_Pattern_GUID = iid
End Function
Public Function Selection_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H66E3B7E8, &HD821, &H4D25, &H87, &H61, &H43, &H5D, &H2C, &H8B, &H25, &H3F)
Selection_Pattern_GUID = iid
End Function
Public Function Value_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H17FAAD9E, &HC877, &H475B, &HB9, &H33, &H77, &H33, &H27, &H79, &HB6, &H37)
Value_Pattern_GUID = iid
End Function
Public Function RangeValue_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H18B00D87, &HB1C9, &H476A, &HBF, &HBD, &H5F, &HB, &HDB, &H92, &H6F, &H63)
RangeValue_Pattern_GUID = iid
End Function
Public Function Scroll_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H895FA4B4, &H759D, &H4C50, &H8E, &H15, &H3, &H46, &H6, &H72, &H0, &H3C)
Scroll_Pattern_GUID = iid
End Function
Public Function ExpandCollapse_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAE05EFA2, &HF9D1, &H428A, &H83, &H4C, &H53, &HA5, &HC5, &H2F, &H9B, &H8B)
ExpandCollapse_Pattern_GUID = iid
End Function
Public Function Grid_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H260A2CCB, &H93A8, &H4E44, &HA4, &HC1, &H3D, &HF3, &H97, &HF2, &HB0, &H2B)
Grid_Pattern_GUID = iid
End Function
Public Function GridItem_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF2D5C877, &HA462, &H4957, &HA2, &HA5, &H2C, &H96, &HB3, &H3, &HBC, &H63)
GridItem_Pattern_GUID = iid
End Function
Public Function MultipleView_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H547A6AE4, &H113F, &H47C4, &H85, &HF, &HDB, &H4D, &HFA, &H46, &H6B, &H1D)
MultipleView_Pattern_GUID = iid
End Function
Public Function Window_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H27901735, &HC760, &H4994, &HAD, &H11, &H59, &H19, &HE6, &H6, &HB1, &H10)
Window_Pattern_GUID = iid
End Function
Public Function SelectionItem_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9BC64EEB, &H87C7, &H4B28, &H94, &HBB, &H4D, &H9F, &HA4, &H37, &HB6, &HEF)
SelectionItem_Pattern_GUID = iid
End Function
Public Function Dock_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9CBAA846, &H83C8, &H428D, &H82, &H7F, &H7E, &H60, &H63, &HFE, &H6, &H20)
Dock_Pattern_GUID = iid
End Function
Public Function Table_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC415218E, &HA028, &H461E, &HAA, &H92, &H8F, &H92, &H5C, &HF7, &H93, &H51)
Table_Pattern_GUID = iid
End Function
Public Function TableItem_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDF1343BD, &H1888, &H4A29, &HA5, &HC, &HB9, &H2E, &H6D, &HE3, &H7F, &H6F)
TableItem_Pattern_GUID = iid
End Function
Public Function Text_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8615F05D, &H7DE5, &H44FD, &HA6, &H79, &H2C, &HA4, &HB4, &H60, &H33, &HA8)
Text_Pattern_GUID = iid
End Function
Public Function Toggle_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB419760, &HE2F4, &H43FF, &H8C, &H5F, &H94, &H57, &HC8, &H2B, &H56, &HE9)
Toggle_Pattern_GUID = iid
End Function
Public Function Transform_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H24B46FDB, &H587E, &H49F1, &H9C, &H4A, &HD8, &HE9, &H8B, &H66, &H4B, &H7B)
Transform_Pattern_GUID = iid
End Function
Public Function ScrollItem_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4591D005, &HA803, &H4D5C, &HB4, &HD5, &H8D, &H28, &H0, &HF9, &H6, &HA7)
ScrollItem_Pattern_GUID = iid
End Function
Public Function LegacyIAccessible_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H54CC0A9F, &H3395, &H48AF, &HBA, &H8D, &H73, &HF8, &H56, &H90, &HF3, &HE0)
LegacyIAccessible_Pattern_GUID = iid
End Function
Public Function ItemContainer_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3D13DA0F, &H8B9A, &H4A99, &H85, &HFA, &HC5, &HC9, &HA6, &H9F, &H1E, &HD4)
ItemContainer_Pattern_GUID = iid
End Function
Public Function VirtualizedItem_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF510173E, &H2E71, &H45E9, &HA6, &HE5, &H62, &HF6, &HED, &H82, &H89, &HD5)
VirtualizedItem_Pattern_GUID = iid
End Function
Public Function SynchronizedInput_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5C288A6, &HC47B, &H488B, &HB6, &H53, &H33, &H97, &H7A, &H55, &H1B, &H8B)
SynchronizedInput_Pattern_GUID = iid
End Function
Public Function ObjectModel_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3E04ACFE, &H8FC, &H47EC, &H96, &HBC, &H35, &H3F, &HA3, &HB3, &H4A, &HA7)
ObjectModel_Pattern_GUID = iid
End Function
Public Function Annotation_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF6C72AD7, &H356C, &H4850, &H92, &H91, &H31, &H6F, &H60, &H8A, &H8C, &H84)
Annotation_Pattern_GUID = iid
End Function
Public Function Text_Pattern2_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H498479A2, &H5B22, &H448D, &HB6, &HE4, &H64, &H74, &H90, &H86, &H6, &H98)
Text_Pattern2_GUID = iid
End Function
Public Function TextEdit_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H69F3FF89, &H5AF9, &H4C75, &H93, &H40, &HF2, &HDE, &H29, &H2E, &H45, &H91)
TextEdit_Pattern_GUID = iid
End Function
Public Function CustomNavigation_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAFEA938A, &H621E, &H4054, &HBB, &H2C, &H2F, &H46, &H11, &H4D, &HAC, &H3F)
CustomNavigation_Pattern_GUID = iid
End Function
Public Function Styles_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1AE62655, &HDA72, &H4D60, &HA1, &H53, &HE5, &HAA, &H69, &H88, &HE3, &HBF)
Styles_Pattern_GUID = iid
End Function
Public Function Spreadsheet_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6A5B24C9, &H9D1E, &H4B85, &H9E, &H44, &HC0, &H2E, &H31, &H69, &HB1, &HB)
Spreadsheet_Pattern_GUID = iid
End Function
Public Function SpreadsheetItem_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H32CF83FF, &HF1A8, &H4A8C, &H86, &H58, &HD4, &H7B, &HA7, &H4E, &H20, &HBA)
SpreadsheetItem_Pattern_GUID = iid
End Function
Public Function Tranform_Pattern2_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8AFCFD07, &HA369, &H44DE, &H98, &H8B, &H2F, &H7F, &HF4, &H9F, &HB8, &HA8)
Tranform_Pattern2_GUID = iid
End Function
Public Function TextChild_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7533CAB7, &H3BFE, &H41EF, &H9E, &H85, &HE2, &H63, &H8C, &HBE, &H16, &H9E)
TextChild_Pattern_GUID = iid
End Function
Public Function Drag_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC0BEE21F, &HCCB3, &H4FED, &H99, &H5B, &H11, &H4F, &H6E, &H3D, &H27, &H28)
Drag_Pattern_GUID = iid
End Function
Public Function DropTarget_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBCBEC56, &HBD34, &H4B7B, &H9F, &HD5, &H26, &H59, &H90, &H5E, &HA3, &HDC)
DropTarget_Pattern_GUID = iid
End Function
Public Function StructuredMarkup_Pattern_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HABBD0878, &H8665, &H4F5C, &H94, &HFC, &H36, &HE7, &HD8, &HBB, &H70, &H6B)
StructuredMarkup_Pattern_GUID = iid
End Function
Public Function Button_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5A78E369, &HC6A1, &H4F33, &HA9, &HD7, &H79, &HF2, &HD, &HC, &H78, &H8E)
Button_Control_GUID = iid
End Function
Public Function Calendar_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8913EB88, &HE5, &H46BC, &H8E, &H4E, &H14, &HA7, &H86, &HE1, &H65, &HA1)
Calendar_Control_GUID = iid
End Function
Public Function CheckBox_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFB50F922, &HA3DB, &H49C0, &H8B, &HC3, &H6, &HDA, &HD5, &H57, &H78, &HE2)
CheckBox_Control_GUID = iid
End Function
Public Function ComboBox_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H54CB426C, &H2F33, &H4FFF, &HAA, &HA1, &HAE, &HF6, &HD, &HAC, &H5D, &HEB)
ComboBox_Control_GUID = iid
End Function
Public Function Edit_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6504A5C8, &H2C86, &H4F87, &HAE, &H7B, &H1A, &HBD, &HDC, &H81, &HC, &HF9)
Edit_Control_GUID = iid
End Function
Public Function Hyperlink_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8A56022C, &HB00D, &H4D15, &H8F, &HF0, &H5B, &H6B, &H26, &H6E, &H5E, &H2)
Hyperlink_Control_GUID = iid
End Function
Public Function Image_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2D3736E4, &H6B16, &H4C57, &HA9, &H62, &HF9, &H32, &H60, &HA7, &H52, &H43)
Image_Control_GUID = iid
End Function
Public Function ListItem_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7B3717F2, &H44D1, &H4A58, &H98, &HA8, &HF1, &H2A, &H9B, &H8F, &H78, &HE2)
ListItem_Control_GUID = iid
End Function
Public Function List_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9B149EE1, &H7CCA, &H4CFC, &H9A, &HF1, &HCA, &HC7, &HBD, &HDD, &H30, &H31)
List_Control_GUID = iid
End Function
Public Function Menu_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2E9B1440, &HEA8, &H41FD, &HB3, &H74, &HC1, &HEA, &H6F, &H50, &H3C, &HD1)
Menu_Control_GUID = iid
End Function
Public Function MenuBar_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCC384250, &HE7B, &H4AE8, &H95, &HAE, &HA0, &H8F, &H26, &H1B, &H52, &HEE)
MenuBar_Control_GUID = iid
End Function
Public Function MenuItem_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF45225D3, &HD0A0, &H49D8, &H98, &H34, &H9A, &H0, &HD, &H2A, &HED, &HDC)
MenuItem_Control_GUID = iid
End Function
Public Function ProgressBar_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H228C9F86, &HC36C, &H47BB, &H9F, &HB6, &HA5, &H83, &H4B, &HFC, &H53, &HA4)
ProgressBar_Control_GUID = iid
End Function
Public Function RadioButton_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3BDB49DB, &HFE2C, &H4483, &HB3, &HE1, &HE5, &H7F, &H21, &H94, &H40, &HC6)
RadioButton_Control_GUID = iid
End Function
Public Function ScrollBar_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDAF34B36, &H5065, &H4946, &HB2, &H2F, &H92, &H59, &H5F, &HC0, &H75, &H1A)
ScrollBar_Control_GUID = iid
End Function
Public Function Slider_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB033C24B, &H3B35, &H4CEA, &HB6, &H9, &H76, &H36, &H82, &HFA, &H66, &HB)
Slider_Control_GUID = iid
End Function
Public Function Spinner_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H60CC4B38, &H3CB1, &H4161, &HB4, &H42, &HC6, &HB7, &H26, &HC1, &H78, &H25)
Spinner_Control_GUID = iid
End Function
Public Function StatusBar_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD45E7D1B, &H5873, &H475F, &H95, &HA4, &H4, &H33, &HE1, &HF1, &HB0, &HA)
StatusBar_Control_GUID = iid
End Function
Public Function Tab_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H38CD1F2D, &H337A, &H4BD2, &HA5, &HE3, &HAD, &HB4, &H69, &HE3, &HB, &HD3)
Tab_Control_GUID = iid
End Function
Public Function TabItem_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2C6A634F, &H921B, &H4E6E, &HB2, &H6E, &H8, &HFC, &HB0, &H79, &H8F, &H4C)
TabItem_Control_GUID = iid
End Function
Public Function Text_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAE9772DC, &HD331, &H4F09, &HBE, &H20, &H7E, &H6D, &HFA, &HF0, &H7B, &HA)
Text_Control_GUID = iid
End Function
Public Function ToolBar_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8F06B751, &HE182, &H4E98, &H88, &H93, &H22, &H84, &H54, &H3A, &H7D, &HCE)
ToolBar_Control_GUID = iid
End Function
Public Function ToolTip_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5DDC6D1, &H2137, &H4768, &H98, &HEA, &H73, &HF5, &H2F, &H71, &H34, &HF3)
ToolTip_Control_GUID = iid
End Function
Public Function Tree_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7561349C, &HD241, &H43F4, &H99, &H8, &HB5, &HF0, &H91, &HBE, &HE6, &H11)
Tree_Control_GUID = iid
End Function
Public Function TreeItem_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H62C9FEB9, &H8FFC, &H4878, &HA3, &HA4, &H96, &HB0, &H30, &H31, &H5C, &H18)
TreeItem_Control_GUID = iid
End Function
Public Function Custom_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF29EA0C3, &HADB7, &H430A, &HBA, &H90, &HE5, &H2C, &H73, &H13, &HE6, &HED)
Custom_Control_GUID = iid
End Function
Public Function Group_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAD50AA1C, &HE8C8, &H4774, &HAE, &H1B, &HDD, &H86, &HDF, &HB, &H3B, &HDC)
Group_Control_GUID = iid
End Function
Public Function Thumb_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H701CA877, &HE310, &H4DD6, &HB6, &H44, &H79, &H7E, &H4F, &HAE, &HA2, &H13)
Thumb_Control_GUID = iid
End Function
Public Function DataGrid_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H84B783AF, &HD103, &H4B0A, &H84, &H15, &HE7, &H39, &H42, &H41, &HF, &H4B)
DataGrid_Control_GUID = iid
End Function
Public Function DataItem_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA0177842, &HD94F, &H42A5, &H81, &H4B, &H60, &H68, &HAD, &HDC, &H8D, &HA5)
DataItem_Control_GUID = iid
End Function
Public Function Document_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3CD6BB6F, &H6F08, &H4562, &HB2, &H29, &HE4, &HE2, &HFC, &H7A, &H9E, &HB4)
Document_Control_GUID = iid
End Function
Public Function SplitButton_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7011F01F, &H4ACE, &H4901, &HB4, &H61, &H92, &HA, &H6F, &H1C, &HA6, &H50)
SplitButton_Control_GUID = iid
End Function
Public Function Window_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE13A7242, &HF462, &H4F4D, &HAE, &HC1, &H53, &HB2, &H8D, &H6C, &H32, &H90)
Window_Control_GUID = iid
End Function
Public Function Pane_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5C2B3F5B, &H9182, &H42A3, &H8D, &HEC, &H8C, &H4, &HC1, &HEE, &H63, &H4D)
Pane_Control_GUID = iid
End Function
Public Function Header_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5B90CBCE, &H78FB, &H4614, &H82, &HB6, &H55, &H4D, &H74, &H71, &H8E, &H67)
Header_Control_GUID = iid
End Function
Public Function HeaderItem_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE6BC12CB, &H7C8E, &H49CF, &HB1, &H68, &H4A, &H93, &HA3, &H2B, &HEB, &HB0)
HeaderItem_Control_GUID = iid
End Function
Public Function Table_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H773BFA0E, &H5BC4, &H4DEB, &H92, &H1B, &HDE, &H7B, &H32, &H6, &H22, &H9E)
Table_Control_GUID = iid
End Function
Public Function TitleBar_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H98AA55BF, &H3BB0, &H4B65, &H83, &H6E, &H2E, &HA3, &HD, &HBC, &H17, &H1F)
TitleBar_Control_GUID = iid
End Function
Public Function Separator_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8767EBA3, &H2A63, &H4AB0, &HAC, &H8D, &HAA, &H50, &HE2, &H3D, &HE9, &H78)
Separator_Control_GUID = iid
End Function
Public Function SemanticZoom_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5FD34A43, &H61E, &H42C8, &HB5, &H89, &H9D, &HCC, &HF7, &H4B, &HC4, &H3A)
SemanticZoom_Control_GUID = iid
End Function
Public Function AppBar_Control_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6114908D, &HCC02, &H4D37, &H87, &H5B, &HB5, &H30, &HC7, &H13, &H95, &H54)
AppBar_Control_GUID = iid
End Function
Public Function Text_AnimationStyle_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H628209F0, &H7C9A, &H4D57, &HBE, &H64, &H1F, &H18, &H36, &H57, &H1F, &HF5)
Text_AnimationStyle_Attribute_GUID = iid
End Function
Public Function Text_BackgroundColor_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFDC49A07, &H583D, &H4F17, &HAD, &H27, &H77, &HFC, &H83, &H2A, &H3C, &HB)
Text_BackgroundColor_Attribute_GUID = iid
End Function
Public Function Text_BulletStyle_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC1097C90, &HD5C4, &H4237, &H97, &H81, &H3B, &HEC, &H8B, &HA5, &H4E, &H48)
Text_BulletStyle_Attribute_GUID = iid
End Function
Public Function Text_CapStyle_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFB059C50, &H92CC, &H49A5, &HBA, &H8F, &HA, &HA8, &H72, &HBB, &HA2, &HF3)
Text_CapStyle_Attribute_GUID = iid
End Function
Public Function Text_Culture_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC2025AF9, &HA42D, &H4CED, &HA1, &HFB, &HC6, &H74, &H63, &H15, &H22, &H2E)
Text_Culture_Attribute_GUID = iid
End Function
Public Function Text_FontName_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H64E63BA8, &HF2E5, &H476E, &HA4, &H77, &H17, &H34, &HFE, &HAA, &HF7, &H26)
Text_FontName_Attribute_GUID = iid
End Function
Public Function Text_FontSize_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDC5EEEFF, &H506, &H4673, &H93, &HF2, &H37, &H7E, &H4A, &H8E, &H1, &HF1)
Text_FontSize_Attribute_GUID = iid
End Function
Public Function Text_FontWeight_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6FC02359, &HB316, &H4F5F, &HB4, &H1, &HF1, &HCE, &H55, &H74, &H18, &H53)
Text_FontWeight_Attribute_GUID = iid
End Function
Public Function Text_ForegroundColor_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H72D1C95D, &H5E60, &H471A, &H96, &HB1, &H6C, &H1B, &H3B, &H77, &HA4, &H36)
Text_ForegroundColor_Attribute_GUID = iid
End Function
Public Function Text_HorizontalTextAlignment_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4EA6161, &HFBA3, &H477A, &H95, &H2A, &HBB, &H32, &H6D, &H2, &H6A, &H5B)
Text_HorizontalTextAlignment_Attribute_GUID = iid
End Function
Public Function Text_IndentationFirstLine_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H206F9AD5, &HC1D3, &H424A, &H81, &H82, &H6D, &HA9, &HA7, &HF3, &HD6, &H32)
Text_IndentationFirstLine_Attribute_GUID = iid
End Function
Public Function Text_IndentationLeading_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5CF66BAC, &H2D45, &H4A4B, &HB6, &HC9, &HF7, &H22, &H1D, &H28, &H15, &HB0)
Text_IndentationLeading_Attribute_GUID = iid
End Function
Public Function Text_IndentationTrailing_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H97FF6C0F, &H1CE4, &H408A, &HB6, &H7B, &H94, &HD8, &H3E, &HB6, &H9B, &HF2)
Text_IndentationTrailing_Attribute_GUID = iid
End Function
Public Function Text_IsHidden_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H360182FB, &HBDD7, &H47F6, &HAB, &H69, &H19, &HE3, &H3F, &H8A, &H33, &H44)
Text_IsHidden_Attribute_GUID = iid
End Function
Public Function Text_IsItalic_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFCE12A56, &H1336, &H4A34, &H96, &H63, &H1B, &HAB, &H47, &H23, &H93, &H20)
Text_IsItalic_Attribute_GUID = iid
End Function
Public Function Text_IsReadOnly_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA738156B, &HCA3E, &H495E, &H95, &H14, &H83, &H3C, &H44, &HF, &HEB, &H11)
Text_IsReadOnly_Attribute_GUID = iid
End Function
Public Function Text_IsSubscript_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF0EAD858, &H8F53, &H413C, &H87, &H3F, &H1A, &H7D, &H7F, &H5E, &HD, &HE4)
Text_IsSubscript_Attribute_GUID = iid
End Function
Public Function Text_IsSuperscript_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDA706EE4, &HB3AA, &H4645, &HA4, &H1F, &HCD, &H25, &H15, &H7D, &HEA, &H76)
Text_IsSuperscript_Attribute_GUID = iid
End Function
Public Function Text_MarginBottom_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7EE593C4, &H72B4, &H4CAC, &H92, &H71, &H3E, &HD2, &H4B, &HE, &H4D, &H42)
Text_MarginBottom_Attribute_GUID = iid
End Function
Public Function Text_MarginLeading_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9E9242D0, &H5ED0, &H4900, &H8E, &H8A, &HEE, &HCC, &H3, &H83, &H5A, &HFC)
Text_MarginLeading_Attribute_GUID = iid
End Function
Public Function Text_MarginTop_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H683D936F, &HC9B9, &H4A9A, &HB3, &HD9, &HD2, &HD, &H33, &H31, &H1E, &H2A)
Text_MarginTop_Attribute_GUID = iid
End Function
Public Function Text_MarginTrailing_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAF522F98, &H999D, &H40AF, &HA5, &HB2, &H1, &H69, &HD0, &H34, &H20, &H2)
Text_MarginTrailing_Attribute_GUID = iid
End Function
Public Function Text_OutlineStyles_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5B675B27, &HDB89, &H46FE, &H97, &HC, &H61, &H4D, &H52, &H3B, &HB9, &H7D)
Text_OutlineStyles_Attribute_GUID = iid
End Function
Public Function Text_OverlineColor_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H83AB383A, &HFD43, &H40DA, &HAB, &H3E, &HEC, &HF8, &H16, &H5C, &HBB, &H6D)
Text_OverlineColor_Attribute_GUID = iid
End Function
Public Function Text_OverlineStyle_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA234D66, &H617E, &H427F, &H87, &H1D, &HE1, &HFF, &H1E, &HC, &H21, &H3F)
Text_OverlineStyle_Attribute_GUID = iid
End Function
Public Function Text_StrikethroughColor_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBFE15A18, &H8C41, &H4C5A, &H9A, &HB, &H4, &HAF, &HE, &H7, &HF4, &H87)
Text_StrikethroughColor_Attribute_GUID = iid
End Function
Public Function Text_StrikethroughStyle_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H72913EF1, &HDA00, &H4F01, &H89, &H9C, &HAC, &H5A, &H85, &H77, &HA3, &H7)
Text_StrikethroughStyle_Attribute_GUID = iid
End Function
Public Function Text_Tabs_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2E68D00B, &H92FE, &H42D8, &H89, &H9A, &HA7, &H84, &HAA, &H44, &H54, &HA1)
Text_Tabs_Attribute_GUID = iid
End Function
Public Function Text_TextFlowDirections_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8BDF8739, &HF420, &H423E, &HAF, &H77, &H20, &HA5, &HD9, &H73, &HA9, &H7)
Text_TextFlowDirections_Attribute_GUID = iid
End Function
Public Function Text_UnderlineColor_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBFA12C73, &HFDE2, &H4473, &HBF, &H64, &H10, &H36, &HD6, &HAA, &HF, &H45)
Text_UnderlineColor_Attribute_GUID = iid
End Function
Public Function Text_UnderlineStyle_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5F3B21C0, &HEDE4, &H44BD, &H9C, &H36, &H38, &H53, &H3, &H8C, &HBF, &HEB)
Text_UnderlineStyle_Attribute_GUID = iid
End Function
Public Function Text_AnnotationTypes_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAD2EB431, &HEE4E, &H4BE1, &HA7, &HBA, &H55, &H59, &H15, &H5A, &H73, &HEF)
Text_AnnotationTypes_Attribute_GUID = iid
End Function
Public Function Text_AnnotationObjects_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFF41CF68, &HE7AB, &H40B9, &H8C, &H72, &H72, &HA8, &HED, &H94, &H1, &H7D)
Text_AnnotationObjects_Attribute_GUID = iid
End Function
Public Function Text_StyleName_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H22C9E091, &H4D66, &H45D8, &HA8, &H28, &H73, &H7B, &HAB, &H4C, &H98, &HA7)
Text_StyleName_Attribute_GUID = iid
End Function
Public Function Text_StyleId_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H14C300DE, &HC32B, &H449B, &HAB, &H7C, &HB0, &HE0, &H78, &H9A, &HEA, &H5D)
Text_StyleId_Attribute_GUID = iid
End Function
Public Function Text_Link_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB38EF51D, &H9E8D, &H4E46, &H91, &H44, &H56, &HEB, &HE1, &H77, &H32, &H9B)
Text_Link_Attribute_GUID = iid
End Function
Public Function Text_IsActive_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF5A4E533, &HE1B8, &H436B, &H93, &H5D, &HB5, &H7A, &HA3, &HF5, &H58, &HC4)
Text_IsActive_Attribute_GUID = iid
End Function
Public Function Text_SelectionActiveEnd_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1F668CC3, &H9BBF, &H416B, &HB0, &HA2, &HF8, &H9F, &H86, &HF6, &H61, &H2C)
Text_SelectionActiveEnd_Attribute_GUID = iid
End Function
Public Function Text_CaretPosition_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB227B131, &H9889, &H4752, &HA9, &H1B, &H73, &H3E, &HFD, &HC5, &HC5, &HA0)
Text_CaretPosition_Attribute_GUID = iid
End Function
Public Function Text_CaretBidiMode_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H929EE7A6, &H51D3, &H4715, &H96, &HDC, &HB6, &H94, &HFA, &H24, &HA1, &H68)
Text_CaretBidiMode_Attribute_GUID = iid
End Function
Public Function Text_BeforeParagraphSpacing_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBE7B0AB1, &HC822, &H4A24, &H85, &HE9, &HC8, &HF2, &H65, &HF, &HC7, &H9C)
Text_BeforeParagraphSpacing_Attribute_GUID = iid
End Function
Public Function Text_AfterParagraphSpacing_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H588CBB38, &HE62F, &H497C, &HB5, &HD1, &HCC, &HDF, &HE, &HE8, &H23, &HD8)
Text_AfterParagraphSpacing_Attribute_GUID = iid
End Function
Public Function Text_LineSpacing_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H63FF70AE, &HD943, &H4B47, &H8A, &HB7, &HA7, &HA0, &H33, &HD3, &H21, &H4B)
Text_LineSpacing_Attribute_GUID = iid
End Function
Public Function Text_BeforeSpacing_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBE7B0AB1, &HC822, &H4A24, &H85, &HE9, &HC8, &HF2, &H65, &HF, &HC7, &H9C)
Text_BeforeSpacing_Attribute_GUID = iid
End Function
Public Function Text_AfterSpacing_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H588CBB38, &HE62F, &H497C, &HB5, &HD1, &HCC, &HDF, &HE, &HE8, &H23, &HD8)
Text_AfterSpacing_Attribute_GUID = iid
End Function
Public Function Text_SayAsInterpretAs_Attribute_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB38AD6AC, &HEEE1, &H4B6E, &H88, &HCC, &H1, &H4C, &HEF, &HA9, &H3F, &HCB)
Text_SayAsInterpretAs_Attribute_GUID = iid
End Function
Public Function TextEdit_TextChanged_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H120B0308, &HEC22, &H4EB8, &H9C, &H98, &H98, &H67, &HCD, &HA1, &HB1, &H65)
TextEdit_TextChanged_Event_GUID = iid
End Function
Public Function TextEdit_ConversionTargetChanged_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3388C183, &HED4F, &H4C8B, &H9B, &HAA, &H36, &H4D, &H51, &HD8, &H84, &H7F)
TextEdit_ConversionTargetChanged_Event_GUID = iid
End Function
Public Function Changes_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7DF26714, &H614F, &H4E05, &H94, &H88, &H71, &H6C, &H5B, &HA1, &H94, &H36)
Changes_Event_GUID = iid
End Function
Public Function Annotation_Custom_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9EC82750, &H3931, &H4952, &H85, &HBC, &H1D, &HBF, &HF7, &H8A, &H43, &HE3)
Annotation_Custom_GUID = iid
End Function
Public Function Annotation_SpellingError_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAE85567E, &H9ECE, &H423F, &H81, &HB7, &H96, &HC4, &H3D, &H53, &HE5, &HE)
Annotation_SpellingError_GUID = iid
End Function
Public Function Annotation_GrammarError_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H757A048D, &H4518, &H41C6, &H85, &H4C, &HDC, &H0, &H9B, &H7C, &HFB, &H53)
Annotation_GrammarError_GUID = iid
End Function
Public Function Annotation_Comment_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFD2FDA30, &H26B3, &H4C06, &H8B, &HC7, &H98, &HF1, &H53, &H2E, &H46, &HFD)
Annotation_Comment_GUID = iid
End Function
Public Function Annotation_FormulaError_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H95611982, &HCAB, &H46D5, &HA2, &HF0, &HE3, &HD, &H19, &H5, &HF8, &HBF)
Annotation_FormulaError_GUID = iid
End Function
Public Function Annotation_TrackChanges_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H21E6E888, &HDC14, &H4016, &HAC, &H27, &H19, &H5, &H53, &HC8, &HC4, &H70)
Annotation_TrackChanges_GUID = iid
End Function
Public Function Annotation_Header_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H867B409B, &HB216, &H4472, &HA2, &H19, &H52, &H5E, &H31, &H6, &H81, &HF8)
Annotation_Header_GUID = iid
End Function
Public Function Annotation_Footer_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCCEAB046, &H1833, &H47AA, &H80, &H80, &H70, &H1E, &HD0, &HB0, &HC8, &H32)
Annotation_Footer_GUID = iid
End Function
Public Function Annotation_Highlighted_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H757C884E, &H8083, &H4081, &H8B, &H9C, &HE8, &H7F, &H50, &H72, &HF0, &HE4)
Annotation_Highlighted_GUID = iid
End Function
Public Function Annotation_Endnote_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7565725C, &H2D99, &H4839, &H96, &HD, &H33, &HD3, &HB8, &H66, &HAB, &HA5)
Annotation_Endnote_GUID = iid
End Function
Public Function Annotation_Footnote_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3DE10E21, &H4125, &H42DB, &H86, &H20, &HBE, &H80, &H83, &H8, &H6, &H24)
Annotation_Footnote_GUID = iid
End Function
Public Function Annotation_InsertionChange_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDBEB3A6, &HDF15, &H4164, &HA3, &HC0, &HE2, &H1A, &H8C, &HE9, &H31, &HC4)
Annotation_InsertionChange_GUID = iid
End Function
Public Function Annotation_DeletionChange_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBE3D5B05, &H951D, &H42E7, &H90, &H1D, &HAD, &HC8, &HC2, &HCF, &H34, &HD0)
Annotation_DeletionChange_GUID = iid
End Function
Public Function Annotation_MoveChange_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9DA587EB, &H23E5, &H4490, &HB3, &H85, &H1A, &H22, &HDD, &HC8, &HB1, &H87)
Annotation_MoveChange_GUID = iid
End Function
Public Function Annotation_FormatChange_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEB247345, &HD4F1, &H41CE, &H8E, &H52, &HF7, &H9B, &H69, &H63, &H5E, &H48)
Annotation_FormatChange_GUID = iid
End Function
Public Function Annotation_UnsyncedChange_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1851116A, &HE47, &H4B30, &H8C, &HB5, &HD7, &HDA, &HE4, &HFB, &HCD, &H1B)
Annotation_UnsyncedChange_GUID = iid
End Function
Public Function Annotation_EditingLockedChange_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC31F3E1C, &H7423, &H4DAC, &H83, &H48, &H41, &HF0, &H99, &HFF, &H6F, &H64)
Annotation_EditingLockedChange_GUID = iid
End Function
Public Function Annotation_ExternalChange_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H75A05B31, &H5F11, &H42FD, &H88, &H7D, &HDF, &HA0, &H10, &HDB, &H23, &H92)
Annotation_ExternalChange_GUID = iid
End Function
Public Function Annotation_ConflictingChange_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H98AF8802, &H517C, &H459F, &HAF, &H13, &H1, &H6D, &H3F, &HAB, &H87, &H7E)
Annotation_ConflictingChange_GUID = iid
End Function
Public Function Annotation_Author_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF161D3A7, &HF81B, &H4128, &HB1, &H7F, &H71, &HF6, &H90, &H91, &H45, &H20)
Annotation_Author_GUID = iid
End Function
Public Function Annotation_AdvancedProofingIssue_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDAC7B72C, &HC0F2, &H4B84, &HB9, &HD, &H5F, &HAF, &HC0, &HF0, &HEF, &H1C)
Annotation_AdvancedProofingIssue_GUID = iid
End Function
Public Function Annotation_DataValidationError_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC8649FA8, &H9775, &H437E, &HAD, &H46, &HE7, &H9, &HD9, &H3C, &H23, &H43)
Annotation_DataValidationError_GUID = iid
End Function
Public Function Annotation_CircularReferenceError_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H25BD9CF4, &H1745, &H4659, &HBA, &H67, &H72, &H7F, &H3, &H18, &HC6, &H16)
Annotation_CircularReferenceError_GUID = iid
End Function
Public Function Annotation_Mathematics_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEAAB634B, &H26D0, &H40C1, &H80, &H73, &H57, &HCA, &H1C, &H63, &H3C, &H9B)
Annotation_Mathematics_GUID = iid
End Function
Public Function Annotation_Sensitive_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H37F4C04F, &HF12, &H4464, &H92, &H9C, &H82, &H8F, &HD1, &H52, &H92, &HE3)
Annotation_Sensitive_GUID = iid
End Function
Public Function Changes_Summary_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H313D65A6, &HE60F, &H4D62, &H98, &H61, &H55, &HAF, &HD7, &H28, &HD2, &H7)
Changes_Summary_GUID = iid
End Function
Public Function StyleId_Custom_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEF2EDD3E, &HA999, &H4B7C, &HA3, &H78, &H9, &HBB, &HD5, &H2A, &H35, &H16)
StyleId_Custom_GUID = iid
End Function
Public Function StyleId_Heading1_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7F7E8F69, &H6866, &H4621, &H93, &HC, &H9A, &H5D, &HC, &HA5, &H96, &H1C)
StyleId_Heading1_GUID = iid
End Function
Public Function StyleId_Heading2_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBAA9B241, &H5C69, &H469D, &H85, &HAD, &H47, &H47, &H37, &HB5, &H2B, &H14)
StyleId_Heading2_GUID = iid
End Function
Public Function StyleId_Heading3_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBF8BE9D2, &HD8B8, &H4EC5, &H8C, &H52, &H9C, &HFB, &HD, &H3, &H59, &H70)
StyleId_Heading3_GUID = iid
End Function
Public Function StyleId_Heading4_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8436FFC0, &H9578, &H45FC, &H83, &HA4, &HFF, &H40, &H5, &H33, &H15, &HDD)
StyleId_Heading4_GUID = iid
End Function
Public Function StyleId_Heading5_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H909F424D, &HDBF, &H406E, &H97, &HBB, &H4E, &H77, &H3D, &H97, &H98, &HF7)
StyleId_Heading5_GUID = iid
End Function
Public Function StyleId_Heading6_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H89D23459, &H5D5B, &H4824, &HA4, &H20, &H11, &HD3, &HED, &H82, &HE4, &HF)
StyleId_Heading6_GUID = iid
End Function
Public Function StyleId_Heading7_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA3790473, &HE9AE, &H422D, &HB8, &HE3, &H3B, &H67, &H5C, &H61, &H81, &HA4)
StyleId_Heading7_GUID = iid
End Function
Public Function StyleId_Heading8_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2BC14145, &HA40C, &H4881, &H84, &HAE, &HF2, &H23, &H56, &H85, &H38, &HC)
StyleId_Heading8_GUID = iid
End Function
Public Function StyleId_Heading9_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC70D9133, &HBB2A, &H43D3, &H8A, &HC6, &H33, &H65, &H78, &H84, &HB0, &HF0)
StyleId_Heading9_GUID = iid
End Function
Public Function StyleId_Title_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H15D8201A, &HFFCF, &H481F, &HB0, &HA1, &H30, &HB6, &H3B, &HE9, &H8F, &H7)
StyleId_Title_GUID = iid
End Function
Public Function StyleId_Subtitle_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB5D9FC17, &H5D6F, &H4420, &HB4, &H39, &H7C, &HB1, &H9A, &HD4, &H34, &HE2)
StyleId_Subtitle_GUID = iid
End Function
Public Function StyleId_Normal_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCD14D429, &HE45E, &H4475, &HA1, &HC5, &H7F, &H9E, &H6B, &HE9, &H6E, &HBA)
StyleId_Normal_GUID = iid
End Function
Public Function StyleId_Emphasis_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCA6E7DBE, &H355E, &H4820, &H95, &HA0, &H92, &H5F, &H4, &H1D, &H34, &H70)
StyleId_Emphasis_GUID = iid
End Function
Public Function StyleId_Quote_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5D1C21EA, &H8195, &H4F6C, &H87, &HEA, &H5D, &HAB, &HEC, &HE6, &H4C, &H1D)
StyleId_Quote_GUID = iid
End Function
Public Function StyleId_BulletedList_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5963ED64, &H6426, &H4632, &H8C, &HAF, &HA3, &H2A, &HD4, &H2, &HD9, &H1A)
StyleId_BulletedList_GUID = iid
End Function
Public Function StyleId_NumberedList_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1E96DBD5, &H64C3, &H43D0, &HB1, &HEE, &HB5, &H3B, &H6, &HE3, &HED, &HDF)
StyleId_NumberedList_GUID = iid
End Function
Public Function Notification_Event_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H72C5A2F7, &H9788, &H480F, &HB8, &HEB, &H4D, &HEE, &H0, &HF6, &H18, &H6F)
Notification_Event_GUID = iid
End Function
Public Function SID_IsUIAutomationObject() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB96FDB85, &H7204, &H4724, &H84, &H2B, &HC7, &H5, &H9D, &HED, &HB9, &HD0)
SID_IsUIAutomationObject = iid
End Function
Public Function SID_ControlElementProvider() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF4791D68, &HE254, &H4BA3, &H9A, &H53, &H26, &HA5, &HC5, &H49, &H79, &H46)
SID_ControlElementProvider = iid
End Function
Public Function IsSelectionPattern2Available_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H490806FB, &H6E89, &H4A47, &H83, &H19, &HD2, &H66, &HE5, &H11, &HF0, &H21)
IsSelectionPattern2Available_Property_GUID = iid
End Function
Public Function Selection2_FirstSelectedItem_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCC24EA67, &H369C, &H4E55, &H9F, &HF7, &H38, &HDA, &H69, &H54, &HC, &H29)
Selection2_FirstSelectedItem_Property_GUID = iid
End Function
Public Function Selection2_LastSelectedItem_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCF7BDA90, &H2D83, &H49F8, &H86, &HC, &H9C, &HE3, &H94, &HCF, &H89, &HB4)
Selection2_LastSelectedItem_Property_GUID = iid
End Function
Public Function Selection2_CurrentSelectedItem_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H34257C26, &H83B5, &H41A6, &H93, &H9C, &HAE, &H84, &H1C, &H13, &H62, &H36)
Selection2_CurrentSelectedItem_Property_GUID = iid
End Function
Public Function Selection2_ItemCount_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBB49EB9F, &H456D, &H4048, &HB5, &H91, &H9C, &H20, &H26, &HB8, &H46, &H36)
Selection2_ItemCount_Property_GUID = iid
End Function
Public Function Selection_Pattern2_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFBA25CAB, &HAB98, &H49F7, &HA7, &HDC, &HFE, &H53, &H9D, &HC1, &H5B, &HE7)
Selection_Pattern2_GUID = iid
End Function
Public Function HeadingLevel_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H29084272, &HAAAF, &H4A30, &H87, &H96, &H3C, &H12, &HF6, &H2B, &H6B, &HBB)
HeadingLevel_Property_GUID = iid
End Function
Public Function IsDialog_Property_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9D0DFB9B, &H8436, &H4501, &HBB, &HBB, &HE5, &H34, &HA4, &HFB, &H3B, &H3F)
IsDialog_Property_GUID = iid
End Function
Public Function PROPID_ACC_NAME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H608D3DF8, &H8128, &H4AA7, &HA4, &H28, &HF5, &H5E, &H49, &H26, &H72, &H91)
 PROPID_ACC_NAME = iid
End Function
Public Function PROPID_ACC_VALUE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H123FE443, &H211A, &H4615, &H95, &H27, &HC4, &H5A, &H7E, &H93, &H71, &H7A)
 PROPID_ACC_VALUE = iid
End Function
Public Function PROPID_ACC_DESCRIPTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D48DFE4, &HBD3F, &H491F, &HA6, &H48, &H49, &H2D, &H6F, &H20, &HC5, &H88)
 PROPID_ACC_DESCRIPTION = iid
End Function
Public Function PROPID_ACC_ROLE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCB905FF2, &H7BD1, &H4C05, &HB3, &HC8, &HE6, &HC2, &H41, &H36, &H4D, &H70)
 PROPID_ACC_ROLE = iid
End Function
Public Function PROPID_ACC_STATE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA8D4D5B0, &HA21, &H42D0, &HA5, &HC0, &H51, &H4E, &H98, &H4F, &H45, &H7B)
 PROPID_ACC_STATE = iid
End Function
Public Function PROPID_ACC_HELP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC831E11F, &H44DB, &H4A99, &H97, &H68, &HCB, &H8F, &H97, &H8B, &H72, &H31)
 PROPID_ACC_HELP = iid
End Function
Public Function PROPID_ACC_KEYBOARDSHORTCUT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7D9BCEEE, &H7D1E, &H4979, &H93, &H82, &H51, &H80, &HF4, &H17, &H2C, &H34)
 PROPID_ACC_KEYBOARDSHORTCUT = iid
End Function
Public Function PROPID_ACC_DEFAULTACTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H180C072B, &HC27F, &H43C7, &H99, &H22, &HF6, &H35, &H62, &HA4, &H63, &H2B)
 PROPID_ACC_DEFAULTACTION = iid
End Function
Public Function PROPID_ACC_HELPTOPIC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H787D1379, &H8EDE, &H440B, &H8A, &HEC, &H11, &HF7, &HBF, &H90, &H30, &HB3)
 PROPID_ACC_HELPTOPIC = iid
End Function
Public Function PROPID_ACC_FOCUS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6EB335DF, &H1C29, &H4127, &HB1, &H2C, &HDE, &HE9, &HFD, &H15, &H7F, &H2B)
 PROPID_ACC_FOCUS = iid
End Function
Public Function PROPID_ACC_SELECTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB99D073C, &HD731, &H405B, &H90, &H61, &HD9, &H5E, &H8F, &H84, &H29, &H84)
 PROPID_ACC_SELECTION = iid
End Function
Public Function PROPID_ACC_PARENT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H474C22B6, &HFFC2, &H467A, &HB1, &HB5, &HE9, &H58, &HB4, &H65, &H73, &H30)
 PROPID_ACC_PARENT = iid
End Function
Public Function PROPID_ACC_NAV_UP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H16E1A2B, &H1A4E, &H4767, &H86, &H12, &H33, &H86, &HF6, &H69, &H35, &HEC)
 PROPID_ACC_NAV_UP = iid
End Function
Public Function PROPID_ACC_NAV_DOWN() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H31670ED, &H3CDF, &H48D2, &H96, &H13, &H13, &H8F, &H2D, &HD8, &HA6, &H68)
 PROPID_ACC_NAV_DOWN = iid
End Function
Public Function PROPID_ACC_NAV_LEFT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H228086CB, &H82F1, &H4A39, &H87, &H5, &HDC, &HDC, &HF, &HFF, &H92, &HF5)
 PROPID_ACC_NAV_LEFT = iid
End Function
Public Function PROPID_ACC_NAV_RIGHT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCD211D9F, &HE1CB, &H4FE5, &HA7, &H7C, &H92, &HB, &H88, &H4D, &H9, &H5B)
 PROPID_ACC_NAV_RIGHT = iid
End Function
Public Function PROPID_ACC_NAV_PREV() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H776D3891, &HC73B, &H4480, &HB3, &HF6, &H7, &H6A, &H16, &HA1, &H5A, &HF6)
 PROPID_ACC_NAV_PREV = iid
End Function
Public Function PROPID_ACC_NAV_NEXT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1CDC5455, &H8CD9, &H4C92, &HA3, &H71, &H39, &H39, &HA2, &HFE, &H3E, &HEE)
 PROPID_ACC_NAV_NEXT = iid
End Function
Public Function PROPID_ACC_NAV_FIRSTCHILD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCFD02558, &H557B, &H4C67, &H84, &HF9, &H2A, &H9, &HFC, &HE4, &H7, &H49)
 PROPID_ACC_NAV_FIRSTCHILD = iid
End Function
Public Function PROPID_ACC_NAV_LASTCHILD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H302ECAA5, &H48D5, &H4F8D, &HB6, &H71, &H1A, &H8D, &H20, &HA7, &H78, &H32)
 PROPID_ACC_NAV_LASTCHILD = iid
End Function
Public Function PROPID_ACC_ROLEMAP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF79ACDA2, &H140D, &H4FE6, &H89, &H14, &H20, &H84, &H76, &H32, &H82, &H69)
 PROPID_ACC_ROLEMAP = iid
End Function
Public Function PROPID_ACC_VALUEMAP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDA1C3D79, &HFC5C, &H420E, &HB3, &H99, &H9D, &H15, &H33, &H54, &H9E, &H75)
 PROPID_ACC_VALUEMAP = iid
End Function
Public Function PROPID_ACC_STATEMAP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43946C5E, &HAC0, &H4042, &HB5, &H25, &H7, &HBB, &HDB, &HE1, &H7F, &HA7)
 PROPID_ACC_STATEMAP = iid
End Function
Public Function PROPID_ACC_DESCRIPTIONMAP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1FF1435F, &H8A14, &H477B, &HB2, &H26, &HA0, &HAB, &HE2, &H79, &H97, &H5D)
 PROPID_ACC_DESCRIPTIONMAP = iid
End Function
Public Function PROPID_ACC_DODEFAULTACTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1BA09523, &H2E3B, &H49A6, &HA0, &H59, &H59, &H68, &H2A, &H3C, &H48, &HFD)
 PROPID_ACC_DODEFAULTACTION = iid
End Function

