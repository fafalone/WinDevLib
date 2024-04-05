Attribute VB_Name = "mDirectX"
Option Explicit

'modDirectX - IIDs for DirectWrite, Direct2D, and Media Foundation

Private Const PID_FIRST_USABLE = 2

Public Const D3D11_DEFAULT_BLEND_FACTOR_ALPHA As Single = (1#)
Public Const D3D11_DEFAULT_BLEND_FACTOR_BLUE    As Single = 1
Public Const D3D11_DEFAULT_BLEND_FACTOR_GREEN = &H3F800000
Public Const D3D11_DEFAULT_BLEND_FACTOR_RED As Single = 1
Public Const D3D11_DEFAULT_BORDER_COLOR_COMPONENT = 0 ' (0.0)
Public Const D3D11_DEFAULT_DEPTH_BIAS_CLAMP = 0   ' (0.0)
Public Const D3D11_DEFAULT_MIP_LOD_BIAS = 0   ' (0.0)
Public Const D3D11_DEFAULT_SLOPE_SCALED_DEPTH_BIAS = 0    ' (0.0)
Public Const D3D11_FLOAT16_FUSED_TOLERANCE_IN_ULP   As Single = (0.6)
Public Const D3D11_FLOAT32_MAX As Single = (3.402823466E+38)
Public Const D3D11_FLOAT32_TO_INTEGER_TOLERANCE_IN_ULP  As Single = (0.6)
Public Const D3D11_FLOAT_TO_SRGB_EXPONENT_DENOMINATOR As Single = (2.4)
Public Const D3D11_FLOAT_TO_SRGB_EXPONENT_NUMERATOR As Single = 1
Public Const D3D11_FLOAT_TO_SRGB_OFFSET As Single = (0.055)
Public Const D3D11_FLOAT_TO_SRGB_SCALE_1 As Single = (12.92)
Public Const D3D11_FLOAT_TO_SRGB_SCALE_2 As Single = (1.055)
Public Const D3D11_FLOAT_TO_SRGB_THRESHOLD As Single = (0.0031308)
Public Const D3D11_FTOI_INSTRUCTION_MAX_INPUT As Single = (2147483647.999)
Public Const D3D11_FTOI_INSTRUCTION_MIN_INPUT As Single = (-2147483648.999)
Public Const D3D11_FTOU_INSTRUCTION_MAX_INPUT As Single = (4294967295.999)
Public Const D3D11_FTOU_INSTRUCTION_MIN_INPUT = 0 ' (0.0)
Public Const D3D11_DEFAULT_VIEWPORT_MAX_DEPTH = 0   ' (0.0)
Public Const D3D11_DEFAULT_VIEWPORT_MIN_DEPTH = 0   ' (0.0)
Public Const D3D11_HS_MAXTESSFACTOR_LOWER_BOUND = (1#)
Public Const D3D11_HS_MAXTESSFACTOR_UPPER_BOUND As Single = (64#)
Public Const D3D11_LINEAR_GAMMA As Single = (1#)
Public Const D3D11_MAX_BORDER_COLOR_COMPONENT   As Single = (1#)
Public Const D3D11_MAX_POSITION_VALUE As Single = (3.402823466E+34)
Public Const D3D11_MIN_BORDER_COLOR_COMPONENT = 0   ' ( 0.0)
Public Const D3D11_MIN_DEPTH = 0    ' ( 0.0)
Public Const D3D11_MIP_LOD_BIAS_MAX As Single = (15.99)
Public Const D3D11_MIP_LOD_BIAS_MIN As Single = (-16#)
Public Const D3D11_MULTISAMPLE_ANTIALIAS_LINE_WIDTH As Single = (1.4)
Public Const D3D11_PS_LEGACY_PIXEL_CENTER_FRACTIONAL_COMPONENT = 0  ' ( 0.0)
Public Const D3D11_PS_PIXEL_CENTER_FRACTIONAL_COMPONENT As Single = (0.5)
Public Const D3D11_REQ_RESOURCE_SIZE_IN_MEGABYTES_EXPRESSION_B_TERM As Single = (0.25)
Public Const D3D11_SPEC_VERSION As Single = (1.07)
Public Const D3D11_SRGB_GAMMA As Single = (2.2)
Public Const D3D11_SRGB_TO_FLOAT_DENOMINATOR_1 As Single = (12.92)
Public Const D3D11_SRGB_TO_FLOAT_DENOMINATOR_2 As Single = (1.055)
Public Const D3D11_SRGB_TO_FLOAT_EXPONENT As Single = (2.4)
Public Const D3D11_SRGB_TO_FLOAT_OFFSET As Single = (0.055)
Public Const D3D11_SRGB_TO_FLOAT_THRESHOLD As Single = (0.04045)
Public Const D3D11_SRGB_TO_FLOAT_TOLERANCE_IN_ULP   As Single = (0.5)
Public Const D3D12_DEFAULT_BLEND_FACTOR_ALPHA As Single = (1#)
Public Const D3D12_DEFAULT_BLEND_FACTOR_BLUE As Single = (1#)
Public Const D3D12_DEFAULT_BLEND_FACTOR_GREEN As Single = (1#)
Public Const D3D12_DEFAULT_BLEND_FACTOR_RED As Single = (1#)
Public Const D3D12_DEFAULT_BORDER_COLOR_COMPONENT = 0 ' (0.0f)
Public Const D3D12_DEFAULT_DEPTH_BIAS_CLAMP = 0 ' (0.0f)
Public Const D3D12_DEFAULT_MIP_LOD_BIAS = 0 '( 0.0f )
Public Const D3D12_DEFAULT_SLOPE_SCALED_DEPTH_BIAS = 0 ' (0.0f)
Public Const D3D12_DEFAULT_VIEWPORT_MAX_DEPTH = 0 ' (0.0f)
Public Const D3D12_DEFAULT_VIEWPORT_MIN_DEPTH = 0 ' (0.0f)
Public Const D3D12_FLOAT16_FUSED_TOLERANCE_IN_ULP   As Single = (0.6)
Public Const D3D12_FLOAT32_MAX As Single = (3.402823466E+38)
Public Const D3D12_FLOAT32_TO_INTEGER_TOLERANCE_IN_ULP As Single = (0.6)
Public Const D3D12_FLOAT_TO_SRGB_EXPONENT_DENOMINATOR As Single = (2.4)
Public Const D3D12_FLOAT_TO_SRGB_EXPONENT_NUMERATOR As Single = 1#
Public Const D3D12_FLOAT_TO_SRGB_OFFSET As Single = (0.055)
Public Const D3D12_FLOAT_TO_SRGB_SCALE_1 As Single = (12.92)
Public Const D3D12_FLOAT_TO_SRGB_SCALE_2 As Single = (1.055)
Public Const D3D12_FLOAT_TO_SRGB_THRESHOLD  As Single = (0.0031308)
Public Const D3D12_FTOI_INSTRUCTION_MAX_INPUT As Single = (2147483647.999)
Public Const D3D12_FTOI_INSTRUCTION_MIN_INPUT As Single = (-2147483648.999)
Public Const D3D12_FTOU_INSTRUCTION_MAX_INPUT As Single = (4294967295.999)
Public Const D3D12_FTOU_INSTRUCTION_MIN_INPUT = 0 ' (0.0)
Public Const D3D12_HS_MAXTESSFACTOR_LOWER_BOUND As Single = (1#)
Public Const D3D12_HS_MAXTESSFACTOR_UPPER_BOUND As Single = (64#)
Public Const D3D12_LINEAR_GAMMA As Single = (1#)
Public Const D3D12_MAX_BORDER_COLOR_COMPONENT As Single = (1#)
Public Const D3D12_MAX_DEPTH As Single = (1#)
Public Const D3D12_MAX_POSITION_VALUE As Single = (3.402823466E+34)
Public Const D3D12_MIN_BORDER_COLOR_COMPONENT = 0 ' (0.0f)
Public Const D3D12_MIN_DEPTH = 0 '  (0.0f)
Public Const D3D12_MIP_LOD_BIAS_MAX As Single = (15.99)
Public Const D3D12_MIP_LOD_BIAS_MIN As Single = (-16#)
Public Const D3D12_MULTISAMPLE_ANTIALIAS_LINE_WIDTH As Single = (1.4)
Public Const D3D12_PS_LEGACY_PIXEL_CENTER_FRACTIONAL_COMPONENT = 0 ' (0.0f)
Public Const D3D12_PS_PIXEL_CENTER_FRACTIONAL_COMPONENT As Single = (0.5)
Public Const D3D12_REQ_RESOURCE_SIZE_IN_MEGABYTES_EXPRESSION_B_TERM As Single = (0.25)
Public Const D3D12_SPEC_VERSION As Single = (1.16)
Public Const D3D12_SRGB_GAMMA As Single = (2.2)
Public Const D3D12_SRGB_TO_FLOAT_DENOMINATOR_1 As Single = (12.92)
Public Const D3D12_SRGB_TO_FLOAT_DENOMINATOR_2 As Single = (1.055)
Public Const D3D12_SRGB_TO_FLOAT_EXPONENT As Single = (2.4)
Public Const D3D12_SRGB_TO_FLOAT_OFFSET As Single = (0.055)
Public Const D3D12_SRGB_TO_FLOAT_THRESHOLD As Single = (0.04045)
Public Const D3D12_SRGB_TO_FLOAT_TOLERANCE_IN_ULP As Single = (0.5)

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

Public Function CLSID_D2D12DAffineTransform() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6AA97485, &H6354, &H4CFC, &H90, &H8C, &HE4, &HA7, &H4F, &H62, &HC9, &H6C)
CLSID_D2D12DAffineTransform = iid
End Function
Public Function CLSID_D2D13DPerspectiveTransform() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC2844D0B, &H3D86, &H46E7, &H85, &HBA, &H52, &H6C, &H92, &H40, &HF3, &HFB)
CLSID_D2D13DPerspectiveTransform = iid
End Function
Public Function CLSID_D2D13DTransform() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE8467B04, &HEC61, &H4B8A, &HB5, &HDE, &HD4, &HD7, &H3D, &HEB, &HEA, &H5A)
CLSID_D2D13DTransform = iid
End Function
Public Function CLSID_D2D1ArithmeticComposite() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFC151437, &H49A, &H4784, &HA2, &H4A, &HF1, &HC4, &HDA, &HF2, &H9, &H87)
CLSID_D2D1ArithmeticComposite = iid
End Function
Public Function CLSID_D2D1Atlas() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H913E2BE4, &HFDCF, &H4FE2, &HA5, &HF0, &H24, &H54, &HF1, &H4F, &HF4, &H8)
CLSID_D2D1Atlas = iid
End Function
Public Function CLSID_D2D1BitmapSource() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5FB6C24D, &HC6DD, &H4231, &H94, &H4, &H50, &HF4, &HD5, &HC3, &H25, &H2D)
CLSID_D2D1BitmapSource = iid
End Function
Public Function CLSID_D2D1Blend() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H81C5B77B, &H13F8, &H4CDD, &HAD, &H20, &HC8, &H90, &H54, &H7A, &HC6, &H5D)
CLSID_D2D1Blend = iid
End Function
Public Function CLSID_D2D1Border() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2A2D49C0, &H4ACF, &H43C7, &H8C, &H6A, &H7C, &H4A, &H27, &H87, &H4D, &H27)
CLSID_D2D1Border = iid
End Function
Public Function CLSID_D2D1Brightness() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8CEA8D1E, &H77B0, &H4986, &HB3, &HB9, &H2F, &HC, &HE, &HAE, &H78, &H87)
CLSID_D2D1Brightness = iid
End Function
Public Function CLSID_D2D1ColorManagement() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1A28524C, &HFDD6, &H4AA4, &HAE, &H8F, &H83, &H7E, &HB8, &H26, &H7B, &H37)
CLSID_D2D1ColorManagement = iid
End Function
Public Function CLSID_D2D1ColorMatrix() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H921F03D6, &H641C, &H47DF, &H85, &H2D, &HB4, &HBB, &H61, &H53, &HAE, &H11)
CLSID_D2D1ColorMatrix = iid
End Function
Public Function CLSID_D2D1Composite() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H48FC9F51, &HF6AC, &H48F1, &H8B, &H58, &H3B, &H28, &HAC, &H46, &HF7, &H6D)
CLSID_D2D1Composite = iid
End Function
Public Function CLSID_D2D1ConvolveMatrix() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H407F8C08, &H5533, &H4331, &HA3, &H41, &H23, &HCC, &H38, &H77, &H84, &H3E)
CLSID_D2D1ConvolveMatrix = iid
End Function
Public Function CLSID_D2D1Crop() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE23F7110, &HE9A, &H4324, &HAF, &H47, &H6A, &H2C, &HC, &H46, &HF3, &H5B)
CLSID_D2D1Crop = iid
End Function
Public Function CLSID_D2D1DirectionalBlur() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H174319A6, &H58E9, &H49B2, &HBB, &H63, &HCA, &HF2, &HC8, &H11, &HA3, &HDB)
CLSID_D2D1DirectionalBlur = iid
End Function
Public Function CLSID_D2D1DiscreteTransfer() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H90866FCD, &H488E, &H454B, &HAF, &H6, &HE5, &H4, &H1B, &H66, &HC3, &H6C)
CLSID_D2D1DiscreteTransfer = iid
End Function
Public Function CLSID_D2D1DisplacementMap() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEDC48364, &H417, &H4111, &H94, &H50, &H43, &H84, &H5F, &HA9, &HF8, &H90)
CLSID_D2D1DisplacementMap = iid
End Function
Public Function CLSID_D2D1DistantDiffuse() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3E7EFD62, &HA32D, &H46D4, &HA8, &H3C, &H52, &H78, &H88, &H9A, &HC9, &H54)
CLSID_D2D1DistantDiffuse = iid
End Function
Public Function CLSID_D2D1DistantSpecular() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H428C1EE5, &H77B8, &H4450, &H8A, &HB5, &H72, &H21, &H9C, &H21, &HAB, &HDA)
CLSID_D2D1DistantSpecular = iid
End Function
Public Function CLSID_D2D1DpiCompensation() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C26C5C7, &H34E0, &H46FC, &H9C, &HFD, &HE5, &H82, &H37, &H6, &HE2, &H28)
CLSID_D2D1DpiCompensation = iid
End Function
Public Function CLSID_D2D1Flood() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H61C23C20, &HAE69, &H4D8E, &H94, &HCF, &H50, &H7, &H8D, &HF6, &H38, &HF2)
CLSID_D2D1Flood = iid
End Function
Public Function CLSID_D2D1GammaTransfer() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H409444C4, &HC419, &H41A0, &HB0, &HC1, &H8C, &HD0, &HC0, &HA1, &H8E, &H42)
CLSID_D2D1GammaTransfer = iid
End Function
Public Function CLSID_D2D1GaussianBlur() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1FEB6D69, &H2FE6, &H4AC9, &H8C, &H58, &H1D, &H7F, &H93, &HE7, &HA6, &HA5)
CLSID_D2D1GaussianBlur = iid
End Function
Public Function CLSID_D2D1Scale() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9DAF9369, &H3846, &H4D0E, &HA4, &H4E, &HC, &H60, &H79, &H34, &HA5, &HD7)
CLSID_D2D1Scale = iid
End Function
Public Function CLSID_D2D1Histogram() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H881DB7D0, &HF7EE, &H4D4D, &HA6, &HD2, &H46, &H97, &HAC, &HC6, &H6E, &HE8)
CLSID_D2D1Histogram = iid
End Function
Public Function CLSID_D2D1HueRotation() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF4458EC, &H4B32, &H491B, &H9E, &H85, &HBD, &H73, &HF4, &H4D, &H3E, &HB6)
CLSID_D2D1HueRotation = iid
End Function
Public Function CLSID_D2D1LinearTransfer() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAD47C8FD, &H63EF, &H4ACC, &H9B, &H51, &H67, &H97, &H9C, &H3, &H6C, &H6)
CLSID_D2D1LinearTransfer = iid
End Function
Public Function CLSID_D2D1LuminanceToAlpha() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H41251AB7, &HBEB, &H46F8, &H9D, &HA7, &H59, &HE9, &H3F, &HCC, &HE5, &HDE)
CLSID_D2D1LuminanceToAlpha = iid
End Function
Public Function CLSID_D2D1Morphology() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEAE6C40D, &H626A, &H4C2D, &HBF, &HCB, &H39, &H10, &H1, &HAB, &HE2, &H2)
CLSID_D2D1Morphology = iid
End Function
Public Function CLSID_D2D1OpacityMetadata() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6C53006A, &H4450, &H4199, &HAA, &H5B, &HAD, &H16, &H56, &HFE, &HCE, &H5E)
CLSID_D2D1OpacityMetadata = iid
End Function
Public Function CLSID_D2D1PointDiffuse() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB9E303C3, &HC08C, &H4F91, &H8B, &H7B, &H38, &H65, &H6B, &HC4, &H8C, &H20)
CLSID_D2D1PointDiffuse = iid
End Function
Public Function CLSID_D2D1PointSpecular() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9C3CA26, &H3AE2, &H4F09, &H9E, &HBC, &HED, &H38, &H65, &HD5, &H3F, &H22)
CLSID_D2D1PointSpecular = iid
End Function
Public Function CLSID_D2D1Premultiply() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6EAB419, &HDEED, &H4018, &H80, &HD2, &H3E, &H1D, &H47, &H1A, &HDE, &HB2)
CLSID_D2D1Premultiply = iid
End Function
Public Function CLSID_D2D1Saturation() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5CB2D9CF, &H327D, &H459F, &HA0, &HCE, &H40, &HC0, &HB2, &H8, &H6B, &HF7)
CLSID_D2D1Saturation = iid
End Function
Public Function CLSID_D2D1Shadow() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC67EA361, &H1863, &H4E69, &H89, &HDB, &H69, &H5D, &H3E, &H9A, &H5B, &H6B)
CLSID_D2D1Shadow = iid
End Function
Public Function CLSID_D2D1SpotDiffuse() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H818A1105, &H7932, &H44F4, &HAA, &H86, &H8, &HAE, &H7B, &H2F, &H2C, &H93)
CLSID_D2D1SpotDiffuse = iid
End Function
Public Function CLSID_D2D1SpotSpecular() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEDAE421E, &H7654, &H4A37, &H9D, &HB8, &H71, &HAC, &HC1, &HBE, &HB3, &HC1)
CLSID_D2D1SpotSpecular = iid
End Function
Public Function CLSID_D2D1TableTransfer() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5BF818C3, &H5E43, &H48CB, &HB6, &H31, &H86, &H83, &H96, &HD6, &HA1, &HD4)
CLSID_D2D1TableTransfer = iid
End Function
Public Function CLSID_D2D1Tile() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB0784138, &H3B76, &H4BC5, &HB1, &H3B, &HF, &HA2, &HAD, &H2, &H65, &H9F)
CLSID_D2D1Tile = iid
End Function
Public Function CLSID_D2D1Turbulence() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCF2BB6AE, &H889A, &H4AD7, &HBA, &H29, &HA2, &HFD, &H73, &H2C, &H9F, &HC9)
CLSID_D2D1Turbulence = iid
End Function
Public Function CLSID_D2D1UnPremultiply() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFB9AC489, &HAD8D, &H41ED, &H99, &H99, &HBB, &H63, &H47, &HD1, &H10, &HF7)
CLSID_D2D1UnPremultiply = iid
End Function


'modDirectX - IIDs for DirectWrite and Direct2D


Public Function IID_ID2D1Factory() As UUID
'{06152247-6F50-465A-9245-118BFD3B6007}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6152247, CInt(&H6F50), CInt(&H465A), &H92, &H45, &H11, &H8B, &HFD, &H3B, &H60, &H7)
IID_ID2D1Factory = iid
End Function
Public Function IID_ID2D1RectangleGeometry() As UUID
'{2CD906A2-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD906A2, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1RectangleGeometry = iid
End Function
Public Function IID_ID2D1Geometry() As UUID
'{2CD906A1-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD906A1, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1Geometry = iid
End Function
Public Function IID_ID2D1Resource() As UUID
'{2CD90691-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD90691, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1Resource = iid
End Function
Public Function IID_ID2D1StrokeStyle() As UUID
'{2CD9069D-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD9069D, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1StrokeStyle = iid
End Function
Public Function IID_ID2D1SimplifiedGeometrySink() As UUID
'{2CD9069E-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD9069E, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1SimplifiedGeometrySink = iid
End Function
Public Function IID_ID2D1TessellationSink() As UUID
'{2CD906C1-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD906C1, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1TessellationSink = iid
End Function
Public Function IID_ID2D1RoundedRectangleGeometry() As UUID
'{2CD906A3-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD906A3, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1RoundedRectangleGeometry = iid
End Function
Public Function IID_ID2D1EllipseGeometry() As UUID
'{2CD906A4-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD906A4, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1EllipseGeometry = iid
End Function
Public Function IID_ID2D1GeometryGroup() As UUID
'{2CD906A6-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD906A6, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1GeometryGroup = iid
End Function
Public Function IID_ID2D1TransformedGeometry() As UUID
'{2CD906BB-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD906BB, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1TransformedGeometry = iid
End Function
Public Function IID_ID2D1PathGeometry() As UUID
'{2CD906A5-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD906A5, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1PathGeometry = iid
End Function
Public Function IID_ID2D1GeometrySink() As UUID
'{2CD9069F-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD9069F, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1GeometrySink = iid
End Function
Public Function IID_ID2D1DrawingStateBlock() As UUID
'{28506E39-EBF6-46A1-BB47-FD85565AB957}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H28506E39, CInt(&HEBF6), CInt(&H46A1), &HBB, &H47, &HFD, &H85, &H56, &H5A, &HB9, &H57)
IID_ID2D1DrawingStateBlock = iid
End Function
Public Function IID_ID2D1RenderTarget() As UUID
'{2CD90694-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD90694, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1RenderTarget = iid
End Function
Public Function IID_ID2D1Bitmap() As UUID
'{A2296057-EA42-4099-983B-539FB6505426}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA2296057, CInt(&HEA42), CInt(&H4099), &H98, &H3B, &H53, &H9F, &HB6, &H50, &H54, &H26)
IID_ID2D1Bitmap = iid
End Function
Public Function IID_ID2D1BitmapBrush() As UUID
'{2CD906AA-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD906AA, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1BitmapBrush = iid
End Function
Public Function IID_ID2D1Brush() As UUID
'{2CD906A8-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD906A8, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1Brush = iid
End Function
Public Function IID_ID2D1SolidColorBrush() As UUID
'{2CD906A9-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD906A9, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1SolidColorBrush = iid
End Function
Public Function IID_ID2D1GradientStopCollection() As UUID
'{2CD906A7-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD906A7, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1GradientStopCollection = iid
End Function
Public Function IID_ID2D1LinearGradientBrush() As UUID
'{2CD906AB-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD906AB, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1LinearGradientBrush = iid
End Function
Public Function IID_ID2D1RadialGradientBrush() As UUID
'{2CD906AC-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD906AC, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1RadialGradientBrush = iid
End Function
Public Function IID_ID2D1BitmapRenderTarget() As UUID
'{2CD90695-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD90695, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1BitmapRenderTarget = iid
End Function
Public Function IID_ID2D1Layer() As UUID
'{2CD9069B-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD9069B, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1Layer = iid
End Function
Public Function IID_ID2D1Mesh() As UUID
'{2CD906C2-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD906C2, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1Mesh = iid
End Function
Public Function IID_ID2D1HwndRenderTarget() As UUID
'{2CD90698-12E2-11DC-9FED-001143A055F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2CD90698, CInt(&H12E2), CInt(&H11DC), &H9F, &HED, &H0, &H11, &H43, &HA0, &H55, &HF9)
IID_ID2D1HwndRenderTarget = iid
End Function
Public Function IID_ID2D1DCRenderTarget() As UUID
'{1C51BC64-DE61-46FD-9899-63A5D8F03950}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1C51BC64, CInt(&HDE61), CInt(&H46FD), &H98, &H99, &H63, &HA5, &HD8, &HF0, &H39, &H50)
IID_ID2D1DCRenderTarget = iid
End Function
Public Function IID_ID2D1GdiInteropRenderTarget() As UUID
'{E0DB51C3-6F77-4BAE-B3D5-E47509B35838}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE0DB51C3, CInt(&H6F77), CInt(&H4BAE), &HB3, &HD5, &HE4, &H75, &H9, &HB3, &H58, &H38)
IID_ID2D1GdiInteropRenderTarget = iid
End Function




Public Function IID_IDWriteFactory() As UUID
'{B859EE5A-D838-4B5B-A2E8-1ADC7D93DB48}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB859EE5A, CInt(&HD838), CInt(&H4B5B), &HA2, &HE8, &H1A, &HDC, &H7D, &H93, &HDB, &H48)
IID_IDWriteFactory = iid
End Function
Public Function IID_IDWriteFontCollection() As UUID
'{A84CEE02-3EEA-4EEE-A827-87C1A02A0FCC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA84CEE02, CInt(&H3EEA), CInt(&H4EEE), &HA8, &H27, &H87, &HC1, &HA0, &H2A, &HF, &HCC)
IID_IDWriteFontCollection = iid
End Function
Public Function IID_IDWriteFontFamily() As UUID
'{DA20D8EF-812A-4C43-9802-62EC4ABD7ADD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDA20D8EF, CInt(&H812A), CInt(&H4C43), &H98, &H2, &H62, &HEC, &H4A, &HBD, &H7A, &HDD)
IID_IDWriteFontFamily = iid
End Function
Public Function IID_IDWriteFontList() As UUID
'{1A0D8438-1D97-4EC1-AEF9-A2FB86ED6ACB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1A0D8438, CInt(&H1D97), CInt(&H4EC1), &HAE, &HF9, &HA2, &HFB, &H86, &HED, &H6A, &HCB)
IID_IDWriteFontList = iid
End Function
Public Function IID_IDWriteFont() As UUID
'{ACD16696-8C14-4F5D-877E-FE3FC1D32737}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HACD16696, CInt(&H8C14), CInt(&H4F5D), &H87, &H7E, &HFE, &H3F, &HC1, &HD3, &H27, &H37)
IID_IDWriteFont = iid
End Function
Public Function IID_IDWriteLocalizedStrings() As UUID
'{08256209-099A-4B34-B86D-C22B110E7771}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8256209, CInt(&H99A), CInt(&H4B34), &HB8, &H6D, &HC2, &H2B, &H11, &HE, &H77, &H71)
IID_IDWriteLocalizedStrings = iid
End Function
Public Function IID_IDWriteFontFace() As UUID
'{5F49804D-7024-4D43-BFA9-D25984F53849}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5F49804D, CInt(&H7024), CInt(&H4D43), &HBF, &HA9, &HD2, &H59, &H84, &HF5, &H38, &H49)
IID_IDWriteFontFace = iid
End Function
Public Function IID_IDWriteRenderingParams() As UUID
'{2F0DA53A-2ADD-47CD-82EE-D9EC34688E75}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2F0DA53A, CInt(&H2ADD), CInt(&H47CD), &H82, &HEE, &HD9, &HEC, &H34, &H68, &H8E, &H75)
IID_IDWriteRenderingParams = iid
End Function
Public Function IID_IDWriteFontCollectionLoader() As UUID
'{CCA920E4-52F0-492B-BFA8-29C72EE0A468}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCCA920E4, CInt(&H52F0), CInt(&H492B), &HBF, &HA8, &H29, &HC7, &H2E, &HE0, &HA4, &H68)
IID_IDWriteFontCollectionLoader = iid
End Function
Public Function IID_IDWriteFontFileEnumerator() As UUID
'{72755049-5FF7-435D-8348-4BE97CFA6C7C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H72755049, CInt(&H5FF7), CInt(&H435D), &H83, &H48, &H4B, &HE9, &H7C, &HFA, &H6C, &H7C)
IID_IDWriteFontFileEnumerator = iid
End Function
Public Function IID_IDWriteFontFile() As UUID
'{739D886A-CEF5-47DC-8769-1A8B41BEBBB0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H739D886A, CInt(&HCEF5), CInt(&H47DC), &H87, &H69, &H1A, &H8B, &H41, &HBE, &HBB, &HB0)
IID_IDWriteFontFile = iid
End Function
Public Function IID_IDWriteFontFileLoader() As UUID
'{727CAD4E-D6AF-4C9E-8A08-D695B11CAA49}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H727CAD4E, CInt(&HD6AF), CInt(&H4C9E), &H8A, &H8, &HD6, &H95, &HB1, &H1C, &HAA, &H49)
IID_IDWriteFontFileLoader = iid
End Function
Public Function IID_IDWriteFontFileStream() As UUID
'{6D4865FE-0AB8-4D91-8F62-5DD6BE34A3E0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6D4865FE, CInt(&HAB8), CInt(&H4D91), &H8F, &H62, &H5D, &HD6, &HBE, &H34, &HA3, &HE0)
IID_IDWriteFontFileStream = iid
End Function
Public Function IID_IDWriteTextFormat() As UUID
'{9C906818-31D7-4FD3-A151-7C5E225DB55A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9C906818, CInt(&H31D7), CInt(&H4FD3), &HA1, &H51, &H7C, &H5E, &H22, &H5D, &HB5, &H5A)
IID_IDWriteTextFormat = iid
End Function
Public Function IID_IDWriteInlineObject() As UUID
'{8339FDE3-106F-47AB-8373-1C6295EB10B3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8339FDE3, CInt(&H106F), CInt(&H47AB), &H83, &H73, &H1C, &H62, &H95, &HEB, &H10, &HB3)
IID_IDWriteInlineObject = iid
End Function
Public Function IID_IDWriteTextRenderer() As UUID
'{EF8A8135-5CC6-45FE-8825-C5A0724EB819}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEF8A8135, CInt(&H5CC6), CInt(&H45FE), &H88, &H25, &HC5, &HA0, &H72, &H4E, &HB8, &H19)
IID_IDWriteTextRenderer = iid
End Function
Public Function IID_IDWritePixelSnapping() As UUID
'{EAF3A2DA-ECF4-4D24-B644-B34F6842024B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEAF3A2DA, CInt(&HECF4), CInt(&H4D24), &HB6, &H44, &HB3, &H4F, &H68, &H42, &H2, &H4B)
IID_IDWritePixelSnapping = iid
End Function
Public Function IID_IDWriteTypography() As UUID
'{55F1112B-1DC2-4B3C-9541-F46894ED85B6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H55F1112B, CInt(&H1DC2), CInt(&H4B3C), &H95, &H41, &HF4, &H68, &H94, &HED, &H85, &HB6)
IID_IDWriteTypography = iid
End Function
Public Function IID_IDWriteGdiInterop() As UUID
'{1EDD9491-9853-4299-898F-6432983B6F3A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1EDD9491, CInt(&H9853), CInt(&H4299), &H89, &H8F, &H64, &H32, &H98, &H3B, &H6F, &H3A)
IID_IDWriteGdiInterop = iid
End Function
Public Function IID_IDWriteBitmapRenderTarget() As UUID
'{5E5A32A3-8DFF-4773-9FF6-0696EAB77267}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5E5A32A3, CInt(&H8DFF), CInt(&H4773), &H9F, &HF6, &H6, &H96, &HEA, &HB7, &H72, &H67)
IID_IDWriteBitmapRenderTarget = iid
End Function
Public Function IID_IDWriteTextLayout() As UUID
'{53737037-6D14-410B-9BFE-0B182BB70961}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H53737037, CInt(&H6D14), CInt(&H410B), &H9B, &HFE, &HB, &H18, &H2B, &HB7, &H9, &H61)
IID_IDWriteTextLayout = iid
End Function
Public Function IID_IDWriteTextAnalyzer() As UUID
'{B7E6163E-7F46-43B4-84B3-E4E6249C365D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB7E6163E, CInt(&H7F46), CInt(&H43B4), &H84, &HB3, &HE4, &HE6, &H24, &H9C, &H36, &H5D)
IID_IDWriteTextAnalyzer = iid
End Function
Public Function IID_IDWriteTextAnalysisSource() As UUID
'{688E1A58-5094-47C8-ADC8-FBCEA60AE92B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H688E1A58, CInt(&H5094), CInt(&H47C8), &HAD, &HC8, &HFB, &HCE, &HA6, &HA, &HE9, &H2B)
IID_IDWriteTextAnalysisSource = iid
End Function
Public Function IID_IDWriteNumberSubstitution() As UUID
'{14885CC9-BAB0-4F90-B6ED-5C366A2CD03D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H14885CC9, CInt(&HBAB0), CInt(&H4F90), &HB6, &HED, &H5C, &H36, &H6A, &H2C, &HD0, &H3D)
IID_IDWriteNumberSubstitution = iid
End Function
Public Function IID_IDWriteTextAnalysisSink() As UUID
'{5810CD44-0CA0-4701-B3FA-BEC5182AE4F6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5810CD44, CInt(&HCA0), CInt(&H4701), &HB3, &HFA, &HBE, &HC5, &H18, &H2A, &HE4, &HF6)
IID_IDWriteTextAnalysisSink = iid
End Function
Public Function IID_IDWriteGlyphRunAnalysis() As UUID
'{7D97DBF7-E085-42D4-81E3-6A883BDED118}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7D97DBF7, CInt(&HE085), CInt(&H42D4), &H81, &HE3, &H6A, &H88, &H3B, &HDE, &HD1, &H18)
IID_IDWriteGlyphRunAnalysis = iid
End Function
Public Function IID_IDWriteLocalFontFileLoader() As UUID
'{B2D9F3EC-C9FE-4A11-A2EC-D86208F7C0A2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB2D9F3EC, CInt(&HC9FE), CInt(&H4A11), &HA2, &HEC, &HD8, &H62, &H8, &HF7, &HC0, &HA2)
IID_IDWriteLocalFontFileLoader = iid
End Function

Public Function IID_IDXGIObject() As UUID
'{aec22fb8-76f3-4639-9be0-28eb43a67a2e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAEC22FB8, CInt(&H76F3), CInt(&H4639), &H9B, &HE0, &H28, &HEB, &H43, &HA6, &H7A, &H2E)
IID_IDXGIObject = iid
End Function
Public Function IID_IDXGIDeviceSubObject() As UUID
'{3d3e0379-f9de-4d58-bb6c-18d62992f1a6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3D3E0379, CInt(&HF9DE), CInt(&H4D58), &HBB, &H6C, &H18, &HD6, &H29, &H92, &HF1, &HA6)
IID_IDXGIDeviceSubObject = iid
End Function
Public Function IID_IDXGIResource() As UUID
'{035f3ab4-482e-4e50-b41f-8a7f8bd8960b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H35F3AB4, CInt(&H482E), CInt(&H4E50), &HB4, &H1F, &H8A, &H7F, &H8B, &HD8, &H96, &HB)
IID_IDXGIResource = iid
End Function
Public Function IID_IDXGIKeyedMutex() As UUID
'{9d8e1289-d7b3-465f-8126-250e349af85d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9D8E1289, CInt(&HD7B3), CInt(&H465F), &H81, &H26, &H25, &HE, &H34, &H9A, &HF8, &H5D)
IID_IDXGIKeyedMutex = iid
End Function
Public Function IID_IDXGISurface() As UUID
'{cafcb56c-6ac3-4889-bf47-9e23bbd260ec}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCAFCB56C, CInt(&H6AC3), CInt(&H4889), &HBF, &H47, &H9E, &H23, &HBB, &HD2, &H60, &HEC)
IID_IDXGISurface = iid
End Function
Public Function IID_IDXGISurface1() As UUID
'{4AE63092-6327-4c1b-80AE-BFE12EA32B86}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4AE63092, CInt(&H6327), CInt(&H4C1B), &H80, &HAE, &HBF, &HE1, &H2E, &HA3, &H2B, &H86)
IID_IDXGISurface1 = iid
End Function
Public Function IID_IDXGIAdapter() As UUID
'{2411e7e1-12ac-4ccf-bd14-9798e8534dc0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2411E7E1, CInt(&H12AC), CInt(&H4CCF), &HBD, &H14, &H97, &H98, &HE8, &H53, &H4D, &HC0)
IID_IDXGIAdapter = iid
End Function
Public Function IID_IDXGIOutput() As UUID
'{ae02eedb-c735-4690-8d52-5a8dc20213aa}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAE02EEDB, CInt(&HC735), CInt(&H4690), &H8D, &H52, &H5A, &H8D, &HC2, &H2, &H13, &HAA)
IID_IDXGIOutput = iid
End Function
Public Function IID_IDXGISwapChain() As UUID
'{310d36a0-d2e7-4c0a-aa04-6a9d23b8886a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H310D36A0, CInt(&HD2E7), CInt(&H4C0A), &HAA, &H4, &H6A, &H9D, &H23, &HB8, &H88, &H6A)
IID_IDXGISwapChain = iid
End Function
Public Function IID_IDXGIFactory() As UUID
'{7b7166ec-21c7-44ae-b21a-c9ae321ae369}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7B7166EC, CInt(&H21C7), CInt(&H44AE), &HB2, &H1A, &HC9, &HAE, &H32, &H1A, &HE3, &H69)
IID_IDXGIFactory = iid
End Function
Public Function IID_IDXGIDevice() As UUID
'{54ec77fa-1377-44e6-8c32-88fd5f44c84c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H54EC77FA, CInt(&H1377), CInt(&H44E6), &H8C, &H32, &H88, &HFD, &H5F, &H44, &HC8, &H4C)
IID_IDXGIDevice = iid
End Function
Public Function IID_IDXGIFactory1() As UUID
'{770aae78-f26f-4dba-a829-253c83d1b387}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H770AAE78, CInt(&HF26F), CInt(&H4DBA), &HA8, &H29, &H25, &H3C, &H83, &HD1, &HB3, &H87)
IID_IDXGIFactory1 = iid
End Function
Public Function IID_IDXGIAdapter1() As UUID
'{29038f61-3839-4626-91fd-086879011a05}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H29038F61, CInt(&H3839), CInt(&H4626), &H91, &HFD, &H8, &H68, &H79, &H1, &H1A, &H5)
IID_IDXGIAdapter1 = iid
End Function
Public Function IID_IDXGIDevice1() As UUID
'{77db970f-6276-48ba-ba28-070143b4392c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H77DB970F, CInt(&H6276), CInt(&H48BA), &HBA, &H28, &H7, &H1, &H43, &HB4, &H39, &H2C)
IID_IDXGIDevice1 = iid
End Function
Public Function IID_IDXGIDisplayControl() As UUID
'{ea9dbf1a-c88e-4486-854a-98aa0138f30c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEA9DBF1A, CInt(&HC88E), CInt(&H4486), &H85, &H4A, &H98, &HAA, &H1, &H38, &HF3, &HC)
IID_IDXGIDisplayControl = iid
End Function
Public Function IID_IDXGIOutputDuplication() As UUID
'{191cfac3-a341-470d-b26e-a864f428319c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H191CFAC3, CInt(&HA341), CInt(&H470D), &HB2, &H6E, &HA8, &H64, &HF4, &H28, &H31, &H9C)
IID_IDXGIOutputDuplication = iid
End Function
Public Function IID_IDXGISurface2() As UUID
'{aba496dd-b617-4cb8-a866-bc44d7eb1fa2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HABA496DD, CInt(&HB617), CInt(&H4CB8), &HA8, &H66, &HBC, &H44, &HD7, &HEB, &H1F, &HA2)
IID_IDXGISurface2 = iid
End Function
Public Function IID_IDXGIResource1() As UUID
'{30961379-4609-4a41-998e-54fe567ee0c1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H30961379, CInt(&H4609), CInt(&H4A41), &H99, &H8E, &H54, &HFE, &H56, &H7E, &HE0, &HC1)
IID_IDXGIResource1 = iid
End Function
Public Function IID_IDXGIDevice2() As UUID
'{05008617-fbfd-4051-a790-144884b4f6a9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5008617, CInt(&HFBFD), CInt(&H4051), &HA7, &H90, &H14, &H48, &H84, &HB4, &HF6, &HA9)
IID_IDXGIDevice2 = iid
End Function
Public Function IID_IDXGISwapChain1() As UUID
'{790a45f7-0d42-4876-983a-0a55cfe6f4aa}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H790A45F7, CInt(&HD42), CInt(&H4876), &H98, &H3A, &HA, &H55, &HCF, &HE6, &HF4, &HAA)
IID_IDXGISwapChain1 = iid
End Function
Public Function IID_IDXGIFactory2() As UUID
'{50c83a1c-e072-4c48-87b0-3630fa36a6d0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H50C83A1C, CInt(&HE072), CInt(&H4C48), &H87, &HB0, &H36, &H30, &HFA, &H36, &HA6, &HD0)
IID_IDXGIFactory2 = iid
End Function
Public Function IID_IDXGIAdapter2() As UUID
'{0AA1AE0A-FA0E-4B84-8644-E05FF8E5ACB5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAA1AE0A, CInt(&HFA0E), CInt(&H4B84), &H86, &H44, &HE0, &H5F, &HF8, &HE5, &HAC, &HB5)
IID_IDXGIAdapter2 = iid
End Function
Public Function IID_IDXGIOutput1() As UUID
'{00cddea8-939b-4b83-a340-a685226666cc}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCDDEA8, CInt(&H939B), CInt(&H4B83), &HA3, &H40, &HA6, &H85, &H22, &H66, &H66, &HCC)
IID_IDXGIOutput1 = iid
End Function
Public Function IID_IDXGIDevice3() As UUID
'{6007896c-3244-4afd-bf18-a6d3beda5023}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6007896C, CInt(&H3244), CInt(&H4AFD), &HBF, &H18, &HA6, &HD3, &HBE, &HDA, &H50, &H23)
IID_IDXGIDevice3 = iid
End Function
Public Function IID_IDXGISwapChain2() As UUID
'{a8be2ac4-199f-4946-b331-79599fb98de7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA8BE2AC4, CInt(&H199F), CInt(&H4946), &HB3, &H31, &H79, &H59, &H9F, &HB9, &H8D, &HE7)
IID_IDXGISwapChain2 = iid
End Function
Public Function IID_IDXGIOutput2() As UUID
'{595e39d1-2724-4663-99b1-da969de28364}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H595E39D1, CInt(&H2724), CInt(&H4663), &H99, &HB1, &HDA, &H96, &H9D, &HE2, &H83, &H64)
IID_IDXGIOutput2 = iid
End Function
Public Function IID_IDXGIFactory3() As UUID
'{25483823-cd46-4c7d-86ca-47aa95b837bd}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H25483823, CInt(&HCD46), CInt(&H4C7D), &H86, &HCA, &H47, &HAA, &H95, &HB8, &H37, &HBD)
IID_IDXGIFactory3 = iid
End Function
Public Function IID_IDXGIDecodeSwapChain() As UUID
'{2633066b-4514-4c7a-8fd8-12ea98059d18}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2633066B, CInt(&H4514), CInt(&H4C7A), &H8F, &HD8, &H12, &HEA, &H98, &H5, &H9D, &H18)
IID_IDXGIDecodeSwapChain = iid
End Function
Public Function IID_IDXGIFactoryMedia() As UUID
'{41e7d1f2-a591-4f7b-a2e5-fa9c843e1c12}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H41E7D1F2, CInt(&HA591), CInt(&H4F7B), &HA2, &HE5, &HFA, &H9C, &H84, &H3E, &H1C, &H12)
IID_IDXGIFactoryMedia = iid
End Function
Public Function IID_IDXGISwapChainMedia() As UUID
'{dd95b90b-f05f-4f6a-bd65-25bfb264bd84}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDD95B90B, CInt(&HF05F), CInt(&H4F6A), &HBD, &H65, &H25, &HBF, &HB2, &H64, &HBD, &H84)
IID_IDXGISwapChainMedia = iid
End Function
Public Function IID_IDXGIOutput3() As UUID
'{8a6bb301-7e7e-41F4-a8e0-5b32f7f99b18}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8A6BB301, CInt(&H7E7E), CInt(&H41F4), &HA8, &HE0, &H5B, &H32, &HF7, &HF9, &H9B, &H18)
IID_IDXGIOutput3 = iid
End Function
Public Function IID_IDXGISwapChain3() As UUID
'{94d99bdb-f1f8-4ab0-b236-7da0170edab1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H94D99BDB, CInt(&HF1F8), CInt(&H4AB0), &HB2, &H36, &H7D, &HA0, &H17, &HE, &HDA, &HB1)
IID_IDXGISwapChain3 = iid
End Function
Public Function IID_IDXGIOutput4() As UUID
'{dc7dca35-2196-414d-9F53-617884032a60}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDC7DCA35, CInt(&H2196), CInt(&H414D), &H9F, &H53, &H61, &H78, &H84, &H3, &H2A, &H60)
IID_IDXGIOutput4 = iid
End Function
Public Function IID_IDXGIFactory4() As UUID
'{1bc6ea02-ef36-464f-bf0c-21ca39e5168a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1BC6EA02, CInt(&HEF36), CInt(&H464F), &HBF, &HC, &H21, &HCA, &H39, &HE5, &H16, &H8A)
IID_IDXGIFactory4 = iid
End Function
Public Function IID_IDXGIAdapter3() As UUID
'{645967A4-1392-4310-A798-8053CE3E93FD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H645967A4, CInt(&H1392), CInt(&H4310), &HA7, &H98, &H80, &H53, &HCE, &H3E, &H93, &HFD)
IID_IDXGIAdapter3 = iid
End Function
Public Function IID_IDXGIOutput5() As UUID
'{80A07424-AB52-42EB-833C-0C42FD282D98}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H80A07424, CInt(&HAB52), CInt(&H42EB), &H83, &H3C, &HC, &H42, &HFD, &H28, &H2D, &H98)
IID_IDXGIOutput5 = iid
End Function
Public Function IID_IDXGISwapChain4() As UUID
'{3D585D5A-BD4A-489E-B1F4-3DBCB6452FFB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3D585D5A, CInt(&HBD4A), CInt(&H489E), &HB1, &HF4, &H3D, &HBC, &HB6, &H45, &H2F, &HFB)
IID_IDXGISwapChain4 = iid
End Function
Public Function IID_IDXGIDevice4() As UUID
'{95B4F95F-D8DA-4CA4-9EE6-3B76D5968A10}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H95B4F95F, CInt(&HD8DA), CInt(&H4CA4), &H9E, &HE6, &H3B, &H76, &HD5, &H96, &H8A, &H10)
IID_IDXGIDevice4 = iid
End Function
Public Function IID_IDXGIFactory5() As UUID
'{7632e1f5-ee65-4dca-87fd-84cd75f8838d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7632E1F5, CInt(&HEE65), CInt(&H4DCA), &H87, &HFD, &H84, &HCD, &H75, &HF8, &H83, &H8D)
IID_IDXGIFactory5 = iid
End Function
Public Function IID_IDXGIAdapter4() As UUID
'{3c8d99d1-4fbf-4181-a82c-af66bf7bd24e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3C8D99D1, CInt(&H4FBF), CInt(&H4181), &HA8, &H2C, &HAF, &H66, &HBF, &H7B, &HD2, &H4E)
IID_IDXGIAdapter4 = iid
End Function
Public Function IID_IDXGIOutput6() As UUID
'{068346e8-aaec-4b84-add7-137f513f77a1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H68346E8, CInt(&HAAEC), CInt(&H4B84), &HAD, &HD7, &H13, &H7F, &H51, &H3F, &H77, &HA1)
IID_IDXGIOutput6 = iid
End Function
Public Function IID_IPrintPreviewDxgiPackageTarget() As UUID
'{1a6dd0ad-1e2a-4e99-a5ba-91f17818290e}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1A6DD0AD, CInt(&H1E2A), CInt(&H4E99), &HA5, &HBA, &H91, &HF1, &H78, &H18, &H29, &HE)
 IID_IPrintPreviewDxgiPackageTarget = iid
End Function
Public Function IID_IPresentationBuffer() As UUID
'{2E217D3A-5ABB-4138-9A13-A775593C89CA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2E217D3A, CInt(&H5ABB), CInt(&H4138), &H9A, &H13, &HA7, &H75, &H59, &H3C, &H89, &HCA)
IID_IPresentationBuffer = iid
End Function
Public Function IID_IPresentationContent() As UUID
'{5668BB79-3D8E-415C-B215-F38020F2D252}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5668BB79, CInt(&H3D8E), CInt(&H415C), &HB2, &H15, &HF3, &H80, &H20, &HF2, &HD2, &H52)
IID_IPresentationContent = iid
End Function
Public Function IID_IPresentationSurface() As UUID
'{956710FB-EA40-4EBA-A3EB-4375A0EB4EDC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H956710FB, CInt(&HEA40), CInt(&H4EBA), &HA3, &HEB, &H43, &H75, &HA0, &HEB, &H4E, &HDC)
IID_IPresentationSurface = iid
End Function
Public Function IID_IPresentStatistics() As UUID
'{B44B8BDA-7282-495D-9DD7-CEADD8B4BB86}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB44B8BDA, CInt(&H7282), CInt(&H495D), &H9D, &HD7, &HCE, &HAD, &HD8, &HB4, &HBB, &H86)
IID_IPresentStatistics = iid
End Function
Public Function IID_IPresentationManager() As UUID
'{FB562F82-6292-470A-88B1-843661E7F20C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFB562F82, CInt(&H6292), CInt(&H470A), &H88, &HB1, &H84, &H36, &H61, &HE7, &HF2, &HC)
IID_IPresentationManager = iid
End Function
Public Function IID_IPresentationFactory() As UUID
'{8FB37B58-1D74-4F64-A49C-1F97A80A2EC0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8FB37B58, CInt(&H1D74), CInt(&H4F64), &HA4, &H9C, &H1F, &H97, &HA8, &HA, &H2E, &HC0)
IID_IPresentationFactory = iid
End Function
Public Function IID_IPresentStatusPresentStatistics() As UUID
'{C9ED2A41-79CB-435E-964E-C8553055420C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC9ED2A41, CInt(&H79CB), CInt(&H435E), &H96, &H4E, &HC8, &H55, &H30, &H55, &H42, &HC)
IID_IPresentStatusPresentStatistics = iid
End Function
Public Function IID_ICompositionFramePresentStatistics() As UUID
'{AB41D127-C101-4C0A-911D-F9F2E9D08E64}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB41D127, CInt(&HC101), CInt(&H4C0A), &H91, &H1D, &HF9, &HF2, &HE9, &HD0, &H8E, &H64)
IID_ICompositionFramePresentStatistics = iid
End Function
Public Function IID_IIndependentFlipFramePresentStatistics() As UUID
'{8C93BE27-AD94-4DA0-8FD4-2413132D124E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8C93BE27, CInt(&HAD94), CInt(&H4DA0), &H8F, &HD4, &H24, &H13, &H13, &H2D, &H12, &H4E)
IID_IIndependentFlipFramePresentStatistics = iid
End Function

Public Function IID_IDCompositionDevice() As UUID
'{C37EA93A-E7AA-450D-B16F-9746CB0407F3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC37EA93A, CInt(&HE7AA), CInt(&H450D), &HB1, &H6F, &H97, &H46, &HCB, &H4, &H7, &HF3)
IID_IDCompositionDevice = iid
End Function
Public Function IID_IDCompositionTarget() As UUID
'{eacdd04c-117e-4e17-88f4-d1b12b0e3d89}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEACDD04C, CInt(&H117E), CInt(&H4E17), &H88, &HF4, &HD1, &HB1, &H2B, &HE, &H3D, &H89)
IID_IDCompositionTarget = iid
End Function
Public Function IID_IDCompositionVisual() As UUID
'{4d93059d-097b-4651-9a60-f0f25116e2f3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4D93059D, CInt(&H97B), CInt(&H4651), &H9A, &H60, &HF0, &HF2, &H51, &H16, &HE2, &HF3)
IID_IDCompositionVisual = iid
End Function
Public Function IID_IDCompositionEffect() As UUID
'{EC81B08F-BFCB-4e8d-B193-A915587999E8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEC81B08F, CInt(&HBFCB), CInt(&H4E8D), &HB1, &H93, &HA9, &H15, &H58, &H79, &H99, &HE8)
IID_IDCompositionEffect = iid
End Function
Public Function IID_IDCompositionTransform3D() As UUID
'{71185722-246B-41f2-AAD1-0443F7F4BFC2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H71185722, CInt(&H246B), CInt(&H41F2), &HAA, &HD1, &H4, &H43, &HF7, &HF4, &HBF, &HC2)
IID_IDCompositionTransform3D = iid
End Function
Public Function IID_IDCompositionTransform() As UUID
'{FD55FAA7-37E0-4c20-95D2-9BE45BC33F55}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFD55FAA7, CInt(&H37E0), CInt(&H4C20), &H95, &HD2, &H9B, &HE4, &H5B, &HC3, &H3F, &H55)
IID_IDCompositionTransform = iid
End Function
Public Function IID_IDCompositionTranslateTransform() As UUID
'{06791122-C6F0-417d-8323-269E987F5954}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6791122, CInt(&HC6F0), CInt(&H417D), &H83, &H23, &H26, &H9E, &H98, &H7F, &H59, &H54)
IID_IDCompositionTranslateTransform = iid
End Function
Public Function IID_IDCompositionScaleTransform() As UUID
'{71FDE914-40EF-45ef-BD51-68B037C339F9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H71FDE914, CInt(&H40EF), CInt(&H45EF), &HBD, &H51, &H68, &HB0, &H37, &HC3, &H39, &HF9)
IID_IDCompositionScaleTransform = iid
End Function
Public Function IID_IDCompositionRotateTransform() As UUID
'{641ED83C-AE96-46c5-90DC-32774CC5C6D5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H641ED83C, CInt(&HAE96), CInt(&H46C5), &H90, &HDC, &H32, &H77, &H4C, &HC5, &HC6, &HD5)
IID_IDCompositionRotateTransform = iid
End Function
Public Function IID_IDCompositionSkewTransform() As UUID
'{E57AA735-DCDB-4c72-9C61-0591F58889EE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE57AA735, CInt(&HDCDB), CInt(&H4C72), &H9C, &H61, &H5, &H91, &HF5, &H88, &H89, &HEE)
IID_IDCompositionSkewTransform = iid
End Function
Public Function IID_IDCompositionMatrixTransform() As UUID
'{16CDFF07-C503-419c-83F2-0965C7AF1FA6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H16CDFF07, CInt(&HC503), CInt(&H419C), &H83, &HF2, &H9, &H65, &HC7, &HAF, &H1F, &HA6)
IID_IDCompositionMatrixTransform = iid
End Function
Public Function IID_IDCompositionEffectGroup() As UUID
'{A7929A74-E6B2-4bd6-8B95-4040119CA34D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA7929A74, CInt(&HE6B2), CInt(&H4BD6), &H8B, &H95, &H40, &H40, &H11, &H9C, &HA3, &H4D)
IID_IDCompositionEffectGroup = iid
End Function
Public Function IID_IDCompositionTranslateTransform3D() As UUID
'{91636D4B-9BA1-4532-AAF7-E3344994D788}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H91636D4B, CInt(&H9BA1), CInt(&H4532), &HAA, &HF7, &HE3, &H34, &H49, &H94, &HD7, &H88)
IID_IDCompositionTranslateTransform3D = iid
End Function
Public Function IID_IDCompositionScaleTransform3D() As UUID
'{2A9E9EAD-364B-4b15-A7C4-A1997F78B389}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2A9E9EAD, CInt(&H364B), CInt(&H4B15), &HA7, &HC4, &HA1, &H99, &H7F, &H78, &HB3, &H89)
IID_IDCompositionScaleTransform3D = iid
End Function
Public Function IID_IDCompositionRotateTransform3D() As UUID
'{D8F5B23F-D429-4a91-B55A-D2F45FD75B18}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD8F5B23F, CInt(&HD429), CInt(&H4A91), &HB5, &H5A, &HD2, &HF4, &H5F, &HD7, &H5B, &H18)
IID_IDCompositionRotateTransform3D = iid
End Function
Public Function IID_IDCompositionMatrixTransform3D() As UUID
'{4B3363F0-643B-41b7-B6E0-CCF22D34467C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4B3363F0, CInt(&H643B), CInt(&H41B7), &HB6, &HE0, &HCC, &HF2, &H2D, &H34, &H46, &H7C)
IID_IDCompositionMatrixTransform3D = iid
End Function
Public Function IID_IDCompositionClip() As UUID
'{64AC3703-9D3F-45ec-A109-7CAC0E7A13A7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H64AC3703, CInt(&H9D3F), CInt(&H45EC), &HA1, &H9, &H7C, &HAC, &HE, &H7A, &H13, &HA7)
IID_IDCompositionClip = iid
End Function
Public Function IID_IDCompositionRectangleClip() As UUID
'{9842AD7D-D9CF-4908-AED7-48B51DA5E7C2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9842AD7D, CInt(&HD9CF), CInt(&H4908), &HAE, &HD7, &H48, &HB5, &H1D, &HA5, &HE7, &HC2)
IID_IDCompositionRectangleClip = iid
End Function
Public Function IID_IDCompositionSurface() As UUID
'{BB8A4953-2C99-4F5A-96F5-4819027FA3AC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBB8A4953, CInt(&H2C99), CInt(&H4F5A), &H96, &HF5, &H48, &H19, &H2, &H7F, &HA3, &HAC)
IID_IDCompositionSurface = iid
End Function
Public Function IID_IDCompositionVirtualSurface() As UUID
'{AE471C51-5F53-4A24-8D3E-D0C39C30B3F0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAE471C51, CInt(&H5F53), CInt(&H4A24), &H8D, &H3E, &HD0, &HC3, &H9C, &H30, &HB3, &HF0)
IID_IDCompositionVirtualSurface = iid
End Function
Public Function IID_IDCompositionDevice2() As UUID
'{75F6468D-1B8E-447C-9BC6-75FEA80B5B25}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H75F6468D, CInt(&H1B8E), CInt(&H447C), &H9B, &HC6, &H75, &HFE, &HA8, &HB, &H5B, &H25)
IID_IDCompositionDevice2 = iid
End Function
Public Function IID_IDCompositionDesktopDevice() As UUID
'{5F4633FE-1E08-4CB8-8C75-CE24333F5602}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5F4633FE, CInt(&H1E08), CInt(&H4CB8), &H8C, &H75, &HCE, &H24, &H33, &H3F, &H56, &H2)
IID_IDCompositionDesktopDevice = iid
End Function
Public Function IID_IDCompositionDeviceDebug() As UUID
'{A1A3C64A-224F-4A81-9773-4F03A89D3C6C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA1A3C64A, CInt(&H224F), CInt(&H4A81), &H97, &H73, &H4F, &H3, &HA8, &H9D, &H3C, &H6C)
IID_IDCompositionDeviceDebug = iid
End Function
Public Function IID_IDCompositionSurfaceFactory() As UUID
'{E334BC12-3937-4E02-85EB-FCF4EB30D2C8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE334BC12, CInt(&H3937), CInt(&H4E02), &H85, &HEB, &HFC, &HF4, &HEB, &H30, &HD2, &HC8)
IID_IDCompositionSurfaceFactory = iid
End Function
Public Function IID_IDCompositionVisual2() As UUID
'{E8DE1639-4331-4B26-BC5F-6A321D347A85}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE8DE1639, CInt(&H4331), CInt(&H4B26), &HBC, &H5F, &H6A, &H32, &H1D, &H34, &H7A, &H85)
IID_IDCompositionVisual2 = iid
End Function
Public Function IID_IDCompositionVisualDebug() As UUID
'{FED2B808-5EB4-43A0-AEA3-35F65280F91B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFED2B808, CInt(&H5EB4), CInt(&H43A0), &HAE, &HA3, &H35, &HF6, &H52, &H80, &HF9, &H1B)
IID_IDCompositionVisualDebug = iid
End Function
Public Function IID_IDCompositionVisual3() As UUID
'{2775F462-B6C1-4015-B0BE-B3E7D6A4976D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2775F462, CInt(&HB6C1), CInt(&H4015), &HB0, &HBE, &HB3, &HE7, &HD6, &HA4, &H97, &H6D)
IID_IDCompositionVisual3 = iid
End Function
Public Function IID_IDCompositionDevice3() As UUID
'{0987CB06-F916-48BF-8D35-CE7641781BD9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H987CB06, CInt(&HF916), CInt(&H48BF), &H8D, &H35, &HCE, &H76, &H41, &H78, &H1B, &HD9)
IID_IDCompositionDevice3 = iid
End Function
Public Function IID_IDCompositionFilterEffect() As UUID
'{30C421D5-8CB2-4E9F-B133-37BE270D4AC2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H30C421D5, CInt(&H8CB2), CInt(&H4E9F), &HB1, &H33, &H37, &HBE, &H27, &HD, &H4A, &HC2)
IID_IDCompositionFilterEffect = iid
End Function
Public Function IID_IDCompositionGaussianBlurEffect() As UUID
'{45D4D0B7-1BD4-454E-8894-2BFA68443033}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H45D4D0B7, CInt(&H1BD4), CInt(&H454E), &H88, &H94, &H2B, &HFA, &H68, &H44, &H30, &H33)
IID_IDCompositionGaussianBlurEffect = iid
End Function
Public Function IID_IDCompositionBrightnessEffect() As UUID
'{6027496E-CB3A-49AB-934F-D798DA4F7DA6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6027496E, CInt(&HCB3A), CInt(&H49AB), &H93, &H4F, &HD7, &H98, &HDA, &H4F, &H7D, &HA6)
IID_IDCompositionBrightnessEffect = iid
End Function
Public Function IID_IDCompositionColorMatrixEffect() As UUID
'{C1170A22-3CE2-4966-90D4-55408BFC84C4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC1170A22, CInt(&H3CE2), CInt(&H4966), &H90, &HD4, &H55, &H40, &H8B, &HFC, &H84, &HC4)
IID_IDCompositionColorMatrixEffect = iid
End Function
Public Function IID_IDCompositionShadowEffect() As UUID
'{4AD18AC0-CFD2-4C2F-BB62-96E54FDB6879}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4AD18AC0, CInt(&HCFD2), CInt(&H4C2F), &HBB, &H62, &H96, &HE5, &H4F, &HDB, &H68, &H79)
IID_IDCompositionShadowEffect = iid
End Function
Public Function IID_IDCompositionHueRotationEffect() As UUID
'{6DB9F920-0770-4781-B0C6-381912F9D167}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6DB9F920, CInt(&H770), CInt(&H4781), &HB0, &HC6, &H38, &H19, &H12, &HF9, &HD1, &H67)
IID_IDCompositionHueRotationEffect = iid
End Function
Public Function IID_IDCompositionSaturationEffect() As UUID
'{A08DEBDA-3258-4FA4-9F16-9174D3FE93B1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA08DEBDA, CInt(&H3258), CInt(&H4FA4), &H9F, &H16, &H91, &H74, &HD3, &HFE, &H93, &HB1)
IID_IDCompositionSaturationEffect = iid
End Function
Public Function IID_IDCompositionTurbulenceEffect() As UUID
'{A6A55BDA-C09C-49F3-9193-A41922C89715}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA6A55BDA, CInt(&HC09C), CInt(&H49F3), &H91, &H93, &HA4, &H19, &H22, &HC8, &H97, &H15)
IID_IDCompositionTurbulenceEffect = iid
End Function
Public Function IID_IDCompositionLinearTransferEffect() As UUID
'{4305EE5B-C4A0-4C88-9385-67124E017683}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4305EE5B, CInt(&HC4A0), CInt(&H4C88), &H93, &H85, &H67, &H12, &H4E, &H1, &H76, &H83)
IID_IDCompositionLinearTransferEffect = iid
End Function
Public Function IID_IDCompositionTableTransferEffect() As UUID
'{9B7E82E2-69C5-4EB4-A5F5-A7033F5132CD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9B7E82E2, CInt(&H69C5), CInt(&H4EB4), &HA5, &HF5, &HA7, &H3, &H3F, &H51, &H32, &HCD)
IID_IDCompositionTableTransferEffect = iid
End Function
Public Function IID_IDCompositionCompositeEffect() As UUID
'{576616C0-A231-494D-A38D-00FD5EC4DB46}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H576616C0, CInt(&HA231), CInt(&H494D), &HA3, &H8D, &H0, &HFD, &H5E, &HC4, &HDB, &H46)
IID_IDCompositionCompositeEffect = iid
End Function
Public Function IID_IDCompositionBlendEffect() As UUID
'{33ECDC0A-578A-4A11-9C14-0CB90517F9C5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H33ECDC0A, CInt(&H578A), CInt(&H4A11), &H9C, &H14, &HC, &HB9, &H5, &H17, &HF9, &HC5)
IID_IDCompositionBlendEffect = iid
End Function
Public Function IID_IDCompositionArithmeticCompositeEffect() As UUID
'{3B67DFA8-E3DD-4E61-B640-46C2F3D739DC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3B67DFA8, CInt(&HE3DD), CInt(&H4E61), &HB6, &H40, &H46, &HC2, &HF3, &HD7, &H39, &HDC)
IID_IDCompositionArithmeticCompositeEffect = iid
End Function
Public Function IID_IDCompositionAffineTransform2DEffect() As UUID
'{0B74B9E8-CDD6-492F-BBBC-5ED32157026D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB74B9E8, CInt(&HCDD6), CInt(&H492F), &HBB, &HBC, &H5E, &HD3, &H21, &H57, &H2, &H6D)
IID_IDCompositionAffineTransform2DEffect = iid
End Function
Public Function IID_IDCompositionAnimation() As UUID
'{CBFD91D9-51B2-45e4-B3DE-D19CCFB863C5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCBFD91D9, CInt(&H51B2), CInt(&H45E4), &HB3, &HDE, &HD1, &H9C, &HCF, &HB8, &H63, &HC5)
IID_IDCompositionAnimation = iid
End Function


Public Function IID_ID3D10Blob() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8BA5FB08, &H5195, &H40E2, &HAC, &H58, &HD, &H98, &H9C, &H3A, &H1, &H2)
IID_ID3D10Blob = iid
End Function
Public Function WKPDID_D3DDebugObjectName() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H429B8C22, &H9188, &H4B0C, &H87, &H42, &HAC, &HB0, &HBF, &H85, &HC2, &H0)
WKPDID_D3DDebugObjectName = iid
End Function
Public Function WKPDID_D3DDebugObjectNameW() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4CCA5FD8, &H921F, &H42C8, &H85, &H66, &H70, &HCA, &HF2, &HA9, &HB7, &H41)
WKPDID_D3DDebugObjectNameW = iid
End Function
Public Function WKPDID_CommentStringW() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD0149DC0, &H90E8, &H4EC8, &H81, &H44, &HE9, &H0, &HAD, &H26, &H6B, &HB2)
WKPDID_CommentStringW = iid
End Function
Public Function D3D11_DECODER_PROFILE_MPEG2_MOCOMP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE6A9F44B, &H61B0, &H4563, &H9E, &HA4, &H63, &HD2, &HA3, &HC6, &HFE, &H66)
D3D11_DECODER_PROFILE_MPEG2_MOCOMP = iid
End Function
Public Function D3D11_DECODER_PROFILE_MPEG2_IDCT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBF22AD00, &H3EA, &H4690, &H80, &H77, &H47, &H33, &H46, &H20, &H9B, &H7E)
D3D11_DECODER_PROFILE_MPEG2_IDCT = iid
End Function
Public Function D3D11_DECODER_PROFILE_MPEG2_VLD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEE27417F, &H5E28, &H4E65, &HBE, &HEA, &H1D, &H26, &HB5, &H8, &HAD, &HC9)
D3D11_DECODER_PROFILE_MPEG2_VLD = iid
End Function
Public Function D3D11_DECODER_PROFILE_MPEG1_VLD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6F3EC719, &H3735, &H42CC, &H80, &H63, &H65, &HCC, &H3C, &HB3, &H66, &H16)
D3D11_DECODER_PROFILE_MPEG1_VLD = iid
End Function
Public Function D3D11_DECODER_PROFILE_MPEG2and1_VLD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H86695F12, &H340E, &H4F04, &H9F, &HD3, &H92, &H53, &HDD, &H32, &H74, &H60)
D3D11_DECODER_PROFILE_MPEG2and1_VLD = iid
End Function
Public Function D3D11_DECODER_PROFILE_H264_MOCOMP_NOFGT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B81BE64, &HA0C7, &H11D3, &HB9, &H84, &H0, &HC0, &H4F, &H2E, &H73, &HC5)
D3D11_DECODER_PROFILE_H264_MOCOMP_NOFGT = iid
End Function
Public Function D3D11_DECODER_PROFILE_H264_MOCOMP_FGT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B81BE65, &HA0C7, &H11D3, &HB9, &H84, &H0, &HC0, &H4F, &H2E, &H73, &HC5)
D3D11_DECODER_PROFILE_H264_MOCOMP_FGT = iid
End Function
Public Function D3D11_DECODER_PROFILE_H264_IDCT_NOFGT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B81BE66, &HA0C7, &H11D3, &HB9, &H84, &H0, &HC0, &H4F, &H2E, &H73, &HC5)
D3D11_DECODER_PROFILE_H264_IDCT_NOFGT = iid
End Function
Public Function D3D11_DECODER_PROFILE_H264_IDCT_FGT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B81BE67, &HA0C7, &H11D3, &HB9, &H84, &H0, &HC0, &H4F, &H2E, &H73, &HC5)
D3D11_DECODER_PROFILE_H264_IDCT_FGT = iid
End Function
Public Function D3D11_DECODER_PROFILE_H264_VLD_NOFGT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B81BE68, &HA0C7, &H11D3, &HB9, &H84, &H0, &HC0, &H4F, &H2E, &H73, &HC5)
D3D11_DECODER_PROFILE_H264_VLD_NOFGT = iid
End Function
Public Function D3D11_DECODER_PROFILE_H264_VLD_FGT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B81BE69, &HA0C7, &H11D3, &HB9, &H84, &H0, &HC0, &H4F, &H2E, &H73, &HC5)
D3D11_DECODER_PROFILE_H264_VLD_FGT = iid
End Function
Public Function D3D11_DECODER_PROFILE_H264_VLD_WITHFMOASO_NOFGT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD5F04FF9, &H3418, &H45D8, &H95, &H61, &H32, &HA7, &H6A, &HAE, &H2D, &HDD)
D3D11_DECODER_PROFILE_H264_VLD_WITHFMOASO_NOFGT = iid
End Function
Public Function D3D11_DECODER_PROFILE_H264_VLD_STEREO_PROGRESSIVE_NOFGT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD79BE8DA, &HCF1, &H4C81, &HB8, &H2A, &H69, &HA4, &HE2, &H36, &HF4, &H3D)
D3D11_DECODER_PROFILE_H264_VLD_STEREO_PROGRESSIVE_NOFGT = iid
End Function
Public Function D3D11_DECODER_PROFILE_H264_VLD_STEREO_NOFGT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF9AACCBB, &HC2B6, &H4CFC, &H87, &H79, &H57, &H7, &HB1, &H76, &H5, &H52)
D3D11_DECODER_PROFILE_H264_VLD_STEREO_NOFGT = iid
End Function
Public Function D3D11_DECODER_PROFILE_H264_VLD_MULTIVIEW_NOFGT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H705B9D82, &H76CF, &H49D6, &HB7, &HE6, &HAC, &H88, &H72, &HDB, &H1, &H3C)
D3D11_DECODER_PROFILE_H264_VLD_MULTIVIEW_NOFGT = iid
End Function
Public Function D3D11_DECODER_PROFILE_WMV8_POSTPROC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B81BE80, &HA0C7, &H11D3, &HB9, &H84, &H0, &HC0, &H4F, &H2E, &H73, &HC5)
D3D11_DECODER_PROFILE_WMV8_POSTPROC = iid
End Function
Public Function D3D11_DECODER_PROFILE_WMV8_MOCOMP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B81BE81, &HA0C7, &H11D3, &HB9, &H84, &H0, &HC0, &H4F, &H2E, &H73, &HC5)
D3D11_DECODER_PROFILE_WMV8_MOCOMP = iid
End Function
Public Function D3D11_DECODER_PROFILE_WMV9_POSTPROC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B81BE90, &HA0C7, &H11D3, &HB9, &H84, &H0, &HC0, &H4F, &H2E, &H73, &HC5)
D3D11_DECODER_PROFILE_WMV9_POSTPROC = iid
End Function
Public Function D3D11_DECODER_PROFILE_WMV9_MOCOMP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B81BE91, &HA0C7, &H11D3, &HB9, &H84, &H0, &HC0, &H4F, &H2E, &H73, &HC5)
D3D11_DECODER_PROFILE_WMV9_MOCOMP = iid
End Function
Public Function D3D11_DECODER_PROFILE_WMV9_IDCT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B81BE94, &HA0C7, &H11D3, &HB9, &H84, &H0, &HC0, &H4F, &H2E, &H73, &HC5)
D3D11_DECODER_PROFILE_WMV9_IDCT = iid
End Function
Public Function D3D11_DECODER_PROFILE_VC1_POSTPROC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B81BEA0, &HA0C7, &H11D3, &HB9, &H84, &H0, &HC0, &H4F, &H2E, &H73, &HC5)
D3D11_DECODER_PROFILE_VC1_POSTPROC = iid
End Function
Public Function D3D11_DECODER_PROFILE_VC1_MOCOMP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B81BEA1, &HA0C7, &H11D3, &HB9, &H84, &H0, &HC0, &H4F, &H2E, &H73, &HC5)
D3D11_DECODER_PROFILE_VC1_MOCOMP = iid
End Function
Public Function D3D11_DECODER_PROFILE_VC1_IDCT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B81BEA2, &HA0C7, &H11D3, &HB9, &H84, &H0, &HC0, &H4F, &H2E, &H73, &HC5)
D3D11_DECODER_PROFILE_VC1_IDCT = iid
End Function
Public Function D3D11_DECODER_PROFILE_VC1_VLD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B81BEA3, &HA0C7, &H11D3, &HB9, &H84, &H0, &HC0, &H4F, &H2E, &H73, &HC5)
D3D11_DECODER_PROFILE_VC1_VLD = iid
End Function
Public Function D3D11_DECODER_PROFILE_VC1_D2010() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1B81BEA4, &HA0C7, &H11D3, &HB9, &H84, &H0, &HC0, &H4F, &H2E, &H73, &HC5)
D3D11_DECODER_PROFILE_VC1_D2010 = iid
End Function
Public Function D3D11_DECODER_PROFILE_MPEG4PT2_VLD_SIMPLE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEFD64D74, &HC9E8, &H41D7, &HA5, &HE9, &HE9, &HB0, &HE3, &H9F, &HA3, &H19)
D3D11_DECODER_PROFILE_MPEG4PT2_VLD_SIMPLE = iid
End Function
Public Function D3D11_DECODER_PROFILE_MPEG4PT2_VLD_ADVSIMPLE_NOGMC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HED418A9F, &H10D, &H4EDA, &H9A, &HE3, &H9A, &H65, &H35, &H8D, &H8D, &H2E)
D3D11_DECODER_PROFILE_MPEG4PT2_VLD_ADVSIMPLE_NOGMC = iid
End Function
Public Function D3D11_DECODER_PROFILE_MPEG4PT2_VLD_ADVSIMPLE_GMC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAB998B5B, &H4258, &H44A9, &H9F, &HEB, &H94, &HE5, &H97, &HA6, &HBA, &HAE)
D3D11_DECODER_PROFILE_MPEG4PT2_VLD_ADVSIMPLE_GMC = iid
End Function
Public Function D3D11_DECODER_PROFILE_HEVC_VLD_MAIN() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5B11D51B, &H2F4C, &H4452, &HBC, &HC3, &H9, &HF2, &HA1, &H16, &HC, &HC0)
D3D11_DECODER_PROFILE_HEVC_VLD_MAIN = iid
End Function
Public Function D3D11_DECODER_PROFILE_HEVC_VLD_MAIN10() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H107AF0E0, &HEF1A, &H4D19, &HAB, &HA8, &H67, &HA1, &H63, &H7, &H3D, &H13)
D3D11_DECODER_PROFILE_HEVC_VLD_MAIN10 = iid
End Function
Public Function D3D11_DECODER_PROFILE_VP9_VLD_PROFILE0() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H463707F8, &HA1D0, &H4585, &H87, &H6D, &H83, &HAA, &H6D, &H60, &HB8, &H9E)
D3D11_DECODER_PROFILE_VP9_VLD_PROFILE0 = iid
End Function
Public Function D3D11_DECODER_PROFILE_VP8_VLD() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H90B899EA, &H3A62, &H4705, &H88, &HB3, &H8D, &HF0, &H4B, &H27, &H44, &HE7)
D3D11_DECODER_PROFILE_VP8_VLD = iid
End Function
Public Function D3D11_CRYPTO_TYPE_AES128_CTR() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9B6BD711, &H4F74, &H41C9, &H9E, &H7B, &HB, &HE2, &HD7, &HD9, &H3B, &H4F)
D3D11_CRYPTO_TYPE_AES128_CTR = iid
End Function
Public Function D3D11_DECODER_ENCRYPTION_HW_CENC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H89D6AC4F, &H9F2, &H4229, &HB2, &HCD, &H37, &H74, &HA, &H6D, &HFD, &H81)
D3D11_DECODER_ENCRYPTION_HW_CENC = iid
End Function
Public Function D3D11_KEY_EXCHANGE_HW_PROTECTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB1170D8A, &H628D, &H4DA3, &HAD, &H3B, &H82, &HDD, &HB0, &H8B, &H49, &H70)
D3D11_KEY_EXCHANGE_HW_PROTECTION = iid
End Function
Public Function D3D11_AUTHENTICATED_QUERY_PROTECTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA84EB584, &HC495, &H48AA, &HB9, &H4D, &H8B, &HD2, &HD6, &HFB, &HCE, &H5)
D3D11_AUTHENTICATED_QUERY_PROTECTION = iid
End Function
Public Function D3D11_AUTHENTICATED_QUERY_CHANNEL_TYPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBC1B18A5, &HB1FB, &H42AB, &HBD, &H94, &HB5, &H82, &H8B, &H4B, &HF7, &HBE)
D3D11_AUTHENTICATED_QUERY_CHANNEL_TYPE = iid
End Function
Public Function D3D11_AUTHENTICATED_QUERY_DEVICE_HANDLE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEC1C539D, &H8CFF, &H4E2A, &HBC, &HC4, &HF5, &H69, &H2F, &H99, &HF4, &H80)
D3D11_AUTHENTICATED_QUERY_DEVICE_HANDLE = iid
End Function
Public Function D3D11_AUTHENTICATED_QUERY_CRYPTO_SESSION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2634499E, &HD018, &H4D74, &HAC, &H17, &H7F, &H72, &H40, &H59, &H52, &H8D)
D3D11_AUTHENTICATED_QUERY_CRYPTO_SESSION = iid
End Function
Public Function D3D11_AUTHENTICATED_QUERY_RESTRICTED_SHARED_RESOURCE_PROCESS_COUNT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDB207B3, &H9450, &H46A6, &H82, &HDE, &H1B, &H96, &HD4, &H4F, &H9C, &HF2)
D3D11_AUTHENTICATED_QUERY_RESTRICTED_SHARED_RESOURCE_PROCESS_COUNT = iid
End Function
Public Function D3D11_AUTHENTICATED_QUERY_RESTRICTED_SHARED_RESOURCE_PROCESS() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H649BBADB, &HF0F4, &H4639, &HA1, &H5B, &H24, &H39, &H3F, &HC3, &HAB, &HAC)
D3D11_AUTHENTICATED_QUERY_RESTRICTED_SHARED_RESOURCE_PROCESS = iid
End Function
Public Function D3D11_AUTHENTICATED_QUERY_UNRESTRICTED_PROTECTED_SHARED_RESOURCE_COUNT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H12F0BD6, &HE662, &H4474, &HBE, &HFD, &HAA, &H53, &HE5, &H14, &H3C, &H6D)
D3D11_AUTHENTICATED_QUERY_UNRESTRICTED_PROTECTED_SHARED_RESOURCE_COUNT = iid
End Function
Public Function D3D11_AUTHENTICATED_QUERY_OUTPUT_ID_COUNT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2C042B5E, &H8C07, &H46D5, &HAA, &HBE, &H8F, &H75, &HCB, &HAD, &H4C, &H31)
D3D11_AUTHENTICATED_QUERY_OUTPUT_ID_COUNT = iid
End Function
Public Function D3D11_AUTHENTICATED_QUERY_OUTPUT_ID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H839DDCA3, &H9B4E, &H41E4, &HB0, &H53, &H89, &H2B, &HD2, &HA1, &H1E, &HE7)
D3D11_AUTHENTICATED_QUERY_OUTPUT_ID = iid
End Function
Public Function D3D11_AUTHENTICATED_QUERY_ACCESSIBILITY_ATTRIBUTES() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6214D9D2, &H432C, &H4ABB, &H9F, &HCE, &H21, &H6E, &HEA, &H26, &H9E, &H3B)
D3D11_AUTHENTICATED_QUERY_ACCESSIBILITY_ATTRIBUTES = iid
End Function
Public Function D3D11_AUTHENTICATED_QUERY_ENCRYPTION_WHEN_ACCESSIBLE_GUID_COUNT() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB30F7066, &H203C, &H4B07, &H93, &HFC, &HCE, &HAA, &HFD, &H61, &H24, &H1E)
D3D11_AUTHENTICATED_QUERY_ENCRYPTION_WHEN_ACCESSIBLE_GUID_COUNT = iid
End Function
Public Function D3D11_AUTHENTICATED_QUERY_ENCRYPTION_WHEN_ACCESSIBLE_GUID() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF83A5958, &HE986, &H4BDA, &HBE, &HB0, &H41, &H1F, &H6A, &H7A, &H1, &HB7)
D3D11_AUTHENTICATED_QUERY_ENCRYPTION_WHEN_ACCESSIBLE_GUID = iid
End Function
Public Function D3D11_AUTHENTICATED_QUERY_CURRENT_ENCRYPTION_WHEN_ACCESSIBLE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEC1791C7, &HDAD3, &H4F15, &H9E, &HC3, &HFA, &HA9, &H3D, &H60, &HD4, &HF0)
D3D11_AUTHENTICATED_QUERY_CURRENT_ENCRYPTION_WHEN_ACCESSIBLE = iid
End Function
Public Function D3D11_AUTHENTICATED_CONFIGURE_INITIALIZE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6114BDB, &H3523, &H470A, &H8D, &HCA, &HFB, &HC2, &H84, &H51, &H54, &HF0)
D3D11_AUTHENTICATED_CONFIGURE_INITIALIZE = iid
End Function
Public Function D3D11_AUTHENTICATED_CONFIGURE_PROTECTION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H50455658, &H3F47, &H4362, &HBF, &H99, &HBF, &HDF, &HCD, &HE9, &HED, &H29)
D3D11_AUTHENTICATED_CONFIGURE_PROTECTION = iid
End Function
Public Function D3D11_AUTHENTICATED_CONFIGURE_CRYPTO_SESSION() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6346CC54, &H2CFC, &H4AD4, &H82, &H24, &HD1, &H58, &H37, &HDE, &H77, &H0)
D3D11_AUTHENTICATED_CONFIGURE_CRYPTO_SESSION = iid
End Function
Public Function D3D11_AUTHENTICATED_CONFIGURE_SHARED_RESOURCE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H772D047, &H1B40, &H48E8, &H9C, &HA6, &HB5, &HF5, &H10, &HDE, &H9F, &H1)
D3D11_AUTHENTICATED_CONFIGURE_SHARED_RESOURCE = iid
End Function
Public Function D3D11_AUTHENTICATED_CONFIGURE_ENCRYPTION_WHEN_ACCESSIBLE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H41FFF286, &H6AE0, &H4D43, &H9D, &H55, &HA4, &H6E, &H9E, &HFD, &H15, &H8A)
D3D11_AUTHENTICATED_CONFIGURE_ENCRYPTION_WHEN_ACCESSIBLE = iid
End Function
Public Function D3D11_KEY_EXCHANGE_RSAES_OAEP() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC1949895, &HD72A, &H4A1D, &H8E, &H5D, &HED, &H85, &H7D, &H17, &H15, &H20)
D3D11_KEY_EXCHANGE_RSAES_OAEP = iid
End Function
Public Function IID_ID3D11DeviceChild() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1841E5C8, &H16B0, &H489B, &HBC, &HC8, &H44, &HCF, &HB0, &HD5, &HDE, &HAE)
IID_ID3D11DeviceChild = iid
End Function
Public Function IID_ID3D11DepthStencilState() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3823EFB, &H8D8F, &H4E1C, &H9A, &HA2, &HF6, &H4B, &HB2, &HCB, &HFD, &HF1)
IID_ID3D11DepthStencilState = iid
End Function
Public Function IID_ID3D11BlendState() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H75B68FAA, &H347D, &H4159, &H8F, &H45, &HA0, &H64, &HF, &H1, &HCD, &H9A)
IID_ID3D11BlendState = iid
End Function
Public Function IID_ID3D11RasterizerState() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9BB4AB81, &HAB1A, &H4D8F, &HB5, &H6, &HFC, &H4, &H20, &HB, &H6E, &HE7)
IID_ID3D11RasterizerState = iid
End Function
Public Function IID_ID3D11Resource() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDC8E63F3, &HD12B, &H4952, &HB4, &H7B, &H5E, &H45, &H2, &H6A, &H86, &H2D)
IID_ID3D11Resource = iid
End Function
Public Function IID_ID3D11Buffer() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H48570B85, &HD1EE, &H4FCD, &HA2, &H50, &HEB, &H35, &H7, &H22, &HB0, &H37)
IID_ID3D11Buffer = iid
End Function
Public Function IID_ID3D11Texture1D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF8FB5C27, &HC6B3, &H4F75, &HA4, &HC8, &H43, &H9A, &HF2, &HEF, &H56, &H4C)
IID_ID3D11Texture1D = iid
End Function
Public Function IID_ID3D11Texture2D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6F15AAF2, &HD208, &H4E89, &H9A, &HB4, &H48, &H95, &H35, &HD3, &H4F, &H9C)
IID_ID3D11Texture2D = iid
End Function
Public Function IID_ID3D11Texture3D() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H37E866E, &HF56D, &H4357, &HA8, &HAF, &H9D, &HAB, &HBE, &H6E, &H25, &HE)
IID_ID3D11Texture3D = iid
End Function
Public Function IID_ID3D11View() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H839D1216, &HBB2E, &H412B, &HB7, &HF4, &HA9, &HDB, &HEB, &HE0, &H8E, &HD1)
IID_ID3D11View = iid
End Function
Public Function IID_ID3D11ShaderResourceView() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB0E06FE0, &H8192, &H4E1A, &HB1, &HCA, &H36, &HD7, &H41, &H47, &H10, &HB2)
IID_ID3D11ShaderResourceView = iid
End Function
Public Function IID_ID3D11RenderTargetView() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDFDBA067, &HB8D, &H4865, &H87, &H5B, &HD7, &HB4, &H51, &H6C, &HC1, &H64)
IID_ID3D11RenderTargetView = iid
End Function
Public Function IID_ID3D11DepthStencilView() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9FDAC92A, &H1876, &H48C3, &HAF, &HAD, &H25, &HB9, &H4F, &H84, &HA9, &HB6)
IID_ID3D11DepthStencilView = iid
End Function
Public Function IID_ID3D11UnorderedAccessView() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H28ACF509, &H7F5C, &H48F6, &H86, &H11, &HF3, &H16, &H1, &HA, &H63, &H80)
IID_ID3D11UnorderedAccessView = iid
End Function
Public Function IID_ID3D11VertexShader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3B301D64, &HD678, &H4289, &H88, &H97, &H22, &HF8, &H92, &H8B, &H72, &HF3)
IID_ID3D11VertexShader = iid
End Function
Public Function IID_ID3D11HullShader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8E5C6061, &H628A, &H4C8E, &H82, &H64, &HBB, &HE4, &H5C, &HB3, &HD5, &HDD)
IID_ID3D11HullShader = iid
End Function
Public Function IID_ID3D11DomainShader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF582C508, &HF36, &H490C, &H99, &H77, &H31, &HEE, &HCE, &H26, &H8C, &HFA)
IID_ID3D11DomainShader = iid
End Function
Public Function IID_ID3D11GeometryShader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H38325B96, &HEFFB, &H4022, &HBA, &H2, &H2E, &H79, &H5B, &H70, &H27, &H5C)
IID_ID3D11GeometryShader = iid
End Function
Public Function IID_ID3D11PixelShader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEA82E40D, &H51DC, &H4F33, &H93, &HD4, &HDB, &H7C, &H91, &H25, &HAE, &H8C)
IID_ID3D11PixelShader = iid
End Function
Public Function IID_ID3D11ComputeShader() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4F5B196E, &HC2BD, &H495E, &HBD, &H1, &H1F, &HDE, &HD3, &H8E, &H49, &H69)
IID_ID3D11ComputeShader = iid
End Function
Public Function IID_ID3D11InputLayout() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE4819DDC, &H4CF0, &H4025, &HBD, &H26, &H5D, &HE8, &H2A, &H3E, &H7, &HB7)
IID_ID3D11InputLayout = iid
End Function
Public Function IID_ID3D11SamplerState() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDA6FEA51, &H564C, &H4487, &H98, &H10, &HF0, &HD0, &HF9, &HB4, &HE3, &HA5)
IID_ID3D11SamplerState = iid
End Function
Public Function IID_ID3D11Asynchronous() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4B35D0CD, &H1E15, &H4258, &H9C, &H98, &H1B, &H13, &H33, &HF6, &HDD, &H3B)
IID_ID3D11Asynchronous = iid
End Function
Public Function IID_ID3D11Query() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD6C00747, &H87B7, &H425E, &HB8, &H4D, &H44, &HD1, &H8, &H56, &HA, &HFD)
IID_ID3D11Query = iid
End Function
Public Function IID_ID3D11Predicate() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9EB576DD, &H9F77, &H4D86, &H81, &HAA, &H8B, &HAB, &H5F, &HE4, &H90, &HE2)
IID_ID3D11Predicate = iid
End Function
Public Function IID_ID3D11Counter() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6E8C49FB, &HA371, &H4770, &HB4, &H40, &H29, &H8, &H60, &H22, &HB7, &H41)
IID_ID3D11Counter = iid
End Function
Public Function IID_ID3D11ClassInstance() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA6CD7FAA, &HB0B7, &H4A2F, &H94, &H36, &H86, &H62, &HA6, &H57, &H97, &HCB)
IID_ID3D11ClassInstance = iid
End Function
Public Function IID_ID3D11ClassLinkage() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDDF57CBA, &H9543, &H46E4, &HA1, &H2B, &HF2, &H7, &HA0, &HFE, &H7F, &HED)
IID_ID3D11ClassLinkage = iid
End Function
Public Function IID_ID3D11CommandList() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA24BC4D1, &H769E, &H43F7, &H80, &H13, &H98, &HFF, &H56, &H6C, &H18, &HE2)
IID_ID3D11CommandList = iid
End Function
Public Function IID_ID3D11DeviceContext() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC0BFA96C, &HE089, &H44FB, &H8E, &HAF, &H26, &HF8, &H79, &H61, &H90, &HDA)
IID_ID3D11DeviceContext = iid
End Function
Public Function IID_ID3D11VideoDecoder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3C9C5B51, &H995D, &H48D1, &H9B, &H8D, &HFA, &H5C, &HAE, &HDE, &HD6, &H5C)
IID_ID3D11VideoDecoder = iid
End Function
Public Function IID_ID3D11VideoProcessorEnumerator() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H31627037, &H53AB, &H4200, &H90, &H61, &H5, &HFA, &HA9, &HAB, &H45, &HF9)
IID_ID3D11VideoProcessorEnumerator = iid
End Function
Public Function IID_ID3D11VideoProcessor() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1D7B0652, &H185F, &H41C6, &H85, &HCE, &HC, &H5B, &HE3, &HD4, &HAE, &H6C)
IID_ID3D11VideoProcessor = iid
End Function
Public Function IID_ID3D11AuthenticatedChannel() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3015A308, &HDCBD, &H47AA, &HA7, &H47, &H19, &H24, &H86, &HD1, &H4D, &H4A)
IID_ID3D11AuthenticatedChannel = iid
End Function
Public Function IID_ID3D11CryptoSession() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9B32F9AD, &HBDCC, &H40A6, &HA3, &H9D, &HD5, &HC8, &H65, &H84, &H57, &H20)
IID_ID3D11CryptoSession = iid
End Function
Public Function IID_ID3D11VideoDecoderOutputView() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC2931AEA, &H2A85, &H4F20, &H86, &HF, &HFB, &HA1, &HFD, &H25, &H6E, &H18)
IID_ID3D11VideoDecoderOutputView = iid
End Function
Public Function IID_ID3D11VideoProcessorInputView() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H11EC5A5F, &H51DC, &H4945, &HAB, &H34, &H6E, &H8C, &H21, &H30, &HE, &HA5)
IID_ID3D11VideoProcessorInputView = iid
End Function
Public Function IID_ID3D11VideoProcessorOutputView() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA048285E, &H25A9, &H4527, &HBD, &H93, &HD6, &H8B, &H68, &HC4, &H42, &H54)
IID_ID3D11VideoProcessorOutputView = iid
End Function
Public Function IID_ID3D11VideoContext() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H61F21C45, &H3C0E, &H4A74, &H9C, &HEA, &H67, &H10, &HD, &H9A, &HD5, &HE4)
IID_ID3D11VideoContext = iid
End Function
Public Function IID_ID3D11VideoDevice() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H10EC4D5B, &H975A, &H4689, &HB9, &HE4, &HD0, &HAA, &HC3, &HF, &HE3, &H33)
IID_ID3D11VideoDevice = iid
End Function
Public Function IID_ID3D11Device() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDB6F6DDB, &HAC77, &H4E88, &H82, &H53, &H81, &H9D, &HF9, &HBB, &HF1, &H40)
IID_ID3D11Device = iid
End Function
Public Function IID_ID3D11BlendState1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCC86FABE, &HDA55, &H401D, &H85, &HE7, &HE3, &HC9, &HDE, &H28, &H77, &HE9)
IID_ID3D11BlendState1 = iid
End Function
Public Function IID_ID3D11RasterizerState1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1217D7A6, &H5039, &H418C, &HB0, &H42, &H9C, &HBE, &H25, &H6A, &HFD, &H6E)
IID_ID3D11RasterizerState1 = iid
End Function
Public Function IID_ID3DDeviceContextState() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5C1E0D8A, &H7C23, &H48F9, &H8C, &H59, &HA9, &H29, &H58, &HCE, &HFF, &H11)
IID_ID3DDeviceContextState = iid
End Function
Public Function IID_ID3D11DeviceContext1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBB2C6FAA, &HB5FB, &H4082, &H8E, &H6B, &H38, &H8B, &H8C, &HFA, &H90, &HE1)
IID_ID3D11DeviceContext1 = iid
End Function
Public Function IID_ID3D11VideoContext1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA7F026DA, &HA5F8, &H4487, &HA5, &H64, &H15, &HE3, &H43, &H57, &H65, &H1E)
IID_ID3D11VideoContext1 = iid
End Function
Public Function IID_ID3D11VideoDevice1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H29DA1D51, &H1321, &H4454, &H80, &H4B, &HF5, &HFC, &H9F, &H86, &H1F, &HF)
IID_ID3D11VideoDevice1 = iid
End Function
Public Function IID_ID3D11VideoProcessorEnumerator1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H465217F2, &H5568, &H43CF, &HB5, &HB9, &HF6, &H1D, &H54, &H53, &H1C, &HA1)
IID_ID3D11VideoProcessorEnumerator1 = iid
End Function
Public Function IID_ID3D11Device1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA04BFB29, &H8EF, &H43D6, &HA4, &H9C, &HA9, &HBD, &HBD, &HCB, &HE6, &H86)
IID_ID3D11Device1 = iid
End Function
Public Function IID_ID3DUserDefinedAnnotation() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB2DAAD8B, &H3D4, &H4DBF, &H95, &HEB, &H32, &HAB, &H4B, &H63, &HD0, &HAB)
IID_ID3DUserDefinedAnnotation = iid
End Function
Public Function IID_ID3D11DeviceContext2() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H420D5B32, &HB90C, &H4DA4, &HBE, &HF0, &H35, &H9F, &H6A, &H24, &HA8, &H3A)
IID_ID3D11DeviceContext2 = iid
End Function
Public Function IID_ID3D11Device2() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9D06DFFA, &HD1E5, &H4D07, &H83, &HA8, &H1B, &HB1, &H23, &HF2, &HF8, &H41)
IID_ID3D11Device2 = iid
End Function
Public Function IID_ID3D11Texture2D1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H51218251, &H1E33, &H4617, &H9C, &HCB, &H4D, &H3A, &H43, &H67, &HE7, &HBB)
IID_ID3D11Texture2D1 = iid
End Function
Public Function IID_ID3D11Texture3D1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC711683, &H2853, &H4846, &H9B, &HB0, &HF3, &HE6, &H6, &H39, &HE4, &H6A)
IID_ID3D11Texture3D1 = iid
End Function
Public Function IID_ID3D11RasterizerState2() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6FBD02FB, &H209F, &H46C4, &HB0, &H59, &H2E, &HD1, &H55, &H86, &HA6, &HAC)
IID_ID3D11RasterizerState2 = iid
End Function
Public Function IID_ID3D11ShaderResourceView1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H91308B87, &H9040, &H411D, &H8C, &H67, &HC3, &H92, &H53, &HCE, &H38, &H2)
IID_ID3D11ShaderResourceView1 = iid
End Function
Public Function IID_ID3D11RenderTargetView1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFFBE2E23, &HF011, &H418A, &HAC, &H56, &H5C, &HEE, &HD7, &HC5, &HB9, &H4B)
IID_ID3D11RenderTargetView1 = iid
End Function
Public Function IID_ID3D11UnorderedAccessView1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7B3B6153, &HA886, &H4544, &HAB, &H37, &H65, &H37, &HC8, &H50, &H4, &H3)
IID_ID3D11UnorderedAccessView1 = iid
End Function
Public Function IID_ID3D11Query1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H631B4766, &H36DC, &H461D, &H8D, &HB6, &HC4, &H7E, &H13, &HE6, &H9, &H16)
IID_ID3D11Query1 = iid
End Function
Public Function IID_ID3D11DeviceContext3() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB4E3C01D, &HE79E, &H4637, &H91, &HB2, &H51, &HE, &H9F, &H4C, &H9B, &H8F)
IID_ID3D11DeviceContext3 = iid
End Function
Public Function IID_ID3D11Device3() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA05C8C37, &HD2C6, &H4732, &HB3, &HA0, &H9C, &HE0, &HB0, &HDC, &H9A, &HE6)
IID_ID3D11Device3 = iid
End Function
Public Function IID_ID3D11Device4() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8992AB71, &H2E6, &H4B8D, &HBA, &H48, &HB0, &H56, &HDC, &HDA, &H42, &HC4)
IID_ID3D11Device4 = iid
End Function
Public Function IID_ID3D11On12Device() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H85611E73, &H70A9, &H490E, &H96, &H14, &HA9, &HE3, &H2, &H77, &H79, &H4)
IID_ID3D11On12Device = iid
End Function
Public Function DXGI_DEBUG_D3D11() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4B99317B, &HAC39, &H4AA6, &HBB, &HB, &HBA, &HA0, &H47, &H84, &H79, &H8F)
DXGI_DEBUG_D3D11 = iid
End Function
Public Function IID_ID3D11Debug() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H79CF2233, &H7536, &H4948, &H9D, &H36, &H1E, &H46, &H92, &HDC, &H57, &H60)
IID_ID3D11Debug = iid
End Function
Public Function IID_ID3D11SwitchToRef() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1EF337E3, &H58E7, &H4F83, &HA6, &H92, &HDB, &H22, &H1F, &H5E, &HD4, &H7E)
IID_ID3D11SwitchToRef = iid
End Function
Public Function IID_ID3D11TracingDevice() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1911C771, &H1587, &H413E, &HA7, &HE0, &HFB, &H26, &HC3, &HDE, &H2, &H68)
IID_ID3D11TracingDevice = iid
End Function
Public Function IID_ID3D11RefTrackingOptions() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H193DACDF, &HDB2, &H4C05, &HA5, &H5C, &HEF, &H6, &HCA, &HC5, &H6F, &HD9)
IID_ID3D11RefTrackingOptions = iid
End Function
Public Function IID_ID3D11RefDefaultTrackingOptions() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3916615, &HC644, &H418C, &H9B, &HF4, &H75, &HDB, &H5B, &HE6, &H3C, &HA0)
IID_ID3D11RefDefaultTrackingOptions = iid
End Function
Public Function IID_ID3D11InfoQueue() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6543DBB6, &H1B48, &H42F5, &HAB, &H82, &HE9, &H7E, &HC7, &H43, &H26, &HF6)
IID_ID3D11InfoQueue = iid
End Function
Public Function IID_ID3D11ShaderTrace() As UUID
'{36b013e6-2811-4845-baa7-d623fe0df104}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H36B013E6, CInt(&H2811), CInt(&H4845), &HBA, &HA7, &HD6, &H23, &HFE, &HD, &HF1, &H4)
iid = iid
IID_ID3D11ShaderTrace = iid
End Function
Public Function IID_ID3D11ShaderReflectionType() As UUID
'{6E6FFA6A-9BAE-4613-A51E-91652D508C21}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6E6FFA6A, CInt(&H9BAE), CInt(&H4613), &HA5, &H1E, &H91, &H65, &H2D, &H50, &H8C, &H21)
iid = iid
IID_ID3D11ShaderReflectionType = iid
End Function
Public Function IID_ID3D11ShaderReflectionVariable() As UUID
'{51F23923-F3E5-4BD1-91CB-606177D8DB4C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H51F23923, CInt(&HF3E5), CInt(&H4BD1), &H91, &HCB, &H60, &H61, &H77, &HD8, &HDB, &H4C)
iid = iid
IID_ID3D11ShaderReflectionVariable = iid
End Function
Public Function IID_ID3D11ShaderReflectionConstantBuffer() As UUID
'{EB62D63D-93DD-4318-8AE8-C6F83AD371B8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEB62D63D, CInt(&H93DD), CInt(&H4318), &H8A, &HE8, &HC6, &HF8, &H3A, &HD3, &H71, &HB8)
iid = iid
IID_ID3D11ShaderReflectionConstantBuffer = iid
End Function
Public Function IID_ID3D11ShaderReflection() As UUID
'{8d536ca1-0cca-4956-a837-786963755584}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8D536CA1, CInt(&HCCA), CInt(&H4956), &HA8, &H37, &H78, &H69, &H63, &H75, &H55, &H84)
iid = iid
IID_ID3D11ShaderReflection = iid
End Function
Public Function IID_ID3D11LibraryReflection() As UUID
'{54384F1B-5B3E-4BB7-AE01-60BA3097CBB6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H54384F1B, CInt(&H5B3E), CInt(&H4BB7), &HAE, &H1, &H60, &HBA, &H30, &H97, &HCB, &HB6)
iid = iid
IID_ID3D11LibraryReflection = iid
End Function
Public Function IID_ID3D11FunctionReflection() As UUID
'{207BCECB-D683-4A06-A8A3-9B149B9F73A4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H207BCECB, CInt(&HD683), CInt(&H4A06), &HA8, &HA3, &H9B, &H14, &H9B, &H9F, &H73, &HA4)
iid = iid
IID_ID3D11FunctionReflection = iid
End Function
Public Function IID_ID3D11FunctionParameterReflection() As UUID
'{42757488-334F-47FE-982E-1A65D08CC462}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H42757488, CInt(&H334F), CInt(&H47FE), &H98, &H2E, &H1A, &H65, &HD0, &H8C, &HC4, &H62)
iid = iid
IID_ID3D11FunctionParameterReflection = iid
End Function
Public Function IID_ID3D11Module() As UUID
'{CAC701EE-80FC-4122-8242-10B39C8CEC34}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCAC701EE, CInt(&H80FC), CInt(&H4122), &H82, &H42, &H10, &HB3, &H9C, &H8C, &HEC, &H34)
iid = iid
IID_ID3D11Module = iid
End Function
Public Function IID_ID3D11ModuleInstance() As UUID
'{469E07F7-045A-48D5-AA12-68A478CDF75D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H469E07F7, CInt(&H45A), CInt(&H48D5), &HAA, &H12, &H68, &HA4, &H78, &HCD, &HF7, &H5D)
iid = iid
IID_ID3D11ModuleInstance = iid
End Function
Public Function IID_ID3D11Linker() As UUID
'{59A6CD0E-E10D-4C1F-88C0-63ABA1DAF30E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H59A6CD0E, CInt(&HE10D), CInt(&H4C1F), &H88, &HC0, &H63, &HAB, &HA1, &HDA, &HF3, &HE)
iid = iid
IID_ID3D11Linker = iid
End Function
Public Function IID_ID3D11LinkingNode() As UUID
'{D80DD70C-8D2F-4751-94A1-03C79B3556DB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD80DD70C, CInt(&H8D2F), CInt(&H4751), &H94, &HA1, &H3, &HC7, &H9B, &H35, &H56, &HDB)
iid = iid
IID_ID3D11LinkingNode = iid
End Function
Public Function IID_ID3D11FunctionLinkingGraph() As UUID
'{54133220-1CE8-43D3-8236-9855C5CEECFF}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H54133220, CInt(&H1CE8), CInt(&H43D3), &H82, &H36, &H98, &H55, &HC5, &HCE, &HEC, &HFF)
iid = iid
IID_ID3D11FunctionLinkingGraph = iid
End Function



Public Function IID_ID3D12Object() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC4FEC28F, &H7966, &H4E95, &H9F, &H94, &HF4, &H31, &HCB, &H56, &HC3, &HB8)
IID_ID3D12Object = iid
End Function
Public Function IID_ID3D12DeviceChild() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H905DB94B, &HA00C, &H4140, &H9D, &HF5, &H2B, &H64, &HCA, &H9E, &HA3, &H57)
IID_ID3D12DeviceChild = iid
End Function
Public Function IID_ID3D12RootSignature() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC54A6B66, &H72DF, &H4EE8, &H8B, &HE5, &HA9, &H46, &HA1, &H42, &H92, &H14)
IID_ID3D12RootSignature = iid
End Function
Public Function IID_ID3D12RootSignatureDeserializer() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H34AB647B, &H3CC8, &H46AC, &H84, &H1B, &HC0, &H96, &H56, &H45, &HC0, &H46)
IID_ID3D12RootSignatureDeserializer = iid
End Function
Public Function IID_ID3D12VersionedRootSignatureDeserializer() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7F91CE67, &H90C, &H4BB7, &HB7, &H8E, &HED, &H8F, &HF2, &HE3, &H1D, &HA0)
IID_ID3D12VersionedRootSignatureDeserializer = iid
End Function
Public Function IID_ID3D12Pageable() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H63EE58FB, &H1268, &H4835, &H86, &HDA, &HF0, &H8, &HCE, &H62, &HF0, &HD6)
IID_ID3D12Pageable = iid
End Function
Public Function IID_ID3D12Heap() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6B3B2502, &H6E51, &H45B3, &H90, &HEE, &H98, &H84, &H26, &H5E, &H8D, &HF3)
IID_ID3D12Heap = iid
End Function
Public Function IID_ID3D12Resource() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H696442BE, &HA72E, &H4059, &HBC, &H79, &H5B, &H5C, &H98, &H4, &HF, &HAD)
IID_ID3D12Resource = iid
End Function
Public Function IID_ID3D12CommandAllocator() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6102DEE4, &HAF59, &H4B09, &HB9, &H99, &HB4, &H4D, &H73, &HF0, &H9B, &H24)
IID_ID3D12CommandAllocator = iid
End Function
Public Function IID_ID3D12Fence() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA753DCF, &HC4D8, &H4B91, &HAD, &HF6, &HBE, &H5A, &H60, &HD9, &H5A, &H76)
IID_ID3D12Fence = iid
End Function
Public Function IID_ID3D12PipelineState() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H765A30F3, &HF624, &H4C6F, &HA8, &H28, &HAC, &HE9, &H48, &H62, &H24, &H45)
IID_ID3D12PipelineState = iid
End Function
Public Function IID_ID3D12DescriptorHeap() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8EFB471D, &H616C, &H4F49, &H90, &HF7, &H12, &H7B, &HB7, &H63, &HFA, &H51)
IID_ID3D12DescriptorHeap = iid
End Function
Public Function IID_ID3D12QueryHeap() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD9658AE, &HED45, &H469E, &HA6, &H1D, &H97, &HE, &HC5, &H83, &HCA, &HB4)
IID_ID3D12QueryHeap = iid
End Function
Public Function IID_ID3D12CommandSignature() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC36A797C, &HEC80, &H4F0A, &H89, &H85, &HA7, &HB2, &H47, &H50, &H82, &HD1)
IID_ID3D12CommandSignature = iid
End Function
Public Function IID_ID3D12CommandList() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7116D91C, &HE7E4, &H47CE, &HB8, &HC6, &HEC, &H81, &H68, &HF4, &H37, &HE5)
IID_ID3D12CommandList = iid
End Function
Public Function IID_ID3D12GraphicsCommandList() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5B160D0F, &HAC1B, &H4185, &H8B, &HA8, &HB3, &HAE, &H42, &HA5, &HA4, &H55)
IID_ID3D12GraphicsCommandList = iid
End Function
Public Function IID_ID3D12CommandQueue() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEC870A6, &H5D7E, &H4C22, &H8C, &HFC, &H5B, &HAA, &HE0, &H76, &H16, &HED)
IID_ID3D12CommandQueue = iid
End Function
Public Function IID_ID3D12Device() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H189819F1, &H1DB6, &H4B57, &HBE, &H54, &H18, &H21, &H33, &H9B, &H85, &HF7)
IID_ID3D12Device = iid
End Function
Public Function IID_ID3D12PipelineLibrary() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC64226A8, &H9201, &H46AF, &HB4, &HCC, &H53, &HFB, &H9F, &HF7, &H41, &H4F)
IID_ID3D12PipelineLibrary = iid
End Function
Public Function IID_ID3D12Device1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H77ACCE80, &H638E, &H4E65, &H88, &H95, &HC1, &HF2, &H33, &H86, &H86, &H3E)
IID_ID3D12Device1 = iid
End Function
Public Function IID_ID3D12Debug() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H344488B7, &H6846, &H474B, &HB9, &H89, &HF0, &H27, &H44, &H82, &H45, &HE0)
IID_ID3D12Debug = iid
End Function
Public Function IID_ID3D12Debug1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAFFAA4CA, &H63FE, &H4D8E, &HB8, &HAD, &H15, &H90, &H0, &HAF, &H43, &H4)
IID_ID3D12Debug1 = iid
End Function
Public Function IID_ID3D12DebugDevice1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA9B71770, &HD099, &H4A65, &HA6, &H98, &H3D, &HEE, &H10, &H2, &HF, &H88)
IID_ID3D12DebugDevice1 = iid
End Function
Public Function IID_ID3D12DebugDevice() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3FEBD6DD, &H4973, &H4787, &H81, &H94, &HE4, &H5F, &H9E, &H28, &H92, &H3E)
IID_ID3D12DebugDevice = iid
End Function
Public Function IID_ID3D12DebugCommandQueue() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9E0BF36, &H54AC, &H484F, &H88, &H47, &H4B, &HAE, &HEA, &HB6, &H5, &H3A)
IID_ID3D12DebugCommandQueue = iid
End Function
Public Function IID_ID3D12DebugCommandList1() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H102CA951, &H311B, &H4B01, &HB1, &H1F, &HEC, &HB8, &H3E, &H6, &H1B, &H37)
IID_ID3D12DebugCommandList1 = iid
End Function
Public Function IID_ID3D12DebugCommandList() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9E0BF36, &H54AC, &H484F, &H88, &H47, &H4B, &HAE, &HEA, &HB6, &H5, &H3F)
IID_ID3D12DebugCommandList = iid
End Function
Public Function IID_ID3D12InfoQueue() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H742A90B, &HC387, &H483F, &HB9, &H46, &H30, &HA7, &HE4, &HE6, &H14, &H58)
IID_ID3D12InfoQueue = iid
End Function
Public Function DXGI_DEBUG_D3D12() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCF59A98C, &HA950, &H4326, &H91, &HEF, &H9B, &HBA, &HA1, &H7B, &HFD, &H95)
DXGI_DEBUG_D3D12 = iid
End Function


Public Function IID_ID2D1VertexBuffer() As UUID
'{9b8b1336-00a5-4668-92b7-ced5d8bf9b7b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9B8B1336, CInt(&HA5), CInt(&H4668), &H92, &HB7, &HCE, &HD5, &HD8, &HBF, &H9B, &H7B)
IID_ID2D1VertexBuffer = iid
End Function
Public Function IID_ID2D1ResourceTexture() As UUID
'{688d15c3-02b0-438d-b13a-d1b44c32c39a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H688D15C3, CInt(&H2B0), CInt(&H438D), &HB1, &H3A, &HD1, &HB4, &H4C, &H32, &HC3, &H9A)
IID_ID2D1ResourceTexture = iid
End Function
Public Function IID_ID2D1RenderInfo() As UUID
'{519ae1bd-d19a-420d-b849-364f594776b7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H519AE1BD, CInt(&HD19A), CInt(&H420D), &HB8, &H49, &H36, &H4F, &H59, &H47, &H76, &HB7)
IID_ID2D1RenderInfo = iid
End Function
Public Function IID_ID2D1DrawInfo() As UUID
'{693ce632-7f2f-45de-93fe-18d88b37aa21}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H693CE632, CInt(&H7F2F), CInt(&H45DE), &H93, &HFE, &H18, &HD8, &H8B, &H37, &HAA, &H21)
IID_ID2D1DrawInfo = iid
End Function
Public Function IID_ID2D1ComputeInfo() As UUID
'{5598b14b-9fd7-48b7-9bdb-8f0964eb38bc}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5598B14B, CInt(&H9FD7), CInt(&H48B7), &H9B, &HDB, &H8F, &H9, &H64, &HEB, &H38, &HBC)
IID_ID2D1ComputeInfo = iid
End Function
Public Function IID_ID2D1TransformNode() As UUID
'{b2efe1e7-729f-4102-949f-505fa21bf666}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB2EFE1E7, CInt(&H729F), CInt(&H4102), &H94, &H9F, &H50, &H5F, &HA2, &H1B, &HF6, &H66)
IID_ID2D1TransformNode = iid
End Function
Public Function IID_ID2D1TransformGraph() As UUID
'{13d29038-c3e6-4034-9081-13b53a417992}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H13D29038, CInt(&HC3E6), CInt(&H4034), &H90, &H81, &H13, &HB5, &H3A, &H41, &H79, &H92)
IID_ID2D1TransformGraph = iid
End Function
Public Function IID_ID2D1Transform() As UUID
'{ef1a287d-342a-4f76-8fdb-da0d6ea9f92b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEF1A287D, CInt(&H342A), CInt(&H4F76), &H8F, &HDB, &HDA, &HD, &H6E, &HA9, &HF9, &H2B)
IID_ID2D1Transform = iid
End Function
Public Function IID_ID2D1DrawTransform() As UUID
'{36bfdcb6-9739-435d-a30d-a653beff6a6f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H36BFDCB6, CInt(&H9739), CInt(&H435D), &HA3, &HD, &HA6, &H53, &HBE, &HFF, &H6A, &H6F)
IID_ID2D1DrawTransform = iid
End Function
Public Function IID_ID2D1ComputeTransform() As UUID
'{0d85573c-01e3-4f7d-bfd9-0d60608bf3c3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD85573C, CInt(&H1E3), CInt(&H4F7D), &HBF, &HD9, &HD, &H60, &H60, &H8B, &HF3, &HC3)
IID_ID2D1ComputeTransform = iid
End Function
Public Function IID_ID2D1AnalysisTransform() As UUID
'{0359dc30-95e6-4568-9055-27720d130e93}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H359DC30, CInt(&H95E6), CInt(&H4568), &H90, &H55, &H27, &H72, &HD, &H13, &HE, &H93)
IID_ID2D1AnalysisTransform = iid
End Function
Public Function IID_ID2D1SourceTransform() As UUID
'{db1800dd-0c34-4cf9-be90-31cc0a5653e1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDB1800DD, CInt(&HC34), CInt(&H4CF9), &HBE, &H90, &H31, &HCC, &HA, &H56, &H53, &HE1)
IID_ID2D1SourceTransform = iid
End Function
Public Function IID_ID2D1ConcreteTransform() As UUID
'{1a799d8a-69f7-4e4c-9fed-437ccc6684cc}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1A799D8A, CInt(&H69F7), CInt(&H4E4C), &H9F, &HED, &H43, &H7C, &HCC, &H66, &H84, &HCC)
IID_ID2D1ConcreteTransform = iid
End Function
Public Function IID_ID2D1BlendTransform() As UUID
'{63ac0b32-ba44-450f-8806-7f4ca1ff2f1b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H63AC0B32, CInt(&HBA44), CInt(&H450F), &H88, &H6, &H7F, &H4C, &HA1, &HFF, &H2F, &H1B)
IID_ID2D1BlendTransform = iid
End Function
Public Function IID_ID2D1BorderTransform() As UUID
'{4998735c-3a19-473c-9781-656847e3a347}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4998735C, CInt(&H3A19), CInt(&H473C), &H97, &H81, &H65, &H68, &H47, &HE3, &HA3, &H47)
IID_ID2D1BorderTransform = iid
End Function
Public Function IID_ID2D1OffsetTransform() As UUID
'{3fe6adea-7643-4f53-bd14-a0ce63f24042}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3FE6ADEA, CInt(&H7643), CInt(&H4F53), &HBD, &H14, &HA0, &HCE, &H63, &HF2, &H40, &H42)
IID_ID2D1OffsetTransform = iid
End Function
Public Function IID_ID2D1BoundsAdjustmentTransform() As UUID
'{90f732e2-5092-4606-a819-8651970baccd}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H90F732E2, CInt(&H5092), CInt(&H4606), &HA8, &H19, &H86, &H51, &H97, &HB, &HAC, &HCD)
IID_ID2D1BoundsAdjustmentTransform = iid
End Function
Public Function IID_ID2D1EffectImpl() As UUID
'{a248fd3f-3e6c-4e63-9f03-7f68ecc91db9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA248FD3F, CInt(&H3E6C), CInt(&H4E63), &H9F, &H3, &H7F, &H68, &HEC, &HC9, &H1D, &HB9)
IID_ID2D1EffectImpl = iid
End Function
Public Function IID_ID2D1EffectContext() As UUID
'{3d9f916b-27dc-4ad7-b4f1-64945340f563}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3D9F916B, CInt(&H27DC), CInt(&H4AD7), &HB4, &HF1, &H64, &H94, &H53, &H40, &HF5, &H63)
IID_ID2D1EffectContext = iid
End Function
Public Function IID_ID2D1Device() As UUID
'{47dd575d-ac05-4cdd-8049-9b02cd16f44c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H47DD575D, CInt(&HAC05), CInt(&H4CDD), &H80, &H49, &H9B, &H2, &HCD, &H16, &HF4, &H4C)
IID_ID2D1Device = iid
End Function
Public Function IID_ID2D1Factory1() As UUID
'{bb12d362-daee-4b9a-aa1d-14ba401cfa1f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBB12D362, CInt(&HDAEE), CInt(&H4B9A), &HAA, &H1D, &H14, &HBA, &H40, &H1C, &HFA, &H1F)
IID_ID2D1Factory1 = iid
End Function
Public Function IID_ID2D1Multithread() As UUID
'{31e6e7bc-e0ff-4d46-8c64-a0a8c41c15d3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H31E6E7BC, CInt(&HE0FF), CInt(&H4D46), &H8C, &H64, &HA0, &HA8, &HC4, &H1C, &H15, &HD3)
IID_ID2D1Multithread = iid
End Function
Public Function IID_ID2D1GeometryRealization() As UUID
'{a16907d7-bc02-4801-99e8-8cf7f485f774}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA16907D7, CInt(&HBC02), CInt(&H4801), &H99, &HE8, &H8C, &HF7, &HF4, &H85, &HF7, &H74)
IID_ID2D1GeometryRealization = iid
End Function
Public Function IID_ID2D1DeviceContext1() As UUID
'{d37f57e4-6908-459f-a199-e72f24f79987}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD37F57E4, CInt(&H6908), CInt(&H459F), &HA1, &H99, &HE7, &H2F, &H24, &HF7, &H99, &H87)
IID_ID2D1DeviceContext1 = iid
End Function
Public Function IID_ID2D1Device1() As UUID
'{d21768e1-23a4-4823-a14b-7c3eba85d658}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD21768E1, CInt(&H23A4), CInt(&H4823), &HA1, &H4B, &H7C, &H3E, &HBA, &H85, &HD6, &H58)
IID_ID2D1Device1 = iid
End Function
Public Function IID_ID2D1Factory2() As UUID
'{94f81a73-9212-4376-9c58-b16a3a0d3992}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H94F81A73, CInt(&H9212), CInt(&H4376), &H9C, &H58, &HB1, &H6A, &H3A, &HD, &H39, &H92)
IID_ID2D1Factory2 = iid
End Function
Public Function IID_ID2D1CommandSink1() As UUID
'{9eb767fd-4269-4467-b8c2-eb30cb305743}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9EB767FD, CInt(&H4269), CInt(&H4467), &HB8, &HC2, &HEB, &H30, &HCB, &H30, &H57, &H43)
IID_ID2D1CommandSink1 = iid
End Function
Public Function IID_ID2D1InkStyle() As UUID
'{bae8b344-23fc-4071-8cb5-d05d6f073848}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBAE8B344, CInt(&H23FC), CInt(&H4071), &H8C, &HB5, &HD0, &H5D, &H6F, &H7, &H38, &H48)
IID_ID2D1InkStyle = iid
End Function
Public Function IID_ID2D1Ink() As UUID
'{b499923b-7029-478f-a8b3-432c7c5f5312}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB499923B, CInt(&H7029), CInt(&H478F), &HA8, &HB3, &H43, &H2C, &H7C, &H5F, &H53, &H12)
IID_ID2D1Ink = iid
End Function
Public Function IID_ID2D1GradientMesh() As UUID
'{f292e401-c050-4cde-83d7-04962d3b23c2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF292E401, CInt(&HC050), CInt(&H4CDE), &H83, &HD7, &H4, &H96, &H2D, &H3B, &H23, &HC2)
IID_ID2D1GradientMesh = iid
End Function
Public Function IID_ID2D1ImageSource() As UUID
'{c9b664e5-74a1-4378-9ac2-eefc37a3f4d8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC9B664E5, CInt(&H74A1), CInt(&H4378), &H9A, &HC2, &HEE, &HFC, &H37, &HA3, &HF4, &HD8)
IID_ID2D1ImageSource = iid
End Function
Public Function IID_ID2D1ImageSourceFromWic() As UUID
'{77395441-1c8f-4555-8683-f50dab0fe792}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H77395441, CInt(&H1C8F), CInt(&H4555), &H86, &H83, &HF5, &HD, &HAB, &HF, &HE7, &H92)
IID_ID2D1ImageSourceFromWic = iid
End Function
Public Function IID_ID2D1TransformedImageSource() As UUID
'{7f1f79e5-2796-416c-8f55-700f911445e5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7F1F79E5, CInt(&H2796), CInt(&H416C), &H8F, &H55, &H70, &HF, &H91, &H14, &H45, &HE5)
IID_ID2D1TransformedImageSource = iid
End Function
Public Function IID_ID2D1LookupTable3D() As UUID
'{53dd9855-a3b0-4d5b-82e1-26e25c5e5797}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H53DD9855, CInt(&HA3B0), CInt(&H4D5B), &H82, &HE1, &H26, &HE2, &H5C, &H5E, &H57, &H97)
IID_ID2D1LookupTable3D = iid
End Function
Public Function IID_ID2D1DeviceContext2() As UUID
'{394ea6a3-0c34-4321-950b-6ca20f0be6c7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H394EA6A3, CInt(&HC34), CInt(&H4321), &H95, &HB, &H6C, &HA2, &HF, &HB, &HE6, &HC7)
IID_ID2D1DeviceContext2 = iid
End Function
Public Function IID_ID2D1Device2() As UUID
'{a44472e1-8dfb-4e60-8492-6e2861c9ca8b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA44472E1, CInt(&H8DFB), CInt(&H4E60), &H84, &H92, &H6E, &H28, &H61, &HC9, &HCA, &H8B)
IID_ID2D1Device2 = iid
End Function
Public Function IID_ID2D1Factory3() As UUID
'{0869759f-4f00-413f-b03e-2bda45404d0f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H869759F, CInt(&H4F00), CInt(&H413F), &HB0, &H3E, &H2B, &HDA, &H45, &H40, &H4D, &HF)
IID_ID2D1Factory3 = iid
End Function
Public Function IID_ID2D1CommandSink2() As UUID
'{3bab440e-417e-47df-a2e2-bc0be6a00916}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3BAB440E, CInt(&H417E), CInt(&H47DF), &HA2, &HE2, &HBC, &HB, &HE6, &HA0, &H9, &H16)
IID_ID2D1CommandSink2 = iid
End Function
Public Function IID_ID2D1GdiMetafile1() As UUID
'{2e69f9e8-dd3f-4bf9-95ba-c04f49d788df}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2E69F9E8, CInt(&HDD3F), CInt(&H4BF9), &H95, &HBA, &HC0, &H4F, &H49, &HD7, &H88, &HDF)
IID_ID2D1GdiMetafile1 = iid
End Function
Public Function IID_ID2D1GdiMetafileSink1() As UUID
'{fd0ecb6b-91e6-411e-8655-395e760f91b4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFD0ECB6B, CInt(&H91E6), CInt(&H411E), &H86, &H55, &H39, &H5E, &H76, &HF, &H91, &HB4)
IID_ID2D1GdiMetafileSink1 = iid
End Function
Public Function IID_ID2D1SpriteBatch() As UUID
'{4dc583bf-3a10-438a-8722-e9765224f1f1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4DC583BF, CInt(&H3A10), CInt(&H438A), &H87, &H22, &HE9, &H76, &H52, &H24, &HF1, &HF1)
IID_ID2D1SpriteBatch = iid
End Function
Public Function IID_ID2D1DeviceContext3() As UUID
'{235a7496-8351-414c-bcd4-6672ab2d8e00}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H235A7496, CInt(&H8351), CInt(&H414C), &HBC, &HD4, &H66, &H72, &HAB, &H2D, &H8E, &H0)
IID_ID2D1DeviceContext3 = iid
End Function
Public Function IID_ID2D1Device3() As UUID
'{852f2087-802c-4037-ab60-ff2e7ee6fc01}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H852F2087, CInt(&H802C), CInt(&H4037), &HAB, &H60, &HFF, &H2E, &H7E, &HE6, &HFC, &H1)
IID_ID2D1Device3 = iid
End Function
Public Function IID_ID2D1Factory4() As UUID
'{bd4ec2d2-0662-4bee-ba8e-6f29f032e096}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBD4EC2D2, CInt(&H662), CInt(&H4BEE), &HBA, &H8E, &H6F, &H29, &HF0, &H32, &HE0, &H96)
IID_ID2D1Factory4 = iid
End Function
Public Function IID_ID2D1CommandSink3() As UUID
'{18079135-4cf3-4868-bc8e-06067e6d242d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H18079135, CInt(&H4CF3), CInt(&H4868), &HBC, &H8E, &H6, &H6, &H7E, &H6D, &H24, &H2D)
IID_ID2D1CommandSink3 = iid
End Function
Public Function IID_ID2D1SvgGlyphStyle() As UUID
'{af671749-d241-4db8-8e41-dcc2e5c1a438}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAF671749, CInt(&HD241), CInt(&H4DB8), &H8E, &H41, &HDC, &HC2, &HE5, &HC1, &HA4, &H38)
IID_ID2D1SvgGlyphStyle = iid
End Function
Public Function IID_ID2D1DeviceContext4() As UUID
'{8c427831-3d90-4476-b647-c4fae349e4db}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8C427831, CInt(&H3D90), CInt(&H4476), &HB6, &H47, &HC4, &HFA, &HE3, &H49, &HE4, &HDB)
IID_ID2D1DeviceContext4 = iid
End Function
Public Function IID_ID2D1Device4() As UUID
'{d7bdb159-5683-4a46-bc9c-72dc720b858b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD7BDB159, CInt(&H5683), CInt(&H4A46), &HBC, &H9C, &H72, &HDC, &H72, &HB, &H85, &H8B)
IID_ID2D1Device4 = iid
End Function
Public Function IID_ID2D1Factory5() As UUID
'{c4349994-838e-4b0f-8cab-44997d9eeacc}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC4349994, CInt(&H838E), CInt(&H4B0F), &H8C, &HAB, &H44, &H99, &H7D, &H9E, &HEA, &HCC)
IID_ID2D1Factory5 = iid
End Function
Public Function IID_ID2D1CommandSink() As UUID
'{54d7898a-a061-40a7-bec7-e465bcba2c4f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H54D7898A, CInt(&HA061), CInt(&H40A7), &HBE, &HC7, &HE4, &H65, &HBC, &HBA, &H2C, &H4F)
IID_ID2D1CommandSink = iid
End Function
Public Function IID_ID2D1CommandList() As UUID
'{b4f34a19-2383-4d76-94f6-ec343657c3dc}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB4F34A19, CInt(&H2383), CInt(&H4D76), &H94, &HF6, &HEC, &H34, &H36, &H57, &HC3, &HDC)
IID_ID2D1CommandList = iid
End Function
Public Function IID_ID2D1PrintControl() As UUID
'{2c1d867d-c290-41c8-ae7e-34a98702e9a5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2C1D867D, CInt(&HC290), CInt(&H41C8), &HAE, &H7E, &H34, &HA9, &H87, &H2, &HE9, &HA5)
IID_ID2D1PrintControl = iid
End Function
Public Function IID_ID2D1ImageBrush() As UUID
'{fe9e984d-3f95-407c-b5db-cb94d4e8f87c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFE9E984D, CInt(&H3F95), CInt(&H407C), &HB5, &HDB, &HCB, &H94, &HD4, &HE8, &HF8, &H7C)
IID_ID2D1ImageBrush = iid
End Function
Public Function IID_ID2D1BitmapBrush1() As UUID
'{41343a53-e41a-49a2-91cd-21793bbb62e5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H41343A53, CInt(&HE41A), CInt(&H49A2), &H91, &HCD, &H21, &H79, &H3B, &HBB, &H62, &HE5)
IID_ID2D1BitmapBrush1 = iid
End Function
Public Function IID_ID2D1StrokeStyle1() As UUID
'{10a72a66-e91c-43f4-993f-ddf4b82b0b4a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10A72A66, CInt(&HE91C), CInt(&H43F4), &H99, &H3F, &HDD, &HF4, &HB8, &H2B, &HB, &H4A)
IID_ID2D1StrokeStyle1 = iid
End Function
Public Function IID_ID2D1PathGeometry1() As UUID
'{62baa2d2-ab54-41b7-b872-787e0106a421}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H62BAA2D2, CInt(&HAB54), CInt(&H41B7), &HB8, &H72, &H78, &H7E, &H1, &H6, &HA4, &H21)
IID_ID2D1PathGeometry1 = iid
End Function
Public Function IID_ID2D1Properties() As UUID
'{483473d7-cd46-4f9d-9d3a-3112aa80159d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H483473D7, CInt(&HCD46), CInt(&H4F9D), &H9D, &H3A, &H31, &H12, &HAA, &H80, &H15, &H9D)
IID_ID2D1Properties = iid
End Function
Public Function IID_ID2D1Effect() As UUID
'{28211a43-7d89-476f-8181-2d6159b220ad}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H28211A43, CInt(&H7D89), CInt(&H476F), &H81, &H81, &H2D, &H61, &H59, &HB2, &H20, &HAD)
IID_ID2D1Effect = iid
End Function
Public Function IID_ID2D1Bitmap1() As UUID
'{a898a84c-3873-4588-b08b-ebbf978df041}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA898A84C, CInt(&H3873), CInt(&H4588), &HB0, &H8B, &HEB, &HBF, &H97, &H8D, &HF0, &H41)
IID_ID2D1Bitmap1 = iid
End Function
Public Function IID_ID2D1ColorContext() As UUID
'{1c4820bb-5771-4518-a581-2fe4dd0ec657}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1C4820BB, CInt(&H5771), CInt(&H4518), &HA5, &H81, &H2F, &HE4, &HDD, &HE, &HC6, &H57)
IID_ID2D1ColorContext = iid
End Function
Public Function IID_ID2D1GradientStopCollection1() As UUID
'{ae1572f4-5dd0-4777-998b-9279472ae63b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAE1572F4, CInt(&H5DD0), CInt(&H4777), &H99, &H8B, &H92, &H79, &H47, &H2A, &HE6, &H3B)
IID_ID2D1GradientStopCollection1 = iid
End Function
Public Function IID_ID2D1DrawingStateBlock1() As UUID
'{689f1f85-c72e-4e33-8f19-85754efd5ace}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H689F1F85, CInt(&HC72E), CInt(&H4E33), &H8F, &H19, &H85, &H75, &H4E, &HFD, &H5A, &HCE)
IID_ID2D1DrawingStateBlock1 = iid
End Function
Public Function IID_ID2D1DeviceContext() As UUID
'{e8f7fe7a-191c-466d-ad95-975678bda998}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE8F7FE7A, CInt(&H191C), CInt(&H466D), &HAD, &H95, &H97, &H56, &H78, &HBD, &HA9, &H98)
IID_ID2D1DeviceContext = iid
End Function

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

Public Function IID_IDirectSoundFXGargle() As UUID
'{D616F352-D622-11CE-AAC5-0020AF0B99A3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD616F352, CInt(&HD622), CInt(&H11CE), &HAA, &HC5, &H0, &H20, &HAF, &HB, &H99, &HA3)
IID_IDirectSoundFXGargle = iid
End Function
Public Function IID_IDirectSoundFXParamEq() As UUID
'{C03CA9FE-FE90-4204-8078-82334CD177DA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC03CA9FE, CInt(&HFE90), CInt(&H4204), &H80, &H78, &H82, &H33, &H4C, &HD1, &H77, &HDA)
IID_IDirectSoundFXParamEq = iid
End Function
Public Function IID_IDirectSoundFXI3DL2Reverb() As UUID
'{4B166A6A-0D66-43F3-80E3-EE6280DEE1A4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4B166A6A, CInt(&HD66), CInt(&H43F3), &H80, &HE3, &HEE, &H62, &H80, &HDE, &HE1, &HA4)
IID_IDirectSoundFXI3DL2Reverb = iid
End Function
Public Function IID_IDirectSoundFXWavesReverb() As UUID
'{46858C3A-0DC6-45E3-B760-D4EEF16CB325}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H46858C3A, CInt(&HDC6), CInt(&H45E3), &HB7, &H60, &HD4, &HEE, &HF1, &H6C, &HB3, &H25)
IID_IDirectSoundFXWavesReverb = iid
End Function
Public Function IID_IDirectSoundFXCompressor() As UUID
'{4BBD1154-62F6-4E2C-A15C-D3B6C417F7A0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4BBD1154, CInt(&H62F6), CInt(&H4E2C), &HA1, &H5C, &HD3, &HB6, &HC4, &H17, &HF7, &HA0)
IID_IDirectSoundFXCompressor = iid
End Function
Public Function IID_IDirectSoundFXDistortion() As UUID
'{8ECF4326-455F-4D8B-BDA9-8D5D3E9E3E0B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8ECF4326, CInt(&H455F), CInt(&H4D8B), &HBD, &HA9, &H8D, &H5D, &H3E, &H9E, &H3E, &HB)
IID_IDirectSoundFXDistortion = iid
End Function
Public Function IID_IDirectSoundFXEcho() As UUID
'{8BD28EDF-50DB-4E92-A2BD-445488D1ED42}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8BD28EDF, CInt(&H50DB), CInt(&H4E92), &HA2, &HBD, &H44, &H54, &H88, &HD1, &HED, &H42)
IID_IDirectSoundFXEcho = iid
End Function
Public Function IID_IDirectSoundFXFlanger() As UUID
'{903E9878-2C92-4072-9B2C-EA68F5396783}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H903E9878, CInt(&H2C92), CInt(&H4072), &H9B, &H2C, &HEA, &H68, &HF5, &H39, &H67, &H83)
IID_IDirectSoundFXFlanger = iid
End Function
Public Function IID_IDirectSoundFXChorus() As UUID
'{880842E3-145F-43E6-A934-A71806E50547}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H880842E3, CInt(&H145F), CInt(&H43E6), &HA9, &H34, &HA7, &H18, &H6, &HE5, &H5, &H47)
IID_IDirectSoundFXChorus = iid
End Function
Public Function IID_IDirectSound() As UUID
'{279AFA83-4981-11CE-A521-0020AF0BE560}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H279AFA83, CInt(&H4981), CInt(&H11CE), &HA5, &H21, &H0, &H20, &HAF, &HB, &HE5, &H60)
IID_IDirectSound = iid
End Function
Public Function IID_IDirectSoundBuffer() As UUID
'{279AFA85-4981-11CE-A521-0020AF0BE560}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H279AFA85, CInt(&H4981), CInt(&H11CE), &HA5, &H21, &H0, &H20, &HAF, &HB, &HE5, &H60)
IID_IDirectSoundBuffer = iid
End Function
Public Function IID_IDirectSound8() As UUID
'{C50A7E93-F395-4834-9EF6-7FA99DE50966}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC50A7E93, CInt(&HF395), CInt(&H4834), &H9E, &HF6, &H7F, &HA9, &H9D, &HE5, &H9, &H66)
IID_IDirectSound8 = iid
End Function
Public Function IID_IDirectSoundBuffer8() As UUID
'{6825A449-7524-4D82-920F-50E36AB3AB1E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6825A449, CInt(&H7524), CInt(&H4D82), &H92, &HF, &H50, &HE3, &H6A, &HB3, &HAB, &H1E)
IID_IDirectSoundBuffer8 = iid
End Function
Public Function IID_IDirectSound3DBuffer() As UUID
'{279AFA86-4981-11CE-A521-0020AF0BE560}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H279AFA86, CInt(&H4981), CInt(&H11CE), &HA5, &H21, &H0, &H20, &HAF, &HB, &HE5, &H60)
IID_IDirectSound3DBuffer = iid
End Function
Public Function IID_IDirectSound3DListener() As UUID
'{279AFA84-4981-11CE-A521-0020AF0BE560}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H279AFA84, CInt(&H4981), CInt(&H11CE), &HA5, &H21, &H0, &H20, &HAF, &HB, &HE5, &H60)
IID_IDirectSound3DListener = iid
End Function
Public Function IID_IDirectSoundCapture() As UUID
'{B0210781-89CD-11D0-AF08-00A0C925CD16}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB0210781, CInt(&H89CD), CInt(&H11D0), &HAF, &H8, &H0, &HA0, &HC9, &H25, &HCD, &H16)
IID_IDirectSoundCapture = iid
End Function
Public Function IID_IDirectSoundCaptureBuffer() As UUID
'{B0210782-89CD-11D0-AF08-00A0C925CD16}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB0210782, CInt(&H89CD), CInt(&H11D0), &HAF, &H8, &H0, &HA0, &HC9, &H25, &HCD, &H16)
IID_IDirectSoundCaptureBuffer = iid
End Function
Public Function IID_IDirectSoundCaptureBuffer8() As UUID
'{00990DF4-0DBB-4872-833E-6D303E80AEB6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H990DF4, CInt(&HDBB), CInt(&H4872), &H83, &H3E, &H6D, &H30, &H3E, &H80, &HAE, &HB6)
IID_IDirectSoundCaptureBuffer8 = iid
End Function
Public Function IID_IDirectSoundNotify() As UUID
'{B0210783-89CD-11D0-AF08-00A0C925CD16}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB0210783, CInt(&H89CD), CInt(&H11D0), &HAF, &H8, &H0, &HA0, &HC9, &H25, &HCD, &H16)
IID_IDirectSoundNotify = iid
End Function

Public Function IID_IMultiMediaStream() As UUID
'{B502D1BC-9A57-11d0-8FDE-00C04FD9189D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB502D1BC, CInt(&H9A57), CInt(&H11D0), &H8F, &HDE, &H0, &HC0, &H4F, &HD9, &H18, &H9D)
IID_IMultiMediaStream = iid
End Function
Public Function IID_IMediaStream() As UUID
'{B502D1BD-9A57-11d0-8FDE-00C04FD9189D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB502D1BD, CInt(&H9A57), CInt(&H11D0), &H8F, &HDE, &H0, &HC0, &H4F, &HD9, &H18, &H9D)
IID_IMediaStream = iid
End Function
Public Function IID_IStreamSample() As UUID
'{B502D1BE-9A57-11d0-8FDE-00C04FD9189D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB502D1BE, CInt(&H9A57), CInt(&H11D0), &H8F, &HDE, &H0, &HC0, &H4F, &HD9, &H18, &H9D)
IID_IStreamSample = iid
End Function
Public Function IID_IAMMultiMediaStream() As UUID
'{BEBE595C-9A6F-11d0-8FDE-00C04FD9189D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBEBE595C, CInt(&H9A6F), CInt(&H11D0), &H8F, &HDE, &H0, &HC0, &H4F, &HD9, &H18, &H9D)
IID_IAMMultiMediaStream = iid
End Function
Public Function IID_IAMMediaStream() As UUID
'{BEBE595D-9A6F-11d0-8FDE-00C04FD9189D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBEBE595D, CInt(&H9A6F), CInt(&H11D0), &H8F, &HDE, &H0, &HC0, &H4F, &HD9, &H18, &H9D)
IID_IAMMediaStream = iid
End Function
Public Function IID_IMediaStreamFilter() As UUID
'{BEBE595E-9A6F-11d0-8FDE-00C04FD9189D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBEBE595E, CInt(&H9A6F), CInt(&H11D0), &H8F, &HDE, &H0, &HC0, &H4F, &HD9, &H18, &H9D)
IID_IMediaStreamFilter = iid
End Function
Public Function IID_IDirectDrawMediaSampleAllocator() As UUID
'{AB6B4AFC-F6E4-11d0-900D-00C04FD9189D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB6B4AFC, CInt(&HF6E4), CInt(&H11D0), &H90, &HD, &H0, &HC0, &H4F, &HD9, &H18, &H9D)
IID_IDirectDrawMediaSampleAllocator = iid
End Function
Public Function IID_IDirectDrawMediaSample() As UUID
'{AB6B4AFE-F6E4-11d0-900D-00C04FD9189D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB6B4AFE, CInt(&HF6E4), CInt(&H11D0), &H90, &HD, &H0, &HC0, &H4F, &HD9, &H18, &H9D)
IID_IDirectDrawMediaSample = iid
End Function
Public Function IID_IAMMediaTypeStream() As UUID
'{AB6B4AFA-F6E4-11d0-900D-00C04FD9189D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB6B4AFA, CInt(&HF6E4), CInt(&H11D0), &H90, &HD, &H0, &HC0, &H4F, &HD9, &H18, &H9D)
IID_IAMMediaTypeStream = iid
End Function
Public Function IID_IAMMediaTypeSample() As UUID
'{AB6B4AFB-F6E4-11d0-900D-00C04FD9189D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB6B4AFB, CInt(&HF6E4), CInt(&H11D0), &H90, &HD, &H0, &HC0, &H4F, &HD9, &H18, &H9D)
IID_IAMMediaTypeSample = iid
End Function
Public Function IID_IAudioMediaStream() As UUID
'{f7537560-a3be-11d0-8212-00c04fc32c45}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF7537560, CInt(&HA3BE), CInt(&H11D0), &H82, &H12, &H0, &HC0, &H4F, &HC3, &H2C, &H45)
IID_IAudioMediaStream = iid
End Function
Public Function IID_IAudioStreamSample() As UUID
'{345fee00-aba5-11d0-8212-00c04fc32c45}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H345FEE00, CInt(&HABA5), CInt(&H11D0), &H82, &H12, &H0, &HC0, &H4F, &HC3, &H2C, &H45)
IID_IAudioStreamSample = iid
End Function
Public Function IID_IMemoryData() As UUID
'{327fc560-af60-11d0-8212-00c04fc32c45}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H327FC560, CInt(&HAF60), CInt(&H11D0), &H82, &H12, &H0, &HC0, &H4F, &HC3, &H2C, &H45)
IID_IMemoryData = iid
End Function
Public Function IID_IAudioData() As UUID
'{54c719c0-af60-11d0-8212-00c04fc32c45}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H54C719C0, CInt(&HAF60), CInt(&H11D0), &H82, &H12, &H0, &HC0, &H4F, &HC3, &H2C, &H45)
IID_IAudioData = iid
End Function
Public Function IID_IDirectDrawMediaStream() As UUID
'{F4104FCE-9A70-11d0-8FDE-00C04FD9189D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF4104FCE, CInt(&H9A70), CInt(&H11D0), &H8F, &HDE, &H0, &HC0, &H4F, &HD9, &H18, &H9D)
IID_IDirectDrawMediaStream = iid
End Function
Public Function IID_IDirectDrawStreamSample() As UUID
'{F4104FCF-9A70-11d0-8FDE-00C04FD9189D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF4104FCF, CInt(&H9A70), CInt(&H11D0), &H8F, &HDE, &H0, &HC0, &H4F, &HD9, &H18, &H9D)
IID_IDirectDrawStreamSample = iid
End Function
