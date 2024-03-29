VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLogFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *********************************************************************
'  Copyright ©1998-99 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  Warning: This computer program is protected by copyright law and
'  international treaties. Unauthorized reproduction or distribution
'  of this program, or any portion of it, may result in severe civil
'  and criminal penalties, and will be prosecuted to the maximum
'  extent possible under the law.
' *********************************************************************
Option Explicit

' Logical Font
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64

Private Const CLIP_DEFAULT_PRECIS = 0
Private Const CLIP_CHARACTER_PRECIS = 1
Private Const CLIP_STROKE_PRECIS = 2
Private Const CLIP_MASK = &HF
Private Const CLIP_LH_ANGLES = 16
Private Const CLIP_TT_ALWAYS = 32
Private Const CLIP_EMBEDDED = 128

Private Const DEFAULT_QUALITY = 0
Private Const DRAFT_QUALITY = 1
Private Const PROOF_QUALITY = 2

Private Const DEFAULT_PITCH = 0
Private Const FIXED_PITCH = 1
Private Const VARIABLE_PITCH = 2

Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1
Private Const SYMBOL_CHARSET = 2
Private Const SHIFTJIS_CHARSET = 128
Private Const HANGEUL_CHARSET = 129
Private Const CHINESEBIG5_CHARSET = 136
Private Const OEM_CHARSET = 255

' Font Families
'
Private Const FF_DONTCARE = 0    '  Don't care or don't know.
Private Const FF_ROMAN = 16      '  Variable stroke width, serifed.

' Times Roman, Century Schoolbook, etc.
Private Const FF_SWISS = 32      '  Variable stroke width, sans-serifed.

' Helvetica, Swiss, etc.
Private Const FF_MODERN = 48     '  Constant stroke width, serifed or sans-serifed.

' Pica, Elite, Courier, etc.
Private Const FF_SCRIPT = 64     '  Cursive, etc.
Private Const FF_DECORATIVE = 80 '  Old English, etc.

' Font Weights
Private Const FW_DONTCARE = 0
Private Const FW_THIN = 100
Private Const FW_EXTRALIGHT = 200
Private Const FW_LIGHT = 300
Private Const FW_NORMAL = 400
Private Const FW_MEDIUM = 500
Private Const FW_SEMIBOLD = 600
Private Const FW_BOLD = 700
Private Const FW_EXTRABOLD = 800
Private Const FW_HEAVY = 900

Private Const FW_ULTRALIGHT = FW_EXTRALIGHT
Private Const FW_REGULAR = FW_NORMAL
Private Const FW_DEMIBOLD = FW_SEMIBOLD
Private Const FW_ULTRABOLD = FW_EXTRABOLD
Private Const FW_BLACK = FW_HEAVY

Private Const OUT_DEFAULT_PRECIS = 0
Private Const OUT_STRING_PRECIS = 1
Private Const OUT_CHARACTER_PRECIS = 2
Private Const OUT_STROKE_PRECIS = 3
Private Const OUT_TT_PRECIS = 4
Private Const OUT_DEVICE_PRECIS = 5
Private Const OUT_RASTER_PRECIS = 6
Private Const OUT_TT_ONLY_PRECIS = 7
Private Const OUT_OUTLINE_PRECIS = 8

Private Type LogFont
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName As String * LF_FACESIZE
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LogFont) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private m_Font As StdFont
Private m_hFont As Long
Private m_Rotation As Single

Public Sub CleanUp()
    Class_Terminate
End Sub

Private Sub Class_Terminate()
   '
   ' Clean-up created objects!!!
   '
   If m_hFont Then
      Call DeleteObject(m_hFont)
      Set m_Font = Nothing
   End If
End Sub

Public Property Set LogFont(ByVal NewFont As IFont)
   If m_hFont Then
      Call DeleteObject(m_hFont)
      m_hFont = 0
   End If
   
   Set m_Font = Nothing
   If Not NewFont Is Nothing Then
      '
      ' Stash a copy of the passed object,
      ' to avoid a new reference to it.
      '
      NewFont.Clone m_Font
      m_hFont = CreateLogFont
   End If
End Property

Public Property Get LogFont() As IFont
   Set LogFont = m_Font
End Property

Public Property Let Rotation(ByVal NewVal As Single)
   If NewVal <> m_Rotation Then
      m_Rotation = NewVal
      If m_hFont Then
         Call DeleteObject(m_hFont)
         m_hFont = 0
      End If
      If Not (m_Font Is Nothing) Then
         m_hFont = CreateLogFont
      End If
   End If
End Property

Public Property Get Rotation() As Single
   Rotation = m_Rotation
End Property

Public Property Get Handle() As Long
   Handle = m_hFont
End Property

Private Function CreateLogFont() As Long
   Dim lf As LogFont
   Dim hWnd As Long
   Dim hDC As Long
   
   hWnd = GetDesktopWindow
   hDC = GetDC(hWnd)
   
   With lf
      '
      ' All but two properties are very straight-forward,
      ' even with rotation, and map directly.
      '
      .lfHeight = -(m_Font.Size * GetDeviceCaps(hDC, LOGPIXELSY)) / 72
      .lfWidth = 0
      .lfEscapement = m_Rotation * 10
      .lfOrientation = .lfEscapement
      .lfWeight = m_Font.Weight
      .lfItalic = m_Font.Italic
      .lfUnderline = m_Font.Underline
      .lfStrikeOut = m_Font.Strikethrough
      .lfClipPrecision = CLIP_DEFAULT_PRECIS
      .lfQuality = PROOF_QUALITY
      .lfPitchAndFamily = DEFAULT_PITCH Or FF_DONTCARE
      .lfFaceName = m_Font.Name & vbNullChar
      '
      ' OEM fonts can't rotate, and we must force
      ' substitution with something ANSI.
      '
      .lfCharSet = m_Font.Charset
      If .lfCharSet = OEM_CHARSET Then
         If (m_Rotation Mod 360) <> 0 Then
            .lfCharSet = ANSI_CHARSET
         End If
      End If
      '
      ' Only TrueType fonts can rotate, so we must
      ' specify TT-only if angle is not zero.
      '
      If (m_Rotation Mod 360) <> 0 Then
         .lfOutPrecision = OUT_TT_ONLY_PRECIS
      Else
         .lfOutPrecision = OUT_DEFAULT_PRECIS
      End If
   End With
   
   CreateLogFont = CreateFontIndirect(lf)
   Call ReleaseDC(hWnd, hDC)
End Function
