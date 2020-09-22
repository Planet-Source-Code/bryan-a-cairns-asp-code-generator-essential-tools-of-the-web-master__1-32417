Attribute VB_Name = "mod_Fonts"
'****************************************************************
'Windows API/Global Declarations for :ShowFonts
'****************************************************************

'     'set in optFontType
Public ShowFontType
'     'Font enumeration types
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64

Type LOGFONT
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
       lfFaceName(LF_FACESIZE) As Byte
End Type


Type NEWTEXTMETRIC
       tmHeight As Long
       tmAscent As Long
       tmDescent As Long
       tmInternalLeading As Long
       tmExternalLeading As Long
       tmAveCharWidth As Long
       tmMaxCharWidth As Long
       tmWeight As Long
       tmOverhang As Long
       tmDigitizedAspectX As Long
       tmDigitizedAspectY As Long
       tmFirstChar As Byte
       tmLastChar As Byte
       tmDefaultChar As Byte
       tmBreakChar As Byte
       tmItalic As Byte
       tmUnderlined As Byte
       tmStruckOut As Byte
       tmPitchAndFamily As Byte
       tmCharSet As Byte
       ntmFlags As Long
       ntmSizeEM As Long
       ntmCellHeight As Long
       ntmAveWidth As Long
End Type

'     'ntmFlags field flags
Public Const NTM_REGULAR = &H40&
Public Const NTM_BOLD = &H20&
Public Const NTM_ITALIC = &H1&
'     'tmPitchAndFamily flags
Public Const TMPF_FIXED_PITCH = &H1
Public Const TMPF_VECTOR = &H2
Public Const TMPF_DEVICE = &H8
Public Const TMPF_TRUETYPE = &H4
Public Const ELF_VERSION = 0
Public Const ELF_CULTURE_LATIN = 0
'     'EnumFonts Masks
Public Const VECTOR_FONTTYPE = &H0
Public Const RASTER_FONTTYPE = &H1
Public Const DEVICE_FONTTYPE = &H2
Public Const TRUETYPE_FONTTYPE = &H4

Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" _
        (ByVal hDC As Long, ByVal lpszFamily As String, _
       ByVal lpEnumFontFamProc As Long, lParam As Any) As Long

Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Declare Function ReleaseDC Lib "user32" _
       (ByVal hwnd As Long, ByVal hDC As Long) As Long


Function EnumFontFamDisplayProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, lParam As ListBox) As Long

        
       'This callback function is identical to EnumFontFamProc from
       '
       'the example on Enumerating Fonts except that it also stores
       '
       '     'the FontType as the list item's ItemData.
       Dim FaceName As String
       Dim FaceType As String
       '     'convert the returned string from UniCode to ANSI
       FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
       '     'add the font to the list
       lParam.AddItem Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
        
       '     'add the FontType to the listbox itemdata property
       lParam.ItemData(lParam.NewIndex) = FontType
        
       '     'return success to the call
       EnumFontFamDisplayProc = 1
End Function



Public Sub FontsToList(CMB As Object)
Dim hDC As Long
CMB.Clear
hDC = GetDC(CMB.hwnd)
EnumFontFamilies hDC, vbNullString, AddressOf EnumFontFamDisplayProc, CMB
ReleaseDC CMB.hwnd, hDC
End Sub

Public Sub SelectDefaultFont(CMB As Object)
Dim I As Integer
For I = 0 To CMB.ListCount - 1
    If LCase(CMB.List(I)) = LCase(GetSetting(App.Title, "Main", "FontFace", "Arial")) Then
    CMB.ListIndex = I
    Exit For
    End If
Next I
End Sub
