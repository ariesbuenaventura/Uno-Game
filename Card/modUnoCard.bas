Attribute VB_Name = "modUnoCard"
 Option Explicit

Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64

Private Type LOGFONT
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

Private Type NEWTEXTMETRIC
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

Private Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As Long, _
                                                                                ByVal lpszFamily As String, _
                                                                                ByVal lpEnumFontFamProc As Long, _
                                                                                LParam As Any) As Long
                                                                                
Private Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _
                                ByVal FontType As Long, _
                                LParam As Variant) As Long

    Dim FaceName As String
   
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
   
    If Left$(FaceName, InStr(FaceName, vbNullChar) - 1) = LParam Then
        EnumFontFamProc = 0
    Else
        EnumFontFamProc = 1
    End If
End Function

Public Function IsFontInstalled(ByVal hdc As Long, ByVal FaceName As String) As Boolean
    If Not CBool(EnumFontFamilies(hdc, vbNullString, AddressOf EnumFontFamProc, CVar(FaceName))) Then
        IsFontInstalled = True
    End If
End Function
