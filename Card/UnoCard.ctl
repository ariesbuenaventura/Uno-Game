VERSION 5.00
Begin VB.UserControl UnoCard 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1230
   ScaleHeight     =   117
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   82
   ToolboxBitmap   =   "UnoCard.ctx":0000
   Begin VB.Timer tmrHover 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "UnoCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const WM_KEYDOWN = &H100

Private Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Const CRD_OFFS_SUIT = 101
Private Const CRD_WILD_CARD = 201
Private Const CRD_DRAW_FOUR = 202
Private Const CRD_OFFS_DECK = 301

Public Enum DeckConstants
    uno_DCDefault
    uno_DCCoconut
    uno_DCBamboo
    uno_Fish
    uno_Flower
End Enum

Public Enum FaceConstants
    uno_FCUp
    uno_FCDown
End Enum

Public Enum RankConstants
    uno_RCZero
    uno_RCOne
    uno_RCTwo
    uno_RCThree
    uno_RCFour
    uno_RCFive
    uno_RCSix
    uno_RCSeven
    uno_RCEight
    uno_RCNine
    uno_RCDrawTwo
    uno_RCReverse
    uno_RCSkip
    uno_RCWild
    uno_RCDrawFour
End Enum

Public Enum SuitConstants
    uno_SCBlue
    uno_SCRed
    uno_SCGreen
    uno_SCYellow
End Enum

Private Type CardProperties
    AutoUpdate    As Boolean
    Data          As String
    Deck          As DeckConstants
    Face          As FaceConstants
    Invert        As Boolean
    Rank          As RankConstants
    Points         As Integer
    ShowFocusRect As Boolean
    Suit          As SuitConstants
End Type

Event Click()
Event KeyPress(KeyAscii As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseOut()
Event MouseOver()
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim MyProp     As CardProperties
Dim IsGotFocus As Boolean
Dim IsHover    As Boolean
Dim UnoFont(2) As String

Public Property Get AutoUpdate() As Boolean
    AutoUpdate = MyProp.AutoUpdate
End Property

Public Property Let AutoUpdate(ByVal New_AutoUpdate As Boolean)
    MyProp.AutoUpdate = New_AutoUpdate
    PropertyChanged "AutoUpdate"
End Property

Public Property Get Data() As String
    Data = MyProp.Data
End Property

Public Property Let Data(ByVal New_Data As String)
    MyProp.Data = New_Data
    PropertyChanged "Data"
End Property

Public Property Get Deck() As DeckConstants
    Deck = MyProp.Deck
End Property

Public Property Let Deck(ByVal New_Deck As DeckConstants)
    MyProp.Deck = New_Deck
    PropertyChanged "Deck"
    
    Call Redraw
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal bVal As Boolean)
    UserControl.Enabled() = bVal
    PropertyChanged "Enabled"
End Property

Public Property Get Face() As FaceConstants
    Face = MyProp.Face
End Property

Public Property Let Face(ByVal New_Face As FaceConstants)
    MyProp.Face = New_Face
    PropertyChanged "Face"
    
    Call Redraw
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Picture() As Picture
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

Public Property Get Rank() As RankConstants
    Rank = MyProp.Rank
End Property

Public Property Let Rank(ByVal New_Rank As RankConstants)
    MyProp.Rank = New_Rank
    PropertyChanged "Rank"
    
    Call Redraw
End Property

Public Property Get Points() As Integer
    Points = MyProp.Points
End Property

Public Property Let Points(ByVal New_Points As Integer)
    MyProp.Points = New_Points
    PropertyChanged "Points"
End Property

Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = MyProp.ShowFocusRect
End Property

Public Property Let ShowFocusRect(ByVal bVal As Boolean)
    MyProp.ShowFocusRect = bVal
    PropertyChanged "ShowFocusRect"
End Property

Public Property Get Suit() As SuitConstants
    Suit = MyProp.Suit
End Property

Public Property Let Suit(ByVal New_Suit As SuitConstants)
    MyProp.Suit = New_Suit
    PropertyChanged "Suit"
    
    Call Redraw
End Property

Private Sub tmrHover_Timer()
    Dim tPT    As POINTAPI
    Dim rcRect As RECT
    
    GetCursorPos tPT
    If WindowFromPoint(tPT.x, tPT.y) = UserControl.hWnd Then
        If Not IsHover Then
            IsHover = True
            DrawFocusRect RGB(229, 151, 0)
            
            RaiseEvent MouseOver
        End If
    Else
        If IsHover Then
            IsHover = False
            
            If IsGotFocus Then
                If MyProp.ShowFocusRect Then
                    DrawFocusRect RGB(118, 142, 239)
                Else
                    UserControl.Cls
                End If
            Else
                UserControl.Cls
            End If
            
            RaiseEvent MouseOut
        End If
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_GotFocus()
    If Not IsHover Then
        If MyProp.ShowFocusRect Then
            DrawFocusRect RGB(137, 173, 228)
        End If
    End If
    
    IsGotFocus = True
End Sub

Private Sub UserControl_Initialize()
    If IsFontInstalled(UserControl.hdc, "Arial") Then
        UnoFont(0) = "Arial"
    Else
        UnoFont(0) = GetFontAvailable
    End If
    
    If IsFontInstalled(UserControl.hdc, "Arial Narrow") Then
        UnoFont(1) = "Arial Narrow"
    Else
        UnoFont(1) = GetFontAvailable
    End If
    
    If IsFontInstalled(UserControl.hdc, "Arial Black") Then
        UnoFont(2) = "Arial Black"
    Else
        UnoFont(2) = GetFontAvailable
    End If

    MyProp.AutoUpdate = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not UserControl.Extender.TabStop Then Exit Sub
    
    Dim hWndParen  As Long
    
    hWndParen = GetParent(UserControl.hWnd)
    
    Select Case KeyCode
    Case Is = vbKeyRight
        KeyCode = 0
        PostMessage hWndParen, WM_KEYDOWN, ByVal &H27, ByVal &H4D0001
    Case Is = vbKeyDown
        KeyCode = 0
        PostMessage hWndParen, WM_KEYDOWN, ByVal &H28, ByVal &H500001
    Case Is = vbKeyLeft
        KeyCode = 0
        PostMessage hWndParen, WM_KEYDOWN, ByVal &H25, ByVal &H4B0001
    Case Is = vbKeyUp
        KeyCode = 0
        PostMessage hWndParen, WM_KEYDOWN, ByVal &H26, ByVal &H480001
    End Select
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_LostFocus()
    If IsHover Then
        DrawFocusRect RGB(229, 151, 0)
    Else
        UserControl.Cls
    End If
    
    IsGotFocus = False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrHover.Enabled = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    MyProp.AutoUpdate = PropBag.ReadProperty("AutoUpdate", True)
    MyProp.Data = PropBag.ReadProperty("Data", "")
    MyProp.Deck = PropBag.ReadProperty("Deck", uno_DCDefault)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    MyProp.Face = PropBag.ReadProperty("Face", uno_FCUp)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", vbDefault)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    MyProp.Rank = PropBag.ReadProperty("Rank", uno_RCZero)
    MyProp.Points = PropBag.ReadProperty("Points", 0)
    MyProp.ShowFocusRect = PropBag.ReadProperty("ShowFocusRect", True)
    MyProp.Suit = PropBag.ReadProperty("Suit", uno_SCBlue)
End Sub

Private Sub UserControl_Resize()
    Call Redraw
End Sub

Private Sub UserControl_Show()
    Call Redraw
End Sub

Private Sub UserControl_Terminate()
    tmrHover.Enabled = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "AutoUpdate", MyProp.AutoUpdate, True
    PropBag.WriteProperty "Data", MyProp.Data, ""
    PropBag.WriteProperty "Deck", MyProp.Deck, uno_DCDefault
    PropBag.WriteProperty "Enabled", UserControl.Enabled, True
    PropBag.WriteProperty "Face", MyProp.Face, uno_FCUp
    PropBag.WriteProperty "MouseIcon", MouseIcon, Nothing
    PropBag.WriteProperty "MousePointer", UserControl.MousePointer, vbDefault
    PropBag.WriteProperty "Rank", MyProp.Rank, uno_RCZero
    PropBag.WriteProperty "Points", MyProp.Points, 0
    PropBag.WriteProperty "ShowFocusRect", MyProp.ShowFocusRect, True
    PropBag.WriteProperty "Suit", MyProp.Suit, uno_SCBlue
    PropBag.WriteProperty "Picture", Picture, Nothing
End Sub

Private Sub Redraw()
    If Not MyProp.AutoUpdate Then Exit Sub
    
    Dim BmpW   As Integer
    Dim BmpH   As Integer
    Dim CardID As Integer

    On Error Resume Next
    
    UserControl.Cls
    
    If MyProp.Face = uno_FCUp Then
        If MyProp.Rank <= uno_RCSkip Then
            CardID = CRD_OFFS_SUIT + MyProp.Suit
        Else
            CardID = IIf(MyProp.Rank = uno_RCWild, CRD_WILD_CARD, _
                                                   CRD_DRAW_FOUR)
        End If
    Else
        CardID = CRD_OFFS_DECK + MyProp.Deck
    End If
    
    BmpW = ScaleX(LoadResPicture(CardID, vbResBitmap).Width, vbHimetric, vbPixels)
    BmpH = ScaleX(LoadResPicture(CardID, vbResBitmap).Height, vbHimetric, vbPixels)
    Set UserControl.Picture = LoadResPicture(CardID, vbResBitmap)
    
    If MyProp.Face = uno_FCUp Then
        If MyProp.Rank <= uno_RCSkip Then
            Dim Rv  As Integer
            Dim Gv  As Integer
            Dim Bv  As Integer
            Dim Rt  As Integer
            Dim Gt  As Integer
            Dim Bt  As Integer
            Dim mx  As Integer
            Dim my  As Integer
            Dim Rnk As String
            Dim hdc As Long
            
            Select Case MyProp.Rank
            Case 0 To 9  ' 0..9
                Rnk = CStr(MyProp.Rank)
            Case Is = 10 ' Draw Two
                Rnk = "Draw Two"
            Case Is = 11 ' Reverse
                Rnk = "Reverse"
            Case Is = 12 ' Skip"
                Rnk = "Skip"
            End Select
            
            If (MyProp.Suit = uno_SCBlue) Or (MyProp.Suit = uno_SCRed) Then
                UserControl.ForeColor = RGB(255, 255, 255)
            Else
                UserControl.ForeColor = RGB(0, 0, 0)
            End If
            
            UserControl.FontSize = 9
            hdc = UserControl.hdc
            
            If MyProp.Rank <= uno_RCNine Then
                UserControl.FontName = UnoFont(0)
                TextOut hdc, 6, 2, Rnk, Len(Rnk)
            Else
                UserControl.FontName = UnoFont(1)
                TextOut hdc, 6, 1, Rnk, Len(Rnk)
            End If
            
            Select Case MyProp.Rank
            Case 0 To 9  ' 0..9
                Rnk = CStr(MyProp.Rank)
            Case Is = 10 ' Draw Two
                Rnk = "D"
            Case Is = 11 ' Reverse
                Rnk = "R"
            Case Is = 12 ' Skip
                Rnk = "S"
            End Select

            Select Case MyProp.Suit
            Case Is = 0 ' Blue
                Rv = 0:   Gv = 0:   Bv = 255
            Case Is = 1 ' Red
                Rv = 255: Gv = 0:   Bv = 0
            Case Is = 2 ' Green
                Rv = 0:   Gv = 255: Bv = 0
            Case Is = 3 ' Yellow
                Rv = 255: Gv = 255: Bv = 0
            End Select
            
            Rt = IIf(Rv = 0, Rv, Rv - &H3F)
            Gt = IIf(Gv = 0, Gv, Gv - &H3F)
            Bt = IIf(Bv = 0, Bv, Bv - &H3F)
            
            UserControl.FontName = UnoFont(2)
            UserControl.FontSize = 32
            
            mx = (BmpW - UserControl.TextWidth(Rnk)) / 2 - 2
            my = (BmpH - UserControl.TextHeight(Rnk)) / 2 - 2
            
            ' outline effect
            SetTextColor hdc, RGB(Rt, Gt, Bt)
            TextOut hdc, mx + 1, my + 1, Rnk, 1
            TextOut hdc, mx + 1, my + 1, Rnk, 1
            TextOut hdc, mx + 1, my - 1, Rnk, 1
            TextOut hdc, mx - 1, my + 1, Rnk, 1
            TextOut hdc, mx - 1, my - 1, Rnk, 1
            
            ' shadow effect
            SetTextColor hdc, RGB(0, 0, 0)
            TextOut hdc, mx + 2, my + 2, Rnk, 1
        
            ' text
            SetTextColor hdc, RGB(Rv, Gv, Bv)
            TextOut hdc, mx, my, Rnk, 1
        End If
    End If
    
    Set UserControl.Picture = UserControl.Image
    
    If IsHover Then
        DrawFocusRect RGB(229, 151, 0)
    End If
    
    Select Case MyProp.Rank
    Case 0 To 9
        Points = MyProp.Rank
    Case Is = 10 ' draw two
        Points = 20
    Case Is = 11 ' reverse
        Points = 20
    Case Is = 12 ' skip
        Points = 20
    Case Is = 13 ' wild draw
        Points = 50
    Case Is = 14 ' draw four
        Points = 50
    End Select
End Sub

Private Sub DrawFocusRect(ByVal Color As Long)
    Dim rcRect As RECT
    
    UserControl.DrawWidth = 2
    GetClientRect UserControl.hWnd, rcRect
    UserControl.Line (1, 1)-(rcRect.Right - 1, _
                             rcRect.Bottom - 1), _
                     Color, B
    UserControl.DrawWidth = 1
End Sub

Private Function GetFontAvailable() As String
    Dim idx         As Integer
    Dim arrFont(8)  As String
    Dim DefaultFont As String * 255
    
    arrFont(0) = "Arial"
    arrFont(1) = "Arial Narrow"
    arrFont(2) = "Arial Black"
    arrFont(3) = "Book Antiqua"
    arrFont(4) = "Courier New"
    arrFont(5) = "Garamond"
    arrFont(6) = "Tahoma"
    arrFont(7) = "Times New Roman"
    arrFont(8) = "MS Sans Serif"
    
    For idx = UBound(arrFont()) To LBound(arrFont()) Step -1
        If IsFontInstalled(UserControl.hdc, arrFont(idx)) Then
            GetFontAvailable = arrFont(idx)
            Exit Function
        End If
    Next idx
    
    GetTextFace UserControl.hdc, Len(DefaultFont), DefaultFont
    GetFontAvailable = DefaultFont
End Function

Public Sub Refresh()
    UserControl.Refresh
End Sub


