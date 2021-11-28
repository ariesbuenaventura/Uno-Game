VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "About Uno Game 2.0"
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   220
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   4
      Left            =   60
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   60
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   60
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   60
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmrAni 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   1320
   End
   Begin VB.Image imgExit 
      Height          =   510
      Left            =   3840
      MousePointer    =   99  'Custom
      Picture         =   "frmAbout.frx":0000
      Top             =   2640
      Width           =   510
   End
   Begin VB.Image imgEmail 
      Height          =   195
      Index           =   0
      Left            =   960
      MousePointer    =   99  'Custom
Tag             =   "ariesbuenaventura2019@gmail.com"
      Top             =   2880
      Width           =   2655
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const HALFTONE = 4

Private Declare Function GetStretchBltMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type AtomInfo
    Angle  As Integer
    XPos   As Single
    YPos   As Single
    Width  As Integer
    Height As Integer
End Type

Dim WinW                   As Long     ' window width
Dim WinH                   As Long     ' window height
Dim B_Atom(1 To 4)         As AtomInfo ' big atom
Dim S_Atom(1 To 4, 1 To 3) As AtomInfo ' small atom

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim i        As Integer
    Dim j        As Integer
    Dim AboutLoc As String

    On Error Resume Next
    
    ' note: i renamed uno.jpg to uno.dat so that no one
    '       will think that it is a jpeg file (nice trick).
    AboutLoc = App.Path & "\uno.dat"
    If Dir$(AboutLoc) <> "" Then
        Set Me.Picture = LoadPicture(AboutLoc)
    End If
    
    Me.PaintPicture imgExit.Picture, imgExit.Left, imgExit.Top
    Set imgExit.Picture = Nothing
    
    Set Me.Picture = Me.Image ' convert image to picture
    Set imgExit.MouseIcon = LoadResPicture(101, vbResCursor)
    Set imgEmail(0).MouseIcon = LoadResPicture(101, vbResCursor)
    Set imgEmail(1).MouseIcon = LoadResPicture(101, vbResCursor)
    
    With frmMain
        For i = 1 To .imlStatIcons.ListImages.Count - 1
            Set picCanvas(i).Picture = .imlStatIcons.ListImages(i).Picture
            
            B_Atom(i).Angle = i * 90 ' 0°, 90°, 180°, 270°
            B_Atom(i).Width = picCanvas(i).ScaleWidth
            B_Atom(i).Height = picCanvas(i).ScaleHeight
        Next i
            
        For i = 1 To 3
            For j = 1 To 4
                S_Atom(j, i).Angle = 30 + ((i - 1) * 120) ' 30°, 150°, 270°
                S_Atom(j, i).Width = picCanvas(j).ScaleWidth / 2
                S_Atom(j, i).Height = picCanvas(j).ScaleHeight / 2
            Next j
        Next i
    End With
    
    WinW = Me.ScaleWidth
    WinH = Me.ScaleHeight
    
    tmrAni.Enabled = True
End Sub

Private Sub imgEmail_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgEmail(Index).MouseIcon = LoadResPicture(102, vbResCursor)
End Sub

Private Sub imgEmail_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    Set imgEmail(Index).MouseIcon = LoadResPicture(101, vbResCursor)
    
    If Index = 0 Then
        ShellExecute 0, "open", "mailto:" & imgEmail(Index).Tag, 0, 0, 0
    Else
        ShellExecute 0, "open", imgEmail(Index).Tag, 0, 0, 0
    End If
End Sub

Private Sub imgExit_Click()
    Unload Me
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgExit.MouseIcon = LoadResPicture(102, vbResCursor)
End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgExit.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub tmrAni_Timer()
    Dim i        As Integer
    Dim j        As Integer
    Dim px       As Single
    Dim py       As Single
    Dim Radian   As Single
    Dim Index    As Integer
    Dim Diameter As Single
    
    Me.Cls
    
    For i = picCanvas.LBound To picCanvas.UBound
        DrawBigAtom B_Atom(i), i, 40, 2
    Next i
    
    For i = 1 To 4
        For j = 1 To 3
            If j = 1 Then
                Index = (i + 4) Mod 4 + 1
            ElseIf j = 2 Then
                Index = (i + 1) Mod 4 + 1
            Else
                Index = (i + 2) Mod 4 + 1
            End If
            
            DrawSmallAtom S_Atom(i, j), Index, _
                          B_Atom(i).XPos, B_Atom(i).YPos, 15, 5
        Next j
    Next i
End Sub

Private Sub DrawBigAtom(ByRef B_Atom As AtomInfo, _
                        ByVal Index As Integer, _
                        ByVal Radius As Integer, _
                        ByVal Increment As Integer)
                     
    
    Dim px       As Single
    Dim py       As Single
    Dim Radian   As Single
    Dim Diameter As Single
    Dim CurPos   As POINTAPI
    
    GetCursorPos CurPos
    ScreenToClient Me.hwnd, CurPos
    
    Radian = Rads(B_Atom.Angle)
    
    ' polar graph formula
    px = Radius * Cos(Radian)
    py = Radius * Sin(Radian)
    
    Diameter = Sqr(px * px + py * py) * 2 ' or Diameter = Radius * 2. However, that one is more accurate.
    B_Atom.XPos = CurPos.X + (Diameter - B_Atom.Width) / 2 + px - Diameter / 2
    B_Atom.YPos = CurPos.Y + (Diameter - B_Atom.Height) / 2 - py - Diameter / 2
    
    BitBlt Me.hdc, B_Atom.XPos, B_Atom.YPos, _
                   B_Atom.Width, B_Atom.Height, _
           picCanvas(Index).hdc, 0, 0, vbSrcInvert
    BitBlt Me.hdc, B_Atom.XPos, B_Atom.YPos, _
                   B_Atom.Width, B_Atom.Height, _
           picCanvas(Index).hdc, 0, 0, vbSrcAnd
    BitBlt Me.hdc, B_Atom.XPos, B_Atom.YPos, _
                   B_Atom.Width, B_Atom.Height, _
           picCanvas(Index).hdc, 0, 0, vbSrcInvert
    
    
    RefreshWindow Me.hwnd ' same as Me.Refresh but more faster.
    
    If B_Atom.Angle > 360 Then
        ' since that 360° = 0°, 375° = 15°...by getting
        ' the reminder of a given angle divided by 360°
        ' we can also get the same result.
        
        '   Ex. 450° mod 360° = 90°
        
        B_Atom.Angle = B_Atom.Angle Mod 360
    Else
        B_Atom.Angle = B_Atom.Angle + Increment
    End If
End Sub

Private Sub DrawSmallAtom(ByRef B_Atom As AtomInfo, _
                          ByVal Index As Integer, _
                          ByVal X As Integer, _
                          ByVal Y As Integer, _
                          ByVal Radius As Integer, _
                          ByVal Increment As Integer)
                                                  
    Dim W          As Integer
    Dim H          As Integer
    Dim px         As Single
    Dim py         As Single
    Dim Radian     As Single
    Dim Diameter   As Single
    Dim OldBltMode As Long
    
    W = picCanvas(Index).ScaleWidth
    H = picCanvas(Index).ScaleHeight
    
    Radian = Rads(B_Atom.Angle)
    
    ' polar graph formula
    px = Radius * Cos(Radian)
    py = Radius * Sin(Radian)
    
    Diameter = Sqr(px * px + py * py) * 2 ' or Diameter = Radius * 2. However, that one is more accurate.
    B_Atom.XPos = X + (Diameter - B_Atom.Width) / 2 + px - Diameter / 4
    B_Atom.YPos = Y + (Diameter - B_Atom.Height) / 2 - py - Diameter / 4
        
    OldBltMode = GetStretchBltMode(Me.hdc)
    SetStretchBltMode Me.hdc, HALFTONE
    
    StretchBlt Me.hdc, B_Atom.XPos, B_Atom.YPos, _
                       B_Atom.Width, B_Atom.Height, _
               picCanvas(Index).hdc, 0, 0, W, H, vbSrcInvert
    StretchBlt Me.hdc, B_Atom.XPos, B_Atom.YPos, _
                       B_Atom.Width, B_Atom.Height, _
               picCanvas(Index).hdc, 0, 0, W, H, vbSrcAnd
    StretchBlt Me.hdc, B_Atom.XPos, B_Atom.YPos, _
                       B_Atom.Width, B_Atom.Height, _
               picCanvas(Index).hdc, 0, 0, W, H, vbSrcInvert
    
    If B_Atom.Angle > 360 Then
        ' since that 360° = 0°, 375° = 15°...by getting
        ' the reminder of a given angle divided by 360°
        ' we can also get the same result.
        
        '   Ex. 450° mod 360° = 90°
        B_Atom.Angle = B_Atom.Angle Mod 360
    Else
        B_Atom.Angle = B_Atom.Angle + Increment
    End If
    
    SetStretchBltMode Me.hdc, OldBltMode
End Sub
