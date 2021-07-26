VERSION 5.00
Begin VB.Form frmSettingAni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "#"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   3465
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   19
      Top             =   2460
      Width           =   795
   End
   Begin VB.Frame fraScatter 
      Height          =   2415
      Left            =   60
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin VB.ComboBox cmbScatterSpeed 
         Height          =   315
         Index           =   0
         ItemData        =   "frmSettingAni.frx":0000
         Left            =   960
         List            =   "frmSettingAni.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbScatterSpeed 
         Height          =   315
         Index           =   1
         ItemData        =   "frmSettingAni.frx":0045
         Left            =   960
         List            =   "frmSettingAni.frx":0067
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblScatterSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed X : "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   420
         Width           =   750
      End
      Begin VB.Label lblScatterSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed Y : "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   780
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1800
      TabIndex        =   15
      Top             =   2460
      Width           =   795
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   960
      TabIndex        =   14
      Top             =   2460
      Width           =   795
   End
   Begin VB.Frame fraBounce 
      Height          =   2415
      Left            =   60
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin VB.ComboBox cmbBounceSpeed 
         Height          =   315
         Index           =   1
         ItemData        =   "frmSettingAni.frx":008A
         Left            =   1140
         List            =   "frmSettingAni.frx":00AC
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox cmbBounceSpeed 
         Height          =   315
         Index           =   0
         ItemData        =   "frmSettingAni.frx":00CF
         Left            =   1140
         List            =   "frmSettingAni.frx":00F1
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cmbBounceDist 
         Height          =   315
         Index           =   1
         ItemData        =   "frmSettingAni.frx":0114
         Left            =   1140
         List            =   "frmSettingAni.frx":0136
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   660
         Width           =   1095
      End
      Begin VB.ComboBox cmbBounceDist 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         ItemData        =   "frmSettingAni.frx":0159
         Left            =   1140
         List            =   "frmSettingAni.frx":017B
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label lblBounceSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed Y : "
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   1740
         Width           =   750
      End
      Begin VB.Label lblBounceSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed X : "
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   1380
         Width           =   750
      End
      Begin VB.Line linLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   120
         X2              =   3180
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Line linLine 
         Index           =   0
         X1              =   120
         X2              =   3240
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label lblBounceDist 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distance Y : "
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   660
         Width           =   915
      End
      Begin VB.Label lblBounceDist 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distance X : "
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame fraSpin 
      Height          =   2415
      Left            =   60
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin VB.ComboBox cmbSpinDist 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmSettingAni.frx":019E
         Left            =   1020
         List            =   "frmSettingAni.frx":01DE
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblSpinDist 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distance : "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   420
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmSettingAni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurWinAni As Integer

Private Sub cmbBounceDist_Click(Index As Integer)
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub cmbBounceDist_GotFocus(Index As Integer)
    lblBounceDist(Index).ForeColor = vbRed
    cmbBounceDist(Index).ForeColor = vbBlue
    cmbBounceDist(Index).BackColor = &HEAFDFD
End Sub

Private Sub cmbBounceDist_LostFocus(Index As Integer)
    lblBounceDist(Index).ForeColor = vbBlack
    cmbBounceDist(Index).ForeColor = vbBlack
    cmbBounceDist(Index).BackColor = vbWhite
End Sub

Private Sub cmbBounceSpeed_Click(Index As Integer)
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub cmbBounceSpeed_GotFocus(Index As Integer)
    lblBounceSpeed(Index).ForeColor = vbRed
    cmbBounceSpeed(Index).ForeColor = vbBlue
    cmbBounceSpeed(Index).BackColor = &HEAFDFD
End Sub

Private Sub cmbBounceSpeed_LostFocus(Index As Integer)
    lblBounceSpeed(Index).ForeColor = vbBlack
    cmbBounceSpeed(Index).ForeColor = vbBlack
    cmbBounceSpeed(Index).BackColor = vbWhite
End Sub

Private Sub cmbScatterSpeed_Click(Index As Integer)
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub cmbScatterSpeed_GotFocus(Index As Integer)
    lblScatterSpeed(Index).ForeColor = vbRed
    cmbScatterSpeed(Index).ForeColor = vbBlue
    cmbScatterSpeed(Index).BackColor = &HEAFDFD
End Sub

Private Sub cmbScatterSpeed_LostFocus(Index As Integer)
    lblScatterSpeed(Index).ForeColor = vbBlack
    cmbScatterSpeed(Index).ForeColor = vbBlack
    cmbScatterSpeed(Index).BackColor = vbWhite
End Sub

Private Sub cmbSpinDist_Click()
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub cmbSpinDist_GotFocus()
    lblSpinDist.ForeColor = vbRed
    cmbSpinDist.ForeColor = vbBlue
    cmbSpinDist.BackColor = &HEAFDFD
End Sub

Private Sub cmbSpinDist_LostFocus()
    lblSpinDist.ForeColor = vbBlack
    cmbSpinDist.ForeColor = vbBlack
    cmbSpinDist.BackColor = vbWhite
End Sub

Private Sub cmdApply_Click()
    With frmOptions
        If CurWinAni = 1 Then
            ' Bounce
            .CustomAni.BounceDistX = Val(cmbBounceDist(0).Text)
            .CustomAni.BounceDistY = Val(cmbBounceDist(1).Text)
            
            .CustomAni.BounceSpeedX = Val(cmbBounceSpeed(0).Text)
            .CustomAni.BounceSpeedY = Val(cmbBounceSpeed(1).Text)
        ElseIf CurWinAni = 2 Then
            ' Scatter
            .CustomAni.ScatterSpeedX = Val(cmbScatterSpeed(0).Text)
            .CustomAni.ScatterSpeedY = Val(cmbScatterSpeed(1).Text)
        ElseIf CurWinAni = 3 Then
            ' Spin
            .CustomAni.SpinDistance = Val(cmbSpinDist.Text)
        End If
        
        .CustomAni.Reset = True
    End With
    
    If cmdApply.Enabled Then cmdApply.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If cmdApply.Enabled Then Call cmdApply_Click
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    With frmOptions
        If .optWinAni(1).Value Then
            CurWinAni = 1
        ElseIf .optWinAni(2).Value Then
            CurWinAni = 2
        ElseIf .optWinAni(3).Value Then
            CurWinAni = 3
        Else
            CurWinAni = 4
        End If
        
        Me.Caption = .optWinAni(CurWinAni).Caption
        Set Me.Icon = .optWinAni(CurWinAni).Picture
        
        ' Bounce
        If CurWinAni = 1 Then
            cmbBounceDist(0).ListIndex = .CustomAni.BounceDistX - 1
            cmbBounceDist(1).ListIndex = .CustomAni.BounceDistY - 1
            
            cmbBounceSpeed(0).ListIndex = .CustomAni.BounceSpeedX - 1
            cmbBounceSpeed(1).ListIndex = .CustomAni.BounceSpeedY - 1
            
            fraBounce.Visible = True
        ElseIf CurWinAni = 2 Then
            cmbScatterSpeed(0).ListIndex = .CustomAni.ScatterSpeedX - 1
            cmbScatterSpeed(1).ListIndex = .CustomAni.ScatterSpeedY - 1
            
            fraScatter.Visible = True
        ElseIf CurWinAni = 3 Then
            cmbSpinDist.ListIndex = .CustomAni.SpinDistance
            
            fraSpin.Visible = True
        End If
    End With
End Sub

