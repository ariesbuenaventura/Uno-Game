VERSION 5.00
Begin VB.Form frmColor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Chooser"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3780
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2760
      TabIndex        =   7
      Top             =   1500
      Width           =   915
   End
   Begin VB.Frame fraColor 
      Caption         =   "Color"
      Height          =   1695
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   1635
      Begin VB.OptionButton optColor 
         Height          =   615
         Index           =   3
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   675
      End
      Begin VB.OptionButton optColor 
         Height          =   615
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Width           =   675
      End
      Begin VB.OptionButton optColor 
         Height          =   615
         Index           =   1
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   300
         Width           =   675
      End
      Begin VB.OptionButton optColor 
         Height          =   615
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   1500
      Width           =   915
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a color for the wild card you just played, then click OK button."
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1920
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelColor As Integer

Private Sub cmdCancel_Click()
    SelColor = -1
    
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = 301 To 304
        Set optColor(i Mod 301).Picture = _
            LoadResPicture(i, vbResBitmap)
    Next i
    
    SelColor = 0
End Sub

Private Sub optColor_Click(Index As Integer)
    SelColor = Index
End Sub
