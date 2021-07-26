VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "*\A..\Card\prjUnoCard.vbp"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList imlIconsDisabled 
      Left            =   2640
      Top             =   5220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":2368
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":2C42
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":351C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":3DF6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIconsGrayed 
      Left            =   2040
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":46D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":4FAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":5884
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":615E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":6A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":7312
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":7BEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":84C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIconsColor 
      Left            =   1440
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":8DA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":967A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":9F54
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":A82E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":B108
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":B9E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":C2BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":CB96
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      Height          =   315
      Left            =   60
      TabIndex        =   52
      Top             =   5280
      Width           =   795
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   4440
      TabIndex        =   51
      Top             =   5280
      Width           =   795
   End
   Begin MSComDlg.CommonDialog dlgOp 
      Left            =   900
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5280
      TabIndex        =   6
      Top             =   5280
      Width           =   795
   End
   Begin VB.Timer tmrAni 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1380
      Top             =   5280
   End
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
      Left            =   6120
      TabIndex        =   2
      Top             =   5280
      Width           =   795
   End
   Begin VB.Frame fraGeneral 
      Height          =   4755
      Left            =   120
      TabIndex        =   32
      Top             =   420
      Width           =   6735
      Begin VB.CheckBox chkShowTooltip 
         Caption         =   "Show &Tooltip"
         Height          =   435
         Left            =   180
         TabIndex        =   50
         Top             =   4200
         Width           =   1395
      End
      Begin VB.Frame fraName 
         Caption         =   "Player names"
         Height          =   1695
         Left            =   2280
         TabIndex        =   37
         Top             =   180
         Width           =   4335
         Begin VB.TextBox txtName 
            BackColor       =   &H00EAFDFD&
            Height          =   285
            Index           =   0
            Left            =   1140
            MaxLength       =   15
            TabIndex        =   41
            Top             =   300
            Width           =   3075
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Index           =   1
            Left            =   1140
            MaxLength       =   15
            TabIndex        =   40
            Top             =   600
            Width           =   3075
         End
         Begin VB.TextBox txtName 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   1140
            MaxLength       =   15
            TabIndex        =   39
            Top             =   900
            Width           =   3075
         End
         Begin VB.TextBox txtName 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   1140
            MaxLength       =   15
            TabIndex        =   38
            Top             =   1200
            Width           =   3075
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "You"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   45
            Top             =   360
            Width           =   285
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Computer I"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   44
            Top             =   660
            Width           =   765
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Computer II"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   43
            Top             =   960
            Width           =   810
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Computer III"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   42
            Top             =   1260
            Width           =   855
         End
      End
      Begin VB.Frame fraOpponents 
         Caption         =   "Opponents"
         Height          =   1695
         Left            =   120
         TabIndex        =   33
         Top             =   180
         Width           =   2115
         Begin VB.OptionButton optOpponent 
            Caption         =   "Three"
            Height          =   255
            Index           =   2
            Left            =   420
            TabIndex        =   36
            Tag             =   "3"
            Top             =   1080
            Width           =   1515
         End
         Begin VB.OptionButton optOpponent 
            Caption         =   "Two"
            Height          =   255
            Index           =   1
            Left            =   420
            TabIndex        =   35
            Tag             =   "2"
            Top             =   720
            Width           =   1515
         End
         Begin VB.OptionButton optOpponent 
            Caption         =   "One"
            Height          =   255
            Index           =   0
            Left            =   420
            TabIndex        =   34
            Tag             =   "1"
            Top             =   360
            Value           =   -1  'True
            Width           =   1515
         End
      End
      Begin MSComctlLib.Slider sldSpeed 
         Height          =   270
         Left            =   180
         TabIndex        =   63
         Top             =   2400
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   476
         _Version        =   393216
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "Speed"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   2100
         Width           =   465
      End
   End
   Begin VB.Frame fraDifficulty 
      Caption         =   "Select Difficulty"
      Height          =   4695
      Left            =   120
      TabIndex        =   46
      Top             =   480
      Width           =   6735
      Begin VB.OptionButton optLevel 
         Caption         =   "Normal"
         Height          =   435
         Index           =   1
         Left            =   420
         TabIndex        =   49
         Top             =   900
         Width           =   975
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "Easy"
         Height          =   435
         Index           =   0
         Left            =   420
         TabIndex        =   48
         Top             =   540
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "Difficult"
         Height          =   435
         Index           =   2
         Left            =   420
         TabIndex        =   47
         Top             =   1260
         Width           =   975
      End
   End
   Begin VB.Frame fraAni 
      Height          =   4755
      Left            =   120
      TabIndex        =   4
      Top             =   420
      Width           =   6735
      Begin VB.PictureBox picViewer 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00008000&
         Height          =   3435
         Left            =   60
         ScaleHeight     =   225
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   349
         TabIndex        =   5
         Top             =   600
         Width           =   5295
         Begin VB.PictureBox picCanvas 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   0
            ScaleHeight     =   25
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   21
            TabIndex        =   61
            Top             =   60
            Visible         =   0   'False
            Width           =   315
         End
      End
      Begin VB.Frame fraWinType 
         Height          =   4755
         Left            =   5400
         TabIndex        =   54
         Top             =   0
         Width           =   1335
         Begin VB.OptionButton optWinAni 
            Caption         =   "Wind"
            Height          =   555
            Index           =   5
            Left            =   120
            Picture         =   "frmOptions.frx":D470
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   2880
            Width           =   1095
         End
         Begin VB.OptionButton optWinAni 
            Caption         =   "Fall"
            Height          =   555
            Index           =   4
            Left            =   120
            Picture         =   "frmOptions.frx":DE2A
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   2340
            Width           =   1095
         End
         Begin VB.ComboBox cmbMaxCards 
            Height          =   315
            ItemData        =   "frmOptions.frx":E7E4
            Left            =   120
            List            =   "frmOptions.frx":E824
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   4320
            Width           =   1095
         End
         Begin VB.OptionButton optWinAni 
            Caption         =   "(None)"
            Height          =   555
            Index           =   0
            Left            =   120
            Picture         =   "frmOptions.frx":E86F
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   180
            Width           =   1095
         End
         Begin VB.OptionButton optWinAni 
            Caption         =   "Spin"
            Height          =   555
            Index           =   3
            Left            =   120
            Picture         =   "frmOptions.frx":E9C1
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   1800
            Width           =   1095
         End
         Begin VB.OptionButton optWinAni 
            Caption         =   "Scatter"
            Height          =   555
            Index           =   2
            Left            =   120
            Picture         =   "frmOptions.frx":EF4B
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   1260
            Width           =   1095
         End
         Begin VB.OptionButton optWinAni 
            Caption         =   "Bounce"
            Height          =   555
            Index           =   1
            Left            =   120
            Picture         =   "frmOptions.frx":F4D5
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   720
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.Label lblCards 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max. Cards"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   4080
            Width           =   795
         End
      End
      Begin MSComctlLib.Toolbar tblToolbar 
         Height          =   570
         Left            =   60
         TabIndex        =   65
         Top             =   4080
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlIconsGrayed"
         DisabledImageList=   "imlIconsDisabled"
         HotImageList    =   "imlIconsColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "pause"
               Object.ToolTipText     =   "Pause"
               ImageIndex      =   7
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "play"
               Object.ToolTipText     =   "Play"
               ImageIndex      =   6
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "trail"
               Object.ToolTipText     =   "Trail"
               ImageIndex      =   5
               Style           =   1
               Value           =   1
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "from_bottom"
               Object.ToolTipText     =   "From Bottom"
               ImageIndex      =   1
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "from_top"
               Object.ToolTipText     =   "From Top"
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "from_right"
               Object.ToolTipText     =   "From Right"
               ImageIndex      =   3
               Style           =   2
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "from_left"
               Object.ToolTipText     =   "From Left"
               ImageIndex      =   4
               Style           =   2
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "setting"
               Object.ToolTipText     =   "Settings"
               ImageIndex      =   8
            EndProperty
         EndProperty
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Select the type of animation would you like to play when user wins the game."
         Height          =   435
         Index           =   0
         Left            =   120
         TabIndex        =   53
         Top             =   180
         Width           =   5175
      End
   End
   Begin VB.Frame fraBk 
      Height          =   4755
      Left            =   120
      TabIndex        =   7
      Top             =   420
      Width           =   6735
      Begin VB.CheckBox chkBkPreview 
         Caption         =   "Preview"
         Height          =   195
         Left            =   360
         TabIndex        =   31
         Top             =   4440
         Value           =   2  'Grayed
         Width           =   915
      End
      Begin VB.PictureBox picEye 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   2
         Left            =   120
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   30
         Top             =   4440
         Width           =   210
      End
      Begin VB.Frame fraBkOp 
         Caption         =   "Options"
         Height          =   4095
         Left            =   4920
         TabIndex        =   10
         Top             =   180
         Width           =   1695
         Begin VB.CommandButton cmdColor 
            Caption         =   "Change to Color"
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
            Left            =   120
            TabIndex        =   12
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CommandButton cmdBitmap 
            Caption         =   "Change to Bitmap"
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
            Left            =   120
            TabIndex        =   11
            Top             =   2040
            Width           =   1455
         End
      End
      Begin VB.Frame fraPreview 
         Height          =   4095
         Left            =   120
         TabIndex        =   8
         Top             =   180
         Width           =   4755
         Begin VB.PictureBox picPreview 
            AutoRedraw      =   -1  'True
            Height          =   3615
            Left            =   120
            ScaleHeight     =   237
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   297
            TabIndex        =   9
            Top             =   300
            Width           =   4515
         End
      End
   End
   Begin MSComctlLib.TabStrip tbsOp 
      Height          =   5175
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9128
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Animation"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Background"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Deck"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "D&ifficulty"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Sort"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraDeck 
      Caption         =   "Select Card Back"
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6735
      Begin VB.PictureBox picEye 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   120
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   29
         Top             =   4380
         Width           =   210
      End
      Begin VB.CheckBox chkDeckPreview 
         Caption         =   "Preview"
         Height          =   195
         Left            =   360
         TabIndex        =   27
         Top             =   4380
         Value           =   2  'Grayed
         Width           =   915
      End
      Begin prjUnoCard.UnoCard crdDeck 
         Height          =   855
         Index           =   0
         Left            =   480
         TabIndex        =   64
         Top             =   540
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1508
         Face            =   1
         ShowFocusRect   =   0   'False
         Picture         =   "frmOptions.frx":FA5F
      End
      Begin VB.Shape shpRect 
         BorderColor     =   &H00EE8269&
         BorderWidth     =   7
         Height          =   495
         Left            =   3840
         Top             =   1200
         Width           =   555
      End
   End
   Begin VB.Frame fraSort 
      Height          =   4755
      Left            =   120
      TabIndex        =   13
      Top             =   420
      Width           =   6735
      Begin VB.PictureBox picEye 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   3360
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   28
         Top             =   4320
         Width           =   210
      End
      Begin VB.CheckBox chkSortPreview 
         Caption         =   "Preview"
         Height          =   195
         Left            =   3600
         TabIndex        =   26
         Top             =   4320
         Value           =   2  'Grayed
         Width           =   1035
      End
      Begin VB.PictureBox picActiveSort 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1320
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   120
         TabIndex        =   25
         Top             =   4200
         Width           =   1800
      End
      Begin VB.CheckBox chkAutoSort 
         Caption         =   "Auto Sort"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   4320
         Width           =   1035
      End
      Begin VB.PictureBox picSortMode 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   405
         Index           =   4
         Left            =   600
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   23
         ToolTipText     =   "Wild Card, Draw Four, Blue, Red, Green, Yellow"
         Top             =   2340
         Width           =   1340
      End
      Begin VB.OptionButton optSortMode 
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   22
         Top             =   2460
         Width           =   195
      End
      Begin VB.OptionButton optSortMode 
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   21
         Top             =   1980
         Width           =   195
      End
      Begin VB.OptionButton optSortMode 
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   20
         Top             =   1500
         Width           =   195
      End
      Begin VB.OptionButton optSortMode 
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   19
         Top             =   1020
         Width           =   195
      End
      Begin VB.OptionButton optSortMode 
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   18
         Top             =   540
         Width           =   195
      End
      Begin VB.PictureBox picSortMode 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   405
         Index           =   0
         Left            =   600
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   17
         ToolTipText     =   "Blue, Red, Green, Yellow, Wild Card, Draw Four"
         Top             =   420
         Width           =   1340
      End
      Begin VB.PictureBox picSortMode 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   405
         Index           =   1
         Left            =   600
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   16
         ToolTipText     =   "Red, Green, Yellow, Wild Card, Draw Four, Blue"
         Top             =   900
         Width           =   1340
      End
      Begin VB.PictureBox picSortMode 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   405
         Index           =   2
         Left            =   600
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   15
         ToolTipText     =   "Green, Yellow, Wild Card, Draw Four, Blue, Red"
         Top             =   1380
         Width           =   1340
      End
      Begin VB.PictureBox picSortMode 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   405
         Index           =   3
         Left            =   600
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   14
         ToolTipText     =   "Yellow, Wild Card, Draw Four, Blue, Red, Green"
         Top             =   1860
         Width           =   1340
      End
      Begin VB.Line linHoriz 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   120
         X2              =   6660
         Y1              =   4155
         Y2              =   4155
      End
      Begin VB.Line linHoriz 
         Index           =   0
         X1              =   120
         X2              =   6600
         Y1              =   4140
         Y2              =   4140
      End
      Begin VB.Shape shpFocusRect 
         BorderColor     =   &H00EE8269&
         BorderWidth     =   5
         Height          =   495
         Left            =   4200
         Top             =   360
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Temp_Set As GAME_SETTING
Dim Temp_Col As New Collection
Dim Temp_Obj As Object

Public CustomAni As New clsAni

Private Sub chkAutoSort_Click()
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub chkBkPreview_Click()
    picEye(2).Visible = CBool(chkBkPreview.Value)
End Sub

Private Sub chkDeckPreview_Click()
    picEye(1).Visible = CBool(chkDeckPreview.Value)
End Sub

Private Sub chkShowTooltip_Click()
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub chkSortPreview_Click()
    picEye(0).Visible = CBool(chkSortPreview.Value)
End Sub

Private Sub cmdApply_Click()
    Dim i As Integer
    
    With frmMain
        If Temp_Set.Opponents > 0 Then
            For i = 0 To Temp_Set.Opponents
                Setting.PlayerName(i) = txtName(i).Text
            Next i
        End If
        
        If Temp_Set.BkFileLoc <> "" Then
            Dim rcRect As RECT
            
            Setting.BkFileLoc = Temp_Set.BkFileLoc
            Set Setting.BkPicture = LoadPicture(Temp_Set.BkFileLoc)
            GetClientRect .picTable.hwnd, rcRect
            Background .picTable.hdc, LoadPicture(Temp_Set.BkFileLoc), _
                       rcRect.Right, rcRect.Bottom
            RefreshWindow .picTable.hwnd
        Else
            .picTable.Cls
            .picTable.BackColor = picPreview.BackColor
            Set Setting.BkPicture = Nothing
            Setting.BkFileLoc = ""
        End If
        
        If Setting.Deck <> Temp_Set.Deck Then
            ChangeDeck Temp_Set.Deck
            Setting.Deck = Temp_Set.Deck
        End If
        
        If Setting.Difficulty <> Temp_Set.Difficulty Then
            Setting.Difficulty = Temp_Set.Difficulty
        End If
        
        If Temp_Set.Opponents > 0 Then
            If Setting.Opponents <> Temp_Set.Opponents Then
                .tblToolbar.Buttons("opponent").Image = Temp_Set.Opponents
            End If
        End If
        
        If Setting.WinnerSelAni <> Temp_Set.WinnerSelAni Then
            Setting.WinnerSelAni = Temp_Set.WinnerSelAni
            .WinnerAni.Reset = True
        End If
        
        If Setting.ShowTrail <> Temp_Set.ShowTrail Then
            Setting.ShowTrail = Temp_Set.ShowTrail
        End If
        
        If Setting.FallType <> Temp_Set.FallType Then
            Setting.FallType = Temp_Set.FallType
        End If
        
        If Setting.WindType <> Temp_Set.WindType Then
            Setting.WindType = Temp_Set.WindType
        End If
        
        If Setting.BkColor <> Temp_Set.BkColor Then
            Setting.BkColor = Temp_Set.BkColor
        End If
        
        If Setting.SortMode <> Temp_Set.SortMode Then
            ChangeSort Temp_Set.SortMode
            .imcSortMode.ComboItems(Temp_Set.SortMode + 1).Selected = True
            Setting.SortMode = Temp_Set.SortMode
            Set picActiveSort.Picture = picSortMode(Temp_Set.SortMode).Picture
        End If
        
        If Setting.Speed <> Temp_Set.Speed Then
            Setting.Speed = Temp_Set.Speed
            .LA.StopAni = True
            .LA.Speed = Setting.Speed
        End If
        
        If Setting.MaxCard <> Temp_Set.MaxCard Then
            Setting.MaxCard = Temp_Set.MaxCard
            .WinnerAni.Reset = True
            .WinnerAni.MaxCards = Setting.MaxCard
        End If
        
        Setting.BounceDistX = CustomAni.BounceDistX
        Setting.BounceDistY = CustomAni.BounceDistY
        Setting.BounceSpeedX = CustomAni.BounceSpeedX
        Setting.BounceSpeedY = CustomAni.BounceSpeedY
        Setting.ScatterSpeedX = CustomAni.ScatterSpeedX
        Setting.ScatterSpeedY = CustomAni.ScatterSpeedY
        Setting.SpinDist = CustomAni.SpinDistance
        
        .tblToolbar.Buttons("tooltip").Value = chkShowTooltip.Value
        .chkAutoSort.Value = chkAutoSort.Value
    End With
    
    If cmdApply.Enabled Then cmdApply.Enabled = False
End Sub

Private Sub cmdBitmap_Click()
    On Error GoTo ErrHandler
    
    With dlgOp
        .Filter = "Bitmap Files (*.bmp) | *.bmp; |" _
                & "JPEG (*.JPG,*.JPEG) | *.jpg; *.jpeg; |" _
                & "GIF (*.GIF) | *.GIF; |" _
                & "All Picture Files | *.bmp; *.gif; *.jpg; *.jpeg; |" _
                & "All Files (*.*) |*.*"
                
        .FilterIndex = 4
        .InitDir = ""
        .Filename = ""
        .ShowOpen
            
        If .Filename <> "" Then
            Dim rcRect As RECT
            
            Temp_Set.BkFileLoc = .Filename
            GetClientRect picPreview.hwnd, rcRect
            Background picPreview.hdc, LoadPicture(.Filename), _
                       rcRect.Right, rcRect.Bottom
            RefreshWindow picPreview.hwnd
        End If
    End With

    If Not cmdApply.Enabled Then cmdApply.Enabled = True
    Exit Sub
    
ErrHandler:
    If Err.Number = 32755 Then
        ' Cancel Selected
    Else
        MsgBox Err.Description, vbOKOnly Or vbInformation, "Uno Game 2.0"
    End If
End Sub

Private Sub cmdCancel_Click()
    If Setting.Deck <> Temp_Set.Deck Then
        ChangeDeck Setting.Deck
    End If
    
    If Setting.SortMode <> Temp_Set.SortMode Then
        ChangeSort Setting.SortMode
    End If
    
    Unload Me
End Sub

Private Sub cmdColor_Click()
    On Error GoTo ErrHandler 'Resume Next
    
    dlgOp.ShowColor
    picPreview.BackColor = dlgOp.Color
    Temp_Set.BkColor = dlgOp.Color
    Temp_Set.BkFileLoc = ""

    If Not cmdApply.Enabled Then cmdApply.Enabled = True
    Exit Sub
    
ErrHandler:
    ' cancel was selected
End Sub

Private Sub cmbMaxCards_Click()
    tmrAni.Enabled = False
    CustomAni.Reset = True
    CustomAni.MaxCards = Val(cmbMaxCards.Text)
    Temp_Set.MaxCard = Val(cmbMaxCards.Text)
    tmrAni.Enabled = True
    
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub cmdDefault_Click()
    ' General
    optOpponent(0).Value = True
    txtName(0).Text = "Juan"
    txtName(1).Text = "Pedro"
    txtName(2).Text = "Maria"
    txtName(3).Text = "Piso"
    chkShowTooltip.Value = vbChecked
    sldSpeed.Value = 1
    
    Temp_Set.Opponents = 1
    Temp_Set.PlayerName(0) = "Juan"
    Temp_Set.PlayerName(1) = "Pedro"
    Temp_Set.PlayerName(2) = "Maria"
    Temp_Set.PlayerName(3) = "Piso"
    Temp_Set.Speed = 1
    
    ' Animation
    optWinAni(1).Value = True
    cmbMaxCards.ListIndex = 4
    tblToolbar.Buttons("trail").Value = tbrUnpressed
    tblToolbar.Buttons("from_right").Value = tbrPressed
    
    Temp_Set.FallType = 0
    Temp_Set.ShowTrail = False
    
    ' Background
    picPreview.Cls
    picPreview.BackColor = &H8000&
    chkBkPreview.Value = vbChecked
    
    Temp_Set.BkColor = &H8000&
    Temp_Set.BkFileLoc = ""
    Set Temp_Set.BkPicture = Nothing
    
    ' Deck
    Call crdDeck_Click(0)
    crdDeck(0).SetFocus
    
    ' Difficulty
    optLevel(0).Value = True
    
    Temp_Set.Difficulty = 0
    
    ' Sort
    optSortMode(0).Value = True
    chkAutoSort.Value = vbChecked
    chkSortPreview.Value = vbChecked
    
    Temp_Set.SortMode = 0
    
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub cmdOk_Click()
    If cmdApply.Enabled Then Call cmdApply_Click
    Unload Me
End Sub

Private Sub crdDeck_Click(Index As Integer)
    If CBool(chkDeckPreview.Value) Then
        ChangeDeck Index
    End If
    
    Temp_Set.Deck = Index
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub crdDeck_GotFocus(Index As Integer)
    shpRect.Move crdDeck(Index).Left, crdDeck(Index).Top, _
                 crdDeck(Index).Width, crdDeck(Index).Height
End Sub

Private Sub Form_Load()
    Dim i       As Integer
    Dim j       As Integer
    Dim TempW   As Integer
    Dim curDeck As Integer
    
    crdDeck(0).Width = ScaleX(CardW, vbPixels, vbTwips)
    crdDeck(0).Height = ScaleY(CardH, vbPixels, vbTwips)
    crdDeck(0).Move (fraDeck.Width - crdDeck(0).Width * 5 - 480) / 2, _
                    (fraDeck.Height - crdDeck(0).Height) / 2

    For curDeck = 1 To 4
        Load crdDeck(curDeck)
        
        crdDeck(curDeck).Visible = True
        crdDeck(curDeck).Deck = curDeck
        crdDeck(curDeck).Left = crdDeck(curDeck - 1).Left + _
                                crdDeck(curDeck).Width + 120
    Next curDeck
    
    For i = 0 To 20
        If i <> 0 Then
            Load picCanvas(picCanvas.Count)
        End If
        
        Set picCanvas(i).Picture = _
            LoadResPicture(201 + i Mod 4, vbResBitmap)
    Next i
    
    For i = 0 To 4
        picSortMode(i).Width = 1800
        
        For j = 301 To 305
            picSortMode(i).PaintPicture LoadResPicture((j Mod 301 + i) Mod 5 + 301, _
                                                          vbResBitmap), _
                                          24 * (j Mod 301) + 3, 0
            Set picSortMode(i).Picture = picSortMode(i).Image
        Next j
    Next i

    BitBlt picActiveSort.hdc, 0, 0, picActiveSort.ScaleWidth, picActiveSort.ScaleHeight, _
           picSortMode(Uno.GetSortMode).hdc, 0, 0, vbSrcCopy
    picActiveSort.Refresh
    Set picActiveSort.Picture = picActiveSort.Image
    picActiveSort.ToolTipText = picSortMode(Uno.GetSortMode).ToolTipText
    
    Set picEye(0).Picture = LoadResPicture(401, vbResBitmap)
    Set picEye(1).Picture = LoadResPicture(401, vbResBitmap)
    Set picEye(2).Picture = LoadResPicture(401, vbResBitmap)
        
    optOpponent(Setting.Opponents - 1).Value = True
    
    With frmMain
        txtName(0).Text = .lblPlayerName(0).Caption
        txtName(1).Text = .lblPlayerName(2).Caption
        txtName(2).Text = .lblPlayerName(1).Caption
        txtName(3).Text = .lblPlayerName(3).Caption
        
        Temp_Set.Deck = .crdStock(0).Deck
        Temp_Set.BkColor = .picTable.BackColor
        Temp_Set.SortMode = Uno.GetSortMode
        Temp_Set.Opponents = Setting.Opponents
        
        If Setting.BkFileLoc <> "" Then
            Background picPreview.hdc, Setting.BkPicture, _
                       picPreview.ScaleWidth, picPreview.ScaleHeight
        Else
            picPreview.BackColor = .picTable.BackColor
        End If
        
        chkAutoSort.Value = .chkAutoSort.Value
        chkShowTooltip.Value = .tblToolbar.Buttons("tooltip").Value
        optLevel(Setting.Difficulty).Value = True
        optOpponent(Setting.Opponents - 1).Value = True
        cmbMaxCards.ListIndex = Setting.MaxCard - 1
        sldSpeed.Value = Setting.Speed
        optWinAni(Setting.WinnerSelAni).Value = vbChecked
        tblToolbar.Buttons("setting").Enabled = CBool(Setting.WinnerSelAni)
        
        tblToolbar.Buttons("trail").Value = IIf(Setting.ShowTrail, tbrPressed, tbrUnpressed)
        
        If Setting.FallType = 0 Then
            tblToolbar.Buttons("from_bottom").Value = tbrPressed
        ElseIf Setting.FallType = 1 Then
            tblToolbar.Buttons("from_top").Value = tbrPressed
        ElseIf Setting.FallType = 2 Then
            tblToolbar.Buttons("from_right").Value = tbrPressed
        Else
            tblToolbar.Buttons("from_left").Value = tbrPressed
        End If
        
        If Setting.WinnerSelAni = 4 Then
            tblToolbar.Buttons("setting").Enabled = False
        Else
            tblToolbar.Buttons("setting").Enabled = CBool(Setting.WinnerSelAni)
        End If
        
        CustomAni.BounceDistX = Setting.BounceDistX
        CustomAni.BounceDistY = Setting.BounceDistY
        CustomAni.BounceSpeedX = Setting.BounceSpeedX
        CustomAni.BounceSpeedY = Setting.BounceSpeedY
        CustomAni.ScatterSpeedX = Setting.ScatterSpeedX
        CustomAni.ScatterSpeedY = Setting.ScatterSpeedY
        CustomAni.SpinDistance = Setting.SpinDist
        CustomAni.ShowTrail = Setting.ShowTrail
    End With
    
    cmbMaxCards.ListIndex = cmbMaxCards.TopIndex
    cmdApply.Enabled = False
End Sub

Private Sub optLevel_Click(Index As Integer)
    Temp_Set.Difficulty = Index
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub optOpponent_Click(Index As Integer)
    Dim i As Integer
    
    For i = txtName.LBound To txtName.UBound
        If i <= Val(optOpponent(Index).Tag) Then
            txtName(i).Enabled = True
            lblName(i).Enabled = True
        Else
            txtName(i).Enabled = False
            lblName(i).Enabled = False
        End If
    Next i
    
    Temp_Set.Opponents = Index + 1
    
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub optSortMode_Click(Index As Integer)
    shpFocusRect.Move picSortMode(Index).Left, _
                      picSortMode(Index).Top, _
                      picSortMode(Index).Width, _
                      picSortMode(Index).Height
                    
    If CBool(chkSortPreview.Value) Then
        ChangeSort Index
    End If
    
    Temp_Set.SortMode = Index
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub optWinAni_Click(Index As Integer)
    Dim i As Integer
    
    cmbMaxCards.Enabled = True
    
    Select Case Index
    Case Is = 0
        For i = 1 To tblToolbar.Buttons.Count
            tblToolbar.Buttons(i).Enabled = False
        Next i
        
        cmbMaxCards.Enabled = False
    Case Is = 4
        For i = 1 To tblToolbar.Buttons.Count - 1
            tblToolbar.Buttons(i).Enabled = True
        Next i
        
        tblToolbar.Buttons("setting").Enabled = False
    Case Is = 5
        tblToolbar.Buttons(6).Enabled = True
        tblToolbar.Buttons(7).Enabled = True
        tblToolbar.Buttons(8).Enabled = False
        tblToolbar.Buttons(9).Enabled = False
        
        tblToolbar.Buttons("setting").Enabled = False
    Case Else
        For i = 1 To tblToolbar.Buttons.Count - 1
            If i <= 4 Then
                tblToolbar.Buttons(i).Enabled = True
            Else
                tblToolbar.Buttons(i).Enabled = False
            End If
        Next i

        tblToolbar.Buttons("setting").Enabled = True
    End Select
    
    Temp_Set.WinnerSelAni = Index
    CustomAni.Reset = True
    tmrAni.Enabled = True
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub picSortMode_Click(Index As Integer)
    optSortMode(Index).Value = True
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub picSortMode_GotFocus(Index As Integer)
    optSortMode(Index).SetFocus
End Sub

Private Sub sldSpeed_Change()
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
    Temp_Set.Speed = sldSpeed.Value
End Sub

Private Sub tblToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case Is = 1 ' Pause
        tmrAni.Enabled = False
    Case Is = 2 ' Play
        tmrAni.Enabled = True
    Case Is = 3 ' Separator
    Case Is = 4 ' Trail
        CustomAni.ShowTrail = Button.Value
        Temp_Set.ShowTrail = CustomAni.ShowTrail
    Case Is = 5 ' Seperator
    Case 6 To 9
        If optWinAni(4).Value Then
            CustomAni.FallType = Button.Index - 6
            Temp_Set.FallType = CustomAni.FallType
        ElseIf optWinAni(5).Value Then
            CustomAni.WindType = Button.Index - 6
            Temp_Set.WindType = CustomAni.WindType
        End If
    Case Is = 10 ' Separator
    Case Is = 11 ' Settings
        frmSettingAni.Show vbModal, Me
    End Select
        
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub tbsOp_Click()
    Dim Frame As Object
    
    Select Case tbsOp.SelectedItem.Index
    Case Is = 1 ' General
        fraGeneral.ZOrder 0
    Case Is = 2 ' Animation
        fraAni.ZOrder 0
    Case Is = 3 ' Background
        fraBk.ZOrder 0
    Case Is = 4 ' Deck
        fraDeck.ZOrder 0
        crdDeck(frmMain.crdStock(0).Deck).SetFocus
    Case Is = 5 ' Difficulty
        fraDifficulty.ZOrder 0
    Case Is = 6 ' Sort
        optSortMode(Uno.GetSortMode).SetFocus
        fraSort.ZOrder 0
    End Select
End Sub

Private Sub tmrAni_Timer()
    If tbsOp.SelectedItem.Index = 2 Then
        With CustomAni
            If optWinAni(1).Value Then
                .Bounce picViewer, picCanvas
            ElseIf optWinAni(2).Value Then
                .Scatter picViewer, picCanvas
            ElseIf optWinAni(3).Value Then
                .Spin picViewer, picCanvas
            ElseIf optWinAni(4).Value Then
                .Fall picViewer, picCanvas
            ElseIf optWinAni(5).Value Then
                .Wind picViewer, picCanvas
            End If
        End With
    End If
End Sub


Private Sub txtName_Change(Index As Integer)
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub ChangeDeck(ByVal SelDeck As Integer)
    Dim Card As Object
        
    With frmMain
        .crdStock(0).Deck = crdDeck(SelDeck).Deck
            
        For Each Card In .Controls
            If TypeName(Card) = "UnoCard" Then
                If Card.Face = uno_FCDown Then
                    Card.AutoUpdate = False
                    Set Card.Picture = .crdStock(0).Picture
                    Card.Deck = .crdStock(0).Deck
                    Card.AutoUpdate = True
                Else
                    Card.Deck = .crdStock(0).Deck
                End If
            End If
        Next Card
    End With
End Sub


