VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "*\A..\Card\prjUnoCard.vbp"
Begin VB.Form frmMain 
   Caption         =   "Uno Game 2.2 by Aris Buenaventura"
   ClientHeight    =   6330
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9420
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlToolbarGrayed 
      Left            =   1260
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1180
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2334
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":469C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5850
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":612A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A04
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":72DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrCheerLeader 
      Interval        =   100
      Left            =   2400
      Top             =   1920
   End
   Begin MSComctlLib.ImageList imlLogoIcon 
      Left            =   660
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   31
      ImageHeight     =   31
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7BB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":846E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8D24
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":95DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrIconAni 
      Interval        =   400
      Left            =   1860
      Top             =   1920
   End
   Begin VB.Timer tmrWinAni 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2340
      Top             =   1440
   End
   Begin MSComDlg.CommonDialog dlgUno 
      Left            =   2820
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1260
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9E90
            Key             =   "opponent1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A76A
            Key             =   "opponent2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B044
            Key             =   "opponent3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B91E
            Key             =   "deal"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C1F8
            Key             =   "cheat"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CAD2
            Key             =   "open"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D3AC
            Key             =   "tooltip"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DC86
            Key             =   "save"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E560
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EE3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F714
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FFEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":108C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlStatIcons 
      Left            =   660
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":111A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1149C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11796
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11D8A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSortIcons 
      Left            =   60
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Timer tmrValidCard 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1860
      Top             =   1440
   End
   Begin VB.PictureBox picTray 
      Height          =   5415
      Left            =   6480
      ScaleHeight     =   5355
      ScaleWidth      =   2835
      TabIndex        =   3
      Top             =   660
      Width           =   2895
      Begin VB.PictureBox picTrackWin 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00EAFDFD&
         Height          =   1755
         Left            =   60
         ScaleHeight     =   113
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   177
         TabIndex        =   5
         Top             =   3540
         Width           =   2715
         Begin VB.PictureBox picViewer 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   720
            ScaleHeight     =   97
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   73
            TabIndex        =   21
            Top             =   60
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Shape shpRectR 
            BackColor       =   &H00C0FFC0&
            BorderColor     =   &H00008000&
            BorderWidth     =   2
            FillColor       =   &H00008000&
            FillStyle       =   7  'Diagonal Cross
            Height          =   135
            Index           =   1
            Left            =   1920
            Top             =   1080
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Shape shpRectR 
            BackColor       =   &H00C0FFC0&
            BorderColor     =   &H00008000&
            BorderWidth     =   2
            FillColor       =   &H00008000&
            FillStyle       =   7  'Diagonal Cross
            Height          =   135
            Index           =   0
            Left            =   1920
            Top             =   300
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Shape shpRectL 
            BackColor       =   &H00C0FFC0&
            BorderColor     =   &H00008000&
            BorderWidth     =   2
            FillColor       =   &H00008000&
            FillStyle       =   7  'Diagonal Cross
            Height          =   135
            Index           =   1
            Left            =   -1860
            Top             =   1080
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Shape shpRectL 
            BackColor       =   &H00C0FFC0&
            BorderColor     =   &H00008000&
            BorderWidth     =   2
            FillColor       =   &H00008000&
            FillStyle       =   7  'Diagonal Cross
            Height          =   135
            Index           =   0
            Left            =   -1860
            Top             =   300
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Line linLine 
            BorderColor     =   &H0000C000&
            BorderWidth     =   3
            Index           =   1
            Visible         =   0   'False
            X1              =   128
            X2              =   128
            Y1              =   0
            Y2              =   112
         End
         Begin VB.Line linLine 
            BorderColor     =   &H0000C000&
            BorderWidth     =   3
            Index           =   0
            Visible         =   0   'False
            X1              =   28
            X2              =   28
            Y1              =   0
            Y2              =   112
         End
      End
      Begin MSComctlLib.ListView lvwHistory 
         Height          =   2955
         Left            =   60
         TabIndex        =   4
         Top             =   300
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   5212
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         Icons           =   "imlBallIcons"
         SmallIcons      =   "imlStatIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "History"
            Object.Width           =   2540
         EndProperty
         Picture         =   "frmMain.frx":12084
      End
      Begin VB.Label lblHistory 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "HISTORY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   60
         Width           =   2715
      End
      Begin VB.Label lblTrackCard 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "*** W E L C O M E ***"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   3300
         Width           =   2715
      End
   End
   Begin VB.PictureBox picTable 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      DrawStyle       =   5  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   7  'Diagonal Cross
      Height          =   5475
      Left            =   0
      ScaleHeight     =   361
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   429
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   6495
      Begin VB.PictureBox picCanvas 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   60
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   10
         Top             =   1380
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   2160
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin prjUnoCard.UnoCard crdStock 
         Height          =   675
         Index           =   0
         Left            =   2760
         TabIndex        =   15
         Top             =   2580
         Visible         =   0   'False
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   1191
         ShowFocusRect   =   0   'False
         Picture         =   "frmMain.frx":129BE
      End
      Begin prjUnoCard.UnoCard crdWaste 
         Height          =   675
         Left            =   3420
         TabIndex        =   16
         Top             =   2580
         Visible         =   0   'False
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   1191
         ShowFocusRect   =   0   'False
         Picture         =   "frmMain.frx":13DC0
      End
      Begin prjUnoCard.UnoCard crdPlayerOne 
         Height          =   675
         Index           =   0
         Left            =   4080
         TabIndex        =   17
         Top             =   2580
         Visible         =   0   'False
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   1191
         Picture         =   "frmMain.frx":151C2
      End
      Begin prjUnoCard.UnoCard crdPlayerTwo 
         Height          =   675
         Index           =   0
         Left            =   4740
         TabIndex        =   18
         Top             =   2580
         Visible         =   0   'False
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   1191
         ShowFocusRect   =   0   'False
         Picture         =   "frmMain.frx":165C4
      End
      Begin prjUnoCard.UnoCard crdPlayerThree 
         Height          =   675
         Index           =   0
         Left            =   2760
         TabIndex        =   19
         Top             =   3360
         Visible         =   0   'False
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   1191
         ShowFocusRect   =   0   'False
         Picture         =   "frmMain.frx":179C6
      End
      Begin prjUnoCard.UnoCard crdPlayerFour 
         Height          =   675
         Index           =   0
         Left            =   3420
         TabIndex        =   20
         Top             =   3360
         Visible         =   0   'False
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   1191
         ShowFocusRect   =   0   'False
         Picture         =   "frmMain.frx":18DC8
      End
      Begin VB.Label lblPlayerName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "######"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Tag             =   "PlayerName"
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Shape shpCircle 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H0000C000&
         BorderWidth     =   10
         FillColor       =   &H00C0FFC0&
         FillStyle       =   7  'Diagonal Cross
         Height          =   630
         Left            =   4020
         Shape           =   3  'Circle
         Top             =   60
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Shape shpWaste 
         BackColor       =   &H00808080&
         BorderColor     =   &H00000000&
         FillColor       =   &H00404040&
         FillStyle       =   7  'Diagonal Cross
         Height          =   675
         Left            =   3420
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape shpStock 
         BackColor       =   &H00808080&
         BorderColor     =   &H00000000&
         FillColor       =   &H00404040&
         FillStyle       =   7  'Diagonal Cross
         Height          =   675
         Left            =   2880
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin MSComctlLib.Toolbar tblToolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarGrayed"
      HotImageList    =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deal"
            Object.ToolTipText     =   "Deal"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "opponent"
            Object.ToolTipText     =   "1 Opponent"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "1 Opponent"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "2 Opponents"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "3 Opponents"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cheat"
            Object.ToolTipText     =   "Cheat"
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Hint"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tooltip"
            Object.ToolTipText     =   "Show Tooltip"
            ImageIndex      =   7
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sound"
            Object.ToolTipText     =   "Sound On"
            ImageIndex      =   10
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   13
         EndProperty
      EndProperty
      Begin VB.Frame fraSort 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   6600
         TabIndex        =   11
         Top             =   -60
         Width           =   3195
         Begin VB.CheckBox chkAutoSort 
            Height          =   435
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Auto Sort"
            Top             =   120
            Value           =   1  'Checked
            Width           =   375
         End
         Begin MSComctlLib.ImageCombo imcSortMode 
            Height          =   330
            Left            =   420
            TabIndex        =   13
            Top             =   120
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
         End
         Begin VB.Label lblSort 
            AutoSize        =   -1  'True
            Caption         =   "Sort : "
            Height          =   195
            Left            =   0
            TabIndex        =   14
            Top             =   240
            Width           =   420
         End
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6075
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8414
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5080
            MinWidth        =   5080
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "1/8/2004"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameDeal 
         Caption         =   "&Deal"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuGameBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuGameSave 
         Caption         =   "Sa&ve"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuGameBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuGameBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameCheat 
         Caption         =   "&Cheat"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuGameShowTooltip 
         Caption         =   "Show &Tooltip"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuGameShowDemo 
         Caption         =   "Show &Demo"
      End
      Begin VB.Menu mnuGameSound 
         Caption         =   "Sound"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuGameBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuGameHistoryTrackCard 
         Caption         =   "History and Track Card"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewCheerLeader 
         Caption         =   "Cheer Leader"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rotation          As Integer
Dim TotalStockLeft    As Integer

Dim CurrentPlayer     As Object
Dim PreviousPlayer    As Object
Dim LastPlayer        As Object

Dim IsGameExit        As Boolean
Dim ResetGame         As Boolean
Dim IsLinearAniOn     As Boolean
Dim IsDeal            As Boolean
Dim SelCardAni        As Object
Dim IsDone            As Boolean
Dim IsLogo            As Boolean

Dim HelpWinHwnd       As Long
Dim OldWinnerAni      As Integer
Dim WinSelCardAni     As New Collection
Dim OldSettingMaxCard As Integer

Dim cl_1              As Integer
Dim cl_2              As Integer

Dim cookie            As Long

Public WinnerAni      As New clsAni
Public LA             As New clsLinearAni
Attribute LA.VB_VarHelpID = -1
    
Private Sub crdPlayerFour_MouseOut(Index As Integer)
    If crdPlayerFour(Index).ToolTipText <> "" Then
        crdPlayerFour(Index).ToolTipText = ""
    End If
End Sub

Private Sub crdPlayerFour_MouseOver(Index As Integer)
    TrackCard crdPlayerFour(Index)
End Sub

Private Sub crdPlayerOne_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeySpace) Then
        crdPlayerOne_MouseUp Index, vbLeftButton, 0, 0, 0
    End If
End Sub

Private Sub crdPlayerOne_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set crdPlayerOne(Index).MouseIcon = LoadResPicture(102, vbResCursor)
End Sub

Private Sub crdPlayerOne_MouseOut(Index As Integer)
    If crdPlayerOne(Index).ToolTipText <> "" Then
        crdPlayerOne(Index).ToolTipText = ""
    End If
    
    picTrackWin.MousePointer = vbDefault
    crdPlayerOne(Index).MousePointer = vbDefault
End Sub

Private Sub crdPlayerOne_MouseOver(Index As Integer)
    If Uno.IsMoveValid(crdPlayerOne, Index, _
                       crdWaste.Rank, crdWaste.Suit) Then
        
        crdPlayerOne(Index).MousePointer = vbCustom
        Set crdPlayerOne(Index).MouseIcon = _
            LoadResPicture(101, vbResCursor)
        
        linLine(0).X1 = -linLine(0).BorderWidth
        linLine(0).X2 = -linLine(0).BorderWidth
        linLine(1).X1 = picTrackWin.ScaleWidth + linLine(0).BorderWidth
        linLine(1).X2 = picTrackWin.ScaleWidth + linLine(0).BorderWidth
        
        shpRectL(0).Left = -shpRectL(0).Width - shpRectL(0).BorderWidth + 1
        shpRectL(1).Left = -shpRectL(1).Width - shpRectL(1).BorderWidth + 1
        shpRectR(0).Left = picTrackWin.ScaleWidth + shpRectR(0).BorderWidth + 1
        shpRectR(1).Left = picTrackWin.ScaleWidth + shpRectR(1).BorderWidth + 1
        
        linLine(0).Visible = True
        linLine(1).Visible = True
        
        shpRectL(0).Visible = True
        shpRectL(1).Visible = True
        shpRectR(0).Visible = True
        shpRectR(1).Visible = True
        
        tmrValidCard.Enabled = True
    Else
        linLine(0).Visible = False
        linLine(1).Visible = False
        
        shpRectL(0).Visible = False
        shpRectL(1).Visible = False
        shpRectR(0).Visible = False
        shpRectR(1).Visible = False
        
        tmrValidCard.Enabled = False
    End If
    
    TrackCard crdPlayerOne(Index)
End Sub

Private Sub crdPlayerOne_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then
        If CurrentPlayer(0).Name = "crdPlayerOne" Then
            If crdPlayerOne(Index).MousePointer = vbCustom Then
                Set crdPlayerOne(Index).MouseIcon = _
                    LoadResPicture(101, vbResCursor)
            End If
            
            PlayerMove Index
        End If
    Else
        On Error Resume Next
        crdPlayerOne(Index).Rank = CInt(InputBox("Rank", "")) ' ****
        crdPlayerOne(Index).Suit = CInt(InputBox("Suit", "")) ' ****
    End If
End Sub

Private Sub crdPlayerThree_MouseOut(Index As Integer)
    If crdPlayerThree(Index).ToolTipText <> "" Then
        crdPlayerThree(Index).ToolTipText = ""
    End If
End Sub

Private Sub crdPlayerThree_MouseOver(Index As Integer)
    TrackCard crdPlayerThree(Index)
End Sub

Private Sub crdPlayerTwo_MouseOut(Index As Integer)
    If crdPlayerTwo(Index).ToolTipText <> "" Then
        crdPlayerTwo(Index).ToolTipText = ""
    End If
End Sub

Private Sub crdPlayerTwo_MouseOver(Index As Integer)
    TrackCard crdPlayerTwo(Index)
End Sub

Private Sub crdStock_Click(Index As Integer)
    If Uno.Stock.Count > 0 Then
        If Index = 0 Then
            If crdStock(2).Visible Or crdStock(1).Visible Then
                Exit Sub
            End If
        ElseIf Index = 1 Then
            If crdStock(2).Visible Then
                Exit Sub
            End If
        End If
        
        Uno.Pick CurrentPlayer
        Call UpdateStockPile
                
        If Uno.Stock.Count > 0 Then
            crdStock(GetCurrentStock).Rank = Uno.Stock(Uno.Stock.Count).Rank
            crdStock(GetCurrentStock).Suit = Uno.Stock(Uno.Stock.Count).Suit
        Else
            crdStock(0).Visible = False
        End If
        
        CardAni CurrentPlayer(CurrentPlayer.Count - 1), 0
        History CurrentPlayer(0).Name, "Draws a card", CurrentPlayer(CurrentPlayer.Count - 1).Rank, _
                                                       CurrentPlayer(CurrentPlayer.Count - 1).Suit
        AlignPlayerCards CurrentPlayer
        
        If CBool(chkAutoSort.Value) Then
            Uno.SortCards CurrentPlayer
        End If
        
        Set LastPlayer = CurrentPlayer
        Set CurrentPlayer = GetNextPlayer(CurrentPlayer, Rotation)
        
        Call ShowActivePlayer
        TrackCard crdStock(Index)
    End If
End Sub

Private Sub crdStock_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set crdStock(GetCurrentStock).MouseIcon = LoadResPicture(102, vbResCursor)
End Sub

Private Sub crdStock_MouseOut(Index As Integer)
    Dim CurStock As Integer
    
    If Index = 0 Then
        If crdStock(2).Visible Or crdStock(1).Visible Then
            Exit Sub
        End If
    ElseIf Index = 1 Then
        If crdStock(2).Visible Then
            Exit Sub
        End If
    End If

    If crdStock(Index).ToolTipText <> "" Then
        crdStock(Index).ToolTipText = ""
    End If
End Sub

Private Sub crdStock_MouseOver(Index As Integer)
    If Index = 0 Then
        If crdStock(2).Visible Or crdStock(1).Visible Then
            Exit Sub
        End If
    ElseIf Index = 1 Then
        If crdStock(2).Visible Then
            Exit Sub
        End If
    End If

    TrackCard crdStock(Index)
End Sub

Private Sub crdStock_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set crdStock(GetCurrentStock).MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub crdWaste_MouseOut()
    If crdWaste.ToolTipText <> "" Then
        crdWaste.ToolTipText = ""
    End If
End Sub

Private Sub crdWaste_MouseOver()
    TrackCard crdWaste
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        mnuGameShowDemo.Checked = False
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim j As Integer
    
    HtmlHelp 0, vbNullString, HH_INITIALIZE, cookie
    
    ' set track card window
    picViewer.Width = CardW + 2
    picViewer.Height = CardH + 2
    picViewer.BackColor = RGB(253, 206, 102)
    picViewer.Move (picTrackWin.ScaleWidth - picViewer.ScaleWidth) / 2, _
                   (picTrackWin.ScaleHeight - picViewer.ScaleHeight) / 2
    picTrackWin.BackColor = RGB(253, 206, 102)
    Set picTrackWin.MouseIcon = LoadResPicture(101, vbResCursor)
    
    ' set stock pile
    For i = 1 To 2
        Load crdStock(i)
        crdStock(i).Left = crdStock(i - 1).Left + 20
        crdStock(i).Rank = Int(15 * Rnd)
        crdStock(i).Suit = (i - 1) Mod 4
        crdStock(i).ZOrder 0
    Next i
    
    ' set all cards
    Dim Card As Object
    
    For Each Card In Me.Controls
        If TypeName(Card) = "UnoCard" Then
            Card.Move Card.Left, Card.Top, CardW, CardH
        End If
    Next Card
    
    ' set combo box (sort)
    For i = 0 To 4
        picCanvas.Cls
        picCanvas.Width = 120
        picCanvas.Height = 29
        
        For j = 301 To 305
        
            picCanvas.PaintPicture LoadResPicture((j Mod 301 + i) Mod 5 + 301, _
                                                   vbResBitmap), _
                                   24 * (j Mod 301) + 3, 1
        Next j
        
        Set picCanvas.Picture = picCanvas.Image
        imlSortIcons.ListImages.Add , , picCanvas.Picture
    Next i
    
    Set imcSortMode.ImageList = imlSortIcons
    Set chkAutoSort.Picture = LoadResPicture(401, vbResBitmap)
    
    For i = 1 To 5
        imcSortMode.ComboItems.Add , , , i
    Next i
    
    For i = 0 To 3
        If i <> 0 Then
            Load lblPlayerName(i)
        End If
        
        lblPlayerName(i).Caption = Setting.PlayerName(i)
    Next i
    
    lvwHistory.ColumnHeaders(1).Width = lvwHistory.Width
    
    ' set settings
    On Error Resume Next
    
    Call OpenSettings
    If Setting.BkFileLoc <> "" Then
        If Dir$(Setting.BkFileLoc) Then
            Set Setting.BkPicture = LoadPicture(Setting.BkFileLoc)
        End If
    End If
    
    For Each Card In Me.Controls
        If TypeName(Card) = "UnoCard" Then
            Card.Deck = Setting.Deck
        End If
    Next Card
    
    lblPlayerName(0).Caption = Setting.PlayerName(0)
    lblPlayerName(1).Caption = Setting.PlayerName(2)
    lblPlayerName(2).Caption = Setting.PlayerName(1)
    lblPlayerName(3).Caption = Setting.PlayerName(3)
    
    picTable.BackColor = Setting.BkColor
    tblToolbar.Buttons("opponent").Image = Setting.Opponents
    imcSortMode.ComboItems(Setting.SortMode + 1).Selected = True
    Uno.SetSortMode = Setting.SortMode
    LA.Speed = Setting.Speed
                
    OldSettingMaxCard = Setting.MaxCard
    WinnerAni.FallType = Setting.FallType
    WinnerAni.MaxCards = Setting.MaxCard
    WinnerAni.WindType = Setting.WindType
    
    cl_1 = 0: cl_2 = 4
    
    IsLogo = True
    IsGameExit = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HtmlHelp Me.hwnd, "", HH_CLOSE_ALL, 0&
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Dim ToolBarH As Integer
    Dim TableW   As Integer

    ToolBarH = IIf(mnuViewToolbar.Checked, tblToolbar.Height, 0)
    TableW = IIf(mnuGameHistoryTrackCard.Checked, picTray.ScaleWidth, 0)
    
    If Me.WindowState <> vbMinimized Then
        LockWindowUpdate Me.hwnd
        picTray.Move Me.ScaleWidth - picTray.ScaleWidth - 75, _
                     ToolBarH, picTray.Width, _
                     Me.ScaleHeight - sbStatusBar.Height - ToolBarH
        picTable.Move 0, ToolBarH, _
                      Me.ScaleWidth - TableW, _
                      Me.ScaleHeight - sbStatusBar.Height - ToolBarH
        lvwHistory.Height = picTray.Height - picTrackWin.Height - lblTrackCard.Height - _
                            lblHistory.Height - 360
        lblTrackCard.Top = lvwHistory.Top + lvwHistory.Height + 45
        picTrackWin.Top = lvwHistory.Top + lvwHistory.Height + lblHistory.Height + _
                          lblTrackCard.Height - 75
        
        LockWindowUpdate 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Are you sure you want to exit?", vbQuestion Or vbOKCancel, "Exit") = vbOK Then
        HtmlHelp 0, vbNullString, HH_UNINITIALIZE, cookie
        
        IsGameExit = True
        LA.StopAni = True
        tmrWinAni.Enabled = False
        
        Call SaveSettings
        End
    Else
        Cancel = True
    End If
End Sub

Private Sub imcSortMode_Click()
    ChangeSort imcSortMode.SelectedItem.Index - 1
End Sub

Private Sub mnuGameCheat_Click()
    Dim Card As Object
    
    mnuGameCheat.Checked = Not mnuGameCheat.Checked
    tblToolbar.Buttons("cheat").Value = _
        IIf(mnuGameCheat.Checked, tbrPressed, tbrUnpressed)
    
    For Each Card In Me.Controls
        If TypeName(Card) = "UnoCard" Then
            If (Card.Name <> "crdPlayerOne") And _
               (Card.Name <> "crdWaste") Then
                Card.Face = IIf(mnuGameCheat.Checked, _
                                uno_FCUp, uno_FCDown)
            End If
        End If
    Next Card
    
    Call SortPlayerCards
End Sub

Private Sub mnuGameDeal_Click()
    Dim i         As Integer
    Dim j         As Integer
    Dim oTemp     As Object
    Dim Player(4) As Object
    Dim SW        As Integer
    Dim SH        As Integer
    
    tmrWinAni.Enabled = False
    
    ResetGame = True
    Me.Enabled = False
    picTable.Enabled = False
    
    picTable.Cls
    Set picTable.Picture = Nothing
     
    If Setting.BkPicture.handle <> 0 Then
        Dim rcRect As RECT
        
        GetClientRect picTable.hwnd, rcRect
        Background picTable.hdc, Setting.BkPicture, _
                   rcRect.Right, rcRect.Bottom
        RefreshWindow picTable.hwnd
    End If
     
    If Not mnuGameCheat.Checked Then
        crdStock(0).Face = uno_FCDown
        crdStock(1).Face = uno_FCDown
        crdStock(2).Face = uno_FCDown
        crdPlayerTwo(0).Face = uno_FCDown
        crdPlayerThree(0).Face = uno_FCDown
        crdPlayerFour(0).Face = uno_FCDown
    End If
    
    SW = picTable.ScaleWidth
    SH = picTable.ScaleHeight
    
    For i = crdStock.LBound To crdStock.UBound
        crdStock(i).Move (SW - CardW) / 2 - CardW * 0.6 + i * 2 - picColor.ScaleWidth * 0.5, _
                         (SH - CardH) / 2 + i * 2, CardW, CardH
        crdStock(i).ZOrder 0
    Next i
    
    For Each oTemp In Me.Controls
        If TypeName(oTemp) = "UnoCard" Then
            If oTemp.Visible And (oTemp.Name <> "crdStock") Then
                Dim OldSpeed As Integer
                
                If LA.Speed <> 0 Then
                    OldSpeed = LA.Speed
                    LA.Speed = LA.Speed + 1
                End If
                
                Do While Not LA.Linear(oTemp, oTemp.Left, _
                                              oTemp.Top, _
                                              crdStock(GetCurrentStock).Left, _
                                              crdStock(GetCurrentStock).Top)
                Loop
                
                If LA.Speed <> 0 Then LA.Speed = OldSpeed
            End If
        End If
    Next oTemp
    
    picColor.Cls
    picColor.Visible = False
    shpStock.Visible = True
    shpWaste.Visible = True
    shpCircle.Visible = True
    Setting.Opponents = tblToolbar.Buttons(5).Image
    
    lvwHistory.ListItems.Clear
    sbStatusBar.Panels(1).Text = ""
    Set sbStatusBar.Panels(1).Picture = Nothing
                
    For Each oTemp In Me.Controls
        If TypeName(oTemp) = "UnoCard" Then
            If oTemp.Name <> "crdStock" Then
                oTemp.Visible = False
            End If
            
            oTemp.Enabled = True
        End If
        
        If oTemp.Tag = "PlayerName" Then
            oTemp.Visible = False
        End If
    Next oTemp
    
    Select Case Setting.Difficulty
    Case Is = 0 ' Easy
        sbStatusBar.Panels(2).Text = "Level : EASY"
    Case Is = 1 ' Normal
        sbStatusBar.Panels(2).Text = "Level : NORMAL"
    Case Is = 2 ' Difficult
        sbStatusBar.Panels(2).Text = "Level : DIFFICULT"
    End Select
    
    Call Uno.Shuffle
    
    For i = LBound(Player()) To UBound(Player())
        Select Case i
        Case Is = 0 ' Player One
            Set Player(0) = crdPlayerOne
        Case Is = 1 ' Player Two
            Set Player(1) = crdPlayerTwo
        Case Is = 2 ' Player Three
            Set Player(2) = crdPlayerThree
        Case Is = 3 ' Player Four
            Set Player(3) = crdPlayerFour
        End Select
    Next i
    
    For i = 0 To 3
        For j = 1 To Player(i).Count - 1
            Unload Player(i)(j)
        Next j
    Next i
    
    For i = 0 To crdStock.Count - 1
        crdStock(i).Visible = True
    Next i
    
    For i = 0 To lblPlayerName.Count - 1
        If i <= Setting.Opponents Then
            lblPlayerName(i).Move -lblPlayerName(i).Width, _
                                  -lblPlayerName(i).Height
            lblPlayerName(i).Visible = True
        Else
            lblPlayerName(i).Visible = False
        End If
    Next i
    
    For i = lblPlayerName.LBound To lblPlayerName.UBound
        lblPlayerName(0).ForeColor = vbWhite
    Next i
    
    For i = 0 To Setting.Opponents
        Player(i)(0).Move crdStock(GetCurrentStock).Left, _
                          crdStock(GetCurrentStock).Top
        Player(i)(0).Data = 0
        Player(i)(0).Rank = Uno.Stock(Uno.Stock.Count).Rank
        Player(i)(0).Suit = Uno.Stock(Uno.Stock.Count).Suit
        Player(i)(0).Visible = True
        Uno.Stock.Remove Uno.Stock.Count
        
        For j = 1 To 6
            Uno.Pick Player(i)
        Next j
    Next i
    
    Call SortPlayerCards
    crdWaste.Rank = CInt(Rnd * uno_RCNine)
    crdWaste.Suit = CInt(Rnd * uno_SCYellow)
    crdStock(GetCurrentStock).Rank = Player(0)(0).Rank
    crdStock(GetCurrentStock).Suit = Player(0)(0).Suit
    
    For i = crdStock.LBound To crdStock.UBound
        If i < 2 Then
            crdStock(i).MousePointer = vbDefault
        Else
            crdStock(i).MousePointer = vbCustom
            Set crdStock(i).MouseIcon = LoadResPicture(101, vbResCursor)
        End If
    Next i
    
    Select Case Setting.Opponents
    Case Is = 1
        lblPlayerName(0).Caption = Setting.PlayerName(0)
        lblPlayerName(1).Caption = Setting.PlayerName(1)
    Case Is = 2
        lblPlayerName(0).Caption = Setting.PlayerName(0)
        lblPlayerName(1).Caption = Setting.PlayerName(2)
        lblPlayerName(2).Caption = Setting.PlayerName(1)
    Case Is = 3
        lblPlayerName(0).Caption = Setting.PlayerName(0)
        lblPlayerName(1).Caption = Setting.PlayerName(2)
        lblPlayerName(2).Caption = Setting.PlayerName(1)
        lblPlayerName(3).Caption = Setting.PlayerName(3)
    End Select
        
    Uno.SetLevelMode = Setting.Difficulty
    
    Setting.MaxCard = OldSettingMaxCard
    WinnerAni.MaxCards = OldSettingMaxCard
    WinnerAni.FallType = Setting.FallType
    WinnerAni.ShowTrail = Setting.ShowTrail
    
    Rotation = 0
    IsDeal = True
    IsDone = False
    IsLogo = False
    
    TotalStockLeft = Uno.Stock.Count
    
    Call picTable_Resize
    crdStock(GetCurrentStock).ZOrder 0
    picColor.Visible = True
    Call ShowActiveColor
    
    Dim SelectedPlayer As Integer
    
    SelectedPlayer = Random_Number(0, Setting.Opponents)
    Select Case SelectedPlayer
    Case Is = 0
        Set CurrentPlayer = crdPlayerOne
    Case Is = 1
        Set CurrentPlayer = crdPlayerTwo
    Case Is = 2
        Set CurrentPlayer = crdPlayerThree
    Case Is = 3
        Set CurrentPlayer = crdPlayerFour
    End Select
    
    Set PreviousPlayer = CurrentPlayer
    Set LastPlayer = CurrentPlayer
    
    If CBool(chkAutoSort.Value) Then
        Uno.SortCards CurrentPlayer
    End If
    
    If Setting.Opponents = 1 Then
        If SelectedPlayer = 0 Then
            SelectedPlayer = 0
        Else
            SelectedPlayer = 1
        End If
    ElseIf Setting.Opponents = 2 Then
        If SelectedPlayer = 0 Then
            SelectedPlayer = 0
        ElseIf SelectedPlayer = 1 Then
            SelectedPlayer = 2
        Else
            SelectedPlayer = 1
        End If
    Else
        If SelectedPlayer = 0 Then
            SelectedPlayer = 0
        ElseIf SelectedPlayer = 1 Then
            SelectedPlayer = 2
        ElseIf SelectedPlayer = 2 Then
            SelectedPlayer = 1
        Else
            SelectedPlayer = 3
        End If
    End If

    lblPlayerName(SelectedPlayer).ForeColor = vbYellow
    MsgBox lblPlayerName(SelectedPlayer).Caption & " starts!", vbOKOnly
    
    IsDeal = False
    picTable.Enabled = True
    Me.Enabled = True
    
    ResetGame = False
    Call GameStart
End Sub

Private Sub mnuGameExit_Click()
    Unload Me
End Sub

Private Sub mnuGameHistoryTrackCard_Click()
    mnuGameHistoryTrackCard.Checked = Not mnuGameHistoryTrackCard.Checked
    picTray.Visible = mnuGameHistoryTrackCard.Checked
    Call Form_Resize
End Sub

Private Sub mnuGameOpen_Click()
    On Error GoTo OpenErr
    
    With dlgUno
        .Filter = "Uno Game File (*.uno) | *.uno; |All Files (*.*) | *.*"
        .FilterIndex = 1
        .Filename = ""
        .InitDir = App.Path & "\Save"
        .ShowOpen
        
        If .Filename <> "" Then
            Dim InFile As Long
            Dim sData  As String
            
            InFile = FreeFile
            Open .Filename For Input As InFile
                Input #InFile, sData ' Signature
                
                If sData <> Signature Then
                    If InFile <> 0 Then Close InFile
                    MsgBox "File format error!", vbCritical Or vbOKOnly, "Uno Game 2.0"
                    Exit Sub
                End If
                
                Input #InFile, sData ' Version
                
                If sData <> Version Then
                    If InFile <> 0 Then Close InFile
                    MsgBox "Invalid version!", vbCritical Or vbOKOnly, "Uno Game 2.0"
                    Exit Sub
                End If
                
                Dim i         As Integer
                Dim j         As Integer
                Dim Status    As String
                Dim arrData() As String
                Dim bVal      As Boolean
                Dim Player(4) As Object
                Dim oTemp     As Object
            
                ResetGame = False
                Set Uno.Stock = Nothing
                Set Uno.Waste = Nothing
                lvwHistory.ListItems.Clear
    
                For i = LBound(Player()) To UBound(Player())
                    Select Case i
                    Case Is = 0 ' Player One
                        Set Player(0) = crdPlayerOne
                    Case Is = 1 ' Player Two
                        Set Player(1) = crdPlayerTwo
                    Case Is = 2 ' Player Three
                        Set Player(2) = crdPlayerThree
                    Case Is = 3 ' Player Four
                        Set Player(3) = crdPlayerFour
                    End Select
                Next i
        
                For i = 0 To Setting.Opponents
                    For j = 1 To Player(i).Count - 1
                        Unload Player(i)(j)
                    Next j
                Next i
                
                For i = lblPlayerName.LBound To lblPlayerName.UBound
                    lblPlayerName(i).Visible = False
                Next i
                
                Input #InFile, Setting.Opponents
                Input #InFile, sData ' Currentplayer
                
                Select Case sData
                Case Is = "crdPlayerOne"
                    Set CurrentPlayer = crdPlayerOne
                Case Is = "crdPlayerTwo"
                    Set CurrentPlayer = crdPlayerTwo
                Case Is = "crdPlayerThree"
                    Set CurrentPlayer = crdPlayerThree
                Case Is = "crdPlayerFour"
                    Set CurrentPlayer = crdPlayerFour
                End Select
                
                Do While Not EOF(InFile)
                    Input #InFile, sData
                    If sData = "[Stock]" Then
                        Status = 0: bVal = True
                    ElseIf sData = "[Waste]" Then
                        Status = 1: bVal = True
                    ElseIf sData = "[Player]" Then
                        Status = 2: bVal = True
                    ElseIf sData = "[Card]" Then
                        Status = 3: bVal = True
                    ElseIf sData = "[History]" Then
                        Status = 4: bVal = True
                    End If
                    
                    If bVal Then
                        bVal = False
                    Else
                        arrData = Split(sData, ",")
                        
                        Select Case Status
                        Case Is = 0 ' load stock
                            Set oTemp = New clsCardInfo
                            oTemp.Data = arrData(0)
                            oTemp.Rank = CInt(arrData(1))
                            oTemp.Suit = CInt(arrData(2))
                            oTemp.Tag = arrData(3)
                            Uno.Stock.Add oTemp
                        Case Is = 1 ' load waste
                            Set oTemp = New clsCardInfo
                            oTemp.Data = arrData(0)
                            oTemp.Rank = CInt(arrData(1))
                            oTemp.Suit = CInt(arrData(2))
                            oTemp.Tag = arrData(3)
                            Uno.Waste.Add oTemp
                        Case Is = 2 ' load players name
                            lblPlayerName(CInt(arrData(1))).Caption = arrData(0)
                            lblPlayerName(CInt(arrData(1))).Visible = CBool(arrData(2))
                            lblPlayerName(CInt(arrData(1))).ForeColor = CLng(arrData(3))
                        Case Is = 3 ' load cards
                            If arrData(0) <> "crdWaste" Then
                                If Val(arrData(1)) <> 0 Then
                                    Select Case arrData(0)
                                    Case Is = "crdPlayerOne"
                                        Load crdPlayerOne(Val(arrData(1)))
                                        Set oTemp = crdPlayerOne(Val(arrData(1)))
                                    Case Is = "crdPlayerTwo"
                                        Load crdPlayerTwo(Val(arrData(1)))
                                        Set oTemp = crdPlayerTwo(Val(arrData(1)))
                                    Case Is = "crdPlayerThree"
                                        Load crdPlayerThree(Val(arrData(1)))
                                        Set oTemp = crdPlayerThree(Val(arrData(1)))
                                    Case Is = "crdPlayerFour"
                                        Load crdPlayerFour(Val(arrData(1)))
                                        Set oTemp = crdPlayerFour(Val(arrData(1)))
                                    Case Is = "crdStock"
                                        Set oTemp = crdStock(Val(arrData(1)))
                                    End Select
                                Else
                                    Select Case arrData(0)
                                    Case Is = "crdPlayerOne"
                                        Set oTemp = crdPlayerOne(0)
                                    Case Is = "crdPlayerTwo"
                                        Set oTemp = crdPlayerTwo(0)
                                    Case Is = "crdPlayerThree"
                                        Set oTemp = crdPlayerThree(0)
                                    Case Is = "crdPlayerFour"
                                        Set oTemp = crdPlayerFour(0)
                                    Case Is = "crdStock"
                                        Set oTemp = crdStock(0)
                                    End Select
                                End If
                                
                                oTemp.Rank = CInt(arrData(2))
                                oTemp.Suit = CInt(arrData(3))
                                oTemp.Visible = CBool(arrData(4))
                            Else
                                crdWaste.Rank = CInt(arrData(1))
                                crdWaste.Suit = CInt(arrData(2))
                                crdWaste.Visible = CBool(arrData(3))
                            End If
                        Case Is = 4 ' list history
                            lvwHistory.ListItems.Add , , arrData(0), , CInt(arrData(1))
                            lvwHistory.ListItems(lvwHistory.ListItems.Count).ForeColor = CLng(arrData(2))
                            lvwHistory.ListItems(lvwHistory.ListItems.Count).ToolTipText = arrData(1)
                        End Select
                    End If
                Loop
            Close #InFile
            
            Call picTable_Resize
            Call ShowActiveColor
            
            For i = crdStock.LBound To crdStock.UBound
                If i < 2 Then
                    crdStock(i).MousePointer = vbDefault
                Else
                    crdStock(i).MousePointer = vbCustom
                    Set crdStock(i).MouseIcon = LoadResPicture(101, vbResCursor)
                End If
            Next i
            
            If lvwHistory.ListItems.Count > 0 Then
                sbStatusBar.Panels(1).Text = _
                    lvwHistory.ListItems(lvwHistory.ListItems.Count).Text
                sbStatusBar.Panels(1).Picture = _
                    imlStatIcons.ListImages(lvwHistory.ListItems(lvwHistory.ListItems.Count).SmallIcon).Picture
            Else
                sbStatusBar.Panels(1).Text = ""
                Set sbStatusBar.Panels(1).Picture = Nothing
            End If
            
            tblToolbar.Buttons(5).Image = Setting.Opponents
            
            Call GameStart
            ResetGame = True
        End If
    End With
    Exit Sub
    
OpenErr:
    If InFile <> 0 Then Close InFile
    MsgBox Err.Description, vbOKOnly Or vbCritical, "Error"
End Sub

Private Sub mnuGameOptions_Click()
    Setting.MaxCard = OldSettingMaxCard
    WinnerAni.MaxCards = OldSettingMaxCard
    
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuGameSave_Click()
    On Error GoTo SaveErr
    
    Dim i      As Integer
    Dim Card   As Object
    Dim sData  As String
    Dim InFile As Long
        
    With dlgUno
        .Filter = "Uno Game File (*.uno) | *.uno"
        .FilterIndex = 1
        .Filename = ""
        .InitDir = App.Path & "\Save"
        .Flags = cdlOFNOverwritePrompt
        .ShowSave
        
        If .Filename = "" Then Exit Sub
    
        InFile = FreeFile
        Open .Filename For Output As InFile
            Write #InFile, Signature
            Write #InFile, Version
            Write #InFile, Setting.Opponents
            Write #InFile, CurrentPlayer(0).Name
            
            Write #InFile, "[Player]"
            For i = 0 To Setting.Opponents
                Write #InFile, lblPlayerName(i).Caption & "," & _
                               lblPlayerName(i).Index & "," & _
                               lblPlayerName(i).Visible & "," & _
                               lblPlayerName(i).ForeColor
            Next i
            
            Write #InFile, "[Stock]"
            For i = 1 To Uno.Stock.Count
                Write #InFile, Uno.Stock(i).Data & "," & _
                               Uno.Stock(i).Rank & "," & _
                               Uno.Stock(i).Suit & "," & _
                               Uno.Stock(i).Tag
            Next i
            
            Write #InFile, "[Waste]"
            For i = 1 To Uno.Waste.Count
                Write #InFile, Uno.Waste(i).Data & "," & _
                               Uno.Waste(i).Rank & "," & _
                               Uno.Waste(i).Suit & "," & _
                               Uno.Waste(i).Tag
            Next i
            
            Write #InFile, "[Card]"
            For Each Card In Me.Controls
                If TypeName(Card) = "UnoCard" Then
                    ' save all cards
                    If Card.Name <> "crdWaste" Then
                        sData = Card.Name & "," & _
                                Card.Index & "," & _
                                Card.Rank & "," & _
                                Card.Suit & "," & _
                                Card.Visible
                    Else
                        sData = Card.Name & "," & _
                                Card.Rank & "," & _
                                Card.Suit & "," & _
                                Card.Visible
                    End If
                    
                    Write #InFile, sData
                End If
            Next Card
            
            Write #InFile, "[History]"
            For i = 1 To lvwHistory.ListItems.Count
                Write #InFile, lvwHistory.ListItems(i).Text & "," & _
                               lvwHistory.ListItems(i).SmallIcon & "," & _
                               lvwHistory.ListItems(i).ForeColor
            Next i
        Close InFile
    End With
    Exit Sub
    
SaveErr:
    If InFile <> 0 Then Close InFile
    MsgBox Err.Description, vbOKOnly Or vbCritical, "Error"
End Sub

Private Sub mnuGameShowDemo_Click()
    mnuGameShowDemo.Checked = Not mnuGameShowDemo.Checked
    If mnuGameShowDemo.Checked Then
        MsgBox "Press escape to cancel demo", vbInformation, "Demo"
    End If
End Sub

Private Sub mnuGameShowTooltip_Click()
    mnuGameShowTooltip.Checked = Not mnuGameShowTooltip.Checked
    tblToolbar.Buttons("tooltip").Value = _
        IIf(mnuGameShowTooltip.Checked, tbrPressed, tbrUnpressed)
End Sub

Private Sub mnuGameSound_Click()
    mnuGameSound.Checked = Not mnuGameSound.Checked
    tblToolbar.Buttons("sound").Value = IIf(mnuGameSound.Checked, tbrPressed, tbrUnpressed)
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub AlignPlayerCards(ByRef PlayerCards As Object)
    Dim i    As Integer
    Dim SW   As Integer
    Dim SH   As Integer
    Dim cx   As Integer
    Dim cy   As Integer
    Dim NC   As Integer
    Dim temp As Integer
    
    SW = picTable.ScaleWidth
    SH = picTable.ScaleHeight
    NC = PlayerCards.Count - 1
       
    For i = 0 To PlayerCards.Count - 1
        If Setting.Opponents = 1 Then
            Select Case PlayerCards(i).Name
            Case Is = "crdPlayerOne"
                cx = (SW - (CardW + CardW * 0.3 * NC)) / 2 + i * CardW * 0.3
                cy = SH - CardH - CardH * 0.05
            Case Is = "crdPlayerTwo"
                cx = (SW - (CardW + CardW * 0.3 * NC)) / 2 + i * CardW * 0.3
                cy = CardH * 0.05
            End Select
        ElseIf Setting.Opponents = 2 Then
            Select Case PlayerCards(i).Name
            Case Is = "crdPlayerOne"
                cx = (SW - (CardW + CardW * 0.3 * NC)) / 2 + i * CardW * 0.3
                cy = SH - CardH - CardH * 0.05
            Case Is = "crdPlayerTwo"
                cx = i * CardW * 0.14
                cy = (SH - CardH) / 2
            Case Is = "crdPlayerThree"
                cx = (SW - (CardW + CardW * 0.3 * NC)) / 2 + i * CardW * 0.3
                cy = CardH * 0.05
            End Select
        Else
            Select Case PlayerCards(0).Name
            Case Is = "crdPlayerOne"
                cx = (SW - (CardW + CardW * 0.3 * NC)) / 2 + i * CardW * 0.3
                cy = SH - CardH - CardH * 0.05
            Case Is = "crdPlayerTwo"
                cx = i * CardW * 0.14
                cy = (SH - CardH) / 2
            Case Is = "crdPlayerThree"
                cx = (SW - (CardW + CardW * 0.3 * NC)) / 2 + i * CardW * 0.3
                cy = CardH * 0.05
            Case Is = "crdPlayerFour"
                cx = SW - CardW - i * CardW * 0.14
                cy = (SH - CardH) / 2
            End Select
        End If
            
        If IsDeal Then
            Dim OldSpeed As Integer
              
            If LA.Speed <> 0 Then
                OldSpeed = LA.Speed
                LA.Speed = LA.Speed + 1
            End If
            
            PlaySound SND_BLIP
            
            crdStock(GetCurrentStock).ZOrder 0
            If i < PlayerCards.Count - 1 Then
                crdStock(GetCurrentStock).Rank = PlayerCards(i + 1).Rank
                crdStock(GetCurrentStock).Suit = PlayerCards(i + 1).Suit
            Else
                If PlayerCards(0).Name = "crdPlayerOne" Then
                    crdStock(GetCurrentStock).Rank = crdPlayerTwo(0).Rank
                    crdStock(GetCurrentStock).Suit = crdPlayerTwo(0).Suit
                ElseIf PlayerCards(0).Name = "crdPlayerTwo" Then
                    If Setting.Opponents = 2 Then
                        crdStock(GetCurrentStock).Rank = crdPlayerThree(0).Rank
                        crdStock(GetCurrentStock).Suit = crdPlayerThree(0).Suit
                    End If
                ElseIf PlayerCards(0).Name = "crdPlayerThree" Then
                    If Setting.Opponents = 3 Then
                        crdStock(GetCurrentStock).Rank = crdPlayerFour(0).Rank
                        crdStock(GetCurrentStock).Suit = crdPlayerFour(0).Suit
                    End If
                Else
                    crdStock(GetCurrentStock).Rank = crdWaste.Rank
                    crdStock(GetCurrentStock).Suit = crdWaste.Suit
                End If
            End If
                        
            Do While Not LA.Linear(PlayerCards(i), _
                                   crdStock(GetCurrentStock).Left, _
                                   crdStock(GetCurrentStock).Top, _
                                   cx, cy)
            Loop
                        
            If LA.Speed <> 0 Then LA.Speed = OldSpeed
        Else
            PlayerCards(i).Move cx, cy
            PlayerCards(i).ZOrder 0
        End If
    Next i
    
    If IsDeal Then
        Dim bVal As Boolean
        
        If Setting.Opponents = 1 Then
            If PlayerCards(0).Name = "crdPlayerTwo" Then
                bVal = True
            End If
        ElseIf Setting.Opponents = 2 Then
            If PlayerCards(0).Name = "crdPlayerThree" Then
                bVal = True
            End If
        ElseIf Setting.Opponents = 3 Then
            If PlayerCards(0).Name = "crdPlayerFour" Then
                bVal = True
            End If
        End If
        
        If bVal Then
            crdStock(GetCurrentStock).ZOrder 0
            crdStock(GetCurrentStock).Rank = Uno.Stock(Uno.Stock.Count).Rank
            crdStock(GetCurrentStock).Suit = Uno.Stock(Uno.Stock.Count).Suit
            
            crdWaste.ZOrder 0
            crdWaste.Visible = True
            CardAni crdWaste, 2
        End If
    End If
End Sub

Private Sub AlignCards()
    If Setting.Opponents = 1 Then
        Call AlignPlayerCards(crdPlayerOne)
        Call AlignPlayerCards(crdPlayerTwo)
    ElseIf Setting.Opponents = 2 Then
        Call AlignPlayerCards(crdPlayerOne)
        Call AlignPlayerCards(crdPlayerTwo)
        Call AlignPlayerCards(crdPlayerThree)
    Else
        Call AlignPlayerCards(crdPlayerOne)
        Call AlignPlayerCards(crdPlayerTwo)
        Call AlignPlayerCards(crdPlayerThree)
        Call AlignPlayerCards(crdPlayerFour)
    End If
End Sub

Private Sub mnuHelpContents_Click()
    ' HTML Help Workshop is available in Visual Studio.Net
    HelpWinHwnd = HtmlHelp(Me.hwnd, App.Path & "\uno.chm", HH_DISPLAY_TOPIC, ByVal "overview.htm")
End Sub

Private Sub mnuViewCheerLeader_Click()
    mnuViewCheerLeader.Checked = Not mnuViewCheerLeader.Checked
    tmrCheerLeader.Enabled = mnuViewCheerLeader.Checked
    picTrackWin.Cls
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tblToolbar.Visible = mnuViewToolbar.Checked
    Call Form_Resize
End Sub

Private Sub picTable_Resize()
    On Error Resume Next
        
    If IsLogo Then
        Dim temp As Boolean
        
        temp = tmrWinAni.Enabled
        tmrWinAni.Enabled = False
        picTable.Cls
        Set picTable.Picture = Nothing
        ShowLogo picTable.hdc, picTable.ScaleWidth, picTable.ScaleHeight
        Set picTable.Picture = picTable.Image
        tmrWinAni.Enabled = temp
    End If
    
    Call AlignCards
    
    Dim i  As Integer
    Dim SW As Integer
    Dim SH As Integer
    
    SW = picTable.ScaleWidth
    SH = picTable.ScaleHeight
    
    For i = crdStock.LBound To crdStock.UBound
        crdStock(i).Move (SW - CardW) / 2 - CardW * 0.6 + i * 2 - picColor.ScaleWidth * 0.5, _
                         (SH - CardH) / 2 + i * 2, CardW, CardH
        crdStock(i).ZOrder 0
    Next i
    
    crdWaste.Move (SW - CardW) / 2 + CardW * 0.6 - picColor.ScaleWidth * 0.5, _
                  (SH - CardH) / 2, CardW, CardH
    picColor.Move crdWaste.Left + crdWaste.Width, crdWaste.Top + 1, _
                  picColor.ScaleWidth, crdWaste.Height - 2
    shpStock.Move crdStock(0).Left, crdStock(0).Top, CardW, CardH
    shpWaste.Move crdWaste.Left, crdWaste.Top, CardW, CardH
    shpCircle.Move crdStock(0).Left + (CardW - shpCircle.Width) / 2, _
                   crdStock(1).Top + (CardH - shpCircle.Height) / 2
    Call AlignPlayersName
    
    If Setting.BkPicture.handle <> 0 Then
        Dim rcRect As RECT
        
        GetClientRect picTable.hwnd, rcRect
        Background picTable.hdc, Setting.BkPicture, _
                   rcRect.Right, rcRect.Bottom
        RefreshWindow picTable.hwnd
    End If
End Sub

Private Function GetCardToolTip(ByVal Rank As Integer, ByVal Suit As Integer, ByVal Points As Integer) As String
    Dim arrRank() As Variant
    Dim arrSuit() As Variant
            
    arrRank = Array("ZERO", "ONE", "TWO", "THREE", _
                    "FOUR", "FIVE", "SIX", "SEVEN", _
                    "EIGHT", "NINE", "DRAW TWO", _
                    "REVERSE", "SKIP", "WILD CARD", _
                    "DRAW FOUR")
            
    arrSuit = Array("BLUE", "RED", "GREEN", "YELLOW")
    
    GetCardToolTip = "Rank = " & arrRank(Rank) & " : " & _
                     "Suit = " & arrSuit(Suit) & " : " & _
                     "Score = " & Points
End Function

Private Sub TrackCard(Card As Object)
    If Card.Face = uno_FCUp Then
        If mnuGameShowTooltip.Checked Then
            Card.ToolTipText = _
                GetCardToolTip(Card.Rank, Card.Suit, Card.Points)
        End If
    
        Dim cx As Integer
        Dim cy As Integer
            
        cx = (picViewer.ScaleWidth - CardW) / 2 - 1
        cy = (picViewer.ScaleHeight - CardH) / 2 - 1
        
        If Not picViewer.Visible Then picViewer.Visible = True
        
        picViewer.Line (cx + 2, cy + 2)- _
                         (cx + CardW + 2, cy + CardH + 2), &H0, BF
        picViewer.PaintPicture Card.Picture, cx, cy, CardW, CardH
    
        Select Case Card.Name
        Case Is = "crdPlayerOne"
            lblTrackCard.BackColor = &H800000
            lblTrackCard.Caption = "Player : " & Setting.PlayerName(0)
        Case Is = "crdPlayerTwo"
            lblTrackCard.BackColor = &H80&
            lblTrackCard.Caption = "Player : " & Setting.PlayerName(1)
        Case Is = "crdPlayerThree"
            lblTrackCard.BackColor = &H80&
            lblTrackCard.Caption = "Player : " & Setting.PlayerName(2)
        Case Is = "crdPlayerFour"
            lblTrackCard.BackColor = &H80&
            lblTrackCard.Caption = "Player : " & Setting.PlayerName(3)
        Case Is = "crdStock"
            lblTrackCard.BackColor = &H808000
            lblTrackCard.Caption = "Stock Pile"
        Case Is = "crdWaste"
            lblTrackCard.BackColor = &H8000&
            lblTrackCard.Caption = "Waste Pile"
        End Select
    End If
End Sub

Private Sub ShowActiveColor()
    Dim i        As Integer
    Dim cy       As Integer
    Dim dv       As Integer
    Dim OldTimer As Single
    
    With picColor
        cy = .ScaleHeight / 2
        dv = 255 / cy
    
        For i = 0 To cy
            Select Case crdWaste.Suit
            Case Is = 0 ' Blue
                .ForeColor = RGB(0, 0, i * dv)
            Case Is = 1 ' Red
                .ForeColor = RGB(i * dv, 0, 0)
            Case Is = 2 ' Green
                .ForeColor = RGB(0, i * dv, 0)
            Case Is = 3 ' Yellow
                .ForeColor = RGB(i * dv, i * dv, 0)
            End Select
            
            picColor.Line (0, i)-(.ScaleWidth, i)
            picColor.Line (0, .ScaleHeight - i)- _
                          (.ScaleWidth, .ScaleHeight - i)
            
            If i Mod 4 = 0 Then
                OldTimer = Timer
                Do While Abs(Timer - OldTimer) < 0.01
                    DoEvents
                Loop
           End If
        Next i
    End With
End Sub

Private Function PlayerMove(ByVal SelCard As Integer) As Boolean
    On Error GoTo ErrHandler
    
    If Uno.IsMoveValid(CurrentPlayer, SelCard, crdWaste.Rank, crdWaste.Suit) Then
        Dim CurSuit As Integer
        Dim OldSuit As Integer
        Dim OldX    As Integer
        Dim OldY    As Integer
        
        OldX = CurrentPlayer(SelCard).Left
        OldY = CurrentPlayer(SelCard).Top
            
        Set LastPlayer = CurrentPlayer
        CardAni CurrentPlayer(SelCard), 1
        CurSuit = CurrentPlayer(SelCard).Suit
        
        If Not mnuGameShowDemo.Checked Then
            If CurrentPlayer(0).Name = "crdPlayerOne" Then
                If (CurrentPlayer(SelCard).Rank = uno_RCWild) Or _
                   (CurrentPlayer(SelCard).Rank = uno_RCDrawFour) Then
                    
                    frmColor.Show vbModal, Me
                    If frmColor.SelColor = -1 Then
                        ' cancel selected

                        Do While Not LA.Linear(CurrentPlayer(SelCard), _
                                               crdWaste.Left, _
                                               crdWaste.Top, _
                                               OldX, _
                                               OldY)
                        Loop
                        
                        AlignPlayerCards CurrentPlayer
                        Exit Function
                    Else
                        CurSuit = frmColor.SelColor
                    End If
                End If
            End If
        End If
        
        OldSuit = crdWaste.Suit
        crdWaste.Rank = CurrentPlayer(SelCard).Rank
        crdWaste.Suit = CurSuit
        History CurrentPlayer(SelCard).Name, "plays", _
                crdWaste.Rank, crdWaste.Suit
        Uno.Throw CurrentPlayer, SelCard
        AlignPlayerCards CurrentPlayer
        
        If (crdWaste.Rank = uno_RCDrawTwo) Or (crdWaste.Rank = uno_RCDrawFour) Then
            Dim i         As Integer
            Dim AddedCard As Integer
            Dim oTemp     As Object
            
            AddedCard = IIf(crdWaste.Rank = uno_RCDrawTwo, 2, 4)
            AddedCard = IIf(Uno.Stock.Count >= AddedCard, AddedCard, _
                                                          Uno.Stock.Count)
            Set oTemp = GetNextPlayer(CurrentPlayer, Rotation)
            
            For i = 1 To AddedCard
                If Uno.Stock.Count > 1 Then
                    crdStock(GetCurrentStock).Rank = Uno.Stock(Uno.Stock.Count - 1).Rank
                    crdStock(GetCurrentStock).Suit = Uno.Stock(Uno.Stock.Count - 1).Suit
                Else
                    crdStock(GetCurrentStock).Visible = False
                End If
                
                Uno.Pick oTemp
                CardAni oTemp(oTemp.Count - 1), 0
                Call UpdateStockPile
                
                AlignPlayerCards oTemp
                If oTemp(0).Face = uno_FCUp Then
                    Uno.SortCards oTemp
                End If
            Next i
            
            If AddedCard > 0 Then
                If AddedCard = 1 Then
                    History oTemp(0).Name, "Draws a card", oTemp(oTemp.Count - 1).Rank, _
                                                           oTemp(oTemp.Count - 1).Suit
                Else
                    History oTemp(0).Name, "Draws " & AddedCard & " cards", oTemp(oTemp.Count - 1).Rank, _
                                                                            oTemp(oTemp.Count - 1).Suit
                End If
            End If
            Set CurrentPlayer = GetNextPlayer(oTemp, Rotation)
        ElseIf crdWaste.Rank = uno_RCReverse Then
            If Setting.Opponents = 1 Then
                ' do nothing
            Else
                Rotation = IIf(Rotation = 0, 1, 0)
                Set CurrentPlayer = GetNextPlayer(CurrentPlayer, Rotation)
            End If
        ElseIf crdWaste.Rank = uno_RCSkip Then
            If Setting.Opponents = 1 Then
                ' do nothing
            Else
                Set CurrentPlayer = _
                    GetNextPlayer(GetNextPlayer(CurrentPlayer, Rotation), Rotation)
            End If
            
            AlignPlayerCards CurrentPlayer
        Else
            Set CurrentPlayer = GetNextPlayer(CurrentPlayer, Rotation)
        End If
        
        Call ShowActivePlayer
        If CBool(chkAutoSort.Value) Then
            If CurrentPlayer(0).Face = uno_FCUp Then
                Uno.SortCards CurrentPlayer
            End If
        End If
        
        If OldSuit <> crdWaste.Suit Then
            Call ShowActiveColor
        End If
        
        PlayerMove = True
    Else
        PlayerMove = False
        If mnuGameSound.Checked Then
            Call Beep
        End If
    End If
    Exit Function
ErrHandler:
End Function

Private Sub History(ByVal Player As String, ByVal Message As String, ByVal Rank As Integer, ByVal Suit As Integer)
    Dim i        As Integer
    Dim sTemp    As String
    Dim CurIcon  As Integer
    
    Select Case Setting.Opponents
    Case Is = 1
        Select Case Player
        Case Is = "crdPlayerOne"
            Player = lblPlayerName(0).Caption
        Case Is = "crdPlayerTwo"
            Player = lblPlayerName(1).Caption
        End Select
    Case Is = 2
        Select Case Player
        Case Is = "crdPlayerOne"
            Player = lblPlayerName(0).Caption
        Case Is = "crdPlayerTwo"
            Player = lblPlayerName(2).Caption
        Case Is = "crdPlayerThree"
            Player = lblPlayerName(1).Caption
        End Select
    Case Else
        Select Case Player
        Case Is = "crdPlayerOne"
            Player = lblPlayerName(0).Caption
        Case Is = "crdPlayerTwo"
            Player = lblPlayerName(2).Caption
        Case Is = "crdPlayerThree"
            Player = lblPlayerName(1).Caption
        Case Is = "crdPlayerFour"
            Player = lblPlayerName(3).Caption
        End Select
    End Select
    
    Player = "[" & Player & "]"
    
    If Left$(Message, 5) <> "Draws" Then
        Select Case Suit
        Case Is = uno_SCBlue
            sTemp = "Blue"
        Case Is = uno_SCRed
            sTemp = "Red"
        Case Is = uno_SCGreen
            sTemp = "Green"
        Case Is = uno_SCYellow
            sTemp = "Yellow"
        End Select
    
        sTemp = Player & " " & Message & " " & sTemp & " "
    
        Select Case Rank
        Case uno_RCZero To uno_RCNine
            sTemp = sTemp & Rank
        Case Is = uno_RCDrawTwo
            sTemp = sTemp & "Draw 2"
        Case Is = uno_RCReverse
            sTemp = sTemp & "Reverse"
        Case Is = uno_RCSkip
            sTemp = sTemp & "Skip"
        Case Is = uno_RCWild
            sTemp = sTemp & "Wild Card"
        Case Is = uno_RCDrawFour
            sTemp = sTemp & "Draw 4"
        End Select
        
        CurIcon = Suit + 1
    Else
        CurIcon = 5
        sTemp = Player & " " & Message
    End If
    
    With lvwHistory
        LockWindowUpdate .hwnd
        .ListItems.Add , , sTemp, , CurIcon
        
        If Player = "[" & lblPlayerName(0).Caption & "]" Then
            .ListItems(.ListItems.Count).ForeColor = vbRed
        Else
            .ListItems(.ListItems.Count).ForeColor = vbBlack
        End If
        
        .ListItems(.ListItems.Count).ToolTipText = sTemp
        .ListItems(.ListItems.Count).Selected = True
        .ListItems(.ListItems.Count).EnsureVisible
        LockWindowUpdate 0
    End With
    
    sbStatusBar.Panels(1).Picture = imlStatIcons.ListImages(Suit + 1).Picture
    sbStatusBar.Panels(1).Text = sTemp
End Sub

Private Sub UpdateStockPile()
    Dim temp As Integer
    
    If crdStock(2).Visible Then
        temp = CInt(TotalStockLeft / 3) * 2
        If temp >= Uno.Stock.Count Then
            crdStock(1).MousePointer = vbCustom
            Set crdStock(1).MouseIcon = LoadResPicture(101, vbResCursor)
            crdStock(2).Visible = False
        End If
    ElseIf crdStock(1).Visible Then
        temp = CInt(TotalStockLeft / 3)
        If temp >= Uno.Stock.Count Then
            crdStock(0).MousePointer = vbCustom
            Set crdStock(0).MouseIcon = LoadResPicture(101, vbResCursor)
            crdStock(1).Visible = False
        End If
    End If
End Sub

Private Function GetNextPlayer(Player As Object, Rotation As Integer) As Object
    If Setting.Opponents = 1 Then
        Select Case Player(0).Name
        Case Is = "crdPlayerOne"
            Set GetNextPlayer = crdPlayerTwo
        Case Is = "crdPlayerTwo"
            Set GetNextPlayer = crdPlayerOne
        End Select
    ElseIf Setting.Opponents = 2 Then
        Select Case Player(0).Name
        Case Is = "crdPlayerOne"
            Set GetNextPlayer = _
                IIf(Rotation = 0, crdPlayerTwo, crdPlayerThree)
        Case Is = "crdPlayerTwo"
            Set GetNextPlayer = _
                IIf(Rotation = 0, crdPlayerThree, crdPlayerOne)
        Case Is = "crdPlayerThree"
            Set GetNextPlayer = _
                IIf(Rotation = 0, crdPlayerOne, crdPlayerTwo)
        End Select
    Else
        Select Case Player(0).Name
        Case Is = "crdPlayerOne"
            Set GetNextPlayer = _
                IIf(Rotation = 0, crdPlayerTwo, crdPlayerFour)
        Case Is = "crdPlayerTwo"
            Set GetNextPlayer = _
                IIf(Rotation = 0, crdPlayerThree, crdPlayerOne)
        Case Is = "crdPlayerThree"
            Set GetNextPlayer = _
                IIf(Rotation = 0, crdPlayerFour, crdPlayerTwo)
        Case Is = "crdPlayerFour"
            Set GetNextPlayer = _
                IIf(Rotation = 0, crdPlayerOne, crdPlayerThree)
        End Select
    End If
End Function

Private Sub tblToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case Is = 1 ' Open
        Call mnuGameOpen_Click
    Case Is = 2 ' Save
        Call mnuGameSave_Click
    Case Is = 4 ' Deal
        Call mnuGameDeal_Click
    Case Is = 5 ' Opponent
        Button.Image = IIf(Button.Image + 1 > 3, 1, Button.Image + 1)
        Button.ToolTipText = _
            Button.Image & " " & IIf(Button.Image = 1, _
                                     "Opponent", "Opponents")
    Case Is = 7 ' Cheat
        Call mnuGameCheat_Click
    Case Is = 8 ' Hint
        Dim SelMove  As Integer
        Dim OldTimer As Single
        Dim Card     As Object
        Dim OldLevel As Integer
        
        OldLevel = Uno.GetLevelMode
        Uno.SetLevelMode = LEVEL_HARD
        SelMove = Uno.AI(crdPlayerOne, crdWaste.Rank, crdWaste.Suit)
        Uno.SetLevelMode = OldLevel
        
        If SelMove <> -1 Then
            Set Card = crdPlayerOne(SelMove)
        Else
            Set Card = crdStock(GetCurrentStock)
        End If
            
        PatBlt Card.hdc, 0, 0, CardW, CardH, DSTINVERT
        Card.Refresh
        
        Me.Enabled = False
        OldTimer = Timer
        Do While Abs(Timer - OldTimer) < 0.5
            DoEvents
        Loop
        Me.Enabled = True
            
        PatBlt Card.hdc, 0, 0, CardW, CardH, DSTINVERT
        Card.Refresh
    Case Is = 9 ' Show ToolTip
        Call mnuGameShowTooltip_Click
    Case Is = 10 ' Separator
    Case Is = 11 ' Sound
        If Button.Value = tbrPressed Then
            Button.ToolTipText = "Sound On"
        Else
            Button.ToolTipText = "Sound Off"
        End If
        
        mnuGameSound.Checked = IIf(Button.Value = tbrPressed, True, False)
    Case Is = 12 ' Separator
    Case Is = 13 ' Help
        Call mnuHelpContents_Click
    Case Is = 14 ' Exit
        Call mnuGameExit_Click
    End Select
End Sub

Private Sub tblToolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    tblToolbar.Buttons("opponent").Image = ButtonMenu.Index
    
    Setting.Opponents = ButtonMenu.Index
End Sub

Private Sub GameStart()
    Dim EndGame As Boolean
        
    Do While True
        If Not IsGameExit Then
            If ResetGame Then Exit Do
            
            If Uno.Stock.Count = 0 Then
                EndGame = True
                Exit Do
            End If
            
            Select Case Setting.Opponents
            Case Is = 1
                If Not crdPlayerOne(0).Visible Then
                    EndGame = True
                ElseIf Not crdPlayerTwo(0).Visible Then
                    EndGame = True
                Else
                    EndGame = False
                End If
            Case Is = 2
                If Not crdPlayerOne(0).Visible Then
                    EndGame = True
                ElseIf Not crdPlayerTwo(0).Visible Then
                    EndGame = True
                ElseIf Not crdPlayerThree(0).Visible Then
                    EndGame = True
                Else
                    EndGame = False
                End If
            Case Is = 3
                If Not crdPlayerOne(0).Visible Then
                    EndGame = True
                ElseIf Not crdPlayerTwo(0).Visible Then
                    EndGame = True
                ElseIf Not crdPlayerThree(0).Visible Then
                    EndGame = True
                ElseIf Not crdPlayerFour(0).Visible Then
                    EndGame = True
                Else
                    EndGame = False
                End If
            End Select
            
            If EndGame Then Exit Do
            
            If Me.WindowState <> vbMinimized Then
                If Not mnuGameShowDemo.Checked Then
                    If CurrentPlayer(0).Name = "crdPlayerOne" Then
                    Else
                        Call Computer
                    End If
                Else
                    Call Computer
                End If
                
                Set PreviousPlayer = CurrentPlayer
            End If
    
            DoEvents
        Else
            Exit Do
        End If
    Loop
    
    If EndGame Then
        picViewer.Visible = False
        shpRectL(0).Visible = False
        shpRectL(1).Visible = False
        shpRectR(0).Visible = False
        shpRectR(1).Visible = False
        linLine(0).Visible = False
        linLine(1).Visible = False
        lblTrackCard.Caption = "*** UNO GAME 2.2 ***"
        
        If Not IsDone Then
            Dim Card As Object
                
            For Each Card In Me.Controls
                If TypeName(Card) = "UnoCard" Then
                    If Card.Visible Then
                        Card.Enabled = False
                    End If
                End If
            Next Card

            IsDone = True
            frmScore.Show vbModal, Me
            
            If frmScore.IsWinner And Not mnuGameShowDemo.Checked And Setting.WinnerSelAni <> 0 Then
                ' play animation
                Dim i    As Integer
                Dim temp As Object
                Dim Cntr As Integer
                
                picColor.Visible = False
                shpStock.Visible = False
                shpWaste.Visible = False
                shpCircle.Visible = False
                
                For i = lblPlayerName.LBound To lblPlayerName.UBound
                    lblPlayerName(i).Visible = False
                Next i
                
                Set WinSelCardAni = Nothing
                
                ShowLogo picTable.hdc, picTable.ScaleWidth, picTable.ScaleHeight
                Set picTable.Picture = picTable.Image
                
                OldSettingMaxCard = Setting.MaxCard
                
                For Each temp In Me.Controls
                    If TypeName(temp) = "UnoCard" Then
                        If temp.Visible And Cntr < Setting.MaxCard Then
                            temp.Face = uno_FCUp
                            WinSelCardAni.Add temp
                            Cntr = Cntr + 1
                        Else
                            temp.Visible = False
                        End If
                    End If
                Next temp
                
                WinnerAni.Reset = True
                WinnerAni.MaxCards = WinSelCardAni.Count
                WinnerAni.SaveCardPosition WinSelCardAni
                tmrWinAni.Enabled = True
            End If
        End If
    End If
End Sub

Private Function GetCurrentStock() As Integer
    If crdStock(2).Visible Then
        GetCurrentStock = 2
    ElseIf crdStock(1).Visible Then
        GetCurrentStock = 1
    Else
        GetCurrentStock = 0
    End If
End Function

Private Sub SortPlayerCards()
    Uno.SortCards crdPlayerOne
    
    If mnuGameCheat.Checked And CBool(chkAutoSort.Value) Then
        If Setting.Opponents = 1 Then
            Uno.SortCards crdPlayerTwo
        ElseIf Setting.Opponents = 2 Then
            Uno.SortCards crdPlayerTwo
            Uno.SortCards crdPlayerThree
        Else
            Uno.SortCards crdPlayerTwo
            Uno.SortCards crdPlayerThree
            Uno.SortCards crdPlayerFour
        End If
    End If
End Sub

Private Sub ShowActivePlayer()
    Dim i As Integer
    
    For i = lblPlayerName.LBound To lblPlayerName.UBound
        lblPlayerName(i).ForeColor = RGB(255, 255, 255)
    Next i
    
    Select Case Setting.Opponents
    Case Is = 1
        Select Case CurrentPlayer(0).Name
        Case Is = "crdPlayerOne"
            lblPlayerName(0).ForeColor = RGB(255, 255, 0)
        Case Is = "crdPlayerTwo"
            lblPlayerName(1).ForeColor = RGB(255, 255, 0)
        End Select
    Case Is = 2
        Select Case CurrentPlayer(0).Name
        Case Is = "crdPlayerOne"
            lblPlayerName(0).ForeColor = RGB(255, 255, 0)
        Case Is = "crdPlayerTwo"
            lblPlayerName(2).ForeColor = RGB(255, 255, 0)
        Case Is = "crdPlayerThree"
            lblPlayerName(1).ForeColor = RGB(255, 255, 0)
        End Select
    Case Else
        Select Case CurrentPlayer(0).Name
        Case Is = "crdPlayerOne"
            lblPlayerName(0).ForeColor = RGB(255, 255, 0)
        Case Is = "crdPlayerTwo"
            lblPlayerName(2).ForeColor = RGB(255, 255, 0)
        Case Is = "crdPlayerThree"
            lblPlayerName(1).ForeColor = RGB(255, 255, 0)
        Case Is = "crdPlayerFour"
            lblPlayerName(3).ForeColor = RGB(255, 255, 0)
        End Select
    End Select
End Sub

Public Sub AlignPlayersName()
    Dim SW As Integer
    Dim SH As Integer
    
    SW = picTable.ScaleWidth
    SH = picTable.ScaleHeight
    
    If lblPlayerName(0).Visible Then
        lblPlayerName(0).Move (SW - lblPlayerName(0).Width) / 2, _
                               SH - CardH - lblPlayerName(0).Height * 1.5
    End If
    
    If lblPlayerName(1).Visible Then
        lblPlayerName(1).Move (SW - lblPlayerName(1).Width) / 2, _
                               CardH + lblPlayerName(1).Height * 0.5
    End If
    
    If lblPlayerName(2).Visible Then
        lblPlayerName(2).Move 5, (SH - CardH) / 2 - CardH * 0.2 + 4
    End If
    
    If lblPlayerName(3).Visible Then
        lblPlayerName(3).Move SW - lblPlayerName(3).Width - 5, _
                             (SH - CardH) / 2 - CardH * 0.2 + 4
    End If
End Sub

Private Sub CardAni(Card As Object, Op As Integer)
    Set SelCardAni = Card
    
    PlaySound SND_BLIP
    
    If Op = 0 Then
        ' stock
        Dim xmid As Integer
        
        xmid = (picTable.ScaleWidth - Card.Width) / 2
        
        Select Case Setting.Opponents
        Case Is = 1 ' Player 1, Player 2
            Do While Not LA.Linear(SelCardAni, _
                                   crdStock(GetCurrentStock).Left, _
                                   crdStock(GetCurrentStock).Top, _
                                   xmid, _
                                   SelCardAni.Top)
            Loop
        Case Is = 2 ' Player 1, Player 2, Player 3
            If SelCardAni.Name = "crdPlayerTwo" Then
                
                Do While Not LA.Linear(SelCardAni, _
                                       crdStock(GetCurrentStock).Left, _
                                       crdStock(GetCurrentStock).Top, _
                                      (crdPlayerTwo.Count - 1) * CardW * 0.14, _
                                       crdPlayerTwo(0).Top)
                Loop
            Else
                Do While Not LA.Linear(SelCardAni, _
                                       crdStock(GetCurrentStock).Left, _
                                       crdStock(GetCurrentStock).Top, _
                                       xmid, _
                                       SelCardAni.Top)
                Loop
            End If
        Case Is = 3   ' Player 1, Player 2, Player 3, Player 4
            If (SelCardAni.Name = "crdPlayerOne") Or _
               (SelCardAni.Name = "crdPlayerThree") Then
                Do While Not LA.Linear(SelCardAni, _
                                       crdStock(GetCurrentStock).Left, _
                                       crdStock(GetCurrentStock).Top, _
                                       xmid, _
                                       SelCardAni.Top)
                Loop
            ElseIf SelCardAni.Name = "crdPlayerTwo" Then
                Do While Not LA.Linear(SelCardAni, _
                                       crdStock(GetCurrentStock).Left, _
                                       crdStock(GetCurrentStock).Top, _
                                      (crdPlayerTwo.Count - 1) * CardW * 0.14, _
                                       crdPlayerTwo(0).Top)
                Loop
            Else
                Dim SW As Integer
                
                SW = picTable.ScaleWidth
                Do While Not LA.Linear(SelCardAni, _
                                       crdStock(GetCurrentStock).Left, _
                                       crdStock(GetCurrentStock).Top, _
                                       SW - CardW - (crdPlayerFour.Count - 1) * CardW * 0.14, _
                                       crdPlayerTwo(0).Top)
                Loop
            End If
        End Select
    ElseIf Op = 1 Then
        ' waste
        Do While Not LA.Linear(SelCardAni, _
                               SelCardAni.Left, _
                               SelCardAni.Top, _
                               crdWaste.Left, _
                               crdWaste.Top)
        Loop
    ElseIf Op = 2 Then
        ' stock to waste
        Do While Not LA.Linear(SelCardAni, _
                               crdStock(GetCurrentStock).Left, _
                               crdStock(GetCurrentStock).Top, _
                               shpWaste.Left, _
                               shpWaste.Top)
                          
        Loop
    End If
End Sub

Private Sub Computer()
    Dim GetMove As Integer
            
    picTable.Enabled = False
    GetMove = Uno.AI(CurrentPlayer, crdWaste.Rank, _
                                    crdWaste.Suit)
    If GetMove <> -1 Then
        If Not PlayerMove(GetMove) Then
            If Uno.Stock.Count >= 1 Then
                crdStock_Click GetCurrentStock
            Else
                ' no more stock
            End If
        End If
    Else
        If Uno.Stock.Count >= 1 Then
            crdStock_Click GetCurrentStock
        End If
    End If
        
    picTable.Enabled = True
End Sub

Private Sub tmrCheerLeader_Timer()
    Dim Filename_1 As String
    Dim Filename_2 As String
    Dim PicW       As Integer
    Dim PicH       As Integer
    
    PicW = 59
    PicH = 70
    
    cl_1 = IIf(cl_1 + 1 > 7, cl_1 = 0, cl_1 + 1)
    cl_2 = IIf(cl_2 + 1 > 7, cl_2 = 0, cl_2 + 1)
    
    Filename_1 = Dir$(App.Path & "\image\cheerleader0" & cl_1 + 1 & ".jpg")
    Filename_2 = Dir$(App.Path & "\image\cheerleader0" & cl_2 + 1 & ".jpg")
    
    If Filename_1 <> "" Then
        picTrackWin.PaintPicture LoadPicture(App.Path & "\image\" & Filename_1), _
                                 -3, (picTrackWin.ScaleHeight - PicH) / 2, PicW, PicH
        picTrackWin.PaintPicture LoadPicture(App.Path & "\image\" & Filename_1), _
                                 picTrackWin.ScaleWidth - PicW + 3, (picTrackWin.ScaleHeight - PicH) / 2, PicW, PicH
    End If
    
    If Filename_2 <> "" And Not picViewer.Visible Then
        picTrackWin.PaintPicture LoadPicture(App.Path & "\image\" & Filename_2), _
                                 (picTrackWin.ScaleWidth - PicW) / 2, (picTrackWin.ScaleHeight - PicH) / 2, PicW, PicH
    End If
End Sub

Private Sub tmrIconAni_Timer()
    Static i As Integer
    
    i = IIf(i + 1 > 3, i = 0, i + 1)
    
    Select Case i
    Case Is = 0
        Set Me.Icon = imlLogoIcon.ListImages(1).Picture
    Case Is = 1
        Set Me.Icon = imlLogoIcon.ListImages(2).Picture
    Case Is = 2
        Set Me.Icon = imlLogoIcon.ListImages(3).Picture
    Case Is = 3
        Set Me.Icon = imlLogoIcon.ListImages(4).Picture
    End Select
End Sub

Private Sub tmrValidCard_Timer()
    If IsGameExit Then tmrValidCard.Enabled = False
    
    If Me.WindowState <> vbMinimized Then
        linLine(0).X1 = linLine(0).X1 + 3
        linLine(0).X2 = linLine(0).X2 + 3
        linLine(1).X1 = linLine(1).X1 - 3
        linLine(1).X2 = linLine(1).X2 - 3
        
        shpRectL(0).Left = shpRectL(0).Left + 3
        shpRectL(1).Left = shpRectL(1).Left + 3
        shpRectR(0).Left = shpRectR(0).Left - 3
        shpRectR(1).Left = shpRectR(1).Left - 3
        
        If linLine(0).X1 > picTrackWin.ScaleWidth / 3 - 10 Then
            tmrValidCard.Enabled = False
        End If
    End If
End Sub

Private Sub tmrWinAni_Timer()
    WinnerAni.FallType = Setting.FallType
    WinnerAni.WindType = Setting.WindType
    WinnerAni.ShowTrail = Setting.ShowTrail
    
    WinnerAni.BounceDistX = Setting.BounceDistX
    WinnerAni.BounceDistY = Setting.BounceDistY
    WinnerAni.BounceSpeedX = Setting.BounceSpeedX
    WinnerAni.BounceSpeedY = Setting.BounceSpeedY
    WinnerAni.ScatterSpeedX = Setting.ScatterSpeedX
    WinnerAni.ScatterSpeedY = Setting.ScatterSpeedY
    WinnerAni.SpinDistance = Setting.SpinDist
    
    If Setting.WinnerSelAni <> OldWinnerAni Then
        WinnerAni.Reset = True
        OldWinnerAni = Setting.WinnerSelAni
    Else
        Select Case Setting.WinnerSelAni
        Case Is = 1
            WinnerAni.Bounce picTable, WinSelCardAni
        Case Is = 2
            WinnerAni.Scatter picTable, WinSelCardAni
        Case Is = 3
            WinnerAni.Spin picTable, WinSelCardAni
        Case Is = 4
            WinnerAni.Fall picTable, WinSelCardAni
        Case Is = 5
            WinnerAni.Wind picTable, WinSelCardAni
        End Select
    End If
End Sub
