VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   220
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim AboutLoc As String

    On Error Resume Next

    AboutLoc = App.Path & "\uno.dat"
    If Dir$(AboutLoc) <> "" Then
        Set Me.Picture = LoadPicture(AboutLoc)
    End If
End Sub
