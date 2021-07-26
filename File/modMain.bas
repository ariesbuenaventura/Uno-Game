Attribute VB_Name = "modMain"
Option Explicit
                                                 
Public Const SND_BLIP = 0

Private Const LS = 101 ' logo sprite
Private Const LM = 102 ' logo mask

Public Uno As New clsUno
                                  
Public Sub Main()
    Dim OldTimer As Single
    
    OldTimer = Timer
    
    frmSplash.Show
    frmSplash.Refresh
    
    Load frmMain
    App.HelpFile = App.Path & "\uno.chm"
    
    ' make sure that the splash screen will be shown
    ' exactly 2 seconds or more.
    Do While Abs(Timer - OldTimer) < 2
        DoEvents
    Loop
    
    Unload frmSplash
    frmMain.Show
End Sub

Public Sub PlaySound(ByVal Sound_Type As Integer)
    If frmMain.mnuGameSound.Checked Then
        Dim SoundFile As String
    
        If Sound_Type = 0 Then
            SoundFile = "blip.wav"
        End If
        
        If Dir$(App.Path & "\sound\" & SoundFile) <> "" Then
            sndPlaySound App.Path & "\sound\" & SoundFile, &H1 Or &H2 Or &H2000
        End If
    End If
End Sub

Public Sub ShowLogo(ByVal hDestDC As Long, ByVal WindowW As Integer, ByVal WindowH As Integer)
    Dim lhDCs As Long
    Dim lhDCm As Long
    Dim tBMPs As BITMAP
    Dim tBMPm As BITMAP

    GetObjectAPI LoadResPicture(LS, vbResBitmap).handle, _
                 Len(tBMPs), tBMPs
    GetObjectAPI LoadResPicture(LM, vbResBitmap).handle, _
                 Len(tBMPm), tBMPm

    lhDCs = CreateCompatibleDC(hDestDC)
    lhDCm = CreateCompatibleDC(hDestDC)

    SelectObject lhDCs, LoadResPicture(LS, vbResBitmap).handle
    SelectObject lhDCm, LoadResPicture(LM, vbResBitmap).handle

    BitBlt hDestDC, (WindowW - tBMPm.bmWidth) / 2, _
                    (WindowH - tBMPm.bmHeight) / 2, _
                     tBMPm.bmWidth, tBMPm.bmHeight, _
           lhDCm, 0, 0, vbSrcAnd
    BitBlt hDestDC, (WindowW - tBMPs.bmWidth) / 2, _
                    (WindowH - tBMPs.bmHeight) / 2, _
                     tBMPs.bmWidth, tBMPs.bmHeight, _
           lhDCs, 0, 0, vbSrcPaint
    DeleteDC lhDCs
    DeleteDC lhDCm
End Sub

Public Sub Background(ByVal hDestDC As Long, Wallpaper As IPictureDisp, ByVal WindowW As Integer, ByVal WindowH As Integer)
    Dim hBMP   As Long
    Dim hBrush As Long
    Dim tBMP   As BITMAP
    Dim rcRect As RECT
    
    GetObjectAPI Wallpaper.handle, Len(tBMP), tBMP
    hBMP = CopyImage(Wallpaper.handle, ByVal 0&, tBMP.bmWidth, tBMP.bmHeight, ByVal 0&)
    SetRect rcRect, 0, 0, WindowW, WindowH
    
    hBrush = CreatePatternBrush(hBMP)
    FillRect hDestDC, rcRect, hBrush
    DeleteObject hBrush
    
    DeleteObject hBMP
End Sub

Public Sub RefreshWindow(hwnd As Long)
    Dim rcRect As RECT
    
    GetClientRect hwnd, rcRect
    InvalidateRect hwnd, rcRect, False
End Sub

Public Function Rads(ByVal deg As Single) As Single
    Rads = deg * PI / 180
End Function

Public Function PI() As Single
    PI = Atn(1) * 4
End Function

Public Function Generate_Random_Number(ByVal lowerbound As Integer, upperbound As Integer) As Collection
    Dim i       As Integer
    Dim Rnd_Val As Long
    Dim Cntr    As Integer
    Dim IsExist As Boolean
    Dim Data    As New Collection
    
    Call Randomize
    Do While Cntr < Abs(upperbound - lowerbound) + 1
        IsExist = False
        Rnd_Val = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
        
        For i = 1 To Cntr
            If Data(i) = Rnd_Val Then
                IsExist = True
                Exit For
            End If
        Next i
        If Not IsExist Then
            Data.Add Rnd_Val
            Cntr = Cntr + 1
        End If
    Loop
    
    Set Generate_Random_Number = Data
End Function

Public Function Random_Number(ByVal lowerbound As Integer, upperbound) As Integer
    Call Randomize
    
    Random_Number = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function

Private Sub DefaultSettings()
    With Setting
        .BkFileLoc = ""
        .Deck = 0
        .Difficulty = 1
        .FallType = 0
        .Opponents = 1
        .PlayerName(0) = "Juan"
        .PlayerName(1) = "Pedro"
        .PlayerName(2) = "Maria"
        .PlayerName(3) = "Piso"
        .ShowTrail = 1
        .SortMode = 0
        .WinnerSelAni = 1
        .BkColor = &H8000&
        .Speed = 1
        .MaxCard = 5
        .BounceDistX = 5
        .BounceDistY = 5
        .BounceSpeedX = 2
        .BounceSpeedY = 2
        .ScatterSpeedX = 2
        .ScatterSpeedY = 2
        .SpinDist = 6
        .WindType = 0
    End With
End Sub

Public Sub OpenSettings()
    Dim Filename As String
    
    On Error GoTo OpenErr
    
    Filename = App.Path & "\setting.dat"
    
    If Dir$(Filename) <> "" Then
        Dim InFile As Integer
        
        InFile = FreeFile
        Open Filename For Input As InFile
            With Setting
                Input #InFile, .BkFileLoc
                Input #InFile, .Deck
                Input #InFile, .Difficulty
                Input #InFile, .FallType
                Input #InFile, .Opponents
                Input #InFile, .PlayerName(0)
                Input #InFile, .PlayerName(1)
                Input #InFile, .PlayerName(2)
                Input #InFile, .PlayerName(3)
                Input #InFile, .ShowTrail
                Input #InFile, .SortMode
                Input #InFile, .WinnerSelAni
                Input #InFile, .BkColor
                Input #InFile, .Speed
                Input #InFile, .MaxCard
                Input #InFile, .BounceDistX
                Input #InFile, .BounceDistY
                Input #InFile, .BounceSpeedX
                Input #InFile, .BounceSpeedY
                Input #InFile, .ScatterSpeedX
                Input #InFile, .ScatterSpeedY
                Input #InFile, .SpinDist
                Input #InFile, .WindType
            End With
        Close InFile
    Else
        ' if setting.dat does not exist then use default settings
        Call DefaultSettings
    End If
    Exit Sub

OpenErr:
    If InFile <> 0 Then Close InFile
    MsgBox Err.Description, vbOKOnly & vbCritical, "Error"
End Sub

Public Sub SaveSettings()
    Dim Filename As String
    Dim i        As Integer
    Dim InFile   As Integer
        
    On Error GoTo SaveErr
    
    Filename = App.Path & "\setting.dat"

    InFile = FreeFile
    Open Filename For Output As InFile
        With Setting
            Write #InFile, .BkFileLoc
            Write #InFile, .Deck
            Write #InFile, .Difficulty
            Write #InFile, .FallType
            Write #InFile, .Opponents
            Write #InFile, .PlayerName(0)
            Write #InFile, .PlayerName(1)
            Write #InFile, .PlayerName(2)
            Write #InFile, .PlayerName(3)
            Write #InFile, .ShowTrail
            Write #InFile, .SortMode
            Write #InFile, .WinnerSelAni
            Write #InFile, .BkColor
            Write #InFile, .Speed
            Write #InFile, .MaxCard
            Write #InFile, .BounceDistX
            Write #InFile, .BounceDistY
            Write #InFile, .BounceSpeedX
            Write #InFile, .BounceSpeedY
            Write #InFile, .ScatterSpeedX
            Write #InFile, .ScatterSpeedY
            Write #InFile, .SpinDist
            Write #InFile, .WindType
        End With
    Close InFile
    Exit Sub
    
SaveErr:
    If InFile <> 0 Then Close InFile
    MsgBox Err.Description, vbOKOnly Or vbCritical, "Error"
End Sub
