Attribute VB_Name = "modUno"
Option Explicit

Public Const Signature = "UNO"
Public Const Version = "2.0"

Public Const LEVEL_EASY = 0
Public Const LEVEL_NORMAL = 1
Public Const LEVEL_HARD = 2

Public Const CardW = 60 ' Card Width
Public Const CardH = 90 ' Card Height

Public Const CNT_Rank = 0
Public Const CNT_Suit = 1

Public Type GAME_SETTING
    Opponents     As Integer
    PlayerName(3) As String
    BkFileLoc     As String
    BkPicture     As New StdPicture
    Deck          As Integer
    Difficulty    As Integer
    ShowTrail     As Integer
    SortMode      As Integer
    FallType      As Integer
    WinnerSelAni  As Integer
    BkColor       As Long
    Speed         As Integer
    MaxCard       As Integer
    BounceDistX   As Integer
    BounceDistY   As Integer
    BounceSpeedX  As Integer
    BounceSpeedY  As Integer
    ScatterSpeedX As Integer
    ScatterSpeedY As Integer
    SpinDist      As Integer
    WindType      As Integer
End Type

Public Setting As GAME_SETTING

Public Function CountCard(ByRef PlayerCards As Object, ByVal Compare As Integer, ByVal OpCount As Integer)
    Dim Sum  As Integer
    Dim Card As Object
    
    For Each Card In PlayerCards
        If OpCount = 0 Then
            ' Rank only
            If Compare = Card.Rank Then Sum = Sum + 1
        Else
            ' Suit only
            If Card.Rank < uno_RCWild Then
                If Compare = Card.Suit Then Sum = Sum + 1
            End If
        End If
    Next Card
    
    CountCard = Sum
End Function

Public Function GetSuit(ByRef PlayerCards As Object, ByVal Suit As Integer) As Collection
    Dim Card As Object
    Dim temp As Object
    
    Set GetSuit = New Collection
    For Each Card In PlayerCards
        If Card.Rank < 13 Then
            If Card.Suit = Suit Then
                Set temp = New clsCardInfo
                
                temp.Data = Card.Data
                temp.Rank = Card.Rank
                temp.Points = Card.Points
                temp.Suit = Card.Suit
                temp.Tag = Card.Tag
                GetSuit.Add temp
            End If
        End If
    Next Card
End Function

Public Function GetWildCard(ByRef PlayerCards As Object) As Collection
    Dim Card As Object
    Dim temp As Object
    
    Set GetWildCard = New Collection
    For Each Card In PlayerCards
        If Card.Rank > 12 Then
            Set temp = New clsCardInfo
                    
            temp.Data = Card.Data
            temp.Rank = Card.Rank
            temp.Points = Card.Points
            temp.Suit = Card.Suit
            temp.Tag = Card.Tag
            GetWildCard.Add temp
        End If
    Next Card
End Function

Public Sub Sort(ByRef Data As Object)
    Dim i As Integer
    Dim j As Integer
    Dim z As Integer
    Dim S As String
    Dim b As StdPicture
    
    For i = 1 To Data.Count
        For j = 1 To Data.Count - 1
            If Data(j).Rank > Data(j + 1).Rank Then
                S = Data(j).Data
                Data(j).Data = Data(j + 1).Data
                Data(j + 1).Data = S
                
                z = Data(j).Rank
                Data(j).Rank = Data(j + 1).Rank
                Data(j + 1).Rank = z
                
                z = Data(j).Points
                Data(j).Points = Data(j + 1).Points
                Data(j + 1).Points = z
                
                z = Data(j).Suit
                Data(j).Suit = Data(j + 1).Suit
                Data(j + 1).Suit = z
                
                S = Data(j).Tag
                Data(j).Tag = Data(j + 1).Tag
                Data(j + 1).Tag = S
            End If
        Next j
    Next i
End Sub

Public Function SearchCard(ByRef PlayerCards As Object, ByVal Rank As Integer, ByVal Suit As Integer, ByVal WhatRank As Integer) As Integer
    Dim CurPos As Integer
    Dim Card   As Object
    
    CurPos = -1

    For Each Card In PlayerCards
        If Card.Suit = Suit Then
            If Card.Rank = WhatRank Then
                CurPos = Card.Index
            End If
        End If
        
        If CurPos <> -1 Then Exit For
    Next Card
    
    If CurPos = -1 Then
        If Rank = WhatRank Then
            For Each Card In PlayerCards
                If Card.Rank = WhatRank Then
                    CurPos = Card.Index
                End If
            
                If CurPos <> -1 Then Exit For
            Next Card
        End If
    End If
    
    SearchCard = CurPos
End Function

Public Function SearchMove(ByRef PlayerCards As Object, ByVal Rank As Integer, ByVal Suit As Integer) As Integer
    Dim CurPos As Integer
    Dim Card   As Object
    
    CurPos = -1
    
    For Each Card In PlayerCards
        ' wild card and draw four are both not included in the search
        If (Card.Rank <> uno_RCWild) And _
           (Card.Suit <> uno_RCDrawFour) Then
            If (Card.Rank = Rank) Or (Card.Suit = Suit) Then
                CurPos = Card.Index
            End If
        End If
        
        If CurPos <> -1 Then Exit For
    Next Card
    
    SearchMove = CurPos
End Function

Public Function SearchWildCard(ByRef PlayerCards As Object, ByVal Rank As Integer) As Integer
    Dim CurPos As Integer
    Dim Card   As Object
    
    CurPos = -1
    
    For Each Card In PlayerCards
        If Card.Rank = uno_RCWild Then
            CurPos = Card.Index
        End If
        
        If CurPos <> -1 Then Exit For
    Next Card
    
    SearchWildCard = CurPos
End Function

Public Function SearchDrawTwo(ByRef PlayerCards As Object, Rank As Integer, Suit As Integer) As Integer
    Dim CurPos As Integer
    Dim Card   As Object
    
    CurPos = -1

    For Each Card In PlayerCards
        If (Card.Rank = uno_RCDrawTwo) Then
            If (Card.Rank = Rank) Or (Card.Suit = Suit) Then
                CurPos = Card.Index
            End If
        End If
        
        If CurPos <> -1 Then Exit For
    Next Card
    
    SearchDrawTwo = CurPos
End Function

Public Function SearchDrawFour(ByRef PlayerCards As Object, Rank As Integer, Suit As Integer) As Integer
    Dim CurPos As Integer
    Dim Card  As Object
    
    CurPos = -1
    
    If CountCard(PlayerCards, Suit, 1) = 0 Then
        For Each Card In PlayerCards
            If Card.Rank = uno_RCDrawFour Then
                CurPos = Card.Index
            End If
                
            If CurPos <> -1 Then Exit For
        Next Card
    End If
    
    SearchDrawFour = CurPos
End Function

Public Function GetLargestRankSuit(PlayerCards As Object, ByVal Suit As Integer)
    Dim CurPos As Integer
    Dim Card   As Object
    Dim l      As Integer
    
    CurPos = -1: l = -1
    
    For Each Card In PlayerCards
        If (Card.Rank >= uno_RCZero) And (Card.Rank <= uno_RCNine) Then
            If Card.Suit = Suit Then
                If l = -1 Then
                    l = Card.Rank
                ElseIf l < Card.Rank Then
                    l = Card.Rank
                End If
            End If
        End If
    Next Card
    
    For Each Card In PlayerCards
        If (Card.Rank = l) And (Card.Suit = Suit) Then
            CurPos = Card.Index
        End If
            
        If CurPos <> -1 Then Exit For
    Next Card
    
    GetLargestRankSuit = CurPos
End Function

Public Function GetLargestRank(PlayerCards As Object) As Integer
    Dim cb   As Integer
    Dim cr   As Integer
    Dim cg   As Integer
    Dim cy   As Integer
    Dim br   As Integer
    Dim gy   As Integer
    Dim lbr  As Integer
    Dim lgy  As Integer
    Dim Card As Object
    
    For Each Card In PlayerCards
        If Card.Rank < 13 Then
            If Card.Suit = 0 Then     ' blue
                cb = cb + Card.Rank
            ElseIf Card.Suit = 1 Then ' red
                cr = cr + Card.Rank
            ElseIf Card.Suit = 2 Then ' green
                cg = cg + Card.Rank
            Else                      ' yellow
                cy = cy + Card.Rank
            End If
        End If
    Next Card
    
    If cb < cr Then
        br = cr
        lbr = 1 ' red
    Else
        br = cb
        lbr = 0 ' blue
    End If
    
    If cg < cy Then
        gy = cy
        lgy = 3 ' yellow
    Else
        gy = cg
        lgy = 2 ' green
    End If
    
    GetLargestRank = IIf(br < gy, lgy, lbr)
End Function

Public Function GetLargestSuit(PlayerCards As Object) As Integer
    Dim cb  As Integer
    Dim cr  As Integer
    Dim cg  As Integer
    Dim cy  As Integer
    Dim br  As Integer
    Dim gy  As Integer
    Dim lbr As Integer
    Dim lgy As Integer
    
    cb = CountCard(PlayerCards, 0, 1) ' blue
    cr = CountCard(PlayerCards, 1, 1) ' red
    cg = CountCard(PlayerCards, 2, 1) ' green
    cy = CountCard(PlayerCards, 3, 1) ' yellow
    
    If cb < cr Then
        br = cr
        lbr = 1 ' red
    Else
        br = cb
        lbr = 0 ' blue
    End If
    
    If cg < cy Then
        gy = cy
        lgy = 3 ' yellow
    Else
        gy = cg
        lgy = 2 ' green
    End If
    
    GetLargestSuit = IIf(br < gy, lgy, lbr)
End Function

Public Sub ChangeSort(ByVal SelSort As Integer)
    With frmMain
        Uno.SetSortMode = SelSort
        Uno.SortCards .crdPlayerOne
        
        If .crdPlayerTwo(0).Face = uno_FCUp Then
            Uno.SortCards .crdPlayerTwo
        End If
            
        If .crdPlayerThree(0).Face = uno_FCUp Then
            Uno.SortCards .crdPlayerThree
        End If
            
        If .crdPlayerFour(0).Face = uno_FCUp Then
            Uno.SortCards .crdPlayerFour
        End If
        
        Select Case SelSort + 1
        Case Is = 1
            .imcSortMode.ToolTipText = _
                "Blue, Red, Green, Yellow, Wild Card, Draw Four"
        Case Is = 2
            .imcSortMode.ToolTipText = _
                "Red, Green, Yellow, Wild Card, Draw Four, Blue"
        Case Is = 3
            .imcSortMode.ToolTipText = _
                "Green, Yellow, Wild Card, Draw Four, Blue, Red"
        Case Is = 4
            .imcSortMode.ToolTipText = _
                "Yellow, Wild Card, Draw Four, Blue, Red, Green"
        Case Is = 5
            .imcSortMode.ToolTipText = _
                "Wild Card, Draw Four, Blue, Red, Green, Yellow"
        End Select
    End With
End Sub
