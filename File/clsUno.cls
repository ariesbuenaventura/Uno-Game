VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type CARD_INFO
    Data    As String
    Rank    As Integer
    Points   As Integer
    Suit    As Integer
    Tag     As String
    Picture As New StdPicture
End Type

Public Stock As New Collection
Public Waste As New Collection

Dim Easy      As New clsEasy
Dim Normal    As New clsNormal
Dim Difficult As New clsDifficult

Dim SortMode  As Integer
Dim LevelMode As Integer

Public Function IsMoveValid(ByRef PlayerCards As Object, ByVal SC As Integer, ByVal LR As Integer, ByVal LS As Integer)
    Dim bTemp   As Boolean
    Dim SelRank As Integer
    Dim SelSuit As Integer
    
    SelRank = PlayerCards(SC).Rank
    SelSuit = PlayerCards(SC).Suit
    
    If (SelRank <> 13) And (SelRank <> 14) Then
        If Not bTemp Then bTemp = IIf(SelRank = LR, True, False)
        If Not bTemp Then bTemp = IIf(SelSuit = LS, True, False)
    Else
        If Not bTemp Then bTemp = IIf(SelRank = 13, True, False)
        If Not bTemp Then bTemp = IIf(CountCard(PlayerCards, LS, CNT_Suit) = 0, True, False)
    End If
    
    IsMoveValid = bTemp
End Function

Public Property Get GetSortMode() As Integer
    GetSortMode = SortMode
End Property

Public Property Let SetSortMode(ByVal Mode As Integer)
    SortMode = Mode
End Property

Public Sub SortCards(ByRef PlayerCards As Object)
    Dim i        As Integer
    Dim j        As Integer
    Dim Card     As Object
    Dim Temp     As New Collection
    Dim allCards As New Collection

    For i = 0 To 4
        If i = Abs(GetSortMode - 4) Mod 5 Then
            Set Temp = GetWildCard(PlayerCards)
        Else
            Set Temp = GetSuit(PlayerCards, (j + GetSortMode) Mod 4)
            j = j + 1
        End If
        
        Call Sort(Temp)
        
        For Each Card In Temp
            allCards.Add Card
        Next Card
    Next i
    
    Dim CI() As CARD_INFO
    
    ReDim CI(PlayerCards.Count - 1) As CARD_INFO
        
    For i = LBound(CI()) To UBound(CI())
        CI(i).Data = PlayerCards(i).Data
        CI(i).Rank = PlayerCards(i).Rank
        CI(i).Points = PlayerCards(i).Points
        CI(i).Suit = PlayerCards(i).Suit
        CI(i).Tag = PlayerCards(i).Tag
        Set CI(i).Picture = PlayerCards(i).Picture
    Next i
    
    For i = 1 To allCards.Count
        If (PlayerCards(i - 1).Rank <> allCards(i).Rank) Or _
           (PlayerCards(i - 1).Suit <> allCards(i).Suit) Then
            For j = LBound(CI()) To UBound(CI())
                If (allCards(i).Rank = CI(j).Rank) And _
                   (allCards(i).Suit = CI(j).Suit) Then
                    PlayerCards(i - 1).AutoUpdate = False
                    PlayerCards(i - 1).Data = CI(j).Data
                    PlayerCards(i - 1).Rank = CI(j).Rank
                    PlayerCards(i - 1).Points = CI(j).Points
                    PlayerCards(i - 1).Suit = CI(j).Suit
                    PlayerCards(i - 1).Tag = CI(j).Tag
                    Set PlayerCards(i - 1).Picture = CI(j).Picture
                    PlayerCards(i - 1).AutoUpdate = True
                    Exit For
                End If
            Next j
        End If
    Next i
End Sub

Public Sub Pick(ByRef PlayerCards As Object)
    If Stock.Count > 0 Then
        Dim CurIndex As Integer
        
        Load PlayerCards(PlayerCards.Count)
        
        CurIndex = PlayerCards.Count - 1
        PlayerCards(CurIndex).Rank = Stock(Stock.Count).Rank
        PlayerCards(CurIndex).Suit = Stock(Stock.Count).Suit
        PlayerCards(CurIndex).Visible = True
        Stock.Remove Stock.Count
        
        For CurIndex = 0 To PlayerCards.Count - 1
            PlayerCards(CurIndex).TabIndex = CurIndex
        Next CurIndex
    End If
End Sub

Public Sub Throw(ByRef PlayerCards As Object, ByVal SC As Integer)
    Dim i    As Integer
    Dim Temp As Object
    
    Set Temp = New clsCardInfo
    Temp.Rank = PlayerCards(SC).Rank
    Temp.Suit = PlayerCards(SC).Suit
    Waste.Add Temp
    
    For i = SC To PlayerCards.Count - 2
        PlayerCards(i).AutoUpdate = False
        PlayerCards(i).Data = PlayerCards(i + 1).Data
        PlayerCards(i).Rank = PlayerCards(i + 1).Rank
        PlayerCards(i).Points = PlayerCards(i + 1).Points
        PlayerCards(i).Suit = PlayerCards(i + 1).Suit
        PlayerCards(i).Tag = PlayerCards(i + 1).Tag
        Set PlayerCards(i).Picture = PlayerCards(i + 1).Picture
        PlayerCards(i).AutoUpdate = True
    Next i
    
    If PlayerCards.Count > 1 Then
        Unload PlayerCards(PlayerCards.Count - 1)
    Else
        PlayerCards(0).Visible = False
    End If
End Sub

Public Sub Shuffle()
    Dim i    As Integer
    Dim Card As Object
    Dim Temp As New Collection
    
    Set Stock = Nothing
    Set Waste = Nothing
    Set Temp = Generate_Random_Number(1, 108)
    
    For i = 1 To Temp.Count
        Set Card = New clsCardInfo
        Card.Rank = Temp(i) Mod 27
        If Card.Rank > 18 Then
            Select Case Card.Rank
            Case 19 To 20 ' Draw Two
                Card.Rank = 10
            Case 21 To 22 ' Reverse
                Card.Rank = 11
            Case 23 To 24 ' Skip
                Card.Rank = 12
            Case 25       ' Wild
                Card.Rank = 13
            Case 26       ' Draw Four
                Card.Rank = 14
            End Select
        Else
            Card.Rank = Card.Rank Mod 10
        End If
        
        Card.Suit = Temp(i) Mod 4
        Stock.Add Card
    Next i
End Sub

Public Property Let SetLevelMode(ByVal New_Level As Integer)
    LevelMode = New_Level
End Property

Public Property Get GetLevelMode() As Integer
    GetLevelMode = LevelMode
End Property

Public Function AI(ByRef PlayerCards As Object, ByVal Rank As Integer, ByVal Suit As Integer) As Integer
    Select Case GetLevelMode
    Case Is = LEVEL_EASY
        AI = Easy.Move(PlayerCards, Rank, Suit)
    Case Is = LEVEL_NORMAL
        AI = Normal.Move(PlayerCards, Rank, Suit)
    Case Is = LEVEL_HARD
        AI = Difficult.Move(PlayerCards, Rank, Suit)
    End Select
End Function

Public Function TotalPoints(ByRef PlayerCards As Object) As Integer
    Dim Card  As Object
    Dim Total As Integer
    
    For Each Card In PlayerCards
        If Card.Visible Then
            Total = Total + Card.Points
        End If
    Next Card
    
    TotalPoints = Total
End Function

Public Function TotalCards(ByRef PlayerCards As Object) As Integer
    Dim Card  As Object
    Dim Total As Integer
    
    For Each Card In PlayerCards
        If Card.Visible Then
            Total = Total + 1
        End If
    Next Card
    
    TotalCards = Total
End Function

Private Sub Class_Initialize()
    SortMode = 0
    LevelMode = LEVEL_EASY
End Sub
