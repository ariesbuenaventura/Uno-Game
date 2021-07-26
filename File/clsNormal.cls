VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Move(PlayerCards As Object, ByVal Rank As Integer, ByVal Suit As Integer)
    Dim CurCard As Integer
    
    CurCard = -1
    
    CurCard = SearchWildCard(PlayerCards, Rank)
    If CurCard = -1 Then CurCard = SearchDrawTwo(PlayerCards, Rank, Suit)
    If CurCard = -1 Then CurCard = SearchMove(PlayerCards, Rank, Suit)
    If CurCard = -1 Then CurCard = SearchWildCard(PlayerCards, Rank)
    If CurCard = -1 Then CurCard = SearchDrawFour(PlayerCards, Rank, Suit)

    If CurCard <> -1 Then
        If (PlayerCards(CurCard).Rank = uno_RCWild) Or _
           (PlayerCards(CurCard).Rank = uno_RCDrawFour) Then
           
            Dim Temp As Integer
    
            If PlayerCards.Count - 1 = 1 Then
                Temp = Int(4 * Rnd)
            Else
                Temp = GetLargestSuit(PlayerCards)
            End If
    
            PlayerCards(CurCard).Suit = Temp
        End If
    End If
    
    Move = CurCard
End Function

