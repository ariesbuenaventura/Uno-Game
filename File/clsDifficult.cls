VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDifficult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Move(PlayerCards As Object, ByVal Rank As Integer, ByVal Suit As Integer)
    Dim CurCard As Integer
    
    CurCard = -1
    
    If CurCard = -1 Then CurCard = SearchDrawFour(PlayerCards, Rank, Suit)
    If CurCard = -1 Then CurCard = SearchDrawTwo(PlayerCards, Rank, Suit)
    If CurCard = -1 Then CurCard = SearchCard(PlayerCards, Rank, Suit, uno_RCSkip)
    If CurCard = -1 Then CurCard = SearchCard(PlayerCards, Rank, Suit, uno_RCReverse)
    If CurCard = -1 Then CurCard = GetLargestRankSuit(PlayerCards, Suit)
    If CurCard = -1 Then CurCard = SearchMove(PlayerCards, Rank, Suit)
    If CurCard = -1 Then CurCard = SearchWildCard(PlayerCards, Rank)
    
    If CurCard <> -1 Then
        If (PlayerCards(CurCard).Rank = uno_RCWild) Or _
           (PlayerCards(CurCard).Rank = uno_RCDrawFour) Then
           
            Dim temp As Integer
            
            temp = -1
            
            If PlayerCards.Count - 1 = 0 Then
                temp = Int(4 * Rnd)
            Else
                Dim cc As Integer
                
                cc = CountCard(PlayerCards, 0, 1)                ' count blue
                If cc = 0 Then cc = CountCard(PlayerCards, 1, 1) ' count red
                If cc = 0 Then cc = CountCard(PlayerCards, 2, 1) ' count green
                If cc = 0 Then cc = CountCard(PlayerCards, 3, 1) ' count yellow
                
                If cc > 0 Then
                    If (CountCard(PlayerCards, 1, 1) = cc) Or (CountCard(PlayerCards, 0, 1) = 0) Then              ' blue
                        If (CountCard(PlayerCards, 1, 1) = cc) Or (CountCard(PlayerCards, 1, 1) = 0) Then          ' red
                            If (CountCard(PlayerCards, 2, 1) = cc) Or (CountCard(PlayerCards, 2, 1) = 0) Then      ' green
                                If (CountCard(PlayerCards, 3, 1) = cc) Or (CountCard(PlayerCards, 1, 3) = 0) Then  ' yellow
                                    temp = GetLargestRank(PlayerCards)
                                End If
                            End If
                        End If
                    End If
                End If
                
                If temp = -1 Then
                    temp = GetLargestSuit(PlayerCards)
                End If
            End If
            
            PlayerCards(CurCard).Suit = temp
        End If
    End If
    
    Move = CurCard
End Function
