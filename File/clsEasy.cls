VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEasy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Move(PlayerCards As Object, ByVal Rank As Integer, ByVal Suit As Integer)
    Dim CurCard As Integer
    
    CurCard = -1
    
    CurCard = SearchMove(PlayerCards, Rank, Suit)
    
    If CurCard = -1 Then
        CurCard = SearchWildCard(PlayerCards, Rank)
    End If
    
    If CurCard = -1 Then
        CurCard = SearchDrawFour(PlayerCards, Rank, Suit)
    End If
    
    Move = CurCard
End Function
