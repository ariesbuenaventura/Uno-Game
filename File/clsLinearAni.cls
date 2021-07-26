VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinearAni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public varStopAni As Boolean
Public varSpeed   As Single

Public Property Get Speed() As Single
    Speed = varSpeed
End Property

Public Property Let Speed(nSpeed As Single)
    varSpeed = nSpeed
End Property

Public Property Get StopAni() As Boolean
    StopAni = varStopAni
End Property

Public Property Let StopAni(bVal As Boolean)
    varStopAni = bVal
End Property

Public Function Linear(Card As Object, _
                       X1 As Integer, _
                       Y1 As Integer, _
                       X2 As Integer, _
                       Y2 As Integer) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim X        As Single
    Dim Y        As Single
    Dim dx       As Single
    Dim dy       As Single
    
    Static px    As Single
    Static py    As Single
    Static sdx   As Single
    Static sdy   As Single
    Static dxabs As Single
    Static dyabs As Single
    Static bInit As Boolean
    
    If Not bInit Or StopAni Then
        X = 0: Y = 0
        px = X1: py = Y1
        dx = (X2 - X1) / Speed: dy = (Y2 - Y1) / Speed
        sdx = Sgn(dx): sdy = Sgn(dy)
        dxabs = Abs(dx): dyabs = Abs(dy)
        bInit = True: Linear = False: StopAni = False
    Else
        Dim i          As Single
        Dim OldTimer   As Single
                
        Card.Parent.Enabled = False
        
        Card.ZOrder 0
        If dxabs >= dyabs Then
            Do While (i <= dxabs) And Not StopAni
                If Not StopAni Then
                    OldTimer = Timer
                    
                    Y = Y + dyabs * Speed
                    If Y >= dxabs * Speed Then
                        Y = Y - dxabs * Speed
                        py = py + sdy * Speed
                    End If
                    
                    Card.Move px, py
                    DoEvents
                    
                    i = i + 1
                    px = px + sdx * Speed
                    
                    Do While Abs(Timer - OldTimer) < 0.00001
                    Loop
                Else
                    bInit = False: Linear = True
                    Exit Do
                End If
            Loop
            
            Card.Move X2, Y2
            bInit = False: Linear = True
        Else
            Do While (i <= dyabs) And Not StopAni
                If Not StopAni Then
                    OldTimer = Timer
                    
                    X = X + dxabs * Speed
                    If X >= dyabs * Speed Then
                        X = X - dyabs * Speed
                        px = px + sdx * Speed
                    End If
                    
                    Card.Move px, py
                    DoEvents
                    
                    i = i + 1
                    py = py + sdy * Speed
                    
                    Do While Abs(Timer - OldTimer) < 0.00001
                    Loop
                Else
                    bInit = False: Linear = True
                    Exit Do
                End If
            Loop
            
            Card.Move X2, Y2
            bInit = False: Linear = True
        End If
        
        Card.Parent.Enabled = True
    End If
    Exit Function
    
ErrHandler:
    Linear = True: StopAni = True
End Function

Private Sub Class_Initialize()
    varSpeed = 1
End Sub
