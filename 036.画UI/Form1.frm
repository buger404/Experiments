VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3792
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6576
   LinkTopic       =   "Form1"
   ScaleHeight     =   316
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   548
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Points
    xMin As Long
    yMin As Long
    xMax As Long
    yMax As Long
End Type
Dim P As Points
Dim R() As Points
Dim lX As Long, lY As Long

Private Sub Form_DblClick()
    Me.Cls
    ReDim R(0)
End Sub

Private Sub Form_Load()
    ReDim R(0)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button > 0 Then
        If lX = 0 And lY = 0 Then lX = X: lY = Y: P.xMin = X: P.xMax = X: P.yMin = Y: P.yMax = Y
        Me.Line (lX, lY)-(X, Y)
        If X < P.xMin Then P.xMin = X
        If X > P.xMax Then P.xMax = X
        If Y < P.yMin Then P.yMin = Y
        If Y > P.yMax Then P.yMax = Y
        lX = X: lY = Y
    End If
End Sub
Function Suit(v1, v2, Optional er As Single = 0.3) As Boolean
    Dim f As Single
    f = Abs(v1 - v2) / IIf(v1 > v2, v1, v2)
    Suit = (f > 1 - er And f < 1 + er)
    Suit = Abs(v1 - v2) < 20 'IIf(v1 > v2, v1, v2) * er
    Debug.Print Now, Abs(v1 - v2), IIf(v1 > v2, v1, v2) * er
End Function
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lX = 0: lY = 0
    'Me.Cls
    Dim Xn As Long, Yn As Long, W As Long, H As Long
    If Button = 1 Then
        Xn = P.xMin
        Yn = P.yMin
        W = P.xMax - P.xMin
        H = P.yMax - P.yMin
        For i = 1 To UBound(R)
            If Suit(R(i).xMin, Xn) Then
                If Suit(R(i).xMax, W) Then W = R(i).xMax
                Xn = R(i).xMin
            End If
            If Suit(R(i).yMin, Yn) Then
                If Suit(R(i).yMax, H) Then H = R(i).yMax
                Yn = R(i).yMin
            End If
        Next
        Me.Line (Xn, Yn)-(Xn + W, Yn + H), RGB(255, 0, 0), B
    Else
        Xn = (P.xMin + P.xMax) / 2
        Yn = (P.yMin + P.yMax) / 2
        W = Sqr((P.xMax - P.xMin) ^ 2 + (P.yMax - P.yMin) ^ 2) / 3.14
        Me.Circle (Xn, Yn), W, RGB(255, 0, 0)
    End If
    
    P.xMin = Xn: P.xMin = Yn: P.xMax = W: P.yMax = H
    ReDim Preserve R(UBound(R) + 1)
    R(UBound(R)) = P
End Sub
