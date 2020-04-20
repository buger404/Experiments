VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecognizeImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub RenderRecognizeImg(ByVal DC As Long, ByVal Width As Long, ByVal Height As Long)
    Dim Color As Long, X As Long, Y As Long
    Dim Buff(3) As Byte
    
    For X = 1 To Width
        For Y = 1 To Height
            Color = GetPixel(DC, X, Y)
            CopyMemory Buff(0), Color, 4
            Color = ConverColor(Buff)
            SetPixel DC, X, Y, Color
        Next
        If X Mod 100 = 0 Then
            MainWindow.Outputs "Rendered " & Int(X / Width * 1000) / 10 & "% ..."
            DoEvents
        End If
    Next
End Sub
Public Function ConverColor(Buff() As Byte) As Long
    Dim I As Integer, T As Long, temp As Long
    For I = 0 To UBound(Buff)
        temp = Buff(I): temp = temp * temp
        T = T + temp
    Next
    T = T / (255 ^ 2 + 255 ^ 2 + 255 ^ 2) * 255
    ConverColor = RGB(T, T, T)
End Function

