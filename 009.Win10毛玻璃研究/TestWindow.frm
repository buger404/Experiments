VERSION 5.00
Begin VB.Form TestWindow 
   BorderStyle     =   0  'None
   Caption         =   "TestWindow"
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "TestWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim o As WinShadow
Private Sub Form_Load()
    SetWindowLongA Me.hwnd, GWL_EXSTYLE, GetWindowLongA(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Call Win10Blur(Me.hwnd, argb(200, 255, 255, 255))
    
    Set o = New WinShadow
    With o
        If .Shadow(Me) Then
            .Color = RGB(0, 0, 0)
            .Depth = 16
            .Transparency = 16
        End If
    End With
End Sub
