VERSION 5.00
Begin VB.Form MainWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "贝塞尔曲线"
   ClientHeight    =   5040
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   6732
   DrawWidth       =   3
   Icon            =   "MainWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Params 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00404040&
      Height          =   312
      Index           =   3
      Left            =   3432
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "1"
      Top             =   4488
      Width           =   684
   End
   Begin VB.TextBox Params 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00404040&
      Height          =   312
      Index           =   2
      Left            =   2616
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "0.17"
      Top             =   4488
      Width           =   684
   End
   Begin VB.TextBox Params 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00404040&
      Height          =   312
      Index           =   1
      Left            =   1800
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "0"
      Top             =   4488
      Width           =   684
   End
   Begin VB.TextBox Params 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00404040&
      Height          =   312
      Index           =   0
      Left            =   984
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "0.77"
      Top             =   4476
      Width           =   684
   End
   Begin VB.Label LinePad 
      Alignment       =   2  'Center
      BackColor       =   &H00FF3E64&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   276
      Left            =   6192
      TabIndex        =   6
      Top             =   2544
      Width           =   276
   End
   Begin VB.Line SetLine 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   80
      X2              =   386
      Y1              =   288
      Y2              =   288
   End
   Begin VB.Label ColorPad 
      Alignment       =   2  'Center
      BackColor       =   &H00CEDB1A&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   276
      Left            =   5760
      TabIndex        =   5
      Top             =   2544
      Width           =   276
   End
   Begin VB.Line XLine 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   80
      X2              =   386
      Y1              =   330
      Y2              =   330
   End
   Begin VB.Line YLine 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   80
      X2              =   80
      Y1              =   24
      Y2              =   330
   End
   Begin VB.Label StartBtn 
      Alignment       =   2  'Center
      BackColor       =   &H00CEDB1A&
      Caption         =   "Launch"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   276
      Left            =   4824
      TabIndex        =   0
      Top             =   4512
      Width           =   1236
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()
    'Me.Caption = "Experiment " & Split(Split(App.Path, ".")(0), "Experiments\")(1)
    Dim t As Single, y As Single
    t = Timer
    For i = 1 To 100000
        y = HtmlCubic(Rnd, Rnd, Rnd, Rnd, Rnd)
    Next
    MsgBox "十万次：" & Timer - t & "s"
End Sub
Private Sub Test()
    'Test Code
    Me.Cls
    Me.Refresh
    
    Dim x As Single, y As Single, pro As Single
    Dim w As Long, h As Long, ly As Single, lx As Single, cx As Single, pr As Single
    w = Abs(XLine.x2 - XLine.x1)
    h = Abs(YLine.Y2 - YLine.Y1)
    w = h
    ly = -1
    
    Dim t As Long
    t = GetTickCount
    
    Do While GetTickCount - t <= 3000
        pro = (GetTickCount - t) / 3000
        x = XLine.x1 + (XLine.x2 - XLine.x1) * pro
        y = HtmlCubic(pro, Val(Params(0).Text), Val(Params(1).Text), Val(Params(2).Text), Val(Params(3).Text))
        pr = y
        cx = x
        If y > 1 Then y = 1
        SetLine.BorderColor = RGB(160 + 90 * y, 160 - 160 * y, 160 - 160 * y)
        If ly <> -1 Then Me.Line (lx, YLine.Y2 - h * ly)-(cx, YLine.Y2 - h * y), SetLine.BorderColor
        ly = y: lx = x
        ColorPad.Top = YLine.Y2 - h * pr - ColorPad.Height / 2
        LinePad.Top = YLine.Y2 - h * pro - LinePad.Height / 2
        ColorPad.BackColor = RGB(255 - 255 * pr, 255 - (255 - 176) * pr, 255 - (255 - 240) * pr)
        SetLine.x1 = cx + 3: SetLine.Y1 = YLine.Y2 - h * pr
        SetLine.Y2 = SetLine.Y1: SetLine.x2 = ColorPad.Left - 3
        StartBtn.Caption = Format(Int(y * 10000) / 10000, "0.0000")
        DoEvents
    Loop

End Sub
Private Sub StartBtn_Click()
    Call Test
End Sub
Function Cubic(t As Single, arg0 As Single, arg1 As Single, arg2 As Single, arg3 As Single) As Single
    'Formula:B(t)=P_0(1-t)^3+3P_1t(1-t)^2+3P_2t^2(1-t)+P_3t^3
    'Attention:all the args must in this area (0~1)
    Cubic = (arg0 * ((1 - t) ^ 3)) + (3 * arg1 * t * ((1 - t) ^ 2)) + (3 * arg2 * (t ^ 2) * (1 - t)) + (arg3 * (t ^ 3))
End Function
Function HtmlCubic(t As Single, p0 As Single, p1 As Single, p2 As Single, p3 As Single) As Single
    'Formula:B(t)=P_0(1-t)^3+3P_1t(1-t)^2+3P_2t^2(1-t)+P_3t^3
    'Attention:all the args must in this area (0~1)
    'Made by 极光
    
    If (t = 0) Then HtmlCubic = 0: Exit Function
    If (t = 1) Then HtmlCubic = 1: Exit Function
    
    Dim a As Single, b As Single, c As Single, d As Single
    a = 1 - 3 * p2 + 3 * p0
    b = 3 * p2 - 6 * p0
    c = 3 * p0
    d = -t
    
    Dim x1 As Single, x2 As Single, tt As Single, fx1 As Single, fx2 As Single, fx0 As Single
    x1 = 0: x2 = 1: fx1 = d: fx2 = a + b + c + d
    fx0 = 1
    
    Do While Abs(fx0) >= 0.00001
        tt = (x1 + x2) / 2
        fx0 = ((a * tt + b) * tt + c) * tt + d
        If fx0 * fx1 < 0 Then
            x2 = tt: fx2 = fx0
        Else
            x1 = tt: fx1 = fx0
        End If
    Loop
    
    Dim z As Single
    z = 1 - tt
    HtmlCubic = (3 * p1 * tt * z * z + 3 * p3 * tt * tt * z + tt * tt * tt)
End Function
