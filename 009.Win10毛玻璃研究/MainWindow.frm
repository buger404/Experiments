VERSION 5.00
Begin VB.Form MainWindow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Experiments"
   ClientHeight    =   3396
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   8184
   Icon            =   "MainWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3396
   ScaleWidth      =   8184
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.HScrollBar AlphaScroll 
      Height          =   180
      Left            =   4608
      Max             =   255
      TabIndex        =   4
      Top             =   2328
      Value           =   128
      Width           =   2436
   End
   Begin VB.PictureBox ColorBox 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1548
      Left            =   4608
      Picture         =   "MainWindow.frx":000C
      ScaleHeight     =   1524
      ScaleWidth      =   2364
      TabIndex        =   3
      Top             =   672
      Width           =   2388
   End
   Begin VB.PictureBox MoveBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1692
      Left            =   720
      ScaleHeight     =   1668
      ScaleWidth      =   2556
      TabIndex        =   0
      Top             =   600
      Width           =   2580
      Begin VB.Label DateText 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "[Date]"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   216
         Left            =   984
         TabIndex        =   2
         Top             =   1176
         Width           =   576
      End
      Begin VB.Label TitleText 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Experiment"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CEDB1A&
         Height          =   216
         Left            =   816
         TabIndex        =   1
         Top             =   936
         Width           =   960
      End
      Begin VB.Image testIcon 
         Height          =   624
         Left            =   984
         Picture         =   "MainWindow.frx":125FE
         Stretch         =   -1  'True
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000006&
      X1              =   2232
      X2              =   1728
      Y1              =   2808
      Y2              =   2808
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      X1              =   1992
      X2              =   1992
      Y1              =   2256
      Y2              =   2808
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BColor(2) As Byte

Private Sub AlphaScroll_Change()
    Call Win10Blur(TestWindow.hwnd, argb(AlphaScroll.Value, BColor(0), BColor(1), BColor(2)))
End Sub

Private Sub ColorBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 0 Then Exit Sub
    Dim C As Long
    C = ColorBox.Point(x, y)
    CopyMemory BColor(0), C, 3
    
    Call Win10Blur(TestWindow.hwnd, argb(AlphaScroll.Value, BColor(0), BColor(1), BColor(2)))
End Sub

Private Sub Form_Load()
    Me.Caption = "Experiment " & Split(Split(App.Path, ".")(0), "Experiments\")(1)
    TitleText.Caption = Me.Caption
    DateText.Caption = Now
    Set Me.Icon = testIcon.Picture
    
    TestWindow.Show
End Sub
Private Sub Outputs(Text As String)
    outputBox.Text = outputBox.Text & Now & "   " & Text & vbCrLf
    outputBox.SelLength = 1
    outputBox.SelStart = Len(outputBox.Text)
End Sub
Private Sub Test()
    'Test Code
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub MoveBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 0 Then Exit Sub
    TestWindow.Move x / MoveBox.Width * Screen.Width - TestWindow.Width / 2, _
                    y / MoveBox.Height * Screen.Height - TestWindow.Height / 2
End Sub

Private Sub StartBtn_Click()
    Call Test
End Sub
