VERSION 5.00
Begin VB.Form MainWindow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Experiments"
   ClientHeight    =   5040
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10608
   Icon            =   "MainWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   884
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.PictureBox ScrImg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2220
      Left            =   7440
      ScaleHeight     =   2196
      ScaleWidth      =   2268
      TabIndex        =   5
      Top             =   2592
      Width           =   2292
   End
   Begin VB.PictureBox ConstImg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2220
      Left            =   7440
      ScaleHeight     =   2196
      ScaleWidth      =   2268
      TabIndex        =   4
      Top             =   168
      Width           =   2292
   End
   Begin VB.TextBox outputBox 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0C0&
      Height          =   1764
      Left            =   2664
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1656
      Width           =   3636
   End
   Begin VB.Label DateText 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Date]"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   2664
      TabIndex        =   2
      Top             =   1296
      Width           =   648
   End
   Begin VB.Label TitleText 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Experiment"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CEDB1A&
      Height          =   276
      Left            =   2664
      TabIndex        =   1
      Top             =   936
      Width           =   1320
   End
   Begin VB.Image testIcon 
      Height          =   1536
      Left            =   648
      Picture         =   "MainWindow.frx":000C
      Top             =   984
      Width           =   1536
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
      Left            =   5088
      TabIndex        =   0
      Top             =   4032
      Width           =   1236
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As New RecognizeImg
Private Sub Form_Load()
    Me.Caption = "Experiment " & Split(Split(App.Path, ".")(0), "Experiments\")(1)
    TitleText.Caption = Me.Caption
    DateText.Caption = Now
    Set Me.Icon = testIcon.Picture
End Sub
Public Sub Outputs(Text As String)
    outputBox.Text = outputBox.Text & Now & "   " & Text & vbCrLf
    outputBox.SelLength = 1
    outputBox.SelStart = Len(outputBox.Text)
End Sub
Private Sub Test()
    'Test Code
    ConstImg.Cls
    ConstImg.PaintPicture testIcon.Picture, 0, 0
    ConstImg.Refresh
    Outputs "Picture painted ."
    Outputs "Picture rendering ..."
    r.RenderRecognizeImg ConstImg.hDC, testIcon.Width, testIcon.Height
    Outputs "Picture rendered ."
    ScrImg.Width = Screen.Width / Screen.TwipsPerPixelX: ScrImg.Height = Screen.Height / Screen.TwipsPerPixelY
    BitBlt ScrImg.hDC, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, GetDC(0), 0, 0, vbSrcCopy
    ScrImg.Refresh
    Outputs "Screen catched ."
    Outputs "Picture rendering ..."
    r.RenderRecognizeImg ScrImg.hDC, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY
    Outputs "Picture rendered ."
    ScrImg.Refresh
    ConstImg.Refresh
End Sub
Private Sub StartBtn_Click()
    Call Test
End Sub
