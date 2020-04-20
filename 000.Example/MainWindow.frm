VERSION 5.00
Begin VB.Form MainWindow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Experiments"
   ClientHeight    =   5040
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   7140
   Icon            =   "MainWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   595
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
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
      Left            =   624
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
Private Sub Form_Load()
    Me.Caption = "Experiment " & Split(Split(App.Path, ".")(0), "Experiments\")(1)
    TitleText.Caption = Me.Caption
    DateText.Caption = Now
    Set Me.Icon = testIcon.Picture
End Sub
Private Sub Outputs(Text As String)
    outputBox.Text = outputBox.Text & Now & "   " & Text & vbCrLf
    outputBox.SelLength = 1
    outputBox.SelStart = Len(outputBox.Text)
End Sub
Private Sub Test()
    'Test Code
End Sub
Private Sub StartBtn_Click()
    Call Test
End Sub
