VERSION 5.00
Begin VB.Form MainWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
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
Public Sub Download(ByVal nUrl As String, ByVal nFile As String)
     Dim XmlHttp, b() As Byte
     Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
     XmlHttp.Open "GET", nUrl, False
     XmlHttp.Send
     If XmlHttp.ReadyState = 4 Then
         b() = XmlHttp.responseBody
         Open nFile For Binary As #1
         Put #1, , b()
         Close #1
     End If
     Set XmlHttp = Nothing
End Sub
Public Function NetContent(ByVal nUrl As String) As String
     Dim XmlHttp, b() As Byte
     Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
     XmlHttp.Open "GET", nUrl, False
     XmlHttp.Send
     NetContent = StrConvEx(XmlHttp.responseBody)
     Set XmlHttp = Nothing
End Function
Function StrConvEx(b, Optional Charset As String = "GB2312")
    Dim o As Object
    Set o = CreateObject("Adodb.Stream")
    With o
        .Type = 1: .Mode = 3
        .Open: .Write b
        .Position = 0: .Type = 2
        .Charset = Charset
    End With
    StrConvEx = o.ReadText: o.Close
    Set o = Nothing
End Function
Public Function GetLoginQQ() As String()
    Dim Hwnd As Long, QQ As String, Size As Integer, Class As String
    Dim Ret() As String
    ReDim Ret(0)
    Hwnd = FindWindowA("CTXOPConntion_Class", vbNullString)
    If Hwnd = 0 Then Exit Function
    
    Do While Hwnd <> 0
        QQ = String(255, vbNullChar)
        GetWindowTextA Hwnd, QQ, Len(QQ)
        QQ = Left(QQ, InStr(QQ, vbNullChar) - 1)
        If InStr(QQ, "OP_") = 1 Then
            QQ = Mid(QQ, 4)
            ReDim Preserve Ret(UBound(Ret) + 1)
            Ret(UBound(Ret)) = QQ
        End If
        Hwnd = GetWindow(Hwnd, GW_HWNDNEXT)
    Loop
    GetLoginQQ = Ret
End Function
Private Sub Outputs(Text As String)
    outputBox.Text = outputBox.Text & Now & "   " & Text & vbCrLf
    outputBox.SelLength = 1
    outputBox.SelStart = Len(outputBox.Text)
End Sub
Private Sub Test()
    Dim QQ() As String, Ret As String, Temp As String, Args() As String
    QQ = GetLoginQQ
    
    Me.Cls
    
    For i = 1 To UBound(QQ)
        Temp = NetContent("https://users.qzone.qq.com/fcg-bin/cgi_get_portrait.fcg?uins=" & QQ(i))
        Args = Split(Temp, """")
        Outputs "QQ " & QQ(i) & "(" & Args(5) & ")"
        Download Args(3), App.Path & "\" & QQ(i) & ".bmp"
        Me.PaintPicture LoadPicture(App.Path & "\" & QQ(i) & ".bmp"), (i - 1) * 140 + 50, 50
        Me.CurrentX = (i - 1) * 140 + 50: Me.CurrentY = 150: Me.Print Args(5)
    Next
    
    Me.Refresh
    
End Sub
Private Sub StartBtn_Click()
    Call Test
End Sub
