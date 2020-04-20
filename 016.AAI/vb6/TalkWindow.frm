VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form TalkWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artifical Artifical Intelligent"
   ClientHeight    =   9900
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   9228
   Icon            =   "TalkWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   9228
   StartUpPosition =   2  '屏幕中心
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   30
      Left            =   720
      TabIndex        =   3
      Top             =   2640
      Width           =   30
      ExtentX         =   53
      ExtentY         =   53
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox LogBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   9015
   End
   Begin VB.TextBox Words 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Text            =   "说点什么吧"
      Top             =   8040
      Width           =   8415
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "发送"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   360
      TabIndex        =   7
      Top             =   7440
      Width           =   480
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "正在和帅气的人工智障聊天"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3120
      TabIndex        =   6
      Top             =   120
      Width           =   2880
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9255
   End
   Begin VB.Label CommunicateBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "沟通"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   9120
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   0
      TabIndex        =   4
      Top             =   7200
      Width           =   9255
   End
End
Attribute VB_Name = "TalkWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Wait As Boolean
Private Sub CommunicateBtn_Click()
    CommunicateBtn.Enabled = False
    CommunicateBtn.Caption = "思考中..."

    web.Navigate "https://zhidao.baidu.com/search?lm=0&rn=10&pn=0&fr=search&ie=gbk&word=" & Words.Text
    Wait = True
    web.Silent = True
    NewLog "你(" & Now & ")：" & vbCrLf & Words.Text
    
    On Error Resume Next
    Do While Wait
        Sleep 10: DoEvents
    Loop
    
    Dim b As Object, link() As String
    ReDim link(0)
    
    For Each b In web.Document.getelementsbytagname("a")
        If b.classname <> "ti" Then
        Else
            ReDim Preserve link(UBound(link) + 1)
            link(UBound(link)) = b.href
        End If
    Next
    
    Dim Num As Long
    Randomize
    Num = Int(Rnd * UBound(link))
    If UBound(link) = Num Then Num = Num - 1
    
    web.Navigate link(Num + 1)
    Wait = True
    
    Do While Wait
        Sleep 10: DoEvents
    Loop
    
    For Each b In web.Document.getelementsbytagname("div")
        If InStr(b.classname, "showbtn") = 0 Then
        Else
            b.Click
        End If
    Next
    
    For i = 1 To 100
        Sleep 10: DoEvents
    Next
    
    Dim temp() As String, str As String
    Dim ans() As String
    ReDim ans(0)
    
    For Each b In web.Document.getelementsbytagname("div")
        If InStr(b.Id, "content-") = 0 Then
        Else
            str = Replace(Replace(Replace(b.innertext, "展开全部", ""), "采纳", "夸奖"), "向左转|向右转", "")
            temp = Split(str, vbCrLf)
            str = ""
            For i = 0 To UBound(temp) - 1
                If Trim(temp(i)) <> "" Then str = str & temp(i) & vbCrLf
            Next
            str = str & temp(UBound(temp))
            ReDim Preserve ans(UBound(ans) + 1)
            ans(UBound(ans)) = str
        End If
    Next

    ReDim Preserve ans(UBound(ans) + 1)
    ans(UBound(ans)) = ans(UBound(ans) - 1)

    Dim ra() As String
    ra = Split(ans(Rnd * (UBound(ans) - 1) + 1), "。")
    Dim tts As String, fail As Long
rechoose:
    
    tts = ra(Rnd * UBound(ra))
    
    If Replace(tts, " ", "") = "" Then
        fail = fail + 1
        If fail >= 100 Then Exit Sub
        GoTo rechoose
    End If
    
    NewLog "AAI(" & Now & ")：" & vbCrLf & tts

    CommunicateBtn.Enabled = True
    CommunicateBtn.Caption = "沟通"
End Sub
Sub NewLog(str As String)
    LogBox.Text = LogBox.Text & str & vbCrLf
    LogBox.SelLength = 1
    LogBox.SelStart = Len(LogBox.Text)
End Sub

Private Sub Form_Load()
    If MsgBox("此程序包含令人不适内容，继续吗？", 48 + vbYesNo, "人工智障") = vbNo Then End
End Sub

Private Sub Label3_Click()
    MsgBox "由脑回路新奇的Error 404制作。"
End Sub

Private Sub web_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Wait = False
End Sub

