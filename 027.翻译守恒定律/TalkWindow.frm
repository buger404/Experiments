VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form TalkWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Translation Conservation Law"
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
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.ComboBox Trans 
      Appearance      =   0  'Flat
      Height          =   336
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   72
      Width           =   8820
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   2460
      Left            =   240
      TabIndex        =   3
      Top             =   456
      Width           =   8796
      ExtentX         =   15515
      ExtentY         =   4339
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
      Location        =   "http:///"
   End
   Begin VB.TextBox Words 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2244
      Left            =   336
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "TalkWindow.frx":000C
      Top             =   6528
      Width           =   8772
   End
   Begin VB.TextBox LogBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2268
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3024
      Width           =   9015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "¿½±´"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   3576
      TabIndex        =   7
      Top             =   5400
      Width           =   2148
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ê§Õæ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   324
      Left            =   336
      TabIndex        =   5
      Top             =   6048
      Width           =   480
   End
   Begin VB.Label CommunicateBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "Ò»¼üÊ§Õæ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   408
      TabIndex        =   1
      Top             =   9120
      Width           =   8364
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   4116
      Left            =   0
      TabIndex        =   4
      Top             =   5880
      Width           =   9252
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
    Trans.Enabled = False
    CommunicateBtn.BackColor = RGB(160, 160, 160)

    LogBox.Text = ""

    Select Case Trans.ListIndex
        Case 0
            Call BaiduS
        Case 1
            Call SogoS
        Case 2
            Call GoogleS
        Case 3
            Call BaiduZEZ
    End Select
    
    CommunicateBtn.Enabled = True
    Trans.Enabled = True
    CommunicateBtn.BackColor = RGB(0, 192, 192)
End Sub
Sub GoogleS()
    Dim temp() As String, Text As String, text2 As String
    temp = Split("zh-CN>ru>nl>sw>sl>pt>ja>hi>da>fr>de>zh-CN", ">")
    
    Text = Words.Text
    
    Wait = True
    For i = 1 To UBound(temp)
        NewLog Now & " Connecting to google ..."
    
        MsgBox "https://translate.google.cn/#view=home&op=translate&sl=" & temp(i - 1) & "&tl=" & temp(i) & "&text=" & Replace(Text, vbCrLf, "%0A")
        web.Navigate "https://translate.google.cn/#view=home&op=translate&sl=" & temp(i - 1) & "&tl=" & temp(i) & "&text=" & Replace(Text, vbCrLf, "%0A")
        web.Silent = True
    
        On Error Resume Next
        Do While Wait
            Sleep 10: DoEvents
        Loop
        
        NewLog Now & " Waiting for translation(" & temp(i - 1) & "->" & temp(i) & ") ..."
        
        Do While text2 = "" Or text2 = Text
            text2 = ""
            For s = 1 To 30
                Sleep 10: DoEvents
            Next
            For Each b In web.Document.getelementsbytagname("span")
                If b.classname <> "tlid-translation translation" Then
                Else
                    If text2 <> "" Then text2 = text2 & vbCrLf
                    text2 = text2 & b.innertext
                End If
            Next
        Loop
        
        NewLog Now & " Result : " & text2
        
        Text = text2
    Next

    LogBox.Text = Text
End Sub
Sub SogoS()
    Dim temp() As String, Text As String, text2 As String
    temp = Split("zh-CHS>ru>nl>sw>sl>pt>ja>hi>da>fr>de>zh-CHS", ">")
    
    Text = Words.Text
    
    Wait = True
    For i = 1 To UBound(temp)
        NewLog Now & " Connecting to sogou ..."
    
        web.Navigate "https://fanyi.sogou.com/#" & temp(i - 1) & "/" & temp(i) & "/" & Replace(Text, vbCrLf, "%0A")
        web.Silent = True
    
        On Error Resume Next
        Do While Wait
            Sleep 10: DoEvents
        Loop
        
        NewLog Now & " Waiting for translation(" & temp(i - 1) & "->" & temp(i) & ") ..."
        
        Do While text2 = "" Or text2 = Text
            text2 = ""
            For s = 1 To 30
                Sleep 10: DoEvents
            Next
            For Each b In web.Document.getelementsbytagname("div")
                If b.Id <> "sogou-translate-output" Then
                Else
                    If text2 <> "" Then text2 = text2 & vbCrLf
                    text2 = text2 & b.innertext
                End If
            Next
        Loop
        
        NewLog Now & " Result : " & text2
        
        Text = text2
    Next

    LogBox.Text = Text
End Sub
Sub BaiduZEZ()
    Dim temp() As String, Text As String, text2 As String
    temp = Split("zh>en>jp>de>zh", ">")
    Dim c As Long, i As Integer
    
    Text = Words.Text
    
    i = 1
    
    NewLog "Ô­ÎÄ£º" & Text
    
    Do
    
        Wait = True
        web.Navigate "http://fanyi.baidu.com/?aldtype=16047#" & temp(i - 1) & "/" & temp(i) & "/" & Replace(Text, vbCrLf, "%0A")
        web.Silent = True
    
        On Error Resume Next
        Do While Wait
            Sleep 10: DoEvents
        Loop
        
        text2 = ""
        Do While text2 = ""
            For s = 1 To 30
                Sleep 10: DoEvents
            Next
            For Each b In web.Document.getelementsbytagname("p")
                If b.classname <> "ordinary-output target-output clearfix" Then
                Else
                    If text2 <> "" Then text2 = text2 & vbCrLf
                    text2 = text2 & b.innertext
                End If
            Next
        Loop
        
        c = c + 1
        NewLog Now & " µÚ" & c & "´ÎµÄ·­Òë½á¹û: " & text2
        
        Text = text2
        
        i = i + 1
        If i > UBound(temp) Then i = 1
    Loop

    LogBox.Text = Text
End Sub
Sub BaiduS()
    Dim temp() As String, Text As String, text2 As String
    'temp = Split("zh>ru>nl>swe>slo>pt>jp>el>dan>fra>de>zh", ">")
    temp = Split("zh>pt>jp>en>fra>zh", ">")
    
    Text = Words.Text
    
    For i = 1 To UBound(temp)
        NewLog Now & " Connecting to baidu ..."
    
        Wait = True
        web.Navigate "http://fanyi.baidu.com/?aldtype=16047#" & temp(i - 1) & "/" & temp(i) & "/" & Replace(Text, vbCrLf, "%0A")
        web.Silent = True
    
        On Error Resume Next
        Do While Wait
            Sleep 10: DoEvents
        Loop
        
        NewLog Now & " Waiting for translation(" & temp(i - 1) & "->" & temp(i) & ") ..."
        
        text2 = ""
        Do While text2 = ""
            For s = 1 To 30
                Sleep 10: DoEvents
            Next
            For Each b In web.Document.getelementsbytagname("p")
                If b.classname <> "ordinary-output target-output clearfix" Then
                Else
                    If text2 <> "" Then text2 = text2 & vbCrLf
                    text2 = text2 & b.innertext
                End If
            Next
        Loop
        
        NewLog Now & " Result : " & text2
        
        Text = text2
    Next

    LogBox.Text = Text
End Sub
Sub NewLog(str As String)
    LogBox.Text = LogBox.Text & str & vbCrLf
    LogBox.SelLength = 1
    LogBox.SelStart = Len(LogBox.Text)
End Sub

Private Sub Form_Load()
    Trans.AddItem "°Ù¶È·­Òë"
    Trans.AddItem "ËÑ¹··­Òë"
    Trans.AddItem "¹È¸è·­Òë"
    Trans.AddItem "°Ù¶È·­Òë£¨ÖÐ->Ó¢->ÖÐÑ­»·£©"
    Trans.ListIndex = 0
    
    Words.Text = "¶ì¶ì¶ì£¬ÇúÏîÏòÌì¸è¡£" & vbCrLf & "°×Ã«¸¡ÂÌË®£¬ºìÕÆ²¦Çå²¨¡£"
End Sub

Private Sub Label2_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText LogBox.Text
End Sub

Private Sub web_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Wait = False
End Sub

