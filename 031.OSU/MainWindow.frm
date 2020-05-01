VERSION 5.00
Begin VB.Form MainWindow 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Osu!Error"
   ClientHeight    =   3876
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   8616
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   323
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   718
   StartUpPosition =   2  '屏幕中心
   Begin VB.HScrollBar Process 
      Height          =   324
      Left            =   312
      TabIndex        =   3
      Top             =   3432
      Width           =   7968
   End
   Begin VB.ListBox LevelList 
      Appearance      =   0  'Flat
      Height          =   744
      Left            =   312
      TabIndex        =   2
      Top             =   2496
      Width           =   6564
   End
   Begin VB.ListBox SongList 
      Appearance      =   0  'Flat
      Height          =   2184
      Left            =   312
      TabIndex        =   1
      Top             =   156
      Width           =   7968
   End
   Begin VB.CommandButton TestBtn 
      Caption         =   "BOOM!"
      Height          =   792
      Left            =   7176
      TabIndex        =   0
      Top             =   2496
      Width           =   1104
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time As Long
Dim OsuPath As String

Private Sub Form_Load()
    ReDim CurrentObjects(0)
    AppWindow.Show

    Set ESave = New GSaving
    ESave.Create "osu!error", "37DED0953BC8D9FA31D04BA428E23B7A"
    
    Dim Song As String
    If ESave.GetData("path") = "" Then
        OsuPath = "C:\Users\" & VBA.Environ("username") & "\AppData\Local\osu!\Songs\"
    Else
        OsuPath = ESave.GetData("path")
    End If
    
    Song = Dir(OsuPath, vbDirectory)
    If Song = "" Then
        If MsgBox("没有发现osu曲谱，您需要指定一个路径吗？", 48 + vbYesNo) <> vbYes Then Exit Sub
    End If
    Do While Dir(OsuPath, vbDirectory) = ""
        OsuPath = InputBox("请指定一个位置。")
        ESave.PutData "path", OsuPath
        DoEvents
    Loop
    Song = Dir(OsuPath, vbDirectory)
    
    Do While Song <> ""
        If Song <> "." And Song <> ".." And (Song Like "*.*" = False) Then
            SongList.AddItem Song
        End If
        Song = Dir(, vbDirectory)
    Loop
    
    Me.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload AppWindow
    End
End Sub

Private Sub Process_Change()
    Music.position = Process.value / Process.max * Music.Length
End Sub

Private Sub SongList_Click()
    Dim Level As String
    Level = Dir(OsuPath & SongList.List(SongList.ListIndex) & "\*.osu")
    LevelList.Clear
    Do While Level <> ""
        LevelList.AddItem Level
        Level = Dir()
    Loop
End Sub

Private Sub TestBtn_Click()
    

    If Dir(OsuPath & SongList.List(SongList.ListIndex) & "\" & LevelList.List(LevelList.ListIndex)) = "" Then
        MsgBox "请先选择上方的歌曲，再选择旁边的难度。", 64
        Exit Sub
    End If
    Dim Code As String, t As String
    Open OsuPath & SongList.List(SongList.ListIndex) & "\" & LevelList.List(LevelList.ListIndex) For Input As #1
    Do While Not EOF(1)
        Line Input #1, t
        Code = Code & t & vbCrLf
    Loop
    Close #1
    o.Load Code
    o.GetObjects
    Music.Dispose
    Music.Create OsuPath & SongList.List(SongList.ListIndex) & "\" & o.Audio
    Music.Play
End Sub

Private Sub TimeBtn_Click()
    time = GetTickCount
End Sub
