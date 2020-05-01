VERSION 5.00
Begin VB.Form AppWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Osu!Error"
   ClientHeight    =   6672
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   9660
   LinkTopic       =   "AppWindow"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   9000
      Top             =   240
   End
End
Attribute VB_Name = "AppWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'   该类模块是由Emerald创建的 界面容器窗口（应用窗口） 模板
'==================================================
'   页面管理器
    Dim EC As GMan
'==================================================
'   在此处放置你的页面控制器类模块声明语句
    Dim AppPage As AppPage
    Dim CloseMark As Boolean
'==================================================

Private Sub DrawTimer_Timer()
    DrawTimer.Enabled = False
    Do While Not CloseMark
        EC.Display
        DoEvents
    Loop
End Sub

Private Sub Form_Load()
    StartEmerald Me.Hwnd, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY   '初始化Emerald（在此处可以修改窗口大小）
    MakeFont "微软雅黑"  '创建字体
   
    Set EC = New GMan   '创建页面管理器
    EC.Layered False
    
    '创建存档（可选），存档key的问题请查看Emerald的wiki
    'Set ESave = New GSaving
    'ESave.Create "emerald.test", "Emerald.test"
    
    '创建音乐列表（可选）
    'Set MusicList = New GMusicList
    'MusicList.Create App.path & "\music"

    '开始显示界面
    Me.Show
    
    '在此处实例化你的页面控制器
    '=============================================
    '示例：TestPage.cls
    '     Set TestPage = New TestPage
    '公共部分：Dim TestPage As TestPage
        Set AppPage = New AppPage
    '=============================================

    '设置活动页面（在此处设置则为你的启动页面）
    EC.ActivePage = "AppPage"
    
    SetWindowPos Me.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    DrawTimer.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '终止绘制
    CloseMark = True
    '释放Emerald资源
    EndEmerald
End Sub

'============================================================
' 事件映射
Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    UpdateMouse X, y, 1, button
End Sub
Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    If Mouse.State = 0 Then
        UpdateMouse X, y, 0, button
    Else
        Mouse.X = X: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    UpdateMouse X, y, 2, button
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    '发送字符输入
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub
'============================================================
