VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   该类模块是由Emerald创建的 页面控制器 模板
'======================================
'   页面绘制器
    Dim Page As GPage
'======================================

Public Sub Update()
    '应用界面每一帧刷新会调用的方法（页面绘制i代码请放在此处）
    
    Page.Clear argb(0, 0, 0, 0) '清空画布

    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    Dim Pass As Long, alpha As Single, CircleAlpha As Single
    Pass = Music.position * 1000 - time

    On Error Resume Next
    For I = 1 To UBound(CurrentObjects)
        With CurrentObjects(I)
            If .time - o.SliderMultiplier * 200 - 1000 > Pass Then Exit For
            If Abs(.time - Pass) <= (o.SliderMultiplier * 200) Then
                If .time > Pass Then alpha = 1 - Cubic((.time - Pass) / (o.SliderMultiplier * 200), 0, 1, 1, 1): CircleAlpha = 0 '1 - alpha
                If .time <= Pass Then alpha = 0: CircleAlpha = 0 '1 - (Pass - .time) / (o.SliderMultiplier * 300)
                Page.DrawImage .Kind & IIf(alpha > 0.3, "lost", "") & ".png", .X / 700 * GW + 50, .y / 560 * GH + 50, alpha:=alpha ', pos:=posOnCenter
                'DrawBox.Circle (.X / 700 * DrawBox.ScaleWidth + 50, .Y / 560 * DrawBox.ScaleHeight + 50), 20, RGB(255 - alpha * 255, 255 - alpha * 255, 255 - alpha * 255)
                'If alpha <> 0 Then DrawBox.Circle (.X / 700 * DrawBox.ScaleWidth + 50, .Y / 560 * DrawBox.ScaleHeight + 50), 20 + 40 * CircleAlpha, RGB(alpha * 255, 255 - alpha * 255, 255 - alpha * 255)
                If (Not .sound Is Nothing) And alpha > 0 Then
                    If .sound.Volume <> 0.6 Then
                        .sound.Volume = 0.6
                        .sound.Play
                    End If
                End If
            End If
            If Not .sound Is Nothing Then
                If .sound.IsPlaying = False And .sound.Volume = 0.6 Then
                    .sound.Dispose
                    Set .sound = Nothing
                End If
            End If
        End With
    Next
    'DrawBox.Refresh
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    Set Page = New GPage  '创建页面绘制器
    Page.Create Me  '将绘制器与控制器绑定
    Page.Res.NewImages App.path & "\assets", 0.7 '导入资源
    ECore.Add Page, "AppPage"  '创建页面
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
