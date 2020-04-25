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
    Private Type Ball
        Vx As Double
        Vy As Double
        Ax As Double
        Ay As Double
        X As Single
        Y As Single
        OX As Single
        OY As Single
        Life As Long
        Time As Single
        Size As Long
        See As Boolean
        Box As New GCrashBox
        R As Long
        G As Long
        B As Long
    End Type
    Dim bs() As Ball
    Dim CD As Long
'======================================
Public Sub AllIntoBox()
    For I = 1 To UBound(bs)
        bs(I).Box.IsCrash
    Next
End Sub
Public Sub Update()
    '应用界面每一帧刷新会调用的方法（页面绘制代码请放在此处）
    
    Page.Clear '清空画布
    
    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    Dim Span As Single, SX As Long, SY As Long
    SX = 20: SY = 20
    For I = 1 To UBound(bs)
        If I > UBound(bs) Then Exit For
        Span = Timer - bs(I).Time
        Span = Span * IIf(AppWindow.Slowed, 0.1, 1)
        bs(I).Vy = bs(I).Vy + bs(I).Ay * Span
        bs(I).Vx = bs(I).Vx + bs(I).Ax * Span
        bs(I).OX = bs(I).X: bs(I).OY = bs(I).Y
        bs(I).X = bs(I).X + bs(I).Vx * Span
        bs(I).Y = bs(I).Y + bs(I).Vy * Span
        bs(I).Time = Timer
        If bs(I).Life > 255 Then bs(I).Life = 255
        Page.Paint 1, bs(I).X, bs(I).Y, bs(I).Size, bs(I).Size, IIf(bs(I).See, argb(bs(I).Life, bs(I).R, bs(I).G, bs(I).B), argb(bs(I).Life, 255, 83, 57)), pos:=posOnCenter
        bs(I).Box.Reset
        If bs(I).See Then
            Page.Writes "生命周期  " & bs(I).Life & vbCrLf & "位置  " & bs(I).X & "," & bs(I).Y & vbCrLf & "水平方向加速度  " & bs(I).Ax & " pixel/s^2" & vbCrLf & "竖直方向加速度  " & bs(I).Ay & " pixel/s^2" & vbCrLf & "水平方向速度  " & bs(I).Vx & " pixel/s" & vbCrLf & "竖直方向速度  " & bs(I).Vy & " pixel/s", SX, SY, 18, argb(255, bs(I).R, bs(I).G, bs(I).B)
            SX = SX + 450
            If SX > GW Then SX = 20: SY = SY + 170
        End If
        If bs(I).Y > GH Or bs(I).Life <= 8 Or bs(I).Y < 0 Or bs(I).X < 0 Or bs(I).X > GW Then
            Set bs(I).Box = Nothing
            bs(I) = bs(UBound(bs))
            ReDim Preserve bs(UBound(bs) - 1)
            I = I - 1
        End If
    Next
    
    For I = 1 To UBound(bs)
        For J = 1 To UBound(bs)
            If I <> J Then
                DrawF.X = bs(J).X: DrawF.Y = bs(J).Y
                DrawF.Width = 30: DrawF.Height = 30
                bs(I).Box.IsCrash
            End If
        Next
    Next
    
    Page.Paint 0, 100, GH / 2, 10, GH * 0.6, argb(255, 0, 0, 0), pos:=posOnCenter
    Call AllIntoBox
    Page.Paint 0, 300, GH / 2, 10, GH * 0.5, argb(255, 0, 0, 0), pos:=posOnCenter
    Call AllIntoBox
    Page.Paint 0, 500, GH / 2, 10, GH * 0.3, argb(255, 0, 0, 0), pos:=posOnCenter
    Call AllIntoBox
    Page.Paint 0, 180, GH * 0.6, 10, GH * 0.4, argb(255, 0, 0, 0)
    Call AllIntoBox
    Page.Paint 0, 400, GH * 0.4, 10, GH * 0.4, argb(255, 0, 0, 0)
    Call AllIntoBox
    Page.Paint 0, 700, GH * 0.5, 10, GH * 0.2, argb(255, 0, 0, 0)
    Call AllIntoBox
    Page.Paint 0, 650, GH * 0.8, 10, GH * 0.9, argb(255, 0, 0, 0)
    Call AllIntoBox
    'Page.Paint 0, GW - 10, 0, 10, GH, argb(255, 100, 100, 100)
    'Call AllIntoBox
    
    For I = 1 To UBound(bs)
        If bs(I).Box.CrashResult Then
            bs(I).X = bs(I).OX: bs(I).Y = bs(I).OY
            bs(I).Vx = -bs(I).Vx * 0.3
            bs(I).Vy = -bs(I).Vy * 0.3
            bs(I).Ax = -bs(I).Ax
            bs(I).Life = bs(I).Life * 0.9
        Else
            bs(I).Life = bs(I).Life * 1.05
        End If
    Next
    
    If IsKeyPress(VK_Z) Then
        If GetTickCount - CD >= 50 Then
            CD = GetTickCount
            ReDim Preserve bs(UBound(bs) + 1)
            With bs(UBound(bs))
                .X = Rnd * GW
                .Y = Rnd * GH
                .OX = .X
                .OY = .Y
                .Size = Int(Rnd * 25) + 15
                .Life = 255
                .Ax = -0.1 * IIf(Int(Rnd * 2) = 1, -1, 1)
                .Ay = 9.8 * 30
                .Time = Timer
                .Vx = Rnd * 10 * 50 * IIf(.Ax < 0, 1, -1)
            End With
        End If
    End If
    If IsKeyPress(VK_Q) Then
        For I = 1 To UBound(bs)
            bs(I).Ax = 2000 * Rnd
        Next
    End If
    If IsKeyPress(VK_W) Then
        For I = 1 To UBound(bs)
            bs(I).Ax = 6000 * Rnd
        Next
    End If
    If IsKeyPress(VK_U) Then
        For I = 1 To UBound(bs)
            bs(I).Ay = -9.8 * 30
        Next
    End If
    If IsKeyPress(VK_F) Then
        For I = 1 To UBound(bs)
            bs(I).Ay = 0
        Next
    End If
    If IsKeyPress(VK_B) Then
        For I = 1 To UBound(bs)
            bs(I).Size = bs(I).Size * 1.05
        Next
    End If
    If IsKeyPress(VK_S) Then
        For I = 1 To UBound(bs)
            bs(I).Size = bs(I).Size * 0.95
        Next
    End If
    If IsKeyPress(VK_G) Then
        For I = 1 To UBound(bs)
            bs(I).Vy = 9.8 * 100
        Next
    End If
    
    If CheckMouse(0, 0, GW, GH) >= mMouseDown Then
        If GetTickCount - CD >= 400 Then
            CD = GetTickCount
            ReDim Preserve bs(UBound(bs) + 1)
            With bs(UBound(bs))
                .X = Mouse.X
                .Y = Mouse.Y
                .OX = .X
                .OY = .Y
                .See = (Mouse.button = 2)
                .R = Rnd * 255
                .G = Rnd * 255
                .B = Rnd * 255
                .Size = Int(Rnd * 25) + 15
                .Life = 255
                .Ax = -0.1 * IIf(Int(Rnd * 2) = 1, -1, 1)
                .Ay = 9.8 * 30
                .Time = Timer
                .Vx = Rnd * 10 * 50 * IIf(.Ax < 0, 1, -1)
            End With
        Else
            Dim apro As Single
            apro = (GetTickCount - CD) / 400
            If apro > 1 Then apro = 1
            Page.PaintArc Mouse.X, Mouse.Y, 60, 60, apro * 400, -90, argb(255 * (1 - apro), 0, 0, 0), style:=2, pos:=posOnCenter
        End If
    Else
        Page.Writes "按住鼠标放置小球（右键创建观测小球）", GW / 2, 70, 24, argb(255 * Abs((GetTickCount Mod 1200) / 600 - 1), 255, 0, 0), align:=StringAlignmentCenter
        Page.Writes "Z：生成大量小球   Q：刮风   W：台风   B：增大   S：缩小   U：反重力   F：悬浮   G：超重力   Space：减慢", GW / 2, GH - 160, 18, argb(255 * Abs((GetTickCount Mod 1200) / 600 - 1), 255, 0, 0), align:=StringAlignmentCenter
    End If
    
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
    ReDim bs(0)

    Set Page = New GPage  '创建页面绘制器
    Page.Create Me  '将绘制器与控制器绑定
    Page.Res.NewImages App.path & "\assets"  '导入资源
    ECore.Add Page, "AppPage"  '创建页面
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
