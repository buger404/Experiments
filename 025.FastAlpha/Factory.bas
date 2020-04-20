Attribute VB_Name = "Factory"

Public Enum FontModes
    Regular = 0
    Bold = 1
    Italic = 2
End Enum
Public Enum PresentMode
    Normal = 0
    Central = 1
End Enum
Public Enum CtrlClass
    Button = 0
    ProgressBar = 1
    SliderBar = 2
    CheckBox = 3
    HScrollBar = 4
    VScrollBar = 5
    Button2 = 6
    EditBox = 7
    VScrollBar2 = 8
End Enum
Public Enum StrAlignment
    near = 0
    center = 1
    far = 2
End Enum
Public ProDraw As Images, ProCore As PageManager, ProFont As Fonts, TargetDC As Long, ProWin As Form, ProPath As String
Public MousePointer As Integer
Dim mWW As Long, mWH As Long, mProFrm As Form
Public MouseX As Long, MouseY As Long, MouseState As Long, MouseType As Integer, Moused As Boolean
Public CtrlX As Long, CtrlY As Long, CtrlW As Long, CtrlH As Long
Public LockX As Long, LockY As Long, LockW As Long, LockH As Long
Public TLockX As Long, TLockY As Long, TLockW As Long, TLockH As Long, TLockState As Boolean, tRet As Boolean
Public GWW As Long, GWH As Long
Public mNowShow As String
Public LockMousePage As String, NowPage As String
Public DrawCloseButton As Boolean
Public UnClicked As Boolean
Public PublicTextBox As Object
Function cubicCurves(t As Single, value0 As Single, value1 As Single, Value2 As Single, value3 As Single) As Single
    cubicCurves = (value0 * ((1 - t) ^ 3)) + (3 * value1 * t * ((1 - t) ^ 2)) + (3 * Value2 * (t ^ 2) * (1 - t)) + (value3 * (t ^ 3))
    '贝塞尔曲线公式： B(t)=P_0(1-t)^3+3P_1t(1-t)^2+3P_2t^2(1-t)+P_3t^3
End Function
Sub ResetEdit()
    TLockX = 0: TLockY = 0: TLockState = False
End Sub
Function GetRetEdit() As String
    GetRetEdit = PublicTextBox.Text: Call ResetEdit
End Function
Public Function ReadXML(ByVal Path As String, ByVal SectionName As String, ByVal PartName As String) As String

    On Error Resume Next
    
    Dim ObjDom As Object
    Set ObjDom = CreateObject("MicroSoft.XMLDom")
    ObjDom.Load Path
    ReadXML = ObjDom.DocumentElement.SelectSingleNode("//" & SectionName & "//" & PartName).Text
    
    If Err.Number <> 0 Then Err.Clear
End Function
Public Function GetCurrentPath(ByVal Path As String) As String

    Dim temp() As String, ret As String
    
    temp = Split(Path, "\")
    
    For I = 0 To UBound(temp) - 1
        ret = ret & temp(I) & "\"
    Next
    
    GetCurrentPath = ret
    
End Function
Function IsRetEdit() As Boolean
    If TLockState = False Then Exit Function
    If CtrlX <> TLockX Or CtrlY <> TLockY Then Exit Function
    
    If tRet Then IsRetEdit = True
End Function
Function IsShowEdit(ByVal DefeatText As String) As Boolean
    If CtrlX = TLockX And CtrlY = TLockY Then
        TLockState = True: IsShowEdit = True
        Exit Function
    End If
    IsShowEdit = IsMouseDownNoKeep
    If IsShowEdit Then
        TLockState = True
        TLockX = CtrlX: TLockY = CtrlY: TLockW = CtrlW: TLockH = CtrlH
        PublicTextBox.Text = DefeatText
        IsShowEdit = True
    End If
End Function
Function IsMouseIn() As Boolean
    If LockX <> -1 Then
        If LockX = CtrlX And LockY = CtrlY And LockW = CtrlW And LockH = CtrlH Then IsMouseIn = True: MousePointer = MousePointerConstants.vbCustom
        Exit Function
    End If
    
    If MouseX >= CtrlX And MouseY >= CtrlY And MouseX <= CtrlX + CtrlW And MouseY <= CtrlY + CtrlH Then
        IsMouseIn = True
        MousePointer = MousePointerConstants.vbCustom
    End If
End Function
Function IsClick() As Boolean
    IsClick = IsMouseUp
End Function
Function IsMouseDown() As Boolean
    If LockMousePage <> "" Then
        If NowPage <> LockMousePage Then Exit Function
    End If

    If LockX <> -1 Then
        If LockX = CtrlX And LockY = CtrlY And LockW = CtrlW And LockH = CtrlH And MouseState = 1 Then Moused = True: IsMouseDown = True
        Exit Function
    End If
    
    If MouseX >= CtrlX And MouseY >= CtrlY And MouseX <= CtrlX + CtrlW And MouseY <= CtrlY + CtrlH Then
        If MouseState = 1 Then
            IsMouseDown = True
            LockX = CtrlX: LockY = CtrlY: LockW = CtrlW: LockH = CtrlH
            Moused = True
        End If
    End If
    If UnClicked Then IsMouseDown = False
End Function
Function IsMouseDownNoKeep() As Boolean
    If LockMousePage <> "" Then
        If NowPage <> LockMousePage Then Exit Function
    End If

    If LockX <> -1 Then
        If LockX = CtrlX And LockY = CtrlY And LockW = CtrlW And LockH = CtrlH And MouseState = 1 Then Moused = True: IsMouseDownNoKeep = True
        Exit Function
    End If
    
    If MouseX >= CtrlX And MouseY >= CtrlY And MouseX <= CtrlX + CtrlW And MouseY <= CtrlY + CtrlH Then
        If MouseState = 1 Then
            Moused = True
            IsMouseDownNoKeep = True
        End If
    End If
    If UnClicked Then IsMouseDownNoKeep = False
End Function
Function IsMouseUp() As Boolean
    
    If LockMousePage <> "" Then
        If NowPage <> LockMousePage Then Exit Function
    End If

    If LockX <> -1 Then
        If LockX = CtrlX And LockY = CtrlY And LockW = CtrlW And LockH = CtrlH And MouseState = 2 Then
            LockX = -1
            IsMouseUp = True
            Moused = True
            Call ResetEdit
        End If
        Exit Function
    End If
    
    If MouseX >= CtrlX And MouseY >= CtrlY And MouseX <= CtrlX + CtrlW And MouseY <= CtrlY + CtrlH Then
        If MouseState = 2 Then
            LockX = -1
            IsMouseUp = True
            Moused = True
            Call ResetEdit
        End If
    End If

    If UnClicked Then IsMouseUp = False
End Function
Sub CreateFolder(ByVal Path As String)
    Dim temp() As String, NowPath As String, FSO As Object
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    temp = Split(Path, "\")
    For I = 0 To UBound(temp) - 1
        If I <> UBound(temp) - 1 Then
            If FSO.FolderExists(NowPath & temp(I)) = False Then Exit Sub
        End If
        NowPath = NowPath & temp(I) & "\"
        If Dir(NowPath, vbDirectory) = "" Then MkDir NowPath
    Next
End Sub
Public Sub BlurTo(DC As Long, Optional Radius As Long = 60)
    If ProWin Is Nothing Then Exit Sub
    Dim Image As Long, Graphics As Long
    ProWin.AutoRedraw = True
    ProDraw.Draw ProWin.hdc, 0, 0
    ProWin.Refresh
    DoEvents
    GdipCreateBitmapFromHBITMAP ProWin.Image.handle, ProWin.Image.hpal, Image
    BlurImage Image, GWW, GWH, Radius
    GdipCreateFromHDC DC, Graphics
    GdipDrawImage Graphics, Image, 0, 0
    DelImage Image
    GdipDeleteGraphics Graphics
    ProWin.AutoRedraw = False
End Sub
Public Sub BlurTo2(DC As Long, orandc As Long, Optional Radius As Long = 60)
    If ProWin Is Nothing Then Exit Sub
    Dim Image As Long, Graphics As Long
    ProWin.AutoRedraw = True
    BitBlt ProWin.hdc, 0, 0, GWW, GWH, orandc, 0, 0, vbSrcCopy
    ProWin.Refresh
    DoEvents
    GdipCreateBitmapFromHBITMAP ProWin.Image.handle, ProWin.Image.hpal, Image
    BlurImage Image, GWW, GWH, Radius
    GdipCreateFromHDC DC, Graphics
    GdipDrawImage Graphics, Image, 0, 0
    DelImage Image
    GdipDeleteGraphics Graphics
    ProWin.AutoRedraw = False
End Sub
Sub BlurImage(Image As Long, Width As Long, Height As Long, Optional Radius As Long = 60)
    Dim Effect As Long
    Dim p As BlurParams
    GdipCreateEffect2 GdipEffectType.Blur, Effect
    p.Radius = Radius
    GdipSetEffectParameters Effect, p, LenB(p)
    GdipBitmapApplyEffect Image, Effect, NewRectL(0, 0, Width, Height), 0, 0, 0
    GdipDeleteEffect Effect
End Sub
Sub SetClickArea2(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long)
    CtrlX = X: CtrlY = Y
    CtrlW = Width: CtrlH = Height
End Sub
Public Property Get WW() As Long
    WW = mWW
End Property
Public Property Get WH() As Long
    WH = mWH
End Property
Public Property Get GS() As Object
    Set GS = ProSave
End Property
Public Property Get NowShow() As String
    NowShow = mNowShow
End Property
Public Property Let NowShow(ByVal NewShow As String)
    mNowShow = NewShow
End Property
Public Sub StartProgram(ProFrm As Object, nProPath As String, Power As Boolean)
    If Power = True Then
        InitGDIPlus
        Set mProFrm = ProFrm
        BASS_Init -1, 44100, BASS_DEVICE_3D, ProFrm.Hwnd, 0
        mWW = ProFrm.ScaleWidth: mWH = ProFrm.ScaleHeight
        GWW = mWW: GWH = mWH
        ProPath = nProPath
        InitDustbin
        Set ProDraw = New Images
        ProDraw.Create ProFrm.hdc, mWW, mWH
        TargetDC = ProFrm.hdc
        Set ProCore = New PageManager
        Set ProWin = ProFrm
    Else
        DoClearing
        TerminateGDIPlus
        BASS_Free
    End If
End Sub
Public Sub UpdateClickTest(ByVal X As Long, ByVal Y As Long, ByVal State As Long)
    If State < MouseState Then Exit Sub
    MouseX = X: MouseY = Y
    MouseState = State
End Sub
Sub ResetClick()
    MouseState = 0
End Sub
Public Function FriendError(ByVal Num As Long) As String
    FriendError = Error(Num)
    Select Case Num
        Case 0
        FriendError = "等一下？哪来的错误？"
        Case 5
        FriendError = "多半是该死的404又忘记把废弃的代码删干净了。"
        Case 6
        FriendError = "shit 404为变量倒果汁的时候不小心溢出来了。"
        Case 7
        FriendError = "真的很抱歉，404没有做好清洁工的职责。"
        Case 9
        FriendError = "404让我在数组的边缘试探...哦豁，掉入悬崖了。"
        Case 11
        FriendError = "对不起，脑残404数学不过关，除以0什么的..."
        Case 13
        FriendError = "404刚才在给别人介绍对象的时候被双方左右各扇了一巴掌。"
        Case 28
        FriendError = "罢工！404一次性让我们做太多的工作了！"
        Case 35
        FriendError = "多半是该死的404又忘记把废弃的代码删干净了。"
        Case 52
        FriendError = "404给错了房间号码。"
        Case 53
        FriendError = "等等...404的GPS出了点问题..."
        Case 55
        FriendError = "丢三落四的404打开一个文件后忘了关上。"
        Case 58
        FriendError = "嗯。。这个文件已经存在了，用TNT炸掉么？"
        Case 70
        FriendError = "抱歉，身在底层的我实在没有权利完成这项工作。"
        Case 75
        FriendError = "404！这个地址是什么毛线啦！"
        Case 76
        FriendError = "404！这里已经拆迁了啦！"
    End Select
End Function
