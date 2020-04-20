VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Screens() As Object, StopWork As Boolean
Public TransTime As Long, TransPage As String, TransEffect As Integer
Dim FadeDraw As New Images, ErrDraw As Images
Sub FadePage(Page As String)
    ProDraw.Draw FadeDraw.CompatibleDC, 0, 0
    TransTime = GetTickCount
    TransPage = Page
    TransEffect = 1
    Call ResetEdit: TLockX = 0
End Sub
Sub ChangePage(Page As String)
    TransTime = GetTickCount
    TransPage = Page
    TransEffect = 0
    Call ResetEdit: TLockX = 0
End Sub
Sub AddScreen(NewScreen As Object)
    ReDim Preserve Screens(UBound(Screens) + 1)
    Set Screens(UBound(Screens)) = NewScreen
End Sub
Sub Display()
    'On Error GoTo CatchError
    
    Moused = False
    
    If StopWork = True Then
        ErrDraw.Draw ProDraw.CompatibleDC, 0, 0, 255
        If DrawCloseButton Then
            CtrlX = -1: CtrlY = -1: CtrlW = 0: CtrlH = 0
            SetClickArea2 20, 0, GWW - 60 - 30, 30 + 40
            ProFont.DrawText ProDraw.Graphics, 20, 20, GWW, 30, ProWin.Caption, argb(255, 189, 189, 189), near, 18, Regular
            If IsMouseDownNoKeep Then
                ReleaseCapture
                SendMessageA ProWin.Hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
                MouseState = 2
            End If
            SetClickArea2 GWW - 30 - 20, 20, 30, 30
            ProFont.DrawText ProDraw.Graphics, GWW - 30 - 10, 20, 30, 30, "x", IIf(IsMouseIn, argb(255, 241, 77, 31), argb(255, 189, 189, 189)), center, 18, Regular
            If IsClick Then Unload MainWindow: Exit Sub
            
            SetClickArea2 GWW - 60 - 20, 20, 30, 30
            ProFont.DrawText ProDraw.Graphics, GWW - 60 - 10, 20, 30, 30, "-", IIf(IsMouseIn, argb(255, 0, 176, 240), argb(255, 189, 189, 189)), center, 18, Regular
            If IsClick Then MainWindow.WindowState = 1
            
            If MouseState = 2 Then ResetClick: LockX = -1
            If MouseState = 0 And LockX <> -1 Then LockX = -1
        End If
        
        ProDraw.Draw TargetDC, 0, 0, 255
        Exit Sub
    End If
    
    CtrlX = -1: CtrlY = -1: CtrlW = 0: CtrlH = 0
    'NowEdit = -1
    TLockState = False
    
    MousePointer = MousePointerConstants.vbDefault
    
    Select Case TransEffect
        Case 0
            If GetTickCount - TransTime < 1000 Then
                'UnClicked = True
                If GetTickCount - TransTime >= 500 Then
                   If mNowShow <> TransPage Then mNowShow = TransPage
                End If
            Else
                'UnClicked = False
            End If
        Case 1
            If GetTickCount - TransTime < 400 Then
                If mNowShow <> TransPage Then mNowShow = TransPage
                'UnClicked = True
            Else
                'UnClicked = False
            End If
    End Select
    
    Dim i As Integer
    For i = 1 To UBound(Screens)
        If Screens(i).Visible Then Screens(i).Display
    Next

    Select Case TransEffect
        Case 0
            If GetTickCount - TransTime < 1000 Then
                If GetTickCount - TransTime < 500 Then
                    'BackImg.ImageByIndex(1).Draw ProDraw.CompatibleDC, 0, 0, cubicCurves((GetTickCount - TransTime) / 500, 0, 1, 1, 1) * 255
                ElseIf GetTickCount - TransTime >= 500 Then
                    'BackImg.ImageByIndex(1).Draw ProDraw.CompatibleDC, 0, 0, 255 - cubicCurves((GetTickCount - TransTime - 500) / 500, 0, 1, 1, 1) * 255
                End If
            End If
        Case 1
            If GetTickCount - TransTime < 400 Then FadeDraw.Draw ProDraw.CompatibleDC, 0, 0, 255 - cubicCurves((GetTickCount - TransTime) / 400, 0, 1, 1, 1) * 255
    End Select
    
    'ProFont.DrawText ProDraw.Graphics, 0, 0, 0, 0, LockMousePage, argb(255, 255, 0, 0), near, 18, Bold
    If DrawCloseButton Then
        SetClickArea2 20, 0, GWW - 60 - 30, 30 + 40
        ProFont.DrawText ProDraw.Graphics, 20, 20, GWW, 30, ProWin.Caption, argb(255, 189, 189, 189), near, 18, Regular
        If IsMouseDownNoKeep Then
            ReleaseCapture
            SendMessageA ProWin.Hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
            MouseState = 2
        End If
        SetClickArea2 GWW - 30 - 20, 20, 30, 30
        ProFont.DrawText ProDraw.Graphics, GWW - 30 - 10, 20, 30, 30, "x", IIf(IsMouseIn, argb(255, 241, 77, 31), argb(255, 189, 189, 189)), center, 18, Regular
        If IsClick Then Unload MainWindow: ResetClick: LockX = -1: Exit Sub
        
        SetClickArea2 GWW - 60 - 20, 20, 30, 30
        ProFont.DrawText ProDraw.Graphics, GWW - 60 - 10, 20, 30, 30, "-", IIf(IsMouseIn, argb(255, 0, 176, 240), argb(255, 189, 189, 189)), center, 18, Regular
        If IsClick Then
            MainWindow.WindowState = 1
            Call MainWindow.OnMinsize
        End If
    End If
    
    
    ProDraw.Draw TargetDC, 0, 0, 255
    
    'If Screen.MousePointer = MousePointerConstants.vbDefault Or Screen.MousePointer = MousePointerConstants.vbCustom Then
        'If Screen.MousePointer <> MousePointer Then Screen.MousePointer = MousePointer
    'End If
    
    If MouseState = 2 Then ResetClick: LockX = -1
    If MouseState = 0 And LockX <> -1 Then LockX = -1
    
    If PublicTextBox.Visible <> TLockState Then
        PublicTextBox.Visible = TLockState
    End If
    If TLockState Then
        If PublicTextBox.Left <> TLockX Or PublicTextBox.top <> TLockY Then
            PublicTextBox.Move TLockX, TLockY, TLockW, TLockH
        End If
    End If
        
    tRet = False
    
CatchError:
    If Err.Number <> 0 Then
        Dim ErrNum As Long
        ErrNum = Err.Number
    '
    '©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
    '©§ * An error blocks the way !                                                                ©§
    '©§                                                                                                            ©§
    '©§                                                                                                            ©§
    '©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿
    ' ERROR  LV 1                     20/20  ¨€¨€¨€¨€¨€¨€¨€¨€¨€
    '©°©¤©¤©¤©¤©¤©¤©¤©¤©´  ©°©¤©¤©¤©¤©¤©¤©¤©¤©´   ©°©¤©¤©¤©¤©¤©¤©¤©¤©´  ©°©¤©¤©¤©¤©¤©¤©¤©¤©´
    '©¦     FIGHT    ©¦  ©¦       ACT      ©¦   ©¦      ITEM     ©¦  ©¦    MERCY   ©¦
    '©¸©¤©¤©¤©¤©¤©¤©¤©¤©¼  ©¸©¤©¤©¤©¤©¤©¤©¤©¤©¼   ©¸©¤©¤©¤©¤©¤©¤©¤©¤©¼  ©¸©¤©¤©¤©¤©¤©¤©¤©¤©¼
    '
    
        Dim temp As New Images, Brush As Long
        temp.Create TargetDC, GWW, GWH
        ProDraw.Draw temp.CompatibleDC, 0, 0
        GdipCreateSolidFill argb(120, 0, 0, 0), Brush
        GdipFillRectangle temp.Graphics, Brush, 0, 0, GWW, GWH
        ProFont.DrawText temp.Graphics, GWW / 2, GWH / 2 - 30, 0, 0, "We're soooorry .....", argb(255, 255, 108, 105), center, 24, Bold
        ProFont.DrawText temp.Graphics, GWW / 2, GWH / 2, 0, 0, ErrNum & " " & FriendError(ErrNum), argb(180, 255, 108, 105), center, 16, Regular
        GdipDeleteBrush Brush
        temp.Draw TargetDC, 0, 0
        Set ErrDraw = temp
        StopWork = True
    End If
End Sub
Private Sub Class_Initialize()
    ReDim Screens(0)
    Set ProCore = Me
    FadeDraw.Create TargetDC, GWW, GWH
End Sub
