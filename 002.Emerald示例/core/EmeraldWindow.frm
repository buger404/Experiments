VERSION 5.00
Begin VB.Form EmeraldWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Emerald Screen"
   ClientHeight    =   2316
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   193
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   302
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox DisplayBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   948
      Left            =   1032
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   117
      TabIndex        =   0
      Top             =   960
      Width           =   1404
   End
   Begin VB.Timer UpdateTimer 
      Interval        =   16
      Left            =   240
      Top             =   144
   End
End
Attribute VB_Name = "EmeraldWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Emerald ��ش���
Public Charge As Object, AcceptMark As Boolean, AcceptMark2 As Boolean
Dim OpenTime As Long, WinAlpha As Long

Private Sub Form_Click()
    Call Accept
End Sub
Public Sub Accept()
    If AcceptMark2 Then Exit Sub
    AcceptMark2 = True
    OpenTime = GetTickCount
    Do While WinAlpha > 0
        WinAlpha = 255 - (GetTickCount - OpenTime) / 500 * 255
        If WinAlpha < 0 Then WinAlpha = 0
        SetLayeredWindowAttributes Me.Hwnd, 0, WinAlpha, LWA_ALPHA
        Sleep 10: DoEvents
    Loop
    Dim f As Object
    For Each f In VB.Forms
        f.Enabled = True
    Next
    AcceptMark = True
End Sub
Private Sub DisplayBox_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    UpdateMouse x, y, 1, button
End Sub

Private Sub DisplayBox_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If Mouse.state = 0 Then UpdateMouse x, y, 0, button
End Sub

Private Sub DisplayBox_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    UpdateMouse x, y, 2, button
End Sub
Private Sub Form_Load()
    
    Me.Move 0, 0, Screen.Width, Screen.Height
    AcceptMark = False
    
    Dim rtn As Long
    rtn = GetWindowLongA(Me.Hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetWindowLongA Me.Hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes Me.Hwnd, 0, WinAlpha, LWA_ALPHA
    OpenTime = GetTickCount
        
End Sub

Private Sub UpdateTimer_Timer()
    If Charge Is Nothing Then Exit Sub
    Charge.Page.Clear
    Charge.Page.Update
    Charge.Page.Display DisplayBox.hdc
    If Mouse.state = 2 Then Mouse.state = 0
End Sub

Public Sub NewFocusWindow(w As Long, h As Long, ch As Object)
    Me.Show
    
    Dim Sc As Long, Dh As Long
    Dh = GetDesktopWindow: Sc = GetDC(Dh)
    
    Dim x As Long, y As Long
    Dim G As Long, b As Long, img As Long, g2 As Long
    PoolCreateFromHdc Me.hdc, G
    
    DisplayBox.Width = w: DisplayBox.Height = h
    DisplayBox.Move Me.ScaleWidth / 2 - w / 2, Me.ScaleHeight / 2 - h / 2
    x = DisplayBox.Left: y = DisplayBox.top + 10
    
    BitBlt Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Sc, 0, 0, vbSrcCopy
    ReleaseDC Dh, Sc
    
    BlurTo Me.hdc, Me.hdc, Me, 100
    
    PoolCreateSolidFill argb(20, 0, 116, 217), b

    GdipFillRectangle G, b, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    PoolDeleteBrush b
    
    GdipCreateBitmapFromScan0 Me.ScaleWidth, Me.ScaleHeight, ByVal 0, PixelFormat32bppARGB, ByVal 0, img
    GdipGetImageGraphicsContext img, g2
    
    PoolCreateSolidFill argb(100, 0, 0, 0), b
    
    GdipFillRectangle g2, b, x, y, w + 1, h + 1
    BlurImg img, 30
    GdipDrawImage G, img, 0, 0
    
    PoolDeleteBrush b
    PoolDisposeImage img
    
    PoolDeleteGraphics G

    Set Charge = ch
    Charge.Page.Clear
    Charge.Page.Update
    Charge.Page.Display DisplayBox.hdc
    
    Me.Refresh

    OpenTime = GetTickCount
    Do While WinAlpha < 255
        WinAlpha = (GetTickCount - OpenTime) / 500 * 255
        If WinAlpha > 255 Then WinAlpha = 255
        SetLayeredWindowAttributes Me.Hwnd, 0, WinAlpha, LWA_ALPHA
        Sleep 10: DoEvents
    Loop
End Sub
