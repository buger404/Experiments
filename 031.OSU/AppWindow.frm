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
   StartUpPosition =   2  '��Ļ����
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
'   ����ģ������Emerald������ �����������ڣ�Ӧ�ô��ڣ� ģ��
'==================================================
'   ҳ�������
    Dim EC As GMan
'==================================================
'   �ڴ˴��������ҳ���������ģ���������
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
    StartEmerald Me.Hwnd, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY   '��ʼ��Emerald���ڴ˴������޸Ĵ��ڴ�С��
    MakeFont "΢���ź�"  '��������
   
    Set EC = New GMan   '����ҳ�������
    EC.Layered False
    
    '�����浵����ѡ�����浵key��������鿴Emerald��wiki
    'Set ESave = New GSaving
    'ESave.Create "emerald.test", "Emerald.test"
    
    '���������б���ѡ��
    'Set MusicList = New GMusicList
    'MusicList.Create App.path & "\music"

    '��ʼ��ʾ����
    Me.Show
    
    '�ڴ˴�ʵ�������ҳ�������
    '=============================================
    'ʾ����TestPage.cls
    '     Set TestPage = New TestPage
    '�������֣�Dim TestPage As TestPage
        Set AppPage = New AppPage
    '=============================================

    '���ûҳ�棨�ڴ˴�������Ϊ�������ҳ�棩
    EC.ActivePage = "AppPage"
    
    SetWindowPos Me.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    DrawTimer.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '��ֹ����
    CloseMark = True
    '�ͷ�Emerald��Դ
    EndEmerald
End Sub

'============================================================
' �¼�ӳ��
Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    UpdateMouse X, y, 1, button
End Sub
Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    If Mouse.State = 0 Then
        UpdateMouse X, y, 0, button
    Else
        Mouse.X = X: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    UpdateMouse X, y, 2, button
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    '�����ַ�����
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub
'============================================================
