VERSION 5.00
Begin VB.Form AppWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ƽ��"
   ClientHeight    =   6672
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9660
   LinkTopic       =   "AppWindow"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   StartUpPosition =   2  '��Ļ����
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
    Dim ExitMark As Boolean
    Public Slowed As Boolean
'==================================================

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then Slowed = Not Slowed: ECore.FancyMode = Slowed: ECore.FancyDeepth = 0.2
End Sub

Private Sub Form_Load()
    StartEmerald Me.Hwnd, 1320, 800  '��ʼ��Emerald���ڴ˴������޸Ĵ��ڴ�С��
    'ScaleGame 0.8, ScaleSuitable
    
    MakeFont "΢���ź�"  '��������
   
    Set EC = New GMan   '����ҳ�������
    'EC.FancyMode = True

    '�����浵����ѡ�����浵key��������鿴Emerald��wiki
    'Set ESave = New GSaving
    'ESave.Create "emerald.test", "Emerald.test"
    
    '���������б���ѡ��
    'Set MusicList = New GMusicList
    'MusicList.Create App.path & "\music"
    
    '�ڴ˴�ʵ�������ҳ�������
    '=============================================
    'ʾ����TestPage.cls
    '     Set TestPage = New TestPage
    '�������֣�Dim TestPage As TestPage
        Set AppPage = New AppPage
    '=============================================

    '���ûҳ�棨�ڴ˴�������Ϊ�������ҳ�棩
    EC.ActivePage = "AppPage"
    
    Me.Show
    Do While Not ExitMark
        EC.Display
        DoEvents
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '��ֹ����
    ExitMark = True
    '�ͷ�Emerald��Դ
    EndEmerald
    End
End Sub

Private Sub DrawTimer_Timer()
    '���ƽ��沢ˢ�´��ڻ���
    EC.Display
End Sub

'============================================================
' �¼�ӳ��
Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
    '���������Ϣ
    UpdateMouse X, Y, 1, button
End Sub
Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    '���������Ϣ
    If Mouse.State = 0 Then
        UpdateMouse X, Y, 0, button
    Else
        Mouse.X = X: Mouse.Y = Y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
    '���������Ϣ
    UpdateMouse X, Y, 2, button
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    '�����ַ�����
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub
'============================================================
