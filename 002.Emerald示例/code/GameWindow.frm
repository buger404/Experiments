VERSION 5.00
Begin VB.Form GameWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   6672
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
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
Attribute VB_Name = "GameWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'   ҳ�������
    Dim EC As GMan
'==================================================
'   �ڴ˴��������ҳ����ģ������
    Dim GamePage As GamePage
'==================================================

Private Sub DrawTimer_Timer()
    '����
    EC.Display
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '�����ַ�����
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub

Private Sub Form_Load()
    '��ʼ��Emerald���ڴ˴������޸Ĵ��ڴ�СӴ~��
    StartEmerald Me.Hwnd, 805, 556
    '��������
    MakeFont "΢���ź�"
    '����ҳ�������
    Set EC = New GMan
    
    '�����浵����ѡ��
    'Set ESave = New GSaving
    'ESave.Create "emerald.test", "Emerald.test"
    
    '���������б�
    Set MusicList = New GMusicList
    MusicList.Create App.Path & "\music"

    '��ʼ��ʾ
    Me.Show
    DrawTimer.Enabled = True
    
    '�ڴ˴���ʼ�����ҳ��
    '=============================================
    'ʾ����TestPage.cls
    '     Set TestPage = New TestPage
    '�������֣�Dim TestPage As TestPage
        Set GamePage = New GamePage
    '=============================================

    '���ûҳ��
    EC.ActivePage = "GamePage"
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
    '���������Ϣ
    UpdateMouse X, Y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    '���������Ϣ
    If Mouse.state = 0 Then
        UpdateMouse X, Y, 0, button
    Else
        Mouse.X = X: Mouse.Y = Y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
    '���������Ϣ
    UpdateMouse X, Y, 2, button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '��ֹ����
    DrawTimer.Enabled = False
    '�ͷ�Emerald��Դ
    EndEmerald
End Sub
