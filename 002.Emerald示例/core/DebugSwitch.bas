Attribute VB_Name = "DebugSwitch"
'   Emerald ������

'======================================================
'   �������Ѿ�Ǩ��
'   ���������ת��Builder�еġ����á�
'======================================================


'==============================================================================
'   �汾����ע������
'   1.�ı���ļ���
'     ���������Ϸ���ڼ���:
'     Private Sub Form_KeyPress(KeyAscii As Integer)
'         '�����ַ�����
'         If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
'     End Sub
'==============================================================================
'   1.�������ܵļ���
'     �������ÿһ����Ϸҳ��ģ����룺
'        Public Sub AnimationMsg(id As String, msg As String)
'            '������Ϣ����
'        End Sub
'   2.�浵���޸�
'     ��Ҫ���ṩ�浵����Կ�����ڴ����浵�ĵڶ��������Ѿ������Կ����ע�⣡
'     ***���ڴ��������Ʊ��������Կ����ֹ�����Ϸ�浵���޸ġ�
'     ***��Ҫ�����޸���Կ�������ᵼ�¾ɵĴ浵��������
'     ***����޷�ȷ����Կ����������������У�����debug.print GetBMKey
'==============================================================================
'   1.������޸�
'      ���������Ϸ�����ҵ����´���
'       Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
'           If Mouse.State = 0 Then UpdateMouse X, Y, 0, button
'       End Sub
'      *****�����޸�Ϊ��
'       Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
'           If Mouse.State = 0 Then
'               UpdateMouse X, Y, 0, button
'           Else
'               Mouse.X = X: Mouse.Y = Y
'           End If
'       End Sub
'   2.������ջ����޸�
'     ������Ļ�ͼ���̼��룺
'       Page.Clear
'==============================================================================
'   1.��Դ���صĸı�
'     ���Page.NewImagesǨ�Ƶ�Page.Res.NewImages
'==============================================================================
'   1.���ش���ĸı�
'     ���ڿ���LOGO�ļ��룬
'     �����������ҳ��Ϳ�������Timer�Ĵ���ת�Ƶ�����ҳ��֮ǰ��������һ�С�Me.Show��
'   *�ò��������Բ���Emerald�ṩ��ģ��
'   2.Emerald��ʼ���ĸı�
'     �����뵽Emerald�Ĵ��ڴ�С��������Emerald��������һ�Ρ�
'==============================================================================




'======================================================
'   ���棺��Ҫ�޸����д���
    Public DebugMode As Integer, DisableLOGO As Integer, HideLOGO As Integer, UpdateCheckInterval As Long, UpdateTimeOut As Long
    Public Debug_focus As Boolean, Debug_pos As Boolean, Debug_data As Boolean
'======================================================
