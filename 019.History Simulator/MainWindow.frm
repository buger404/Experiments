VERSION 5.00
Begin VB.Form MainWindow 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'None
   Caption         =   "History Simulator"
   ClientHeight    =   6105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   575
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   375
      Left            =   5100
      TabIndex        =   10
      Text            =   "1"
      Top             =   4800
      Width           =   3015
   End
   Begin VB.PictureBox Draw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   450
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   375
      Left            =   450
      TabIndex        =   7
      Text            =   "/"
      Top             =   4800
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   2235
      Left            =   450
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2100
      Width           =   7665
   End
   Begin VB.TextBox NationText 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   375
      Left            =   450
      TabIndex        =   2
      Text            =   $"MainWindow.frx":0000
      Top             =   1350
      Width           =   7665
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ͼ�������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BAB539&
      Height          =   285
      Left            =   5100
      TabIndex        =   11
      Top             =   4350
      Width           =   1365
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ID��/Ϊ�����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BAB539&
      Height          =   285
      Left            =   450
      TabIndex        =   8
      Top             =   4350
      Width           =   1665
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E8E8E8&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   285
      Left            =   6900
      TabIndex        =   6
      Top             =   5400
      Width           =   1200
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007C7976&
      Height          =   285
      Left            =   8400
      TabIndex        =   5
      Top             =   150
      Width           =   105
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BAB539&
      Height          =   285
      Left            =   450
      TabIndex        =   3
      Top             =   1800
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "History Simulator"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   285
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������ƣ���;������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BAB539&
      Height          =   285
      Left            =   450
      TabIndex        =   0
      Top             =   900
      Width           =   1800
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oShadow As New aShadow
Dim CultureLevel(100) As String
Private Type Nation
    name As String
    Aera As Single
    People As Single
    CLevel As Long
    LoveBOSS As Single
    Money As Single
End Type
Dim DrawList() As Single
Dim HYear As Long, HMonth As Long, HDay As Long
Sub StartHistory()
    On Error Resume Next

    For i = 0 To 100
        CultureLevel(i) = GetRandomName & "����"
    Next

    Dim ID As Long
    Randomize
    ID = IIf(Text2.Text = "/", Int(Rnd * 10000), Val(Text2.Text))
    Text1.Text = "����һ��IDΪ" & ID & "��ƽ������Ĺ���" & vbCrLf & "�������ĵ����ϣ�����"
    Randomize (ID)
    HYear = 2000: HMonth = 1: HDay = 1
    Dim Nations() As Nation, temp() As String, RmA As Single
    RmA = 1
    temp = Split(NationText.Text, ";")
    ReDim Nations(UBound(temp))
    For i = 0 To UBound(temp)
        Nations(i).name = temp(i)
        Nations(i).Aera = Rnd * RmA
        RmA = RmA - Nations(i).Aera
        Nations(i).People = Int(Rnd * 200000000 + 200000000)
        Nations(i).Money = Int(Rnd * (100000 * Nations(i).People) + (100000 * Nations(i).People))
        Nations(i).LoveBOSS = 0.5
        Text1.Text = Text1.Text & temp(i) & " "
    Next
    
    Text1.Text = Text1.Text & vbCrLf
    
    Dim DayMax(12) As Long, Action As Long, Action2 As Long, Success As Single
    DayMax(1) = 31: DayMax(2) = 28: DayMax(3) = 31: DayMax(4) = 30
    DayMax(5) = 31: DayMax(6) = 30: DayMax(7) = 31: DayMax(8) = 31
    DayMax(9) = 30: DayMax(10) = 31: DayMax(11) = 30: DayMax(12) = 31
    
    Do
        For i = 0 To UBound(Nations)
            If i <= UBound(Nations) Then
            Action = Int(Rnd * 50000)
            Select Case Action
                Case 23333 '����
                    If Int(Rnd * 30) = 13 And Nations(i).People > 100000 Then
                        Action2 = Int(Rnd * (UBound(Nations) + 1))
                        If Action2 > UBound(Nations) Then Action2 = UBound(Nations)
                        If Action2 <> i Then
                            Success = (Nations(i).Money / Nations(Action2).Money) * (Nations(i).People / Nations(Action2).People)
                            If Success < 1 Then
                                NewHistory Nations(i).name & "��" & Nations(Action2).name & "���������ԣ���ʧ�ܸ��ա�"
                                Nations(i).LoveBOSS = Nations(i).LoveBOSS - 0.05
                                Nations(i).Money = Nations(i).Money - Nations(i).Money * (Rnd * 0.3 + 0.3)
                                Nations(i).People = Nations(i).People - Nations(i).People * (Rnd * 0.3 + 0.3)
                                If Int(Rnd * 100) = 44 Then
                                    NewHistory Nations(i).name & "����ͬ��" & Nations(Action2).name & "�����" & GetRandomName & "��ƽ����Լ"
                                    Nations(i).LoveBOSS = Nations(i).LoveBOSS - 0.15
                                    Nations(i).Money = Nations(i).Money - Nations(i).Money * (Rnd * 0.5 + 0.5)
                                End If
                            Else
                                NewHistory Nations(i).name & "��" & Nations(Action2).name & "���������ԣ�" & Nations(Action2).name & "������"
                                Nations(i).LoveBOSS = Nations(i).LoveBOSS - 0.05
                                Nations(i).Money = (Nations(i).Money - Nations(i).Money * (Rnd * 0.3 + 0.3)) + (Nations(Action2).Money - Nations(Action2).Money * (Rnd * 0.3 + 0.3))
                                Nations(i).People = (Nations(i).People - Nations(i).People * (Rnd * 0.3 + 0.3)) + (Nations(Action2).People - Nations(Action2).People * (Rnd * 0.3 + 0.3))
                                Nations(i).Aera = Nations(i).Aera + Nations(Action2).Aera
                                For s = Action2 To UBound(Nations) - 1
                                    Nations(s) = Nations(s + 1)
                                Next
                                ReDim Preserve Nations(UBound(Nations) - 1)
                                If UBound(Nations) = 0 Then
                                    NewHistory "����" & Nations(0).name & "ͳһ��"
                                    Exit Do
                                End If
                                GoTo last
                            End If
                        End If
                    End If
                Case 44444 '����
                    Action2 = Int(Rnd * 30)
                    Select Case Action2
                        Case 0
                        NewHistory Nations(i).name & "��ͳ����������" & GetRandomName & "���ߣ������ܵ��ش졣"
                        Nations(i).Money = Nations(i).Money - Nations(i).Money * (Rnd * 0.3 + 0.3)
                        Nations(i).People = Nations(i).People - Nations(i).People * (Rnd * 0.05 + 0.05)
                        Case 1
                        NewHistory Nations(i).name & "��ͳ��������ʵ��" & GetRandomName & "�������������������������Ĳ�����"
                        Nations(i).Money = Nations(i).Money - Nations(i).Money * (Rnd * 0.1 + 0.1)
                        Nations(i).People = Nations(i).People - Nations(i).People * (Rnd * 0.05 + 0.05)
                        Case 2
                        If Nations(i).CLevel > 0 Then
                            NewHistory Nations(i).name & "��ͳ���߸���" & CultureLevel(Nations(i).CLevel - 1) & "��������������Ĳ�����"
                            Nations(i).Money = Nations(i).Money - Nations(i).Money * (Rnd * 0.1 + 0.1)
                            Nations(i).People = Nations(i).People - Nations(i).People * (Rnd * 0.05 + 0.05)
                            Nations(i).CLevel = Nations(i).CLevel - 1
                        End If
                    End Select
                    Nations(i).LoveBOSS = Nations(i).LoveBOSS - 0.1
                Case 6666 '��ʢ
                    Action2 = Int(Rnd * 90)
                    Select Case Action2
                        Case 33
                        NewHistory Nations(i).name & "��ͳ����������" & GetRandomName & "���ߣ�����Ѹ�ٷ�չ��"
                        Nations(i).Money = Nations(i).Money + Nations(i).Money * (Rnd * 0.3 + 0.3)
                        Nations(i).People = Nations(i).People + Nations(i).People * (Rnd * 0.05 + 0.05)
                        NewHistory Nations(i).name & "����Ʋ���Լ" & format(Int(Nations(i).Money / 1000000) / 100, "0.00") & "��Ԫ���˿ڴﵽ" & format(Int(Nations(i).People / 1000000) / 100, "0.00") & "���ˡ�"
                        Case 66
                            If Nations(i).CLevel < 100 Then
                                NewHistory Nations(i).name & "������һ�����Ļ��˶���" & CultureLevel(Nations(i).CLevel + 1) & "�����γɡ�"
                                Nations(i).Money = Nations(i).Money - Nations(i).Money * (Rnd * 0.1 + 0.1)
                                Nations(i).People = Nations(i).People + Nations(i).People * (Rnd * 0.05 + 0.05)
                                Nations(i).CLevel = Nations(i).CLevel + 1
                            End If
                    End Select
                    Nations(i).LoveBOSS = Nations(i).LoveBOSS + 0.1
            End Select

            Dim OranName As String
            If Nations(i).LoveBOSS < 0.2 Then '�Ʒ�~
                OranName = Nations(i).name
                'Nations(i).name = GetRandomName & "��"
                NewHistory OranName & "�������̲���ȥ�ˣ�����������~" '& Nations(i).name
                Nations(i).LoveBOSS = 0.5
                Nations(i).Money = Nations(i).Money - Nations(i).Money * (Rnd * 0.1 + 0.1)
                Nations(i).People = Nations(i).People - Nations(i).People * (Rnd * 0.1 + 0.1)
            End If
            If Nations(i).Money / Nations(i).People < 5000 And Nations(i).LoveBOSS < 0.35 Then   '��������
                OranName = Nations(i).name
                'Nations(i).name = GetRandomName & "��"
                NewHistory OranName & "�ľ��������»����˾��Ʋ�����" & Int(Nations(i).Money / Nations(i).People) & " Ԫ���������񷢶������ȡ��Ȩ��" '& Nations(i).name
                Nations(i).LoveBOSS = 0.5
                Nations(i).Money = Nations(i).Money - Nations(i).Money * (Rnd * 0.1 + 0.1)
                Nations(i).People = Nations(i).People - Nations(i).People * (Rnd * 0.1 + 0.1)
            End If
            If Nations(i).People < 100 And UBound(Nations) > 0 Then   'û����
                NewHistory OranName & "���˿ڹ��٣�����������״̬���ܿ�������������" '& Nations(i).name
                For s = i To UBound(Nations) - 1
                    Nations(s) = Nations(s + 1)
                Next
                ReDim Preserve Nations(UBound(Nations) - 1)
                GoTo last
            End If
            If Nations(i).People > 1000000000 Then   'û����
                If Int(Rnd * 10000) = 444 Then
                    NewHistory OranName & "���˿ڹ��࣬������ʼ�����˿ڵ�������" '& Nations(i).name
                    Nations(i).People = Nations(i).People - Int(Nations(i).People * 0.1)
                Else
                    NewHistory OranName & "���˿ڹ��࣬������Ȼʹ��ɱ�˵ķ�ʽ�����˿ڣ�" '& Nations(i).name
                    Nations(i).People = Nations(i).People - Int(Nations(i).People * 0.5)
                    Nations(i).LoveBOSS = Nations(i).LoveBOSS - 0.2
                End If
            End If
            
            Nations(i).People = Nations(i).People + Int(Nations(i).People * 0.000000014)
            Nations(i).Money = Nations(i).Money + Int(Nations(i).Money * 0.00000014)
            End If
            
last:
        Next
        HDay = HDay + 1
        If HDay > DayMax(HMonth) Then HDay = 1: HMonth = HMonth + 1
        If HMonth > 12 Then
            HYear = HYear + 1: HMonth = 1
            ReDim Preserve DrawList(HYear - 1)
            DrawList(HYear - 1) = Val(format(Int(Nations(Val(Text3.Text)).Money / 1000000) / 100, "0.00"))
        End If
        DoEvents
    Loop
    
    Draw.Cls
    Draw.Move 0, 0, (HYear - 1) / 100 * 20, Me.ScaleHeight
    
    Dim Heighest As Single
    For i = 2000 To HYear - 1 Step 100
        If DrawList(i) > Heighest Then Heighest = DrawList(i)
    Next
    
    Draw.Visible = True
    
    s = 0
    For i = 2000 To HYear - 1 Step 100
        If i <> 2000 Then
            Draw.Line ((s - 1) * 20, (Me.ScaleHeight - (DrawList(i - 100) / Heighest) * Me.ScaleHeight))-(s * 20, (Me.ScaleHeight - (DrawList(i) / Heighest) * Me.ScaleHeight))
            Draw.Refresh
        End If
        s = s + 1
    Next
    
    SavePicture Draw.Image, App.Path & "\output.bmp"
    Draw.Visible = False
End Sub
Sub NewHistory(ByVal Text As String)
    Text1.Text = Text1.Text & HYear & "��" & HMonth & "��" & HDay & "��     " & Text & vbCrLf
    Text1.SelStart = Len(Text1.Text) - Text1.SelLength
End Sub
Private Sub Form_Load()

    With oShadow
        If .Shadow(Me) Then
            .Color = RGB(0, 0, 0)
            .Depth = 18
            .Transparency = 11
        End If
    End With
    
End Sub

Private Sub Label2_Click()
    Call StartHistory
End Sub

Private Sub Label7_Click()
    Unload Me
End Sub

Function GetRandomName() As String
    Dim WaitChr As String, ChrList As String
    ChrList = "��Ԫ������������Ӣ˼�������������������������ɭɢ����Ƥ����������"
    For i = 0 To Int(Rnd * 2 + 3)
        WaitChr = Mid(ChrList, Int(Rnd * (Len(ChrList) - 1) + 1), 1)
        GetRandomName = GetRandomName & WaitChr
        WaitChr = ""
    Next
End Function

