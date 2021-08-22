VERSION 5.00
Begin VB.Form MainWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Epub Converter"
   ClientHeight    =   3648
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   7512
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3648
   ScaleWidth      =   7512
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton StartBtn 
      Caption         =   "StartWork"
      Height          =   480
      Left            =   624
      TabIndex        =   2
      Top             =   2808
      Width           =   5940
   End
   Begin VB.DriveListBox Drive1 
      Height          =   336
      Left            =   624
      TabIndex        =   1
      Top             =   312
      Width           =   5940
   End
   Begin VB.DirListBox Dir1 
      Height          =   1920
      Left            =   624
      TabIndex        =   0
      Top             =   624
      Width           =   5940
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Click()
    Me.Caption = "Epub Converter: " & Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.List(Drive1.ListIndex)
End Sub

Private Sub Form_Load()
    DataPath = App.Path & "\converter\"
End Sub

Private Sub StartBtn_Click()
    StartBtn.Enabled = False
    Dim f As String, l() As String
    f = Dir(Dir1.List(Dir1.ListIndex) & "\")
    ReDim l(0)
    Do While f <> ""
        If f Like "*.epub" Then
            ReDim Preserve l(UBound(l) + 1)
            l(UBound(l)) = f
        End If
        f = Dir()
        DoEvents
    Loop
    For i = 1 To UBound(l)
        StartBtn.Caption = "Process " & i & "/" & UBound(l) & " ..."
        DoEvents
        Encode Dir1.List(Dir1.ListIndex) & "\" & l(i)
        DoEvents
    Next
    StartBtn.Caption = "StartWork"
    StartBtn.Enabled = True
End Sub
