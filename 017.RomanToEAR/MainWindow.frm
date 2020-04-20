VERSION 5.00
Begin VB.Form MainWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ò»¼ü¿Õ¶ú"
   ClientHeight    =   5592
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   5772
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5592
   ScaleWidth      =   5772
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox TI 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   5172
      Left            =   24
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   24
      Width           =   5772
   End
   Begin VB.Label Btn 
      Alignment       =   2  'Center
      BackColor       =   &H00CEDB1A&
      Caption         =   "GO"
      ForeColor       =   &H00FFFFFF&
      Height          =   276
      Left            =   2112
      TabIndex        =   1
      Top             =   5280
      Width           =   1596
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn_Click()
    Dim Code As String
    TI.Text = LCase(TI.Text)
    
    Open App.Path & "\Library.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, Code
        If Len(Split(Code, "=")(0)) > 1 Then TI.Text = Replace(TI.Text, Split(Code, "=")(0), Split(Code, "=")(1))
    Loop
    Close #1
    
    Open App.Path & "\Library.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, Code
        If Len(Split(Code, "=")(0)) = 1 Then TI.Text = Replace(TI.Text, Split(Code, "=")(0), Split(Code, "=")(1))
    Loop
    Close #1
End Sub

