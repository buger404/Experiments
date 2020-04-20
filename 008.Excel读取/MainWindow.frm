VERSION 5.00
Begin VB.Form MainWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Experiments"
   ClientHeight    =   5208
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   11028
   Icon            =   "MainWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   919
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox outputBox 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0C0&
      Height          =   1764
      Left            =   8736
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3072
      Width           =   2028
   End
   Begin VB.Label DateText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Date]"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   9432
      TabIndex        =   2
      Top             =   2616
      Width           =   648
   End
   Begin VB.Label TitleText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Experiment"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CEDB1A&
      Height          =   276
      Left            =   9048
      TabIndex        =   1
      Top             =   2280
      Width           =   1320
   End
   Begin VB.Image testIcon 
      Height          =   1536
      Left            =   8928
      Picture         =   "MainWindow.frx":000C
      Top             =   504
      Width           =   1536
   End
   Begin VB.Label StartBtn 
      Alignment       =   2  'Center
      BackColor       =   &H00CEDB1A&
      Caption         =   "Launch"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   276
      Left            =   240
      TabIndex        =   0
      Top             =   4680
      Width           =   1236
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Caption = "Experiment " & Split(Split(App.Path, ".")(0), "Experiments\")(1)
    TitleText.Caption = Me.Caption
    DateText.Caption = Now
    Set Me.Icon = testIcon.Picture
End Sub
Private Sub Outputs(Text As String)
    outputBox.Text = outputBox.Text & Now & "   " & Text & vbCrLf
    outputBox.SelLength = 1
    outputBox.SelStart = Len(outputBox.Text)
End Sub
Function ReadExcel(ByVal Path As String, Optional ByVal SheetIndex As Integer = 1) As String()
    Dim App As Object, WBook As Object, WB As Object, Sheet As Object

    Set App = CreateObject("Excel.Application")
    Set WBook = App.WorkBooks
    Set WB = App.WorkBooks.Open(Path)
    Set Sheet = WB.WorkSheets(SheetIndex)
    
    Dim Row As Integer, Line As Integer
    Dim Data(), Ret() As String
    
    Data = Sheet.UsedRange.Value
    ReDim Ret(UBound(Data, 2) - 1, UBound(Data, 1) - 1)
    
    For Row = 1 To UBound(Data, 1)
        For Line = 1 To UBound(Data, 2)
            Ret(Line - 1, Row - 1) = Data(Row, Line)
        Next
    Next
    
    App.Quit
    Set Sheet = Nothing
    Set WBook = Nothing
    Set App = Nothing
    
    ReadExcel = Ret
End Function
Private Sub Test()
    'Test Code
    Dim App As Object, WBook As Object, WB As Object, Sheet As Object
    
    Me.Cls
    Me.CurrentX = 100: Me.CurrentY = 30: Print "正在打开Excel..."
    Me.Refresh: DoEvents
    
    Set App = CreateObject("Excel.Application")
    Set WBook = App.WorkBooks
    Set WB = App.WorkBooks.Open("E:\Experiments\008.Excel读取\Test.xlsx")
    Set Sheet = WB.WorkSheets(1)
    
    Dim Row As Integer, Line As Integer
    Dim Data()
    
    Me.Cls
    Me.CurrentX = 100: Me.CurrentY = 30: Print "正在读取Excel..."
    Me.Refresh: DoEvents
    
    Data = Sheet.UsedRange.Value
    Me.Cls
    For Row = 1 To UBound(Data, 1)
        For Line = 1 To UBound(Data, 2)
            Me.Line (Line * 100, Row * 30 + 30)-(Line * 100 + 100, Row * 30 + 30 + 30), IIf((Row Mod 2 = 1), RGB(242, 242, 242), RGB(255, 255, 255)), BF
            Me.CurrentX = Line * 100 + 10
            Me.CurrentY = Row * 30 + 30 + 5
            Print Data(Row, Line)
        Next
    Next
    
    For i = 1 To WB.WorkSheets.Count
        Me.Line (i * 100, 20)-(i * 100 + 100, 20 + 30), RGB(232, 232, 232), BF
        Me.CurrentX = i * 100 + 10
        Me.CurrentY = 20 + 5
        Print WB.WorkSheets(i).Name
    Next
    
    App.Quit
    Set Sheet = Nothing
    Set WBook = Nothing
    Set App = Nothing
End Sub
Private Sub StartBtn_Click()
    Dim a() As String
    a = ReadExcel("D:\Experiments\008.Excel读取\Test.xlsx")
    'Call Test
End Sub
