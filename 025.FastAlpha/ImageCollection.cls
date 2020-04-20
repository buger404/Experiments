VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImageCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Images() As Images
Sub AddImage(Img As Object)
    ReDim Preserve Images(UBound(Images) + 1)
    Set Images(UBound(Images)) = Img
End Sub
Property Get Count() As Long
    Count = UBound(Images)
End Property
Sub LoadDir(ByVal Path As String, Optional ByVal Scales As Single = 1)
    Dim file As String
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    file = Dir(Path)
    
    Do While file <> ""
        ReDim Preserve Images(UBound(Images) + 1)
        Set Images(UBound(Images)) = New Images
        Images(UBound(Images)).CreateFromPath TargetDC, Path & file, Scales
        Images(UBound(Images)).name = file
        file = Dir()
        DoEvents
    Loop
End Sub
Public Property Get Image(file As String) As Images
    For i = 1 To UBound(Images)
        If Images(i).name = file Then Set Image = Images(i): Exit For
    Next
End Property
Public Property Get ImageByIndex(Index As Integer) As Images
    Set ImageByIndex = Images(Index)
End Property
Private Sub Class_Initialize()
    ReDim Images(0)
End Sub
