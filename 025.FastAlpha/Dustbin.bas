Attribute VB_Name = "Dustbin"
Dim MyMembers() As Object, MyMembers2()
Sub AddMember(Member As Object)
    ReDim Preserve MyMembers(UBound(MyMembers) + 1)
    Set MyMembers(UBound(MyMembers)) = Member
End Sub
Sub AddImageMember(Member As Long)
    ReDim Preserve MyMembers2(UBound(MyMembers2) + 1)
    MyMembers2(UBound(MyMembers2)) = Member
End Sub
Sub DelImageMember(Member As Long)
    For i = 1 To UBound(MyMembers2)
        If MyMembers2(i) = Member Then
            MyMembers2(i) = MyMembers2(UBound(MyMembers2))
            ReDim Preserve MyMembers2(UBound(MyMembers2) - 1)
            Exit For
        End If
    Next
End Sub
Function IsDelAllImage() As Boolean
    If UBound(MyMembers2) = 0 Then IsDelAllImage = True
End Function
Sub InitDustbin()
    ReDim MyMembers(0)
    ReDim MyMembers2(0)
End Sub
Sub DoClearing()
    On Error Resume Next
    
    For i = 1 To UBound(MyMembers2)
        If MyMembers2(i) <> 0 Then
            Debug.Print Now, "正在回收第" & i & "项（1/2）：Gdip Image"
            GdipDisposeImage MyMembers2(i)
        End If
    Next
    ReDim MyMembers2(0)
    
    For i = 1 To UBound(MyMembers)
        Debug.Print Now, "正在回收第" & i & "项（2/2）：" & TypeName(MyMembers(i))
        MyMembers(i).Dispose
    Next
    
End Sub

Public Sub CreateImage(ByVal filename As Long, bitmap As Long)
    Dim Img As Long
    GdipCreateBitmapFromFile filename, Img
    AddImageMember Img
    bitmap = Img
End Sub
Public Sub DelImage(bitmap As Long)
    GdipDisposeImage bitmap
    DelImageMember bitmap
End Sub
