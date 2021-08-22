Attribute VB_Name = "EpubEncoder"
Public DataPath As String
Sub EncodeUTF(Path As String)
    Dim loadStream As ADODB.stream
    Dim saveStream As ADODB.stream
    
    Set loadStream = New ADODB.stream
    Set saveStream = New ADODB.stream

    With saveStream
        .Mode = 8
        .Open
        .Charset = "chinese"
        With loadStream
            .Open
            .Charset = "utf-8"
            .LoadFromFile Path
            .CopyTo saveStream
            .Close
        End With
        .SaveToFile Path & "encode"
        .Close
    End With
    
    Set loadStream = Nothing
    Set saveStream = Nothing
End Sub
Sub Encode(Path As String)
    Dim FSO As Object, name As String, temp() As String, f As String, Data As String, d As String
    Dim dirs(1) As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Dir(App.Path & "\temp\", vbDirectory) <> "" Then FSO.DeleteFolder App.Path & "\temp"
    
    Sleep 1000
    
    Do While Dir(App.Path & "\temp\", vbDirectory) <> ""
        DoEvents
    Loop
    
    Shell App.Path & "\winrar\winrar.exe x -ibck """ & Path & """ *.* """ & App.Path & "\temp\"""
    
    Sleep 1000
    
    Do While Dir(App.Path & "\temp\OEBPS\Images\", vbDirectory) = "" And Dir(App.Path & "\temp\OPS\Images\", vbDirectory) = ""
        DoEvents
    Loop
    
    If Dir(App.Path & "\temp\OEBPS\Images\*.*") = "" Then
        dirs(0) = App.Path & "\temp\OPS\Images"
        dirs(1) = App.Path & "\temp\OPS\"
    Else
        dirs(0) = App.Path & "\temp\OEBPS\Images"
        dirs(1) = App.Path & "\temp\OEBPS\Text\"
    End If
    
    temp = Split(Path, "\")
    name = Replace(temp(UBound(temp)), ".epub", "")
    name = Replace(name, "?", "")
    
    If Dir(DataPath & "\" & name, vbDirectory) = "" Then MkDir DataPath & "\" & name
    If Dir(DataPath & "\" & name & "\Images", vbDirectory) = "" Then MkDir DataPath & "\" & name & "\Images"
    If Dir(DataPath & "\" & name & "\Passage", vbDirectory) = "" Then MkDir DataPath & "\" & name & "\Passage"
    
    CopyInto dirs(0), DataPath & "\" & name & "\Images"
    
    Dim t As String, t2 As String, t3 As String
    f = Dir(dirs(1))
    Do While f <> ""
        Data = "": t = "": d = ""
        EncodeUTF dirs(1) & f
        
        Open dirs(1) & f & "encode" For Input As #1
        Do While Not EOF(1)
            Line Input #1, t2
            d = d & t2 & vbCrLf
        Loop
        Close #1
        temp = Split(d, "</title>")
        If UBound(temp) = 1 Then
            t = Split(temp(0), "<title>")(1)
            temp(1) = Replace(Replace(Replace(temp(1), ">", "<"), "&nbsp;", ""), "<br /<", ">br>")
            temp = Split(temp(1), "<")
        Else
            temp(0) = Replace(Replace(Replace(temp(0), ">", "<"), "&nbsp;", ""), "<br /<", ">br>")
            temp = Split(temp(0), "<")
        End If

        For i = 2 To UBound(temp) Step 2
            If Trim(Replace(temp(i), vbCrLf, "")) <> "" Then
                t2 = Left(temp(i - 1), 1): t3 = Left(temp(i - 1), 3)
                temp(i) = Trim(Replace(Replace(Replace(Replace(Replace(temp(i), vbCrLf, ""), ">br>", vbCrLf), "&lt;", "<"), "&gt;", ">"), "&#160;", "    "))
                Data = Data & IIf(t2 = "p" Or t2 = "h" Or t3 = "div", vbCrLf, "") & temp(i)
                If t = "" Then t = temp(i)
            End If
        Next
        If Data <> "" And Data <> t Then
            Dim ttt() As String
            ttt = Split(Data, vbCrLf)
            For i = 0 To UBound(ttt)
                t = ttt(i)
                If t <> "" Then Exit For
            Next
            t = Replace(Replace(Replace(Replace(t, "/", " "), ",", "£¬"), ":", "£º"), vbCrLf, "")
            Open DataPath & "\" & name & "\Passage\" & Replace(Trim(t), "?", " ") & ".txt" For Output As #1
            Print #1, Data
            Close #1
        End If
        Kill dirs(1) & f & "encode"
        f = Dir()
        DoEvents
    Loop
    
    If Dir(App.Path & "\temp\", vbDirectory) <> "" Then FSO.DeleteFolder App.Path & "\temp"
End Sub
Sub CopyInto(Src As String, Dst As String)
    Dim f As String
    f = Dir(Src & "\")
    Do While f <> ""
        FileCopy Src & "\" & f, Dst & "\" & f
        f = Dir()
    Loop
End Sub

