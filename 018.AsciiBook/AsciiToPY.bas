Attribute VB_Name = "AsciiToPY"
Public Type PYBook
    Ascii As Long
    PY As String
End Type
Function AscToPY(ByVal Book As String, ByVal CH As String, Optional WithSpace As Boolean = True, Optional ReplaceBook As String) As String
    Dim PYCons() As PYBook, NowChr As String, Res As String, Read As String, temp() As String
    Dim RepCons() As PYBook
    Dim Res2 As String
    
    ReDim PYCons(0), RepCons(0)
    Open Book For Input As #1
        Do While Not EOF(1)
            Line Input #1, Read
            temp = Split(Read, "|")
            PYCons(UBound(PYCons)).Ascii = Val(temp(1)): PYCons(UBound(PYCons)).PY = temp(0)
            ReDim Preserve PYCons(UBound(PYCons) + 1)
        Loop
    Close #1
    ReDim Preserve PYCons(UBound(PYCons) - 1)
    
    If ReplaceBook <> "" Then
        Open ReplaceBook For Input As #1
            Do While Not EOF(1)
                Line Input #1, Read
                temp = Split(Read, "|")
                RepCons(UBound(RepCons)).Ascii = Val(temp(1)): RepCons(UBound(RepCons)).PY = temp(0)
                ReDim Preserve RepCons(UBound(RepCons) + 1)
            Loop
        Close #1
        ReDim Preserve RepCons(UBound(RepCons) - 1)
    End If
    
    For i = 1 To Len(CH)
        NowChr = Mid(CH, i, 1)
        Res2 = ""
        If Asc(NowChr) > PYCons(UBound(PYCons)).Ascii Or Asc(NowChr) < PYCons(0).Ascii Then
            Res = Res & NowChr
        Else
            For s = 0 To UBound(PYCons) - 1
                If Asc(NowChr) >= PYCons(s).Ascii And Asc(NowChr) < PYCons(s + 1).Ascii Then
                    If UBound(RepCons) > 0 Then
                        For t = 0 To UBound(RepCons)
                            If RepCons(t).Ascii = PYCons(s).Ascii Then Res2 = RepCons(t).PY & IIf(WithSpace = True, " ", ""): Exit For
                        Next
                    Else
                        If Res2 = "" Then Res2 = PYCons(s).PY & IIf(WithSpace = True, " ", "")
                    End If
                    Exit For
                End If
            Next
            If Res2 = "" Then
                Res2 = NowChr & IIf(WithSpace = True, " ", "")
            End If
            Res = Res & Res2
        End If
    Next
    
    AscToPY = Res
End Function
