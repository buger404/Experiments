VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OsuCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private BPMs() As BPM
Private Objects() As HitObject
Public SliderMultiplier As Long
Public Audio As String
Public Title As String
Public Code As String
Public Sub GetObjects()
    CurrentObjects = Objects
End Sub
Public Sub Load(OsuCode As String)
    Code = OsuCode
    Audio = Property("General", "AudioFilename")
    Audio = Right(Audio, Len(Audio) - 1)
    Title = Property("Metadata", "Title")
    SliderMultiplier = Val(Property("Difficulty", "SliderMultiplier"))
    Debug.Print Now, "Audio", Audio
    Debug.Print Now, "Title", Title
    Debug.Print Now, "SliderMultiplier", SliderMultiplier
    Dim Line() As String
    Line = Split(Property("TimingPoints", ""), vbCrLf)
    ReDim BPMs(0)
    Dim LastO As Long
    For I = 0 To UBound(Line)
        If Line(I) <> "" Then
            ReDim Preserve BPMs(UBound(BPMs) + 1)
            With BPMs(UBound(BPMs))
                .time = Val(Split(Line(I), ",")(0))
                .value = Val(Split(Line(I), ",")(1))
                'Debug.Print Now, "BPM", .Value
                If .value < 0 Then
                    .value = LastO * Abs(100 / .value)
                Else
                    LastO = .value
                End If
            End With
        End If
    Next
    Dim C() As String, ObjType As Integer
    Dim snd(2) As String
    snd(0) = "error": snd(1) = "ding": snd(2) = "warning"
    'snd(0) = "hit.wav": snd(1) = "hit.wav": snd(2) = "hit.wav"
    For I = 1 To UBound(Objects)
        If Not Objects(I).sound Is Nothing Then Objects(I).sound.Dispose: Set Objects(I).sound = Nothing
    Next
    ReDim Objects(0)
    Line = Split(Property("HitObjects", ""), vbCrLf)
    For I = 0 To UBound(Line)
        If Line(I) <> "" Then
            'X,Y,Time,Type,Sound,<OTHER>0:0:0:0: //Ignore the latter
            C = Split(Line(I), ",")
            '1=圆圈 2=滑条 8=转盘
            ObjType = (Val(C(3)) Mod 16)
            If ObjType = 1 Then
                'Debug.Print Now, "发现圆圈！"
                ReDim Preserve Objects(UBound(Objects) + 1)
                With Objects(UBound(Objects))
                    .X = Val(C(0))
                    .y = Val(C(1))
                    .time = Val(C(2))
                    .Kind = snd(I Mod 3)
                    Set .sound = New GMusic
                    .sound.Create App.path & "\" & snd(I Mod 3) & ".mp3"
                End With
            End If
            If ObjType = 2 Then
                'Debug.Print Now, "发现滑条！"
                Dim d() As String, Span As Long, Round As Long, BP As Long
                Dim OX As Long, OY As Long, OT As Long
                Round = Val(C(6))
                For S = 1 To UBound(BPMs)
                    If Val(C(2)) > BPMs(S).time Then BP = BPMs(S).value
                Next
                Span = Abs(Val(C(7)) * BP / SliderMultiplier) / 100
                d = Split(C(5), "|")
                If UBound(d) > 1 Then Span = Span / (UBound(d) - 1)
                Dim od() As String
                od = d
                For S = 2 To Round
                    If S Mod 2 = 0 Then
                        For J = UBound(od) To 1 Step -1
                            ReDim Preserve d(UBound(d) + 1)
                            d(UBound(d)) = od(J)
                        Next
                    Else
                        For J = 1 To UBound(od)
                            ReDim Preserve d(UBound(d) + 1)
                            d(UBound(d)) = od(J)
                        Next
                    End If
                Next
                OX = Val(C(0)): OY = Val(C(1)): OT = Val(C(2))
                For J = 1 To UBound(d)
                    C = Split(d(J), ":")
                    For S = 1 To Span Step 10
                        ReDim Preserve Objects(UBound(Objects) + 1)
                        With Objects(UBound(Objects))
                            .X = OX + (Val(C(0)) - OX) * (S / Span)
                            .y = OY + (Val(C(1)) - OY) * (S / Span)
                            .Kind = snd(I Mod 3)
                            .time = OT + Span * (J - 1) + S
                            If S Mod 100 = 1 Then
                                Set .sound = New GMusic
                                .sound.Create App.path & "\" & snd(I Mod 3) & ".mp3"
                                .sound.rate = 0.8 + Rnd * 0.4
                            End If
                        End With
                    Next
                    'ReDim Preserve Objects(UBound(Objects) + 1)
                    'With Objects(UBound(Objects))
                    '    .X = Val(c(0))
                    '    .y = Val(c(1))
                    '    .time = OT + Span * J
                    'End With
                    OX = Val(C(0)): OY = Val(C(1))
                Next
                'Debug.Print Now, Round, Span & "ms"
           End If
           If ObjType = 8 Then
                'Debug.Print Now, "发现转盘！"
           End If
        End If
    Next
End Sub
Public Function Property(SubName As String, ItemName As String) As String
    Dim temp() As String
    temp = Split(Code, "[" & SubName & "]")
    If ItemName <> "" Then temp = Split(temp(1), ItemName & ":")
    Property = Split(temp(1), IIf(ItemName = "", "[", vbCrLf))(0)
End Function

Private Sub Class_Initialize()
    ReDim Objects(0)
End Sub
