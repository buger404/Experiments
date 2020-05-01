Attribute VB_Name = "OsuModel"
Public Type HitObject
    X As Long
    y As Long
    time As Single
    kind As String
    sound As GMusic
End Type
Public Type BPM
    time As Single
    value As Single
End Type
Public CurrentObjects() As HitObject
