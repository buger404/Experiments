Attribute VB_Name = "OsuModel"
Public Type HitObject
    X As Long
    Y As Long
    time As Single
    sound As GMusic
End Type
Public Type BPM
    time As Single
    value As Single
End Type
Public CurrentObjects() As HitObject
