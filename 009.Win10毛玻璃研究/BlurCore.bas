Attribute VB_Name = "BlurCore"
Public Declare Function SetWindowCompositionAttribute Lib "user32.dll" (ByVal Hwnd As Long, ByRef Data As WindowsCompostionAttributeData) As Long

Public Enum WindowCompositionAttribute
    WCA_ACCENT_POLICY = 19
End Enum

Public Type WindowsCompostionAttributeData
    Attribute As WindowCompositionAttribute
    Data As Long
    SizeOfData As Long
End Type

Public Enum AccentState
    ACCENT_DISABLED = 0
    ACCENT_ENABLE_GRADIENT = 1
    ACCENT_ENABLE_TRANSPARENTGRADIENT = 2
    ACCENT_ENABLE_BLURBEHIND = 3
    ACCENT_ENABLE_ACRYLICBLURBEHIND = 4
    ACCENT_INVALID_STATE = 5
End Enum

Public Type AccentPolicy
    State As AccentState
    Flags As Long
    GradientColor As Long
    Id As Long
End Type

Public Function argb(ByVal A As Byte, ByVal r As Byte, ByVal g As Byte, ByVal B As Byte) As Long
    Dim Color As Long
    
    CopyMemory ByVal VarPtr(Color) + 3, A, 1
    CopyMemory ByVal VarPtr(Color) + 2, B, 1
    CopyMemory ByVal VarPtr(Color) + 1, g, 1
    CopyMemory ByVal VarPtr(Color), r, 1
    argb = Color
End Function


Public Sub Win10Blur(Hwnd As Long, Color As Long)
    Dim Accent As AccentPolicy, Data As WindowsCompostionAttributeData
    
    With Accent
        .State = AccentState.ACCENT_ENABLE_ACRYLICBLURBEHIND
        .GradientColor = Color
    End With
    
    With Data
        .Attribute = WindowCompositionAttribute.WCA_ACCENT_POLICY
        .SizeOfData = 16
        .Data = VarPtr(Accent)
    End With
    
    SetWindowCompositionAttribute ByVal Hwnd, Data
    
End Sub

