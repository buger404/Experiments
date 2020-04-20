VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Fonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
Private Font(2, 1 To 100) As Long, Brush As Long
Private FontFam As Long, StrFormat(2) As Long
Sub SetFont()
    Set ProFont = Me
End Sub
Sub RegFont(ByVal file As String)
    AddFontResource file
End Sub
Sub Create(ByVal FontName As String)
    GdipCreateFontFamilyFromName StrPtr(FontName), 0, FontFam
    If FontFam = 0 Then GdipCreateFontFamilyFromName StrPtr("Î¢ÈíÑÅºÚ"), 0, FontFam
    If FontFam = 0 Then GdipCreateFontFamilyFromName StrPtr("ËÎÌå"), 0, FontFam
    If FontFam = 0 Then MsgBox "Fuck you ..."
    
    For i = 0 To 2
        GdipCreateStringFormat 0, 0, StrFormat(i)
        GdipSetStringFormatAlign StrFormat(i), i
    Next
    GdipCreateSolidFill 0, Brush
End Sub
Sub Dispose()
    RemoveFontResource file
    GdipDeleteBrush Brush
    GdipDeleteFontFamily FontFam
    For i = 0 To 2
        GdipDeleteStringFormat StrFormat(i)
    Next
    For s = 0 To 2
        For i = 1 To 100
            If Font(s, i) <> 0 Then
                GdipDeleteFont Font(s, i)
            End If
        Next
   Next
End Sub
Sub DrawText(Graphics As Long, ByVal x As Long, ByVal y As Long, ByVal W As Long, ByVal H As Long, ByVal Text As String, ForeColor As Long, _
                        Mode As StrAlignment, ByVal FontSize As Long, Optional FontMode As FontModes = 0)
    
    GdipSetSolidFillColor Brush, ForeColor
    If Font(FontMode, FontSize) = 0 Then
        Select Case FontMode
            Case 0
            GdipCreateFont FontFam, FontSize, FontStyle.FontStyleRegular, UnitPixel, Font(FontMode, FontSize)
            Case 1
            GdipCreateFont FontFam, FontSize, FontStyle.FontStyleBold, UnitPixel, Font(FontMode, FontSize)
            Case 2
            GdipCreateFont FontFam, FontSize, FontStyle.FontStyleItalic, UnitPixel, Font(FontMode, FontSize)
        End Select
    End If
    
    GdipDrawString Graphics, StrPtr(Text), -1, Font(FontMode, FontSize), NewRectF(x, y, W, H), StrFormat(Mode), Brush

    Exit Sub

End Sub

Private Sub Class_Initialize()
    AddMember Me
End Sub
