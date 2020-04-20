VERSION 5.00
Begin VB.UserControl ShapeEx 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4035
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   269
End
Attribute VB_Name = "ShapeEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event Click()
Public Enum ShapeEx_Shape
    ShapeRectangle = 0
    ShapeEllipse = 1
    ShapeArc = 2
    ShapeArc2 = 3
End Enum
Dim brush As Long, brush2 As Long, pen As Long, nBordercolor As Long, nBackcolor As Long, nFillcolor As Long, nAngle As Long, nShape As Long
Dim graphics As Long
Public Property Get Angle() As Long
    Angle = nAngle
End Property
Public Property Let Angle(ByVal mAngle As Long)
    nAngle = mAngle
    Call Refresh
End Property
Public Property Get Shape() As ShapeEx_Shape
    Shape = nShape
End Property
Public Property Let Shape(ByVal mShape As ShapeEx_Shape)
    nShape = mShape
    Call Refresh
End Property
Public Property Get BorderColor() As OLE_COLOR
    Dim temp(3) As Byte
    CopyMemory temp(0), nBordercolor, 4
    BorderColor = RGB(temp(2), temp(1), temp(0))
End Property
Public Property Let BorderColor(Color As OLE_COLOR)
    Dim temp(3) As Byte
    CopyMemory temp(0), Color, 4
    nBordercolor = argb(255, temp(0), temp(1), temp(2))
    GdipSetPenColor pen, nBordercolor
    Call Refresh
End Property
Public Property Get FillColor() As OLE_COLOR
    Dim temp(3) As Byte
    CopyMemory temp(0), nFillcolor, 4
    FillColor = RGB(temp(2), temp(1), temp(0))
End Property
Public Property Let FillColor(Color As OLE_COLOR)
    Dim temp(3) As Byte
    CopyMemory temp(0), Color, 4
    nFillcolor = argb(255, temp(0), temp(1), temp(2))
    GdipSetSolidFillColor brush, nFillcolor
    Call Refresh
End Property
Public Property Get BackColor() As OLE_COLOR
    Dim temp(3) As Byte
    CopyMemory temp(0), nBackcolor, 4
    BackColor = RGB(temp(2), temp(1), temp(0))
End Property
Public Property Let BackColor(Color As OLE_COLOR)
    Dim temp(3) As Byte
    CopyMemory temp(0), Color, 4
    nBackcolor = argb(255, temp(0), temp(1), temp(2))
    GdipSetSolidFillColor brush2, nBackcolor
    Call Refresh
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    If Gdip_Inited = False Then InitGDIPlus: Gdip_Inited = True
    GdipCreatePen1 0, 1, UnitPixel, pen
    GdipCreateSolidFill 0, brush
    GdipCreateSolidFill 0, brush2
    GdipCreateFromHDC UserControl.HDC, graphics
    GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
End Sub

Private Sub UserControl_InitProperties()
    Call UserControl_ReadProperties
End Sub

Private Sub UserControl_Paint()
    GdipDeleteGraphics graphics
    GdipCreateFromHDC UserControl.HDC, graphics
    GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
    Call Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    nBordercolor = PropBag.ReadProperty("BorderColor", argb(255, 255, 255, 255))
    nBackcolor = PropBag.ReadProperty("BackColor", argb(255, 255, 255, 255))
    nFillcolor = PropBag.ReadProperty("FillColor", argb(255, 232, 232, 232))
    nAngle = PropBag.ReadProperty("Angle", 0)
    nShape = PropBag.ReadProperty("Shape", ShapeEx_Shape.ShapeEllipse)
    If pen = 0 Then pen = PropBag.ReadProperty("pen")
    If brush = 0 Then brush = PropBag.ReadProperty("brush")
    If brush2 = 0 Then brush2 = PropBag.ReadProperty("brush2")
    GdipSetPenColor pen, nBordercolor
    GdipSetSolidFillColor brush, nFillcolor
    GdipSetSolidFillColor brush2, nBackcolor
    GdipDeleteGraphics graphics
    GdipCreateFromHDC UserControl.HDC, graphics
    GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
    Call Refresh
End Sub

Private Sub UserControl_Resize()
    GdipDeleteGraphics graphics
    GdipCreateFromHDC UserControl.HDC, graphics
    GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
    Call Refresh
End Sub

Private Sub UserControl_Terminate()
    If Gdip_Inited = False Then Exit Sub
    GdipDeletePen pen
    GdipDeleteBrush brush
    GdipDeleteBrush brush2
    GdipDeleteGraphics graphics
End Sub
Sub Refresh()
    GdipFillRectangle graphics, brush2, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    Dim path As Long
    Select Case nShape
        Case ShapeEx_Shape.ShapeRectangle
            GdipFillRectangle graphics, brush, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
            GdipDrawRectangle graphics, pen, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
        Case ShapeEx_Shape.ShapeEllipse
            GdipFillEllipse graphics, brush, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
            GdipDrawEllipse graphics, pen, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
        Case ShapeEx_Shape.ShapeArc
            GdipCreatePath FillModeWinding, path
            GdipAddPathArc path, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, 0, nAngle
            GdipDrawPath graphcis, pen, path
        Case ShapeEx_Shape.ShapeArc2
            GdipCreatePath FillModeWinding, path
            GdipAddPathLine path, UserControl.ScaleWidth / 2, UserControl.ScaleHeight / 2, UserControl.ScaleWidth / 2, UserControl.ScaleHeight - 1
            GdipAddPathArc path, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, 0, nAngle
            GdipClosePathFigure path
            GdipFillPath graphcis, brush, path
            GdipDrawPath graphcis, pen, path
    End Select
    
    If path <> 0 Then GdipDeletePath path
End Sub
Private Function argb(ByVal a As Long, ByVal R As Long, ByVal G As Long, ByVal b As Long) As Long
    Dim Color As Long
    CopyMemory ByVal VarPtr(Color) + 3, a, 1
    CopyMemory ByVal VarPtr(Color) + 2, R, 1
    CopyMemory ByVal VarPtr(Color) + 1, G, 1
    CopyMemory ByVal VarPtr(Color), b, 1
    argb = Color
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BorderColor", nBordercolor
    PropBag.WriteProperty "BackColor", nBackcolor
    PropBag.WriteProperty "FillColor", nFillcolor
    PropBag.WriteProperty "Angle", nAngle
    PropBag.WriteProperty "Shape", nShape
    PropBag.WriteProperty "pen", pen
    PropBag.WriteProperty "brush", brush
    PropBag.WriteProperty "brush2", brush2
End Sub
