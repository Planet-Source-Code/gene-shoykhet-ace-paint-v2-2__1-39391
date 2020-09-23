Attribute VB_Name = "modPaint"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3D9318C601FA"
' *****************************************************************************
' Project:          PaintPro
' Version:          2.2
' Module:           modPaint
' Original Author:  Gene Shoykhet
' Modified By:
' Date:             9/11/02 11:08:33 AM
' *****************************************************************************

Option Explicit

Public gDragPen&

Public Const strTitle = "ACE Paint"
Public Const strVersion = "2.2"
Public Const strNewFile = "Untitled"

Public lngLColor&
Public lngRColor&
Public lngOutColor&

Public intFillStyle As Integer
Public CurentWidth As Single

Public blnModified As Boolean
Public strFilename As String

Public Type UDT_Tool   'UDT for tool used
    Line As Boolean
    FreeLine As Boolean
    Circle As Boolean
    Point As Boolean
    Square As Boolean
    Eraser As Boolean
    Fan As Boolean
    Elipse As Boolean
    ColorPicker As Boolean
    Caligraphy As Boolean
    Filler As Boolean
End Type

Public Type UDT_Square     'UDT for square
    mintSquareX1 As Single
    mintSquareX2 As Single
    mintSquareY1 As Single
    mintSquareY2 As Single
    blnFill As Boolean
End Type

Public Type UDT_Line       'UDT for Line
    mintLineX1 As Single
    mintLineX2 As Single
    mintLineY1 As Single
    mintLineY2 As Single
End Type

Public Type UDT_Circle     'UDT for Circle
    mintCircleX1 As Single
    mintCircleY1 As Single
    mintCircleX2 As Single
    mintCircleY2 As Single
    mdblCircleR As Single
    mblnFill As Boolean
End Type

Public Type UDT_Elipse      'UDT for Elipse
    mintElipseX1 As Single
    mintElipseY1 As Single
    mintElipseX2 As Single
    mintElipseY2 As Single
    mdblElipseR As Single
    msglAspect As Single
    mblnFill As Boolean
End Type

Public Type UDT_Point       'UDT for a two coordinate point
    msglX As Single
    msglY As Single
End Type

Public Function SetAllFalse(ByRef rudtTool As UDT_Tool) As Boolean
    'reset the UDT tool
    With rudtTool
        .FreeLine = False
        .Line = False
        .Point = False
        .Square = False
        .Eraser = False
        .Fan = False
        .Circle = False
        .Elipse = False
        .Caligraphy = False
        .ColorPicker = False
    End With
    SetAllFalse = True
End Function

Public Function distancePoints(x1 As Single, x2 As Single, y1 As Single, y2 As Single) As Double
    On Error GoTo Err
    
    Dim xses As Double
    Dim yses As Double
    
    'find the distance between two points
    xses = (x2 - x1) ^ 2
    yses = (y2 - y1) ^ 2
    'return the distance
    distancePoints = Sqr(xses + yses)
    Exit Function
Err:
    'do nothing
End Function
    
Public Function RandomizeBackground(ByRef frmName As Form)
    Dim x As Integer
    Dim y As Integer
    Dim color As Integer
    Dim Index As Integer
    
    frmName.MousePointer = vbHourglass
    
    DoEvents
    For Index = 1 To 10000
        With frmName
            x = Rnd * .picMain.ScaleWidth
            y = Rnd * .picMain.ScaleHeight
            color = Rnd * 15
            .picMain.PSet (x, y), QBColor(color)
        End With
    Next
    frmName.MousePointer = vbDefault
End Function

Public Function aspectCalculate(x As Single, y As Single, originX As Single, originY As Single) As Single
    On Error GoTo Err
    Dim distX As Single
    Dim distY As Single
    
    distX = Abs(x - originX)
    distY = Abs(y - originY)
    aspectCalculate = distY / distX
    Exit Function
Err:
    aspectCalculate = 100
    'do nothing
End Function

Public Function calculateCaligraphy(x As Single, y As Single, frmName As Form) As UDT_Point
    Dim mudtPoint As UDT_Point
    
    'change the width of the caligraphy pen depending on the current draw width
    With mudtPoint
        If CurentWidth = 1 Or CurentWidth = 3 Then
            frmName.picMain.DrawWidth = 3
            .msglX = x - 100
            .msglY = y + 100
        ElseIf CurentWidth = 5 Then
            .msglX = x - 150
            .msglY = y + 150
        ElseIf CurentWidth = 9 Then
            .msglX = x - 200
            .msglY = y + 200
        ElseIf CurentWidth = 20 Then
            .msglX = x - 300
            .msglY = y + 300
        End If
    End With
    calculateCaligraphy = mudtPoint
End Function
