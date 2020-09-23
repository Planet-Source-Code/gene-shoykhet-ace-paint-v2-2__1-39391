Attribute VB_Name = "modDraw"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3D9318C70256"
' *****************************************************************************
' Project:          PaintPro
' Version:          2.2
' Module:           modPaint
' Original Author:  Gene Shoykhet
' Modified By:
' Date:             9/11/02 11:08:33 AM
' *****************************************************************************

Option Explicit

Private msglOriginX As Single
Private msglOriginY As Single

Private CurentX As Single
Private CurentY As Single
Private mudtCalPoint As UDT_Point

Private mudtLine As UDT_Line
Private mudtCircle As UDT_Circle
Private mudtElipse As UDT_Elipse
Private mudtSquare As UDT_Square

Public Sub InitiateLine(mudtP As UDT_Point, frmName As Form)
    With mudtLine
        .mintLineX1 = mudtP.msglX
        .mintLineY1 = mudtP.msglY
        .mintLineX2 = mudtP.msglX
        .mintLineY2 = mudtP.msglY
    End With
End Sub

Public Sub DrawLine(mudtP As UDT_Point, frmName As Form)
    With mudtLine
        gDragPen& = frmName.picMain.BackColor Xor QBColor(0)
        frmName.picMain.DrawMode = 7    'to enable live drawing
        frmName.picMain.Line (.mintLineX1, .mintLineY1)-(.mintLineX2, .mintLineY2), gDragPen&
                    
        .mintLineX2 = mudtP.msglX
        .mintLineY2 = mudtP.msglY
        frmName.picMain.Line (.mintLineX1, .mintLineY1)-(.mintLineX2, .mintLineY2), gDragPen&
    End With
End Sub

Public Sub FinalizeLine(mudtP As UDT_Point, frmName As Form)
    With mudtLine
        .mintLineX2 = mudtP.msglX
        .mintLineY2 = mudtP.msglY
         blnModified = True
         frmName.picMain.Line (.mintLineX1, .mintLineY1)-(.mintLineX2, .mintLineY2), lngOutColor&
    End With
End Sub

Public Sub InitiateCircle(mudtP As UDT_Point, Fill As Boolean, frmName As Form)
    With mudtCircle
        .mintCircleX1 = mudtP.msglX
        .mintCircleY1 = mudtP.msglY
        .mintCircleX2 = mudtP.msglX
        .mintCircleY2 = mudtP.msglY
        .mblnFill = Fill
                
        gDragPen& = frmName.picMain.BackColor Xor QBColor(0)
        frmName.picMain.DrawMode = 7    'to enable live drawing
                
        .mdblCircleR = distancePoints(.mintCircleX1, .mintCircleX2, _
            .mintCircleY1, .mintCircleY2)
        frmName.picMain.Circle (.mintCircleX1, .mintCircleY1), .mdblCircleR, gDragPen&
    End With
End Sub

Public Sub DrawCircle(mudtP As UDT_Point, frmName As Form)
    With mudtCircle
        gDragPen& = frmName.picMain.BackColor Xor QBColor(0)
        frmName.picMain.DrawMode = 7    'to enable live drawing
        frmName.picMain.FillStyle = vbFSTransparent
        
        .mdblCircleR = distancePoints(.mintCircleX1, .mintCircleX2, _
            .mintCircleY1, .mintCircleY2)
        frmName.picMain.Circle (.mintCircleX1, .mintCircleY1), .mdblCircleR, gDragPen&
                    
        .mintCircleX2 = mudtP.msglX
        .mintCircleY2 = mudtP.msglY
        .mdblCircleR = distancePoints(.mintCircleX1, .mintCircleX2, _
            .mintCircleY1, .mintCircleY2)
        frmName.picMain.Circle (.mintCircleX1, .mintCircleY1), .mdblCircleR, gDragPen&
    End With
End Sub

Public Sub FinalizeCircle(mudtP As UDT_Point, frmName As Form)
    With mudtCircle
        .mintCircleX2 = mudtP.msglX
        .mintCircleY2 = mudtP.msglY
        .mdblCircleR = distancePoints(.mintCircleX1, .mintCircleX2, .mintCircleY1, .mintCircleY2)
        
        If .mblnFill Then
            frmName.picMain.FillColor = lngRColor&
            frmName.picMain.FillStyle = intFillStyle
        End If
        
        blnModified = True
        frmName.picMain.Circle (.mintCircleX1, .mintCircleY1), .mdblCircleR, lngOutColor&
    End With
End Sub

Public Sub InitiateElipse(mudtP As UDT_Point, Fill As Boolean, frmName As Form)
    With mudtElipse
        msglOriginX = mudtP.msglX
        msglOriginY = mudtP.msglY
        .mintElipseX1 = mudtP.msglX
        .mintElipseX2 = mudtP.msglX
        .mintElipseY1 = mudtP.msglY
        .mintElipseY2 = mudtP.msglY
        .mblnFill = Fill
        
        gDragPen& = frmName.picMain.BackColor Xor QBColor(0)
        frmName.picMain.DrawMode = 7    'to enable live drawing
                
        .mdblElipseR = distancePoints(.mintElipseX1, .mintElipseX2, _
            .mintElipseY1, .mintElipseY2)
        frmName.picMain.Circle (.mintElipseX1, .mintElipseY1), .mdblElipseR, gDragPen&
    End With
End Sub

Public Sub DrawElipse(mudtP As UDT_Point, frmName As Form)
    With mudtElipse
        gDragPen& = frmName.picMain.BackColor Xor QBColor(0)
        frmName.picMain.DrawMode = 7    'to enable live drawing
        frmName.picMain.FillStyle = vbFSTransparent
                    
        .mdblElipseR = distancePoints(.mintElipseX1, .mintElipseX2, _
            .mintElipseY1, .mintElipseY2)
        .msglAspect = aspectCalculate(.mintElipseX2, .mintElipseY2, msglOriginX, msglOriginY)
        frmName.picMain.Circle (.mintElipseX1, .mintElipseY1), .mdblElipseR, gDragPen&, , , .msglAspect
                    
        .mintElipseX2 = mudtP.msglX
        .mintElipseY2 = mudtP.msglY
        .mdblElipseR = distancePoints(.mintElipseX1, .mintElipseX2, _
            .mintElipseY1, .mintElipseY2)
        .msglAspect = aspectCalculate(.mintElipseX2, .mintElipseY2, msglOriginX, msglOriginY)
        frmName.picMain.Circle (.mintElipseX1, .mintElipseY1), .mdblElipseR, gDragPen&, , , .msglAspect
    End With
End Sub

Public Sub FinalizeElipse(mudtP As UDT_Point, frmName As Form)
    With mudtElipse
        .mintElipseX2 = mudtP.msglX
        .mintElipseY2 = mudtP.msglY
        frmName.picMain.FillColor = lngRColor&
        
        If .mblnFill Then
            frmName.picMain.FillColor = lngRColor&
            frmName.picMain.FillStyle = intFillStyle
        End If
        
        .mdblElipseR = distancePoints(.mintElipseX1, .mintElipseX2, _
            .mintElipseY1, .mintElipseY2)
        .msglAspect = aspectCalculate(.mintElipseX2, .mintElipseY2, msglOriginX, msglOriginY)
        blnModified = True
        frmName.picMain.Circle (.mintElipseX1, .mintElipseY1), .mdblElipseR, lngOutColor&, , , .msglAspect
    End With
End Sub

Public Sub InitiateSquare(mudtP As UDT_Point, filled As Boolean, frmName As Form)
    frmName.picMain.FillColor = lngRColor&
    With mudtSquare
        .blnFill = filled
        .mintSquareX1 = mudtP.msglX
        .mintSquareY1 = mudtP.msglY
        .mintSquareX2 = mudtP.msglX
        .mintSquareY2 = mudtP.msglY
    End With
End Sub

Public Sub DrawSquare(mudtP As UDT_Point, frmName As Form)
    With mudtSquare
        gDragPen& = frmName.picMain.BackColor Xor QBColor(0)
        frmName.picMain.DrawMode = 7    'to enable live drawing
        frmName.picMain.FillStyle = vbFSTransparent
        
        frmName.picMain.Line (.mintSquareX1, .mintSquareY1)-(.mintSquareX2, .mintSquareY2), gDragPen&, B
                        
        .mintSquareX2 = mudtP.msglX
        .mintSquareY2 = mudtP.msglY
        frmName.picMain.Line (.mintSquareX1, .mintSquareY1)-(.mintSquareX2, .mintSquareY2), gDragPen&, B
    End With
End Sub

Public Sub FinalizeSquare(mudtP As UDT_Point, frmName As Form)
    With mudtSquare
        .mintSquareX2 = mudtP.msglX
        .mintSquareY2 = mudtP.msglY
        If Not .blnFill Then
            frmName.picMain.FillStyle = vbFSTransparent
            blnModified = True
            frmName.picMain.Line (.mintSquareX1, .mintSquareY1)-(.mintSquareX2, .mintSquareY2) _
                , lngOutColor&, B
        Else
            frmName.picMain.FillStyle = intFillStyle
            blnModified = True
            frmName.picMain.Line (.mintSquareX1, .mintSquareY1)-(.mintSquareX2, .mintSquareY2) _
                , lngOutColor&, B
        End If
    End With
End Sub

Public Sub InitiateFreeLine(mudtP As UDT_Point, frmName As Form)
    blnModified = True
    CurentX = mudtP.msglX
    CurentY = mudtP.msglY
End Sub

Public Sub DrawFreeLine(mudtP As UDT_Point, frmName As Form)
    frmName.picMain.Line (CurentX, CurentY)-(mudtP.msglX, mudtP.msglY), lngOutColor&
    CurentX = mudtP.msglX
    CurentY = mudtP.msglY
End Sub

Public Sub InitiateFan(mudtP As UDT_Point, frmName As Form)
    blnModified = True
    CurentX = mudtP.msglX
    CurentY = mudtP.msglY
End Sub

Public Sub DrawFan(mudtP As UDT_Point, frmName As Form)
    frmName.picMain.Line (CurentX, CurentY)-(mudtP.msglX, mudtP.msglY), lngOutColor&
End Sub

Public Sub GetColorPicker(mudtP As UDT_Point, Button As Integer, frmName As Form)
    If Button = vbLeftButton Then
        lngLColor& = frmName.picMain.Point(mudtP.msglX, mudtP.msglY)    'get the color of the location of the mouse pointer
        frmName.picColorL.BackColor = lngLColor&    'and set the foreground color
    Else
        lngRColor& = frmName.picMain.Point(mudtP.msglX, mudtP.msglY)    'set the background color
        frmName.picColorR.BackColor = lngRColor&
    End If
End Sub

Public Sub InitiateCaligraphy(mudtP As UDT_Point, frmName As Form)
    mudtCalPoint = calculateCaligraphy(mudtP.msglX, mudtP.msglY, frmName)
    With mudtCalPoint
        frmName.picMain.Line (mudtP.msglX, mudtP.msglY)-(.msglX, .msglY), lngOutColor&
    End With
End Sub

Public Sub DrawCaligraphy(mudtP As UDT_Point, frmName As Form)
    mudtCalPoint = calculateCaligraphy(mudtP.msglX, mudtP.msglY, frmName)
    With mudtCalPoint
        frmName.picMain.Line (mudtP.msglX, mudtP.msglY)-(.msglX, .msglY), lngOutColor&
    End With
End Sub

Public Sub DoFiller(mudtP As UDT_Point, picHDC As Long, frmName As Form)
    Dim fillerColor&

    fillerColor& = frmPaint.picMain.Point(mudtP.msglX, mudtP.msglY)    'get the color
    
    ExtFloodFill picHDC, mudtP.msglX, mudtP.msglY, fillerColor&, 1      'execute the fill using API
End Sub
