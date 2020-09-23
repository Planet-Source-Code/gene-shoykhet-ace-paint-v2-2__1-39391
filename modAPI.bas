Attribute VB_Name = "modAPI"
' *****************************************************************************
' Project:          PaintPro
' Module:           modAPI
' Original Author:  Gene Shoykhet
' Date:             9/30/02 10:53:13 AM
' *****************************************************************************


Option Explicit

Public Declare Function ExtFloodFill Lib "gdi32" _
    (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
    ByVal crColor As Long, ByVal wFillType As Long) As Long


Public Const SRCCOPY = &HCC0020
Public Const FLOODFILLSURFACE = 1
Public Const DSTINVERT = &H550009
Public Const PATPAINT = &HFB0A09

