VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPaint 
   Caption         =   "PaintPro"
   ClientHeight    =   8160
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10380
   Icon            =   "PaintPro.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "PaintPro.frx":0442
   ScaleHeight     =   8160
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSec 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   6075
      Left            =   1020
      ScaleHeight     =   6045
      ScaleWidth      =   9225
      TabIndex        =   53
      Top             =   120
      Visible         =   0   'False
      Width           =   9255
   End
   Begin VB.Frame Frame5 
      Caption         =   "Fill Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   1020
      TabIndex        =   31
      Top             =   6240
      Width           =   2115
      Begin VB.OptionButton optFillType 
         Caption         =   "Option3"
         Height          =   195
         Index           =   6
         Left            =   1680
         TabIndex        =   44
         Top             =   1620
         Width           =   195
      End
      Begin VB.OptionButton optFillType 
         Caption         =   "Option3"
         Height          =   195
         Index           =   5
         Left            =   1680
         TabIndex        =   37
         Top             =   1380
         Width           =   195
      End
      Begin VB.OptionButton optFillType 
         Caption         =   "Option3"
         Height          =   195
         Index           =   4
         Left            =   1680
         TabIndex        =   36
         Top             =   1140
         Width           =   195
      End
      Begin VB.OptionButton optFillType 
         Caption         =   "Option3"
         Height          =   195
         Index           =   3
         Left            =   1680
         TabIndex        =   35
         Top             =   900
         Width           =   195
      End
      Begin VB.OptionButton optFillType 
         Caption         =   "Option3"
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   34
         Top             =   660
         Width           =   195
      End
      Begin VB.OptionButton optFillType 
         Caption         =   "Option3"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   33
         Top             =   420
         Width           =   195
      End
      Begin VB.OptionButton optFillType 
         Caption         =   "Option3"
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   32
         Top             =   180
         Width           =   195
      End
      Begin VB.Label Label13 
         Caption         =   "Cross Diagonal"
         Height          =   195
         Left            =   60
         TabIndex        =   45
         Top             =   1620
         Width           =   1155
      End
      Begin VB.Label Label12 
         Caption         =   "Cross"
         Height          =   195
         Left            =   60
         TabIndex        =   43
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Down Diagonal"
         Height          =   195
         Left            =   60
         TabIndex        =   42
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Up Diagonal"
         Height          =   195
         Left            =   60
         TabIndex        =   41
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label Label9 
         Caption         =   "Vertical"
         Height          =   195
         Left            =   60
         TabIndex        =   40
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Horizontal"
         Height          =   195
         Left            =   60
         TabIndex        =   39
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Solid"
         Height          =   195
         Left            =   60
         TabIndex        =   38
         Top             =   180
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Background"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   5940
      TabIndex        =   27
      Top             =   6240
      Width           =   1455
      Begin VB.CommandButton cmdBackground 
         Caption         =   "Random"
         Height          =   315
         Index           =   1
         Left            =   300
         TabIndex        =   51
         Top             =   1020
         Width           =   855
      End
      Begin VB.CommandButton cmdBackground 
         Caption         =   "Solid Fill"
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   50
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open..."
      Height          =   375
      Left            =   9060
      TabIndex        =   16
      ToolTipText     =   "Open an existing image"
      Top             =   6300
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Save &As..."
      Height          =   375
      Left            =   7740
      TabIndex        =   15
      ToolTipText     =   "Save current image as new file."
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   7740
      TabIndex        =   14
      ToolTipText     =   "Save the current image."
      Top             =   6300
      Width           =   1095
   End
   Begin VB.CommandButton cmdTool 
      Height          =   375
      Index           =   9
      Left            =   240
      Picture         =   "PaintPro.frx":0488
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Fan"
      Top             =   4080
      Width           =   495
   End
   Begin VB.Frame Frame3 
      Caption         =   "Brush Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   3480
      TabIndex        =   21
      Top             =   6240
      Width           =   2055
      Begin VB.OptionButton optBrush 
         Caption         =   "Option1"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   12
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton optBrush 
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   11
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton optBrush 
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton optBrush 
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   9
         Top             =   480
         Width           =   255
      End
      Begin VB.OptionButton optBrush 
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "GIT!!!"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Larger Yet"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Large"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Medium"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Small"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   20
      Top             =   6420
      Width           =   735
      Begin VB.PictureBox picColorL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000008&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         ScaleHeight     =   585
         ScaleWidth      =   345
         TabIndex        =   47
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox picColorR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         ScaleHeight     =   585
         ScaleWidth      =   345
         TabIndex        =   46
         Top             =   780
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog cdlColor 
      Left            =   9360
      Top             =   6900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   7740
      TabIndex        =   13
      ToolTipText     =   "Clear the current image"
      Top             =   7620
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   9060
      TabIndex        =   17
      ToolTipText     =   "Exit the application"
      Top             =   7620
      Width           =   1095
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   6075
      Left            =   1020
      MousePointer    =   2  'Cross
      ScaleHeight     =   403
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   615
      TabIndex        =   19
      Top             =   120
      Width           =   9255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6195
      Left            =   120
      TabIndex        =   18
      Top             =   60
      Width           =   735
      Begin VB.CommandButton cmdTool 
         Height          =   375
         Index           =   12
         Left            =   120
         Picture         =   "PaintPro.frx":0AA2
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Color Picker"
         Top             =   5280
         Width           =   495
      End
      Begin VB.CommandButton cmdTool 
         Height          =   375
         Index           =   11
         Left            =   120
         Picture         =   "PaintPro.frx":10BC
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Color Picker"
         Top             =   4860
         Width           =   495
      End
      Begin VB.CommandButton cmdTool 
         Height          =   375
         Index           =   10
         Left            =   120
         Picture         =   "PaintPro.frx":16D6
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Color Picker"
         Top             =   4440
         Width           =   495
      End
      Begin VB.CommandButton cmdTool 
         Height          =   375
         Index           =   5
         Left            =   120
         Picture         =   "PaintPro.frx":1CF0
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Elipse"
         Top             =   2340
         Width           =   495
      End
      Begin VB.CommandButton cmdTool 
         Height          =   375
         Index           =   6
         Left            =   120
         Picture         =   "PaintPro.frx":230A
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Filled Elipse"
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton cmdTool 
         Height          =   375
         Index           =   3
         Left            =   120
         Picture         =   "PaintPro.frx":2924
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Circle"
         Top             =   1500
         Width           =   495
      End
      Begin VB.CommandButton cmdTool 
         Height          =   375
         Index           =   15
         Left            =   120
         Picture         =   "PaintPro.frx":2F3E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Eraser"
         Top             =   5700
         Width           =   495
      End
      Begin VB.CommandButton cmdTool 
         Height          =   375
         Index           =   2
         Left            =   120
         Picture         =   "PaintPro.frx":3760
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Free Line"
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdTool 
         Height          =   375
         Index           =   4
         Left            =   120
         Picture         =   "PaintPro.frx":3D7A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Filled Circle"
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton cmdTool 
         Height          =   375
         Index           =   8
         Left            =   120
         Picture         =   "PaintPro.frx":4394
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Filled Rectangle"
         Top             =   3600
         Width           =   495
      End
      Begin VB.CommandButton cmdTool 
         Height          =   375
         Index           =   7
         Left            =   120
         Picture         =   "PaintPro.frx":49AE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Rectangle"
         Top             =   3180
         Width           =   495
      End
      Begin VB.CommandButton cmdTool 
         Height          =   375
         Index           =   0
         Left            =   120
         Picture         =   "PaintPro.frx":4FC8
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Point"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdTool 
         Height          =   375
         Index           =   1
         Left            =   120
         Picture         =   "PaintPro.frx":55E2
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Straight Line"
         Top             =   660
         Width           =   495
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuTool 
         Caption         =   "Point"
         Index           =   0
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Line"
         Index           =   1
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Free Line"
         Index           =   2
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Circle"
         Index           =   3
      End
      Begin VB.Menu mnuTool 
         Caption         =   "CircleFill"
         Index           =   4
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Elipse"
         Index           =   5
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Elipse Fill"
         Index           =   6
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Square"
         Index           =   7
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Square Fill"
         Index           =   8
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Fan"
         Index           =   9
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Color Picker"
         Index           =   10
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Caligraphy"
         Index           =   11
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Eraser"
         Index           =   15
      End
   End
   Begin VB.Menu mnuBack 
      Caption         =   "Background"
      Begin VB.Menu mnuBG 
         Caption         =   "Solid"
         Index           =   0
      End
      Begin VB.Menu mnuBG 
         Caption         =   "Random"
         Index           =   1
      End
   End
   Begin VB.Menu mnuClr 
      Caption         =   "Color"
      Begin VB.Menu mnuForeColor 
         Caption         =   "Fore Color"
      End
      Begin VB.Menu mnuBackColor 
         Caption         =   "Back Color"
      End
   End
   Begin VB.Menu mnuFillType 
      Caption         =   "Fill Type"
      Begin VB.Menu mnuFill 
         Caption         =   "Solid"
         Index           =   0
      End
      Begin VB.Menu mnuFill 
         Caption         =   "Horizontal"
         Index           =   1
      End
      Begin VB.Menu mnuFill 
         Caption         =   "Vertical"
         Index           =   2
      End
      Begin VB.Menu mnuFill 
         Caption         =   "Up Diagonal"
         Index           =   3
      End
      Begin VB.Menu mnuFill 
         Caption         =   "Down Diagonal"
         Index           =   4
      End
      Begin VB.Menu mnuFill 
         Caption         =   "Cross"
         Index           =   5
      End
      Begin VB.Menu mnuFill 
         Caption         =   "Cross Diagonal"
         Index           =   6
      End
   End
   Begin VB.Menu mnuBrushSize 
      Caption         =   "Brush Size"
      Begin VB.Menu mnuBrush 
         Caption         =   "Small"
         Index           =   0
      End
      Begin VB.Menu mnuBrush 
         Caption         =   "Medium"
         Index           =   1
      End
      Begin VB.Menu mnuBrush 
         Caption         =   "Large"
         Index           =   2
      End
      Begin VB.Menu mnuBrush 
         Caption         =   "Larger Yet"
         Index           =   3
      End
      Begin VB.Menu mnuBrush 
         Caption         =   "GIT!!!"
         Index           =   4
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3D9318C0010C"
' *****************************************************************************
' Project:          PaintPro
' Version:          2.2
' Module:           frmPaint
' Original Author:  Gene Shoykhet
' Modified By:
' Date:             9/11/02 11:08:33 AM
' *****************************************************************************

Option Explicit

Private blnSquareFill As Boolean
Private blnCircleFill As Boolean
Private blnElipseFill As Boolean

Private mudtTool As UDT_Tool
Private mudtPoint As UDT_Point

Private Sub cmdExit_Click()
   Dim intExitChoice As Integer
    'make sure user wants to exit and/or save
   If blnModified Then
      intExitChoice = MsgBox("Would you like to save before exiting?", vbYesNoCancel, "ACE Paint Exit")
        If intExitChoice = vbYes Then
                'save if user chooses Yes
            cmdSave_Click
            End
        ElseIf intExitChoice = vbNo Then
            End
        Else
                'return to the program if cancel is chosen
            Exit Sub
        End If
   End If
        'if not modified, then just exit
   End
End Sub

Private Sub cmdNew_Click()
        'if picture modified, ask for a save
    If blnModified Then
        If MsgBox("Erase without saving?", vbInformation + vbOKCancel, "ACE Paint New") = vbOK Then
                'clear the picture control
            picMain.Cls
            picMain.Refresh
            strFilename = strNewFile
            frmPaint.Caption = strTitle & " v" & strVersion & " <" & strFilename & ">"
        Else
            Exit Sub
        End If
    Else
        picMain.Cls
        picMain.Refresh
        strFilename = strNewFile
        frmPaint.Caption = strTitle & " v" & strVersion & " <" & strFilename & ">"
        Exit Sub
    End If
End Sub

Private Sub cmdOpen_Click()
    On Error GoTo Err
    
        'bring up the open dialog box
    cdlColor.Filter = "Bitmap" & "(*.bmp)|*.bmp|Jpeg Files (*.jpg)|*.jpg"
    cdlColor.FilterIndex = 1
    cdlColor.ShowOpen
    cdlColor.CancelError = False
    strFilename = cdlColor.FileName
    picMain.Picture = LoadPicture(strFilename)
    frmPaint.Caption = strTitle & " " & strVersion & " <" & cdlColor.FileTitle & ">"
Exit Sub

Err:
    'do nothing here
End Sub

Private Sub cmdSaveAs_Click()
    On Error GoTo Err
    
        'bring up the save dialog box
    cdlColor.Filter = "Bitmap" & "(*.bmp)|*.bmp|Jpeg Files (*.jpg)|*.jpg"
        'set the default to .bmp
    cdlColor.FilterIndex = 1
        'bring up the common dialog box for save
    cdlColor.ShowSave
        'save the image
    SavePicture picMain.Image, cdlColor.FileName
        'display the saved filename as the form caption
    strFilename = cdlColor.FileName
    frmPaint.Caption = strTitle & " " & strVersion & " <" & cdlColor.FileTitle & ">"
Exit Sub

Err:
    'do nothing here
End Sub

Private Sub cmdTool_Click(Index As Integer)
        'set UDT tool to clear previous tool
    SetAllFalse mudtTool
        'set the width in case it chaged with the previous tool
    picMain.DrawWidth = CurentWidth
        'select tool and attributes
    With mudtTool
        Select Case Index
            Case 0: .Point = True     'single point
            Case 1: .Line = True      'straight line
            Case 2: .FreeLine = True  'freehand line
            Case 3      'circle
                picMain.FillStyle = vbFSTransparent
                blnCircleFill = False
                .Circle = True
            Case 4      'filled Circle
                picMain.FillStyle = intFillStyle
                blnCircleFill = True
                .Circle = True
            Case 5      'elipse
                picMain.FillStyle = vbFSTransparent
                blnElipseFill = False
                .Elipse = True
            Case 6      'elipse fill
                picMain.FillStyle = intFillStyle
                blnElipseFill = True
                .Elipse = True
            Case 7      'empty rectangle
                blnSquareFill = False
                picMain.FillStyle = vbFSTransparent
                .Square = True
            Case 8      'filled rectangle
                blnSquareFill = True
                picMain.FillStyle = intFillStyle
                picMain.FillColor = lngRColor&
                .Square = True
            Case 9: .Fan = True     'fan effect
            Case 10: .ColorPicker = True    'colorpicker tool
            Case 11: .Caligraphy = True     'caligraphy tool
            Case 12
                .Filler = True         'filler tool
                picMain.FillStyle = intFillStyle
            Case 15     'eraser
                picMain.DrawWidth = CurentWidth + 5
                .Eraser = True
            Case 20: picMain.BackColor = lngRColor&    'fill background with current selected color
            Case 21: RandomizeBackground Me    'fill background with random color dots
        End Select
    End With
End Sub

Private Sub Form_Load()
    frmPaint.MousePointer = vbDefault
    strFilename = strNewFile
    
        'set the caption to static app title and version defined in modPaint.bas
    frmPaint.Caption = strTitle & " v" & strVersion & " <" & strFilename & ">"
    
        'set the default colors and point width
    picColorL.BackColor = vbBlack
    CurentWidth = 1
    picMain.DrawWidth = CurentWidth
    lngLColor& = vbBlack
    lngRColor& = vbWhite
    picMain.FillColor = vbBlack
    picMain.FillStyle = vbSolid
    
        'set other defaults
    optBrush(0).Value = True
    optFillType(0).Value = True
    blnModified = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
        'if picture modified, ask for a save
    If blnModified Then
        If MsgBox("Save Before Exit?", vbInformation + vbYesNo, "ACE Paint Exit") = vbYes Then
                'check to see if file was already saved and has a name
            If strFilename <> strNewFile Then
                SavePicture picMain.Image, strFilename
            Else
                cmdSaveAs_Click
            End If
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuBackColor_Click()
    picColorR_Click
End Sub

Private Sub mnuBG_Click(Index As Integer)
    cmdBackground_Click Index
End Sub

Private Sub mnuBrush_Click(Index As Integer)
    optBrush(Index).Value = True
    optBrush_Click (Index)
End Sub

Private Sub mnuExit_Click()
    cmdExit_Click
End Sub

Private Sub mnuFill_Click(Index As Integer)
    optFillType(Index).Value = True
    optFillType_Click Index
End Sub

Private Sub mnuForeColor_Click()
    picColorL_Click
End Sub

Private Sub mnuNew_Click()
    cmdNew_Click
End Sub

Private Sub mnuOpen_Click()
    cmdOpen_Click
End Sub

Private Sub mnuSave_Click()
    cmdSave_Click
End Sub

Private Sub mnuSaveAs_Click()
    cmdSaveAs_Click
End Sub

Private Sub mnuTool_Click(Index As Integer)
    cmdTool_Click Index
End Sub

Private Sub optFillType_Click(Index As Integer)
        'select the fill style
    Select Case Index
        Case 0: intFillStyle = vbFSSolid
        Case 1: intFillStyle = vbHorizontalLine
        Case 2: intFillStyle = vbVerticalLine
        Case 3: intFillStyle = vbUpwardDiagonal
        Case 4: intFillStyle = vbDownwardDiagonal
        Case 5: intFillStyle = vbCross
        Case 6: intFillStyle = vbDiagonalCross
    End Select
    picMain.FillStyle = intFillStyle
End Sub

Private Sub optBrush_Click(Index As Integer)
        'select brush size
    Select Case Index
        Case 0: CurentWidth = 1
        Case 1: CurentWidth = 3
        Case 2: CurentWidth = 5
        Case 3: CurentWidth = 9
        Case 4: CurentWidth = 20
    End Select
    If mudtTool.Eraser Then CurentWidth = CurentWidth + 5
    picMain.DrawWidth = CurentWidth
End Sub

Private Sub cmdBackground_Click(Index As Integer)
    Select Case Index
        Case 0: cmdTool_Click 20
        Case 1: cmdTool_Click 21
    End Select
End Sub

Private Sub picColorL_Click()
        'bring up the color dialog box
    cdlColor.ShowColor
    lngLColor& = cdlColor.color
        'set the preview box color
    picColorL.BackColor = lngLColor&
End Sub

Private Sub picColorR_Click()
        'bring up the color dialog box
    cdlColor.ShowColor
    lngRColor& = cdlColor.color
        'set the preview box color
    picColorR.BackColor = lngRColor&
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'check to see which tool is selected and act accordingly
   If Button = vbLeftButton Then
        lngOutColor& = lngLColor&
   ElseIf Button = vbRightButton Then
        lngOutColor& = lngRColor&
   End If
   With mudtPoint
        .msglX = x
        .msglY = y
   End With
   'start to draw the objects
   With mudtTool
        If .Line Then
            InitiateLine mudtPoint, Me
        ElseIf .Point Then
            blnModified = True
            picMain.PSet (mudtPoint.msglX, mudtPoint.msglY), lngOutColor&
        ElseIf .Circle Then
            InitiateCircle mudtPoint, blnCircleFill, Me
        ElseIf .Elipse Then
            InitiateElipse mudtPoint, blnElipseFill, Me
        ElseIf .Square Then
            InitiateSquare mudtPoint, blnSquareFill, Me
        ElseIf .Eraser Then
            picMain.PSet (x, y), picMain.BackColor
        ElseIf .FreeLine Then
            InitiateFreeLine mudtPoint, Me
        ElseIf .Fan Then
            InitiateFan mudtPoint, Me
        ElseIf .ColorPicker Then
            GetColorPicker mudtPoint, Button, Me
        ElseIf .Caligraphy Then
            InitiateCaligraphy mudtPoint, Me
        ElseIf .Filler Then
            picMain.FillColor = lngOutColor&
            DoFiller mudtPoint, picMain.hdc, Me
        End If
   End With
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Or Button = vbRightButton Then
       With mudtPoint
        .msglX = x
        .msglY = y
       End With
        'use live drawing features here
        With mudtTool
            If .FreeLine Then
                DrawFreeLine mudtPoint, Me
            ElseIf .Fan Then
                DrawFan mudtPoint, Me
            ElseIf .Eraser Then
                picMain.PSet (x, y), picMain.BackColor
            ElseIf .Line Then
                DrawLine mudtPoint, Me
            ElseIf .Circle Then
                DrawCircle mudtPoint, Me
            ElseIf .Elipse Then
                DrawElipse mudtPoint, Me
            ElseIf .Square Then
                DrawSquare mudtPoint, Me
            ElseIf .Caligraphy Then
                DrawCaligraphy mudtPoint, Me
            End If
        End With
    End If
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picMain.DrawMode = 13
   With mudtPoint
        .msglX = x
        .msglY = y
   End With
    'draw the final objects
    With mudtTool
        If .Line Then
            FinalizeLine mudtPoint, Me
        ElseIf .Circle Then
            FinalizeCircle mudtPoint, Me
        ElseIf .Elipse Then
            FinalizeElipse mudtPoint, Me
        ElseIf .Square Then
            FinalizeSquare mudtPoint, Me
        End If
    End With
End Sub

Private Sub cmdSave_Click()
        'check to see if file was already saved and has a name
    If strFilename <> strNewFile Then
        SavePicture picMain.Image, strFilename
    Else
        cmdSaveAs_Click
    End If
End Sub
