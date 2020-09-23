VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmMultiPgPreview 
   Appearance      =   0  'Flat
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   ClientHeight    =   6720
   ClientLeft      =   4650
   ClientTop       =   2220
   ClientWidth     =   5835
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmMultiPgPreview_WithChart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6720
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picFullPage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   4320
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   23
      Top             =   2400
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox picPrintPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   3885
      ScaleHeight     =   435
      ScaleWidth      =   255
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6450
      Left            =   5280
      ScaleHeight     =   6450
      ScaleWidth      =   555
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   555
      Begin VB.CommandButton cmd_print 
         Caption         =   "Print"
         Height          =   615
         Left            =   30
         Picture         =   "frmMultiPgPreview_WithChart.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Send to Printer"
         Top             =   600
         Width           =   525
      End
      Begin VB.CommandButton cmd_quit 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         Height          =   585
         Left            =   30
         Picture         =   "frmMultiPgPreview_WithChart.frx":0405
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Close"
         Top             =   15
         Width           =   525
      End
      Begin VB.CheckBox cmdFullPage 
         Caption         =   "Fit"
         Height          =   255
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1215
         Width           =   525
      End
      Begin VB.CommandButton cmdGoTo 
         Caption         =   "&Goto"
         Height          =   240
         Left            =   45
         TabIndex        =   2
         ToolTipText     =   "Goto Page"
         Top             =   2595
         Width           =   465
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   285
         Picture         =   "frmMultiPgPreview_WithChart.frx":07FE
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Next Page"
         Top             =   2220
         UseMaskColor    =   -1  'True
         Width           =   225
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   45
         Picture         =   "frmMultiPgPreview_WithChart.frx":08B8
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Prev. Page"
         Top             =   2220
         UseMaskColor    =   -1  'True
         Width           =   225
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2220
         LargeChange     =   10
         Left            =   105
         Max             =   100
         Min             =   -20
         TabIndex        =   3
         Top             =   2910
         Width           =   330
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "P 1"
         Height          =   600
         Left            =   45
         TabIndex        =   14
         Top             =   1500
         Width           =   465
      End
   End
   Begin VB.PictureBox picHScroll 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   5835
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6450
      Visible         =   0   'False
      Width           =   5835
      Begin VB.HScrollBar HScroll1 
         Height          =   270
         Left            =   0
         Max             =   100
         TabIndex        =   4
         Top             =   0
         Width           =   3765
      End
   End
   Begin VB.PictureBox picPrintOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      ForeColor       =   &H000000FF&
      Height          =   2355
      Left            =   555
      ScaleHeight     =   2325
      ScaleWidth      =   3150
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   3180
      Begin VB.TextBox txtFrom 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1695
         TabIndex        =   8
         Text            =   "1"
         Top             =   1095
         Width           =   420
      End
      Begin VB.TextBox txtTo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2475
         TabIndex        =   9
         Text            =   "1"
         Top             =   1095
         Width           =   420
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Ok"
         Height          =   360
         Left            =   2145
         TabIndex        =   11
         Top             =   1815
         Width           =   705
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   0
         Left            =   270
         Picture         =   "frmMultiPgPreview_WithChart.frx":0972
         Top             =   450
         Width           =   300
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Copy page to clipboard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   300
         Index           =   0
         Left            =   585
         TabIndex        =   5
         Top             =   480
         Width           =   2250
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Print Current Page"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   300
         Index           =   1
         Left            =   585
         TabIndex        =   6
         Top             =   810
         Width           =   1965
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   1
         Left            =   270
         Picture         =   "frmMultiPgPreview_WithChart.frx":0A0F
         Top             =   780
         Width           =   300
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   2
         Left            =   270
         Picture         =   "frmMultiPgPreview_WithChart.frx":0AAC
         Top             =   1080
         Width           =   300
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   3
         Left            =   270
         Picture         =   "frmMultiPgPreview_WithChart.frx":0B49
         Top             =   1410
         Width           =   300
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Print All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   300
         Index           =   3
         Left            =   585
         TabIndex        =   10
         Top             =   1440
         Width           =   1965
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   270
         Left            =   2175
         TabIndex        =   19
         Top             =   1125
         Width           =   345
      End
      Begin VB.Label lblPrintingPg 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   255
         TabIndex        =   18
         Top             =   1995
         Visible         =   0   'False
         Width           =   2670
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Print Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   315
         Left            =   135
         TabIndex        =   16
         Top             =   30
         Width           =   2865
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000F&
         Height          =   2250
         Left            =   30
         Top             =   30
         Width           =   3090
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Print Pages"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   300
         Index           =   2
         Left            =   585
         TabIndex        =   7
         Top             =   1110
         Width           =   1965
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4845
      Left            =   0
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   3765
   End
   Begin VB.PictureBox PictChart 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   3570
      ScaleHeight     =   900
      ScaleWidth      =   1305
      TabIndex        =   21
      Top             =   5085
      Visible         =   0   'False
      Width           =   1305
      Begin MSChart20Lib.MSChart Chart1 
         Height          =   750
         Left            =   0
         OleObjectBlob   =   "frmMultiPgPreview_WithChart.frx":0BE6
         TabIndex        =   22
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Image optArt 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   1
      Left            =   0
      Picture         =   "frmMultiPgPreview_WithChart.frx":2F3C
      Top             =   4860
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image optArt 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   0
      Left            =   555
      Picture         =   "frmMultiPgPreview_WithChart.frx":2FE9
      Top             =   4875
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "frmMultiPgPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/*************************************/
'/* Author: Morgan Haueisen
'/*         morganh@hartcom.net
'/* Copyright (c) 1999-2003
'/*************************************/
'Legal:
'        This is intended for and was uploaded to www.planetsourcecode.com
'
'        Redistribution of this code, whole or in part, as source code or in binary form, alone or
'        as part of a larger distribution or product, is forbidden for any commercial or for-profit
'        use without the author's explicit written permission.
'
'        Redistribution of this code, as source code or in binary form, with or without
'        modification, is permitted provided that the following conditions are met:
'
'        Redistributions of source code must include this list of conditions, and the following
'        acknowledgment:
'
'        This code was developed by Morgan Haueisen.  <morganh@hartcom.net>
'        Source code, written in Visual Basic, is freely available for non-commercial,
'        non-profit use at www.planetsourcecode.com.
'
'        Redistributions in binary form, as part of a larger project, must include the above
'        acknowledgment in the end-user documentation.  Alternatively, the above acknowledgment
'        may appear in the software itself, if and wherever such third-party acknowledgments
'        normally appear.

Option Explicit

'/* Used for Manifest files (Win XP)
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Public PageNumber As Integer
Private ViewPage As Integer
Private TempDir As String
Private OptionV As Integer

Private Type PanState
   x As Long
   y As Long
End Type
Dim PanSet As PanState

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
    ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
    ByVal dwRop As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVersionInfo) As Long

Private Type OSVersionInfo
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long
    PlatformID As Long
    szCSDVersion As String * 128
End Type
Private UseStretchBit As Boolean

Private Sub cmdFullPage_Click()
  Dim xmin As Single
  Dim ymin As Single
  Dim wid As Single
  Dim hgt As Single
  Dim aspect As Single
 
    '/* If already here then restore original
    If cmdFullPage.Value = 0 Then
        Picture1.Visible = True
        Picture1.SetFocus
        picFullPage.Visible = False
        Set picFullPage.Picture = Nothing
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    '/* Clear any picture and set the size and loaction
    Set picFullPage.Picture = Nothing
    If Not picHScroll.Visible Then
        picFullPage.Height = Me.Height - 100
        picFullPage.Width = picFullPage.Height * 0.773
        picFullPage.Move ((Me.Width - Picture2.Width) - picFullPage.Width) \ 2, 0
    Else
        picFullPage.Top = 50
        picFullPage.Left = 50
        picFullPage.Width = Me.Width - Picture2.Width - 100
        picFullPage.Height = picFullPage.Width * 0.773
    End If
        
    '/* Get the scale values
    aspect = Picture1.ScaleHeight / Picture1.ScaleWidth
    wid = picFullPage.ScaleWidth
    hgt = picFullPage.ScaleHeight
    
    '/* MaintainRatio
    If hgt / wid > aspect Then
        hgt = aspect * wid
        xmin = picFullPage.ScaleLeft
        ymin = (picFullPage.ScaleHeight - hgt) / 2
    Else
        wid = hgt / aspect
        xmin = (picFullPage.ScaleWidth - wid) / 2
        ymin = picFullPage.ScaleTop
    End If
    
    If UseStretchBit Then '/* NT platform
        StretchBlt picFullPage.hdc, _
            xmin, ymin, wid, hgt, _
            Picture1.hdc, _
            0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbSrcCopy
    Else
        picFullPage.PaintPicture Picture1.Picture, _
          xmin, ymin, wid, hgt, _
          0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbSrcCopy
    End If


    Picture1.Visible = False
    picFullPage.Visible = True
    picFullPage.SetFocus
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmd_print_Click()
    txtTo.Text = PageNumber + 1
    OptionV = 3
    Call optText_Click(OptionV)
    picPrintOptions.Left = Me.Width - (Picture2.Width + picPrintOptions.Width + 50)
    picPrintOptions.Visible = True
End Sub

Private Function IsNumber(ByVal CheckString As String, Optional KeyAscii As Integer = 0, Optional AllowDecPoint As Boolean = False, Optional AllowNegative As Boolean = False) As Boolean
    If KeyAscii > 0 And KeyAscii <> 8 Then
        If Not AllowNegative And KeyAscii = 45 Then KeyAscii = 0
        If Not AllowDecPoint And KeyAscii = 46 Then KeyAscii = 0
        If Not IsNumeric(CheckString & Chr(KeyAscii)) Then
            KeyAscii = False
            IsNumber = False
        Else
            IsNumber = True
        End If
    Else
        IsNumber = IsNumeric(CheckString)
    End If
End Function

Private Sub cmd_quit_Click()
    cPrint.SendToPrinter = False
    Unload Me
End Sub

Private Sub cmdGoTo_Click()
  Dim NewPageNo As Variant
    On Local Error Resume Next
    
    
    cmd_print.SetFocus
    
    NewPageNo = InputBox("Enter page number", "GoTo Page", 1)
    NewPageNo = Val(NewPageNo)
    
    If NewPageNo = 0 Then Exit Sub
    
    NewPageNo = NewPageNo - 1
    If NewPageNo > PageNumber Then NewPageNo = PageNumber
    ViewPage = NewPageNo
        
    Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
    
    picPrintOptions.Visible = False
    VScroll1.Value = 0
    HScroll1.Value = 0
    Call DisplayPages

End Sub

Private Sub cmdPrint_Click()
  Dim i As Integer
  
    '/* Prevent printing again until done
    cmd_print.SetFocus
    picPrintOptions.Enabled = False
    lblPrintingPg.Visible = True
    cmdPrint.Visible = False
    
    Select Case OptionV
    Case 0 '/* Copy to clipboard
        Clipboard.Clear
        Clipboard.SetData Picture1.Picture, vbCFBitmap
    Case 1 '/* Print current page
        lblPrintingPg.Caption = "Printing page " & ViewPage + 1
        lblPrintingPg.Refresh
        Call PrintPictureBox(Picture1, True, False)
    Case 2 '/* Print range
        For i = Val(txtFrom) - 1 To Val(txtTo) - 1
            lblPrintingPg.Caption = "Printing page " & CStr(i + 1) & " of " & txtTo
            lblPrintingPg.Refresh
            Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(i) & ".bmp")
            Call PrintPictureBox(Picture1, True, False)
        Next i
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
    Case Else '/* Print all
        cPrint.SendToPrinter = True '/* Send to Printer */
        Unload Me
    End Select
    
    '/* Restore normal view
    picPrintOptions.Enabled = True
    cmdPrint.Visible = True
    picPrintOptions.Visible = False
    lblPrintingPg.Visible = False

End Sub

Private Sub Command1_Click(Index As Integer)
    On Local Error Resume Next
    If Index = 0 Then
        ViewPage = ViewPage - 1
        If ViewPage < 0 Then ViewPage = 0
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
    Else
        ViewPage = ViewPage + 1
        If ViewPage > PageNumber Then ViewPage = PageNumber
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
    End If
    
    Picture1.Top = 0
    'Picture1.Refresh
    picPrintOptions.Visible = False
    VScroll1.Value = 0
    HScroll1.Value = 0
    Call DisplayPages
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    Call DisplayPages
    If Picture1.Width < Me.Width - Picture2.Width Then
        Picture1.Move ((Me.Width - Picture2.Width) - Picture1.Width) \ 2, 0
    End If
End Sub

Private Sub Form_Click()
    picPrintOptions.Visible = False
End Sub


Private Sub Form_Initialize()
    '/* Used for Manifest files (Win XP)
    Call InitCommonControls
End Sub

Private Sub Form_LinkOpen(Cancel As Integer)

End Sub

Private Sub Form_Load()
  Dim OSV As OSVersionInfo
  Const VER_PLATFORM_WIN32_NT = 2
    
    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
        If OSV.PlatformID = VER_PLATFORM_WIN32_NT Then
            UseStretchBit = True
        Else
            UseStretchBit = False
        End If
    End If

    Me.Move 0, 0, Screen.Width, Screen.Height
    Picture1.Move 0, 0

    VScroll1.Height = Me.Height - cmdGoTo.Top - cmdGoTo.Height - 500
    HScroll1.Width = Me.Width - Picture2.Width - 500
    
    TempDir = Environ("TEMP") & "\"
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim tFileName As String
    
    '/* Remove preview pages
    tFileName = Dir(TempDir & "PPview*.bmp")
    If tFileName > vbNullString Then
        Do
            Kill TempDir & tFileName
            tFileName = Dir(TempDir & "PPview*.bmp")
        Loop Until tFileName = vbNullString
    End If
    
    PageNumber = 0
    ViewPage = 0
    Set frmMultiPgPreview = Nothing
End Sub


Private Sub HScroll1_Change()
    On Local Error Resume Next
    Picture1.Left = -(HScroll1.Value)
    HScroll1.SetFocus
    On Local Error GoTo 0
End Sub

Private Sub HScroll1_KeyUp(KeyCode As Integer, Shift As Integer)
   On Local Error Resume Next
    Select Case KeyCode
    Case 38 '/* Arrow up
        VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
    Case 40 '/* Arrow down
        VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
    Case 33 '/* PageUp
        Call Command1_Click(0)
    Case 34 '/* PageDown
        Call Command1_Click(1)
    Case 71 '/* G
        Call cmdGoTo_Click
    Case 35, 36 '/* Home, End
      Dim NewPageNo As Long
        If KeyCode = 36 Then
            NewPageNo = 0
        Else
            NewPageNo = PageNumber
        End If
        ViewPage = NewPageNo
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
        picPrintOptions.Visible = False
        VScroll1.Value = 0
        HScroll1.Value = 0
        Call DisplayPages
    End Select

End Sub


Private Sub optPrint_Click(Index As Integer)
  Dim i As Byte
    OptionV = Index
    For i = 0 To 3
        If Index = i Then
            optPrint(i).Picture = optArt(1).Picture
        Else
            optPrint(i).Picture = optArt(0).Picture
        End If
    Next i

End Sub

Private Sub optText_Click(Index As Integer)
  Dim i As Byte
    OptionV = Index
    For i = 0 To 3
        If Index = i Then
            optPrint(i).Picture = optArt(1).Picture
        Else
            optPrint(i).Picture = optArt(0).Picture
        End If
    Next i

End Sub


Private Sub picFullPage_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Decode_KeyUp(KeyCode, Shift)
End Sub


Private Sub Decode_KeyUp(KeyCode As Integer, Shift As Integer)
   On Local Error Resume Next
    Select Case KeyCode
    Case 38 '/* Arrow up
        VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
    Case 40 '/* Arrow down
        VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
    Case 37 '/* Arrow left
        If HScroll1.Visible = False Then
            Call Command1_Click(0)
        Else
            HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
        End If
    Case 39 '/* Arrow right
        If HScroll1.Visible = False Then
            Call Command1_Click(1)
        Else
            HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
        End If
    Case 33 '/* PageUp
        Call Command1_Click(0)
    Case 34 '/* PageDown
        Call Command1_Click(1)
    Case 71 '/* G
        Call cmdGoTo_Click
    Case 35, 36 '/* Home, End
      Dim NewPageNo As Long
        If KeyCode = 36 Then
            NewPageNo = 0
        Else
            NewPageNo = PageNumber
        End If
        ViewPage = NewPageNo
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
        picPrintOptions.Visible = False
        VScroll1.Value = 0
        HScroll1.Value = 0
        Call DisplayPages
    End Select
End Sub

Private Sub Picture1_Click()
    picPrintOptions.Visible = False
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Decode_KeyUp(KeyCode, Shift)
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Local Error Resume Next
   If Button = vbLeftButton And Shift = 0 Then
      PanSet.x = x
      PanSet.y = y
      MousePointer = vbSizePointer
   End If
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim nTop As Integer, nLeft As Integer

   On Local Error Resume Next

   If Button = vbLeftButton And Shift = 0 Then

      '/* new coordinates?
      With Picture1
         nTop = -(.Top + (y - PanSet.y))
         nLeft = -(.Left + (x - PanSet.x))
      End With

      '/* Check limits
      With VScroll1
         If .Visible Then
            If nTop < .Min Then
               nTop = .Min
            ElseIf nTop > .Max Then
               nTop = .Max
            End If
         Else
            nTop = -Picture1.Top
         End If
      End With

      With HScroll1
         If .Visible Then
            If nLeft < .Min Then
               nLeft = .Min
            ElseIf nLeft > .Max Then
               nLeft = .Max
            End If
         Else
            nLeft = -Picture1.Left
         End If
      End With

      Picture1.Move -nLeft, -nTop

   End If

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Local Error Resume Next
   If Button = vbLeftButton And Shift = 0 Then
      If VScroll1.Visible Then VScroll1.Value = -(Picture1.Top)
      If HScroll1.Visible Then HScroll1.Value = -(Picture1.Left)
   End If
   MousePointer = vbDefault
End Sub


Private Sub txtFrom_Change()
    If Val(txtFrom) < 1 Then txtFrom = 1
    If Val(txtFrom) > Val(txtTo) Then txtFrom = txtTo
End Sub

Private Sub txtFrom_GotFocus()
    txtFrom.SelStart = 0
    txtFrom.SelLength = Len(txtFrom)
End Sub


Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 38  '/* "+"
        txtFrom = txtFrom + 1
        KeyCode = False
    Case 40  '/* "-"
        txtFrom = txtFrom - 1
        KeyCode = False
    End Select
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    IsNumber txtFrom, KeyAscii, False, False
End Sub


Private Sub txtTo_Change()
    If Val(txtTo) > PageNumber + 1 Then txtTo = PageNumber + 1
    If Val(txtTo) < Val(txtFrom) Then txtTo = txtFrom
End Sub

Private Sub txtTo_GotFocus()
    txtTo.SelStart = 0
    txtTo.SelLength = Len(txtTo)
End Sub


Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 38  '/* "+"
        txtTo = txtTo + 1
        KeyCode = False
    Case 40  '/* "-"
        txtTo = txtTo - 1
        KeyCode = False
    End Select
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    IsNumber txtTo, KeyAscii, False, False
End Sub


Private Sub VScroll1_Change()
    On Local Error Resume Next
    Picture1.Top = -(VScroll1.Value)
    VScroll1.SetFocus
    On Local Error GoTo 0
End Sub


Private Sub DisplayPages()
    Label1 = CStr(ViewPage + 1) & vbNewLine & "-- of --" & vbNewLine & CStr(PageNumber + 1)
    
    If Picture1.Width > Me.Width - Picture2.Width Then
        picHScroll.Visible = True
    Else
        picHScroll.Visible = False
    End If

    If Picture1.Height >= Me.Height Then
        VScroll1.Visible = True
    Else
        VScroll1.Visible = False
    End If
    
    If picFullPage.Visible Then cmdFullPage_Click

End Sub
Private Sub PrintPictureBox(pBox As PictureBox, _
                           Optional ScaleToFit As Boolean = True, _
                           Optional MaintainRatio As Boolean = True)
 
 Dim xmin As Single
 Dim ymin As Single
 Dim wid As Single
 Dim hgt As Single
 Dim aspect As Single
 
    Screen.MousePointer = vbHourglass
    
    If Not ScaleToFit Then
        wid = Printer.ScaleX(pBox.ScaleWidth, pBox.ScaleMode, Printer.ScaleMode)
        hgt = Printer.ScaleY(pBox.ScaleHeight, pBox.ScaleMode, Printer.ScaleMode)
        xmin = (Printer.ScaleWidth - wid) / 2
        ymin = (Printer.ScaleHeight - hgt) / 2
    Else
        aspect = pBox.ScaleHeight / pBox.ScaleWidth
        wid = Printer.ScaleWidth
        hgt = Printer.ScaleHeight
        
        If MaintainRatio Then
            If hgt / wid > aspect Then
                hgt = aspect * wid
                xmin = Printer.ScaleLeft
                ymin = (Printer.ScaleHeight - hgt) / 2
            Else
                wid = hgt / aspect
                xmin = (Printer.ScaleWidth - wid) / 2
                ymin = Printer.ScaleTop
            End If
        End If
    End If
    
    Printer.PaintPicture pBox.Picture, xmin, ymin, wid, hgt, , , , , vbSrcCopy
    Printer.EndDoc

    Screen.MousePointer = vbDefault

End Sub


Private Sub VScroll1_KeyUp(KeyCode As Integer, Shift As Integer)
   On Local Error Resume Next
    Select Case KeyCode
    Case 37, 33 '/* Arrow left, PageUp
        If HScroll1.Visible = False Then
            Call Command1_Click(0)
        Else
            HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
        End If
    Case 39, 34 '/* Arrow right, PageDown
        If HScroll1.Visible = False Then
            Call Command1_Click(1)
        Else
            HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
        End If
    Case 71 '/* G
        Call cmdGoTo_Click
    Case 35, 36 '/* Home, End
      Dim NewPageNo As Long
        If KeyCode = 36 Then
            NewPageNo = 0
        Else
            NewPageNo = PageNumber
        End If
        ViewPage = NewPageNo
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
        picPrintOptions.Visible = False
        VScroll1.Value = 0
        HScroll1.Value = 0
        Call DisplayPages
    End Select
End Sub


