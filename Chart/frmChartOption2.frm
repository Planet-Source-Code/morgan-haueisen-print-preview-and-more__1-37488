VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmChartOption2 
   Caption         =   "frmChartOption2"
   ClientHeight    =   5055
   ClientLeft      =   3870
   ClientTop       =   2325
   ClientWidth     =   5460
   Icon            =   "frmChartOption2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picToolBar 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   5400
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4470
      Width           =   5460
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   450
         Left            =   180
         TabIndex        =   3
         Top             =   45
         Width           =   1035
      End
   End
   Begin VB.PictureBox PictChart 
      AutoRedraw      =   -1  'True
      Height          =   3000
      Left            =   0
      ScaleHeight     =   2940
      ScaleWidth      =   6135
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6195
      Begin MSChart20Lib.MSChart Chart1 
         Height          =   1680
         Left            =   0
         OleObjectBlob   =   "frmChartOption2.frx":000C
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   2430
      End
   End
End
Attribute VB_Name = "frmChartOption2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias _
           "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
           ByVal wParam As Long, ByVal lParam As Long) As Long

Private Sub cmdPrint_Click()
  Dim WinState As Integer
  
    'WinState = Me.WindowState
    'Me.WindowState = vbMaximized
    cmdPrint.Enabled = False
    Me.MousePointer = vbHourglass
    DoEvents
    
    Set cPrint = New clsMultiPgPreview
    cPrint.Orientation = PagePortrait
    cPrint.SendToPrinter = False
    cPrint.pStartDoc
    Call PrintChart(1.5, 1.5, 4, 6)
    cPrint.pFooter
    cPrint.pEndDoc
    
    Set cPrint = Nothing
    cmdPrint.Enabled = True
    'Me.WindowState = WinState
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Chart1.Visible = True
    PictChart.Visible = True
End Sub

Private Sub Form_Resize()
    On Local Error Resume Next
    With PictChart
        .Top = 0
        .Left = 0
        .Height = Me.Height - 1000
        .Width = Me.Width
    End With
    With Chart1
        .Top = 0
        .Left = 0
        .Height = PictChart.Height
        .Width = PictChart.Width
    End With
End Sub


Public Sub PrintChart(Optional LeftMargin As Single = -1, _
                        Optional TopMargin As Single = -1, _
                        Optional pWidth As Single = 0, _
                        Optional pHeight As Single = 0, _
                        Optional ScaleToFit As Boolean = False, _
                        Optional MaintainRatio As Boolean = True)
                        
  '/* Required to move the chart to a picturebox
  '/*      Private Declare Function SendMessage Lib "user32" Alias _
               "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
               ByVal wParam As Long, ByVal lParam As Long) As Long

  
  Const WM_PAINT = &HF

    PictChart.ScaleMode = vbTwips
    PictChart.AutoRedraw = True

    SendMessage Chart1.hwnd, WM_PAINT, PictChart.hdc, 0
    PictChart.Picture = PictChart.Image
    
    '/* Print/Preview Picture
    cPrint.pPrintPicture PictChart.Picture, LeftMargin, TopMargin, pWidth, pHeight, ScaleToFit, MaintainRatio
    
    PictChart.Picture = Nothing
 
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmChartOption2 = Nothing
End Sub


