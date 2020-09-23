VERSION 5.00
Begin VB.Form frmChartDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmChartDemo"
   ClientHeight    =   3045
   ClientLeft      =   4350
   ClientTop       =   3540
   ClientWidth     =   4830
   Icon            =   "frmChartDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Preview/Print Chart Option 2"
      Height          =   555
      Left            =   1230
      TabIndex        =   2
      Top             =   2280
      Width           =   2490
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Preview/Print Chart"
      Height          =   555
      Left            =   1230
      TabIndex        =   0
      Top             =   1740
      Width           =   2490
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   $"frmChartDemo.frx":000C
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   195
      TabIndex        =   1
      Top             =   135
      Width           =   4485
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmChartDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Command1.Enabled = False
    Set cPrint = New clsMultiPgPreview
    cPrint.Orientation = PagePortrait
    cPrint.SendToPrinter = False
    cPrint.pStartDoc
    cPrint.pChart 4, 3.5, 1.5, 1.5, 4, 6, False, True
    cPrint.pFooter
    cPrint.pEndDoc
    Set cPrint = Nothing
    Command1.Enabled = True
End Sub


Private Sub Command2_Click()
    frmChartOption2.Show
End Sub


