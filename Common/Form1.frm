VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5715
   ClientLeft      =   2745
   ClientTop       =   1785
   ClientWidth     =   4875
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   4875
   Begin VB.TextBox Text1 
      Height          =   1035
      Left            =   255
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Text            =   "Form1.frx":000C
      Top             =   4455
      Width           =   4560
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   4875
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   4875
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":0A43
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   135
         TabIndex        =   7
         Top             =   45
         Width           =   4620
      End
   End
   Begin VB.PictureBox picPrinting 
      BackColor       =   &H80000005&
      Height          =   3180
      Left            =   5295
      ScaleHeight     =   3120
      ScaleWidth      =   3615
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4410
      Visible         =   0   'False
      Width           =   3675
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Printing... Please wait"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   3405
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Room Tags"
      Height          =   495
      Left            =   1215
      TabIndex        =   5
      Top             =   2910
      Width           =   2595
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Example Of EndOfPage Check"
      Height          =   495
      Left            =   1215
      TabIndex        =   1
      Top             =   1920
      Width           =   2595
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Preview Instructions"
      Height          =   495
      Left            =   1215
      TabIndex        =   0
      Top             =   1425
      Width           =   2595
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ignore CF and LF"
      Height          =   495
      Left            =   1215
      TabIndex        =   4
      Top             =   2415
      Width           =   2595
   End
   Begin VB.CommandButton Command5 
      Caption         =   "House Order Form"
      Height          =   375
      Left            =   1215
      TabIndex        =   8
      Top             =   3405
      Width           =   2595
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Print the Multi-Line Textbox Below"
      Height          =   570
      Left            =   1215
      TabIndex        =   9
      Top             =   3780
      Width           =   2595
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   0
      X2              =   4845
      Y1              =   1305
      Y2              =   1305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Const C0 = 1
  Const C1 = 1.1
  Const C2 = 2.2
  Const C3 = 7
  Dim memCurrentYt As Single, memCurrentYb As Single
  
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter:
    Screen.MousePointer = vbHourglass
    picPrinting.Visible = True
    DoEvents
    
    cPrint.pStartDoc
    
    '/* Print Cover page ********************************************************************
    cPrint.pPrintPicture LoadPicture(App.Path & "\cover.jpg"), , , , , , False
    cPrint.pFontName
    cPrint.pBox 0.5, 0.5, cPrint.GetPaperWidth - 1, cPrint.GetPaperHeight - 1
    cPrint.FontSize = 24
    cPrint.FontBold = True
    cPrint.CurrentY = 3
    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HC0E0FF, , vbFSSolid
    cPrint.FontTransparent = True
    cPrint.pCenter " Print Preview and More "
    cPrint.FontSize = 14
    cPrint.pPrint
    cPrint.pCenter "By: Morgan Haueisen"
    cPrint.pPrint
    cPrint.FontBold = False
    cPrint.FontSize = 12

    '/* Two different ways to center a long text string
    cPrint.pMultiline "This is page 1 of 6.  Use the horizontal / vertical scroll bars and buttons " & _
                      "on the side bar to move the pages.  You can also click and drag the page " & _
                      "with your mouse, use the arrow keys to move the page around and " & _
                      "change between pages.  You can also use keys PageUp, PageDown, Home, End, or G (goto).", 2, cPrint.GetPaperWidth - 2

    'cPrint.pCenterMultiline "This is page 1 of 6.  Use the horizontal / vertical scroll bars and buttons " & _
                      "on the side bar to move the pages.  You can also click and drag the page " & _
                      "with your mouse, use the arrow keys to move the page around and " & _
                      "change between pages.  You can also use keys PageUp, PageDown, Home, End, or G (goto).", 2, cPrint.GetPaperWidth - 2

    cPrint.pPrint
    cPrint.ForeColor = vbRed
    cPrint.pCenter "Please look at Readme.htm for additional information."
    cPrint.ForeColor = vbBlack

    cPrint.CurrentY = cPrint.GetPaperHeight - 1.5
    cPrint.pPrintedDate True
    cPrint.pNewPage


    '/* Page 2 *********************************************************************************
    cPrint.pHeader "A Print Preview Demo Project", "By: Morgan Haueisen"

    cPrint.FontSize = 11
    cPrint.pPrint "This Print Preview code may not be the best solution but it gets" & _
                      " the job done and the only page limitation is the size of your hard drive. All of the other alternatives" & _
                      " for print preview that I could think of crash at about 45 pages because" & _
                      " of memory limitations.  All printing locations are entered in inches" & _
                      " (or centimeters depending on the setting in Class_Initialize); no more tab(??).", 0.75

    cPrint.pQuarterSpace
    cPrint.pPrint "For example, this line started printing 1.09 inches from the left margin. " & _
                  "The default scale mode is inches; you can change the default scale mode by " & _
                  "modifying the oScaleMode variable in the class's Class_Initialize sub.", 1.09

    cPrint.FontItalic = True
    cPrint.pHalfSpace
    cPrint.pPrint "Preview functions include:", 1.25
    cPrint.FontItalic = False
    cPrint.pBullet 1.5
    cPrint.pPrint "Horizontal and vertical scroll bars appear as required.  Use these to move the page around the screen or just click and drag the page with your mouse.  You can also use the arrow keys."
    cPrint.pBullet 1.5
    cPrint.pPrint "Print current page, a range of pages, all pages, or copy the page to the clipboard."
    cPrint.pBullet 1.5
    cPrint.pPrint "Move forward or backward through the pages."
    cPrint.pBullet 1.5
    cPrint.pPrint "Go to a specific page."
    cPrint.pBullet 1.5
    cPrint.pPrint "Fit page to screen (click the button again to restore the view)."
    cPrint.pHalfSpace
    cPrint.FontItalic = True
    cPrint.pPrint "Preview weaknesses:", 1.25
    cPrint.FontItalic = False
    cPrint.pBullet 1.5
    cPrint.pPrint "The preview may not always match the printed page; one may have a line or two more then the other."
    cPrint.pBullet 1.5
    cPrint.pPrint "When selecting a print option other then All, the printed page does not look as sharp as it does when you print all.  Also, you can not print double sided."
    cPrint.pPrint
    cPrint.pPrint "I did not add any zoom functionally; the preview is already large enough to see and I have spent far more time on this then I intended.  You can add this functionality by adding a second picture box and do the scaling using API StretchBlt or picturebox.PaintPicture function (see the PrintPictureBox sub in the frmMultiPgPreview form).", 0.75
    cPrint.pPrint
    cPrint.pPrint "I hope you find this code useful; please e-mail me if you have any suggestions on how to improve it.", 0.75
    cPrint.pPrint
    cPrint.pPrint "Cordially,", 1
    cPrint.pPrint "Morgan Haueisen", 1
    cPrint.FontUnderline = True
    cPrint.ForeColor = vbBlue
    cPrint.pPrint "morganh@hartcom.net", 1
    cPrint.pPrint
    cPrint.FontUnderline = False
    cPrint.ForeColor = vbBlack
    cPrint.pPrint "To view all the projects I have submitted go to:", 0.75
    cPrint.FontSize = 8
    cPrint.FontUnderline = True
    cPrint.ForeColor = vbBlue
    cPrint.FontName = "Arial"
    cPrint.pPrint "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=885253927&strAuthorName=Morgan%20Haueisen&txtMaxNumberOfEntriesPerPage=25", 1
    cPrint.FontUnderline = False
    cPrint.p15Space
    cPrint.pFontName
    cPrint.ForeColor = vbRed
    cPrint.FontBold = True
    cPrint.FontSize = 11
    cPrint.pPrint "Scroll down and check out the footer.", 1
    cPrint.pPrint "For additional printing examples, please see the last page.", 1
    cPrint.ForeColor = vbBlack
    cPrint.FontBold = False
    cPrint.pPrint
    cPrint.pPrint "Download a project that uses this print-preview code at: ", 0.5
    cPrint.ForeColor = vbBlue
    cPrint.FontUnderline = True
    cPrint.FontSize = 8
    cPrint.FontName = "Arial"
    cPrint.pPrint "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37077&lngWId=1", 0.75
    cPrint.pFontName
    cPrint.FontUnderline = False
    cPrint.ForeColor = vbBlack
    cPrint.pPrint

    cPrint.pFooter


    '/* Page 3 ***********************************************************************************
    cPrint.pNewPage
    cPrint.pPrint
    cPrint.FontSize = 12
    cPrint.FontBold = True
    cPrint.pPrint "Printing and Print Preview Help  ", 0.5, True
    cPrint.FontSize = 10
    cPrint.FontBold = False
    cPrint.pPrint "(all the printing coordinates are in inches.)"
    cPrint.pPrint
    cPrint.pDoubleLine
    cPrint.FontSize = 12
    cPrint.FontBold = True
    cPrint.pPrint
    cPrint.FontName = "Arial"
    cPrint.pPrint "Printing:", 0.5
    cPrint.pHalfSpace
    cPrint.FontBold = False
    cPrint.FontSize = 9

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HFFFFD0, , vbFSSolid
    cPrint.pPrint "pStartDoc", C1, True
    cPrint.pPrint "start a document", C2

    cPrint.pPrint "pEndDoc", C1, True
    cPrint.pPrint "end a document", C2

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HFFFFD0, , vbFSSolid
    cPrint.pPrint "p15Space", C1, True
    cPrint.pPrint "print a Line feed that is 1-1/2 times your current Font Size", C2

    cPrint.pPrint "pBox", C1, True
    cPrint.pPrint "print a box  (open or filled)", C2

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HFFFFD0, , vbFSSolid
    cPrint.pPrint "pCancled", C1, True
    cPrint.pPrint "print the words '**** PRINTING CANCLED ****'", C2

    cPrint.pPrint "pCenter", C1, True
    cPrint.pPrint "print the text centered on the page ", C2, True
    cPrint.FontSize = 8
    cPrint.pPrint "(uses pCenterMultiLine when necessary)"
    cPrint.FontSize = 9

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HFFFFD0, , vbFSSolid
    cPrint.pPrint "pCircle", C1, True
    cPrint.pPrint "print a circle (open or filled)", C2

    cPrint.pPrint "pDots", C1, True
    cPrint.pPrint "print dots from your current position, you pick the end point", C2

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HFFFFD0, , vbFSSolid
    cPrint.pPrint "pDoubleLine", C1, True
    cPrint.pPrint "print a double line", C2

    cPrint.pPrint "pEndOfPage", C1, True
    cPrint.pPrint "check to see if you are close to the end of the page", C2

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HFFFFD0, , vbFSSolid
    cPrint.pPrint "pFontName", C1, True
    cPrint.pPrint "change the printing font, if no font name is supplied, the default is Times New Roman", C2

    cPrint.pPrint "pFooter", C1, True
    cPrint.pPrint "print date, time, and page number at the bottom of the page", C2

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HFFFFD0, , vbFSSolid
    cPrint.pPrint "pHalfSpace", C1, True
    cPrint.pPrint "print a Line feed at 1/2 your current Font Size", C2

    cPrint.pPrint "pLine", C1, True
    cPrint.pPrint "print a Line, ", C2, True
    cPrint.FontSize = 8
    cPrint.pPrint "(if lineweight=0 then the DrawWidth value is used.)"
    cPrint.FontSize = 9

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HFFFFD0, , vbFSSolid
    cPrint.pPrint "pMultiline", C1, True
    cPrint.pPrint "split a long string into several printed lines.  You provide the left and right margins.", C2

    cPrint.pPrint "pNewPage", C1, True
    cPrint.pPrint "start a new page", C2

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HFFFFD0, , vbFSSolid
    cPrint.pPrint "pPrint", C1, True
    cPrint.pPrint "print a string (uses pMultiline when necessary)", C2

    cPrint.pPrint "pPrintedDate", C1, True
    cPrint.pPrint "print Printed: And today 's date and time", C2

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HFFFFD0, , vbFSSolid
    cPrint.pPrint "pQuarterSpace", C1, True
    cPrint.pPrint "print a Line feed at 1/4 your current Font Size", C2

    cPrint.pPrint "pRightJust", C1, True
    cPrint.pPrint "right justifies your text string", C2

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HFFFFD0, , vbFSSolid
    cPrint.pPrint "pRightTab", C1, True
    cPrint.pPrint "print text using the right margin as the starting point, moving left.", C2

    cPrint.pPrint "pSpaces", C1, True
    cPrint.pPrint "print spaces (used to print a colored stripe by changing the back color)", C2

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HFFFFD0, , vbFSSolid
    cPrint.pPrint "pVerticalLine", C1, True
    cPrint.pPrint "print a vertical line on your page.", C2

    cPrint.pPrint "pPrintPicture", C1, True
    cPrint.pPrint "print a picture at a specified location and size or fit to page", C2

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HFFFFD0, , vbFSSolid
    cPrint.pPrint "pHeader", C1, True
    cPrint.pPrint "print a 1 or 2 line header at the top of the page.", C2

    cPrint.pPrint "pCenterMultiLine", C1, True
    cPrint.pPrint "split a long string into several centered lines.  You provide the left and right margins.", C2

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HFFFFD0, , vbFSSolid
    cPrint.pPrint "pBullet", C1, True
    cPrint.pPrint "print a bullet character", C2

    cPrint.pPrint "pPrintRotate", C1, True
    cPrint.pPrint "print text at any angle.", C2

    cPrint.pPrint
    cPrint.pPrint
    cPrint.pFontName
    cPrint.FontSize = 11
    cPrint.ForeColor = vbRed
    cPrint.pPrint "Uses pBox to print a light blue box around every other line.", 0.5
    cPrint.ForeColor = vbBlack

    cPrint.pFooter


    '/* Page 4 *************************************************************************
    cPrint.pNewPage
    cPrint.pPrint
    cPrint.FontSize = 12
    cPrint.FontBold = True
    cPrint.pPrint "Printing and Print Preview Help ", 0.5, True
    cPrint.FontSize = 10
    cPrint.FontBold = False
    cPrint.pPrint "(all the printing coordinates are in inches.)"
    cPrint.pPrint
    cPrint.pDoubleLine
    cPrint.FontSize = 12
    cPrint.FontBold = True
    cPrint.pPrint
    cPrint.FontName = "Arial"
    cPrint.pPrint "Properties (Set or Get):", 0.5
    cPrint.pHalfSpace
    cPrint.FontBold = False
    cPrint.FontSize = 9

    memCurrentYt = cPrint.CurrentY + 0.03 '/* remember starting point for vertical lines

    cPrint.pLine C0, C3, 4
    cPrint.pPrint "CurrentX", C1, True
    cPrint.pPrint "current X position", C2

    cPrint.pLine C0, C3, 4
    cPrint.pPrint "CurrentY", C1, True
    cPrint.pPrint "current Y position", C2

    cPrint.pLine C0, C3, 4
    cPrint.pPrint "DrawWidth", C1, True
    cPrint.pPrint "for printing boxes and circles", C2

    cPrint.pLine C0, C3, 4
    cPrint.pPrint "FontBold", C1, True
    cPrint.pPrint "bold Font", C2

    cPrint.pLine C0, C3, 4
    cPrint.pPrint "FontItalic", C1, True
    cPrint.pPrint "italic Font", C2

    cPrint.pLine C0, C3, 4
    cPrint.pPrint "FontName", C1, True
    cPrint.pPrint "font name (this is different then pFontName in that you must supply a font name)", C2

    cPrint.pLine C0, C3, 4
    cPrint.pPrint "FontSize", C1, True
    cPrint.pPrint "font Size", C2

    cPrint.pLine C0, C3, 4
    cPrint.pPrint "FontStrikethru", C1, True
    cPrint.pPrint "strike through font", C2

    cPrint.pLine C0, C3, 4
    cPrint.pPrint "FontTransparent", C1, True
    cPrint.pPrint "changes the font's background", C2

    cPrint.pLine C0, C3, 4
    cPrint.pPrint "FontUnderline", C1, True
    cPrint.pPrint "underline Font", C2

    cPrint.pLine C0, C3, 4
    cPrint.pPrint "ForeColor", C1, True
    cPrint.pPrint "font's foreground color", C2

    cPrint.pLine C0, C3, 4
    cPrint.pPrint "BackColor", C1, True
    cPrint.pPrint "font's background color, setting it to -1 is the same as setting FontTransparent = True", C2

    cPrint.pLine C0, C3, 4
    cPrint.pPrint "Orientation", C1, True
    cPrint.pPrint "printer 's orientation (portrait or landscape)", C2

    cPrint.pLine C0, C3, 4
    cPrint.pPrint "PrintCopies", C1, True
    cPrint.pPrint "how many copies the printer will print", C2

    cPrint.pLine C0, C3, 4
    cPrint.pPrint "SendToPrinter", C1, True
    cPrint.pPrint "determines if the output goes to the printer or the screen", C2

    memCurrentYb = cPrint.CurrentY + 0.03 '/* remember ending point for vertical lines

    cPrint.pLine C0, C3, 4
    cPrint.pVerticalLine C0, memCurrentYt, memCurrentYb, 4
    cPrint.pVerticalLine C2 - 0.1, memCurrentYt, memCurrentYb, 4
    cPrint.pVerticalLine C3, memCurrentYt, memCurrentYb, 4

    cPrint.pPrint
    cPrint.pPrint
    cPrint.pFontName
    cPrint.FontSize = 11
    cPrint.ForeColor = vbRed
    cPrint.pPrint "Uses pLine and pVerticalLine to draw the grid with a line weight = 4.", 0.5
    cPrint.ForeColor = vbBlack

    cPrint.pFooter


    '/* Page 5 *********************************************************************************
    cPrint.pNewPage
    cPrint.pPrint
    cPrint.FontSize = 12
    cPrint.FontBold = True
    cPrint.pPrint "Printing and Print Preview Help ", 0.5, True
    cPrint.FontSize = 10
    cPrint.FontBold = False
    cPrint.pPrint "(all the printing coordinates are in inches.)"
    cPrint.pPrint
    cPrint.pDoubleLine

    '/* Show Grid Option #1
    cPrint.FontSize = 12
    cPrint.FontBold = True
    cPrint.pPrint
    cPrint.FontName = "Arial"
    cPrint.pPrint "Functions:", 0.5
    cPrint.pHalfSpace
    cPrint.FontBold = False
    cPrint.FontSize = 9
    cPrint.FontTransparent = True

    memCurrentYt = cPrint.CurrentY + 0.03 '/* remember starting point for vertical lines

    cPrint.pLine C0, cPrint.GetPaperWidth - 1
    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HC0FFFF, , vbFSSolid
    cPrint.pPrint "GetFormalCase", C1, True
    cPrint.pPrint "returns the formal case of a string", C2

    cPrint.pLine C0, cPrint.GetPaperWidth - 1
    cPrint.pPrint "GetPage", C1, True
    cPrint.pPrint "returns the current page number", C2

    cPrint.pLine C0, cPrint.GetPaperWidth - 1
    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HC0FFFF, , vbFSSolid
    cPrint.pPrint "GetPaperHeight", C1, True
    cPrint.pPrint "returns the printer's printable area (height in inches)", C2

    cPrint.pLine C0, cPrint.GetPaperWidth - 1
    cPrint.pPrint "GetPaperWidth", C1, True
    cPrint.pPrint "returns the printer's printable area (width in inches)", C2

    cPrint.pLine C0, cPrint.GetPaperWidth - 1
    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HC0FFFF, , vbFSSolid
    cPrint.pPrint "GetTextWidth", C1, True
    cPrint.pPrint "returns the width of a text string in inches", C2

    cPrint.pLine C0, cPrint.GetPaperWidth - 1
    cPrint.pPrint "GetStripQuotes", C1, True
    cPrint.pPrint "removes the quotes from the beginning and end of a text string", C2

    cPrint.pLine C0, cPrint.GetPaperWidth - 1
    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HC0FFFF, , vbFSSolid
    cPrint.pPrint "GetRemoveCRLF", C1, True
    cPrint.pPrint "removes all line feeds imbedded in a string", C2

    memCurrentYb = cPrint.CurrentY + 0.03 '/* remember ending point for vertical lines

    cPrint.pLine C0, cPrint.GetPaperWidth - 1
    cPrint.pVerticalLine C0, memCurrentYt, memCurrentYb
    cPrint.pVerticalLine C2 - 0.1, memCurrentYt, memCurrentYb
    cPrint.pVerticalLine cPrint.GetPaperWidth - 1, memCurrentYt, memCurrentYb

    cPrint.pPrint
    cPrint.pPrint
    cPrint.pFontName
    cPrint.FontSize = 11
    cPrint.ForeColor = vbRed
    cPrint.pPrint "Uses pLine and pVerticalLine to draw the grid with a line weight = 1 and pBox to draw the colored background.", 0.5
    cPrint.ForeColor = vbBlack

    cPrint.p15Space
    cPrint.p15Space

    '/* Show Grid Option #2
    cPrint.FontSize = 12
    cPrint.FontBold = True
    cPrint.pPrint
    cPrint.FontName = "Arial"
    cPrint.pPrint "Functions:", 0.5
    cPrint.pHalfSpace
    cPrint.FontBold = False
    cPrint.FontSize = 9
    cPrint.FontTransparent = True

    memCurrentYt = cPrint.CurrentY '/* remember starting point for vertical lines

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , , &HC0FFFF, vbFSSolid
    cPrint.pPrint "GetFormalCase", C1, True
    cPrint.pPrint "returns the formal case of a string", C2

    cPrint.pPrint "GetPage", C1, True
    cPrint.pPrint "returns the current page number", C2

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , , &HC0FFFF, vbFSSolid
    cPrint.pPrint "GetPaperHeight", C1, True
    cPrint.pPrint "returns the printers printable area (height in inches)", C2

    cPrint.pPrint "GetPaperWidth", C1, True
    cPrint.pPrint "returns the printers printable area (width in inches)", C2

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , , &HC0FFFF, vbFSSolid
    cPrint.pPrint "GetTextWidth", C1, True
    cPrint.pPrint "returns the width of a text string in inches", C2

    cPrint.pPrint "GetStripQuotes", C1, True
    cPrint.pPrint "removes the quotes from the beginning and end of a text string", C2

    cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , , &HC0FFFF, vbFSSolid
    cPrint.pPrint "GetRemoveCRLF", C1, True
    cPrint.pPrint "removes all line feeds imbedded in a string", C2

    memCurrentYb = cPrint.CurrentY '/* remember ending point for vertical lines

    cPrint.pVerticalLine C0, memCurrentYt, memCurrentYb
    cPrint.pVerticalLine C2 - 0.1, memCurrentYt, memCurrentYb
    cPrint.pVerticalLine cPrint.GetPaperWidth - 1, memCurrentYt, memCurrentYb

    cPrint.pPrint
    cPrint.pPrint
    cPrint.pFontName
    cPrint.FontSize = 11
    cPrint.ForeColor = vbRed
    cPrint.pPrint "Uses pVerticalLine and pBox to draw the colored background and grid.", 0.5
    cPrint.ForeColor = vbBlack

    cPrint.pFooter


    '/* Page 6 *********************************************************************************
    cPrint.pNewPage
    cPrint.pFontName
    cPrint.pHeader "MORE EXAMPLES"
    cPrint.p15Space
    cPrint.pPrint "Right justify a string", 1, True
    cPrint.pDots 4.3
    cPrint.pRightJust "45.00", 5
    cPrint.pPrint "Right justify again", 1, True
    cPrint.pDots 4.3
    cPrint.pRightJust "1,445.00", 5
    cPrint.pPrint "Right justify once more", 1, True
    cPrint.pDots 4.3
    cPrint.pRightJust "21,445.00", 5
    cPrint.pPrint
    cPrint.pRightTab "This text is 1.25 inches from the right margin.", 1.25
    cPrint.pRightTab "Print Portrait and Landscape to see the change.", 1.25
    
    cPrint.p15Space
    cPrint.p15Space
    cPrint.p15Space
    
    cPrint.pCircle 1, cPrint.CurrentY, 0.25, vbBlue
    cPrint.DrawWidth = 6
    cPrint.pCircle 2, cPrint.CurrentY, 0.25, vbBlue
    cPrint.pCircle 3, cPrint.CurrentY, 0.25, vbBlue, vbRed, vbFSSolid
    cPrint.DrawWidth = 1
    cPrint.pCircle 4, cPrint.CurrentY, 0.25, vbBlue, vbRed, vbCross
    cPrint.pCircle 5, cPrint.CurrentY, 0.25, vbBlue, , vbDownwardDiagonal, 1.5
    cPrint.pCircle 6, cPrint.CurrentY, 0.25, vbBlue, , vbDownwardDiagonal, 0.5
        
    cPrint.p15Space
    cPrint.pBox 1, 3, 1, 1, vbGreen
    cPrint.pBox 3, 3, 1, 1, vbGreen, QBColor(13), vbFSSolid
    cPrint.pBox 5, 3, 2, 1, vbGreen, QBColor(7), vbUpwardDiagonal
        
    cPrint.CurrentY = 3.25
    cPrint.pPrint " Box ", 1.35, True
    cPrint.BackColor = vbWhite
    cPrint.pPrint " Box ", 3.35, True
    cPrint.pPrint " Box ", 5.35
    cPrint.pPrint
    cPrint.BackColor = -1
    cPrint.pPrint " Box ", 1.35, True
    cPrint.pPrint " Box ", 3.35, True
    cPrint.pPrint " Box ", 5.35
    
    cPrint.CurrentY = 4.1
    cPrint.FontSize = 12
    cPrint.pPrint "An example of highlighted text:", 1
    cPrint.pPrint "This is ", 1, True
    cPrint.BackColor = vbYellow
    cPrint.ForeColor = vbRed
    cPrint.pPrint "My Son", , True
    cPrint.BackColor = -1
    cPrint.ForeColor = vbBlack
    cPrint.pPrint " who is 2.5 years old."
    
    cPrint.pPrintPicture LoadPicture(App.Path & "\myson.jpg"), 1, 4.5, 3
    cPrint.pPrintPicture LoadPicture(App.Path & "\myson.jpg"), 4, 6, 1
    
    memCurrentYt = cPrint.CurrentY
    cPrint.pMultiline "This text uses pMultiline to print text between two user specified margins.  In this case it starts printing at 3.75 inches from the left margin and is 2 inches wide.", 3.75, 5.75
    
    cPrint.CurrentY = memCurrentYt
    cPrint.pCenterMultiline "This text uses pCenterMultiline to print text between two user specified margins.  In this case it starts printing at 6 inches from the left margin and is 1.5 inches wide.", 6, 7.5, True
    
    cPrint.FontBold = True
    cPrint.FontSize = 16
    cPrint.pPrintRotate "Print Preview - Rotate String", 90
    
    cPrint.FontBold = False
    cPrint.CurrentY = 7.5
    cPrint.FontSize = 12
    cPrint.pPrintRotate "Print Preview - Rotate String", 45, 4.5
    
    picPrinting.Visible = False
    Screen.MousePointer = vbDefault
    
    cPrint.ReportTitle = Command1.Caption
    
    cPrint.pFooter
    cPrint.pEndDoc
    
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    
    Set cPrint = Nothing
    
End Sub


Private Sub Command2_Click()
  Dim i As Integer
    
    Set cPrint = New clsMultiPgPreview
    
    With frmPrinterSetUp
        .fraOrientation.Visible = False
        .cmdPrint.Visible = False
        .Show vbModal
    End With
    If QuitCommand Then
        Set cPrint = Nothing
        Exit Sub
    End If
    
    cPrint.Orientation = PageLandscape
    cPrint.SendToPrinter = False
    
    picPrinting.Visible = True
    MousePointer = vbHourglass
    DoEvents

SendToPrinter:
    cPrint.pStartDoc
    
    cPrint.FontSize = 11
    GoSub PrintHeader
    
    cPrint.pPrint "cPrint.Orientation = PageLandscape", 0.5
    cPrint.pPrint "cPrint.SendToPrinter = False", 0.5
    cPrint.pPrint
    
    Do
        i = i + 1
        cPrint.pRightJust CStr(i), 1, True
        cPrint.pPrint "This line is printed over and over again", 1.1, True
        cPrint.pPrint "At 5 Inches", 5, True
        cPrint.pPrint "At 7 Inches", 7, True
        cPrint.pRightJust "RightJust " & CStr(i), 9.5
        
        If cPrint.pEndOfPage Then
            If cPrint.GetPage = 4 Then Exit Do
            cPrint.pFooter
            cPrint.pNewPage
            GoSub PrintHeader
        End If
    Loop
    cPrint.pFooter
    
    picPrinting.Visible = False
    MousePointer = vbDefault
    
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    
Exit Sub
    
PrintHeader:
    cPrint.pHeader "EndOfPage Check Demo", "(print some junk that runs across several pages)"
Return
    
End Sub

Private Sub Command3_Click()
  Dim tString  As String
  
    tString = "This" & vbCrLf & "is" & vbLf & vbCr & "a" & vbCrLf & "Test"
    
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        Set cPrint = Nothing
        Exit Sub
    End If

    
SendToPrinter:
    picPrinting.Visible = True
    
    cPrint.pStartDoc
    cPrint.pHeader "Ignore CF and LF", , False
    cPrint.FontSize = 12
    cPrint.CurrentY = 1
    cPrint.pPrint "Print a text string with imbedded CRLF characters and still keep the user specified left margin.", 0.75
    cPrint.pPrint
    cPrint.pPrint "Example:", 1.5
    cPrint.pPrint
    cPrint.pPrint tString, 2
    cPrint.pPrint
    cPrint.pPrint
    cPrint.pPrint "Or strip the imbedded characters to make a single line.", 0.75
    cPrint.pPrint
    cPrint.pPrint "Example:", 1.5
    cPrint.pPrint
    cPrint.pPrint cPrint.GetRemoveCRLF(tString), 2
    
    picPrinting.Visible = False
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
End Sub

Private Sub Command4_Click()
  Dim RoomNumber As String
  Dim Title1 As String
  Dim Title2 As String
  Dim TechName As String
  Dim tString As String
  

    Set cPrint = New clsMultiPgPreview
    
    With frmPrinterSetUp
        .fraOrientation.Visible = False
        .Show vbModal
    End With
    If QuitCommand Then
        QuitCommand = False
        Exit Sub
    End If

    
SendToPrinter:
    cPrint.pStartDoc
    cPrint.FontName = "Arial"
    
    Title1 = "ADULT" & vbCrLf & "MEN 1"
    Title2 = "18 - 21 Years Old"
    RoomNumber = "230"
    TechName = "Teacher: Morgan Haueisen"
    Call PrintTag(True, RoomNumber, Title1, Title2, TechName, &H40C0&)
    
    Title1 = "PRESCHOOL" & vbCrLf & "BABIES"
    Title2 = "0 - 1 Year Old"
    RoomNumber = "120"
    TechName = "Teacher: Tammy Brown"
    Call PrintTag(False, RoomNumber, Title1, Title2, TechName)
    
    cPrint.pEndDoc
    
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing

End Sub

Private Sub PrintTag(PageTop As Boolean, _
            ByVal RoomNumber As String, _
            ByVal Title1 As String, _
            ByVal Title2 As String, _
            ByVal TeacherName As String, _
            Optional ByVal fColor As Long = vbBlack)
    
  Dim TopMargin As Single
  
    If PageTop Then
        TopMargin = 0
    Else
        TopMargin = cPrint.GetPaperHeight / 2
    End If
    
    
    cPrint.DrawWidth = 10
    cPrint.pBox 0.25, TopMargin + 0.25, cPrint.GetPaperWidth - 0.5, (cPrint.GetPaperHeight / 2) - 0.5, fColor
    cPrint.DrawWidth = 1
    
    cPrint.pBox 0.25, TopMargin + 0.25, cPrint.GetPaperWidth - 0.5, 1, fColor, fColor, vbFSSolid
    cPrint.FontBold = True
    cPrint.ForeColor = vbWhite
    cPrint.BackColor = fColor
    cPrint.FontSize = 58
    cPrint.CurrentY = TopMargin + 0.3
    cPrint.pCenter RoomNumber
    
    cPrint.BackColor = vbWhite
    cPrint.ForeColor = fColor
    cPrint.CurrentY = TopMargin + 1.5
    cPrint.pCenterMultiline Title1, 0.5, cPrint.GetPaperWidth - 0.5, False
    
    cPrint.FontSize = 36
    cPrint.pCenter Title2
    
    cPrint.FontBold = False
    cPrint.FontItalic = True
    cPrint.pPrint
    cPrint.FontSize = 18
    cPrint.CurrentY = TopMargin + (cPrint.GetPaperHeight / 2) - 1
    cPrint.pCenterMultiline TeacherName, 0.5, cPrint.GetPaperWidth - 0.5, False
    cPrint.FontItalic = False
    

End Sub

Private Sub Command5_Click()
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter:
    Screen.MousePointer = vbHourglass
    picPrinting.Visible = True
    DoEvents
    
    cPrint.pStartDoc
    cPrint.pBox 0, 0
    '/* Section 1
    cPrint.pBox cPrint.GetPaperWidth - 1.25, 0.5, , 1.155
    cPrint.CurrentY = 1.125
    cPrint.pLine cPrint.GetPaperWidth - 1.25, , , False
    cPrint.pBox 0, 0, , 1.655
    cPrint.pBox 0, 0, , 1.675
    
    cPrint.FontName = "Arial"
    cPrint.FontSize = 8
    cPrint.CurrentY = 0.53
    cPrint.pPrint "HOUSE ORDER NO.", cPrint.GetPaperWidth - 1.2
    cPrint.CurrentY = 1.13
    cPrint.pPrint "DATE:", cPrint.GetPaperWidth - 1.2
    
    cPrint.pPrintPicture LoadPicture(App.Path & "\JNJ.jpg"), 0.3, 0.25, , 1.25
    
    cPrint.FontName = "Impact"
    cPrint.FontSize = 18
    cPrint.CurrentY = 0.8
    cPrint.pPrint "R", cPrint.GetPaperWidth - 1.2
    cPrint.FontSize = 22
    cPrint.pPrint "HOUSE ORDER", 0.5, True
    cPrint.FontName = "Arial"
    cPrint.FontSize = 8
    cPrint.p15Space
    cPrint.pRightTab "(WORK IN PROCESS, NON-STANDARD GOODS OR RAW MATERIALS)", 1.5
    
    cPrint.CurrentY = 0.6
    cPrint.pPrintPicture LoadPicture(App.Path & "\arrow.bmp"), 6.2, 0.6, 0.5, 0.4, , False
    cPrint.pMultiline "IN CASE OF INQUIRY, QUOTE THIS NUMBER", 5.25, 6.25
    
    '/* Section 2
    cPrint.pVerticalLine 0.375, 1.675, 3.53
    cPrint.pVerticalLine 3.375, 1.675, 3.53
    cPrint.pVerticalLine 3.75, 1.675, 3.53
    cPrint.pVerticalLine 6.75, 1.675, 3.53
    
    cPrint.CurrentY = 2.138
    cPrint.pLine 0.375, 3.375, , False
    cPrint.CurrentY = 2.602
    cPrint.pLine 0.375, 3.375, , False
    cPrint.CurrentY = 3.066
    cPrint.pLine 0.375, 3.375, , False
    
    cPrint.CurrentY = 2.138
    cPrint.pLine 3.75, 6.75, , False
    cPrint.CurrentY = 2.602
    cPrint.pLine 3.75, 6.75, , False
    cPrint.CurrentY = 3.066
    cPrint.pLine 3.75, 6.75, , False
    
    cPrint.CurrentY = 3.5
    cPrint.pDoubleLine
    
    cPrint.CurrentY = 1.7
    cPrint.pPrint "NAME", 0.43
    cPrint.CurrentY = 2.16
    cPrint.pPrint "STREET", 0.43
    cPrint.CurrentY = 2.627
    cPrint.pPrint "CITY-STATE-ZIP CODE", 0.43
    cPrint.CurrentY = 3.09
    cPrint.pPrint "CUSTOMER ORDER NO.", 0.43
    
    cPrint.CurrentY = 1.7
    cPrint.pPrint "NAME", 3.8
    cPrint.CurrentY = 2.16
    cPrint.pPrint "STREET", 3.8
    cPrint.CurrentY = 2.627
    cPrint.pPrint "CITY-STATE-ZIP CODE", 3.8
    cPrint.CurrentY = 3.09
    cPrint.pPrint "ATTENTION", 3.8
    
    cPrint.CurrentY = 1.7
    cPrint.FontSize = 7
    cPrint.pPrint "FOR OFFICAL USE ONLY", 6.8
    cPrint.pLine 6.75
    
    cPrint.CurrentY = 1.8
    cPrint.FontSize = 10
    cPrint.pPrint "S", 0.12
    cPrint.pPrint "O", 0.12
    cPrint.pPrint "L", 0.12
    cPrint.pPrint "D", 0.12
    cPrint.pPrint
    cPrint.pPrint
    cPrint.pPrint "T", 0.12
    cPrint.pPrint "O", 0.12
    
    cPrint.CurrentY = 1.8
    cPrint.FontSize = 10
    cPrint.pPrint "S", 3.52
    cPrint.pPrint "H", 3.52
    cPrint.pPrint "I", 3.52
    cPrint.pPrint "P", 3.52
    cPrint.pPrint
    cPrint.pPrint
    cPrint.pPrint "T", 3.52
    cPrint.pPrint "O", 3.52
    
    
    
    '/* Section 3
    cPrint.CurrentY = 4.25
    cPrint.pLine
    cPrint.CurrentY = 4.75
    cPrint.pLine
    cPrint.CurrentY = 5.125
    cPrint.pLine
    cPrint.CurrentY = 5.5
    cPrint.pLine
    cPrint.CurrentY = 5.875
    cPrint.pLine
    
    cPrint.pVerticalLine 1, 3.56, 6.24
    cPrint.CurrentY = 3.75
    cPrint.FontSize = 8
    cPrint.pPrint "BILLING", 0.1
    cPrint.pPrint "INSTRUCTIONS", 0.1
    cPrint.CurrentY = 4.4
    cPrint.pPrint "OUTBOUND", 0.1
    cPrint.pPrint "FREIGHT TERMS", 0.1
    cPrint.CurrentY = 4.825
    cPrint.pPrint "SHIPPING", 0.1
    cPrint.pPrint "INSTRUCTIONS", 0.1
    cPrint.CurrentY = 5.2
    cPrint.pPrint "SPECIAL", 0.1
    cPrint.pPrint "INSTRUCTIONS", 0.1
    cPrint.CurrentY = 5.58
    cPrint.pPrint "REASON FOR", 0.1
    cPrint.pPrint "SHIPMENT", 0.1
    cPrint.CurrentY = 5.95
    cPrint.pPrint "POINT OF", 0.1
    cPrint.pPrint "SHIPMENT", 0.1
    
    cPrint.pVerticalLine 1.25, 3.7, 4.79
    cPrint.pVerticalLine 2, 4.28, 4.79
    cPrint.pVerticalLine 2.5, 4.28, 4.79
    cPrint.pVerticalLine 3, 3.56, 4.79
    cPrint.pVerticalLine 3.25, 3.8, 4.79
    
    
    cPrint.CurrentY = 3.67
    cPrint.pLine 1, 1.25
    cPrint.CurrentY = 3.95
    cPrint.pLine 1, 1.25
    
    cPrint.CurrentY = 3.77
    cPrint.pLine 3, 3.25
    cPrint.CurrentY = 4.03
    cPrint.pLine 3, 3.25
    cPrint.CurrentY = 4.5
    cPrint.pLine 3, 3.25
    
    cPrint.pVerticalLine 4.2, 3.56, 4.5
    cPrint.pBox 4.2, 4.28, 0.25, 0.25
    
    cPrint.CurrentY = 3.77
    cPrint.pPrint "NO CHARGES", 1.3
    cPrint.CurrentY = 3.85
    cPrint.pPrint "INVOICE", 3.3
    
    cPrint.CurrentY = 4
    cPrint.pPrint "DEBIT VENDOR FOR", 1.3
    cPrint.pPrint "INBOUND FREIGHT CHGS.$", 1.3
    cPrint.CurrentY = 4.1
    cPrint.pPrint "PREPAY", 3.3
    
    cPrint.CurrentY = 4.3
    cPrint.pPrint "No. Cases", 1.3, True
    cPrint.pPrint "Weight", 2.1, True
    cPrint.pPrint "Charges", 2.55, True
    cPrint.CurrentY = 4.35
    cPrint.pPrint "COLLECT", 3.3
    cPrint.CurrentY = 4.6
    cPrint.pPrint "PREPAY", 3.3
    
    cPrint.CurrentY = 4.3
    cPrint.pMultiline "Prepay & Add Any Charge Freight To", 4.5, 5.25
    cPrint.CurrentY = 4.5
    cPrint.FontBold = True
    cPrint.pPrint "138P", 5.3
    cPrint.FontBold = False
    
    cPrint.pVerticalLine 5.25, 3.56, 4.79
    cPrint.pVerticalLine 5.625, 3.56, 4.79
    cPrint.pVerticalLine 6.75, 3.56, 4.79
    cPrint.pVerticalLine 7.375, 3.56, 4.79
    
    cPrint.CurrentY = 3.6875
    cPrint.pLine 4.2
    cPrint.CurrentY = 3.98
    cPrint.pLine 4.2
    
    cPrint.CurrentY = 3.58
    cPrint.pPrint "P/C", 5.3, True
    cPrint.pPrint "DEPT.", 5.7, True
    cPrint.pPrint "ACCT", 6.8, True
    cPrint.pPrint "REF", 7.4
    
    cPrint.CurrentY = 3.73
    cPrint.pPrint "DEBIT", 4.25
    cPrint.pPrint "ACCOUNT NO.", 4.25
    
    cPrint.CurrentY = 4.01
    cPrint.pPrint "CREDIT", 4.25
    cPrint.pPrint "ACCOUNT NO.", 4.25
    
    
    cPrint.pBox 1, 4.84, 0.25, 0.25
    cPrint.pBox 3, 4.84, 0.25, 0.25
    cPrint.pBox 5, 4.84, 0.25, 0.25
    cPrint.CurrentY = 4.9
    cPrint.pPrint "AT ONCE", 1.3, True
    cPrint.pPrint "ON OR BEFORE", 3.3, True
    cPrint.pPrint "ON OR AFTER", 5.3
    
    
    
    
    cPrint.CurrentY = 6.2
    cPrint.pDoubleLine
    
    
    '/* Section 4
    cPrint.CurrentY = 9.625
    cPrint.pDoubleLine
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    picPrinting.Visible = False
    Screen.MousePointer = vbDefault
    
    cPrint.pEndDoc
    
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    
    Set cPrint = Nothing

End Sub

Private Sub Command6_Click()
  Dim i As Integer
    
    Set cPrint = New clsMultiPgPreview
    
    With frmPrinterSetUp
        .fraOrientation.Visible = False
        .cmdPrint.Visible = False
        .Show vbModal
    End With
    If QuitCommand Then
        Set cPrint = Nothing
        Exit Sub
    End If
    
    cPrint.Orientation = PageLandscape
    cPrint.SendToPrinter = False
    
    picPrinting.Visible = True
    MousePointer = vbHourglass
    DoEvents

SendToPrinter:
    cPrint.pStartDoc
    
    cPrint.FontSize = 11
    GoSub PrintHeader
    
    cPrint.pPrint "cPrint.Orientation = PageLandscape", 0.5
    cPrint.pPrint "cPrint.SendToPrinter = False", 0.5
    cPrint.FontBold = True
    cPrint.pPrint "The following text is in the Multi-Line Textbox (Text1)", 0.5
    cPrint.FontBold = False
    cPrint.pPrint
    
    cPrint.pMultiline Text1.Text, 2, 6, , , True
    
    cPrint.pFooter
    
    picPrinting.Visible = False
    MousePointer = vbDefault
    
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    
Exit Sub
    
PrintHeader:
    cPrint.pHeader "EndOfPage Check Demo", "(print a textbox that runs across several pages)"
Return

End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    picPrinting.Move Command1.Left - 500, Command1.Top - 100
    ChDir App.Path
End Sub


