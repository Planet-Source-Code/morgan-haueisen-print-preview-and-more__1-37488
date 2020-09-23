VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGdiPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/* Class for GDI+ Access
'/* Requires gdi+.tlb

'/* Note: if you are unfamiliar with tlb (type library)
'/* They are files containing Declares, Enums and Constants (and can also contain interfaces and other data structures)
'/* Type Libraries are compiled into the exe, and do NOT need to be available externally.

Event Error(ByVal lGdiError As Long, ByVal sErrorDesc As String)

Dim m_lToken        As Long         '/* Startup/Shutdown token
Dim tGuids()        As GDIPlus.GUID '/* array of GUIDs for codecs
Dim colCodecs       As Collection   '/* search into codec index on file extension
Dim m_lNumCodecs    As Long
'

Private Sub Class_Initialize()
  Dim gsi       As GdiplusStartupInput
  Dim lError    As Long
  
  gsi.GdiplusVersion = 1
  lError = GdiplusStartup(m_lToken, gsi, ByVal 0)
  
  If Not lError Then
    EnumEncoders
  Else
    RaiseEvent Error(lError, GdiErrorString(lError))
  End If
End Sub

Private Sub Class_Terminate()
  If m_lToken Then GdiplusShutdown m_lToken
End Sub

Private Function EnumEncoders() As Long
  Dim lNumEncoders  As Long
  Dim lEncoderSize  As Long
  Dim lError        As Long
  Dim b()           As Byte
  Dim i             As Long
  Dim codecs()      As ImageCodecInfo
  
  lError = GdipGetImageEncodersSize(lNumEncoders, lEncoderSize)
  If Not lError Then
    ReDim codecs(lNumEncoders - 1)
    ReDim b(lEncoderSize - 1)
    
    lError = GdipGetImageEncoders(lNumEncoders, lEncoderSize, b(0))
    If Not lError Then
      RtlMoveMemory codecs(0), b(0), lNumEncoders * LenB(codecs(0))
      ReDim tGuids(lNumEncoders - 1)
      m_lNumCodecs = lNumEncoders
      Set colCodecs = Nothing
      Set colCodecs = New Collection
      
      Do While lNumEncoders
        lNumEncoders = lNumEncoders - 1
        tGuids(lNumEncoders) = codecs(lNumEncoders).Clsid
        ParseOnChar StringFromPointerW(codecs(lNumEncoders).pwszFilenameExtension), ";", lNumEncoders
      Loop
    Else
      RaiseEvent Error(lError, GdiErrorString(lError))
    End If
  Else
    RaiseEvent Error(lError, GdiErrorString(lError))
  End If
End Function

'/* included instead of using Split() for the VB5 set :)
Private Sub ParseOnChar(ByRef sIn As String, ByRef sChar As String, ByVal lGuidIndex As Long)
  Dim lStartPosition As Long
  Dim lFoundPosition As Long
  Dim sItem          As String
  
  lFoundPosition = InStr(sIn, sChar)
  lStartPosition = 1
  
  Do While lFoundPosition
    sItem = Mid$(sIn, lStartPosition, lFoundPosition - lStartPosition)
    colCodecs.Add lGuidIndex, sItem
    lStartPosition = lFoundPosition + 1
    lFoundPosition = InStr(lStartPosition, sIn, sChar)
  Loop
  
  sItem = Trim$(Mid$(sIn, lStartPosition))
  If LenB(sItem) Then colCodecs.Add lGuidIndex, sItem
End Sub

'/* do not compare this to a boolean
'/*  returns -1 for not found, 0-positive GUID index for found
Private Function ExtensionExists(ByRef sKey As String) As Long
  On Error GoTo errorhandler
  ExtensionExists = True '/* invalid index
  
  If Not colCodecs Is Nothing Then
    ExtensionExists = colCodecs.Item(sKey)
  End If
  
  Exit Function
errorhandler:
  '/* exit silently
End Function

Private Function StringToGuid(ByRef sGuid As String) As GDIPlus.GUID
  CLSIDFromString sGuid, StringToGuid
End Function

'/* saves the contents of a picturebox to a file
'/* supports GIF/JPG/TIF/PNG and various others
Public Function PictureBoxToFile(ByVal pic As PictureBox, ByRef sFilename As String, Optional lQuality As Long = 85) As Long
  Dim sExtension As String
  Dim bitmap     As Long
  Dim lError     As Long
  Dim params     As EncoderParameters
  Dim lQual      As Long
  Dim lIndex     As Long
  Dim tguid      As GDIPlus.GUID
  
  sExtension = GetExtension(sFilename)
  
  lIndex = ExtensionExists("*." & sExtension)
  If lIndex > -1 Then
    lError = GdipCreateBitmapFromHBITMAP(pic.Picture.Handle, pic.Picture.hPal, bitmap)
    If Not lError Then
      If (Asc(sExtension) And Not 32) = vbKeyJ Then '/* does the file extention begin with j
        lQual = lQuality
        params.Count = 1
        params.Parameter.GUID = StringToGuid(EncoderQuality)
        params.Parameter.NumberOfValues = 1
        params.Parameter.Type = EncoderParameterValueTypeLong
        params.Parameter.Value = VarPtr(lQual)
        lError = GdipSaveImageToFile(bitmap, sFilename, tGuids(lIndex), params)
      Else '/* Save as 256 color gif
        lError = GdipSaveImageToFile(bitmap, sFilename, tGuids(lIndex), ByVal 0)
      End If
      
      If Not lError Then
        PictureBoxToFile = True
      Else
        RaiseEvent Error(lError, GdiErrorString(lError))
      End If
    Else
      RaiseEvent Error(lError, GdiErrorString(lError))
    End If
  End If
End Function

Private Function GetExtension(ByVal sFile As String) As String
  Dim i As Long
    i = InStrRev(sFile, ".") + 1
    GetExtension = Mid$(sFile, i)
End Function

Private Function StringFromPointerW(ByVal lPointer As Long) As String
  Dim lLength As Long
  
  If lPointer Then
    lLength = lstrlenW(lPointer)
    StringFromPointerW = Space$(lLength)
    RtlMoveMemory ByVal StrPtr(StringFromPointerW), ByVal lPointer, lLength * 2
  End If
End Function

Public Function GdiErrorString(ByVal lError As Status) As String
  Dim s As String
  
  Select Case lError
    Case GenericError:              s = "Generic Error"
    Case InvalidParameter:          s = "Invalid Parameter"
    Case OutOfMemory:               s = "Out Of Memory"
    Case ObjectBusy:                s = "Object Busy"
    Case InsufficientBuffer:        s = "Insufficient Buffer"
    Case NotImplemented:            s = "Not Implemented"
    Case Win32Error:                s = "Win32 Error"
    Case WrongState:                s = "Wrong State"
    Case Aborted:                   s = "Aborted"
    Case FileNotFound:              s = "File Not Found"
    Case ValueOverflow:             s = "Value Overflow"
    Case AccessDenied:              s = "Access Denied"
    Case UnknownImageFormat:        s = "Unknown Image Format"
    Case FontFamilyNotFound:        s = "FontFamily Not Found"
    Case FontStyleNotFound:         s = "FontStyle Not Found"
    Case NotTrueTypeFont:           s = "Not TrueType Font"
    Case UnsupportedGdiplusVersion: s = "Unsupported Gdiplus Version"
    Case GdiplusNotInitialized:     s = "Gdiplus Not Initialized"
    Case PropertyNotFound:          s = "Property Not Found"
    Case PropertyNotSupported:      s = "Property Not Supported"
    Case Else:                      s = "Unknown GDI+ Error"
  End Select
  
  GdiErrorString = s
End Function
