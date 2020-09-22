Attribute VB_Name = "modPrintTextAtAngle"
Option Explicit
'~modPrintTextAtAngle.bas
'Set text roation angle
'******************************************************************************
' modPrintTextAtAngle - The PrintTextAtAngle() subroutine will send a line of
'                       text out to an object that contains a DC (Device Context).
'                       Normally this is a form, a picturebox, or a printer. The
'                       optional Angle parameter allows you to print text at an
'                       angle in degrees, where 0 is normal, 90 from bottom to top,
'                       180 from right to left, upside down, etc.
'EXAMPLE:
'  With Printer
'    .CurrentX = .ScaleWidth / 2   'start at center of printer page
'    .CurrentY = .ScaleHeight / 2
'  End With
'  '
'  ' print some text at a 45-degree angle
'  '
'  PrintTextAtAngle Printer, "This is rotated Text", 45
'******************************************************************************

'------------------------------------------------------------------------------
' API constants, types, and declarations
'------------------------------------------------------------------------------
Private Const LF_FACESIZE = 32
Private Const ANTIALIASED_QUALITY As Long = 4
'
' logical font
'
Private Type LOGFONT
  lfHeight            As Long
  lfWidth             As Long
  lfEscapement        As Long
  lfOrientation       As Long
  lfWeight            As Long
  lfItalic            As Byte
  lfUnderline         As Byte
  lfStrikeOut         As Byte
  lfCharSet           As Byte
  lfOutPrecision      As Byte
  lfClipPrecision     As Byte
  lfQuality           As Byte
  lfPitchAndFamily    As Byte
  lfFacename          As String * LF_FACESIZE
End Type

Private Type TEXTMETRIC
  tmHeight            As Long
  tmAscent            As Long
  tmDescent           As Long
  tmInternalLeading   As Long
  tmExternalLeading   As Long
  tmAveCharWidth      As Long
  tmMaxCharWidth      As Long
  tmWeight            As Long
  tmOverhang          As Long
  tmDigitizedAspectX  As Long
  tmDigitizedAspectY  As Long
  tmFirstChar         As Byte
  tmLastChar          As Byte
  tmDefaultChar       As Byte
  tmBreakChar         As Byte
  tmItalic            As Byte
  tmUnderlined        As Byte
  tmStruckOut         As Byte
  tmPitchAndFamily    As Byte
  tmCharSet           As Byte
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long

'*******************************************************************************
' Subroutine Name   : PrintTextAtAngle
' Purpose           : Output text to Object having an HDC at a specified angle
'                   ' Angle is in Degrees
'*******************************************************************************
Public Sub PrintTextAtAngle(Obj As Object, Txt As String, Optional Angle As Single = 0!)
  Dim DIAngle As Long, oldHDC As Long
  
  DIAngle = CLng(Angle * 10!)                       'compute angle in 10th of a degree
  oldHDC = SelectObject(Obj.hdc, CreateFontIndirect(CreateRotatedFont(Obj, DIAngle)))
  Obj.Print Txt;
  Call DeleteObject(SelectObject(Obj.hdc, oldHDC))  'reset old font, release new
End Sub

'*******************************************************************************
' Function Name     : CreateRotatedFont
' Purpose           : Define new Font with angle in degrees*10, anti-clockwise
'*******************************************************************************
Private Function CreateRotatedFont(Obj As Object, ByVal Angle As Long) As LOGFONT
  Dim tm As TEXTMETRIC
  
  Call GetTextMetrics(Obj.hdc, tm)                    'get text metrics
  With CreateRotatedFont
    .lfFacename = Obj.Font.Name + Chr$(0)             'null terminated font name
    .lfHeight = tm.tmHeight                           'average text character height
    .lfWidth = tm.tmAveCharWidth                      'average text character width
    .lfOrientation = Angle
    .lfEscapement = Angle                             'escapement
    .lfWeight = Obj.Font.Weight                       'bold
    .lfItalic = Obj.Font.Italic                       'italic
    .lfUnderline = Obj.Font.Underline                 'underline
    .lfStrikeOut = Obj.Font.Strikethrough
    .lfQuality = ANTIALIASED_QUALITY
  End With
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

