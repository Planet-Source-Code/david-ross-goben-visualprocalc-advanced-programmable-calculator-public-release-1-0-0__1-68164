Attribute VB_Name = "modGetScreenRes"
Option Explicit
'~modGetScreenRes.bas;
'Enumerate Display sizes to a string array. Get current display setting
'******************************************************************************
' modGetScreenRes - The following functions are provided:
'  GetCurrentDisplaySize() This function returns the string display for the
'                          current display setting.
'  GetXYBits()             This function extracts the x, y, and bits per pixel count
'                          from a provided string.
'******************************************************************************
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const ENUM_CURRENT_SETTINGS As Long = -1
'
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
'
Private Type DEVMODE
  dmDeviceName        As String * CCHDEVICENAME
  dmSpecVersion       As Integer
  dmDriverVersion     As Integer
  dmSize              As Integer
  dmDriverExtra       As Integer
  dmFields            As Long
  dmOrientation       As Integer
  dmPaperSize         As Integer
  dmPaperLength       As Integer
  dmPaperWidth        As Integer
  dmScale             As Integer
  dmCopies            As Integer
  dmDefaultSource     As Integer
  dmPrintQuality      As Integer
  dmColor             As Integer
  dmDuplex            As Integer
  dmYResolution       As Integer
  dmTTOption          As Integer
  dmCollate           As Integer
  dmFormName          As String * CCHFORMNAME
  dmLogPizels         As Integer
  dmBitsPerPel        As Long
  dmPelsWidth         As Long
  dmPelsHeight        As Long
  dmDisplayFlags      As Long
  dmDisplayFrequency  As Long
  dmICMMethod         As Long
  dmICMIntent         As Long
  dmMediaType         As Long
  dmDitherType        As Long
  dmReserved1         As Long
  dmReserved2         As Long
End Type

'*******************************************************************************
' Function Name     : GetCurrentDisplaySize
' Purpose           : Grab string rep of current display setting
'*******************************************************************************
Public Function GetCurrentDisplaySize() As String
  Dim DM As DEVMODE
  
  On Error Resume Next
  Call EnumDisplaySettings(0&, ENUM_CURRENT_SETTINGS, DM)
  GetCurrentDisplaySize = GetResolutionInfo(DM)
End Function

'*******************************************************************************
' Subroutine Name   : GetXYBits
' Purpose           : Extract x, y, bits from string
'*******************************************************************************
Public Sub GetXYBits(ColorString As String, X As Integer, Y As Integer, bits As Integer)
  Dim S As String
  Dim i As Integer
  
  X = 0                                 'init to bad result
  Y = 0
  bits = 0
  S = Trim$(ColorString)                'grab new size
  If Len(S) Then
    i = InStr(1, S, ";")                'find x and y delimiter
    If i Then
      bits = CInt(Val(Mid$(S, i + 2)))  'get bits per pixel
      S = Left$(S, i - 1)               'grab just x and y
      i = InStr(1, S, "x")              'find separator
      If i Then
        X = CInt(Left$(S, i - 1))       'x
        Y = CInt(Mid$(S, i + 1))        'y
      Else
        bits = 0                        'was bad info
      End If
    End If
  End If
End Sub

'*******************************************************************************
' Function Name     : GetResolutionInfo
' Purpose           : Support routine
'*******************************************************************************
Private Function GetResolutionInfo(DM As DEVMODE) As String
  Dim S As String
  Dim i As Integer
  
  S = CStr(DM.dmPelsWidth) & "x" & CStr(DM.dmPelsHeight) & "; " & CStr(DM.dmBitsPerPel) & " bits"
  On Error Resume Next
  i = DM.dmBitsPerPel
  If Err.Number Then i = 0
  On Error GoTo 0
  Select Case i
    Case 4: S = S & ", 16 Colors"
    Case 8: S = S & ", 256 Colors"
    Case 15: S = S & ", 32767 High Color"
    Case 16: S = S & ", 65536 High Color"
    Case 24: S = S & ", 16777216 True Color"
    Case 32: S = S & ", 4294967296 True Color"
    Case Else
      With Screen
        S = CStr(.Width \ .TwipsPerPixelX) & "x" & CStr(.Height \ .TwipsPerPixelY) & "; 8 bits, 256 Colors"
      End With
  End Select
  GetResolutionInfo = S
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

