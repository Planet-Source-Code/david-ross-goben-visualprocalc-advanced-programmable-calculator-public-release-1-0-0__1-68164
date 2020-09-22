Attribute VB_Name = "modGetWinSysCurDir"
Option Explicit
'~modGetWinSysCurDir.bas;
'Get WINDOWS, SYSTEM, and CURRENT directory paths
'*************************************************
' modGetWinSysCurDir:
' get WINDOWS, SYSTEM, and CURRENT directory paths
'
' The following functions are provided:
'
' GetWindowsDir(): get WINDOWS directory
' GetSystemDir():  get SYSTEM directory
' GetCurrentDir(): get current directory
'
'*************************************************

'*************************************************
' common API function calls
'*************************************************
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'*************************************************
' GetWindowsDir(): get WINDOWS directory
'*************************************************
Public Function GetWindowsDir() As String
  Dim S As String, i As Integer
  
  S = String$(260, 0)                   'init dump location
  i = GetWindowsDirectory(S, 260&)      'get directory to blank string
  If i Then
    GetWindowsDir = Left$(S, i)         'set windows directory
  Else
    If InStr(1, S, vbNullChar) = 0 Then GetWindowsDir = S
  End If
End Function

'*************************************************
' GetSystemDir(): get SYSTEM directory
'*************************************************
Public Function GetSystemDir() As String
  Dim S As String, i As Integer
  
  S = String$(260, 0)                   'init dump location
  i = GetSystemDirectory(S, 260&)       'now get system directory
  If i Then
    GetSystemDir = Left$(S, i)          'get system directory
  Else
    If InStr(1, S, vbNullChar) = 0 Then GetSystemDir = S
  End If
End Function

'*************************************************
' GetCurrentDir(): get current directory
'*************************************************
Public Function GetCurrentDir() As String
  Dim S As String, i As Integer
  
  S = String$(260, 0)                   'init dump location
  i = GetCurrentDirectory(260&, S)      'now get current directory
  If i Then
    GetCurrentDir = Left$(S, i)         'get system directory
  Else
    If InStr(1, S, vbNullChar) = 0 Then GetCurrentDir = S
  End If
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

