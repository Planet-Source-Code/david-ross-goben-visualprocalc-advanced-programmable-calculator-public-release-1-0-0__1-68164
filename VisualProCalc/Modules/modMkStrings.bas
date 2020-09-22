Attribute VB_Name = "modMkStrings"
Option Explicit
'~modMkStrings.bas;
'Copy binary numeric data to and from strings
'****************************************************************************
' modMkStrings: Copy numeric data to and from strings, specifying an offset within
'               a string where the data is located. Offset is zero-based, not
'               1-based as a VB string is, so position 1 in a string is position 0
'               in the offset. Useful for reading/writing blocks of data in files.
'
' The following functions are provided:
' StrMkInt(): convert binary data from string to integer(returns 2 byte int)
' StrMkLng(): convert binary data from string to long   (returns 4 byte lng)
' StrMkSng(): convert binary data from string to single (returns 4 byte sng)
' StrMkDbl(): convert binary data from string to double (returns 8 byte dbl)
'
' The following subroutines are provided:
' IntMkStr(): convert integer to binary string  (write 2 byte int to string)
' LngMkStr(): convert long to binary string     (write 4 byte lng to string)
' SngMkStr(): convert single to binary string   (write 4 byte sng to string)
' DblMkStr(): convert double to binary string   (write 8 byte dbl to string)
'****************************************************************************

'David Goben 2000
'*****************************************************
' API call used to copy memory to/from string/variable
'*****************************************************
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'*************************************************
' StrMkInt(): convert binary data from string to integer. Returns Int
'*************************************************
Public Function StrMkInt(ByVal Txt As String, ByVal Posn As Long) As Integer
  Dim ptr As String   'copy of source string for data to convert to integer
  Dim dst As Integer  'temp integer to receive 2 string bytes
  
  ptr = Mid$(Txt, Posn + 1, 2)        'grab character to convert
  Call CopyMemory(dst, ByVal ptr, 2&) 'copy 2 bytes
  StrMkInt = dst                      'return value to user
End Function
  
'*************************************************
' StrMkLng(): convert binary data from string to long. Returns Lng
'*************************************************
Public Function StrMkLng(ByVal Txt As String, ByVal Posn As Long) As Long
  Dim ptr As String       'copy of source string for data to convert to long
  Dim dst As Long         'temp long to receive 4 string bytes
  
  ptr = Mid$(Txt, Posn + 1, 4)            'grab character to convert
  Call CopyMemory(dst, ByVal ptr, 4&)     'copy 4 bytes
  StrMkLng = dst                          'return value to user
End Function
  
'*************************************************
' StrMkSng(): convert binary data from string to single. Returns Sng
'*************************************************
Public Function StrMkSng(ByVal Txt As String, ByVal Posn As Long) As Single
  Dim ptr As String       'copy of source string for data to convert to single
  Dim dst As Single       'temp single to receive 4 string bytes
  
  ptr = Mid$(Txt, Posn + 1, 4)            'grab character to convert
  Call CopyMemory(dst, ByVal ptr, 4&)     'copy 4 bytes
  StrMkSng = dst                          'return value to user
End Function
  
'*************************************************
' StrMkDbl(): convert binary data from string to double. Returns Dbl
'*************************************************
Public Function StrMkDbl(ByVal Txt As String, ByVal Posn As Long) As Double
  Dim ptr As String       'copy of source string for data to convert to double
  Dim dst As Double       'temp double to receive 8 string bytes
  
  ptr = Mid$(Txt, Posn + 1, 8)            'grab character to convert
  Call CopyMemory(dst, ByVal ptr, 8&)     'copy 8 bytes
  StrMkDbl = dst                          'return value to user
End Function
  
'*************************************************
' IntMkStr(): convert integer to binary string. Writes from Source to txt
'*************************************************
Public Sub IntMkStr(ByRef Source As Integer, Txt As String, ByVal Posn As Long)
  Dim ptr As String * 2   'temp string to receive intial bytes
  
  Call CopyMemory(ByVal ptr, Source, 2&)  'copy integer to temp string
  Mid$(Txt, Posn + 1, 2) = ptr            'stuff new string to master
End Sub
  
'*************************************************
' LngMkStr(): convert long to binary string. Writes from Source to txt
'*************************************************
Public Sub LngMkStr(ByRef Source As Long, Txt As String, ByVal Posn As Long)
  Dim ptr As String * 4   'temp string to receive intial bytes
  
  Call CopyMemory(ByVal ptr, Source, 4&)  'copy long to temp string
  Mid$(Txt, Posn + 1, 4) = ptr            'stuff new string to master
End Sub
  
'*************************************************
' SngMkStr(): convert single to binary string. Writes from Source to txt
'*************************************************
Public Sub SngMkStr(ByRef Source As Single, Txt As String, ByVal Posn As Long)
  Dim ptr As String * 4   'temp string to receive intial bytes
  
  Call CopyMemory(ByVal ptr, Source, 4&)  'copy single to temp string
  Mid$(Txt, Posn + 1, 4) = ptr            'stuff new string to master
End Sub
  
'*************************************************
' DblMkStr(): convert double to binary string. Writes from Source to txt
'*************************************************
Public Sub DblMkStr(ByRef Source As Double, Txt As String, ByVal Posn As Long)
  Dim ptr As String * 8   'temp string to receive intial bytes
  
  Call CopyMemory(ByVal ptr, Source, 8&)  'copy double to temp string
  Mid$(Txt, Posn + 1, 8) = ptr            'stuff new string to master
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

