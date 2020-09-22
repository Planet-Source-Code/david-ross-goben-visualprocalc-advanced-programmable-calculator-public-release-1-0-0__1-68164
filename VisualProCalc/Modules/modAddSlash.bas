Attribute VB_Name = "modAddSlash"
Option Explicit
'~modAddSlash.bas;
'Add a terminating backslash to a drive/path if required. Also remove
'********************************************************************************
' modAddSlash: The following functions are provided:
'
' AddSlash():    Add a terminating backslash to a drive/path if required. This function
'                is useful for building paths, and the string you are working with may
'                or may not already have a backslash appended to it.
' RemoveSlash(): Remove any existing terminating backslash from a path.
'********************************************************************************

Public Function AddSlash(str As String) As String
  AddSlash = Trim$(str)
  If Right$(AddSlash, 1) <> "\" Then AddSlash = AddSlash & "\"
End Function

Public Function RemoveSlash(str As String) As String
  RemoveSlash = Trim$(str)
  If Right$(RemoveSlash, 1) = "\" Then RemoveSlash = Left$(RemoveSlash, Len(RemoveSlash) - 1)
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

