Attribute VB_Name = "modIsDimmed"
Option Explicit
'~modIsDimmed.bas;
'Determine if an array is dimensioned
'********************************************************************************
' modIsDimmed - The IsDimmed() function returns True if the specified array is
'               dimensioned.
'EXAMPLE:
' Dim Test() As String
'
' Print IsDimmed(Test)   'prints False
' ReDim Test(5 To 6)
' Print IsDimmed(Test)   'prints True
' ReDim Test(0 To 5)
' Print IsDimmed(Test)   'prints True
' Erase Test
' Print IsDimmed(Test)   'Prints False
'********************************************************************************
Public Function IsDimmed(vArray As Variant) As Boolean
  On Error Resume Next
  IsDimmed = IsNumeric(UBound(vArray))
  On Error GoTo 0
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

