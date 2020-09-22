Attribute VB_Name = "modArySupport"
Option Explicit

'*******************************************************************************
' Function Name     : BuildMDAry
' Purpose           : Build a square multidimensional Array above a base variable
'                   : Dims() is an array of Dim Sizes (count=size+1)
'*******************************************************************************
Public Function BuildMDAry(ByVal VarNum As Long, ByVal DimX As Long, _
                           ByVal DimY As Long, ByVal Redimn As Boolean) As Boolean
  Dim Vptr As clsVarSto, tPtr As clsVarSto
  Dim Idx As Long
  
  If VarNum < 0 Or VarNum > MaxVar Then Exit Function   'ensure variable number is valid
  If DimX < 0 Or DimX > 99 Then Exit Function           'if DimX (main Dim) is not valid
  Set Vptr = Variables(VarNum).Vdata                    'point to root variable object
  Set tPtr = Vptr.BuildAry(DimX, Redimn)                'build initial dim (x value), and
                                                        'return base variable for new list
  If tPtr Is Nothing Then Exit Function                 'oops
'
' if DimY is defined, build 2nd dim from first
'
  If DimY >= 0 And DimY < 100 Then
    For Idx = 0 To DimX
      Call Vptr.PntToLnk(Idx).AddAry(DimY, Redimn)      'add an array to each added X dim variable
    Next Idx
  End If
  
  BuildMDAry = True                                     'ended in success
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

