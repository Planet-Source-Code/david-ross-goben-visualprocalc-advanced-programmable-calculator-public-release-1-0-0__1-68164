Attribute VB_Name = "modGetSetVars"
Option Explicit

'*******************************************************************************
' Function Name     : ExtractValue
' Purpose           : Extract the proper data type from a variable object
'*******************************************************************************
Public Function ExtractValue(Vptr As clsVarSto) As Variant
  Dim Pool() As StructPool
  Dim Idx As Long, Ofst As Long, Sz As Long
  Dim Typ As Vtypes
  
  Typ = Variables(Vptr.VarRoot).VarType       'get storage typ
  
  If Vptr.VarRoot = MaxVar + 1 Then           'if we are really dealing with a structure
    If CBool(ActivePgm) Then
      Pool = ModStPl                          'module structure pool
    Else
      Pool = StructPl                         'main structure pool
    End If
    
    With Pool(Vptr.StPlIdx)                   'with structure parent...
      With .StItems(Vptr.StItmIdx)            'and selected item...
        Ofst = .siOfst                        'get data offset location in Structure buffer
        Sz = .siLen                           'get size of data in structure
      End With
      Select Case Typ
        Case vChar
          ExtractValue = CVar(Asc(Mid$(.StBuf, Ofst + 1, 1)))
        Case vString
          ExtractValue = CVar(RTrim$(Mid$(.StBuf, Ofst + 1, Sz)))
        Case vInteger
          ExtractValue = CVar(StrMkLng(.StBuf, Ofst))
        Case vNumber
          ExtractValue = CVar(StrMkDbl(.StBuf, Ofst))
      End Select
    End With
  Else
'
' else grab data normally from the variable object pointer
'
    Select Case Typ
      Case vNumber  'double
        ExtractValue = CVar(Vptr.VarNum)
      Case vInteger 'long
        ExtractValue = CVar(Vptr.VarInt)
      Case vChar    'byte
        ExtractValue = CVar(Vptr.VarChar)
      Case vString  'string
        ExtractValue = CVar(Vptr.VarStr)
    End Select
  End If
End Function

'*******************************************************************************
' Function Name     : GetVarValue
' Purpose           : Return the value stored in a base variable (0-99 only)
'                   : this does not process dimensioning of variables.
'                   : This is used by active keypad
'*******************************************************************************
Public Function GetVarValue(ByVal VarNum As Double) As Variant
  Dim TV As Double
  
  TV = Fix(VarNum)
  If TV < 0# Or TV > DMaxVar Then  'check for valid variable
    ForcError "Variable number is out of range"
  Else
    GetVarValue = ExtractValue(Variables(CInt(TV)).Vdata)
  End If
End Function

'*******************************************************************************
' Subroutine Name   : StuffValue
' Purpose           : Convert and stuff a variant to the proper type
'*******************************************************************************
Public Sub StuffValue(Vptr As clsVarSto, ByVal Value As Variant)
  Dim S As String
  Dim Pool() As StructPool
  Dim Ofst As Long, Sz As Long, iLng As Long
  Dim iDbl As Double
  Dim Typ As Vtypes

  Typ = Variables(Vptr.VarRoot).VarType       'get storage typ
  
  If Vptr.VarRoot = MaxVar + 1 Then           'if we are really dealing with a structure
    If CBool(ActivePgm) Then
      Pool = ModStPl                          'module structure pool
    Else
      Pool = StructPl                         'main structure pool
    End If
    
    With Pool(Vptr.StPlIdx)                   'with selected Structure...
      With .StItems(Vptr.StItmIdx)            'and selected member...
        Ofst = .siOfst                        'get data offset location
        Sz = .siLen                           'get size of data
      End With
      On Error Resume Next
      Select Case Typ
        Case vChar
          Mid$(.StBuf, Ofst + 1, 1) = Chr$(CByte(Value))
        Case vString
          S = CStr(Value)
          If Len(S) < Sz Then
            Mid$(.StBuf, Ofst + 1, Sz) = S & String$(Sz - Len(S), 32)  'pad right side
          Else
            Mid$(.StBuf, Ofst + 1, Sz) = S
          End If
        Case vInteger
          iLng = CLng(Value)
          If Err.Number = 0 Then Call LngMkStr(iLng, .StBuf, Ofst)
        Case vNumber
          iDbl = CDbl(Value)
          If Err.Number = 0 Then Call DblMkStr(iDbl, .StBuf, Ofst)
      End Select
    End With
  Else                                          'we are using a normal variable
    Sz = Variables(Vptr.VarRoot).VdataLen       'save legnth of data, in case string
    On Error Resume Next
    Select Case Typ
      Case vChar                                'byte
        Vptr.VarChar = CByte(Value)
      Case vString                              'string
        S = CStr(Value)                         'get data
        If CBool(Sz) Then                       'if fixed width...
          If Len(S) < Sz Then S = S & String$(Sz - Len(S), 32)  'pad right side
        End If
        Vptr.VarStr = S                         'stuff result
      Case vInteger                             'long
        Vptr.VarInt = CLng(Value)
      Case vNumber                              'double
        Vptr.VarNum = CDbl(Value)               'stuff value
    End Select
  End If
  
  If CBool(Err.Number) Then                     'check and report errors
    ForcError Err.Description
    Exit Sub
  End If
  On Error GoTo 0
End Sub

'*******************************************************************************
' Function Name     : SetVarValue
' Purpose           : Set a variable's value
'                   : this does not process dimensioning of variables
'                   : This is used by active keypad
'*******************************************************************************
Public Function SetVarValue(ByVal VarNum As Double, ByVal Value As Variant) As Boolean
  Dim Vptr As clsVarSto
  Dim TV As Double
  Dim TI As Long
  
  TV = Fix(VarNum)
  If TV < 0# Or TV > DMaxVar Then       'check for valid variable
    ForcError "Variable number is out of range"
    Exit Function
  End If
  TI = CLng(TV)                         'get long version of variable number
  
  Set Vptr = Variables(TI).Vdata        'point to base variable
  Call StuffValue(Vptr, Value)          'assign to base variable
  SetVarValue = True
End Function

'*******************************************************************************
' Function Name     : PntToVptr
' Purpose           : Point to a specified dimensioned X and optional Y location.
'                   : if X or Y are not used, they contain a value of -1. (If X
'                   : is not used, then Y is not used). If they contain values
'                   : other than -1, they are assumed to be valid offsets (0-99)
'*******************************************************************************
Public Function PntToVptr(ByVal VarNum As Long, ByVal Xd As Long, ByVal Yd As Long) As clsVarSto
  Dim Vptr As clsVarSto
  
  With Variables(VarNum)                            'with variable base...
    If Xd = -1 Then                                 'if X not defined, then simply return base var
      Set PntToVptr = .Vdata
      Exit Function
    End If
    
    Set Vptr = .Vdata                               'else point to child's possible base X dim reference
    If Vptr.LnkNext Is Nothing Then                 'X dim not defined?
      ForcError "1st Dimension not defined for this variable"
      Exit Function
    End If
    Set Vptr = Vptr.PntToLnk(Xd)                    'point to DimX element
    If Vptr Is Nothing Then
      ForcError "1st Dimension element [" & CStr(Xd) & "] is not defined"
      Exit Function
    End If
    
    If Yd <> -1 Then                                'if DimY is defined
      Set Vptr = Vptr.LnkChild                      'point to base of Y array
      If Vptr Is Nothing Then                       'Y dim not defined
        ForcError "2nd Dimension not defined for this variable"
        Exit Function
      End If
      Set Vptr = Vptr.PntToLnk(Yd)                  'point to Y element
      If Vptr Is Nothing Then
        ForcError "2nd Dimension element [" & CStr(Yd) & "] is not defined"
        Exit Function
      End If
    End If
  End With
  
  Set PntToVptr = Vptr                              'indicate sucess
End Function

'*******************************************************************************
' Subroutine Name   : ClearEmAll
' Purpose           : Non-destructive routine to clear dimensioned arrays
'*******************************************************************************
Public Sub ClearEmAll(Vptr As clsVarSto)
  If Vptr Is Nothing Then Exit Sub    'nothing to do
  With Vptr
    .Init                             'clear variable data
    Call ClearEmAll(.LnkNext)         'clear any/all links in chain
    Call ClearEmAll(.LnkChild)        'clear any/all child branches
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : ClearAllVariables
' Purpose           : This function clears all defined varaibles
'*******************************************************************************
Public Sub ClearAllVariables()
  Dim Idx As Long
  
  For Idx = 0 To MaxVar
    Call ClearEmAll(Variables(Idx).Vdata)
  Next Idx
End Sub

'*******************************************************************************
' Function Name     : ClearVariable
' Purpose           : Clear an individual variable
'*******************************************************************************
Public Function ClearVariable(ByVal VarNum As Double) As Boolean
  If VarNum < 0# Or VarNum > DMaxVar Then           'check for valid variable
    ForcError "Variable number is out of range"
    Exit Function
  End If
  Call ClearEmAll(Variables(CLng(VarNum)).Vdata)    'clear root and any dimensions
  ClearVariable = True                              'success
End Function

'*******************************************************************************
' Subroutine Name   : SzOfElmnt
' Purpose           : Count size of all elements in an array
'*******************************************************************************
Public Function SzOfElmnt(Vptr As clsVarSto, ByVal Count As Boolean) As Long
  Dim Size As Long

  If Vptr Is Nothing Then                       'if we do not exist...
    Size = 0                                    'then size is nothing
  Else
    '
    ' first get local variable size
    '
    If Count Then
      Size = 1                                  'we are counting elements, so count self
    Else
      Select Case Variables(Vptr.VarRoot).VarType
        Case vNumber
          Size = 8                              'size of number
        Case vInteger
          Size = 4                              'size of integer
        Case vChar
          Size = 1                              'size of character
        Case vString
          Size = Len(Vptr.VarStr)               'count length of actual data
      End Select
    End If
    '
    ' now recurse count through any child nodes
    '
    Size = Size + SzOfElmnt(Vptr.LnkChild, Count) 'count any/all child branches
    Size = Size + SzOfElmnt(Vptr.LnkNext, Count)  'count any/all links in chain
  End If
  SzOfElmnt = Size                              'return full count
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

