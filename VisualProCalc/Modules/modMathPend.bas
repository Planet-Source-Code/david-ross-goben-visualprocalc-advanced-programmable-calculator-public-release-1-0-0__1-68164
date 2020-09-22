Attribute VB_Name = "modMathPend"
Option Explicit

'*******************************************************************************
' Subroutine Name   : GetHirValue
' Purpose           : Aquire Heirarchy level for Math Operation
'*******************************************************************************
Private Function GetHirValue(ByVal MathOp As Long) As Long
  Dim Opn As Long
  
  Select Case MathOp
    Case iEqual                           ' =         'LOWEST
      Opn = 0
    Case iOr, iNor                        ' || Nor
      Opn = 1
    Case iAnd                             ' &&
      Opn = 2
    Case iOrB                             ' |
      Opn = 3
    Case iXorB                            ' ^
      Opn = 4
    Case iAndB                            ' &
      Opn = 5
    Case iEQ, iNEQ                        ' == !=
      Opn = 6
    Case iLT, iLE, iGT, iGE               ' < <= > >=
      Opn = 7
    Case iShL, iShR                       '<< >>
      Opn = 8
    Case iAdd, iMinus                     ' + -
      Opn = 9
    Case iMult, iDVD, iMod, iBkSlsh       ' x ÷ % \
      Opn = 10
    Case iPower, iRoot, iLogX             ' y^, Root, LogX
      Opn = 11
    Case iNot, iNotB                      ' ! ~
      Opn = 12
    Case iLparen, iRparen, iLbrkt, iRbrkt ' ( ) [ ]
      Opn = 0                         'although this is priority 13, using '=' priority forces inner () [] calcs
    Case Else                         'anything else is HIGHEST
      Opn = 14
  End Select
  GetHirValue = Opn                   'return priority level
End Function

'*******************************************************************************
' Subroutine Name   : PushImmed
' Purpose           : Push display register and current math operation on stack
'*******************************************************************************
Public Sub PushImmed(ByVal MathOp As Long)
  Dim Opn As Long
  
  PendIdx = PendIdx + 1                   'bump index into pending pool
  If PendIdx > PendSize Then              'bump size of storage pool if needed...
    PendSize = PendSize + PendInc
    ReDim Preserve PendValue(PendSize)    'save pending value
    ReDim Preserve PendOpn(PendSize)      'save operation code
    ReDim Preserve PendHir(PendSize)      'save pending priority level
  End If
  '
  ' save pending value, operation, and heirarchy level
  '
  PendValue(PendIdx) = DisplayReg         'save current display register value
  PendOpn(PendIdx) = MathOp               'save current math operation
  PendHir(PendIdx) = GetHirValue(MathOp)  'save its heirarchy level
End Sub

'*******************************************************************************
' Function Name     : IsPending
' Purpose           : Return TRUE if there is a pending operation
'*******************************************************************************
Public Function IsPending() As Boolean
  IsPending = CBool(PendIdx)
End Function

'*******************************************************************************
' Function Name     : IsPendHir
' Purpose           : Determine if pending operation has equal or greater priority
'                   : than current operation.
'*******************************************************************************
Public Function IsPendEG(ByVal MathOp As Long) As Boolean
  If IsPending() Then
    IsPendEG = PendHir(PendIdx) >= GetHirValue(MathOp)
  End If
End Function

'*******************************************************************************
' Subroutine Name   : Pend
' Purpose           : Perform Algabraic Operating System rules on calcs
'*******************************************************************************
Public Sub Pend(ByVal MathOp As Integer)
  Dim PVal As Double
  
  If MathOp <> iLparen Then                   'if not '('
    Do While IsPendEG(MathOp)                 'do while math priority of prev ops are equal or higher
      PVal = PendValue(PendIdx)               'get value from stack
      Select Case PendOpn(PendIdx)            'process math
        '-------------------------------------
        Case iAdd
          On Error Resume Next
          DisplayReg = PVal + DisplayReg                'add to immediate
          Call CheckError
          On Error GoTo 0
        '-------------------------------------
        Case iMinus
          On Error Resume Next
          DisplayReg = PVal - DisplayReg                'subtract immediate
          Call CheckError
          On Error GoTo 0
        '-------------------------------------
        Case iMult
          On Error Resume Next
          DisplayReg = PVal * DisplayReg                'multiple immediate
          Call CheckError
          On Error GoTo 0
        '-------------------------------------
        Case iDVD
          On Error Resume Next
          DisplayReg = PVal / DisplayReg                'divide by immediate
          Call CheckError
          On Error GoTo 0
        '-------------------------------------
        Case iPower
          On Error Resume Next
          DisplayReg = Exp(DisplayReg * Log(PVal))      'PVal to power of DisplayReg
          Call CheckError
          On Error GoTo 0
        '-------------------------------------
        Case iRoot
          On Error Resume Next
          DisplayReg = Exp(1# / DisplayReg * Log(PVal)) 'PVal to root DisplayReg
          Call CheckError
          On Error GoTo 0
        '-------------------------------------
        Case iBkSlsh
          On Error Resume Next
          DisplayReg = Fix(PVal / DisplayReg)           'keep whole portion
          Call CheckError
          On Error GoTo 0
        '-------------------------------------
        Case iMod
          On Error Resume Next
          DisplayReg = PVal Mod DisplayReg              'keep remainder
          Call CheckError
          On Error GoTo 0
        '-------------------------------------
        Case iLogX
          On Error Resume Next
          DisplayReg = Log(PVal) / Log(DisplayReg)      'get logaritm to base X
          Call CheckError
          On Error GoTo 0
        '-------------------------------------
        Case iAndB                                      'get binary AND
          On Error Resume Next
          DisplayReg = Fix(PVal) And Fix(DisplayReg)
          Call CheckError
          On Error GoTo 0
        '-------------------------------------
        Case iOrB                                       'Get Binary OR
          On Error Resume Next
          DisplayReg = Fix(PVal) Or Fix(DisplayReg)
          Call CheckError
          On Error GoTo 0
        '-------------------------------------
        Case iNotB                                      'Get Binary NOT
          On Error Resume Next
          DisplayReg = Not Fix(DisplayReg)
          Call CheckError
          On Error GoTo 0
        '-------------------------------------
        Case iXorB                                      'Get Binary XOR
          On Error Resume Next
          DisplayReg = Fix(PVal) Xor Fix(DisplayReg)
          Call CheckError
          On Error GoTo 0
        '-------------------------------------
        Case iDivEq '/=
          If CurrentVar = -1 Then
            ForcError "Current default variable is not defined"
            DisplayReg = PVal
          Else
            With CurrentVarObj
              On Error Resume Next
              Select Case Variables(CurrentVarObj.VarRoot).VarType
                Case vNumber
                  .VarNum = .VarNum / DisplayReg
                Case vInteger
                  .VarInt = .VarInt / CLng(DisplayReg)
                Case vChar
                  .VarChar = .VarChar / CByte(DisplayReg)
              End Select
              Call CheckError
              On Error GoTo 0
            End With
            DisplayReg = PVal
          End If
        '-------------------------------------
        Case iMulEq '*=
          If CurrentVar = -1 Then
            ForcError "Current default variable is not defined"
            DisplayReg = PVal
          Else
            With CurrentVarObj
              On Error Resume Next
              Select Case Variables(CurrentVarObj.VarRoot).VarType
                Case vNumber
                  .VarNum = .VarNum * DisplayReg
                Case vInteger
                  .VarInt = .VarInt * CLng(DisplayReg)
                Case vChar
                  .VarChar = .VarChar * CByte(DisplayReg)
              End Select
              Call CheckError
              On Error GoTo 0
            End With
            DisplayReg = PVal
          End If
        '-------------------------------------
        Case iSubEq '-=
          If CurrentVar = -1 Then
            ForcError "Current default variable is not defined"
            DisplayReg = PVal
          Else
            With CurrentVarObj
              On Error Resume Next
              Select Case Variables(CurrentVarObj.VarRoot).VarType
                Case vNumber
                  .VarNum = .VarNum - DisplayReg
                Case vInteger
                  .VarInt = .VarInt - CLng(DisplayReg)
                Case vChar
                  .VarChar = .VarChar - CByte(DisplayReg)
              End Select
              Call CheckError
              On Error GoTo 0
            End With
            DisplayReg = PVal
          End If
        '-------------------------------------
        Case iAddEq '+=
          If CurrentVar = -1 Then
            ForcError "Current default variable is not defined"
            DisplayReg = PVal
          Else
            With CurrentVarObj
              On Error Resume Next
              Select Case Variables(CurrentVarObj.VarRoot).VarType
                Case vNumber
                  .VarNum = .VarNum + DisplayReg
                Case vInteger
                  .VarInt = .VarInt + CLng(DisplayReg)
                Case vChar
                  .VarChar = .VarChar + CByte(DisplayReg)
              End Select
              Call CheckError
              On Error GoTo 0
            End With
            DisplayReg = PVal
          End If
        '-------------------------------------
        Case iNot                                       'Get Logical Not
          If CBool(DisplayReg) Then
            DisplayReg = 0#
          Else
            DisplayReg = 1#
          End If
        '-------------------------------------
        Case iOr                                        'get logical OR
          If CBool(PVal) Or CBool(DisplayReg) Then
            DisplayReg = 1#
          Else
            DisplayReg = 0#
          End If
        '-------------------------------------
        Case iNor                                       'get logical NOR
          If CBool(PVal) Or CBool(DisplayReg) Then
            DisplayReg = 0#
          Else
            DisplayReg = 1#
          End If
        '-------------------------------------
        Case iAnd                                       'get logical AND
          If CBool(PVal) And CBool(DisplayReg) Then
            DisplayReg = 1#
          Else
            DisplayReg = 0#
          End If
        '-------------------------------------
        Case iEQ                                        'get logical EQUAL
          If PVal = DisplayReg Then
            DisplayReg = 1#
          Else
            DisplayReg = 0#
          End If
        '-------------------------------------
        Case iNEQ                                       'get logical NOT EQUAL
          If PVal <> DisplayReg Then
            DisplayReg = 1#
          Else
            DisplayReg = 0#
          End If
        '-------------------------------------
        Case iLT                                        'get logical LESS THAN
          If PVal < DisplayReg Then
            DisplayReg = 1#
          Else
            DisplayReg = 0#
          End If
        '-------------------------------------
        Case iLE                                        'get logical LESS OR EQUAL
          If PVal <= DisplayReg Then
            DisplayReg = 1#
          Else
            DisplayReg = 0#
          End If
        '-------------------------------------
        Case iGT                                        'get logical GREATER THAN
          If PVal > DisplayReg Then
            DisplayReg = 1#
          Else
            DisplayReg = 0#
          End If
        '-------------------------------------
        Case iGE                                        'get logical GREATER OR EQUAL
          If PVal >= DisplayReg Then
            DisplayReg = 1#
          Else
            DisplayReg = 0#
          End If
        '-------------------------------------
        Case iLparen
          PendIdx = PendIdx - 1                         'drop pending index in stack
          Exit Do                                       'break out of loop (we need only what is between ( & )
      
      End Select
      PendIdx = PendIdx - 1                             'drop pending index in stack
    Loop
    DisplayText = False
    If ErrorFlag Then Exit Sub
    Call DisplayLine  'update the display value with the value in the Immediate Register
  End If
'
' done with math, or pending operations are of lower priority
'
  Select Case MathOp
    Case iRparen, iEqual                'ignore ')' and '='
    Case Else
      Call PushImmed(MathOp)            'if not ")" or "=", then push DisplayReg and MathOp
  End Select
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

