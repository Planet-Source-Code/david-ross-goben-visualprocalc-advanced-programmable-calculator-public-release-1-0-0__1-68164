Attribute VB_Name = "ModMathSupport"
Option Explicit

'*******************************************************************************
' Function Name     : AngToRad
' Purpose           : Convert angle in user-selected type to Radians
'*******************************************************************************
Public Function AngToRad(ByVal Angle As Double) As Double
  Select Case AngleType
    Case TypDeg
      AngToRad = Angle * vPi / 180#
    Case TypGrad
      AngToRad = Angle * vPi / 200#
    Case TypMil
      AngToRad = Angle * vPi / 6400#
    Case TypRad
      AngToRad = Angle
  End Select
End Function

'*******************************************************************************
' Function Name     : RadToAng
' Purpose           : Covert angle in Radians to user-selected type
'*******************************************************************************
Public Function RadToAng(ByVal RadAngle As Double) As Double
  Select Case AngleType   'also convert angle to radians
    Case TypDeg
      RadToAng = RadAngle * 180# / vPi
    Case TypGrad
      RadToAng = RadAngle * 200# / vPi
    Case TypMil
      RadToAng = RadAngle * 6400# / vPi
    Case TypRad
      RadToAng = RadAngle
  End Select
End Function

'*******************************************************************************
' Function Name     : Factorial
' Purpose           : Compute the factorial of a number
'*******************************************************************************
Public Function Factorial(Value As Double) As Boolean
  Dim TV As Double
  Dim Idx As Long
  
  TV = Fix(Value)
  If TV < 0# Or TV > 69# Then
    ForcError "Factorial range is 0 - 69"
    Exit Function
  Else
    If TV > 2 Then                        '0 and 1 and 2 will not change
      For Idx = CLng(TV) - 1 To 2 Step -1 'add (n) * (n-1) * (n-2) ...
        TV = TV * CDbl(Idx)
      Next Idx
    End If
  End If
  Value = TV
  Factorial = True
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

