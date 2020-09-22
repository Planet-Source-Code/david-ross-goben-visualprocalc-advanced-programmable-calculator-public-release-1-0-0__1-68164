Attribute VB_Name = "modAtan2"
Option Explicit
'~modATan2.bas;
'Returns the angle (in radians) from the X axis to a point
'*****************************************************************************
' modATan2: The ATan2() function returns the angle (in radians) from the X axis
'           to a point (y,x).
'*****************************************************************************

Public Function ATan2(ByVal Y As Double, ByVal X As Double) As Double
  Static Pie As Double
  
  If Pie = 0# Then    'keep a local copy of Pi
    Pie = Atn(1) * 4#
  End If
  
  On Error Resume Next
  If X = 0 Then
    If Y = 0 Then
      ATan2 = 0#
    ElseIf Y > 0# Then
      ATan2 = Pie / 2#
    Else
      ATan2 = -Pie / 2#
    End If
  ElseIf X > 0# Then
    If Y = 0# Then
      ATan2 = 0#
    Else
      ATan2 = Atn(Y / X)
    End If
  Else
    If Y = 0# Then
      ATan2 = Pie
    Else
      ATan2 = Atn(Y / X) - Pie
    End If
  End If
End Function

