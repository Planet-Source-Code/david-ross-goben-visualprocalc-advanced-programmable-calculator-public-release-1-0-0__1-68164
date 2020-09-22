Attribute VB_Name = "modVal"
Option Explicit
'~modVal.bas;
'Override for Val() function to allow for international settings
'***************************************************************************
' modVal -- The Val() function provides an override for the VBA.Val() function
'           that will allow for international functions. Blank input is addressed.
'***************************************************************************

Public Function Val(InString As String) As Double
  Static IntlDate As String                 'local Date symbol (US='/')
  Static IntlTime As String                 'local Time symbol (US=':')
  
  Dim S As String
  Dim Idx As Integer, Idy As Integer, i As Integer
  Dim Sn As Double, TV As Double
'
' get local copies of date and time constants
'
  If IntlDate = vbNullString Then
    IntlDate = GetInternationalInfo(INTLsDate)
    IntlTime = GetInternationalInfo(INTLsTime)
  End If
  
  S = Trim$(InString)                       'clean up act
  If CBool(Len(S)) Then                     'if data present
    If Left$(S, 1) = "&" Then               'if possible Hex (&H), Oct (&O) or Bin (&B) value
      Select Case UCase$(Mid$(S, 2, 1))
        Case "H", "O", "B"                  'allow Hex, Octal, and Binary defs to slip through
        Case Else
          Do
            Idx = InStr(1, S, "%")          'replace Integer "%" or Long "&" with "$" to...
            If Not CBool(Idx) Then Idx = InStr(1, S, "&")
            If Not CBool(Idx) Then Exit Do
            Mid$(S, Idx, 1) = "$"           '... prevent type mismatch
          Loop
      End Select
    End If
'
' check for possible time
'
    If CBool(InStr(1, S, IntlTime)) Then        'found time tag?
      On Error Resume Next
      Val = CDbl(CDate(S))                      'check for time
      If CBool(Err.Number) Then Val = 0#
      On Error GoTo 0
      Exit Function
    End If
'
' check for possible Date
'
    Idx = InStr(1, S, IntlDate)                 'found a date tag?
    If CBool(Idx) Then
      Idy = InStr(Idx + 1, S, IntlDate)         'found another Date tag?
      If CBool(Idy) Then
        On Error Resume Next
        Val = CDbl(CDate(S))                    'check for date
        If CBool(Err.Number) Then Val = 0#
        On Error GoTo 0
        Exit Function
      End If
    End If
'
' check for fractional expression, such as "3-3/16"
'
    Idx = InStr(1, S, "/")                      'found a possible denominator?
    If CBool(Idx) Then
      i = InStrRev(S, " ", Idx)               'check for " " separator (3 3/16)
      If i = 0 Then i = InStrRev(S, "·", Idx) 'check for "·" separator (3·3/16)
      If i = 0 Then i = InStrRev(S, "-", Idx) 'check for "-" separator (3-3/16)
      If i = 0 Then i = InStrRev(S, "+", Idx) 'check for "+" separator (3+3/16)
      On Error Resume Next
      Select Case i
        Case Is > 1
          TV = DoVal(Left$(S, i - 1))
          Sn = Sgn(TV)
          Val = (Abs(TV) + DoVal(Mid$(S, i + 1, Idx - i - 1)) / DoVal(Mid$(S, Idx + 1))) * Sn
          If CBool(Err.Number) Then Val = 0#
          Exit Function
        Case Else
          Val = DoVal(Left$(S, Idx - 1)) / DoVal(Mid$(S, Idx + 1))
          If CBool(Err.Number) Then Val = 0#
          Exit Function
      End Select
      On Error GoTo 0
    End If
    Val = DoVal(S)
  End If
End Function

Private Function DoVal(ByVal Value As String) As Double
  On Error Resume Next
  DoVal = CDbl(Trim$(Value))                       'convert to double (intl aware)
  If CBool(Err.Number) Then DoVal = VBA.Val(Value) 'if error, go through safer, non-Intl route
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

