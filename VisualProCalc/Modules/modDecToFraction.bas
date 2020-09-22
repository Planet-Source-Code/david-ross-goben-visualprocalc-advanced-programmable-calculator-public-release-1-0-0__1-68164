Attribute VB_Name = "modDecToFraction"
Option Explicit
'~modDecToFraction.bas;
'Convert Decimal number to fractional string. "1.125" becomes "1-1/8"
'******************************************************************************
' modDecToFraction - The DecToFraction() function takes a value, either as
'                    any numeric or string, and returns a string with the
'                    fractional equivalent. If a whole number is also part
'                    of the value, then separate the whole from the fraction
'                    with the optional WholeSep parameter string. By default
'                    this separator is "-". With this function, a value of
'                    .25 becomes "1/4", and 1.125 becomes "1-1/8".
'
'EXAMPLE. This will debug print "3 3/16", "4·1/500", and "1/8"
'  Debug.Print DecToFraction(3.1875, " ") 'display with space separator
'  Debug.Print DecToFraction(4.002)       'display with default separator
'  Debug.Print DecToFraction(.125)
'******************************************************************************

Public Function DecToFraction(Dec As Variant, Optional WholeSep As String = "·") As String
  Dim dDec As Double, Whole As Double, tDec As Double, Tmp As Double
  Dim TwoCnt As Long, Idx As Long
  Dim Result As String
  
  Result = vbNullString
  If IsNumeric(Dec) Then                  'if value is in fact numeric
    dDec = CDbl(Dec)                      'get numeric value
    Whole = Fix(dDec)                     'get whole number
    If Whole = dDec Then                  'if no fractions
      Result = CStr(dDec)                 'result = value
    Else
      If CBool(Whole) Then                'if a whole number ALSO present
        Result = CStr(Whole) & WholeSep   'apply Whole and separator
      End If
      dDec = Abs(Round(dDec - Whole, 12))  'get decimal part to dDec
      tDec = 1# / dDec                    'see if 1/value creates whole
      If Fix(tDec) = tDec Then            'does, so is 1/value
        Result = Result & "1/" & CStr(tDec)
      Else
        TwoCnt = 1                        'init denominator value
        tDec = dDec                       'grab copy of decimal value
        Do While Fix(tDec) <> tDec        'now search for fraction
          TwoCnt = TwoCnt * 2             'double denominator
          tDec = tDec * 2#                'multiply value times 2
          If Fix(tDec) > 1# Then          'has become greater than 1?
            If Fix(tDec) = tDec Then Exit Do  'is base 2 denominator
            tDec = tDec / 2#              'is > 1, so back it off
            For Idx = TwoCnt To TwoCnt * 2 - 1 'scan for intermediate fraction
              Tmp = Round(CDbl(Idx) * dDec + 0.000005, 5)
              If Fix(Tmp) = Tmp Then      'make whole?
                TwoCnt = Idx              'yes, so found denominator
                tDec = Tmp                'get numerator
                Exit For
              End If
            Next Idx
          End If
        Loop
        Result = Result & CStr(tDec) & "/" & CStr(TwoCnt)
      End If
    End If
  Else
    Result = vbNullString                 'value was not numeric
  End If
  DecToFraction = Result
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

