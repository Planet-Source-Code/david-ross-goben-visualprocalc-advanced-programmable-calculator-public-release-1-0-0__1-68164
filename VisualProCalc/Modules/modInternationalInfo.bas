Attribute VB_Name = "modInternationalInfo"
Option Explicit
'~modInternationalInfo.bas;
'Return local system formats for items, such as date, separators, etc.
'*******************************************************************************
' modInternationalInfo - The GetInternationalInfo() function Returns local
'                        system formats for items, such as date, separators, etc.,
'                        for the current user or as the system defaults.
'                        The SetInternationalInfo() function allows you to change
'                        the international data for the current user or as the
'                        system defaults.
'EXAMPLES:
'  Debug.Print "Local Currency Symbol =" & GetInternationalInfo(INTLsCurrency)
'  Debug.Print "Local Long Date Format =" & GetInternationalInfo(INTLsLongDate)
'  Debug.Print "Local Long Country Name=" & GetInternationalInfo(INTLsCountry)
'
' NOTE: GetInternationalInfo() returns all values as strings.
' NOTE: SetInternationalInfo() sends all values as strings.
'*******************************************************************************

Public Enum IntlValues
  INTLLocale = &H1          'locale code: US=409
  INTLsLanguage = &H3       'ENU (Abbreviated for English, US)
  INTLsLLanguage = &H1001   'English
  INTLsCountry = &H1002     'United States
  INTLiCountry = &H5        '1 (United States)
  INTLsList = &HC           ', (list separator)
  INTLiMeasure = &HD        '1 (US), 0=Metric
  INTLsDecimal = &HE        '. (decimal character)
  INTLsThousand = &HF       ', (thousands separator)
  INTLiDigits = &H11        '2 (minimum number of number locations)
  INTLiLZero = &H12         '1 (leading zero digit location)
  INTLsPosSign = &H50       '+ (Sign for positive)
  INTLsNegSign = &H51       '- (Sign for negative)
  INTLsCurrency = &H14      '$ (Currency symbol)
  INTLiCurrDigits = &H19    '2 (Pennies digits)
  INTLiCurrency = &H1B      '0
  INTLiNegCurr = &H1C       '0
  INTLsDate = &H1D          '/ (Date separator)
  INTLsTime = &H1E          ': (Time separator)
  INTLsShortDate = &H1F     'M/d/yy
  INTLsLongDate = &H20      'dddd, MMMM dd, yyyy
  INTLsTimeFormat = &H1003  'HH:mm:ss
  INTLiDate = &H21          '0 (order for short date; 0=mdy, 1=dmy, 2=ymd)
  INTLiLDate = &H22         '0 (order for long date)
  INTLiTime = &H23          '0 (0=AM/PM 12-hour, 1=24-hour)
  INTLiCentury = &H24       '0=2-digit century, 1=4-digit century
  INTLiTLZero = &H25        '0=no leading zero in time field, 1=use leading zeros
  INTLiDLZero = &H26        '1=use leading zero on days less than 2 digits
  INTLiMLZero = &H27        '1=use leading zero on months less than 2 digits
  INTLs1159 = &H28          'AM (Ante Meridian symbol)
  INTLs2359 = &H29          'PM (Post Meridian symbol)
End Enum

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Private Const LOCALE_SYSTEM_DEFAULT = &H800   'system default
Private Const LOCALE_USER_DEFAULT = &H400     'user default

Public Function GetInternationalInfo(INTL As IntlValues, Optional SystemDefault As Boolean = False) As String
  Dim i As Long, Loc As Long
  Dim S As String
  
  If SystemDefault Then                       'set system or user defaults
    Loc = LOCALE_SYSTEM_DEFAULT
  Else
    Loc = LOCALE_USER_DEFAULT
  End If
  
  S = vbNullString
  i = GetLocaleInfo(Loc, CLng(INTL), S, 0&)   'get space needed
  If i Then
    S = String$(i, 0)                         'set aside space
    i = GetLocaleInfo(Loc, CLng(INTL), S, i)  'get data
    If i Then
      i = InStr(1, S, vbNullChar)
      If i = 0 Then i = Len(S) + 1
      S = Left$(S, i - 1)
    Else
      S = vbNullString
    End If
  End If
  GetInternationalInfo = S
End Function

Public Function SetInternationalInfo(INTL As IntlValues, Value As String, Optional SystemDefault As Boolean = False) As String
  Dim Strretdate As String, StrFormat As String, strItem As String, S As String
  Dim i As Long, Loc As Long

  If SystemDefault Then                       'set system or user defaults
    Loc = LOCALE_SYSTEM_DEFAULT
  Else
    Loc = LOCALE_USER_DEFAULT
  End If
  SetInternationalInfo = SetLocaleInfo(Loc, CLng(INTL), Value)
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

