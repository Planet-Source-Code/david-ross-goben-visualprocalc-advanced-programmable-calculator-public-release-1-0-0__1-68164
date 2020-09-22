Attribute VB_Name = "modDateFunctions"
Option Explicit
'~modDateFunctions.bas;
'Various Date-related functions
'*****************************************************************************
' modDateFunctions - Various function providing date-related information:
'
' GetLastDayOfMonth(): Returns the day number of the last day of month for a date
' IsLeapYear()       : Returns TRUE if the supplied date is in a leap year
' IsWeekEnd()        : Returns TRUE if the supplied date is a weekend
' IsWeekDay()        : Returns TRUE if the supplied date is a weekday
' AddDays()          : Returns the date after adding/subtracting a number of days
' AddMonths()        : Returns the date after adding/subtracting a number of months
' AddYears()         : Returns the date after adding/subtracting a number of year
' AddWeeks()         : Returns the date after adding/subtracting a number of weeks
' GetNextWeekDay()   : Return the date of the next weekday following the date
' GetNextDay()       : Return the date of the next specified day of week following the date
'*****************************************************************************

'*****************************************************************************
' GetLastDayOfMonth(): Returns the day number of the last day of month for a date
'*****************************************************************************
Public Function GetLastDayOfMonth(Dt As Date) As Date
  GetLastDayOfMonth = Day(DateSerial(Year(Dt), Month(Dt) + 1, 1) - 1)
End Function

'*****************************************************************************
' IsLeapYear(): Returns TRUE if the supplied date is in a leap year
'*****************************************************************************
Public Function IsLeapYear(Dt As Date) As Boolean
  IsLeapYear = IsDate("02/29/" & CStr(Year(Dt)))
End Function

'*****************************************************************************
' IsWeekEnd(): Returns TRUE if the supplied date is a weekend
'*****************************************************************************
Public Function IsWeekend(Dt As Date) As Boolean
  IsWeekend = Weekday(Dt) Mod 6 = 1
End Function

'*****************************************************************************
' IsWeekDay(): Returns TRUE if the supplied date is a weekday
'*****************************************************************************
Public Function IsWeekDay(Dt As Date) As Boolean
  IsWeekDay = Weekday(Dt) Mod 6 <> 1
End Function

'*****************************************************************************
' AddDays(): Returns the date after adding/subtracting a number of days
'*****************************************************************************
Public Function AddDays(Dt As Date, Dy As Integer) As Date
  AddDays = DateSerial(Year(Dt), Month(Dt), Day(Dt) + Dy)
End Function

'*****************************************************************************
' AddMonths(): Returns the date after adding/subtracting a number of months
'*****************************************************************************
Public Function AddMonths(Dt As Date, Mn As Integer) As Date
  AddMonths = DateSerial(Year(Dt), Month(Dt) + Mn, Day(Dt))
End Function

'*****************************************************************************
' AddYears(): Returns the date after adding/subtracting a number of year
'*****************************************************************************
Public Function AddYears(Dt As Date, yr As Integer) As Date
  AddYears = DateSerial(Year(Dt) + yr, Month(Dt), Day(Dt))
End Function

'*****************************************************************************
' AddWeeks(): Returns the date after adding/subtracting a number of weeks
'*****************************************************************************
Public Function AddWeeks(Dt As Date, wk As Integer) As Date
  AddWeeks = DateSerial(Year(Dt), Month(Dt), Day(Dt) + 7 * wk)
End Function

'*****************************************************************************
' GetNextWeekDay(): Return the date of the next weekday following the date
'*****************************************************************************
Public Function GetNextWeekDay(Dt As Date) As Date
  Dim dt1 As Date
  
  dt1 = Dt
  Do
    Dt = Dt + 1             'add 1 day
  Loop While IsWeekend(dt1) 'while this date is a weekend
  GetNextWeekDay = dt1
End Function

'*****************************************************************************
' GetNextDay(): Return the date of the next specified day of week following the date
'*****************************************************************************
Public Function GetNextDay(Dt As Date, DayWeek As VbDayOfWeek) As Date
  Dim dt1 As Date
  dt1 = Dt
  
  Do
    dt1 = dt1 + 1
  Loop While Not Weekday(dt1) <> DayWeek
  GetNextDay = dt1
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

