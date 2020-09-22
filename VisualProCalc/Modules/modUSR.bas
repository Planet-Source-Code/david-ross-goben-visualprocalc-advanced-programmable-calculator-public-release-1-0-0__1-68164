Attribute VB_Name = "modUSR"
'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************
'This shell is copyrighted, as is all other code in this application. However,
'you are free to alter the MaxUSR constant, add invocations to the ProcessUSR()
'subroutine, and add your own USRxx() subroutines to this module, and distribute
'them at no cost, without violating the copyright restrictions that hold over
'all other modules in this application.
'******************************************************************************

Option Explicit

Public Const MaxUSR As Long = 2  'max USR operations supported (set to the number you have defined)

'*******************************************************************************
' m o d U S R
'*******************************************************************************
'This Module is set aside for USER-DEFINED functions. This module is similar to the OP command,
'but is reserved for program-user-defined functions.
'*******************************************************************************

'*******************************************************************************
' Subroutine Name   : ProcessUSR
' Purpose           : Invoke USER DEFINED OPERATIONS
'*******************************************************************************
Public Sub ProcessUSR(ByVal Index As Integer)
  'USR function invocation go here
  Select Case Index
    Case 0
      Call USR00
    Case 1
      Call USR01
    Case 2
      Call USR02
'    Case 3
'      Call USR03
'    Case 4
'      Call USR04
'    Case 5
'      Call USR05
'    Case 6
'      Call USR06
'    Case 7
'      Call USR07
'    Case 8
'      Call USR08
'    Case 9
'      Call USR09
'    Case 10
'      Call USR10
'    Case 11
'      Call USR11
'    Case 12
'      Call USR12
'    Case 13
'      Call USR13
'    Case 14
'      Call USR14
'    Case 15
'      Call USR15
'    Case 16
'      Call USR16
'    Case 17
'      Call USR17
'    Case 18
'      Call USR18
'    Case 19
'      Call USR19
'    Case 20
'      Call USR20
'    Case 21
'      Call USR21
'    Case 22
'      Call USR22
'    Case 23
'      Call USR23
'    Case 24
'      Call USR24
'    Case 25
'      Call USR25
'    Case 26
'      Call USR26
'    Case 27
'      Call USR27
'    Case 28
'      Call USR28
'    Case 29
'      Call USR29
'    Case 30
'      Call USR30
'    Case 31
'      Call USR31
'    Case 32
'      Call USR32
'    Case 33
'      Call USR33
'    Case 34
'      Call USR34
'    Case 35
'      Call USR35
'    Case 36
'      Call USR36
'    Case 37
'      Call USR37
'    Case 38
'      Call USR38
'    Case 39
'      Call USR39
'    Case 40
'      Call USR40
'    Case 41
'      Call USR41
'    Case 42
'      Call USR42
'    Case 43
'      Call USR43
'    Case 44
'      Call USR44
'    Case 45
'      Call USR45
'    Case 46
'      Call USR46
'    Case 47
'      Call USR47
'    Case 48
'      Call USR48
'    Case 49
'      Call USR49
'    Case 50
'      Call USR50
'    Case 51
'      Call USR51
'    Case 52
'      Call USR52
'    Case 53
'      Call USR53
'    Case 54
'      Call USR54
'    Case 55
'      Call USR55
'    Case 56
'      Call USR56
'    Case 57
'      Call USR57
'    Case 58
'      Call USR58
'    Case 59
'      Call USR59
'    Case 60
'      Call USR60
'    Case 61
'      Call USR61
'    Case 62
'      Call USR62
'    Case 63
'      Call USR63
'    Case 64
'      Call USR64
'    Case 65
'      Call USR65
'    Case 66
'      Call USR66
'    Case 67
'      Call USR67
'    Case 68
'      Call USR68
'    Case 69
'      Call USR69
'    Case 70
'      Call USR70
'    Case 71
'      Call USR71
'    Case 72
'      Call USR72
'    Case 73
'      Call USR73
'    Case 74
'      Call USR74
'    Case 75
'      Call USR75
'    Case 76
'      Call USR76
'    Case 77
'      Call USR77
'    Case 78
'      Call USR78
'    Case 79
'      Call USR79
'    Case 80
'      Call USR80
'    Case 81
'      Call USR81
'    Case 82
'      Call USR82
'    Case 83
'      Call USR83
'    Case 84
'      Call USR84
'    Case 85
'      Call USR85
'    Case 86
'      Call USR86
'    Case 87
'      Call USR87
'    Case 88
'      Call USR88
'    Case 89
'      Call USR89
'    Case 90
'      Call USR90
'    Case 91
'      Call USR91
'    Case 92
'      Call USR92
'    Case 93
'      Call USR93
'    Case 94
'      Call USR94
'    Case 95
'      Call USR95
'    Case 96
'      Call USR96
'    Case 97
'      Call USR97
'    Case 98
'      Call USR98
'    Case 99
'      Call USR99
    '----------------------------------------
    'enter your owen Case tests of Index here
    '----------------------------------------
  End Select
End Sub

'*******************************************************************************
' Subroutine Name   : USR00
' Purpose           : Sample User-Defined Operation. Change as you see fit.
'*******************************************************************************
Private Sub USR00()
  DspTxt = "Test USR Operations"    'set Displat Text variable
  DisplayText = True                'indicate we will display DspTxt, not default numeric DisplayReg
  Call DisplayLine                  'display data on the display line
End Sub

'*******************************************************************************
' Subroutine Name   : USR01
' Purpose           : Sample User-Defined Operation. Change as you see fit.
'*******************************************************************************
Private Sub USR01()
  DisplayMsg "Use [X><T] to see Test Register Results"
  DisplayReg = 12345                'set default displayregister
  TestReg = 54321                   'set the test register (resides "behind" [X><T])
  DisplayText = False               'ensure numeric register (DisplayReg) it will be displayed
  Call ForceDisplay                 'force data to display on the display line, even in RUN Mode
End Sub

'*******************************************************************************
' Subroutine Name   : USR02
' Purpose           : Sample User-Defined Operation. Change as you see fit.
'*******************************************************************************
Private Sub USR02()
  ForcError "This is a test error created by USR 02"  'test error reporting
End Sub

