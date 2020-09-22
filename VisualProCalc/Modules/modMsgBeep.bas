Attribute VB_Name = "modMsgBeep"
Option Explicit
'~modMsgBeep.bas;
'Play a wav file instead of the system default beep (if supported)
'**********************************************************************
' modMsgBeep:
' The MsgBeep() function allows you to play a wav file instead of
' the system default beep. If the function returs TRUE, then the wave file
' was played. If it returns FALSE, then a wave file could not be played,
' and so the system default beep was sounded.
'
' The MsgType parameter has the following values:
'  beepSystemDefault: System default beep
'  beepSystemAsterisk: wav file associated with the Asterisk event
'  beepSystemExclamation: wav file associated with the Exclamation event
'  beepSystemHand: wav file associated with the Hand event
'  beepSystemQuestion: wav file associated with the Question event
'  beepSystemDefault: play the system default beep (same as the Beep command)
'
'EXAMPLE:
'  Call MsgBeep(beepSystemExclamation)
'
'**********************************************************************

'**********************************************************************
' Types ands API calls
'**********************************************************************
Public Enum BeepType
  beepSystemDefault = &HFFFFFFFF  'same as using the VB Beep command
  beepSystemAsterisk = &H40&
  beepSystemExclamation = &H30&
  beepSystemHand = &H10&
  beepSystemQuestion = &H20&
'  beepSystemDefault = &H0&
End Enum

Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm" () As Long

'**********************************************************************
' Beep function
'**********************************************************************
Public Function MsgBeep(MsgType As BeepType) As Boolean
  If waveOutGetNumDevs() Then
    Call MessageBeep(MsgType)
    MsgBeep = True              'we could sound off
  Else
    Beep
    MsgBeep = False             'sounded off with default beep
  End If
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

