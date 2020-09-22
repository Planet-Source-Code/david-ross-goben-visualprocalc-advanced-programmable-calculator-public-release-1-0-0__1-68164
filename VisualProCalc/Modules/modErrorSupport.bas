Attribute VB_Name = "modErrorSupport"
Option Explicit

'*******************************************************************************
' Subroutine Name   : ForcError
' Purpose           : Force an Error
'*******************************************************************************
Public Sub ForcError(Txt As String)
  Dim S As String
  
  InstrErr = InstrPtr + 1               'Program Step where error was encountered
  
  If Not flags(8) Or Not RunMode Then   'if flag 8 not set, we MUST suspend execution
    With frmVisualCalc
      S = " Error Description:" & vbCrLf & " " & Txt
      If Right$(S, 1) <> "?" Then S = S & "."
      S = S & vbCrLf & vbCrLf
      If RunMode Or Not CBool(ActivePgm) Or Preprocessing Or Compressing Then
        If RunMode Or Preprocessing Or Compressing Then
          S = S & " At Program Step: " & CStr(InstrPtr) & vbCrLf & vbCrLf
        End If
      ElseIf CBool(ActivePgm) Then
          S = S & " In Module Pgm: " & CStr(ActivePgm) & vbCrLf & vbCrLf
      End If
      .txtError(1).Text = S & " Press the [CE] key to continue..."
      
      .PicPlot.Visible = False            'hide plot, if present
      PlotTrigger = False
      RunMode = False
      MRunMode = 0
      ModPrep = 0
      Call Reset_Support                  'invoke RESET command
      InstrErr = InstrPtr + 1             'reset instrerr to offending line
      Call UpdateStatus
      Call MsgBeep(beepSystemExclamation) 'warn user
      .txtError(0).Visible = True         'display error prompts
      .txtError(1).Visible = True
      ErrorPause = True                   'Pause until cleared
      ErrorFlag = True                    'flag to program to stop running
    End With
    Call CloseAll                         'close all opened files
  End If
End Sub

'*******************************************************************************
' Function Name     : CheckError
' Purpose           : Check for error conditions
'*******************************************************************************
Public Sub CheckError()
  ErrorFlag = False                           'reset error flag
  If CBool(Err.Number) Then                   'if an error was encountered...
    DisplayReg = 0#                           'nullify display register
    DisplayText = False
    Call ForcError(Err.Description)           'force error processing
  End If
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

