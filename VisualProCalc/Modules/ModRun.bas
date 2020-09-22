Attribute VB_Name = "ModRun"
Option Explicit

'*******************************************************************************
' Subroutine Name   : Run
' Purpose           : Handle Run Mode Code Execution
'*******************************************************************************
Public Sub Run()
  Dim Errstr As String                        'error reporting
  Dim Code As Integer                         'active instruction
  Dim Bol As Boolean                          'flag used by some instructions
  Dim Idx As Integer, i As Integer
'---------------------------------------------
  Tron = TraceFlag And ActivePgm = 0          'set trace mode, if active
  StopMode = False                            'turn off stop mode
  ErrorFlag = False                           'turn off error flag
  InstrErr = 0                                'clear last encountered error location
  Errstr = vbNullString                       'turn off error string
  Cancel = False                              'init test flag
  TextEntry = False                           'turn off text entry mode
  RS_Pressed = False                          'turn off R/S Pressed flag
  Call ResetPnd                               'make sure CharLimit and CharCount are null
  Call ResetValueAccum                        'and accumulator
  If Not CBool(ActivePgm) Then                'do not try to Compress module programs
    If Not Preprocessd Then                   'if program not Preprocessed...
      Call Preprocess                         'pre-process it
      If Not Preprocessd Then                 'if program still not Preprocessed...
        RunMode = False                       'ensure we turn off running modes
        SSTmode = False
        Cancel = False                        'reset Cancel flag
        Exit Sub                              'then leave
      End If
    End If
  
    RunMode = True                            'all OK, so ensure we turn on Run Mode
    
    If Not SSTmode Then                       'if not Single-Step mode...
      Call UpdateStatus                       'update status reports
      
      IgnoreClick = True                      'ignore clicks (next instruction will fire Click event)
      With frmVisualCalc.lstDisplay
        .List(.ListIndex) = vbNullString      'stuff blank in active display
      End With
      IgnoreClick = False
      
      With frmVisualCalc                      'turn off menu options
        .mnuFile.Enabled = False
        .mnuWindow.Enabled = False
        .mnuHelp.Enabled = False
        .mnuMemStk.Enabled = False
      End With
    End If
    If InstrPtr >= GetInstrCnt() Then InstrPtr = 0  'reset run mode
  Else
    RunMode = True                            'turn on run mode for module program
  End If
  SetTip vbNullString
  DoEvents
  DoDebug = False                             'avoid entering Debug unexpectedly
  
  If CBool(HMRunMode) Then
    MRunMode = HMRunMode
    HMRunMode = 0
  End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'    Debug.Assert False            'stop the train until we get this thing working
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  Do While InstrPtr < GetInstrCnt()           'need to check only user program
    Code = GetInstruction(0)                  'get code to process
'
' if DEF processing active and the flag is not true, we will allow only DELSE and ENDDEF
'
'  Debug.Print CStr(InstrPtr) & ": " & GetInst(Code)
    If CBool(DefDef) Then                     'if DEF processing going on
      If Not DefTrue(DefDef) Then             'if disabled code being processed
        Select Case Code
          Case iEDef, iDelse                  'allow only Edef and Delse to continue
          Case Else
            Code = iNOP                       'ignore any other codes
        End Select
      End If
    End If
'---------------------------------------------
    If Code <> iNOP Then                      'if not NOP, then process it
      If CBool(ModPrep) Then                  'if external program will be invoked...
        Select Case Code
          Case iCall                          'calling a subroutine is OK
          Case Is > 900                       'calling a user-key is OK
          Case Else                           'but nothing else
            Errstr = "Invalid instrution. Can only be Call or User-Key after invoking a Pgm"
            Exit Do
        End Select
      End If
      
      Call RunCmd(Code, Errstr)               'process command if OK so far
'
' if single-step mode, we are in Pgm 00, and we will exit running
'
      If SSTmode Then                           'if Single-step mode active...
        If Not CBool(ActivePgm) Then            'ignore if Module program
          i = InstrPtr + 1                      'pre-check next instruction
          If i = InstrCnt Then                  'If at top of program, we will stop anyway
            MsgBeep beepSystemHand              'let single-step user know end was met
            SSTmode = False
          Else
            For Idx = 0 To InstCnt3 - 1         'else scan tables
              If InstMap3(Idx) = i Then         'found a match?
                IncInstrPtr                     'yes, bump instruction pointer...
                Exit Do                         'and exit for SST operation
              ElseIf i < InstMap3(Idx) Then     'else not at next instruction?
                Exit For                        'no, so continue normally
              End If
            Next Idx
          End If
        End If
      End If 'SSTmode
    End If 'Code <> iNOP
'
' now check for errors, and exit if any encountered
'
    If Not flags(8) Or Not RunMode Then
      If CBool(Len(Errstr)) Or CBool(InstrErr) Then
        Exit Do 'if error (keep InstrPtr on error)
      End If
    End If
'---------------------------------------------
' process background tasks, then if R/S was pressed, stop running
'
    If RS_Pressed Then
      RunMode = False
      MRunMode = 0
      ModPrep = 0
      Exit Do
    End If
'---------------------------------------------
' bump the instruction pointer
'
    If IncInstrPtr() Then                     'bump instruction pointer (return True if loop-around)
      If Not CBool(ActivePgm) Then Exit Do    'exit if user pgm pointer reset, or pgm 00
      If Not RunMode Then Exit Do             'we were not co-running pgm 00
      ActivePgm = 0                           'if module reset, reactivate pgm 0 (user program)
      MRunMode = 0                            'we will loop, so ensure we turn off module run mode
      ModPrep = 0
    End If
    If DoDebug Then Exit Do                   'see if we want to debug
    If PmtFlag Then Exit Do                   'user is to be prompted
    If Not RunMode Then Exit Do               'if run mode was turned off
  Loop                                        'and process its data
'-------------------------------------------------------------------------------
  If InstrPtr = GetInstrCnt() Then InstrPtr = 0 'reset base if we must "loop" around from top to bottom
'
' if there is a yet unreported error, report it now
'
  If CBool(Len(Errstr)) Or CBool(InstrErr) Then 'if InstrErr is set, then an error already reported
    PmtFlag = False                           'ignore prompting if errors
    DoDebug = False                           'ensure debug mode is not active if any error
    If Not CBool(InstrErr) Then               'if we have not had error reported, but one exists...
      ForcError Errstr
    End If
    ActivePgm = 0                             'return to user level
  End If
'
' turn off just about any mode that could possibly be on
'
  SSTmode = False                             'ensure we turn off single step mode
  RunMode = False                             'ensure RUN mode is off
  ModPrep = 0
  ForInit = False
'
' if not stacked via pgm invokes
'
  If Not CBool(SbrInvkIdx) Then
'
' enable menu items
'
    With frmVisualCalc                      'turn off menu options
      If Not .mnuFile.Enabled Then          'if menus not already enabled
        .mnuFile.Enabled = True
        .mnuWindow.Enabled = True
        .mnuHelp.Enabled = True
        .mnuMemStk.Enabled = True
'
' clean up peripheral things on the display
'
        Call UpdateStatus                    'update status
        Call DisplayLine                     'redo display data
      End If
    End With
  End If
'
' handle forcing the entry into the Debug Mode
'
  If DoDebug Then                             'if we want to debug
    DoDebug = False                           'turn the flag off
    Bol = Key2nd                              'save state of 2nd key
    Key2nd = False                            'force it off
    Call MainKeyPad(1)                        'force LRN key
    Key2nd = Bol                              'reset state of 2nd key
  End If
  Tron = TraceFlag And ActivePgm = 0          'set trace mode, if active
  DoEvents
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

