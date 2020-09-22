Attribute VB_Name = "modKybdMain"
Option Explicit

'*******************************************************************************
' Subroutine Name   : MainKeyPad
' Purpose           : Keypad click. Note that codes are +128 (2nds are +512).
'*******************************************************************************
Public Sub MainKeyPad(Index As Integer)
  Dim S As String, Ary() As String
  Dim Idx As Integer, i As Integer
'
' process OPCODES as > 128. 2nd keys are 128 higher than non-2nd keys
'
  If Index = 0 Then Exit Sub      'ignore any useless data
  DisplayText = False             'turn off display text mode
  If Key2nd Then
    LastTypedInstr = Index + 256  'mark 2nd key opcodes as 256+
  Else
    LastTypedInstr = Index + 128  'mark normal opcodes as 128+ (128 is ignored)
  End If
'
' if Running, ignore anything from the keyboard, except R/S
'
  If RunMode Then
    If LastTypedInstr <> iRunStop Then Exit Sub
    RunMode = False                     'disable modes
    ModPrep = 0
    SSTmode = False
    Call ResetPnd                       'reset data
    Call UpdateStatus                   'update status
    Call DisplayLine                    'show display register
    Exit Sub
  End If
'
' if an error is being displayed, then do nothing until [CE] is pressed
'
  If ErrorPause Then                          'if error pausing...
    Select Case LastTypedInstr
      Case CEKey
      Case LRNKey
        Call CE_Support
      Case Else
        Call CmdNotActive
        Exit Sub                              'exit if we did not press [CE] or [LRN]
    End Select
  End If
'
' check for help being displayed
'
  If frmVisualCalc.rtbInfo.Visible Then
    If LastTypedInstr = CEKey Then
      Call CE_Support
    Else
      CmdNotActive
    End If
    Exit Sub
  End If
'
' check for DspLocked and CLR key
'
  If Not LrnMode And DspLocked Then
    Select Case LastTypedInstr
      Case LRNKey, iList, iSbr, iLbl, iUkey, iVar, iConst, iStruct, iCP, iPgm, iStyle
      Case 220, 207, 208, 209, 294, 295, 296, 181, 182, 183 '0-9
        If CBool(PndIdx) Then
          If PndStk(PndIdx) <> iStyle Then
            CmdNotActive
            Exit Sub
          End If
        End If
      Case CEKey
        DspLocked = False
        Call DspBackground
        LastTypedInstr = iCLR
      Case Else
        CmdNotActive
        Exit Sub
    End Select
  End If
'
' translate numeric keys to 0-9
'
  Select Case LastTypedInstr
    Case 220  '0
      LastTypedInstr = 0
    Case 207  '1
      LastTypedInstr = 1
    Case 208  '2
      LastTypedInstr = 2
    Case 209  '3
      LastTypedInstr = 3
    Case 194  '4
      LastTypedInstr = 4
    Case 195  '5
      LastTypedInstr = 5
    Case 196  '6
      LastTypedInstr = 6
    Case 181  '7
      LastTypedInstr = 7
    Case 182  '8
      LastTypedInstr = 8
    Case 183  '9
      LastTypedInstr = 9
  End Select
'
' if text entry mode, then reset it if non-numeric keypad entry
'
  If LrnMode Then
    Select Case LastTypedInstr
      Case Is > 127
        CharLimit = 0                           'force off so Display Line will work as expected
        PmtFlag = False
        REMmode = 0
    End Select
  Else
    If TextEntry Then
      Select Case LastTypedInstr
        '
        ' check for digits, and translate to keystrokes
        '
        Case 0 To 9
          S = Chr$(LastTypedInstr + 48)
          
          If VarLbl Then            'if Variable Labeling is active
            If CharCount = 0 Then   'if no other character has been entered
              VarLbl = False        'turn off option
              CharLimit = 2         'force limit to 2
            End If
          End If
        
        Case iDot
          S = "."
        
        Case iComma
          S = ","
        '
        'allow other keys to terminate text entry, without costing the user an instruction
        '
        Case Else
          S = vbNullString                            'no digit input
          AllowSpace = False                          'turn off text entry mode
          Call frmVisualCalc.checkTextEntry(False)
          'reset LastTypedInstr if it was [Txt] or [=]
          If LastTypedInstr = iTXT Or LastTypedInstr = iEqual Then
            LastTypedInstr = 128                      'disable key action
            If CBool(PndIdx) Then Call CheckPnd(0)
            If CharCount = 0 And REMmode <> 3 Then    'if nothing typed...
              REMmode = 0                             'ensure remarks mode is disabled
              CharLimit = 0
              Call ResetPnd                           'reset anything pending
              Call DisplayLine                        'display DisplayReg value
            ElseIf PmtFlag Then                       'something in DspText, are we prompting?
              Call NewLine
              CharLimit = 0                           'force off so Display Line will work as expected
              StopMode = False                        'ensure Stop mode is not active
              LastTypedInstr = 219                    'force R/S, to continue running
              PmtFlag = False                         'turn off prompt flag
              DisplayText = True                      'ensure text will be processed
            Else
              PmtFlag = False
            End If
          Else
           CharLimit = 0                           'force off so Display Line will work as expected
           PmtFlag = False
           REMmode = 0
          End If
          '
          ' special display handling if we were in either REMARKS modes
          '
          If CBool(REMmode) Then                      'if remarks mode
            With frmVisualCalc.lstDisplay
              Select Case REMmode
                Case 1  '[Rem ]
                  .List(.ListIndex) = "Rem " & DspTxt 'standard
                  Call ResetPndAll
                  Call NewLine                        'advance line so our remark is maintained
                Case 2  '[']
                  .List(.ListIndex) = "'" & DspTxt    'enhanced
                  Call ResetPndAll
                  Call NewLine                        'advance line so our remark is maintained
                Case 3  'Fmt
                  If LastTypedInstr = iIND Then       'allow Fmt Ind
                    PushPendKey                       'push data to stack
                    PndStk(PndIdx) = iFmt             'park Fmt Ind xx (later PushPendKey will place IND)
                  Else
                    DspFmt = DspTxt                   'employ user-supplied format
                    DspFmtFix = -1                    'disabled fixed format
                    Call ResetPndAll
                    Call DisplayLine
                  End If
              End Select
            End With
            REMmode = 0                             'turn off remarks mode
          End If
      End Select
  '
  ' if numerics entered, then process them
  '
      If CBool(Len(S)) Then
        LastTypedInstr = 128            'turn off key if digit for text entry
        DspTxt = DspTxt & S             'append digit
        CharCount = Len(DspTxt)         'get length of data
        With frmVisualCalc.lstDisplay
          .List(.ListIndex) = String$(DisplayWidth - Len(DspTxt), 32) & DspTxt
          DisplayHasText = True
        End With
      End If
    End If
  '
  ' If LastTypedInstr was nullified, then nothing more to do here
  '
    If LastTypedInstr = 128 Then Exit Sub
  '
  ' reset numeric input if not numeric keys
  '
    If LastTypedInstr <> CEKey Then 'if not CE key
      Select Case LastTypedInstr
        Case EEKey                                        'EE Key...
          If IsNumbers Then                               'if we were entering a number...
            AllowExp = True                               'then allow typing the exponent
          End If
          
        Case 0 To 9, 221, 222                             '0-9, [.], or [+/-]
          IsNumbers = True                                'indicate numeric data
          
          If VarLbl Then                                  'if Variable Labeling is active
            If CharCount = 0 Then                         'if no other character has been entered
              VarLbl = False                              'turn off option
              CharLimit = 2                               'force limit to 2
            End If
          End If
          
          If CBool(CharLimit) Then                        'if background gathering of value...
            CharCount = CharCount + 1                     'bump character count gathered
          End If
        
        Case Else
          If BaseType = TypHex Then                       'if we are entering a hex value...
            Select Case LastTypedInstr
              Case iArc, 167, 180, 193, 206, 219          'A-F substitution keys?
                IsNumbers = True                          'if so, indicate numeric data
                If CBool(CharLimit) Then                  'if background gathering of value...
                  CharCount = CharCount + 1               'bump character count gathered
                End If
              Case Else
                Call CheckPnd(LastTypedInstr)             'else check if it is a pending-type op
                IsNumbers = False
            End Select
          Else
            Call CheckPnd(LastTypedInstr)                 'see if it is a pending-type op
            IsNumbers = False
          End If
      End Select
    End If
  End If
'
' special case if LRN key pressed. LRN must be able to flip control of learning
'
  If LastTypedInstr = LRNKey Then
    LrnMode = Not LrnMode             'toggle learn mode
    PlayWavResource "Swoosh"
    INSmode = LrnMode                 'disable insert mode, if on, in any case
    Call UpdateStatus                 'update status data for both changes
    With frmVisualCalc
      .cmdBtnHelp.Enabled = Not LrnMode And HaveVHelp Or frmVisualCalc.mnuHelpSepHlp.Checked
      .mnuWindow.Enabled = Not LrnMode    'enable/disable some menu options
      .mnuWinASCII.Enabled = Not LrnMode And CBool(InstrCnt)
      .mnuFileListDir.Enabled = Not LrnMode And CBool(Len(StorePath))
      .mnuMemStk.Enabled = Not LrnMode
      .mnuFilePaste.Enabled = Not LrnMode
      .mnuFileImportSegment.Enabled = Not LrnMode
      
      If Not TextEntry Then .chkShift.Value = vbUnchecked 'turn off shift
      TextEntry = False
      CharLimit = 0
      RedoAlphaPad
      .chkShift.Enabled = False
      .lstDisplay.ToolTipText = vbNullString
      If LrnMode Then                 'if we just turned the LRN mode on...
        If ActivePgm <> 0 Then InstrPtr = 0
        ActivePgm = 0                 'set User program as active program
        .PicPlot.Visible = False      'ensure plot form is not visible
        Call RedoAlphaPad             'reset alpha pad
        LrnDsp = GetDisplayText()     'save display data
        With .lstDisplay
          LrnSav = .ListIndex         'save selected line
          LrnTop = .TopIndex          'save top displayed line
        End With
        Call BuildInstrList           'build instruction list
        .lblLoc.BackStyle = 1         'show learn mode headers
        .lblCode.BackStyle = 1
        .lblInstr.BackStyle = 1
        Call DspBackground
        
        If frmVisualCalc.mnuFileTglCoDisplay.Checked Then
          Call InitCoDisplay
        End If
      Else
        If frmCDLoaded Then
          Unload frmCoDisplay
        End If
        Erase FmtLst, FmtMap
        Call ResetListSupport
        LockControlRepaint .lstDisplay
        With .lstDisplay
          .Clear                      'clean out list
          
          If Not Preprocessd Then     'if no longer Preprocessed
            If DspLocked Then         'and display was locked
              LrnSav = InstrPtr
              Call Preprocess         'try to reCompress
              DspLocked = Preprocessd 'determine fate
              If Not Preprocessd Then
                LrnDsp = String$(DisplayWidth - 2, 32) & "0." & vbCrLf
                LrnTop = 0            'reset display
                LrnSav = 0
              Else
                .Clear                'ensure data is clear
                InstrPtr = LrnSav     'reset pointer
                Ary = BuildInstrArray()
                LrnDsp = Join(Ary, vbCrLf)
              End If
            ElseIf AutoPprc Then
              Call Preprocess         'try to reCompress
            End If
          End If
          
          Call DspBackground
          Ary = Split(LrnDsp, vbCrLf) 'get lines of data
          LrnDsp = vbNullString       'release resources
          .Clear
          i = UBound(Ary)
          Do While Not CBool(Len(Trim$(Ary(i)))) And i > LrnSav
            i = i - 1
          Loop
          For Idx = 0 To i
            .AddItem Ary(Idx)
          Next Idx
          Erase Ary                   'release resources
          .TopIndex = LrnTop          'reset top index
          If DspLocked And DspPgmList Then
            Call RepointIdx
          Else
            Call SelectOnly(LrnSav)   'select only this item
          End If
        End With
        UnlockControlRepaint .lstDisplay
        .lblLoc.BackStyle = 0         'hide learn mode headers
        .lblCode.BackStyle = 0
        .lblInstr.BackStyle = 0
        LastTypedInstr = 128          'reset last key
        Exit Sub
      End If
    End With
  End If
'
' handle other keys
'
  If LrnMode Then
    Call LrnKeypad                                    'if we are in the LRN mode
  Else
    Call ActiveKeypad                                 'if we have a 'hot' keyboard
    If IsNumbers And CBool(CharCount) Then            'if something numeric is typed...
      If CharCount = CharLimit And Not TextEntry Then 'and key input limit is met...
        Call CheckPnd(0)                              'force execution of pending commands
      End If
    End If
  End If
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

