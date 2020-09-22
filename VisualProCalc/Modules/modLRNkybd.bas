Attribute VB_Name = "modLRNkybd"
Option Explicit

'*******************************************************************************
' Subroutine Name   : LrnKeypad
' Purpose           : Handle LRN Mode Keyboard
'*******************************************************************************
Public Sub LrnKeypad()
  Dim dwColor As Long
  
  Select Case LastTypedInstr
'-------------------------------------------------------------------------------
    Case 128  '2nd    'Ignore
'-------------------------------------------------------------------------------
    Case iStyle  'Style  'Ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case LRNKey 'LRN    'handled by main keyboard processor
'-------------------------------------------------------------------------------
    Case 363, 364 'Pcmp, Comp 'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 137  'INS
      INSmode = Not INSmode
      Call UpdateStatus
'-------------------------------------------------------------------------------
    Case 265  'DEL
      Call DeleteInstruction
'-------------------------------------------------------------------------------
    Case 138  'Cut    'cut selected instructions to an internal buffer
      Call CutInstruction
'-------------------------------------------------------------------------------
    Case 139  'Copy   'copy selected instructions to an internal buffer
      Call CopyInstruction
'-------------------------------------------------------------------------------
     Case 266  'Paste
      Call PasteInstruction
'-------------------------------------------------------------------------------
    Case 136  'SST
      If InstrPtr < InstrCnt Then
        Call frmVisualCalc.Form_KeyDown(39, 0)
      End If
'-------------------------------------------------------------------------------
    Case 235  'NxLbl
      Call NxtPrvLabel(True)
'-------------------------------------------------------------------------------
    Case 236  'PvLbl
      Call NxtPrvLabel(False)
'-------------------------------------------------------------------------------
    Case iList  'List 'Ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 264  'BST
      Call LrnBST                                 'rull back up instruction pointer
'-------------------------------------------------------------------------------
    Case iLCbrace                                 'user typed left curly brace?
      Call AddInstruction(LastTypedInstr)         'first add self
      If ConDefFlg Or _
         (Instructions(InstrCnt - 2) = iLCbrace And Instructions(InstrCnt - 2) = 361) Or _
         Instructions(InstrCnt - 2) = iRCbrace Then 'defining a constant or possible Enum?
        ConDefFlg = False                         'yes, so first turn flag off
        If Not TextEntry Then                     'if not in text entry mode
          DspTxt = vbNullString                   'initialize pending text
          CharCount = 0                           'init character counter
          CharLimit = DisplayWidth                'set character limit
          AllowSpace = False                      'do not allow typing of a space
          Call frmVisualCalc.checkTextEntry(True) 'set up keyboard
        End If
      End If
'-------------------------------------------------------------------------------
    Case iPrint, iPrintx
      Call AddInstruction(LastTypedInstr)       'first add self
      DspTxt = vbNullString                     'initialize pending text
      CharCount = 0                             'init character counter
      CharLimit = DisplayWidth                  'set character limit
      AllowSpace = True                         'allow typing of a space
      Call frmVisualCalc.checkTextEntry(True)   'set up keyboard
'-------------------------------------------------------------------------------
    Case iTXT  'Txt  'allow text or numeric input
      If TextEntry Then                         'turn text mode off...
        Call ResetPnd
        AllowSpace = False                      'do not allow typing of a space
        REMmode = 0
        Call frmVisualCalc.checkTextEntry(False) 'set up keyboard
      Else                                      'turn text mode on...
        DspTxt = vbNullString                   'initialize pending text
        CharCount = 0                           'init character counter
        CharLimit = DisplayWidth                'set character limit
        AllowSpace = True                       'allow typing of a space
        Call frmVisualCalc.checkTextEntry(True) 'set up keyboard
        HaveTxt = True                          'force ASCII data
      End If
'-------------------------------------------------------------------------------
    Case iRem, iRem2 'Rem or [']
      Call AddInstruction(LastTypedInstr)       'first add self
      AllowSpace = True                         'allow typing of a space
      Call frmVisualCalc.checkTextEntry(True)   'set up keyboard
      DspTxt = vbNullString                     'initialize pending text
      CharCount = 0                             'init character counter
      If LastTypedInstr = iRem Then
        REMmode = 1
        CharLimit = DisplayWidth - 4            'set character limit
      Else
        REMmode = 2
        CharLimit = DisplayWidth - 1            'set character limit
      End If
'-------------------------------------------------------------------------------
    'Sbr, Lbl, Call, Gto, Const, Struct, AdrOf, Incr, Decr, iPrintf
    Case iSbr, iLbl, iCall, iGTO, iConst, iStruct, iAdrOf, iIncr, iDecr, iPrintf
      REMmode = 0
      If LastTypedInstr = iConst Then ConDefFlg = True 'we are defining a constant
      Call AddInstruction(LastTypedInstr)       'first add self
      DspTxt = vbNullString                     'initialize pending text
      CharCount = 0                             'init character counter
      CharLimit = LabelWidth                    'set character limit
      AllowSpace = False                        'do not allow typing of a space
      Call frmVisualCalc.checkTextEntry(True)   'set up keyboard
'-------------------------------------------------------------------------------
    Case iPmt, iFmt
      REMmode = 0
      Call AddInstruction(LastTypedInstr)       'first add self
      DspTxt = vbNullString                     'initialize pending text
      CharCount = 0                             'init character counter
      CharLimit = DisplayWidth                  'set character limit
      AllowSpace = True                         'allow typing of a space
      Call frmVisualCalc.checkTextEntry(True)   'set up keyboard
      HaveTxt = True                            'force ASCII data
'-------------------------------------------------------------------------------
    Case iComma
      Call AddInstruction(LastTypedInstr)       'first add self
      If InstrPtr > 2 Then
        Select Case Instructions(InstrPtr - 2)
          Case 48 To 122
            DspTxt = vbNullString                     'initialize pending text
            CharCount = 0                             'init character counter
            CharLimit = DisplayWidth                  'set character limit
            AllowSpace = True                         'do not allow typing of a space
            Call frmVisualCalc.checkTextEntry(True)   'set up keyboard
        End Select
      End If
'-------------------------------------------------------------------------------
    Case iUkey                                  'Ukey
      Call AddInstruction(LastTypedInstr)       'first add self
      DspTxt = vbNullString                     'initialize pending text
      CharCount = 0                             'init character counter
      CharLimit = 1                             'set character limit
      AllowSpace = False                        'do not allow typing of a space
      Call frmVisualCalc.checkTextEntry(True)   'set up keyboard
      Upcase = True
'-------------------------------------------------------------------------------
    'Nvar, Tvar, Ivar, Cvar
    Case iNvar, iTvar, iIvar, iCvar
      Call AddInstruction(LastTypedInstr)       'first add self
      DspTxt = vbNullString                     'initialize pending text
      CharCount = 0                             'init character counter
      VarLbl = False
      CharLimit = 2                            'set character limit
      AllowSpace = False                        'do not allow typing of a space
      If InstrCnt > 1 Then                     'if defining at start of program
        If Instructions(InstrCnt - 2) = iLCbrace Then 'if a structure, we will expect a label
          CharLimit = LabelWidth                'set character limit
          Call frmVisualCalc.checkTextEntry(True)   'set up keyboard
          VarLbl = True                         'allow text title, or 2-digit variable number
          Exit Sub
        End If
      End If
      Call frmVisualCalc.checkTextEntry(True)   'set up keyboard
'-------------------------------------------------------------------------------
    'Var, STO, RCL, EXC, SUM, MUL, SUB, DIV, IND, Trim, LTrim, RTrim, With
    Case iVar, iSTO, iRCL, iEXC, iSUM, iMUL, iSUB, iDIV, iIND, iTrim, iLTrim, iRTrim, iWith
      Call AddInstruction(LastTypedInstr)       'first add self
      DspTxt = vbNullString                     'initialize pending text
      CharCount = 0                             'init character counter
      CharLimit = LabelWidth                    'set character limit
      AllowSpace = False                        'do not allow typing of a space
      Call frmVisualCalc.checkTextEntry(True)   'set up keyboard
      VarLbl = True                             'allow text title, or 2-digit variable number
'-------------------------------------------------------------------------------
    'Digits
    Case 0 To 9
      If CharCount = 0 And REMmode = 0 And VarLbl Then 'if nothing has been typed yet
        Call frmVisualCalc.checkTextEntry(False)
        CharLimit = 2                           'assume number, so limit entry to 2 digits
      ElseIf CharCount = CharLimit And CBool(CharLimit) Then 'if we are at limit, force end of text entry
        LastTypedInstr = 128
      ElseIf CBool(REMmode) Or (TextEntry And HaveTxt) Or VarLbl Then 'if we are entering text, convert number to text
        LastTypedInstr = LastTypedInstr + 48    'apply ASCII code for "0"
      End If
      AddInstruction LastTypedInstr             'and add instruction to list
'-------------------------------------------------------------------------------
    Case iSelect  'check for RGB Select
      If InstrPtr > 0 Then
        If Instructions(InstrPtr - 1) = 339 Then        'RGB precedes it?
          If GetColor(frmVisualCalc, dwColor) Then      'if color selected
            Call AddInstruction(iLparen)                'add (
            Call BreakUpVal(dwColor And &HFF)           'add red
            Call AddInstruction(iComma)                 'add ,
            Call BreakUpVal((dwColor \ 256) And &HFF)   'add green
            Call AddInstruction(iComma)                 'add ,
            Call BreakUpVal((dwColor \ 65536) And &HFF) 'add blue
            LastTypedInstr = iRparen                    'add )
          Else
            LastTypedInstr = 128                        'force ignore
          End If
        End If
      End If
      AddInstruction LastTypedInstr                     'add instruction to list
      
'-------------------------------------------------------------------------------
    Case Else
      AddInstruction LastTypedInstr             'else simply add instruction to list
'-------------------------------------------------------------------------------
  End Select
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

