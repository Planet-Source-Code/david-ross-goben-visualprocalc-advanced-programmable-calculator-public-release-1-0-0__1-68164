Attribute VB_Name = "modBuildPgmLine"
Option Explicit

Private Const Limit As Integer = 90 'instruction search limit for status bar code reflection
Private Acm As String               'status bar code reflection text accumulator
Public lIndLvl As Integer           'local indent level index for format co-display
Public ResetPnt As Boolean          'used when updating the CoDisplay form

'*******************************************************************************
' Subroutine Name   : AddPrv
' Purpose           : Append text with or without a space separator
'*******************************************************************************
Private Sub AddPrv(NewTxt As String, ByVal AddSpace As Boolean)
  If AddSpace Then                      'if we want a space separator...
    Acm = Acm & " " & NewTxt            'append data from higher line with space sepatator
  Else
    Acm = Acm & NewTxt                  'append data from higher line without a separator
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : BuildPgmLine
' Purpose           : Display command data up to instruction point in status bar
'                   : This features helps the programmer visualize the code by
'                   : also allowing them to see the code at least horizontally
'*******************************************************************************
Public Sub BuildPgmLine()
  Dim Code As Integer, Idx As Integer, PvCode As Integer, Lbnd As Integer
  Dim C As String
  '
  ' parse through the program up to the instruction pointer (active, selected line)
  '
  Code = 10000                      'init code out of bounds
  Acm = vbNullString                'init null accumulator
  
  Lbnd = InstrPtr - ((Limit \ 8) * 3) 'set lower bounds of checks
  If Lbnd < 0 Then Lbnd = 0
'-----------------------------------
  For Idx = Lbnd To InstrPtr
    If Idx = InstrCnt Then          'if we are actually beyond added code
      Exit For                      'done
    End If
    
    PvCode = Code                   'save previous code
    Code = Instructions(Idx)        'get new instruction
    '
    ' check alph and numeric strings
    '
    Select Case Code
      Case Is < 10                  'numeric?
        C = Chr$(Code + 48)         'yes, convert to ASCII
      Case Is < 128                 'other ASCII?
        C = Chr$(Code)              'convert it as well, as text
      Case Else
        C = GetInst(Code)           'otherwise, descramble the token to text format
    End Select
    
    Select Case Code
      Case Is < 10, iDot            'numeric?
        Select Case PvCode          'previous code numbers?
          Case Is < 10, iDot
            Acm = Acm & C           'accumulate them if so
            C = vbNullString        'nothing more to do
          Case Else
            Acm = Acm & " " & C     'otherwise, separate text form tokens
            C = vbNullString
        End Select
      Case Is < 128                 'ASCII?
        Select Case PvCode
          Case Is < 10              'ignore numeric
          Case Is < 128             'other alpha
            Acm = Acm & C           'so merge
            C = vbNullString
          Case Else
            Acm = Acm & " " & C     'else separate from tokens
            C = vbNullString
        End Select
    End Select
    '
    ' now parse the data for space separations or not
    '
    If CBool(Len(C)) Then
      Select Case Code                'check current code
        '÷, x, -, +, =, &, |, ^, %, >>, <<, <, >, +/-, (case) Else, Pi, e, <=, >=, ==, !=
        Case 171, 184, 197, 210, 223, 161, 174, 200, 213, 226, 354, 274, 275, 222, _
             iCaseElse, iPi, iEp, iEQ, iNEQ, iLE, iGE, iWhile, iUntil
          AddPrv C, True               'append current instruction(s) to the previous with no space sep.
        '(, ), ;, various special end parens
        Case iLparen, iRparen, iUparen, iIparen, iWparen, iFparen, iCparen, iEparen, iSparen, _
             iCCBrace, iSemiC, iSemiColon, iComma, iRbrkt, iColon, iDWparen, iLbrkt
          AddPrv C, False               'append current instruction(s) to the previous with no space sep.
        'Lnx, eX, X!, Int, Frac, Abs, Sgn, 1/X, LogX, Log, 10^, Root, Sqrt,
        'STO, RCL, EXC, SUM, MUL, SUB, DIV, X==T, X>=T, X>T, X!=T, X<=T, X<T
        '+/-, Nor, Else, Hyp, Arc, Sin, Cos, Tan, Sec, Csc, Cot, ]
        Case 149, 277, 152, 175, 303, 176, 304, 158, 286, 355, 356, _
             141, 142, 143, 144, 145, 272, 273, 164, 165, 166, 292, 293, 294, _
             222, iNor, iElse, iHyp, 155, 156, 157, iArc, 283, 284, 285, iRbrkt
          Select Case PvCode
            'special end braces (do not merge)
            Case iLCbrace, iLparen
              AddPrv C, False             'append current instruction(s) to the previous with no space sep.
            Case Else
              AddPrv C, True              'append current instruction(s) to the previous with a space sep.
          End Select
        Case iLCbrace '{
          Select Case PvCode
            Case iCparen, iCaseElse     'ignore if previous was ) or Else for Case
              AddPrv C, False
            Case Else
              AddPrv C, True              'append current instruction(s) to the previous with a space sep.
          End Select
        Case iDBG, iNOP
          AddPrv C, True
        Case Else
          Select Case PvCode         'else check previous instruction
            '(, ÷, x, -, +, ÷=, x=, -=, += ,^, !, %, ~, &, |, &&, >>, <<, \, <, >, various special end parens
            Case iLparen, 171, 184, 197, 210, 299, 312, 325, 338, 200, 315, 213, 187, 161, _
                 174, 289, 302, 226, 354, 341, 274, 275
              AddPrv C, False           'append current instruction(s) to it with no space sep.
            Case Else
              AddPrv C, True
          End Select
      End Select
    End If
    If Len(Acm) > Limit Then Acm = Right$(Acm, Limit)
  Next Idx
'-----------------------------------
  '
  ' all done parsing, so display what we have
  '
  SetTip Acm
End Sub

'*******************************************************************************
' Subroutine Name   : BldFmt
' Purpose           : Build Displayed format for co-display window
'*******************************************************************************
Public Sub BldFmt(Code As Integer)
  Dim PrvCode As Integer, Idx As Integer
  Dim BmpInd As Boolean
  Dim DecInd As Boolean
  Dim S As String
  
  FmtMap(FmtCnt) = FmtIdx                     'address for mapping
  If CBool(FmtIdx) Then                       'if not step 0
    PrvCode = Instructions(FmtIdx - 1)        'get previous instruction
  Else
    PrvCode = 1000                            'else set out of bounds
  End If
  
  If Code > 31 And Code < 128 Then            'current is ASCII?
    If PrvCode < 10 Or PrvCode > 127 Then     'and previous not?
      If PrvCode = iRem Or PrvCode = iRem2 Then 'allow remarks
        FmtLst(FmtCnt) = Chr$(Code)
      Else                                    'otherwise, init embracing text in quotes
        FmtLst(FmtCnt) = String$(lIndLvl * IndSpc, 32) & """" & Chr$(Code)
      End If
    Else
      FmtLst(FmtCnt) = Chr$(Code)
    End If
  Else
    If Code = iRem Then                       'make sure "REM' has trailing space
      FmtLst(FmtCnt) = String$(lIndLvl * IndSpc, 32) & "Rem "
    Else
      FmtLst(FmtCnt) = String$(lIndLvl * IndSpc, 32) & GetInst(Code)
    End If
    If PrvCode > 31 And PrvCode < 128 Then    'ASCII?
      Idx = FmtCnt - 1                        'yes, so search back for previous
      If Idx >= 0 Then                        'if still in bounds...
        Do While FmtMap(Idx) = -1             'search for previous entry
          Idx = Idx - 1
        Loop
        S = LTrim$(FmtLst(Idx))               'grab valid line
        If Left$(S, 1) <> "'" And Left$(S, 4) <> "Rem " Then  'remark?
          FmtLst(Idx) = FmtLst(Idx) & Chr$(34)  'no, so append quote on text
        End If
      End If
    End If
  End If
'
' For formatting, see if we can merge some lines
'
  Select Case Code                    'check current code
      'special end braces (do not merge)
    Case iBCBrace, iSCBrace, iICBrace, iDCBrace, iWCBrace, _
         iFCBrace, iCCBrace, iDWBrace, iDUBrace, iEIBrace, _
         iENBrace, iSTBrace, iSIBrace, iCNBrace, iRCbrace
      DecInd = True
    Case iLbl
      If PrvCode < 128 Then
        ApndFmt True, PrvCode
      End If
    Case 32 To 127                    'current is ASCII?
      Select Case PrvCode
        Case 32 To 127, iRem, iRem2   'previous is ASCII or remark?
          ApndFmt False, PrvCode      'append without added space
        Case Else
          ApndFmt True, PrvCode
      End Select
    Case 0 To 9, iDot                 'merge digits
      Select Case PrvCode
        Case 0 To 9, iDot
          ApndFmt False, PrvCode
        Case Else
          ApndFmt True, PrvCode
      End Select
    Case iDBG, iNOP, 128              'do nothing
    Case iWhile, iUntil
      Select Case PrvCode
        Case iDWBrace, iDUBrace
          ApndFmt True, PrvCode       'append current instruction(s) to the previous with space sep.
      End Select
    '÷, x, -, +, =, &, |, ^, %, >>, <<, <, >, +/-, (case) Else, Pi, e, <=, >=, ==, !=
    Case 171, 184, 197, 210, 223, 161, 174, 200, 213, 226, 354, 274, 275, 222, _
         iCaseElse, iPi, iEp, iEQ, iNEQ, iLE, iGE, iVar
      ApndFmt True, PrvCode           'append current instruction(s) to the previous with space sep.
    '(, ), ;, ], :, various special end parens
    Case iLparen, iRparen, iUparen, iIparen, iWparen, iFparen, iCparen, iEparen, iSparen, _
         iCCBrace, iSemiC, iSemiColon, iComma, iRbrkt, iColon, iDWparen, iLbrkt
      ApndFmt False, PrvCode          'append current instruction(s) to the previous with no space sep.
    'Lnx, eX, X!, Int, Frac, Abs, Sgn, 1/X, LogX, Log, 10^, Root, Sqrt,
    'STO, RCL, EXC, SUM, MUL, SUB, DIV, X==T, X>=T, X>T, X!=T, X<=T, X<T
    '+/-, Nor, Else, Hyp, Arc, Sin, Cos, Tan, Sec, Csc, Cot, ], OP, Val
    Case 149, 277, 152, 175, 303, 176, 304, 158, 286, 355, 356, _
         141, 142, 143, 144, 145, 272, 273, 164, 165, 166, 292, 293, 294, _
         222, iNor, iElse, iHyp, 155, 156, 157, iArc, 283, 284, 285, iRbrkt, iOP, iVal
      Select Case PrvCode
        'special end braces (do not merge)
        Case iBCBrace, iSCBrace, iICBrace, iDCBrace, iWCBrace, _
             iFCBrace, iCCBrace, iDWBrace, iDUBrace, iEIBrace, _
             iENBrace, iSTBrace, iSIBrace, iCNBrace, iRCbrace
        Case iDWparen, iUparen, iSemiC
        Case iLCbrace, iLparen
          ApndFmt False, PrvCode      'append current instruction(s) to the previous with no space sep.
        Case Else
          ApndFmt True, PrvCode       'append current instruction(s) to the previous with a space sep.
      End Select
    Case iLCbrace '{
      Select Case PrvCode
        Case iCparen, iCaseElse       'ignore if previous was ) or Else for Case
        Case Else
          ApndFmt True, PrvCode       'append current instruction(s) to the previous with a space sep.
          BmpInd = True
      End Select
    Case Else
      Select Case PrvCode         'else check previous instruction
        Case iLparen
          ApndFmt False, PrvCode
        '÷, x, -, +, ÷=, x=, -=, += ,^, !, %, ~, &, |, &&, >>, <<, \, <, >, various special end parens
        Case 171, 184, 197, 210, 299, 312, 325, 338, 200, 315, 213, 187, 161, _
             174, 289, 302, 226, 354, 341, 274, 275, iPlot
          ApndFmt True, PrvCode       'append current instruction(s) to it with space sep.
        'Nor, Dfn, Pvt, Pub, Else, [;], [,], LogX, Root, Sqrt, IND
        Case iNor, iDfn, iPvt, iPub, iElse, iSemiColon, iComma, 286, 355, 356, iIND
          ApndFmt True, PrvCode       'append current instruction(s) to it, with a space separator
      End Select
  End Select
'
' see if next line should be in-dented
'
  If BmpInd Then
    lIndLvl = lIndLvl + 1
    BmpInd = False
  End If
'
' see if next line should be out-dented
'
  If DecInd Then
    If CBool(lIndLvl) Then
      lIndLvl = lIndLvl - 1
    End If
    DecInd = False
  End If
'
' point to next formatted text location
'
  FmtCnt = FmtCnt + 1
  If FmtCnt > FmtSize Then
    FmtSize = FmtSize + 128
    ReDim Preserve FmtLst(FmtSize)
    ReDim Preserve FmtMap(FmtSize)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : ApndFmt
' Purpose           : Append the indicated line to the previous valid line
'*******************************************************************************
Public Sub ApndFmt(ByVal AddSpace As Boolean, ByVal PrvCode As Integer)
  Dim Inst As Integer, Code As Integer
  Dim Txt As String, Src As String, Tmp As String
  Dim Bol As Boolean
  
  If PrvCode = iNOP Then Exit Sub       'allow NOP to force blank separator lines
  If PrvCode = iSemiC Then Exit Sub
  If EnDefFlg Or StDefFlg Then Exit Sub 'if Enum or Struct being constructed, then do nothing
  
  Inst = FmtCnt                         'get local copy of instruction index
  If Inst = 0 Then Exit Sub             'if we cannot move backward
  
  Code = Instructions(Inst)
  If Code > 31 And Code < 128 And Not AddSpace Then
    Txt = FmtLst(Inst)                  'set text to append
  Else
    Txt = LTrim$(FmtLst(Inst))          'set text to append
  End If
  
  If PrvCode = iAdv And Txt <> ";" Then Exit Sub
  Inst = Inst - 1                       'back up 1
  Do While FmtMap(Inst) = -1            'while prev data already deleted
    Inst = Inst - 1                     'back up another line
    If Inst < 0 Then Exit Sub           'if we cannot back up anymore
  Loop
  
  If Code > 31 And Code < 128 Then
    Src = RTrim$(FmtLst(Inst))          'grab previous text, remove any trailing blanks
    If Src = vbNullString Then
      Src = Chr$(Code)
    End If
    Tmp = Src
  Else
    Src = RTrim$(FmtLst(Inst))          'grab previous text, remove any trailing blanks
    Tmp = LTrim$(Src)
  End If
'
' check for merging ASCII data
'
  Bol = False
  Select Case Code
    Case 32 To 127
      Select Case PrvCode
        Case 32 To 127, iRem, iRem2
          Bol = True                      'we are merging ASCII
      End Select
  End Select
'
' if not processing ASCII data...
'
  If Not Bol Then
    If Right$(Src, 1) = "{" Then
      If CaseIdx = 0 Then Exit Sub        'do not append of we have a right bracket and Case being defined
    End If
    If Left$(Tmp, 1) = "'" Or Left$(Tmp, 4) = "Rem " Then
      Exit Sub                            'do not append if previous was a remark
    End If
    If Right$(Src, 1) = "(" Or Right$(Src, 1) = "[" Then
      AddSpace = False
    End If
  End If
  
  If AddSpace Then                      'if we want a space separator...
    Tmp = Src & " " & Txt               'append data from higher line with space sepatator
  Else
    If FmtLst(Inst) = "'" Or FmtLst(Inst) = "Rem " And Code > 31 And Code < 128 Then
      Tmp = FmtLst(Inst) & Chr$(Code)
    Else
      Tmp = FmtLst(Inst) & Txt            'append data from higher line without a separator
    End If
  End If
  
  FmtLst(Inst) = Tmp                    'set updated data
  FmtMap(FmtCnt) = -1                   'and delete the current line
End Sub

'*******************************************************************************
' Subroutine Name   : InitCoDisplay
' Purpose           : Initialize co-display of source code
'*******************************************************************************
Public Sub InitCoDisplay()
'
' initialize for building new list
'
  FmtSize = SizeInst
  ReDim FmtLst(SizeInst)
  ReDim FmtMap(SizeInst)
  FmtCnt = 0
  lIndLvl = 0
'
' build new list
'
  For FmtIdx = 0 To InstrCnt - 1
    BldFmt Instructions(FmtIdx)
  Next FmtIdx
  FmtIdx = InstrPtr
'
' now load co-display form
'
  If Not frmCDLoaded Then
    frmCoDisplay.Show vbModeless, frmVisualCalc
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : SetUpCoDisplay
' Purpose           : Update list when already displays
'*******************************************************************************
Public Sub SetUpCoDisplay()
  Dim Idx As Integer, SelLn As Integer, i As Integer
  Dim S As String
  Dim Ary() As String
'
' initialize for building list
'
  FmtCnt = 0
  lIndLvl = 0
'
' build new list
'
  If InstrCnt = 0 Then Exit Sub 'prevent underflow
  For FmtIdx = 0 To InstrCnt - 1
    BldFmt Instructions(FmtIdx)
  Next FmtIdx
  FmtIdx = InstrPtr
'
' build list of all valid lines, and find select line
' using this method is MUCH faster than manipulating the listbox
'
  ReDim Ary(FmtCnt - 1)
  SelLn = 0                      'hold line that matches instruction pointer
  i = 0                          'valid line index
  For Idx = 0 To FmtCnt - 1
    If FmtMap(Idx) <> -1 Then    'valid line?
      Ary(i) = FmtLst(Idx)
      If FmtMap(Idx) <= InstrPtr Then SelLn = i 'set nearest match for pgm step
      i = i + 1                   'bump valid index
    End If
  Next Idx
'
' now update the listbox. Normally, one would clear the listbox, then add new
' data (While .Listcount / .RemoveItem .0 / Wend), but as fast as this is
' (several times faster than .Clear), the following technique is like lightening
'
  LockControlRepaint frmCoDisplay.lstSrc
  ResetPnt = True                 'prevent ping-pong refreshes (also speeds update)
  With frmCoDisplay.lstSrc
    Select Case .ListCount        'only address overhead in listbox
      Case Is < i
        Do While .ListCount < i   'add lines to listbox until it is full-sized
          .AddItem vbNullString
        Loop
      Case Is > i
        Do While .ListCount > i   'else remove excess lines until new full size
          .RemoveItem .ListCount - 1
        Loop
    End Select                    'or do nothing if count is already correct
'
' now add only lines that are different
'
    For Idx = 0 To FmtCnt - 1
      If Ary(Idx) <> .List(Idx) Then .List(Idx) = Ary(Idx)
    Next Idx
    Erase Ary                     'get rid of temp array
'
' now update display list pointers
'
    .ListIndex = SelLn            'set select line
    Idx = SelLn - .Height \ frmVisualCalc.lblChkSize.Height \ 2
    If Idx < 0 Then Idx = 0
    .TopIndex = Idx               'set top index to move select line down
    UnlockControlRepaint frmCoDisplay.lstSrc  'refresh display
    .Refresh
    ResetPnt = False              'again allow updates
  End With
  DoEvents                        'let screen catch up
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

