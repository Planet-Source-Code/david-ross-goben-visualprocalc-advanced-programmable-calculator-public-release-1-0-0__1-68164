Attribute VB_Name = "modProcessSupport"
Option Explicit
'-------------------------------------------------------------------------------
'Locally stored variables for use by this module
'-------------------------------------------------------------------------------
Dim HaveDigit As Boolean  'if the data is a digit string
Dim HaveAlpha As Boolean  'if the data is alphanumeric (begins with alpha)
Dim HaveDec As Boolean    'if a digit string has a decimal place
Dim HaveEE As Boolean     'if a digit string has an exponent entry
Dim HaveSign As Boolean   'if a digit string has a sign applied

'*******************************************************************************
' Function Name     : CheckForValue
' Purpose           : Check for a valid numeric expression entry
'                   : Success set the instruction pointer to the last character
'                   : of the accepted text (in preparation for the next).
'                   : TxtData contains the accepted data.
'                   : TstData contains the numeric conversion of the text.
'*******************************************************************************
Public Function CheckForValue(ByVal Inst As Integer) As Boolean
  Dim Code As Integer, j As Integer
  Dim S As String
  Dim Idx As Long
  
  HaveDigit = False                                     'we do not yet have a digit
  HaveDec = False                                       'we do not yet have a decimal
  HaveSign = False                                      'we do not yet have a sign
  TxtData = vbNullString                                'init accumulator
  HaveEE = False                                        'we do not yet have EE
  
  Do                                                    'being a parsing loop
    Inst = Inst + 1                                     'point to an op code
    Code = GetInstructionAt(Inst)                       'grab it
    Select Case Code
      Case -1
        If CBool(Len(TxtData)) Then Exit Do             'we have data
        Exit Function                                   'error (out of memory)
      Case 0 To 9                                       'digits
        HaveDigit = True                                'indicate we have digits
        TxtData = TxtData & Chr$(Code + 48)             'append accumulator
      Case iDot '[.]
        If HaveDec Then Exit Function                   'error if we already have decimal
        HaveDec = True                                  'else indicate that we have the decimal
        TxtData = TxtData & "."                         'append to accumulator
        HaveDigit = True                                'also assume digits
      Case 222  '+/-
        If HaveSign Or Not HaveDigit Then Exit Function 'if already have it, or no digits
        If HaveEE Then                                  'if we have EE, append the sign after E
          j = InStr(1, TxtData, "E")                    'find the "E" in the digit string
          TxtData = Left$(TxtData, j) & "-" & Mid$(TxtData, j + 1)
          HaveSign = True                               'indicate we have the sign
        Else
          TxtData = TxtData & " +/- "                   'else post-pend sign
          HaveSign = True                               'and mark having it
        End If
      Case 179  'EE
        If HaveEE Or Not HaveDigit Then Exit Function   'check for error conditions
        HaveEE = True                                   'indicate we have it
        HaveSign = False                                'reset sign for after E
        TxtData = TxtData & "E"                         'append E for EE
      Case Is < 128                                     'ASCII?
        If HaveDigit Or HaveDec Then Exit Function      'error if numerics (running into text is bad, anyway)
        If CheckForLabel(Inst - 1, LabelWidth) Then     'found a label there?
          If Preprocessing Then                         'if not Preprocessing
            TstData = 1                                 'assume all is well
            CheckForValue = True                        'indicate success
          Else
            Idx = FindLbl(TxtData, TypEnum)             'check for Enum
            If CBool(Idx) Then                          'found Enum
              If CBool(ActivePgm) Then
                TstData = CDbl(ModLbls(Idx).LblValue)   'get module enum value
              Else
                TstData = CDbl(Lbls(Idx).LblValue)      'get main program enum value
              End If
              CheckForValue = True                      'indicate success
            End If
          End If
        End If
        Exit Function                                   'anything else is bad news
      Case Else
        Exit Do                                         'otherwise, assume valid program instructions
    End Select
  Loop
  
  On Error Resume Next
  TstData = CDbl(TxtData)                               'try converting value
  Call CheckError
  On Error GoTo 0
  If ErrorFlag Then Exit Function                       'error

  InstrPtr = Inst - 1                                   'set master pointer to point after data-1
  CheckForValue = True                                  'indicate success
End Function

'*******************************************************************************
' Function Name     : CheckForLabel
' Purpose           : Check for valid label-type entry
'                   : Success set the instruction pointer to the last character
'                   : of the accepted text (in preparation for the next).
'                   : TxtData contains the accepted data.
'*******************************************************************************
Public Function CheckForLabel(ByVal Inst As Integer, ByVal Limit As Integer) As Boolean
  Dim Code As Integer
  Dim S As String, SCode As String
  
  HaveAlpha = False                                 'init no alpha
  TxtData = vbNullString                            'init accumulator
  
  Do
    Inst = Inst + 1                                 'point to an op code
    Code = GetInstructionAt(Inst)                   'get it
    Select Case Code
      Case -1
        If CBool(Len(TxtData)) Then Exit Do         'we have data
        Exit Function                               'error (out of memory)
      Case Is < 10                                  '0-9 (not ASCII)
        Exit Do                                     'assume non-text numerics
      Case Is < 128
        SCode = Chr$(Code)
        Select Case UCase$(SCode)                   'check string character
          Case "0" To "9"                           'digit
            If Not HaveAlpha Then Exit Function     'error
            TxtData = TxtData & SCode               'accumulate
            If Len(TxtData) = Limit Then
              Inst = Inst + 1                       'increment for later decrement
              Exit Do                               'auto-exit if we have limit
            End If
          Case "A" To "Z"                           'allow only A-Z and "_" and "-" and "."
            HaveAlpha = True                        'indicate we have a valid start
            TxtData = TxtData & SCode               'accumulate
            If Len(TxtData) = Limit Then
              Inst = Inst + 1                       'increment for later decrement
              Exit Do                               'auto-exit if we have limit
            End If
          Case "-", ".", "_"
            If Not HaveAlpha Then Exit Function     'must begin with A-Z
            TxtData = TxtData & SCode               'accumulate
            If Len(TxtData) = Limit Then
              Inst = Inst + 1                       'increment for later decrement
              Exit Do                               'auto-exit if we have limit
            End If
          Case Else
            Exit Function                           'all others are invalid
        End Select
      Case Else
        Exit Do                                     'else assume op code follows
    End Select
  Loop
  InstrPtr = Inst - 1                               'set master pointer to point after data-1
  CheckForLabel = True                              'all is OK
End Function

'*******************************************************************************
' Function Name     : CheckForValueOrLabel
' Purpose           : Allow input of a value expression or a label name
'                   : Success set the instruction pointer to the last character
'                   : of the accepted text (in preparation for the next).
'                   : TxtData contains the accepted data.
'*******************************************************************************
Public Function CheckForValueOrLabel(ByVal Inst As Integer, ByVal Limit As Integer) As Boolean
  Dim Code As Integer
  
  HaveAlpha = False                               'init no alpha
  HaveDigit = False                               'init no number
  Code = GetInstructionAt(Inst + 1)               'grab an op code
  
  Select Case Code
    Case -1
      Exit Function                               'error (out of memory)
    Case 0 To 9, iDot                             'digit or "."
      Limit = DisplayWidth                        'set limit to DisplayWidth if digit but no alpha
      HaveDigit = True                            'we have a value
    Case Is < 128                                 'check ASCII codes
      Select Case Chr$(Code)                      'check string character
        Case "A" To "Z", "_"                      'allow only A-Z and "_"
          HaveAlpha = True                        'indicate we have a valid start
        Case Else
          Exit Function                           'all others ASCII codes are invalid
      End Select
  End Select
'
' check for label or value. Ignore if a function, as processor will scan it
'
  TstData = -1
  
  If HaveAlpha Then
    CheckForValueOrLabel = CheckForLabel(Inst, Limit) 'text
  ElseIf HaveDigit Then
    CheckForValueOrLabel = CheckForValue(Inst)        'value
  Else
    TxtData = vbNullString
    CheckForValueOrLabel = True                       'function
  End If
End Function

'*******************************************************************************
' Function Name     : CheckForNumber
' Purpose           : Check for valid digit-only entry
'                   : Success set the instruction pointer to the last character
'                   : of the accepted text (in preparation for the next).
'                   : TxtData contains the accepted data.
'                   : TstData contains the numeric conversion of the text.
'*******************************************************************************
Public Function CheckForNumber(ByVal Inst As Integer, ByVal Limit As Integer, ByVal MaxVal As Long) As Boolean
  Dim Code As Integer, Iptr As Integer
  Dim S As String
  Dim Idx As Long
  
  HaveDigit = False                                 'init no digits
  TxtData = vbNullString                            'init accumulator
  Iptr = InstrPtr                                   'save instruction pointer
  
  Do
    Inst = Inst + 1                                 'point to an item
    Code = GetInstructionAt(Inst)                   'get an instruction
    Select Case Code
      Case -1
        If CBool(Len(TxtData)) Then Exit Do         'we have data
        Exit Function                               'error (out of memory)
      Case 0 To 9                                   'digit?
        HaveDigit = True                            'indicate we have it
        TxtData = TxtData & Chr$(Code + 48)         'append it
        If Len(TxtData) = Limit Then
          Inst = Inst + 1                           'increment for later decrement
          Exit Do                                   'auto-exit if we have limit
        End If
      Case Is < 128
        If HaveDigit Then Exit Do                   'if we have data, then OK
        If Code < 128 Then                          'ASCII?
          If CheckForLabel(Iptr, LabelWidth) Then   'found a label?
            If Preprocessing Then                   'if Preprocessing, we simply want to check for a label
              TstData = 1                           'give it something for a result (will be tossed)
              CheckForNumber = True                 'indicate all ok
              Exit Function
            End If
            Idx = FindLbl(TxtData, TypEnum)         'not preprocessing, so is found label an Enum?
            If CBool(Idx) Then                      'yes, grab long value stored in it
              If CBool(ActivePgm) Then
                Idx = CStr(ModLbls(Idx).LblValue)   'either from module
              Else
                Idx = CStr(Lbls(Idx).LblValue)      'or from main program memory
              End If
              If CBool(MaxVal) Then                 'check for limits
                If Idx > MaxVal Then Exit Function  'too high
              End If
              TstData = CDbl(Idx)                   'grab Dbl version
              CheckForNumber = True                 'indicate all ok
            End If
          End If
        End If
        Exit Function                               'leave, either in success or failure
      Case Else
        If HaveDigit Then Exit Do                   'if we have data, then OK
        Exit Function                               'else error
    End Select
  Loop
'-----
  TstData = CDbl(TxtData)                           'grab Dbl version
  If CBool(MaxVal) Then
    If CLng(TstData) > MaxVal Then Exit Function    'if value exceedes max value, then error
  End If
  InstrPtr = Inst - 1                               'set master pointer to point after data-1
  CheckForNumber = True                             'all ok
End Function

'*******************************************************************************
' Function Name     : CheckForAlNum
' Purpose           : Check for valid Variable type entry
'                   : Success set the instruction pointer to the last character
'                   : of the accepted text (in preparation for the next).
'                   : TxtData contains the accepted data.
'                   : TstData contains the numeric conversion of the text. -1 if not.
'*******************************************************************************
Public Function CheckForAlNum(ByVal Inst As Integer, ByVal Limit As Integer) As Boolean
  Dim Code As Integer, HldInst As Integer
  Dim Idx As Long
  Dim Pool As Labels
  
  HaveDigit = False                                 'init no digits
  HaveAlpha = False                                 'init no alpha
  TxtData = vbNullString                            'init accumulator
  TstData = -1
  
  Do
    Inst = Inst + 1
    Code = GetInstructionAt(Inst)                   'grab an op code
    Select Case Code
      Case -1
        If CBool(Len(TxtData)) Then Exit Do         'we have data
        Exit Function                               'error (out of memory)
      
      Case 0 To 9                                   'digit
        If HaveAlpha Then Exit Do                   'if was alpha data, this is something new, so OK
        If Not HaveAlpha And Not HaveDigit Then     'if no alpha data entered yet...
          Limit = 2                                 'set limit to 2 if digit but no alpha
          HaveDigit = True                          'indicate we have a value here
        End If
        TxtData = TxtData & Chr$(Code + 48)         'add to accumulator
        If Len(TxtData) = Limit Then
          Inst = Inst + 1                           'increment for later decrement
          Exit Do                                   'auto-exit if we have limit
        End If
      
      Case Is < 128 'check ASCII codes
        If HaveDigit Then Exit Do                   'if we began with digit, then was value, so done
        Select Case UCase$(Chr$(Code))              'check string character
          Case "A" To "Z", "_"                      'allow only A-Z and "_"
            HaveAlpha = True                        'indicate we have a valid start
            TxtData = TxtData & Chr$(Code)          'accumulate
            If Len(TxtData) = Limit Then
              Inst = Inst + 1                       'increment for later decrement
              Exit Do                               'auto-exit if we have limit
            End If
          Case "-", ".", "0" To "9"                 '[-]. [.], or 0-9
            If Not HaveAlpha Then Exit Function     'must begin with A-Z or "_"
            TxtData = TxtData & Chr$(Code)          'accumulate
            If Len(TxtData) = Limit Then
              Inst = Inst + 1                       'increment for later decrement
              Exit Do                               'auto-exit if we have limit
            End If
          Case Else
            Exit Function                           'all other ASCII codes are invalid
        End Select
      Case Else
        Exit Do                                     'else assume op code follows
    End Select
  Loop
    
  If HaveDigit Then
    TstData = CDbl(TxtData)                         'get Dbl version if numeric
  Else                                              'else may be Const or Enum
    If Not Preprocessing Then                       'if Preprocessing, assume all OK
      Idx = FindLbl(TxtData, TypEnum)               'not preprocessing, so is found label an Enum?
      If CBool(Idx) Then                            'yes, grab long value stored in it
        If CBool(ActivePgm) Then
          Idx = CStr(ModLbls(Idx).LblValue)         'either from module
        Else
          Idx = CStr(Lbls(Idx).LblValue)            'or from main program memory
        End If
        TstData = CDbl(Idx)                         'grab Dbl version
      End If
    End If
  End If
  
  InstrPtr = Inst - 1                               'set master pointer to point at end of data
  CheckForAlNum = True                              'all is OK
End Function

'*******************************************************************************
' Function Name     : CheckForText
' Purpose           : Check for valid text entry
'                   : Success set the instruction pointer to the last character
'                   : of the accepted text (in preparation for the next).
'                   : TxtData contains the accepted data.
'*******************************************************************************
Public Function CheckForText(ByVal Inst As Integer, ByVal Limit As Integer) As Boolean
  Dim Code As Integer
  Dim S As String
  Dim HaveText As Boolean
  Dim Idx As Long
  
  TxtData = vbNullString                            'init accumulator
  HaveText = False
  Do
    Inst = Inst + 1                                 'point to an op code
    Code = GetInstructionAt(Inst)                   'grab it
    Select Case Code
      Case -1
        If CBool(Len(TxtData)) Then Exit Do         'we have data
        Exit Function                               'error (out of memory)
      Case 0 To 9                                   'digit
        Exit Do
      Case Is < 128
        HaveText = True
        TxtData = TxtData & Chr$(Code)              'absorb any ASCII
        If Len(TxtData) > Limit Then Exit Function
      Case Else
        Exit Do                                     'assume valid opcode
    End Select
  Loop
  If Not HaveText Then Exit Function                'error if no text data
  InstrPtr = Inst - 1                               'else set master pointer to point after data-1
  '
  ' see if Constant used
  '
  If Len(TxtData) <= LabelWidth Then                'if LabelWidth or less, see if constant name
    If Not Preprocessing Then                       'ignore if Preprocessing
      Idx = FindLbl(TxtData, TypConst)              'check for a constant
      If CBool(Idx) Then
        If CBool(ActivePgm) Then
          TxtData = ModLbls(Idx).lblCmt             'grab module constant text
        Else
          TxtData = Lbls(Idx).lblCmt                'else grab main memory constant text
        End If
        If Len(TxtData) > Limit Then Exit Function  'cannot exceed limit
      End If
    End If
  End If
  CheckForText = True                               'all is OK
End Function

'*******************************************************************************
' Function Name     : ApndTxt2
' Purpose           : Support public subs ApndInst() and ApndTxt()
'                   : This determines how expression text is combined.
'*******************************************************************************
Private Function ApndTxt2(ByVal Txt1 As String, ByVal Txt2 As String, Optional Quoted As Boolean = False) As String
  Dim Bol As Boolean
  
  If Quoted Then Txt2 = EnQuoteTxt(Txt2)  'surround Txt2 with Double quotes
  
  If Not CBool(Len(Txt1)) Then
    ApndTxt2 = Txt2
    Exit Function
  End If
  
  Select Case Right$(Txt1, 1)     'check right end of left text for special characters
    Case "(", "{", "[", "]"
      Bol = True
    Case Else
      Bol = False
  End Select
  
  If Not Bol And Not Quoted Then
    Select Case Left$(Txt2, 1)    'check left end of right text for special characters
      Case ")", "}", "[", "]"
        Bol = True
      Case Else
        Bol = False
    End Select
  End If
  
  If Bol Then                     'if special characters encountered...
    ApndTxt2 = Txt1 & Txt2        'then simply combine the text
  Else
    ApndTxt2 = Txt1 & " " & Txt2  'else separate them with a space
  End If
End Function

'*******************************************************************************
' Subroutine Name   : ApndInst
' Purpose           : Append a single instruction to the formatted string
'*******************************************************************************
Public Sub ApndInst(ByVal Ofst As Integer)
  If CBool(Ofst) Then InstrPtr = InstrPtr + Ofst
  InstTxt = ApndTxt2(InstTxt, GetInst(Instructions(InstrPtr)))
End Sub

'*******************************************************************************
' Subroutine Name   : ApndTxt
' Purpose           : Append InstTxt to the formatted string
'*******************************************************************************
Public Sub ApndTxt()
  If CBool(Len(TxtData)) Then
    InstTxt = ApndTxt2(InstTxt, TxtData, Not IsNumeric(TxtData))
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : ApndQTx
' Purpose           : Append InstTxt to the formatted string
'*******************************************************************************
Public Sub ApndQTx()
  If CBool(Len(TxtData)) Then
    InstTxt = ApndTxt2(InstTxt, TxtData, True)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : ApndPrv
' Purpose           : Append the indicated line to the previous valid line
'*******************************************************************************
Public Sub ApndPrv(ByVal AddSpace As Boolean)
  Dim Inst As Integer
  Dim Txt As String, Src As String, Tmp As String
  
  If PrvCode = iNOP Then Exit Sub       'allow NOP to force blank separator lines
  If PrvCode = iSemiC Then Exit Sub
  If EnDefFlg Or StDefFlg Then Exit Sub 'if Enum or Struct being constructed, then do nothing
  
  Inst = InstCnt                        'get local copy of instruction index
  If Inst = 0 Then Exit Sub             'if we cannot move backward
  Txt = LTrim$(InstFmt(Inst))           'set text to append (remove any added indenting)
  If PrvCode = iAdv And Txt <> ";" Then Exit Sub
  Inst = Inst - 1                       'back up 1
  Do While InstMap(Inst) = -1           'while prev data already deleted
    Inst = Inst - 1                     'back up another line
    If Inst < 0 Then Exit Sub           'if we cannot back up anymore
  Loop
  
  Src = RTrim$(InstFmt(Inst))           'grab previous text, remove any trailing blanks
  Tmp = LTrim$(Src)
  
  If Right$(Src, 1) = "{" Then
    If CaseIdx = 0 Then Exit Sub        'do not append of we have a right bracket and Case being defined
  End If
  
  If Left$(Tmp, 1) = "'" Or Left$(Tmp, 4) = "Rem " Then
    Exit Sub                            'do not append if previous was a remark
  End If
  
  If AddSpace Then                      'if we want a space separator...
    Tmp = Src & " " & Txt               'append data from higher line with space sepatator
  Else
    Tmp = InstFmt(Inst) & Txt           'append data from higher line without a separator
  End If
  
  If Len(Tmp) <= DisplayWidth Then      'if display will not extend beyond display
    InstFmt(Inst) = Tmp                 'set updated data
    InstMap(InstCnt) = -1               'and delete the current line
    HaveDels = True                     'indicate we have deletions
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : CloseFmt
' Purpose           : Close up the Format pool by removing any deleted lines
'*******************************************************************************
Public Sub CloseFmt()
  Dim Idx As Integer, Iptr As Integer
  
  If HaveDels Then
    Iptr = 0
    For Idx = 0 To InstCnt - 1
        If InstMap(Idx) <> -1 Then      'if current is not deleted...
          InstFmt(Iptr) = InstFmt(Idx)  'shift instructions down
          InstMap(Iptr) = InstMap(Idx)  'shift mapping down
          Iptr = Iptr + 1               'bump valid mapping index
        End If
    Next Idx
    InstCnt = Iptr                      'set new ubound of pool+1
    HaveDels = False                    'reset flag
  End If
'
' even if no deletions, shrink size of pools (they are NORMALLY much smaller)
'
  ReDim Preserve InstFmt(InstCnt - 1) 'trim size of pools
  ReDim Preserve InstMap(InstCnt - 1)
End Sub

'*******************************************************************************
' Subroutine Name   : DefineLables
' Purpose           : Define an object to the Lbls() pool (used only by Pgm 00)
'*******************************************************************************
Public Sub DefineLables(ByVal Typ As LblTypes, ByVal Iptr As Integer, Optional KeyCode As Integer = 0)
  Dim i As Integer, j As Integer
  
  i = KeyCode                               'get index into labels (used only for Ukeys by SYSTEM)
  If i = 0 Then i = LblCnt                  'append if no system index provided
  With Lbls(i)
    If KeyCode = 0 Then                     'if new, then check for definition
      j = FindLblMatch(TxtData)             'search for matching name
    Else
      j = .LblDat                           'else check if data address set
    End If
    If CBool(j) Then                        'Attempting redefinition?
      ForcError "Label '" & TxtData & "' has already been defined"
      Exit Sub                              'return error state
    End If
    If Typ = TypKey Then
      .lblName = TxtData                    'all OK, so assign the name as the keypad label
    Else
      .lblName = UCase$(TxtData)            'all OK, so assign the name (uppercase)
    End If
    If Typ = TypSbr Then
      If StrComp(TxtData, "MAIN", vbTextCompare) = 0 Then
        HaveMain = i                        'indicate index of Main() routine
      End If
    End If
    .LblTyp = Typ                           'assign the type of label
    If Typ = TypKey Then                    'Ukey?
      .LblScope = Pub                       'Userkeys are by default public
    Else
      .LblScope = Pvt                       'else assume private
    End If
    If CBool(Iptr) Then                     'if not at start of program
      Select Case Instructions(Iptr - 1)    'check token before SBR, LBL, UKEY, etc.
        Case iPub
          .LblScope = Pub
        Case iPvt
          .LblScope = Pvt
      End Select
    End If
    .lblAddr = Iptr                         'definition of label
    .LblDat = InstrPtr                      'start of actual data block
    .lblUdef = True                         'mark as user-defined
    If Typ = TypSbr Or Typ = TypKey Then
      .LblEnd = FindEbrace()                'store end brace address
    End If
  End With
'
' added definition, now bump indexer
'
  If KeyCode = 0 Then                       'if user did not supply a key
    Call BumpLblCnt                         'bump size of label pool
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : ResetBracing
' Purpose           : Reset bracing in the program
'*******************************************************************************
Public Sub ResetBracing()
  Dim Idx As Integer
  
  For Idx = 0 To InstrCnt - 1
    Select Case Instructions(Idx)
      'convert special end braces to normal
      Case iBCBrace, iSCBrace, iICBrace, iDCBrace, iWCBrace, _
           iFCBrace, iCCBrace, iDWBrace, iDUBrace, iEIBrace, _
           iENBrace, iSTBrace, iSIBrace, iCNBrace
        Instructions(Idx) = iRCbrace
      'conver special parens to normal
      Case iUparen, iIparen, iWparen, iFparen, iCparen, iEparen, iSparen, iDWparen
        Instructions(Idx) = iRparen
      'convert special Else for Case back to normal
      Case iCaseElse
        Instructions(Idx) = iElse
      'convert ';' in FOR statements
      Case iSemiColon
        Instructions(Idx) = iSemiC
    End Select
  Next Idx
End Sub

'*******************************************************************************
' Subroutine Name   : ResetFmt
' Purpose           : Clean out Formatted Text buffer
'*******************************************************************************
Public Sub ResetFmt()
  InstCnt = 0
  Erase InstFmt, InstMap, InstFmt3, InstMap3
End Sub

'*******************************************************************************
' Function Name     : FindLblMatch
' Purpose           : Search for a matching label name in the Lbls() list
'*******************************************************************************
Public Function FindLblMatch(Txt As String) As Long
  Dim Idx As Long
  Dim Test As String * LabelWidth
  
  Test = Txt                          'make sure we are upper case
  If CBool(ActivePgm) Then
    For Idx = ModLblMap(ActivePgm - 1) + 1 To ModLblMap(ActivePgm) - 1
      If StrComp(ModLbls(Idx).lblName, Test, vbTextCompare) = 0 Then
        FindLblMatch = Idx          'return the index
        Exit Function
      End If
    Next Idx
  Else                                'else scan Module's Lbls() pool
    For Idx = 1 To LblCnt - 1         'scan all defined entries
      If StrComp(Lbls(Idx).lblName, Test, vbTextCompare) = 0 Then
        FindLblMatch = Idx            'yes, return its index
        Exit Function
      End If
    Next Idx
  End If
  FindLblMatch = 0                    'else indicate it was not found (0=new label)
End Function

'*******************************************************************************
' Function Name     : GetDimValue
' Purpose           : Extract a dimension value
'*******************************************************************************
Public Function GetDimValue(ErrorStr As String) As Integer
  Dim DM As Integer
  
  DM = -1
  ApndInst 1                                  'append assumed '['
  If CheckForNumber(InstrPtr, 2, 99) Then     'if valid value
    DM = CInt(TxtData)                        'get dim value
    ApndTxt                                   'append data
    If Instructions(InstrPtr + 1) = iRbrkt Then
      ApndInst 1                              'add ']'
    Else
      ErrorStr = "Expected ']'"
      DM = -1
    End If
  Else
    ErrorStr = "Invalid array definition"
  End If
  GetDimValue = DM
End Function

'*******************************************************************************
' Function Name     : CheckDims
' Purpose           : Check for variable dimensioning
'*******************************************************************************
Public Function CheckDims(ByVal Vr As Long, ByRef Dm1 As Long, _
                          ByRef Dm2 As Long, ErrorStr As String) As Boolean
  Dm1 = -1
  Dm2 = -1
  If Instructions(InstrPtr + 1) = iLbrkt Then     'found a dim baracket?
    Dm1 = GetDimValue(ErrorStr)                   'init 1st dim
    If Dm1 <> -1 Then
      If Instructions(InstrPtr + 1) = iLbrkt Then 'found another baracket?
        Dm2 = GetDimValue(ErrorStr)               'init 2nd dim
        If Dm2 = -1 Then Dm1 = -1                 'if error on 2nd, reflect to 1st
      End If
    End If
  Else
    Exit Function
  End If
  
  If Dm1 = -1 Then ErrorStr = "Invalid array definition"
  
  CheckDims = Not CBool(Len(ErrorStr))
End Function

'*******************************************************************************
' Subroutine Name   : BumpLblCnt
' Purpose           : Increase the size of the Label pool
'*******************************************************************************
Public Sub BumpLblCnt()
  LblCnt = LblCnt + 1             'added 1 to list
  If LblCnt > LblSize Then        'outside bounds?
    LblSize = LblSize + DefInc    'yes, so bump pool by increment
    ReDim Preserve Lbls(LblSize)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : CheckLrnDim
' Purpose           : Check Dimensioning
'*******************************************************************************
Public Sub CheckLrnDim(ErrorStr As String)
  ApndInst 1                                'add [
  If Instructions(InstrPtr + 1) = iIND Then 'check for IND
    ApndInst 1                              'add IND
  End If
  If Instructions(InstrPtr + 1) = iVar Then 'check for Var
    ApndInst 1                              'add Var
  End If
  If Not CheckForAlNum(InstrPtr, LabelWidth) Then 'check for label or 0-99
    ErrorStr = "Invalid array reference"
  Else
    ApndTxt                                 'append text
    If Instructions(InstrPtr + 1) <> iRbrkt Then
      ErrorStr = "Ending bracket ']' expected"
    Else
      ApndInst 1                            'add ]
    End If
  End If
End Sub
  
'*******************************************************************************
' Subroutine Name   : CheckLrnDim2
' Purpose           : Check for possible 1 or 2-D array designation
'*******************************************************************************
Public Sub CheckLrnDim2(ErrorStr As String)
  If Instructions(InstrPtr + 1) = iLbrkt Then     'bracket for variable
    Call CheckLrnDim(ErrorStr)                    'check Dim reference
    If Len(ErrorStr) = 0 Then                     'if it was OK
      If Instructions(InstrPtr + 1) = iLbrkt Then 'a second bracket?
        Call CheckLrnDim(ErrorStr)                'check Dim reference
      End If
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : CheckForVar
' Purpose           : Check for expected variable, but also check for IND and
'                   : for possible 1 or 2-D array designation
'*******************************************************************************
Public Function CheckForVar(EStr As String) As Boolean
  Dim ErrorStr As String
  
  If Instructions(InstrPtr + 1) = iIND Then ApndInst 1  'skip IND
  If Instructions(InstrPtr + 1) = iVar Then ApndInst 1  'skip Var
  If CheckForAlNum(InstrPtr, LabelWidth) Then
    ApndTxt                                             'append variable name/number
    If Instructions(InstrPtr + 1) = iLbrkt Then         'bracket for dimensioning
      Call CheckLrnDim2(ErrorStr)                       'check Dim references
    ElseIf Instructions(InstrPtr + 1) = iDot Then       'dot for structure name
      ApndInst 1                                        'allow dot
      If CheckForLabel(InstrPtr, LabelWidth) Then
        ApndTxt                                         'allow label
      Else
        ErrorStr = "Structure member error"
      End If
    End If
  Else
    ErrorStr = "Invalid Variable Definition"
  End If
  If CBool(Len(ErrorStr)) Then
    EStr = ErrorStr
  Else
    CheckForVar = True
  End If
End Function

'*******************************************************************************
' Subroutine Name   : ResetListSupport
' Purpose           : Disable lists
'*******************************************************************************
Public Sub ResetListSupport()
  With frmVisualCalc
    .mnuWinVar.Enabled = False
    .mnuWinLbl.Enabled = False
    .mnuWinUkey.Enabled = False
    .mnuWinSbr.Enabled = False
    .mnuWinConst.Enabled = False
    .mnuWinStruct.Enabled = False
  End With
End Sub

'*******************************************************************************
' Function Name     : CheckForXY
' Purpose           : Check for (x,y) format
'*******************************************************************************
Public Function CheckForXY(ErrorStr As String) As Boolean
  Dim i As Integer
  
  ErrorStr = "Invalid parameter"
  CheckForXY = False                              'return FALSE if ERROR
  
  If Instructions(InstrPtr + 1) <> iLparen Then Exit Function
  ApndInst 1                                      'add (
  i = FindEPar(1)
  If i = -1 Then
    ErrorStr = vbNullString 'error was already reported
    Exit Function
  End If
  Instructions(i - 1) = iEparen                   'mark end
  If Instructions(InstrPtr + 1) = iIND Then       'variable
    If Not CheckForVar(ErrorStr) Then Exit Function
  ElseIf Not CheckForNumber(InstrPtr, 3, PlotWidth) Then
    Exit Function
  End If
  ApndTxt                                         'append X variable or value
  If Instructions(InstrPtr + 1) <> iComma Then Exit Function
  ApndInst 1                                      'add [,]
  If Instructions(InstrPtr + 1) = iIND Then       'variable
    If Not CheckForVar(ErrorStr) Then Exit Function
  ElseIf Not CheckForNumber(InstrPtr, 3, PlotHeight) Then
    Exit Function
  End If
  ApndTxt                                         'append Y variable or value
  If Instructions(InstrPtr + 1) <> iEparen Then Exit Function
  ApndInst 1                                      'add ')'
  ErrorStr = vbNullString                         'remove error status
  CheckForXY = True                               'we are OK
End Function

'*******************************************************************************
' Function Name     : CheckCommaValue
' Purpose           : Check for a comma, followed by a value
'*******************************************************************************
Public Function CheckCommaValue(ErrorStr As String) As Boolean
  ErrorStr = "Invalid parameter"
  CheckCommaValue = True                      'return TRUE if ERROR
  If Instructions(InstrPtr + 1) = iComma Then
    ApndInst 1                                'add [,]
    If CheckForValue(InstrPtr) Then
      ApndTxt                                 'add value
      ErrorStr = vbNullString                 'remove error status
      CheckCommaValue = False                 'we are OK
    End If
  End If
End Function

'*******************************************************************************
' Subroutine Name   : ChkCompVbl
' Purpose           : Check for expected variable
'*******************************************************************************
Public Function ChkCompVbl() As String
  Dim Vn As Integer
  
  If Instructions(InstrPtr + 1) = iVar Then IncInstrPtr 'skip Var definition
  Call CheckForAlNum(InstrPtr, LabelWidth)              'grab number
  If TstData < 0# Then                                  'if we have a label...
    Vn = FindVblMatch(TxtData)                          'find a match in the variables list
    If Vn = -1 Then                                     'error?
      ForcError "Invalid Variable name"                 'yes, so report and exit
      Exit Function
    End If
  Else
    Vn = CInt(TstData)                                  'else grab variable number from Dbl
  End If
  ChkCompVbl = CStr(Vn)                                 'return variable number as a string
End Function

'*******************************************************************************
' Function Name     : EnQuoteTxt
' Purpose           : Embrace text within quotes. COnvert internal quotes to double
'*******************************************************************************
Public Function EnQuoteTxt(Txt As String) As String
  Dim S As String, T As String
  Dim Idx As Long
  
  If Len(Txt) = 0 Then Exit Function    'if nothing to actually enquote
  T = Txt                               'copy text, in case of no quotes
  Idx = InStr(1, T, """")               'find an imbedded quote
  If CBool(Idx) Then
    S = vbNullString                    'init accumulator
    Do While CBool(Idx)
      S = S & Left$(T, Idx - 1) & """""" 'capture and append 2 quotes
      T = Mid$(T, Idx + 1)
      Idx = InStr(1, T, """")
    Loop
    If CBool(Len(T)) Then S = S & T     'append any trailing data
  Else
    S = Txt                             'no embedded quotes, so copy source data
  End If
  
  EnQuoteTxt = """" & S & """"          'embrace final result within quotes
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

