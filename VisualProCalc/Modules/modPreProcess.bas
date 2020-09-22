Attribute VB_Name = "modPreProcess"

Option Explicit

'*******************************************************************************
' This is the program Preprocessor, which just ensures that syntax is not weird
' It also gathers up gathers up numeric and text constants (but does nothing with
' then -- the Compressr will do that), and gathers programs transfers such as
' Lbl, Sbr, and Ukey definitions, as well as user-declared constants and structures.
'*******************************************************************************

'*******************************************************************************
' Subroutine Name   : Preprocess
' Purpose           : Performs Quickie Compress and build instruction format
'*******************************************************************************
Public Sub Preprocess()
  Dim HldIptr As Integer        'hold Instruction Pointer while it is modified
  Dim Errstr As String          'error reporting
  Dim BmpInd As Boolean         'flag indicating that indenting must be bumped
  Dim DecInd As Boolean         'indicates if indent needs to drop back
  Dim IndLvl As Integer         'indent level (also Brace Level)
  Dim KeyName As String         'used to hold user-key reference name
  Dim KeyLbl As String          'used to hold user-key label, displayed on key
  Dim KeyTip As String          'tooltip to be assigned to user key
  Dim Vn As Long                'variable number
  Dim X As Long, Y As Long
  Dim Ln As Integer             'len value
  Dim Nm As String              'name value
  Dim i As Integer, j As Integer, K As Integer, L As Integer
  Dim Typ As Vtypes
  Dim Bol As Boolean
  Dim HDisplayReg As Double, HTestReg As Double
  Dim HInstrPtr As Integer
'--------------------------------
  If CBool(ActivePgm) Then InstrPtr = 0
  HInstrPtr = InstrPtr          'hold values
  HDisplayReg = DisplayReg
  HTestReg = TestReg
  HaveMain = 0                  'no main routine present
  
  ActivePgm = 0                 'force active pgm to the user program (0)
  If Preprocessd Then Exit Sub  'we do not need to Preprocess (already done)
  Preprocessd = False           'else make sure flag is turned off
  Compressd = False              'also turn off full Compress, since we are Preprocessing
  If InstrCnt = 0 Then Exit Sub 'if nothing to Compress, then simply exit
  
  Call Cmn_CP_Support           'mini CP
  
  HaveDels = False              'init flag to indicate format text deletions must be done
  ErrorFlag = False             'turn off error flag
  InstrErr = 0                  'clear last encountered error location
  Errstr = vbNullString         'turn off error string
  
  ReDim InstFmt(InstrCnt)       'init formatted text pool
  ReDim InstMap(InstrCnt)       'init mapping pool
  ReDim InstFmt3(InstrCnt)
  ReDim InstMap3(InstrCnt)
  InstCnt = 0                   'init formatted data counter
  
  SbrDefFlg = False             'turn off Sbr/Ukey definition
  UtlDefFlg = False             'turn off Unitl definition
  ForDefflg = False             'turn off For flags
  ForIdx = 0                    '
  WhiDefFlg = False             'turn off While flags
  DoWhiDefFlg = False
  WhiIdx = 0
  DoIdx = 0
  SelDefFlg = False             'turn off Select flags
  SelIdx = 0
  CaseDefFlg = False            'turn off Case flags
  CaseIdx = 0
  IfDefFlg = False              'turn off If flags
  IfIdx = 0
  StDefFlg = False              'turn off Struct definition
  
  EnDefFlg = False              'turn off Enum definition
  EnumIdx = 0                   'reset enum index
  ConDefFlg = False             'turn off constant definition
  
  IndLvl = 0                    'init indent level for formatted text
  BmpInd = False                'no indenting yet
  DecInd = False
  
  Call ResetBracing             'reset bracing in the program
  InstrPtr = 0                  'init to base
  
  Call RenewLabels              'reset user-defined labels
  Call ResetListSupport
  Preprocessing = True          'indicate we are Preprocessing
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'    Debug.Assert False            'stop the train until we get this thing working
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  Do While InstrPtr < InstrCnt    'Compress while the instruction pointer can find data
    If CBool(InstrPtr) Then       'if not at base...
      PrvCode = Instructions(InstrPtr - 1) 'save previous code
    Else
      PrvCode = -1                'else no previous code
    End If
    Code = Instructions(InstrPtr) 'grab code
    
    InstMap(InstCnt) = InstrPtr   'begin line map
    InstMap3(InstCnt) = InstrPtr
    InstTxt = GetInst(Code)       'init formatted text
'-------------------------------------------------------------------------------
    Select Case Code              'check it

' 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, [.]---------------------------------------------
      Case 0 To 9, iDot
        If CheckForValue(InstrPtr - 1) Then
          InstTxt = TxtData       'set data to instruction text
        Else
          Errstr = "Invalid value"
        End If
' Ascii Text--------------------------------------------------------------------
      Case Is < 128
        If Not CheckForText(InstrPtr - 1, DisplayWidth) Then
          Errstr = "Invalid text usage"
        Else
          InstTxt = EnQuoteTxt(TxtData) 'set data to instruction text
        End If

'-------------------------------------------------------------------------------
'      case 128  '2nd key  'ignored
'      Case 129  ' LRN      'ignored

' Pgm --------------------------------------------------------------------------
      Case 130
        If Instructions(InstrPtr + 1) = iIND Then
          Call CheckForVar(Errstr)                                'check for variable, IND, and arrays
        Else
          If CheckForNumber(InstrPtr, 2, 99) Then                 'else get absolute number
            ApndTxt
          Else
            Errstr = "Invalid program number"
            Exit Do
          End If
        End If
        Select Case Instructions(InstrPtr + 1)                    'see what this is followed by
          Case iCall, Is > 900                                    'ensure Call or UserKey applied
          Case Else
            Errstr = "Invoking module programs requires User-key or Call"
        End Select

' Load, Save, Lapp, ASCII ------------------------------------------------------
      Case iLoad, iSave, iLapp, iASCII
        If CheckForNumber(InstrPtr, 2, 99) Then
          ApndTxt
        Else
          Errstr = "Bad program number"
        End If

'-------------------------------------------------------------------------------
'      Case 133  ' CE   'no need to check (these are auto-added to InstFmt()
'      Case 134  ' CLR  'no need to check

' Op ---------------------------------------------------------------------------
      Case 135
        If Instructions(InstrPtr + 1) = iIND Then
          Call CheckForVar(Errstr)                                'check for variable, IND, and arrays
        Else
          If CheckForNumber(InstrPtr, 2, MaxOps) Then
            ApndTxt
          Else
            Errstr = "Invalid Operation number"
          End If
        End If

'-------------------------------------------------------------------------------
'      Case 136  ' SST  'ignored
'      Case 137  ' INS  'ignored
'      Case 138  ' Cut  'ignored
'      Case 139  ' Copy 'ignored

'-------------------------------------------------------------------------------
'      Case 140  'PtoR 'no need to check

' STO, RCL, EXC, SUM, SUB, MUL, DIV, INC, DEC, TRIM, LTRIM, RTRIM, Var ----------
' Incr, Decr
      Case 141, 142, 143, 144, 145, 272, 273, 329, 330, 309, 310, 311, 287
        Call CheckForVar(Errstr)                      'check for var, Indirection, and 1/2D arrays

' IND --------------------------------------------------------------------------
      Case 146
        Errstr = "Unexpected instruction"

'-------------------------------------------------------------------------------
'      Case 147  ' Reset  'no need to check

'Hkey, Skey --------------------------------------------------------------------
      Case 148, 276
        HldIptr = InstrPtr                            'save instruction pointer
        If CheckForText(InstrPtr, DisplayWidth) Then  'check for text listing
          ApndQTx
          For i = 1 To Len(TxtData)                   'ensure text is valid
            Select Case Mid$(TxtData, i, 1)
              Case "A" To "Z", ",", "-"               'allow A-Z, [,], and [-]
              Case Else
                Errstr = "Invaid text character"
                InstrPtr = HldIptr                    'reset point for error index
                Exit For
            End Select
          Next i
        End If

'-------------------------------------------------------------------------------
'      Case 149  ' lnX  'no need to check
'      Case 150  ' E+   'no need to check
'      Case 151  ' Mean 'no need to check
'      Case 152  ' X!   'no need to check
'      Case 153  ' X><T 'no need to check

' Hyp --------------------------------------------------------------------------
      Case iHyp
        Select Case Instructions(InstrPtr + 1)
          Case iArc  'Arc
            ApndInst 1  'append Arc
            Select Case Instructions(InstrPtr + 1)  'for Hyp Arc fns
              'Sin, Cos, Tan, Sec, Csc, Cot
              Case 155, 156, 157, 283, 284, 285
                ApndInst 1
            End Select
          'Sin, Cos, Tan, Sec, Csc, Cot             'for Hyp fns
          Case 155, 156, 157, 283, 284, 285
            ApndInst 1
          Case Else
            Errstr = "Unexpected instruction"
        End Select

' Arc ---------------------------------------------------------------------------
      Case iArc
        Select Case Instructions(InstrPtr + 1)
          Case 155, 156, 157, 283, 284, 285 'Sin, Cos, Tan, Sec, Csc, Cot
            ApndInst 1
          Case Else
            Errstr = "Unexpected instruction"
        End Select

'--------------------------------------------------------------------------------
'      Case 155  ' Sin  'no need to check
'      Case 156  ' Cos  'no need to check
'      Case 157  ' Tan  'no need to check
'      Case 283  ' Sec  'no need to check
'      Case 284  ' Csc  'no need to check
'      Case 285  ' Cot  'no need to check
'      Case 158  ' 1/X  'no need to check
'      Case 159  'Txt  'This instruction is nevered entered into the LRN mode
'      Case 160  ' Hex  'no need to check
'      Case 161  ' &    'no need to check

' StFlg, RFlg ------------------------------------------------------------------
      Case 162, 290
        If CheckForNumber(InstrPtr, 1, MaxOps) Then
          ApndTxt
        Else
          Errstr = "Invalid flag number"
        End If

' IfFlg, !Flg ------------------------------------------------------------------
      Case 163, 291
        If CheckForNumber(InstrPtr, 1, 9) Then
          ApndTxt
          If Instructions(InstrPtr + 1) = iLCbrace Then
            ApndInst 1
            BmpInd = True   'we should bump indent on next line
            i = FindEbrace  'find matching end brace
            If i < 0 Then
              Errstr = "Cannot find a matching ending brace '}'"
            Else
              Instructions(i) = iICBrace
              IfIdx = IfIdx + 1 'bump depth of embedded Ifs
            End If
          Else
            Errstr = "Opening brace '{' expected"
          End If
        Else
          Errstr = "Invalid flag number"
        End If


' X==T, X>=T, X>T, X!=T, X<=T, X<T----------------------------------------------
      Case 164, 165, 166, 292, 293, 294
        If Instructions(InstrPtr + 1) = iLCbrace Then
          ApndInst 1
          BmpInd = True
          i = FindEbrace  'find matching end brace
          If i < 0 Then
            Errstr = "Cannot find a matching ending brace '}'"
          Else
            Instructions(i) = iICBrace
            IfIdx = IfIdx + 1 'bump depth of embedded Ifs
          End If
        Else
          Errstr = "Opening brace '{' expected"
        End If

' Dfn --------------------------------------------------------------------------
      Case 167
        Select Case Instructions(InstrPtr + 1)
          'Pvt, Pub, Lbl, Sbr, Ukey, Const, Struct, Enum, Ary, Nvar, Tvar, Ivar, Cvar
          Case 232, 360, 193, 180, 206, 233, 234, 361, 212, 225, 340, 353
          Case Else
            ForcError "Invalid Dfn parameter"
        End Select

' [:] --------------------------------------------------------------------------
      Case iColon '[:]
          Errstr = "Unexpected Instruction"

'-------------------------------------------------------------------------------
      '    ÷, x, -, +, ÷=, x=, -=, +=, &, |, ~, ^, %, &&, ||, !, Nor
      '    y^, Root, <, <=, >, >=, ==, !=, LogX. \
      Case 171, 184, 197, 210, 299, 312, 325, 338, 161, 174, 187, 200, _
           213, 289, 302, 315, 328, 227, 355, 274, 327, 275, 314, _
           288, 301, 286, 341
        
        InstFmt(InstCnt) = String$(IndLvl * IndSpc, 32) & InstTxt     'stuff formatted line
        InstFmt3(InstCnt) = InstFmt(InstCnt)
        'Debug.Print InstFmt(InstCnt)
        ApndPrv True                      'append current to previous with space
        InstCnt = InstCnt + 1             'point to next formatted text location
        InstMap(InstCnt) = InstrPtr + 1
        InstMap3(InstCnt) = InstrPtr + 1
        InstTxt = vbNullString
        If CheckForValueOrLabel(InstrPtr, LabelWidth) Then
          If TstData = -1 Then
            ApndQTx                       'assume text data
          Else
            ApndTxt                       'else append value/label
          End If
        Else
          Errstr = "Unexpected Instruction"
        End If
        
' [(] --------------------------------------------------------------------------
      Case iLparen  ' ( 'find matching ending paren and replace with New end paren mark,
                        'to allow us to track matching parentheses.
        i = FindEPar(0)                           'find matching end paren
        If i < 0 Then
          Errstr = "Cannot find a matching end paren ')'"
        Else
          Instructions(i - 1) = iEparen           'replace ')' with special NEW paren
        End If
        
' [)] --------------------------------------------------------------------------
      Case iRparen  'this token should never be encountered 'freely, because '(' and all other tokens that
                    'employ an opening paren will convert this token into a special new one for tracking
                    'and error trapping purposes.
        Errstr = "Ending parentheses ')' without a matching '('"
      
'      Case iEparen  ' ) special replacement paren (this replaces normal paren so we can be sure
                     '   that we have all matches working. This is ignored, because it is the expected
                     '   token for normal parenthesized expressions.

      Case iUparen   ' ) for Do...Until definition
        UtlDefFlg = False     'turn off Until def
      
      Case iDWparen
        DoWhiDefFlg = False   'turn off DO...While def
        
      ' ) special case parens expecting a left brace to follow
      Case iIparen, iWparen, iFparen, iSparen, iCparen, iCaseElse
        Select Case Code
          
          Case iIparen        ' ) for If definition
            IfDefFlg = False  'turn off expression definition flag
            j = iICBrace      'braced terminator
            IfIdx = IfIdx + 1 'bump depth of embedded Ifs
          
          Case iWparen        ' ) for While definition
            WhiDefFlg = False 'turn off While definition flag
            If Not DoWhiDefFlg Then 'if not Do...While flag...
              j = iWCBrace    'then assume While block
              WhiIdx = WhiIdx + 1
            Else
              Instructions(InstrPtr) = iDWparen 'ensure marked as Do-While() end paren
            End If
          
          Case iFparen        ' ) for For Definition
            ForDefflg = False
            j = iFCBrace
            ForIdx = ForIdx + 1
          
          Case iSparen        ' For Select definition
            SelDefFlg = False
            j = iSCBrace
            SelIdx = SelIdx + 1
          
          Case iCparen, iCaseElse '[)] or [Else] for Case definition
            CaseDefFlg = False
            j = iCCBrace
            CaseIdx = CaseIdx + 1 'broaden depth of cases
            
        End Select
        '
        ' we will expect that "{" will follow, execpt in the case of a Do...While()
        '
        If DoWhiDefFlg Then                               'if Do...While active
          DoWhiDefFlg = False                             'turn off defs
        ElseIf Instructions(InstrPtr + 1) = iLCbrace Then 'if def for/while/if/select/case, look for "{"
          BmpInd = True                                   'indicate that following lines will indent
          Select Case Code
            Case iCparen, iCaseElse                       'do nothing more for case [)] and [Case Else]
            Case Else
              ApndInst 1                                    'add the brace
              i = FindEbrace()                              'find matching end brace
              If i < 0 Then
                Errstr = "Cannot find a matching ending brace '}'"
              Else
                Instructions(i) = j                         'stuff appropriate end brace token
              End If
          End Select
        Else
          Errstr = "Opening brace '{' expected"
        End If
        
'-------------------------------------------------------------------------------
'      Case iStyle  ' Style 'Ignored
'      Case 173  ' Dec  'no need to check
'      Case 174  ' |    'no need to check
'      Case 175  ' Int  'no need to check
'      Case 176  ' Abs  'no need to check

' Fix --------------------------------------------------------------------------
      Case 177
        If Instructions(InstrPtr + 1) = iIND Then
          Call CheckForVar(Errstr)                        'check for variable, IND, and arrays
        ElseIf Instructions(InstrPtr + 1) = iIND Then     'allow Fix EE to enable Eng Mode
          ApndInst 1
        ElseIf CheckForNumber(InstrPtr, 1, 9) Then        'not IND, so check 0-9
          ApndTxt
        Else
          Errstr = "Invalid parameter"
        End If

'-------------------------------------------------------------------------------
'      Case 178  ' D.MS 'no need to check
'      Case 179  ' EE   'no need to check

' Sbr --------------------------------------------------------------------------
      Case 180
        j = InstrPtr                                      'save start of definition
        If SbrDefFlg Then                                 'subroutine active...
          Errstr = "Subroutine definition already active"
        Else
          If CheckForLabel(InstrPtr, LabelWidth) Then     'verify label found
            If Len(TxtData) = 1 Then
              Errstr = "You cannot define subroutines names of 1 character"
            Else
              ApndQTx                                       'found, so add
              If Instructions(InstrPtr + 1) = iLCbrace Then  'found '{'?
                ApndInst 1                                  'then add brace to instruction list
                Call DefineLables(TypSbr, j)                'yes, so first declare label
                BmpInd = True
                SbrDefFlg = True                            'tag subroutine active
                i = Lbls(LblCnt - 1).LblEnd                 'get matching end brace address
                If i < 0 Then
                  Errstr = "Cannot find a matching ending brace '}'"
                Else
                  Instructions(i) = iBCBrace                'force returning end brace
                  frmVisualCalc.mnuWinSbr.Enabled = True
                End If
              Else
                Errstr = "Opening brace '{' expected"
              End If
            End If
          Else
            Errstr = "Invaid Named definition"
          End If
        End If

' Rem --------------------------------------------------------------------------
      Case iRem
        If CheckForText(InstrPtr, DisplayWidth - 4) Then
'          ApndTxt
          InstTxt = InstTxt & " " & TxtData 'do not add quotes
        Else
          Errstr = "Parameter error"
        End If

' ['] --------------------------------------------------------------------------
      Case iRem2
        If CheckForText(InstrPtr, DisplayWidth - 1) Then
          InstTxt = InstTxt & TxtData 'do not add quotes or separating space
        Else
          Errstr = "Parameter error"
        End If

'-------------------------------------------------------------------------------
'      Case 186  ' Oct  'no need to check

' Select -----------------------------------------------------------------------
      Case 188
        If Instructions(InstrPtr + 1) <> iLCbrace Then            'Select {? (Using T-Reg for tests)
          If Instructions(InstrPtr + 1) = iLparen Then            'left paren?
            ApndInst 1                                            'append "("
            i = FindEPar(0)                                       'find end paren
            If i < 0 Then
              Errstr = "Cannot find a matching end paren ')'"
            Else
              Instructions(i - 1) = iSparen                       'replace ')' with special CASE paren
              SelDefFlg = True
            End If
          Else
            Errstr = "Expected '('"
          End If
        Else
          Instructions(InstrPtr) = iSelectT               'set tag for easy reference
          ApndInst 1                                      'append '{'
          BmpInd = True                                   'we will bump indent starting on next line
          SelIdx = SelIdx + 1                             'bump Select depth flag
          i = FindEbrace()                                'now find a matching end brace
          If i < 0 Then
            Errstr = "Cannot find a matching end brace '}'"
          Else
            Instructions(i) = iSCBrace                    'special Select Case block terminator
          End If
        End If
        
' Case -------------------------------------------------------------------------
      Case 189
        If CBool(SelIdx) Then                           'if we are within a select block...
          Select Case Instructions(InstrPtr - 1)        'look at previous code
            Case iLCbrace, iCCBrace                     'only items that can precede "Case"
            Case Else
              Errstr = "Invalid Case definition"
              Exit Do
          End Select
          If Instructions(InstrPtr + 1) = iLparen Then  'Case (?
            ApndInst 1                                  'yes, so add "("
            CaseDefFlg = True                           'mark defining case
            i = FindEPar(0)                             'find end paren
            If i < 0 Then
              Errstr = "Cannot find a matching end paren ')'"
            Else
              Instructions(i - 1) = iCparen             'replace ')' with special CASE paren
            End If
          ElseIf Instructions(InstrPtr + 1) = iElse Then  'allow Case Else {
            Instructions(InstrPtr + 1) = iCaseElse        'stuff new Else command for Case
            CaseDefFlg = True
          Else
            Errstr = "Expected '('"
          End If
        Else
          Errstr = "No Select block to contain Case"
        End If

' [{] --------------------------------------------------------------------------
      Case iLCbrace '190
        If Not StDefFlg And Not EnDefFlg And CaseIdx <> 1 Then     'allowed only with Struct and Enum Items
          Errstr = "Opening brace '{' without a vaid instruction associated with it"
        Else
          j = InstrPtr                              'save address for definition
          i = FindEbrace()                          'locate ending brace for item
          If i < 0 Then                             'if not found, then it is an error
            Errstr = "Cannot find a matching end brace '}'"
          Else
            If EnDefFlg Then  'Enum Item Definition
              Instructions(i) = iEIBrace            'mark end of Enum Item block
              If Not CheckForLabel(InstrPtr, LabelWidth) Then 'get name to define in enum
                Errstr = "Invalid Enum name"
              Else
                If CBool(FindLblMatch(TxtData)) Then 'search for matching label
                  ForcError "Label '" & TxtData & "' has already been defined"
                Else
                  Nm = TxtData                      'save name
                  ApndQTx                           'append name
                  If Instructions(InstrPtr + 1) <> iEIBrace Then
                    Errstr = "Closing brace '}' expected"
                  Else                              'else all is OK
                    ApndInst 1                      'append end brace
                    Select Case Instructions(InstrPtr + 1)
                      Case iLCbrace, iENBrace       'allow only '{' or Enum '}" to follow
                      Case Else
                        Errstr = "Enumerations can only contain Enum Item definitions"
                        Exit Do
                    End Select
                    
                    With Lbls(LblCnt)
                      .lblName = UCase$(Nm)         'apply name for Enum item
                      .LblTyp = TypEnum             'define as enumerator
                      .LblScope = Pvt               'always private
                      .LblValue = EnumIdx           'apply current enumeration value
                      .lblAddr = j                  'set definition address
                      .lblUdef = True               'mark as user-defined
                    End With
                    
                    EnumIdx = EnumIdx + 1           'bump for next enum entry
                    Call BumpLblCnt                 'bump size of label pool
                  End If
                End If
              End If
            ElseIf StDefFlg Then
              Instructions(i) = iSIBrace            'mark end of Structure Item block
              Select Case Instructions(InstrPtr + 1)
                Case 212  'Nvar
                  Typ = vNumber
                  i = 8               'size in bytes
                Case 225  'Tvar
                  Typ = vString
                  i = DisplayWidth    'size in bytes
                Case 340  'Ivar
                  Typ = vInteger
                  i = 4               'size in bytes
                Case 353  'Cvar
                  Typ = vChar
                  i = 1               'size in bytes
                Case Else
                  Errstr = "Expected Struture Item type definition"
                  Exit Do
              End Select
              ApndInst 1                            'append value type
              
              If Not CheckForLabel(InstrPtr, LabelWidth) Then 'get name to define in Structure
                Errstr = "Invalid Structure Item name"
              Else
                If CBool(FindLblMatch(TxtData)) Then 'search for matching label
                  ForcError "Label '" & TxtData & "' has already been defined"
                Else
                  Nm = TxtData                      'save name
                  ApndQTx                           'append name
                  
                  Ln = i                            'init with actual length of data
                  If Typ = vString And Instructions(InstrPtr + 1) = iLen Then
                    ApndInst 1                      'add LEN
                    If Not CheckForNumber(InstrPtr, 2, DisplayWidth) Then
                      Errstr = "Invaid Len definition"
                      Exit Do
                    End If
                    ApndTxt                         'add data
                    Ln = CInt(TxtData)              'user-defined length for string
                  End If
                  
                  If Instructions(InstrPtr + 1) <> iSIBrace Then
                    Errstr = "Closing brace '}' expected"
                  Else                              'else all is OK
                    ApndInst 1                      'add }
                    Select Case Instructions(InstrPtr + 1)
                      Case iLCbrace, iSTBrace       'allow only '{' or Struct '}" to follow
                      Case Else
                        Errstr = "Structures can only contain Structure Item definitions"
                        Exit Do
                    End Select
                    
                    With StructPl(StructCnt)        'point to latest structure
                      ReDim Preserve .StItems(.StItmCnt)  'add an item
                      L = .StSiz                    'get copy of current size
                      .StSiz = .StSiz + i           'bump size off I/O buffer
                      With .StItems(.StItmCnt)
                        .SiName = Nm                'set name to it
                        .siType = Typ               'and type
                        .siLen = Ln                 'and length (if String, actual will still be DisplayWidth)
                        .siOfst = L                 '0ffset within I/O buffer
                      End With
                      .StItmCnt = .StItmCnt + 1     'set new structure item count
                    End With
                  End If
                End If
              End If
            Else
              Instructions(i) = iCCBrace          'stuff end brace token for case
            End If
          End If
        End If

' [}] --------------------------------------------------------------------------
      Case iRCbrace '191
        Errstr = "Unexpected end brace '}' encountered" '(it should ALWAYS be converted before being found)
      
      Case iDWBrace  '} for Do...While block
        DoIdx = DoIdx - 1                           'back up definition for the main Do block
        DecInd = True                               'we should drop indent back on next line
        
      Case iDUBrace  '} for DO...Until block
        DoIdx = DoIdx - 1
        DecInd = True                               'we should drop indent back on next line
      
      Case iICBrace  '} for If blocks
        IfIdx = IfIdx - 1
        DecInd = True                               'we should drop indent back on next line
      
      Case iDCBrace  '} for Do Block
        DoIdx = DoIdx - 1
        DecInd = True                               'we should drop indent back on next line

      Case iWCBrace  '} for While block
        WhiIdx = WhiIdx - 1
        DecInd = True                               'we should drop indent back on next line
      
      Case iFCBrace  '} for For block
        ForIdx = ForIdx - 1
        DecInd = True                               'we should drop indent back on next line
      
      Case iBCBrace  '} for Sbr/Ukey block
        SbrDefFlg = False                           'just turn flag off
        DecInd = True                               'we should drop indent back on next line
      
      Case iSCBrace  '} for Select block
        SelIdx = SelIdx - 1                         'just decrement index
        DecInd = True                               'we should drop indent back on next line
        
      Case iCCBrace  '} for Case block
        CaseIdx = CaseIdx - 1
        Select Case Instructions(InstrPtr + 1)
          Case iCase, iSCBrace                        'case block followed by Case or Select '}'
            DecInd = True                             'we should drop indent back on next line
          Case Else
            Errstr = "Case block followed by invalid instruction"
        End Select
       
      Case iSTBrace '} for Struct block
        With StructPl(StructCnt)                    'point to latest structure
          .StBuf = String$(.StSiz, 0)               'make buffer size of data
        End With
        StDefFlg = False                            'turn off structure definition flag
        DecInd = True                               'we should drop indent back on next line
        
      Case iSIBrace ') for struct item
        Errstr = "Unexpected end brace '}' encountered" '(it should ALWAYS be converted before being found)
      
      Case iENBrace ') for Enum block
        EnDefFlg = False
        DecInd = True                               'we should drop indent back on next line
      
      Case iEIBrace ') for Enum item
        Errstr = "Unexpected end brace '}' encountered" '(it should ALWAYS be converted before being found)
      
'-------------------------------------------------------------------------------
'      Case 192  ' Deg  'no need to check

' Lbl --------------------------------------------------------------------------
      Case 193
        i = InstrPtr                                    'save instruction pointer
        If CheckForLabel(InstrPtr, LabelWidth) Then     'verify label found
          ApndQTx                                       'found, so add
          If Instructions(InstrPtr + 1) = iColon Then   'found ':'?
            ApndInst 1                                  'yes, then add colon to instruction list
            Call DefineLables(TypLbl, i)                'then declare label
            frmVisualCalc.mnuWinLbl.Enabled = True
          Else
            Errstr = "Invalid Named Definition"
          End If
        Else
          Errstr = "Invaid Named definition"
        End If

'-------------------------------------------------------------------------------
'      Case 198  ' Beep 'no need to check
'      Case 199  ' Bin  'no need to check

' For --------------------------------------------------------------------------
      Case 201
        If Instructions(InstrPtr + 1) = iLparen Then      'expected '('?
          i = FindEPar(1)                                 'find matching paren
          If i < 0 Then
            Errstr = "Cannot find a matching end paren ')'"
          Else
            Instructions(i - 1) = iFparen 'replace ')' with special FOR paren
            Call FindForInfo(j, K, L)                     'get For def pointers
            If j = -1 Then                                'if Init data is invalid
              Errstr = "Invalid For definition"
            Else
              ApndInst 1                                  'Add FOR
              ForDefflg = True                            'mark definition
            End If
          End If
        Else
          Errstr = "Expected '('"
        End If
        
' Do ---------------------------------------------------------------------------
      Case 202
        If Instructions(InstrPtr + 1) = iLCbrace Then     'expected '{'?
          ApndInst 1                                      'yes
          BmpInd = True                                   'we will bump indent starting on next line
          DoIdx = DoIdx + 1                               'bump Do depth flag
          i = FindEbrace()                                'now find a matching end brace
          If i < 0 Then
            Errstr = "Cannot find a matching end brace '}'"
          Else
            Instructions(i) = iDCBrace                    'special Do block terminator
            Select Case Instructions(i + 1)
              Case iWhile                                 'do..while?
                Instructions(i) = iDWBrace                'set mark for easy run processing
              Case iUntil                                 'do...Until?
                Instructions(i) = iDUBrace                'set mark for easy run processing
            End Select
          End If
        Else
          Errstr = "Opening brace '{' expected"
        End If
      
' While ------------------------------------------------------------------------
      Case 203
        If PrvCode = iDWBrace Then                        'if prev instr. was a DO-WHILE end brace...
          DoWhiDefFlg = True                              'turn on Do...While definition flag
        End If
        
        If Instructions(InstrPtr + 1) = iLparen Then      'expected '('?
          ApndInst 1                                      'yes, so add it
          i = FindEPar(0)                                 'find matching paren
          If i < 0 Then
            Errstr = "Cannot find a matching end paren ')'"
          Else
            Instructions(i - 1) = iWparen                 'mark end of While expression
            WhiDefFlg = True                              'while definition active
          End If
        Else
          Errstr = "Expected '('"
        End If
' Pmt --------------------------------------------------------------------------
      Case 204
        If CheckForText(InstrPtr, DisplayWidth) Then
          ApndQTx                                         'append user response
        Else
          Errstr = "Parameter error"
        End If
'-------------------------------------------------------------------------------
'      Case 205  ' Rad  'no need to check

' Ukey -------------------------------------------------------------------------
      Case 206
        If SbrDefFlg Then                             'subroutine active...
          Errstr = "Subroutine definition already active"
          Exit Do
        End If
        i = InstrPtr                                  'save copy of InstrPtr
        If Not CheckForLabel(InstrPtr, 1) Then        'verify label found
          Errstr = "Invaid Named definition"
          Exit Do
        End If
        
        KeyName = UCase$(TxtData)                     'save a copy of name
        ApndQTx                                       'so add to formatted text
        KeyLbl = vbNullString                         'init key label to nothing
        KeyTip = vbNullString                         'init tip to nothing
        If Instructions(InstrPtr + 1) = iLbl Then     'found Lbl token?
          ApndInst 1                                  'yes, add it
          If Not CheckForText(InstrPtr, LabelWidth) Then 'grab label for keys
            Errstr = "Invaid key label"
            Exit Do
          End If
          KeyLbl = TxtData                            'save label definition
          ApndQTx                                     'append key label
          If Instructions(InstrPtr + 1) = iComma Then 'found comma, leading comment?
            ApndInst 1                                'yes, add comma to list
            If Not CheckForText(InstrPtr, DisplayWidth) Then  'grab comment
              Errstr = "Invaid comment"
              Exit Do
            End If
            KeyTip = TxtData                          'save comment definition
            ApndQTx                                   'add add to buffer
          End If
        End If
        
        If Len(KeyLbl) = 0 Then KeyLbl = KeyName      'use key letter as label, if not defined
        
        j = Asc(KeyName) - 64                         'get index (1-26) for key in Lbls() array
        If Instructions(InstrPtr + 1) = iLCbrace Then 'found '{'?
          ApndInst 1                                  'then add brace to instruction list
          Call DefineLables(TypKey, i, j)             'yes, so first reset existing label (J)
          With Lbls(j)                                'point to user-key definition
            .lblName = KeyLbl                         'stuff label and tip
            .lblCmt = KeyTip
            .lblUdef = True                           'user defined the key
            i = .LblEnd                               'get matching end brace address
          End With
          If i < 0 Then
            Errstr = "Cannot find a matching ending brace '}'"
          Else
            Instructions(i) = iBCBrace                'force returning end brace
            BmpInd = True
            SbrDefFlg = True                          'tag subroutine active
            frmVisualCalc.mnuWinUkey.Enabled = True
          End If
        Else
          Errstr = "Opening brace '{' expected"
        End If

' Plot -------------------------------------------------------------------------
      Case 211
        ' Check for: Plot CLR
        If Instructions(InstrPtr + 1) = iCLR Then     'check for Plot Clear
          ApndInst 1                                  'add CLR
        ' Check for: Plot Sbr Reset or Plot Sbr LABEL
        ElseIf Instructions(InstrPtr + 1) = iSbr Then 'check for Plot Sbr
          ApndInst 1
          If Instructions(InstrPtr + 1) = iReset Then 'if Plot Sbr Reset, then reset Sbr ref
            ApndInst 1
          ElseIf CheckForLabel(InstrPtr, LabelWidth) Then 'else check for Plot Sbr LABEL
            ApndQTx
          Else
            Errstr = "Invalid parameter"
          End If
        ' Check for: Plot Open or Plot Close
        ElseIf Instructions(InstrPtr + 1) = iOpen Then
          ApndInst 1
        ElseIf Instructions(InstrPtr + 1) = iClose Then
          ApndInst 1
        Else
        ' Check for: [-](x,y)[,z]
          If Instructions(InstrPtr + 1) = iMinus Then       'else check for -(x,y)  (Draw to)
            ApndInst 1                                      'add [-]
          End If
          If Instructions(InstrPtr + 1) = iLparen Then      'check for point (x,y)  (dot at)
            If Not CheckForXY(Errstr) Then Exit Do
            If Instructions(InstrPtr + 1) = iComma Then     'check for ,1 (fill)
              ApndInst 1
              If CheckForNumber(InstrPtr, 1, 1) Then        'if 0, no fill
                ApndTxt
                If Instructions(InstrPtr + 1) = iComma Then 'check for border color
                  ApndInst 1
                  If Instructions(InstrPtr + 1) = iIND Then 'check for IND var
                    If CheckForVar(Errstr) Then
                      ApndTxt                               'accept variable
                    End If
                  ElseIf CheckForValue(InstrPtr) Then       'try value
                    ApndTxt                                 'accept value
                  End If
                End If
              Else
                Errstr = "Invalid parameter"
              End If
            End If
          Else
            Errstr = "Invalid parameter"
          End If
        End If
' Nvar -------------------------------------------------------------------------
      Case 212
        i = InstrPtr                                    'save definition location
        Nm = vbNullString
        If Not CheckForNumber(InstrPtr, 2, 99) Then     'check for variable number
          Errstr = "Invalid variable definition"
        Else
          Vn = CLng(TxtData)                            'grab variable number
          ApndTxt                                       'did, so append text
          If Instructions(InstrPtr + 1) = iLbl Then     'found Lbl?
            ApndInst 1                                  'yes, so add it
            If CheckForLabel(InstrPtr, LabelWidth) Then 'found the label
              If Len(TxtData) = 1 Then
                Errstr = "Cannot define variables names of 1 character"
                Exit Do
              End If
              If FindVblMatch(TxtData) <> -1 Then       'search for matching name
                ForcError "Variable name '" & TxtData & "' has already been defined"
                Exit Do
              End If
              Nm = TxtData                              'OK, so grab name
              ApndQTx                                   'add name text
            Else
              Errstr = "Invaid Named definition"
              Exit Do
            End If
          End If
          
          With Variables(Vn)                            'now apply changes to variable
            .VarType = vNumber                          'double precision
            .VName = Nm                                 'apply name if defined
            Set .Vdata = Nothing                        'clear any defined classes (and children)
            Set .Vdata = New clsVarSto                  'init brand new variable storage
            .Vdata.VarRoot = Vn                         'set root variable it is associated with
            .VuDef = True
            .Vaddr = i
          End With
          
          If CheckDims(Vn, X, Y, Errstr) Then           'if dimension dims found...
            Call BuildMDAry(Vn, X, Y, False)            'process them
          End If
        End If
        frmVisualCalc.mnuWinVar.Enabled = True

' If ---------------------------------------------------------------------------
      Case 214
        If Instructions(InstrPtr + 1) <> iLparen Then
          Errstr = "Expected '('"
        Else
          ApndInst 1                      'add opening paren (
          i = FindEPar(0)
          If i < 0 Then
            Errstr = "Cannot find a matching end paren ')'"
          Else
            Instructions(i - 1) = iIparen 'mark end of expression check
            IfDefFlg = True               'indicate we are defining it
          End If
        End If
        
' Else -------------------------------------------------------------------------
      Case 215
        If Instructions(InstrPtr - 1) = iICBrace And Instructions(InstrPtr + 1) = iLCbrace Then
          ApndInst 1        'add opening brace {
          BmpInd = True     'we should bump indent on next line
          i = FindEbrace()  'find matching end brace
          If i < 0 Then
            Errstr = "Cannot find a matching ending brace '}'"
          Else
            Instructions(i) = iICBrace
            IfIdx = IfIdx + 1 'bump depth of embedded Ifs
          End If
        Else
          Errstr = "Else was expected to be formatted '} Else {'"
        End If

' Cont -------------------------------------------------------------------------
      Case 216
        If ForIdx + WhiIdx + DoIdx + CaseIdx = 0 Then 'if not in any of these
          Errstr = "Instruction must be defined within a Loop or Case block"
        End If

' Break ------------------------------------------------------------------------
      Case 217
        If ForIdx + WhiIdx + DoIdx = 0 Then 'if not in any of these
          Errstr = "Instruction must be defined within a Loop block"
        End If
        
'-------------------------------------------------------------------------------
'      Case 218  ' Grad 'no need to check
'      Case 219  ' R/S  'no need to check
'      Case 222  ' +/-  'no need to check
'      Case 223  ' =    'no need to check

' Print, Print; ------------------------------------------------------------------------
      Case 224, 352
        Do  'use this DO loop so we can break out of checks (it NEVER actually loops)
          ' check for: Print ADV, Print Reset
          Select Case Instructions(InstrPtr + 1)
            Case iAdv, iReset
              ApndInst 1
              Exit Do
          End Select
        '
        ' first check for (x,y[,dir]) 'assumes dir=0
        '
          If Instructions(InstrPtr + 1) = iLparen Then
            ApndInst 1                                    'add (
            Errstr = "Invalid parameter"                  'init error trap
            If Instructions(InstrPtr + 1) = iIND Then
              ApndInst 1
              If Not CheckForAlNum(InstrPtr, LabelWidth) Then Exit Do
            Else
              If Not CheckForNumber(InstrPtr, 3, PlotWidth) Then Exit Do
            End If
            ApndTxt                                       'add X value
            If Instructions(InstrPtr + 1) <> iComma Then Exit Do
            ApndInst 1                                    'add [,]
            If Instructions(InstrPtr + 1) = iIND Then
              ApndInst 1
              If Not CheckForAlNum(InstrPtr, LabelWidth) Then Exit Do
            Else
              If Not CheckForNumber(InstrPtr, 3, PlotHeight) Then Exit Do
            End If
            ApndTxt                                       'add Y value
            ' check for optional DIR parameter
            If Instructions(InstrPtr + 1) = iComma Then
              ApndInst 1                                  'add [,]
              If Instructions(InstrPtr + 1) = iIND Then
                ApndInst 1
                If Not CheckForAlNum(InstrPtr, LabelWidth) Then Exit Do
              Else
                If Not CheckForNumber(InstrPtr, 1, 7) Then Exit Do
              End If
              ApndTxt                                     'add Direction
            End If
            If Instructions(InstrPtr + 1) <> iRparen Then Exit Do
            ApndInst 1                                    'add ')'
            Errstr = vbNullString
          End If
          '
          ' now check for optional text to print to plot
          '
          If Instructions(InstrPtr + 1) = iIND Then 'check for indirect variable def of text
            Call CheckForVar(Errstr)
          Else
            If Instructions(InstrPtr + 1) < 128 Then        'text data follows?
              If CheckForText(InstrPtr, DisplayWidth) Then  'check for text if so
                ApndQTx                                     'add it if found
              Else
                Errstr = "Invalid parameter"
              End If
            End If
          End If
          Exit Do
        Loop
' Tvar -------------------------------------------------------------------------
      Case 225
        i = InstrPtr                                    'save definition location
        Nm = vbNullString                               'init no variable name
        Ln = 0                                          'init no data length

        If Not CheckForNumber(InstrPtr, 2, 99) Then     'check for variable number
          Errstr = "Invalid variable definition"
        Else
          Vn = CInt(TstData)                            'grab variable number
          ApndTxt                                       'did, so append text
          If Instructions(InstrPtr + 1) = iLbl Then     'found Lbl?
            ApndInst 1                                  'yes, so add it
            If CheckForLabel(InstrPtr, LabelWidth) Then 'found the label
              If Len(TxtData) = 1 Then
                Errstr = "Cannot define variables names of 1 character"
                Exit Do
              End If
              If FindVblMatch(TxtData) <> -1 Then       'search for matching name
                ForcError "Variable name '" & TxtData & "' has already been defined"
                Exit Do
              End If
              Nm = TxtData                              'OK, so grab name
              ApndQTx                                   'add name text
            Else
              Errstr = "Invaid Named definition"
              Exit Do
            End If
          End If
          If Instructions(InstrPtr + 1) = iLen Then     'found Len?
            ApndInst 1                                  'yes, so add it
            If CheckForNumber(InstrPtr, 2, DisplayWidth) Then 'found the length
              Ln = CInt(TxtData)                        'grab it
              ApndTxt                                   'add name text
            Else
              Errstr = "Invaid Len definition"
              Exit Do
            End If
          End If
          
          With Variables(Vn)                            'now apply changes to variable
            .VarType = vString                          'text string
            .VName = Nm                                 'apply name if defined
            .VdataLen = Ln                              'set len, if non-zero
            Set .Vdata = Nothing                        'clear any defined classes (and children)
            Set .Vdata = New clsVarSto                  'init brand new variable storage
            .Vdata.VarRoot = Vn                         'set root variable it is associated with
            .VuDef = True
            .Vaddr = i
          End With
          
          If CheckDims(Vn, X, Y, Errstr) Then           'if dimension dims found...
            Call BuildMDAry(Vn, X, Y, False)            'process them
          End If
        End If
        frmVisualCalc.mnuWinVar.Enabled = True

' >> ---------------------------------------------------------------------------
    Case 226
        Errstr = "Invalid Shift value (1-9)"            'init with error
        If CheckForNumber(InstrPtr, 1, 9) Then          'check value
          If TxtData <> "0" Then                        'allow only 1-9, not 0
            ApndTxt                                     'append value if 1-9
            Errstr = vbNullString                       'remove error
          End If
        End If

'-------------------------------------------------------------------------------
'      Case 228  ' X²   'no need to check
'      Case 229  ' Pi   'no need to check
'      Case 230  ' Rnd  'no need to check
'      Case 231  ' Mil  'no need to check

' Pvt, Pub ---------------------------------------------------------------------
      Case 232, 360 '
        Select Case Instructions(InstrPtr + 1)
          Case 180, 193, 206 'Sbr, Lbl, Ukey
          Case Else
            Errstr = "Invalid parameter"
        End Select

' Const ------------------------------------------------------------------------
      Case 233
        j = InstrPtr                          'save start of definition
        If Not CheckForLabel(InstrPtr, LabelWidth) Then
          Errstr = "Invaid Named definition"
        Else
          If CBool(FindLblMatch(TxtData)) Then 'search for matching label
            ForcError "Label '" & TxtData & "' has already been defined"
          Else
            Nm = TxtData                      'save name
            ApndQTx                           'append name
            If Instructions(InstrPtr + 1) <> iLCbrace Then
              Errstr = "Invalid Named Definition"
            Else
              ApndInst 1                      'add {
              i = InstrPtr                    'save pointer to '{'
              K = FindEbrace()                'find }
              If K < 0 Then
                Errstr = "Cannot find a matching ending brace '}'"
              Else
                Instructions(K) = iCNBrace    'mark end
                If CheckForText(InstrPtr, DisplayWidth) Then
                  ApndQTx                     'grab data
                  If Instructions(InstrPtr + 1) <> iCNBrace Then
                    Errstr = "Closing brace '}' expected"
                  Else
                    ApndInst 1                'add }
                    '
                    ' now add constant to label pool
                    '
                    With Lbls(LblCnt)
                      .lblName = UCase$(Nm)   'apply name for Const item
                      .LblTyp = TypConst      'type as constant
                      .LblScope = Pvt         'always private
                      .lblAddr = j
                      .LblDat = i + 1
                      .LblEnd = K + 1
                      .lblCmt = TxtData       'store data in un-used comment area
                      .lblUdef = True         'mark as user-defined
                    End With
                    Call BumpLblCnt           'bump size of label pool
                    frmVisualCalc.mnuWinConst.Enabled = True
                  End If
                End If
              End If
            End If
          End If
        End If

' Struct -----------------------------------------------------------------------
      Case 234
        j = InstrPtr                                    'save address for definition
        If Not CheckForLabel(InstrPtr, LabelWidth) Then 'check for structure name
          Errstr = "Invaid Named definition"
          Exit Do
        End If
        Nm = TxtData                                    'save name
        ApndQTx                                         'append name
        
        If Instructions(InstrPtr + 1) <> iLCbrace Then  'check for '{'
          Errstr = "Opening brace '{' expected"
          Exit Do
        End If
        ApndInst 1                                      'add {
        K = InstrPtr + 1                                'save address to '{'+1
        
        i = FindEbrace()                                'find matching end brace
        If i < 0 Then
          Errstr = "Cannot find a matching ending brace '}'"
          Exit Do
        End If
'
' before we define our structure, be sure that the next instruction is another open brace
'
        If Instructions(InstrPtr + 1) <> iLCbrace Then
          Errstr = "Structures can only contain Structure Item definitions"
          Exit Do
        End If
'
' we now have all information available for starting a structure definition
'
        Instructions(i) = iSTBrace                      'mark end of main body
        StructCnt = StructCnt + 1                       'bump structure counter
        With Lbls(LblCnt)
          .lblName = UCase$(Nm)                         'set name
          .LblTyp = TypStruct                           'set type
          .LblScope = Pvt                               'always private
          .LblValue = StructCnt                         'index into structure array
          .lblAddr = j                                  'set definition address
          .LblDat = K                                   'to data address
          .lblUdef = True                               'mark as user-defined
        End With
        
        ReDim Preserve StructPl(StructCnt)              'set aside space for new struct
        With StructPl(StructCnt)                        'point to latest structure
          .StSiz = 0                                    'init size to zero
          .StItmCnt = 0                                 'no items
        End With
        
        Call BumpLblCnt                                 'bump size of label pool
        StDefFlg = True                                 'we are now defining a structure
        BmpInd = True                                   'bump indent
        frmVisualCalc.mnuWinStruct.Enabled = True
        
'-------------------------------------------------------------------------------
'      Case 235  ' NxLbl  'ignored
'      Case 236  ' PvLnl  'ignored

' Line -------------------------------------------------------------------------
      Case 237  ' Line
        Do  'use this loop for easy early exits
          If Instructions(InstrPtr + 1) <> iLparen Then 'check for start (x,y)
            Errstr = "Invalid parameter"
            Exit Do
          End If
          If Not CheckForXY(Errstr) Then Exit Do
          If Instructions(InstrPtr + 1) <> iMinus Then  'ensure -(x,y) follows
            Errstr = "Invalid parameter"
            Exit Do
          End If
          ApndInst 1                                    'append '-'
          If Not CheckForXY(Errstr) Then Exit Do
          
          If Instructions(InstrPtr + 1) = iComma Then   'check for optional ',1' or ',2' (,0 is default)
            ApndInst 1                                  'append ','
            If CheckForNumber(InstrPtr, 1, 2) Then
              ApndTxt                                   'append value
            Else
              Errstr = "Invalid parameter"
            End If
          End If
          Exit Do                                       'always exit the loop
        Loop
        
' [, ]  ------------------------------------------------------------------------
      Case iLbrkt, iRbrkt
        Errstr = "Unexpected instruction"

' ClrVar -----------------------------------------------------------------------
      Case 240
        If Instructions(InstrPtr + 1) = iAll Then       'allow ClrVar All
          ApndInst 1
        Else
          Call CheckForVar(Errstr)                      'check for var, Indirection, and 1/2D arrays
        End If

' SzOf  ------------------------------------------------------------------------
      Case 241
        If Instructions(InstrPtr + 1) = i1X Then ApndInst 1 'check for optional 1/X parameter
        Call CheckForVar(Errstr)                        'check for var, Indirection, and 1/2D arrays
      
' Def, IfDef, Udef, !Def -------------------------------------------------------
      Case 242, 243, 370, 371
        If CheckForLabel(InstrPtr, LabelWidth) Then
          ApndQTx
        Else
          Errstr = "Unexpencted instruction"
        End If

'-------------------------------------------------------------------------------
'      Case 244  'Edef 'ignored

'-------------------------------------------------------------------------------
                '--2nd keys---------------------------------------------
'-------------------------------------------------------------------------------

' MDL --------------------------------------------------------------------------
      Case iMDL
        Select Case Instructions(InstrPtr + 1)  'check next
          Case iLoad                            'LOAD, for MDL Load
            ApndInst 1                          'append valid instruction to InstTxt
                                                'and bump InstrPtr by the supplied count
            If CheckForNumber(InstrPtr, 4, 9999) Then 'check data
               ApndTxt                          'append TxtData to format string
            Else
              Errstr = "Invalid Module number"
            End If
          Case iRCL, iLbl                       'Rcl, Lbl
            ApndInst 1                          'append valid instruction to InstTxt
          Case Else
            Errstr = "Invalid Run-Time MDL command"
        End Select

'      Case 258  ' CMM  'no need to check
'      Case 261  ' CMs  'no need to check
'      Case 262  ' CP   'no need to check
'      Case iList' List 'Not active
'      Case 264  ' BST    'ignored
'      Case 265  ' DEL    'ignored
'      Case 266  ' Paste  'ignored

' USR  -------------------------------------------------------------------------
      Case iUSR
        If Instructions(InstrPtr + 1) = iIND Then
          Call CheckForVar(Errstr)                      'check for variable, IND, and arrays
        Else
          If CheckForNumber(InstrPtr, 2, MaxUSR) Then
            ApndTxt
          Else
            Errstr = "Invalid USR Operation number"
          End If
        End If

'-------------------------------------------------------------------------------
'      Case 268  ' RtoP   'no need to check
'      Case 269  ' Push   'no need to check
'      Case 270  ' Pop    'no need to check
'      Case 271  ' StkEx  'no need to check
'      Case 277  ' eX     'no need to check
'      Case 278  ' E-     'no need to check
'      Case 279  ' StDev  'no need to check
'      Case 280  ' Varnc  'no need to check
'      Case 281  ' Yint   'no need to check
'      Case 295  ' NOP  'ignored
      
' [;] --------------------------------------------------------------------------
'      Case iSemic     'ignore (primary statement separator)
'      Case iSemiColon 'for For Statements
'      Case 297  ' Log  'no need to check
'      Case 298  ' 10^  'no need to check

' Fmt --------------------------------------------------------------------------
      Case 300
        If Instructions(InstrPtr + 1) = iIND Then       'if formatting is via indirect variable
          CheckForVar (Errstr)                          'grab indirect variable
        Else
          If CheckForText(InstrPtr, DisplayWidth) Then  'otherwise scan for text
            ApndQTx                                     'append it as needed
          Else
            Errstr = "Parameter Error"                  'oops
          End If
        End If
'-------------------------------------------------------------------------------
'      Case 303  ' Frac   'no need to check
'      Case 304  ' Sgn    'no need to check
'      Case 305  ' !Fix   'no need to check
'      Case 306  ' D.ddd  'no need to check
'      Case 307  ' !EE    'no need to check

' Call -------------------------------------------------------------------------
      Case 308
        If Instructions(InstrPtr + 1) = iIND Then       'allow Call IND var
          Call CheckForVar(Errstr)                      'check for variable, IND, and arrays
        ElseIf CheckForLabel(InstrPtr, LabelWidth) Then 'else allow Call Label
          ApndQTx
        Else
          Errstr = "Unexpencted instruction"
        End If

' Open -------------------------------------------------------------------------
      Case 316
        If Instructions(InstrPtr + 1) = iIND Then       'first gather filename...
          Call CheckForVar(Errstr)                      'check for variable, IND, and arrays
        Else
          If CheckForLabel(InstrPtr, LabelWidth) Then   'get filename
            ApndQTx
          Else
            Errstr = "Expected filename"
          End If
        End If
        If CBool(Len(Errstr)) Then Exit Do
        '
        ' now check for expected For Command (R|W|A|B)
        '
        If Instructions(InstrPtr + 1) = iFor Then       'next instruction For?
          ApndInst 1                                    'yes, accept it
          If Instructions(InstrPtr + 1) = iIND Then
            Call CheckForVar(Errstr)                    'check for variable, IND, and arrays
          ElseIf CheckForLabel(InstrPtr, 1) Then        'next is R|W|A|B?
            If CBool(InStr(1, "RWAB", TxtData)) Then
              ApndQTx                                   'accept it
            Else
              Errstr = "Invalid parameter. Expected R|W|A|B"
            End If
          Else
            Errstr = "Unexpencted instruction"
          End If
        Else
          Errstr = "Expected FOR parameter"
        End If
        If CBool(Len(Errstr)) Then Exit Do
        '
        ' Now check for expected AS parameter
        '
        If Instructions(InstrPtr + 1) = iAs Then        'next instruction AS?
          ApndInst 1                                    'yes, accept it
          If Instructions(InstrPtr + 1) = iIND Then
            Call CheckForVar(Errstr)                    'check for variable, IND, and arrays
          ElseIf CheckForNumber(InstrPtr, 1, 9) Then    'allow 0-9
            If TxtData <> "0" Then                        'allow only 1-9, not 0
              ApndTxt                                     'append value if 1-9
            Else
              Errstr = "Expected value 1-9"
            End If
          Else
            Errstr = "Expected value 1-9"
          End If
        Else
          Errstr = "Expencted AS parameter"
        End If
        If CBool(Len(Errstr)) Then Exit Do
        '
        ' now check for optional LEN parameter
        '
        If Instructions(InstrPtr + 1) = iLen Then       'next instruction LEN?
          ApndInst 1                                    'yes, accept it
          If Instructions(InstrPtr + 1) = iIND Then     'indirection?
            Call CheckForVar(Errstr)                    'check for variable, IND, and arrays
          ElseIf CheckForNumber(InstrPtr, 5, 32767) Then 'allow large data
            ApndTxt
          Else
            Errstr = "LEN parameter value is invalid"
          End If
        End If

' Close ------------------------------------------------------------------------
      Case 317
        Select Case Instructions(InstrPtr + 1)
          Case iAll, 0 To 9
            ApndInst 1
          Case iIND
            Call CheckForVar(Errstr)                    'check for variable, IND, and arrays
        End Select

' Read, Write, Get, Put --------------------------------------------------------
      Case 318, 319, 323, 324
        If Instructions(InstrPtr + 1) = iIND Then       'indirection?
          Call CheckForVar(Errstr)                      'check for variable, IND, and arrays
        ElseIf CheckForNumber(InstrPtr, 1, 9) Then      'allow 0-9
          If TxtData <> "0" Then                        'allow only 1-9, not 0
            ApndTxt                                     'append value if 1-9
          Else
            Errstr = "Expected value 1-9"
          End If
        Else
          Errstr = "Expected value 1-9"
        End If
        If CBool(Len(Errstr)) Then Exit Do
        If Code = iGet Or Code = iPut Then
          If Instructions(InstrPtr + 1) = iComma Then
            ApndInst 1                                  'append ','
            If Instructions(InstrPtr + 1) = iIND Then   'indirection?
              Call CheckForVar(Errstr)                  'check for variable, IND, and arrays
            ElseIf CheckForNumber(InstrPtr, 5, 32767) Then  'allow 1-32767
              If TstData = "0" Then
                Errstr = "Parameter is out of range (1-32767)"
              Else
                ApndTxt                                 'append if ok
              End If
            Else
              Errstr = "Parameter is out of range (1-32767)"
            End If
          End If
        End If
        If Instructions(InstrPtr + 1) = iAll And _
                   (Code = iWrite Or Code = iRead) Then 'Read/Rrite x ALL
          ApndInst 1                                    'add ALL
        ElseIf Instructions(InstrPtr + 1) = iWith Then  'Read/Write/Get/Put x With?
          ApndInst 1                                    'add WITH
          Call CheckForVar(Errstr)                      'check for a parameter
          ApndTxt
        Else
          Errstr = "Expected WITH or ALL parameter"
        End If

' Swap -------------------------------------------------------------------------
      Case 320
        Call CheckForVar(Errstr)                                'check for variable, IND, and arrays
        If Len(Errstr) = 0 Then
          If Instructions(InstrPtr + 1) <> iComma Then
            Errstr = "Comma [,] expected"
          Else
            ApndInst 1                                            'add comma
            Call CheckForVar(Errstr)                              'check for variable, IND, and arrays
          End If
        End If

' Gto --------------------------------------------------------------------------
      Case 321
        If Instructions(InstrPtr + 1) = iIND Then     'allow indirection
          Call CheckForVar(Errstr)                    'check for variable, IND, and arrays
        Else
          If CheckForLabel(InstrPtr, LabelWidth) Then 'else we are expecting a label
            ApndQTx                                   'and we got it
          Else
            Errstr = "Unexpencted instruction"
          End If
        End If

' Lof --------------------------------------------------------------------------
      Case 322
        If Instructions(InstrPtr + 1) = iIND Then       'indirection?
          Call CheckForVar(Errstr)                      'check for variable, IND, and arrays
        ElseIf CheckForNumber(InstrPtr, 1, 9) Then      'allow 0-9
          If TxtData <> "0" Then                        'allow only 1-9, not 0
            ApndTxt                                     'append value if 1-9
          Else
            Errstr = "Expected value 1-9"
          End If
        Else
          Errstr = "Expected value 1-9"
        End If

' SysBP ------------------------------------------------------------------------
      Case 326
        If CheckForNumber(InstrPtr, 1, 4) Then
          ApndInst 1
        Else
          Errstr = "Invaid parameter"
        End If

' Dsz, Dsnz --------------------------------------------------------------------
      Case 331, 332
        If CheckForVar(Errstr) Then                               'check for variable, IND, and arrays
          If Instructions(InstrPtr + 1) = iLCbrace Then
            ApndInst 1                                            'block for TRUE
            BmpInd = True                                         'we should bump indent on next line
            i = FindEbrace()                                      'find ending brace
            If i < 0 Then
              Errstr = "Cannot find a matching ending brace '}'"
            Else
              Instructions(i) = iICBrace                          'mark end of If block
              IfIdx = IfIdx + 1 'bump depth of embedded Ifs
            End If
          Else
            Errstr = "Opening brace '{' expected"
          End If
        Else
          Errstr = "Invalid parameter"
        End If

' All --------------------------------------------------------------------------
      Case 333
        Errstr = "Unexpected Instruction"

' Rtn --------------------------------------------------------------------------
'      Case 334  'ignore

' LSet, RSet -------------------------------------------------------------------
      Case 335, 335
        Call CheckForVar(Errstr)  'check for variable, IND, and arrays
        
' Printf -----------------------------------------------------------------------
      Case 337
        Call CheckForVar(Errstr)  'check for variable, IND, and arrays

' RGB ------------------------------------------------------------------------
      Case 339
        If Instructions(InstrPtr + 1) <> iLparen Then       'Check for (
          Errstr = "Expected '('"
          Exit Do
        End If
        ApndInst 1                                          'add (
        Errstr = "Invalid parameter"                        'init error flag
        If Instructions(InstrPtr + 1) = iIND Then           'check for Red
          If Not CheckForVar(Errstr) Then Exit Do           'check variable
        ElseIf Not CheckForNumber(InstrPtr, 3, 255) Then    'else check for number
          Exit Do
        End If
        ApndTxt                                             'add Red value
        If Instructions(InstrPtr + 1) <> iComma Then Exit Do 'check for comma
        ApndInst 1                                          'add comma
        If Instructions(InstrPtr + 1) = iIND Then           'check for Green
          If Not CheckForVar(Errstr) Then Exit Do           'check variable
        ElseIf Not CheckForNumber(InstrPtr, 3, 255) Then    'else check for number
          Exit Do
        End If
        ApndTxt                                             'add Green value
        If Instructions(InstrPtr + 1) <> iComma Then Exit Do 'check for comma
        ApndInst 1                                          'add comma
        If Instructions(InstrPtr + 1) = iIND Then           'check for Blue
          If Not CheckForVar(Errstr) Then Exit Do           'check variable
        ElseIf Not CheckForNumber(InstrPtr, 3, 255) Then    'else check for number
          Exit Do
        End If
        ApndTxt                                             'add Blue value
        If Instructions(InstrPtr + 1) = iRparen Then        'Check for )
          ApndInst 1
          Errstr = vbNullString                             'remove error string (we are OK)
        Else
          Errstr = "Cannot find a matching end paren ')'"
        End If

' Ivar -------------------------------------------------------------------------
      Case 340
        i = InstrPtr                                    'save definition location
        Nm = vbNullString
        If Not CheckForNumber(InstrPtr, 2, 99) Then     'check for variable number
          Errstr = "Invalid variable definition"
        Else
          Vn = CInt(TxtData)                            'grab variable number
          ApndTxt                                       'did, so append text
          If Instructions(InstrPtr + 1) = iLbl Then     'found Lbl?
            ApndInst 1                                  'yes, so add it
            If CheckForLabel(InstrPtr, LabelWidth) Then 'found the label
              If Len(TxtData) = 1 Then
                Errstr = "Cannot define variables names of 1 character"
                Exit Do
              End If
              If FindVblMatch(TxtData) <> -1 Then       'search for matching name
                ForcError "Variable name '" & TxtData & "' has already been defined"
                Exit Do
              End If
              Nm = TxtData                              'OK, so grab name
              ApndQTx                                   'add name text
            Else
              Errstr = "Invaid Named definition"
              Exit Do
            End If
          End If
          
          With Variables(Vn)                            'now apply changes to variable
            .VarType = vInteger                         'long integer
            .VName = Nm                                 'apply name if defined
            Set .Vdata = Nothing                        'clear any defined classes (and children)
            Set .Vdata = New clsVarSto                  'init brand new variable storage
            .Vdata.VarRoot = Vn                         'set root variable it is associated with
            .VuDef = True
            .Vaddr = i
          End With
          
          If CheckDims(Vn, X, Y, Errstr) Then           'if dimension dims found...
            Call BuildMDAry(Vn, X, Y, False)            'process them
          End If
        End If
        frmVisualCalc.mnuWinVar.Enabled = True

' As ---------------------------------------------------------------------------
      Case 342
        Errstr = "Unexpected Instruction"
      
' ElseIf -----------------------------------------------------------------------
      Case 343
        If Instructions(InstrPtr - 1) = iICBrace And Instructions(InstrPtr + 1) = iLparen Then
          ApndInst 1                      'add opening paren (
          i = FindEPar(0)                 'find end of expression
          If i < 0 Then
            Errstr = "Cannot find a matching end paren ')'"
          Else
            Instructions(i - 1) = iIparen 'mark end of IF expression
            IfDefFlg = True               'indicate we are defining the expression
          End If
        Else
          Errstr = "ElseIf was expected to be formatted '} ElseIf ('"
        End If

'-------------------------------------------------------------------------------
      Case 344  ' DBG
      Select Case Instructions(InstrPtr + 1)
        Case iOpen, iClose
          ApndInst 1
      End Select


'-------------------------------------------------------------------------------
'      Case 345  ' Gfree  'No Need to check

' Len --------------------------------------------------------------------------
      Case 346
        Errstr = "Unexpected instruction"

'-------------------------------------------------------------------------------
'      Case 347  ' Stop 'no need to check

' With ------------------------------------------------------------------------
      Case 348
        Errstr = "Unexpected instruction"

' ',' --------------------------------------------------------------------------
'      Case iComma  'no need to check (alternate statement separator)

'-------------------------------------------------------------------------------
'      Case 350  ' Val  'No need to check
'      Case 351  ' Adv  'no need to check

' Cvar -------------------------------------------------------------------------
      Case 353
        i = InstrPtr                                    'save definition location
        Nm = vbNullString
        If Not CheckForNumber(InstrPtr, 2, 99) Then     'check for variable number
          Errstr = "Invalid variable definition"
        Else
          Vn = CInt(TxtData)                            'grab variable number
          ApndTxt                                       'did, so append text
          If Instructions(InstrPtr + 1) = iLbl Then     'found Lbl?
            ApndInst 1                                  'yes, so add it
            If CheckForLabel(InstrPtr, LabelWidth) Then 'found the label
              If Len(TxtData) = 1 Then
                Errstr = "Cannot define variables names of 1 character"
                Exit Do
              End If
              If FindVblMatch(TxtData) <> -1 Then       'search for matching name
                ForcError "Variable name '" & TxtData & "' has already been defined"
                Exit Do
              End If
              Nm = TxtData                              'OK, so grab name
              ApndQTx                                   'add name text
            Else
              Errstr = "Invaid Named definition"
              Exit Do
            End If
          End If
          
          With Variables(Vn)                            'now apply changes to variable
            .VarType = vChar                            'character
            .VName = Nm                                 'apply name if defined
            Set .Vdata = Nothing                        'clear any defined classes (and children)
            Set .Vdata = New clsVarSto                  'init brand new variable storage
            .Vdata.VarRoot = Vn                         'set root variable it is associated with
            .VuDef = True
            .Vaddr = i
          End With
          
          If CheckDims(Vn, X, Y, Errstr) Then           'if dimension dims found...
            Call BuildMDAry(Vn, X, Y, False)            'process them
          End If
        End If
        frmVisualCalc.mnuWinVar.Enabled = True

' << ---------------------------------------------------------------------------
    Case 354
        If CheckForNumber(InstrPtr, 1, 9) Then
          If TxtData = "0" Then
            Errstr = "Invalid Shift value (1-9)"
          Else
            ApndTxt                                     'append value
          End If
        Else
          Errstr = "Invalid Shift value (1-9)"
        End If

'-------------------------------------------------------------------------------
'      Case 356  ' Sqrt 'no need to check
'      Case 357  ' e    'no need to check
'      Case 358  'Rnd#  'no need to check

' Until ------------------------------------------------------------------------
      Case 359
        If PrvCode = iDUBrace Then                          'if prev instr. was a DO-UNTIL end brace...
          If Instructions(InstrPtr + 1) = iLparen Then      'expected '('?
            ApndInst 1                                      'yes, so add it
            i = FindEPar(0)                                 'find matching paren
            If i < 0 Then
              Errstr = "Cannot find a matching end paren ')'"
            Else
              Instructions(i - 1) = iUparen                 'mark end of Until expression
              UtlDefFlg = True                              'Until definition active
            End If
          Else
            Errstr = "Expected '('"
          End If
        Else
          Errstr = "Invalid use of Until"
        End If

' Enum -------------------------------------------------------------------------
      Case 361
        EnumIdx = 0                                       'init enumeration index @ zero
        If Instructions(InstrPtr + 1) < 10 Then           'initial value for enumerations?
          If CheckForNumber(InstrPtr, LabelWidth, 0) Then
            ApndTxt                                       'append data
            EnumIdx = CLng(TstData)                       'set start of enumeration values
          Else
            Errstr = "Enum initiator must be an integer value"
            Exit Do
          End If
        End If
        If Instructions(InstrPtr + 1) = iLCbrace Then     'expected '{'?
          ApndInst 1                                      'yes, so add it
          i = FindEbrace()                                'find matching brace
          If i < 0 Then
              Errstr = "Cannot find a matching ending brace '}'"
          Else
            Instructions(i) = iENBrace                    'mark end of Enum definition
            EnDefFlg = True                               'Enum definition active
            BmpInd = True                                 'bump indenting
          End If
        Else
          Errstr = "Opening brace '{' expected"
        End If

' AdrOf ------------------------------------------------------------------------
      Case 362
        If CheckForLabel(InstrPtr, LabelWidth) Then
          ApndQTx
        Else
          Errstr = "Invalid parameter"
        End If

'-------------------------------------------------------------------------------
'      Case 363  ' Pcmp   'ignored
'      Case 364  ' Comp   'ignored

' Circle -----------------------------------------------------------------------
      Case 365
        Do  'use this loop for easy early exits
          If Not CheckForXY(Errstr) Then Exit Do      'check for (x,y)
          If CheckCommaValue(Errstr) Then Exit Do     'check for ,Radius
          If CheckCommaValue(Errstr) Then Exit Do     'check for ,Start
          If CheckCommaValue(Errstr) Then Exit Do     'check for ,End
          If CheckCommaValue(Errstr) Then Exit Do     'check for ,Aspect
          If Instructions(InstrPtr + 1) = iComma Then 'check for optional fill value
            ApndInst 1                                'add comma
            If CheckForNumber(InstrPtr, 1, 1) Then    'check for 0 (default) or 1
              ApndTxt                                 'add it if OK
            Else
              Errstr = "Invalid parameter"
            End If
          End If
          Exit Do
        Loop
' Split, Join ------------------------------------------------------------------
      Case 366, 367
        Select Case Instructions(InstrPtr + 1)
          Case 1 To 9
            ApndInst 1
          Case iIND
            ApndInst 1
            If CheckForVar(Errstr) Then
              ApndTxt
            End If
          Case Else
            Errstr = "Invalid parameter"
        End Select
        If Len(Errstr) = 0 Then
          If Instructions(InstrPtr + 1) = iIND Then
            If Not CheckForVar(Errstr) Then Exit Do
          ElseIf CheckForLabel(InstrPtr + 1, LabelWidth) Then
            ApndQTx
          Else
            Errstr = "Invalid parameter"
          End If
        End If
' ReDim ------------------------------------------------------------------------
      Case 368
        Call CheckForVar(Errstr)

' Mid --------------------------------------------------------------------------
      Case 369
        Errstr = "Invalid parameter"
        If Instructions(InstrPtr + 1) <> iLparen Then Exit Do
        ApndInst 1
        If Not CheckForVar(Errstr) Then Exit Do 'target
        If Instructions(InstrPtr + 1) <> iComma Then Exit Do
        ApndInst 1
        If Not CheckForVar(Errstr) Then Exit Do 'position
        If Instructions(InstrPtr + 1) <> iComma Then Exit Do
        ApndInst 1
        If Not CheckForVar(Errstr) Then Exit Do 'length
        If Instructions(InstrPtr + 1) <> iRparen Then Exit Do
        ApndInst 1
        Errstr = vbNullString

'-------------------------------------------------------------------------------
'      Case 372  'Delse  'Ignored
'      Case Is > 900 'no need to check

'-------------------------------------------------------------------------------
    End Select
    
    If CBool(InstrErr) Then Exit Do    'if an error was already reported, skip out
    If CBool(Len(Errstr)) Then Exit Do 'if a non-reported error found, let's skip out of the loop
    InstrPtr = InstrPtr + 1           'bump the instruction pointer for next instruction
    
    If CBool(Len(InstTxt)) Then         'if data exists
'      Debug.Print InstTxt
      InstFmt(InstCnt) = String$(IndLvl * IndSpc, 32) & InstTxt     'stuff formatted line
      InstFmt3(InstCnt) = InstFmt(InstCnt)
'
' adjust indenting, if required
'
      If BmpInd Then
        BmpInd = False
        IndLvl = IndLvl + 1             'adjust indenting level out
      End If
      If DecInd Then
        DecInd = False
        If CBool(IndLvl) Then
          IndLvl = IndLvl - 1           'adjust indenting level in
        End If
      End If
    Else
      InstCnt = InstCnt - 1             'null data, so back up instruction index
      Code = 128                        'make skipable code
    End If
'
' For extended style formatting, see if we can merge some lines
'
    Select Case Code                    'check current code
      Case iDBG, iNOP, 128              'do nothing
      Case iWhile, iUntil
        Select Case PrvCode
          Case iDWBrace, iDUBrace
            ApndPrv True              'append current instruction(s) to the previous with space sep.
        End Select
      '÷, x, -, +, =, &, |, ^, %, >>, <<, <, >, +/-, (case) Else, Pi, e, <=, >=, ==, !=
      Case 171, 184, 197, 210, 223, 161, 174, 200, 213, 226, 354, 274, 275, 222, _
           iCaseElse, iPi, iEp, iEQ, iNEQ, iLE, iGE
        ApndPrv True                'append current instruction(s) to the previous with space sep.
      '(, ), ;, ], :, various special end parens
      Case iLparen, iRparen, iUparen, iIparen, iWparen, iFparen, iCparen, iEparen, iSparen, _
           iCCBrace, iSemiC, iSemiColon, iComma, iRbrkt, iColon, iDWparen
        ApndPrv False               'append current instruction(s) to the previous with no space sep.
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
            ApndPrv False             'append current instruction(s) to the previous with no space sep.
          Case Else
            ApndPrv True              'append current instruction(s) to the previous with a space sep.
        End Select
      Case iLCbrace '{
        Select Case PrvCode
          Case iCparen, iCaseElse     'ignore if previous was ) or Else for Case
          Case Else
            ApndPrv True              'append current instruction(s) to the previous with a space sep.
        End Select
      Case Else
        Select Case PrvCode         'else check previous instruction
          Case iLparen
            ApndPrv False
          '÷, x, -, +, ÷=, x=, -=, += ,^, !, %, ~, &, |, &&, >>, <<, \, <, >, various special end parens
          Case 171, 184, 197, 210, 299, 312, 325, 338, 200, 315, 213, 187, 161, _
               174, 289, 302, 226, 354, 341, 274, 275
            ApndPrv True            'append current instruction(s) to it with space sep.
          'Nor, Dfn, Pvt, Pub, Else, [;], [,], LogX, Root, Sqrt, IND,[
          Case iNor, iDfn, iPvt, iPub, iElse, iSemiColon, iComma, 286, 355, 356, iIND, iLbrkt
            ApndPrv True            'append current instruction(s) to it, with a space separator
        End Select
    End Select
    
    InstCnt = InstCnt + 1             'point to next formatted text location
  Loop                                'process all data
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'
' do additional terminating error checks
'
  If Len(Errstr) = 0 And InstrErr = 0 Then
    If SbrDefFlg Then
      Errstr = "Subroutine/UserKey"
    ElseIf UtlDefFlg Then
      Errstr = "Do...Until()"
    ElseIf WhiDefFlg Then
      Errstr = "Do..While"
    ElseIf WhiDefFlg Or CBool(WhiIdx) Then
      Errstr = "While"
    ElseIf ForDefflg Or CBool(ForIdx) Then
      Errstr = "For"
    ElseIf SelDefFlg Or CBool(SelIdx) Then
      Errstr = "Select"
    ElseIf CaseDefFlg Or CBool(CaseIdx) Then
      Errstr = "Case"
    ElseIf IfDefFlg Or CBool(IfIdx) Then
      Errstr = "If"
    End If
    
    If CBool(Len(Errstr)) Then
      Errstr = Errstr & " block still in definition at the end of the program code"
    End If
  End If
  
  If CBool(Len(Errstr)) Or CBool(InstrErr) Then 'if ErrorFlag is set, then an error is already reported
    If Not CBool(InstrErr) Then                 'if we have not had an error reported, though one exists...
      ForcError Errstr
    End If
    Call ResetFmt                     'toss formatted data
    Call RenewLabels                  'reset default user key labels
    Call ResetBracing                 'reset bracing and paren codes in the program
  Else
    Preprocessd = True                'else we succeed. We can now run it
    ReDim Preserve InstFmt3(InstCnt - 1)
    ReDim Preserve InstMap3(InstCnt - 1)
    InstCnt3 = InstCnt
    Call CloseFmt                     'close up formatted text pool
    InstrPtr = 0                      'start program at beginning
  End If
  Call UpdateStatus
  TextEntry = False
  CharLimit = 0
  RedoAlphaPad
  Preprocessing = False               'no longer preprocessing
  
  If Not CBool(InstrErr) Then
    InstrPtr = HInstrPtr              'reset held values
    DisplayReg = HDisplayReg
    TestReg = HTestReg
  End If
'
' see if a Main() routine is defined. If so, execute it
'
  If Not Compressing Then
    If CBool(HaveMain) Then
      InstrPtr = Lbls(HaveMain).LblDat + 1 'get code block address+1
      Call Run                            'run it
      If PmtFlag Then                     'if user prompting turned on...
        LastTypedInstr = iTXT             'set TXT command
        Call ActiveKeypad                 'activate it (TXT or [=] (ENTER)) will start R/S cmd
      Else
        Call DisplayLine                  'else terminating run...
      End If
    End If
  End If
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

