Attribute VB_Name = "modRunSupport"
Option Explicit

'*******************************************************************************
' Function Name     : FindVblMatch
' Purpose           : Search for a matching Variable name in the Variables() list
'*******************************************************************************
Public Function FindVblMatch(Vbl As String) As Integer
  Dim Idx As Integer
  Dim Vrbl As String * LabelWidth
  
  Vrbl = Vbl                                                    'get space-padded reference
  For Idx = 0 To MaxVar                                         'scan all base, named entries
    If StrComp(Variables(Idx).VName, Vrbl, vbTextCompare) = 0 Then 'found a match?
      FindVblMatch = Idx                                        'yes, return its index
      Exit Function
    End If
  Next Idx                                                      'scan all
  FindVblMatch = -1                                             'no match found
End Function

'*******************************************************************************
' Function Name     : FindDefMatch
' Purpose           : Search for a matching definition in the DefName() list
'*******************************************************************************
Public Function FindDefMatch(Txt As String) As Integer
  Dim Idx As Integer

  For Idx = 1 To DefCnt
    If StrComp(Txt, DefName(Idx), vbTextCompare) = 0 Then       'found a match?
      FindDefMatch = Idx                                        'yes, so return it
      Exit Function
    End If
  Next Idx
  FindDefMatch = 0                                              'no match found
End Function

'*******************************************************************************
' Subroutine Name   : CheckSngVbl
' Purpose           : Check for expected variable, but also check for IND
'                   : No checks for dimensioning done here.
'*******************************************************************************
Public Function CheckSngVbl() As clsVarSto
  Dim Vn As Integer
  
  If GetInstruction(1) = iIND Then IncInstrPtr          'skip IND (used to indicate variable, not const)
  If GetInstruction(1) = iVar Then IncInstrPtr          'skip Var, if present
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
  Set CheckSngVbl = Variables(Vn).Vdata                 'point to base variable
End Function

'*******************************************************************************
' Subroutine Name   : CheckRunDim
' Purpose           : Check Dimensioning. Return the dimension index (0-99)
'                   : This is used to check for each dimension. '[' is assumed.
'                   : Although IND variables can be used in place of absolute
'                   : array offsets, the IND variable cannot contain its own
'                   : dimensioning references.
'*******************************************************************************
Private Function CheckRunDim() As Long
  Dim Vptr As clsVarSto
  Dim Vn As Long
  
  CheckRunDim = -1                            'init to invalid value
  Vn = -1
  Call IncInstrPtr                            'bump instruction pointer to '['
  If GetInstruction(1) = iIND Then            'check for IND (but sub-dimensioning not allowed)
    Set Vptr = CheckSngVbl()                  'get base variable storage object
    If Vptr Is Nothing Then Exit Function     'if nothing, then error
    Call IncInstrPtr                          'point to ']' (we know it is there)
    On Error Resume Next                      'trap invalid conversions
    Vn = CLng(ExtractValue(Vptr))             'grab value there (should be 0-99)
    If CBool(Err.Number) Or Vn < 0 Or Vn > 99 Then 'error during conversion?
      ForcError "Invaid Array Dimension value"
    Else
      CheckRunDim = Vn                        'else report success by returning valid value
    End If
  Else
    If CheckForNumber(InstrPtr, 2, 99) Then   'assume 0-99 dimension
      Call IncInstrPtr                        'point to ']' (we know it is there)
      CheckRunDim = CInt(TstData)             'grab dimension number
    End If
  End If
End Function
  
'*******************************************************************************
' Subroutine Name   : CheckRunDim2
' Purpose           : Check for possible 1-D or 2-D array designation
'                   : On return, if result is True, then
'                   :    Xd = X array Dim (-1 if not used)
'                   :    Yd = Y array Dim (-1 if not used)
'*******************************************************************************
Public Function CheckRunDim2(Xd As Long, Yd As Long) As Boolean
  Xd = -1                                 'init to not used
  Yd = -1
  If GetInstruction(1) = iLbrkt Then      'bracket following variable used?
    Xd = CheckRunDim()                    'check X Dim reference
    If Xd <> -1 Then                      'if it was OK
      If GetInstruction(1) = iLbrkt Then  'a second bracket?
        Yd = CheckRunDim()                'check Y Dim reference
      End If
    End If
  Else
    Exit Function
  End If
  CheckRunDim2 = Not ErrorFlag            'success if no errors reported
End Function

'*******************************************************************************
' Subroutine Name   : CheckVbl
' Purpose           : Check for expected variable, but also check for IND.
'                   : Also check for 1D or 2D array. Return Pointer to Desired object.
'                   : Note here that IND simply means that a variable is used
'                   : instead of an absolute value. Use CheckIndVar() if you
'                   : want to check for an Indirected variable (the provided
'                   : variable contains the variable number to actually use)
'*******************************************************************************
Public Function CheckVbl() As clsVarSto
  Dim Xd As Long, Yd As Long
  Dim Vptr As clsVarSto
  
  Set Vptr = CheckSngVbl()                      'get variable reference
  If Vptr Is Nothing Then Exit Function         'error occurred
  If CheckRunDim2(Xd, Yd) Then                  'if no errors for dim check...
    Set Vptr = PntToVptr(Vptr.VarRoot, Xd, Yd)  'point to exact storage object
  End If
  Set CheckVbl = Vptr                           'return exact variable storage object
End Function

'*******************************************************************************
' Function Name     : GrabInd
' Purpose           : Get Indirected variable storage reference
'                   : (IND is assumed, if not actually present)
'*******************************************************************************
Public Function GrabInd() As clsVarSto
  Dim Vptr As clsVarSto
  Dim TV As Double
  
  Set Vptr = CheckVbl()                   'get storage object
  If Vptr Is Nothing Then Exit Function   'if error
  TV = Fix(Val(ExtractValue(Vptr)))       'now get indirected variable number
  If TV < 0# Or TV > DMaxVar Then         'variable value in range?
    ForcError "Invalid Indirection value" 'no, so error
    Exit Function
  End If
  Set GrabInd = Variables(CInt(TV)).Vdata 'else return storage object for result
End Function

'*******************************************************************************
' Function Name     : CheckIndVar
' Purpose           : Get assumed Variable, but allow for IND reference to it.
'                   : IND referencing in this case means that the provided variable
'                   : will contain the actual variable number to ultimately use.
'*******************************************************************************
Public Function CheckIndVar() As clsVarSto
  If GetInstruction(1) = iIND Then  'indirection used?
    Set CheckIndVar = GrabInd()     'return storage object for final result
  Else
    Set CheckIndVar = CheckVbl()    'else variable, so get storage object
  End If
End Function

'*******************************************************************************
' Subroutine Name   : CE_RunSupport
' Purpose           : Support CE command in Run Mode
'*******************************************************************************
Public Sub CE_RunSupport()
  flags(7) = False                  'turn off error flag indicator
  ErrorFlag = False                 'turn off error flag
  CharCount = 0                     'reset character count
  CharLimit = 0
  AllowExp = False                  'disable allowing exponent entry
  REMmode = 0                       'ensure remarks mode is off
  DisplayText = False
  StoreList = False
  ModuleList = False
End Sub

'*******************************************************************************
' Subroutine Name   : CLR_RunSupport
' Purpose           : clear display and display reset line
'*******************************************************************************
Public Sub CLR_RunSupport()
  frmVisualCalc.PicPlot.Visible = False 'turn off plot view
  PlotTrigger = False                   'disable plot picking, if active
  EEMode = False                        'Reset EE Mode
  EngMode = False                       'Reset Eng Mode
  If Not Tron Then                      'if we are not in Trace Mode...
    Call Clear_Screen                   'reset display screen
  End If
  Call CE_RunSupport                    'clean up other registers
  DisplayReg = 0#                       'reset display register
  PushIdx = 0                           'remove items from the Push Stack
End Sub

'*******************************************************************************
' Subroutine Name   : STO_Run
' Purpose           : Support Run-Time STO and STO IND
'*******************************************************************************
Public Sub STO_Run(Vptr As clsVarSto)
  If Not Vptr Is Nothing Then                   'if object defined....
    If DisplayText Then                         'if text is active
      If Variables(Vptr.VarRoot).VarType = vString Then
        Vptr.VarStr = DspTxt                    'if string, stuff text
      Else
        If IsNumeric(DspTxt) Then               'if text data is actually numeric...
          Call StuffValue(Vptr, CVar(DspTxt))   'if string, stuff text
        Else
          Call StuffValue(Vptr, CVar(0))        'else stull zero
        End If
      End If
    Else
      Call StuffValue(Vptr, CVar(DisplayReg))   'numeric, so stuff it as number or Str()
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : RCL_run
' Purpose           : Support Run-Time RCL and RCL IND
'*******************************************************************************
Public Sub RCL_run(Vptr As clsVarSto)
  If Not Vptr Is Nothing Then
    If Variables(Vptr.VarRoot).VarType = vString Then
      DspTxt = Vptr.VarStr
      DisplayText = True                    'DisplayLine will show DspTxt
    Else
      DisplayReg = CDbl(ExtractValue(Vptr))
      DspTxt = CStr(DisplayReg)
      DisplayText = False                   'DisplayLine will show DisplayReg
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : EXC_Run
' Purpose           : Support Run-Time EXC and EXC IND
'*******************************************************************************
Public Sub EXC_Run(Vptr As clsVarSto)
  Dim TV As Double
  
  If Not Vptr Is Nothing Then
    If Variables(Vptr.VarRoot).VarType = vString Then
      ForcError "This operation cannot be performed with Text variables"
    Else
      TV = DisplayReg                       'value to exchange
      DisplayReg = Val(ExtractValue(Vptr))  'get value from variable
      Call StuffValue(Vptr, CVar(TV))       'excange values
    End If
    DisplayText = False
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : SUM_Run
' Purpose           : Support Run-Time SUM and SUM IND
'*******************************************************************************
Public Sub SUM_Run(Vptr As clsVarSto)
  Dim TV As Double
  
  If Not Vptr Is Nothing Then
    If Variables(Vptr.VarRoot).VarType = vString Then 'if text, merge variable with DspData\
      Vptr.VarStr = Vptr.VarStr & DspTxt
    Else
      TV = DisplayReg                       'value to exchange
      TV = TV + Val(ExtractValue(Vptr))     'get value from variable and add to DisplayReg
      Call StuffValue(Vptr, CVar(TV))       'stuff SUM
    End If
    DisplayText = False
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : MUL_Run
' Purpose           : Support Run-Time MUL and MUL IND
'*******************************************************************************
Public Sub MUL_Run(Vptr As clsVarSto)
  Dim TV As Double
  
  If Not Vptr Is Nothing Then
    If Variables(Vptr.VarRoot).VarType = vString Then
      ForcError "This operation cannot be performed with Text variables"
    Else
      TV = DisplayReg                       'value to exchange
      TV = TV * Val(ExtractValue(Vptr))     'get value from variable and add to DisplayReg
      Call StuffValue(Vptr, CVar(TV))       'stuff MUL
    End If
    DisplayText = False
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : SUB_Run
' Purpose           : Support Run-Time SUB and SUB IND
'*******************************************************************************
Public Sub SUB_Run(Vptr As clsVarSto)
  Dim TV As Double
  
  If Not Vptr Is Nothing Then
    If Variables(Vptr.VarRoot).VarType = vString Then
      ForcError "This operation cannot be performed with Text variables"
    Else
      TV = Val(ExtractValue(Vptr)) - DisplayReg   'get value from variable and add to DisplayReg
      Call StuffValue(Vptr, CVar(TV))             'stuff SUB
    End If
    DisplayText = False
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : DIV_Run
' Purpose           : Support Run-Time DIV and DIV IND
'*******************************************************************************
Public Sub DIV_Run(Vptr As clsVarSto)
  Dim TV As Double
  
  If Not Vptr Is Nothing Then
    If Variables(Vptr.VarRoot).VarType = vString Then
      ForcError "This operation cannot be performed with Text variables"
    Else
      On Error Resume Next
      TV = Val(ExtractValue(Vptr)) / DisplayReg 'get value from variable and add to DisplayReg
      Call CheckError
      On Error GoTo 0
      If Not ErrorFlag Then
        Call StuffValue(Vptr, CVar(TV))         'stuff DIV result
      End If
    End If
    DisplayText = False
  End If
End Sub

'*******************************************************************************
' Function Name     : Trim_Run
' Purpose           : Return a valid variable object for text trimming
'*******************************************************************************
Public Function Trim_Run() As clsVarSto
  Dim Vptr As clsVarSto
  
  Set Vptr = CheckIndVar()
  If Vptr Is Nothing Then Exit Function
  With Variables(Vptr.VarRoot)
    If .VarType <> vString Then
      ForcError "Variable must be a Text type"
    ElseIf CBool(.VdataLen) Then
      ForcError "A fixed-length string cannot be trimmed"
    Else
      Set Trim_Run = Vptr
    End If
  End With
End Function

'*******************************************************************************
' Subroutine Name   : BuildBraceStk
' Purpose           : Bruild a new brace level in the brace pool
'*******************************************************************************
Public Sub BuildBraceStk(ByVal OpnBrc As Integer, _
                         ByVal Prc As Integer, _
                         ByVal Cnd As Integer, _
                         Optional Truth As Boolean = False)
  BraceIdx = BraceIdx + 1               'bump index for brace pool
  If BraceIdx > BraceSize Then          'if we exceeded current pool size
    BraceSize = BraceSize + BraceDepth  'bump the pool size
    ReDim BracePool(BraceSize)
  End If
  
  With BracePool(BraceIdx)              'now add data to new brace item
    .LpStart = OpnBrc                   'start of data (location of [{])
    .LpTerm = FindEblock(OpnBrc)        'find termination of block
    .LpProcess = Prc                    'process location-1 (-1 if not used)
    .LpCond = Cnd                       'conditional location-1 (-1 if not used)
    .LpTrue = Truth                     'conditional result used by If blocks
    .LpDspReg = DisplayReg              'save current value of Displayregister
    .LpLoop = False                     'this will later be set True by actual looping routines
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : ProcessEparen
' Purpose           : Process End Parens ')' as they are encountered
'*******************************************************************************
Public Sub ProcessEparen(ByVal Paren As Integer)
  Select Case Paren
    '===================================
    Case iDWparen     'Do{...}While(..)  end paren (Brace block already defined)
      With BracePool(BraceIdx)
        If CBool(DisplayReg) Then       'if expression is True...
          DisplayReg = .LpDspReg        'reset DisplayRegister to saved value
          InstrPtr = .LpStart           'set beginning location of [{]
        Else
          DisplayReg = .LpDspReg        'reset DisplayRegister to saved value
          BraceIdx = BraceIdx - 1       'else back off brace (ptr already at termination point)
        End If
      End With
    '-----------------------------------
    Case iSparen      'Select (..){      end paren (Brace block already defined)
      With BracePool(BraceIdx)
        .LpSelect = DisplayReg          'save expression result looked for in select block
        IncInstrPtr                     'point to [{]
        DisplayReg = .LpDspReg          'reset DisplayRegister to saved value
      End With
    '-----------------------------------
    Case iUparen      'Do{...}Until(..)  end paren (Brace block already defined)
      With BracePool(BraceIdx)
        If Not CBool(DisplayReg) Then   'if expression is Not True...
          DisplayReg = .LpDspReg        'reset DisplayRegister to saved value
          InstrPtr = .LpStart           'set beginning location of [{]
        Else
          DisplayReg = .LpDspReg        'reset DisplayRegister to saved value
          BraceIdx = BraceIdx - 1       'else back off brace (ptr already at termination point)
        End If
      End With
    '-----------------------------------
    Case iCparen      'Case (..) {       end paren (Brace block already defined)
      With BracePool(BraceIdx)
        If DisplayReg = .LpSelect Then  'if expression matches needed value
          IncInstrPtr                   'point to '{' of case block
        Else
          InstrPtr = .LpTerm            'if no match, then point to end of case block
        End If
        DisplayReg = .LpDspReg          'reset DisplayRegister to saved value
        BraceIdx = BraceIdx - 1         'back off to select block index regardless
      End With
    '-----------------------------------
    Case iIparen      'If(..){           end paren (Brace block already defined)
      With BracePool(BraceIdx)
        .LpTrue = CBool(DisplayReg)     'set truth of data
        DisplayReg = .LpDspReg          'reset DisplayRegister to saved value
      End With
    '-----------------------------------
    Case iWparen      'While(..) {       end paren (Brace block already defined)
      With BracePool(BraceIdx)
        If CBool(DisplayReg) Then       'if expression is True
          IncInstrPtr                   'point to [{], beginning of block
          DisplayReg = .LpDspReg        'reset DisplayRegister to saved value
        Else
          InstrPtr = .LpTerm            'else point to termination of block
          DisplayReg = .LpDspReg        'reset DisplayRegister to saved value
          BraceIdx = BraceIdx - 1       'remove block
        End If
      End With
    '-----------------------------------
    Case iFparen      'For (..) {        end paren (Brace block already defined)
      With BracePool(BraceIdx)
        If .LpCond = -1 Then            'if no Conditional, assume TRUE
          InstrPtr = .LpStart           'point back to start
          DisplayReg = .LpDspReg        'recover Display Register
        Else
          InstrPtr = .LpCond            'else point to conditional
          Call Pend(iLparen)            'set up for conditional
        End If
      End With
    '===================================
  End Select
End Sub

'*******************************************************************************
' Subroutine Name   : ProcessEbrace
' Purpose           : Process End Braces '}' as they are encountered
'*******************************************************************************
Public Sub ProcessEbrace(ByVal Brc As Integer, ErrorStr As String)
  Dim Idx As Long
  
  Select Case Brc
    '========================
'    Case iRCbrace  '} this and others will never be encountered here...
'    Case iEIBrace  '}        'special end brace for Enum item blocks
'    Case iENBrace  '}        'special end brace for Enum blocks
'    Case iSIBrace  '}        'special end brace for Struct item blocks
'    Case iSTBrace  '}        'special end brace for Struct blocks
'    Case iCNBrace  '}        'Special end brace for Const blocks
    '------------------------
    Case iDWBrace, iDUBrace, iWCBrace 'special end braces for Do-While, Do-Until, and While(){..} blocks
      With BracePool(BraceIdx)
        If .LpCond = -1 Then
          InstrPtr = .LpStart     'point back to start if no condition
        Else
          .LpDspReg = DisplayReg  'save DisplayReg
          InstrPtr = .LpCond      'point to condition and process it ('('-1)
        End If
      End With
    '------------------------
    Case iICBrace  '}             'special end brace for If blocks
      With BracePool(BraceIdx)
        Select Case GetInstruction(1)
          Case iElse, iElseIf     'if else or ElseIf, they will use the brace location and set to defs
          Case Else
            BraceIdx = BraceIdx - 1 'dec brace index (we are done with braceing for this IF block
        End Select
      End With
    '------------------------
    Case iDCBrace  '}             'special end brace for Do blocks
      With BracePool(BraceIdx)
        InstrPtr = .LpStart       'point back to start
      End With
    '------------------------
    Case iFCBrace  '}             'special end brace for For blocks
      With BracePool(BraceIdx)
        .LpDspReg = DisplayReg    'save DisplayReg
        If .LpProcess = -1 Then   'if Process is not used...
          InstrPtr = .LpStart     'simply point back to start
        Else
          InstrPtr = .LpProcess   'else process
          Call Pend(iLparen)      'prepare for processing
        End If
      End With
    '------------------------
    Case iSCBrace  '}             'special end brace for Select blocks
      With BracePool(BraceIdx)
        BraceIdx = BraceIdx - 1   'already pointing where we need to, so just dec brace index
      End With
    '------------------------
    Case iCCBrace  '}             'special end brace for Case blocks
      With BracePool(BraceIdx)
        InstrPtr = .LpTerm        'point to Select's terminator if we encounter this
        BraceIdx = BraceIdx - 1   'decrement brace index (we were pointing to Select brace block)
      End With
    '------------------------
    Case iBCBrace  '}             'special end brace that acts like a RTN command for Sbr and Ukey
      If CBool(SbrInvkIdx) Then   'if we called a subroutine
        Idx = SbrInvkIdx          'copy current stack location
        SbrInvkIdx = SbrInvkIdx - 1 'decrement invoke
        With SbrInvkStk(Idx)
          ActivePgm = .Pgm        'set the previous active program, in case different
          BraceIdx = .PgmBrcIdx   'clear brace index to what it was before (quick clean stack)
          If Not .PgmLiveInvk Then 'if this routine was not invoked from live keyboard
            InstrPtr = .PgmInst   'set return location from invoke to the previous location
            Exit Sub              'done (SbrInvkIdx=0 will fall below...)
          End If
        End With
      ElseIf CBool(BraceIdx) Then 'if any bracing exists...
        BraceIdx = BraceIdx - 1   'decrement brace index (we are pointing to end of block)
      End If
      RunMode = False             'turn off run mode
      ModPrep = 0                 'turn from prep
      Call ResetPndAll            'reset data
      StopMode = Not flags(9)     'set stop mode if flag 9 not set
    '========================
  End Select
End Sub

'*******************************************************************************
' Subroutine Name   : PushCall
' Purpose           : Push a location onto the call stack
'*******************************************************************************
Public Sub PushCall(ByVal SbrIndex As Long, ByVal Newptr As Integer, ByVal Pgm As Integer, Optional LiveInvk As Boolean = False)
  Dim Idx As Integer, i As Integer
  Dim S As String
  
  SbrInvkIdx = SbrInvkIdx + 1               'point to new sbr stack location
  If SbrInvkIdx > SbrInvkSize Then          'if beyond pool
    SbrInvkSize = SbrInvkSize + 16          'bump new size
    ReDim Preserve SbrInvkStk(SbrInvkSize)  'set new pool size
  End If
  
  With SbrInvkStk(SbrInvkIdx)
    .Pgm = Pgm                              'save program invoking call
    .PgmInst = InstrPtr                     'save instruction to return to (next-1)
    .SbrIdx = SbrIndex                      'set index into Lbls array
    .PgmBrcIdx = BraceIdx                   'save brace index
    .PgmLiveInvk = LiveInvk                 'if routine invoked from keyboard
    Call BuildBraceStk(Newptr, -1, -1, True) 'build brace stack for current subroutine
    InstrPtr = Newptr                       'now set instruction pointer to new destination
                                            'next instruction after PushCall() will be Call Run(),
                                            'so we will execute the current location's instruction.
'
' if SbrIndex was not supplied (Call IND, for example), then we will need to find
' the reference by tracking backward through the code for the Sbr or Ukey tokens.
'
    If SbrIndex = 0 Then                    'if index was not defined...
      If CBool(ActivePgm) Then
        For Idx = ModLblMap(ActivePgm - 1) To ModLblMap(ActivePgm) - 1
          If ModLbls(Idx).lblAddr = Newptr Then
            .SbrIdx = Idx                   'set definition Index
            Exit For
          End If
        Next Idx
      Else
        For Idx = 1 To LblCnt - 1
          If Lbls(Idx).lblAddr = Newptr Then
            .SbrIdx = Idx
            Exit For
          End If
        Next Idx
      End If
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : FileIO
' Purpose           : Handle File Read, Write, Get, Put
'*******************************************************************************
Public Sub FileIO(ByVal Code As Integer, Errstr As String)
  Dim Rd As Boolean   'True if Read, False if Write
  Dim Blk As Boolean  'True if Block, False if Stream
  Dim Iptr As Integer, J As Integer, Iint As Integer, iTs As Integer, i As Integer
  Dim S As String
  Dim TV As Double, iDbl As Double
  Dim Vptr As clsVarSto
  Dim Vt As Vtypes
  Dim iLng As Long, Rcd As Long, JL As Long
  Dim Idx As Long, Ofst As Long, Sz As Long
  Dim iByte As Byte
  Dim Pool() As StructPool
'
' first set I/O flags
'
  Select Case Code
    Case iRead
      Blk = False 'stream
      Rd = True   'reading
    Case iWrite
      Blk = False 'streaming
      Rd = False  'writing
    Case iGet
      Blk = True  'block I/O
      Rd = True   'reading
    Case iPut
      Blk = True  'block I/O
      Rd = False  'writing
  End Select
'
' get which port to read/write
'
  iTs = GetInstruction(1)               'check next instruction
  Select Case iTs
    Case 1 To 9                         'allow ports 1-9
      IncInstrPtr                       'bump instruction pointer
    Case iIND                           'allow specifying an indirect variable
      Set Vptr = CheckVbl()
      If Vptr Is Nothing Then Exit Sub
      TV = Val(ExtractValue(Vptr))      'get value
      If TV < 1# Or TV > 9# Then
        Errstr = "Parameter is out of range (1-9)"
        Exit Sub
      End If
      iTs = CInt(TV)                    'get file port number
    Case Else
      Errstr = "Invalid parameter"
      Exit Sub
  End Select
'
' verify the file buffer is open
'
  With Files(iTs)
    If .FileNum = 0 Then
      Errstr = "File buffer " & CStr(iTs) & " is not open"
      Exit Sub
    End If
'
' cannot stream blocked data, and cannot block stramed data
'
    If CBool(.FileLen) And Not Blk Then
      Errstr = "Cannot Read/Write Block data. Use Get/Put"
      Exit Sub
    ElseIf Not CBool(.FileLen) And Blk Then
      Errstr = "Cannot Get/Put stream data. Use Read/Write"
      Exit Sub
    End If
'
' for Block I/O, allow specifying a record number after a comma
'
  If Blk Then
    If GetInstruction(1) = iComma Then      'comma present?
      IncInstrPtr                           'prepare for next
      If GetInstruction(1) = iIND Then      'Indirection?
        Set Vptr = CheckVbl()               'yes, get variage storage
        If Vptr Is Nothing Then Exit Sub    'oops
        TV = Val(ExtractValue(Vptr))        'get data
      Else
        Call CheckForNumber(InstrPtr, 5, 32767)
        TV = TstData                        'get numeric input
      End If
      If TV < 1# Or TV > 32767# Then        'in range?
        Errstr = "Parameter is out of range (1-32767)"
        Exit Sub
      End If
      Rcd = CLng(TV)                        'set record number
    Else
      Rcd = .FileRec                        'no comma, so get current record
    End If
    .FileRec = Rcd + 1                      'bump record number, if not specified next time
  End If
'
' see what to process the data through
'
    IncInstrPtr                           'bump instruction pointer
    Select Case GetInstruction(0)         'instruction will be ALL or WITH
      Case iAll                           'read/write whole file via buffer?
        Select Case Code
          '----------
          Case iRead
           On Error Resume Next
           .FileBufT = Tstrm(iTs).ReadAll 'read everything
           Call CheckError
           On Error GoTo 0
           If ErrorFlag Then Exit Sub
          '----------
          Case iWrite
           On Error Resume Next
            Tstrm(iTs).Write .FileBufT    'write everything
           Call CheckError
           On Error GoTo 0
           If ErrorFlag Then Exit Sub
          '----------
          Case Else
            Errstr = "Invalid parameter"
            Exit Sub
        End Select
      '-----------
      Case iWith                     'allow With to fall through
      '-----------
      Case Else
        Errstr = "Expected 'With' token"
        Exit Sub
    End Select
    
    Iptr = InstrPtr                   'get instruction pointer at 'With'
    If GetInstruction(1) = iVar Then Iptr = Iptr + 1  'skip Var, if present
    Call CheckForAlNum(Iptr, LabelWidth)
    If TstData = -1 Then              'if label specified
      J = FindVblMatch(TxtData)       'variable?
    Else
      J = CInt(TstData)               'else number, so get variable number
    End If
    If J <> -1 Then                   'is a variable
'-------------------------------------
'------- PROCESS VARIABLES -----------
'-------------------------------------
      InstrPtr = Iptr                   'reset pointer
      Do                                'begin variable list
        Set Vptr = CheckVbl()           'get variable target (collect array dimensioning also)
        If Vptr Is Nothing Then Exit Sub
        With Variables(Vptr.VarRoot)
          Vt = .VarType                 'save variable type
          i = .VdataLen                 'get possible string length
        End With
        Select Case Code
          Case iRead  '--------------------------------
            On Error Resume Next
            S = Tstrm(iTs).ReadLine     'read a line from file
            Call CheckError
            If ErrorFlag Then Exit Sub  'read error
            Select Case Vt
              Case vChar
                Vptr.VarChar = CByte(S) 'covert to a byte
              Case vInteger
                Vptr.VarInt = CLng(S)   'convert to a long
              Case vNumber
                Vptr.VarNum = CDbl(S)   'convert to a double
              Case vString
                If CBool(i) Then        'if fixed width...
                  If Len(S) < i Then S = S & String$(i - Len(S), 32)  'pad right side
                End If
                Vptr.VarStr = S         'convert to a string
            End Select
            Call CheckError
            On Error GoTo 0
            If ErrorFlag Then Exit Sub
          Case iWrite '--------------------------------
            Select Case Vt
              Case vChar
                S = CStr(Vptr.VarChar)  'convert all to strings
              Case vInteger
                S = CStr(Vptr.VarInt)
              Case vNumber
                S = CStr(Vptr.VarNum)
              Case vString
                S = Vptr.VarStr
            End Select
            On Error Resume Next
            Tstrm(iTs).WriteLine S      'write string to a line
            Call CheckError
            On Error GoTo 0
            If ErrorFlag Then Exit Sub
          Case iGet '-----------------------------------
            On Error Resume Next
            Select Case Vt
              Case vChar
                Get #.FileNum, Rcd, iByte       'read to a byte variable
                Vptr.VarChar = iByte            'stuff to Variable class
              Case vInteger
                Get #.FileNum, Rcd, iLng
                Vptr.VarInt = iLng
              Case vNumber
                Get #.FileNum, Rcd, iDbl
                Vptr.VarNum = iDbl
              Case vString
                S = String$(.FileLen, 32)
                Get #.FileNum, Rcd, S
                Vptr.VarStr = S
            End Select
            Call CheckError
            On Error GoTo 0
            If ErrorFlag Then Exit Sub
          Case iPut '-----------------------------------
            On Error Resume Next
            Select Case Vt
              Case vChar
                Put #.FileNum, Rcd, Vptr.VarChar
              Case vInteger
                Put #.FileNum, Rcd, Vptr.VarInt
              Case vNumber
                Put #.FileNum, Rcd, Vptr.VarNum
              Case vString
                Put #.FileNum, Rcd, Vptr.VarStr
            End Select
            Call CheckError
            On Error GoTo 0
            If ErrorFlag Then Exit Do       'if error, then exit loop
        End Select
        If GetInstruction(1) = iComma Then  'another variable coming?
          IncInstrPtr                       'yes, so prepare for it
          If Blk Then                       'if blocked I/O...
            Rcd = Rcd + 1                   'bump target record number
            .FileRec = Rcd + 1              'establish next record for default record next time
          End If
        Else
          Exit Do                           'else we are done with list
        End If
      Loop
      Exit Sub
    End If
'-------------------------------------
'------- PROCESS STRUCTURES ----------
'-------------------------------------
    JL = FindLbl(TxtData, TypStruct)  'find Label for type Struct
    If JL = 0 Then                    'oops, did not find it
      Errstr = "Invalid parameter"
      Exit Sub
    End If
    If CBool(ActivePgm) Then
      JL = ModLbls(JL).LblValue       'get its index into StructPl()
      Pool = ModStPl                  'point to module pool
    Else
      JL = Lbls(JL).LblValue          'get its index into StructPl()
      Pool = StructPl                 'point to local pool
    End If
    J = .FileLen                      'get I/O buffer size
'
' now ensure that blocked I/O has data length, and streaming does not
'
    Select Case Code
      Case iGet, iPut
        If J = 0 Then
          Errstr = "Get/Put requires record length specification"
          Exit Sub
        End If
      Case Else
        If CBool(J) Then
          Errstr = "Read/Write cannot process fixed-length records"
          Exit Sub
        End If
    End Select
  End With
'
' now process I/O to structure
'
  With Pool(JL)
    Select Case Code
      Case iGet
        Get #iTs, Rcd, .StBuf             'read to buffer
      Case iPut
        Put #iTs, Rcd, .StBuf             'write from buffer
      Case iRead
        For J = 0 To .StItmCnt            'process each structure item individually
          Ofst = .StItems(J).siOfst       'save offset within buffer
          Sz = .StItems(J).siLen          'and size of object
          S = Tstrm(iTs).ReadLine         'now read an item
          Select Case .StItems(J).siType
            Case vChar
              Mid$(.StBuf, Ofst + 1, 1) = Chr$(CByte(S))
            Case vString
              Mid$(.StBuf, Ofst + 1, Sz) = Left$(S, Sz)
            Case vInteger
              iLng = CLng(S)
              Call LngMkStr(iLng, .StBuf, Ofst)
            Case vNumber
              iDbl = CDbl(S)
              Call DblMkStr(iDbl, .StBuf, Ofst)
          End Select
        Next J
      Case iWrite
        For J = 0 To .StItmCnt
          Ofst = .StItems(J).siOfst       'save offset within buffer
          Sz = .StItems(J).siLen          'and size of object
          Select Case .StItems(J).siType
            Case vChar
              S = CStr(Asc(Mid$(.StBuf, Ofst + 1, 1)))
            Case vString
              S = Mid$(.StBuf, Ofst + 1, Sz)
            Case vInteger
              S = CStr(StrMkLng(.StBuf, Ofst))
            Case vNumber
              S = CStr(StrMkDbl(.StBuf, Ofst))
          End Select
          Tstrm(iTs).WriteLine S
        Next J
    End Select
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : SplitJoinCommon
' Purpose           : Common run-time support for the Split and Join tokens
'*******************************************************************************
Private Sub SplitJoinCommon(iTs As Integer, vNum As Long, Vptr As clsVarSto, Errstr As String)
  Dim TV As Double
  
  iTs = GetInstruction(1)               'check next instruction
  Select Case iTs
    Case 1 To 9                         'allow ports 1-9
      IncInstrPtr                       'bump instruction pointer
    Case iIND                           'allow specifying an indirect variable
      Set Vptr = CheckVbl()
      If Vptr Is Nothing Then Exit Sub
      TV = Val(ExtractValue(Vptr))      'get value
      If TV < 1# Or TV > 9# Then
        Errstr = "Invalid parameter"
        Exit Sub
      End If
      iTs = CInt(TV)                     'get file port number
    Case Else
      Errstr = "Invalid parameter"
      Exit Sub
  End Select
'
' see if the file port is open for business
'
  If Files(iTs).FileNum = 0 Then
    Errstr = "File buffer " & CStr(iTs) & " is not open"
    Exit Sub
  End If
'
' check base variable to process
'
  If GetInstruction(1) = iIND Then
    Set Vptr = CheckIndVar()        'check for variable and for indirection
  Else
    Set Vptr = CheckSngVbl()        'get variable reference (cannot be array)
  End If
  If Vptr Is Nothing Then Exit Sub
  vNum = Vptr.VarRoot               'get root variable
End Sub

'*******************************************************************************
' Subroutine Name   : Split_Support
' Purpose           : Provide run-time support for the Split token
'*******************************************************************************
Public Sub Split_Support(Errstr As String)
  Dim vNum As Long, iX As Long, Idx As Long
  Dim sPtr As clsVarSto, Vptr As clsVarSto
  Dim iTs As Integer
  Dim Ary() As String, S As String
'
' init the display register to invalid value
'
  DisplayReg = -1#
  DisplayText = False
'
' gather parameters provided to token
'
  ErrorFlag = False
  Call SplitJoinCommon(iTs, vNum, Vptr, Errstr)
  If CBool(Len(Errstr)) Or ErrorFlag Or Vptr Is Nothing Then Exit Sub
'
' build the local array
'
  Ary = Split(Files(iTs).FileBufT, vbCrLf)  'split array
  If Not IsDimmed(Ary) Then                 'if not dimmed, the error
    Errstr = "Split has no data to split"
    Exit Sub
  End If
  iX = UBound(Ary)                          'get upper bounds of array
'
' copy the array to the variable list
'
  With Variables(vNum)
    Set .Vdata = Nothing                    'reset variable
    Set .Vdata = New clsVarSto              'set new variable
  End With
  Call BuildMDAry(vNum, iX, -1, False)      'build new array
    
  ErrorFlag = False
  For Idx = 0 To iX                         'now fill variables withe the data
    Set sPtr = PntToVptr(vNum, Idx, -1)     'point to target object
    If sPtr Is Nothing Then Exit Sub        'error (unlikely)
    Call StuffValue(sPtr, CVar(Ary(Idx)))   'fill object with data
    If ErrorFlag Then Exit Sub
  Next Idx
'
' set the display register to the maximum dimension
'
  DisplayReg = CDbl(iX)
End Sub

'*******************************************************************************
' Subroutine Name   : Join_Support
' Purpose           : Provide run-time support for the Join token
'*******************************************************************************
Public Sub Join_Support(Errstr As String)
  Dim vNum As Long, iX As Long, Idx As Long
  Dim sPtr As clsVarSto, Vptr As clsVarSto
  Dim iTs As Integer
  Dim Ary() As String, S As String
'
' init the display register to invalid value
'
  DisplayReg = -1#
  DisplayText = False
'
' gather parameters provided to token
'
  ErrorFlag = False
  Call SplitJoinCommon(iTs, vNum, Vptr, Errstr)
  If CBool(Len(Errstr)) Or ErrorFlag Or Vptr Is Nothing Then Exit Sub
'
' size temp array to size of elements in array
'
  iX = Vptr.GetMaxDim()                   'get max dim
  ReDim Ary(iX)                           'size array to it
'
' extract variable array data to Ary()
'
  vNum = Vptr.VarRoot                     'set root variable
  ErrorFlag = False
  For Idx = 0 To iX                       'now fill variables withe the data
    Set sPtr = PntToVptr(vNum, Idx, -1)   'point to target object
    If sPtr Is Nothing Then Exit Sub      'error (unlikely)
    Ary(Idx) = CStr(ExtractValue(sPtr))   'build array
    If ErrorFlag Then Exit Sub
  Next Idx
'
' Join the array to the I/O buffer
'
  Files(iTs).FileBufT = Join(Ary, vbCrLf)
'
' set the display register to the maximum dimension
'
  DisplayReg = CDbl(iX)
End Sub

'*******************************************************************************
' Function Name     : CheckStruct
' Purpose           : If a structure Item, return it
'*******************************************************************************
Public Function CheckStruct() As clsVarSto
  Dim Sn As Long, Si As Long, Idx As Long
  Dim Pool() As StructPool
  Dim Iptr As Integer
  
  Iptr = InstrPtr                                       'save instruction pointer
  If CheckForLabel(InstrPtr, LabelWidth) Then           'label found?
    Sn = FindLbl(TxtData, TypStruct)                    'yes, is it a struct?
    If CBool(Sn) Then
      If GetInstruction(1) <> iDot Then                 'yes. Dot follows?
        ForcError "Structure must include member '.' operator" 'no, so error
        Exit Function
      End If
      IncInstrPtr                                       'account for dot
      If CheckForLabel(InstrPtr, LabelWidth) Then       'get structure item text
        If CBool(ActivePgm) Then                        'account for modules
          Sn = ModLbls(Sn).LblValue                     'structure pool index
          Pool = ModStPl                                'structure pool to use
        Else
          Sn = Lbls(Sn).LblValue
          Pool = StructPl
        End If
              
        With Pool(Sn)
          For Idx = 0 To .StItmCnt - 1                  'scan for matching structure item
            If StrComp(RTrim$(.StItems(Idx).SiName), TxtData, vbTextCompare) = 0 Then
              Variables(100).VarType = .StItems(Idx).siType 'found it, so save data type
              Variables(100).VdataLen = 0                   'ensure data length is null
              With Variables(100).Vdata
                .StPlIdx = Sn                               'save struct pool index
                .StItmIdx = Idx                             'and structure item matched
              End With
              Set CheckStruct = Variables(100).Vdata        'return variable object
              Exit Function
            End If
          Next Idx
        End With
      Else
        ForcError "Structure member error"
        Exit Function
      End If
    End If
  End If
  InstrPtr = Iptr
End Function

'*******************************************************************************
' Subroutine Name   : GTO_IND
' Purpose           : Support GTO IND
'*******************************************************************************
Public Sub GTO_IND(Vptr As clsVarSto)
  Dim TV As Double
  
  If Not Vptr Is Nothing Then                     'if object defined....
    TV = Fix(Val(ExtractValue(Vptr)))             'no, so grab value stored there
    If TV < 0# Or TV > CDbl(GetInstrCnt()) Then   'in range?
      ForcError "Indirection out of range"        'no
    Else
      InstrPtr = CInt(TV) - 1                     'else set new pointer to 1 less start point
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : StFlg_IND
' Purpose           : Support StFlg IND
'*******************************************************************************
Public Sub StFlg_IND(Vptr As clsVarSto)
  Dim TV As Double
  
  If Vptr Is Nothing Then Exit Sub
  TV = Fix(Val(ExtractValue(Vptr)))     'get value there
  If TV < 0# Or TV > 9# Then            'in range?
    ForcError "Parameter is out of range (0-9)"
  Else
    flags(CInt(TV)) = True              'it is, so set the flag
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : RFlg_IND
' Purpose           : Support StFlg IND
'*******************************************************************************
Public Sub RFlg_IND(Vptr As clsVarSto)
  Dim TV As Double
  
  If Vptr Is Nothing Then Exit Sub
  TV = Fix(Val(ExtractValue(Vptr)))     'get value there
  If TV < 0# Or TV > 9# Then            'in range?
    ForcError "Parameter is out of range (0-9)"
  Else
    flags(CInt(TV)) = False             'it is, so reset the flag
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : IfFlg_IND
' Purpose           : Support IfFlg IND
'*******************************************************************************
Public Sub IfFlg_IND(Vptr As clsVarSto)
  Dim TV As Double
  Dim i As Integer
  
  If Vptr Is Nothing Then Exit Sub
  TV = Fix(Val(ExtractValue(Vptr)))   'get value there
  If TV < 0# Or TV > 9# Then          'in range?
    ForcError "Parameter is out of range (0-9)"
  Else
    i = CInt(TV)                      'get flag number to process
    Call BuildBraceStk(InstrPtr, -1, -1, flags(i))
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : NotFlg_IND
' Purpose           : Support IfFlg IND
'*******************************************************************************
Public Sub NotFlg_IND(Vptr As clsVarSto)
  Dim TV As Double
  Dim i As Integer
  
  If Vptr Is Nothing Then Exit Sub
  TV = Fix(Val(ExtractValue(Vptr)))   'get value there
  If TV < 0# Or TV > 9# Then          'in range?
    ForcError "Parameter is out of range (0-9)"
  Else
    i = CInt(TV)                      'get flag number to process
    Call BuildBraceStk(InstrPtr, -1, -1, Not flags(i))
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Incr_IND
' Purpose           : Support Incr IND
'*******************************************************************************
Public Sub Incr_IND(Vptr As clsVarSto)
  Dim TV As Double
  
  If Not Vptr Is Nothing Then
    If Variables(Vptr.VarRoot).VarType = vString Then
      ForcError "This operation cannot be performed with Text variables"
    Else
      TV = CDbl(ExtractValue(Vptr)) + 1#      'add 1 to variable value
      Call StuffValue(Vptr, CVar(TV))         'stuff result to variable
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Decr_IND
' Purpose           : Support Decr IND
'*******************************************************************************
Public Sub Decr_IND(Vptr As clsVarSto)
  Dim TV As Double
  
  If Not Vptr Is Nothing Then
    If Variables(Vptr.VarRoot).VarType = vString Then
      ForcError "This operation cannot be performed with Text variables"
    Else
      TV = CDbl(ExtractValue(Vptr)) - 1#      'add 1 to variable value
      Call StuffValue(Vptr, CVar(TV))         'stuff result to variable
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Dsz_IND
' Purpose           : Support Dsz IND
'*******************************************************************************
Public Sub Dsz_IND(Vptr As clsVarSto)
  Dim TV As Double
  Dim Bol As Boolean
  
  If Not Vptr Is Nothing Then
    If Variables(Vptr.VarRoot).VarType = vString Then
      ForcError "This operation cannot be performed with Text variables"
    Else
      TV = CDbl(ExtractValue(Vptr))           'sub 1 from variable value
      TV = TV - 1#                            'decrement 1
      If TV < 0# Then TV = 0#
      Call StuffValue(Vptr, CVar(TV))         'stuff result to variable
      Bol = CBool(TV)                         'set flag (execute block until TV=0)
      Call BuildBraceStk(InstrPtr, -1, -1, Bol)
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Dsnz_IND
' Purpose           : Support Dsnz IND
'*******************************************************************************
Public Sub Dsnz_IND(Vptr As clsVarSto)
  Dim TV As Double
  Dim Bol As Boolean
  
  If Not Vptr Is Nothing Then
    If Variables(Vptr.VarRoot).VarType = vString Then
      ForcError "This operation cannot be performed with Text variables"
    Else
      TV = CDbl(ExtractValue(Vptr))           'sub 1 frm variable value
      TV = TV - 1#                            'decrement 1
      If TV <= 0# Then TV = 0#
      Call StuffValue(Vptr, CVar(TV))         'stuff result to variable
      Bol = Not CBool(TV)                     'set flag (skip execute block until TV=0)
      Call BuildBraceStk(InstrPtr, -1, -1, Bol)
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Var_IND
' Purpose           : Support Var IND
'*******************************************************************************
Public Sub Var_IND(Vptr As clsVarSto)
  If Vptr Is Nothing Then
    CurrentVar = -1                       'error
  Else
    CurrentVar = Vptr.VarRoot             'set root variable number (not <0)
    CurrentVarTyp = Variables(CurrentVar).VarType
    Set CurrentVarObj = Vptr              'keyboard mode uses ONLY base variables
    
    Call RCL_run(Vptr)                    'now treat like RCL
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : CloseAll
' Purpose           : Close All opened files ports
'*******************************************************************************
Public Sub CloseAll()
  Dim Idx As Integer
  
  For Idx = 1 To 9
    With Files(Idx)
      If CBool(.FileNum) Then   'file open?
        If Tstrm(Idx) Is Nothing Then
          Close #Idx            'is block I/O, so close it
        Else
          Tstrm(Idx).Close      'close it
          Set Tstrm(Idx) = Nothing
        End If
        .FileNum = 0            'mark as closed
      End If
    End With
  Next Idx
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

