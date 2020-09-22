Attribute VB_Name = "modRunCmd"
Option Explicit

'*******************************************************************************
' Subroutine Name   : RunCmd
' Purpose           : Process provided code to run-time processor
'*******************************************************************************
Public Sub RunCmd(ByVal Code As Integer, Errstr As String)
'===============================================================================
  Dim i As Integer, j As Integer, K As Integer
  Dim Iptr As Integer, OldPgm As Integer, Ln As Integer
  Dim Idx As Long, iX As Long, iY As Long, JL As Long, Vn As Long
  Dim Xs As Long, Ys As Long, Xe As Long, Ye As Long, nClr As Long, Clr As Long
  Dim Bol As Boolean
  Dim TV As Double, vDeg As Double, vMin As Double, vSec As Double
  Dim Radius As Single, ArcStart As Single, ArcEnd As Single, Aspect As Single
  Dim X As Double, Y As Double
  Dim S As String, T As String, Ary() As String, SS As String, Nm As String
  Dim Vptr As clsVarSto, sPtr As clsVarSto
  Dim Pool() As Labels
  Dim ts As TextStream
  Dim Tmp As Variant
  Dim Typ As Vtypes
'===============================================================================
  Iptr = InstrPtr                           'get the instruction pointer for the current location.
'
' see if instruction tracing is active
'
  If Tron Then
    For Idx = 0 To InstCnt3 - 1
      If InstMap3(Idx) = Iptr Then
        DisplayMsg Format(Iptr, "0000  ") & InstFmt3(Idx)
        Exit For
      ElseIf InstMap3(Idx) > Iptr Then
        Exit For
      End If
    Next Idx
  End If
  
  Select Case Code                          'process the provided instruction code. Although it might
                                            'seem to make sense to derive the code data at this
                                            'point, this allows this subroutine to be invoked from
                                            'within itself, such as for extended hyperbolic functions.
'-------------------------------------------------------------------------------

' 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, [.]---------------------------------------------
    Case 0 To 9, iDot
      Call CheckForValue(Iptr - 1)          'grab value (we know format is correct)
      DisplayReg = TstData                  'set display register to that value
      DisplayText = False                   'we will display DisplayReg, not DspTxt

' Ascii Text--------------------------------------------------------------------
    Case Is < 128                                 'text data
      JL = 0                                      'init flag
      If CheckForLabel(Iptr - 1, LabelWidth) Then 'first see if the text qualifies as a label
        JL = FindLbl(TxtData, TypEnum)            'if it does, see if it is an Enum
        If CBool(JL) Then                         'enum?
          If CBool(ActivePgm) Then                'yes, grab enum value
            DisplayReg = CDbl(ModLbls(JL).LblValue) 'grab module enum value
          Else
            DisplayReg = CDbl(Lbls(JL).LblValue)  'grab main program enum value
          End If
          DisplayText = False                     'set numeric display
          Exit Sub                                'and we are done
        End If
      End If
      Call CheckForText(Iptr - 1, DisplayWidth) 'extract to TxtData normally if not a constant
      DspTxt = TxtData                          'store found data as the pending text
      DisplayText = True                        'we will display DspTxt, not DisplayReg

'-------------------------------------------------------------------------------
      Case 128         'Ignore
      Case 129  ' LRN  'ignore

' Pgm --------------------------------------------------------------------------
    Case 130  ' Pgm
      DisplayText = False
      Do
        Select Case GetInstruction(1)                   'check the next instruction
          Case iLoad, iSave, iLapp, iASCII              'Pgm Load, Pgm Save, Pgm Lapp, or Pgm ASCII?
            Exit Do                                     'do nothing now. Let them take care of things
          Case iIND                                     'if we are using a variable (Pgm IND xx)
            Set Vptr = CheckVbl()                       'get variable object storage
            If Vptr Is Nothing Then Exit Do             'error
            TstData = Val(ExtractValue(Vptr))           'get value there to temp variable
          Case Else
            Call CheckForNumber(Iptr, 2, 99)            'else get absolute value to temp variable
        End Select
        
        If TstData = 0# Then                            'if selecting program 0 (User program)
          ActivePgm = 0                                 'set it
          ModPrep = 0
          Exit Do
        End If
        
        If TstData < 0# Or TstData > CDbl(ModCnt) Then  'check value range
          Errstr = "Invalid program number for module " & CStr(ModName)
          If ModName = 0 Then
            Errstr = Errstr & " (no module loaded)"
          Else
            Errstr = Errstr & " (1-" & CStr(ModCnt) & ")"
          End If
          Exit Do
        End If
        
        ModPrep = CInt(TstData)   'get program number (activated when Call or Ukey cmd is invoked, next)
        Exit Do
      Loop

' Load, Save, Lapp, ASCII ------------------------------------------------------
      Case iLoad, iSave, iLapp, iASCII
        Call CheckForNumber(Iptr, 2, 99)              'get program number to load
        If Len(StorePath) = 0 Then
          Errstr = "No storage path yet defined"
          Exit Sub
        End If
        T = Format(DisplayReg, "00")                  'format and save as 2-digit string
        SS = RemoveSlash(StorePath)
        S = SS & "\PGM\Pgm" & T                       'form path, less file extension
        j = FreeFile(0)
        Select Case Code
          '-----------------------------
          Case iLoad  ' Load  Load Binary
            If Not Fso.FileExists(S & ".pgm") Then
              Errstr = " Pgm" & T & ".pgm" & vbCrLf & _
                      " Program does not exist in path:" & vbCrLf & _
                      " " & SS & "\PGM"
            Else
              Call CP_Support                         'clean up memory and buffers
              On Error Resume Next
              Open S & ".pgm" For Random Access Read As #j Len = 2
              Call CheckError                         'check for errors
              On Error GoTo 0
              If Not ErrorFlag Then                   'if no errors
                InstrCnt = LOF(j) / 2                 'get # of instructions
                If InstrCnt > InstrSize Then          'if we will need to bump our pool
                  Do While InstrCnt > InstrSize
                    InstrSize = InstrSize + InstrInc  'increment by offset increment
                  Loop
                  ReDim Preserve Instructions(InstrSize) 'resize pool
                End If
                For Idx = 0 To InstrCnt - 1
                  Get #j, , Instructions(Idx)         'get all new instructions
                Next Idx
                Call ResetBracing                     'reset any special bracing in program
                DisplayMsg "Loaded Binary Pgm" & T & ".pgm OK"
                frmVisualCalc.mnuWinASCII.Enabled = True
                Call UpdateStatus
                PgmName = CInt(T)                     'save user-defined program name (always local=0)
                IsDirty = False                       'indicate not dirty
                Call UpdateStatus
              End If
              Close #j                                'close file
            End If
            Preprocessd = False                       'not Preprocessed
            Compressd = False
            If AutoPprc Then
              Call Preprocess
            Else
              Call ResetListSupport
            End If
            InstrPtr = 0
          '-----------------------------
          Case iSave  ' Save Binary code
            If InstrCnt = 0 Then                    'if nothing to do
              Errstr = "No LEARNED code exists"
            Else
              On Error Resume Next
              If Fso.FileExists(S & ".pgm") Then Fso.DeleteFile S & ".pgm" 'delete old version
              Open S & ".pgm" For Random Access Write As #j Len = 2
              Call CheckError                       'check for errors
              On Error GoTo 0
              If Not ErrorFlag Then                 'if we are OK
                For Idx = 0 To InstrCnt - 1
                  Put #j, , Instructions(Idx)       'save all instructions
                Next Idx
                DisplayMsg "Saved Binary Pgm" & T & ".pgm OK"
                PgmName = CInt(T)
                Call UpdateStatus
                IsDirty = False                     'no longer dirty
              End If
            End If
            Close #j
            DisplayReg = PndImmed                   'reset display to prior value
          '-----------------------------
          Case iLapp  ' Lapp  Load and Append to existing code
            On Error Resume Next
            If Fso.FileExists(S & ".pgm") Then
              Open S & ".pgm" For Random Access Read As #j Len = 2
              Call CheckError
              On Error GoTo 0
              If Not ErrorFlag Then
                K = LOF(j) / 2                        'get # of instructions
                i = InstrCnt                          'save start location
                InstrCnt = i + K                      'set new instruction count
                If InstrCnt > InstrSize Then          'if we will need to bump our pool
                  Do While InstrCnt > InstrSize
                    InstrSize = InstrSize + InstrInc  'increment by 100
                  Loop
                  ReDim Preserve Instructions(InstrSize) 'resize pool
                End If
                For Idx = i To InstrCnt
                  Get #j, , Instructions(Idx)         'get all instructions
                Next Idx
                Call ResetBracing                     'reset any special bracing
                DisplayMsg "Load Appended Pgm" & T & ".pgm OK"
                Call UpdateStatus
                IsDirty = True                        'indicate pool is now dirty
                Preprocessd = False                   'not Preprocessed
                Compressd = False
                If AutoPprc Then
                  Call Preprocess
                Else
                  Call ResetListSupport
                End If
              End If
              Close #j
            Else
              Errstr = " Pgm" & T & ".pgm" & vbCrLf & _
                       " Program does not exist in path:" & vbCrLf & _
                       " " & SS & "\PGM"
            End If
            Preprocessd = False
            Compressd = False
            Call Preprocess
            InstrPtr = 0
          '-----------------------------
          Case iASCII  ' ASCII 'save to ASCII file
            On Error Resume Next
            If Fso.FileExists(S & ".txt") Then Fso.DeleteFile (S & ".txt")  'delete old
            Set ts = Fso.OpenTextFile(S & ".txt", ForWriting, True)
            Call CheckError                   'check for errors
            On Error GoTo 0
            If Not ErrorFlag Then             'if OK
              Ary = BuildInstrArray()         'get array list
              ts.Write Join(Ary, vbCrLf)      'write data
              DisplayMsg "Saved ASCII Pgm" & T & ".txt OK"
            End If
            ts.Close                          'close file
          '-----------------------------
        End Select
        DisplayText = False

' CE ---------------------------------------------------------------------------
    Case 133  ' CE
      Call CE_RunSupport

' CLR --------------------------------------------------------------------------
    Case 134  ' CLR
    Call CLR_RunSupport

' Op ---------------------------------------------------------------------------
    Case 135  ' OP
      If GetInstruction(1) = iIND Then            'if Op IND xx...
        Set Vptr = CheckVbl()                     'get storage object
        If Vptr Is Nothing Then Exit Sub
        TstData = Val(ExtractValue(Vptr))         'get value there to temp variable
        If TstData < 0# Or TstData > CDbl(MaxOps) Then
          Errstr = "Invalid Operation number"
          Exit Sub
        End If
      Else
        Call CheckForNumber(Iptr, 2, MaxOps)      'get value to TstData
      End If
      Call ResetPndAll                            'OP must do this here
      Call ProcessOP(CInt(TstData))               'now invoke operation

'-------------------------------------------------------------------------------
    Case 136  ' SST   'ignore
    Case 137  ' INS   'ignore
    Case 138  ' Cut   'ignore
    Case 139  ' Copy  'ignore

' PtoR -------------------------------------------------------------------------
    Case 140  ' PtoR
      DisplayReg = AngToRad(DisplayReg)       'get angle in Radians
      TV = TestReg * Cos(DisplayReg)          'get x
      DisplayReg = TestReg * Sin(DisplayReg)  'get y to display register
      TestReg = TV                            'set x to test register
      DisplayText = False

' STO --------------------------------------------------------------------------
    Case 141  ' STO
      ErrorFlag = False
      Set Vptr = CheckStruct()
      If ErrorFlag Then Exit Sub
      If Vptr Is Nothing Then Set Vptr = CheckIndVar()
      Call STO_Run(Vptr)
      
' RCL --------------------------------------------------------------------------
    Case 142  ' RCL
      ErrorFlag = False
      Set Vptr = CheckStruct()
      If ErrorFlag Then Exit Sub
      If Vptr Is Nothing Then Set Vptr = CheckIndVar()
      Call RCL_run(Vptr)

' EXC --------------------------------------------------------------------------
    Case 143  ' EXC
      ErrorFlag = False
      Set Vptr = CheckStruct()
      If ErrorFlag Then Exit Sub
      If Vptr Is Nothing Then Set Vptr = CheckIndVar()
      Call EXC_Run(Vptr)

' SUM --------------------------------------------------------------------------
    Case 144  ' SUM
      ErrorFlag = False
      Set Vptr = CheckStruct()
      If ErrorFlag Then Exit Sub
      If Vptr Is Nothing Then Set Vptr = CheckIndVar()
      Call SUM_Run(Vptr)

' MUL --------------------------------------------------------------------------
    Case 145  ' MUL
      ErrorFlag = False
      Set Vptr = CheckStruct()
      If ErrorFlag Then Exit Sub
      If Vptr Is Nothing Then Set Vptr = CheckIndVar()
      Call MUL_Run(Vptr)

'-------------------------------------------------------------------------------
      Case 146  ' IND  'Ignore

' Reset ------------------------------------------------------------------------
    Case 147  ' Reset
      InstrErr = 0        'reset instruction error line (if one set).
      Call Reset_Support  'reset instruction pointer and other flags
      InstrPtr = -1       'Use -1 for later INC

' Hkey -------------------------------------------------------------------------
    Case 148  ' Hkey
      Call CheckForText(Iptr - 1, DisplayWidth)
      DspTxt = TxtData                    'stuff data to DspTxt for processing
      Call HkeySkeySupport
      For i = 1 To 26
        Hidden(i) = Kyz(i)                'mark visibility flags
      Next i
      Call RedoAlphaPad
      DisplayText = True

' lnX --------------------------------------------------------------------------
    Case 149  ' lnX
      DisplayReg = Log(DisplayReg)
      DisplayText = False
      
' E+ ---------------------------------------------------------------------------
    Case 150  ' E+
      Call StatSUMadd

' Mean -------------------------------------------------------------------------
    Case 151  ' Mean
      Call StatMean

' X! ---------------------------------------------------------------------------
    Case 152  ' X!
      Call Factorial(DisplayReg)
      DisplayText = False

' X><T -------------------------------------------------------------------------
    Case 153  ' X><T
      TV = DisplayReg         'swap display register with the test register
      DisplayReg = TestReg
      TestReg = TV
      DisplayText = False

' HYP --------------------------------------------------------------------------
    Case iHyp  ' Hyp
      Call IncInstrPtr      'skip Past Hyp
      i = GetInstruction(0) 'Get next instruction
      If i = iArc Then      'Hyp Arc?
        Call IncInstrPtr    'yes, so skip Arc
        i = GetInstruction(0) 'Get next instruction
        If i < 158 Then     'Sin, Cos, Tan
          i = i + 251       'I - 155 + 400 + 6  'Hyp Arc versions
        Else                'Sec, Csc, Cot
          i = i + 132       'I - 283 + 409 + 6
        End If
      Else                  'Hyp only
        If i < 158 Then     'Sin, Cos, Tan
          i = i + 248       'I - 155 + 400 + 3  'Hyp versions
        Else                'Sec, Csc, Cot
          i = i + 129       'I - 283 + 409 + 3
        End If
      End If
      Call RunCmd(i, Errstr) 'process command

' Arc --------------------------------------------------------------------------
    Case iArc  ' Arc
      Call IncInstrPtr      'skip Past Arc
      i = GetInstruction(0) 'Get next instruction
      If i < 158 Then       'Sin, Cos, Tan
        i = i + 245         'I - 155 + 400  'Arc versions
      Else                  'Sec, Csc, Cot
        i = i + 126         'I - 283 + 409
      End If
      Call RunCmd(i, Errstr) 'process command

' Sin --------------------------------------------------------------------------
    Case 155  ' Sin
      DisplayReg = Sin(AngToRad(DisplayReg))
      DisplayText = False

' Cos --------------------------------------------------------------------------
    Case 156  ' Cos
      DisplayReg = Cos(AngToRad(DisplayReg))
      DisplayText = False

' Tan --------------------------------------------------------------------------
    Case 157  ' Tan
      DisplayReg = Tan(AngToRad(DisplayReg))
      DisplayText = False

' Sec --------------------------------------------------------------------------
    Case 283  ' Sec
      On Error Resume Next
      DisplayReg = 1# / Cos(AngToRad(DisplayReg))
      Call CheckError
      On Error GoTo 0
      
' Csc --------------------------------------------------------------------------
    Case 284  ' Csc
      On Error Resume Next
      DisplayReg = 1# / Sin(AngToRad(DisplayReg))
      Call CheckError
      On Error GoTo 0

' Cot --------------------------------------------------------------------------
    Case 285  ' Cot
      On Error Resume Next
      DisplayReg = 1# / Tan(AngToRad(DisplayReg))
      Call CheckError
      On Error GoTo 0

' 1/X --------------------------------------------------------------------------
    Case 158  ' 1/X
      On Error Resume Next
      DisplayReg = 1# / DisplayReg
      Call CheckError
      On Error GoTo 0

'-------------------------------------------------------------------------------
    Case 159  ' Txt  'Ignore

' Hex --------------------------------------------------------------------------
    Case 160  ' Hex
      If BaseType <> TypHex Then              'ignore if alread type
        If BaseType = TypDec Then DisplayReg = Fix(DisplayReg)  'remove possible decimal
        BaseType = TypHex                     'set new type
        Call UpdateStatus
      End If
    
' & ----------------------------------------------------------------------------
    Case 161  ' &
      Call Pend(iAndB)
      
' StFlg ------------------------------------------------------------------------
    Case 162  ' StFlg
      i = GetInstruction(1)                     'get next code
      Select Case i
        Case 0 To 9                             '0-9
          flags(i) = True                       'set flag
          IncInstrPtr                           'bump instruction pointer
        Case iIND                               'indirection?
          Set Vptr = CheckVbl()                 'point to desired storage object
          Call StFlg_IND(Vptr)
      End Select
    
' IfFlg ------------------------------------------------------------------------
    Case 163  ' IfFlg
      i = GetInstruction(1)                     'get next code (in case 0-9)
      Select Case i
        Case 0 To 9                             '0-9
          IncInstrPtr                           'point to '{'-1
          Call BuildBraceStk(InstrPtr, -1, -1, flags(i))
        Case iIND                               'indirection?
          Set Vptr = CheckVbl()                 'point to desired storage object
          Call IfFlg_IND(Vptr)
      End Select
    
' X==T -------------------------------------------------------------------------
    Case 164  ' X==T
      Bol = DisplayReg = TestReg            'set result
      Call BuildBraceStk(InstrPtr + 1, -1, -1, Bol)
    
' X>=T -------------------------------------------------------------------------
    Case 165  ' X>=T
      Bol = DisplayReg >= TestReg           'set result
      Call BuildBraceStk(InstrPtr + 1, -1, -1, Bol)
    
' X>T --------------------------------------------------------------------------
    Case 166  ' X>T
      Bol = DisplayReg > TestReg            'set result
      Call BuildBraceStk(InstrPtr + 1, -1, -1, Bol)
    
'-------------------------------------------------------------------------------
    Case 167     ' Dfn  'ignore
    Case iColon  ' :    'ignore
    
' ( ----------------------------------------------------------------------------
    Case 169  ' (
      Call Pend(iLparen)
    
' ) ----------------------------------------------------------------------------
'    Case 170  ' )        'never encountered

'-------------------------------------------------------------------------------
    Case iEparen          'new substitution (set by Preprocessr)
      Call Pend(iRparen)  'use old ) for calcs
    
'-------------------------------------------------------------------------------
'    Case iDWparen ')        'special end parentheses for Do{...}While(..)
'    Case iSparen  ')        'special end parentheses for Select definition
'    Case iUparen  ')        'special end parentheses for Until definition
'    Case iCparen  ')        'special end parentheses for Case definition
'    Case iIparen  ')        'special end parentheses for If definition
'    Case iWparen  ')        'special end parentheses for While definition
'    Case iFparen  ')        'special end parentheses for For definition
    Case iDWparen, iSparen, iUparen, iCparen, iIparen, iWparen, iFparen
      Call Pend(iRparen)  'use old ) for calcs
      Call ProcessEparen(Code)  'now process what we should do next

' / ----------------------------------------------------------------------------
    Case 171  ' /
      Call Pend(iDVD)
    
'-------------------------------------------------------------------------------
    Case iStyle  ' Style  'ignore

' Dec --------------------------------------------------------------------------
    Case 173  ' Dec
      If BaseType <> TypDec Then              'ignore if already type
        BaseType = TypDec                     'set type
        Call UpdateStatus
      End If
    
' | ----------------------------------------------------------------------------
    Case 174  ' |
      Call Pend(iOrB)
    
' Int --------------------------------------------------------------------------
    Case 175  ' Int
      DisplayReg = Fix(DisplayReg)
    
' Abs --------------------------------------------------------------------------
    Case 176  ' Abs
      DisplayReg = Abs(DisplayReg)
    
' Fix --------------------------------------------------------------------------
    Case 177  ' Fix
      Do
        i = GetInstruction(1)                   'get next code
        Select Case i
          Case EEKey                            'Fix Engineering mode?
            EngMode = True                      'enable engineering mode
            EEMode = True                       'made possible through EE flag
          Case 0 To 9                           '0-9
            IncInstrPtr
          Case iIND                             'indirection?
            Set Vptr = CheckVbl()               'point to desired storage object
            If Vptr Is Nothing Then Exit Do     'error
            TV = Fix(Val(ExtractValue(Vptr)))   'get value there
            If TV < 0# Or TV > 9# Then          'in range?
              Errstr = "Parameter is out of range (0-9)"
              Exit Do
            End If
            i = CInt(TV)                        'get value
        End Select
        DspFmtFix = i                           'save decimal count
        DspFmt = "0." & String$(DspFmtFix, "0") 'set format
        ScientifEE = DspFmt & "E+00"
        Exit Do
      Loop
      
' D.MS -------------------------------------------------------------------------
    Case 178  ' D.MS
      vDeg = Fix(DisplayReg)                        'get DDD
      DisplayReg = (DisplayReg - vDeg) * 100#       'get MM.SSsss
      vMin = Fix(DisplayReg)                        'get MM
      vSec = (DisplayReg - vMin) * 100#             'get SS.dddd
      DisplayReg = vDeg + vMin / 60# + vSec / 3600# 'get dd.ddddd
          
' EE ---------------------------------------------------------------------------
    Case 179  ' EE
      EEMode = True
      
' Sbr --------------------------------------------------------------------------
    Case 180  ' Sbr 'simply pass over definition (already defined via preprocessor)
      Call CheckForLabel(InstrPtr, LabelWidth)          'skip label
      Call IncInstrPtr                                  'point to '{'
      Call BuildBraceStk(InstrPtr, -1, -1, True)        'build subroutine block
    
' x ----------------------------------------------------------------------------
    Case 184  ' x
      Call Pend(iMult)
    
' Rem --------------------------------------------------------------------------
    Case iRem, iRem2 ' Rem and [']
      Do
        Select Case GetInstruction(1)
          Case 0 To 9, Is > 127             'regular ops?
            Exit Do                         'yes, so done (next IncInstrPtr will point to it)
          Case Else
            IncInstrPtr                     'else bump pointer to current checked instruction
        End Select
      Loop
      
' Oct --------------------------------------------------------------------------
    Case 186  ' Oct
      If BaseType <> TypOct Then                                'ignore if already type
        If BaseType = TypDec Then DisplayReg = Fix(DisplayReg)  'remove decimal
        BaseType = TypOct                                       'set type
        Call UpdateStatus                                       'display it in status
      End If
    
' ~ ----------------------------------------------------------------------------
    Case 187  ' ~
      Call Pend(iNotB)
    
' Select -----------------------------------------------------------------------
    Case 188  ' Select
      i = FindEPar2(1)                                  'find ending [)] and point to [{]
      Call BuildBraceStk(i, -1, -1)                     'build select block
      'LpSelect value will be set when Select's [)] is encountered
      
' SelectT ----------------------------------------------------------------------
    Case 432  ' SelectT 'Select with T-register value
      IncInstrPtr                                       'point to [{]
      Call BuildBraceStk(InstrPtr, -1, -1)              'build select block
      BracePool(BraceIdx).LpSelect = TestReg            'set test value
      
' Case -------------------------------------------------------------------------
    Case 189  ' Case
      Select Case GetInstruction(1)                     'check next instruction
        Case iLparen                                    ' '('
          TV = BracePool(BraceIdx).LpSelect             'save Select's master value
          i = FindEPar2(1)                              'find ending [)] and point to [{]
          Call BuildBraceStk(i, -1, -1)                 'build Case block
          BracePool(BraceIdx).LpSelect = TV             'set test value for case
        Case iCaseElse                                  'Case Else?
          IncInstrPtr                                   'yes, point to [{] (no need to create block)
      End Select
    
' { ----------------------------------------------------------------------------
    Case 190  ' { 'encountered only during IF processing (others are pre-handled)
      With BracePool(BraceIdx)
        If Not .LpTrue Then         'if we should NOT process the block...
          InstrPtr = .LpTerm - 1    'set control to the end of the block - 1 ('}'-1)
        End If                      'otherwise the pointer is right where it should be
      End With
    
' } ----------------------------------------------------------------------------
    Case iBCBrace, iSCBrace, iICBrace, iDCBrace, iWCBrace, _
         iFCBrace, iCCBrace, iDWBrace, iDUBrace, iEIBrace, _
         iENBrace, iSTBrace, iSIBrace, iCNBrace, iRCbrace
      Call ProcessEbrace(Code, Errstr)
    
' Deg --------------------------------------------------------------------------
    Case 192  ' Deg
      AngleType = TypDeg
      Call UpdateStatus
    
' Lbl --------------------------------------------------------------------------
    Case 193  ' Lbl 'simply pass over definition (already defined via preprocessor)
      Call CheckForLabel(InstrPtr, LabelWidth)        'skip label
      Call IncInstrPtr                                'bump to ':'
    
' - ----------------------------------------------------------------------------
    Case 197  ' -
      Call Pend(iMinus)
    
' Beep -------------------------------------------------------------------------
    Case 198  ' Beep
      Beep
      
' Bin --------------------------------------------------------------------------
    Case 199  ' Bin
      If BaseType <> TypBin Then                                'ignore if already type
        If BaseType = TypDec Then DisplayReg = Fix(DisplayReg)  'remove decimal
        BaseType = TypBin                                       'set type
        Call UpdateStatus                                       'display it in status
      End If
    
' ^ ----------------------------------------------------------------------------
    Case 200  ' ^
      Call Pend(iXorB)
    
' For --------------------------------------------------------------------------
    Case 201  ' For
      Call FindForInfo(i, j, K)                         'find 3 parts of the FOR data
      Iptr = FindEPar2(1)                               'find ending [)]+1 (points to [{])
      Call BuildBraceStk(Iptr, j, K)                    'build While()[...} block
      With BracePool(BraceIdx)
        .LpDspReg = DisplayReg                          'save Display Register
        .LpLoop = True                                  'this is a looping block
        ForInit = False                                 'ensure For Init is turned off
        If i = -1 Then                                  'if not init code...
          If .LpCond = -1 Then                          'if no condition code...
            InstrPtr = Iptr                             'point to '{' after ')'
          Else
            InstrPtr = .LpCond                          'point to condition
            Call Pend(iLparen)                          'begin conditional testing
          End If
        Else
          ForInit = True                                'else we are pointing to init data
          Call Pend(iLparen)                            'begin Init processing
        End If
      End With
      
' Do ---------------------------------------------------------------------------
    Case 202  ' Do
      IncInstrPtr                                       'point to [{]
      Call BuildBraceStk(InstrPtr, -1, -1)              'build block for basic DO
      With BracePool(BraceIdx)
        .LpLoop = True                                  'this is a looping block
        i = .LpTerm + 1                                 'point to possible Until or While
        Select Case GetInstructionAt(i)                 'check the instruction at '}' + 1
          Case iWhile, iUntil                           'Do-While or Do-Until?
            .LpCond = i                                 'save address for condition at '('-1
            Do
              i = i + 1                                 'bump pointer
              Select Case GetInstructionAt(i)           'check instruction at that location
                Case iDWparen, iUparen                  'end paren for do-while or do-Until?
                  Exit Do                               'yes, so exit
              End Select
            Loop                                        'else keep searching
            .LpTerm = i                                 'set new terminator for block
        End Select
      End With
      
' While ------------------------------------------------------------------------
    Case 203  ' While 'will not be encountered in DO..While loops (already processed in block)
      j = InstrPtr                                      'point to [(]-1
      i = FindEPar2(1)                                  'find ending [)]+1 (points to [{])
      Call BuildBraceStk(i, -1, j)                      'build While()[...} block
      BracePool(BraceIdx).LpLoop = True                 'this is a looping block
    
' Pmt --------------------------------------------------------------------------
    Case 204  ' Pmt
      Call CheckForText(InstrPtr, DisplayWidth)
      DspTxt = TxtData      'grab data to prompt user with
      DisplayText = True
      Call ForceDisplay     'display the data
      DisplayText = False
      Call NewLine          'advance to new line for user response
      Call ResetPndAll      'reset pending commands
      HMRunMode = MRunMode  'save status of MRunMode
      PmtFlag = True        'enable prompt flag
      
' Rad --------------------------------------------------------------------------
    Case 205  ' Rad
      AngleType = TypRad
      Call UpdateStatus
    
' Ukey -------------------------------------------------------------------------
    Case 206  ' UKey
      Call CheckForLabel(InstrPtr, 1)                       'get Key (A-Z)
      JL = Asc(UCase$(TxtData)) - 64                        'get base index
      If CBool(ActivePgm) Then
        JL = JL + ModLblMap(ActivePgm - 1)                  'index to module offset
        Pool = ModLbls                                      'use the module list
      Else
        Pool = Lbls                                         'the pool uses the user pgm label list
      End If
      frmVisualCalc.cmdUsrA(i).Caption = Trim$(Pool(JL).lblName)  'set the name for the key
      InstrPtr = Pool(JL).LblDat                            'now skip past the rest
      Call BuildBraceStk(InstrPtr, -1, -1, True)            'build subroutine block
    
' + ----------------------------------------------------------------------------
    Case 210  ' +
      Call Pend(iAdd)
    
' Plot -------------------------------------------------------------------------
    Case 211  ' Plot
      IncInstrPtr                               'pre-skip to next instruction
      Select Case GetInstruction(0)             'now check it
        Case iCLR                               'CLR
          Call PlotClr                          'clear Plot screen
          Exit Sub
        Case iClose                             'hide plot screen
          frmVisualCalc.PicPlot.Visible = False
          PlotTrigger = False
          Exit Sub
        Case iOpen                              'show plot screen
          frmVisualCalc.PicPlot.Visible = True
          PlotTrigger = False
          Exit Sub
        Case iSbr                               'Plot Sbr?
          If GetInstruction(1) = iReset Then    'Plot Sbr Reset?
            IncInstrPtr
            PlotTrigger = False                 'yes
            Exit Sub
          Else
            Call CheckForLabel(InstrPtr, LabelWidth)
            '
            ' get index for label
            '
            Idx = FindLbl(TxtData, TypSbr)      'allow subroutines...
            If Idx = 0 Then
              Idx = FindLbl(TxtData, TypKey)    'or user-defined keys
            End If
            If Idx = 0 Then                     'if nothing found
              Errstr = "Invalid parameter"
              Exit Sub
            End If
            '
            ' now grab sub definition address for index
            '
            If CBool(ActivePgm) Then
              Idx = ModLbls(Idx).lblAddr
            Else
              Idx = Lbls(Idx).lblAddr
            End If
            '
            ' store plot trigger subroutine address
            '
            PlotTrigger = True                      'enable plot tigger
            PlotTriggerSbr = CInt(Idx)              'set plot trigger subtoutine
            Exit Sub
          End If
      End Select
'
' Plot point (or to point)
'
      Bol = (GetInstruction(0) = iMinus)            'are we drawing TO a point?
      If Bol Then IncInstrPtr                       'skip instruction if so
'
' get X and Y values for Plot
'
      If Not GetPlotXY(Xs, Ys, Errstr) Then Exit Sub
'
' if Paint fill is specified
'
      With frmVisualCalc
        If GetInstruction(1) = iComma Then
          IncInstrPtr
          IncInstrPtr                               'point to 0 or 1
          i = GetInstruction(0)                     'get plot flag (0 or 1)
          JL = -1                                   'init border color
          If GetInstruction(1) = iComma Then        'user wants to specify a border color?
            IncInstrPtr                             'point to comma
            If GetInstruction(1) = iIND Then        'check for Variable indirection
              Set Vptr = CheckVbl()                 'get variable pointer
              If Vptr Is Nothing Then Exit Sub      'error
              TstData = Val(ExtractValue(Vptr))     'else grab its value
            Else
              Call CheckForValue(InstrPtr)          'else find number
            End If
            On Error Resume Next
            JL = CLng(TstData)                      'get long version of color value
            On Error GoTo 0
          Else
          End If
          
          If i = 1 Then                             'paint fill?
            Call PaintIt(Xs, Ys, PlotColor, JL)     'fill it if so
            Exit Sub
          End If
        End If
        
        If Bol Then                                 'draw line to new point?
          Call DrawLine(PlotX, PlotY, Xs, Ys, PlotColor)
        Else                                        'draw plot
          Call SetPoint(Xs, Ys, PlotColor)          'draw dot
        End If
      End With
    
' Nvar, Tvar, Ivar, Cvar -------------------------------------------------------
    Case iNvar, iTvar, iIvar, iCvar
      Select Case Code
        Case iNvar
          Typ = vNumber
        Case iTvar
          Typ = vString
        Case iIvar
          Typ = vInteger
        Case iCvar
          Typ = vChar
      End Select
      i = InstrPtr                                      'save definition location
      Nm = vbNullString                                 'init no variable name
      Ln = 0                                            'init no data length
      Call CheckForNumber(InstrPtr, 2, 99)              'check variable number
      Vn = CLng(TstData)                                'grab variable number
      If GetInstruction(1) = iLbl Then                  'found Lbl?
        Call CheckForLabel(InstrPtr + 1, LabelWidth)    'yes, so grab label
        Nm = TxtData                                    'and stor it
      End If
      If GetInstruction(1) = iLen Then                  'data length defined?
        Call CheckForNumber(InstrPtr, 2, DisplayWidth)  'get length
        Ln = CInt(TxtData)                              'store it
      End If
      
      With Variables(Vn)                                'now apply changes to variable
        .VName = Nm                                     'apply name if defined
        .VdataLen = Ln                                  'set len, if non-zero
        Set .Vdata = Nothing                            'clear any defined classes (and children)
        Set .Vdata = New clsVarSto                      'init brand new variable storage
        .Vdata.VarRoot = Vn                             'set root variable
        .VuDef = True                                   'mark as user-defined variables
        .Vaddr = i                                      'set definition address
        .VarType = Typ                                  'set variable type
      End With
        
      If CheckRunDim2(iX, iY) Then                      'if dimension dims found...
        Call BuildMDAry(Vn, iX, iY, False)             'process them
      End If
    
' % ----------------------------------------------------------------------------
    Case 213  ' %
      Call Pend(iMod)
    
' If ---------------------------------------------------------------------------
    Case 214  ' If
      i = FindEPar2(1)                                  'find ending [)] and point to [{]
      Call BuildBraceStk(i, -1, -1)                     'build IF block
    
' Else -------------------------------------------------------------------------
    Case 215  ' Else
      With BracePool(BraceIdx)
        .LpTrue = Not .LpTrue                           'simply flip current block's True flag
        .LpTerm = FindEblock(InstrPtr + 1)              'get term of Else block
      End With
    
' ElseIf -------------------------------------------------------------
    Case 343  ' ElseIf
      Bol = BracePool(BraceIdx).LpTrue                  'get current condition
      BraceIdx = BraceIdx - 1                           'quickly remove old block (re-use saves no time)
      i = FindEPar2(1)                                  'find ending [)] and point to [{]
      Call BuildBraceStk(i, -1, InstrPtr)               'build brand new IF block
      If Bol Then                                       'if last block WAS processed (True)
        InstrPtr = BracePool(BraceIdx).LpTerm - 1       'then point to end of the block (ignore new block)
      End If
      
' Cont -------------------------------------------------------------------------
    Case 216  ' Cont
      Do While Not BracePool(BraceIdx).LpLoop 'find a looping brace block
        BraceIdx = BraceIdx - 1
      Loop
      With BracePool(BraceIdx)
        If .LpProcess <> -1 Then    'if a process defined, go to it, first
          InstrPtr = .LpProcess
          Call Pend(iLparen)        'prepare for conditional
        ElseIf .LpCond <> -1 Then   'else if a condition defined, go to it, secondly
          InstrPtr = .LpCond
          Call Pend(iLparen)        'prepare for conditional
        Else
          InstrPtr = .LpStart       'else go to the top of the block
        End If
      End With
    
' Break ------------------------------------------------------------------------
    Case 217  ' Break
      Do While Not BracePool(BraceIdx).LpLoop 'find a looping brace block
        BraceIdx = BraceIdx - 1
      Loop
      With BracePool(BraceIdx)
        InstrPtr = .LpTerm          'go to end of data
        BraceIdx = BraceIdx - 1
      End With
    
' Grad -------------------------------------------------------------------------
    Case 218  ' Grad
      AngleType = TypGrad
      Call UpdateStatus
    
' R/S --------------------------------------------------------------------------
    Case 219  ' R/S
      RunMode = False           'disable Run mode
      ModPrep = 0
      Call ResetPndAll          'reset data
    
' +/- --------------------------------------------------------------------------
    Case 222  ' +/-
      DisplayReg = -DisplayReg
      DisplayText = False
      
' = ----------------------------------------------------------------------------
    Case 223  ' =
      Call Pend(iEqual)
    
' Print, Print;-----------------------------------------------------------------
    Case iPrint, iPrintx  ' Print, Print;
      Select Case GetInstruction(1)
        '---------
        Case iReset             'Allow 'Print Reset'
          Call PrintReset
          Exit Sub
        '---------
        Case iAdv               'Allow 'Print Adv'
          PlotXDef = PlotXDef + CLng(CDbl(LineHeight) * Sin(LastDir))
          PlotYDef = PlotYDef + CLng(CDbl(LineHeight) * Cos(LastDir))
          PlotX = PlotXDef
          PlotY = PlotYDef
          Exit Sub
        '---------
        Case iLparen            'Allow 'Print(x,y[,dir])
          IncInstrPtr
          If Not GetPlotValue("X", PlotWidth, PlotX, Errstr) Then Exit Sub  'get X
          IncInstrPtr                       'skip ',' known to be here
          If Not GetPlotValue("Y", PlotHeight, PlotY, Errstr) Then Exit Sub  'get Y
          If GetInstruction(1) = iComma Then                                'Dir is optional
            IncInstrPtr                                                     'skip ','
            If Not GetPlotValue("Dir", 7&, Clr, Errstr) Then Exit Sub       'get Dir
            LastDir = CDbl(Clr) * Atn(1)                                    'compute new def. angle
          End If
          IncInstrPtr                       'skip ')' known to be here
      End Select
'
' now check for text data to print. This can be a normal text, a variable (via IND),
' a Constant, or immediate data (DspTxt)
'
      If GetInstruction(1) = iIND Then              'indirection for output text?
        Set Vptr = CheckVbl()                       'yes, so get variable
        If Vptr Is Nothing Then Exit Sub
        TxtData = CStr(ExtractValue(Vptr))          'else grab text
      Else
        If GetInstruction(1) < 128 Then             'next instruction is text?
          JL = 0                                    'init flag
          i = InstrPtr                              'yes, so save current address
          If CheckForLabel(i, LabelWidth) Then      'first see if the text qualifies as a label
            JL = FindLbl(TxtData, TypConst)         'if it does, see if it is a constant
          End If
          If CBool(JL) Then                         'did we find a constant?
            If CBool(ActivePgm) Then                'yes, so extract the constant text as the data
              TxtData = RTrim$(ModLbls(JL).lblCmt)  'either from the module...
            Else
              TxtData = RTrim$(Lbls(JL).lblCmt)     'or the user area
            End If
          Else
            Call CheckForText(i, LabelWidth)        'extract to TxtData normally if not a constant
          End If
        Else
          TxtData = DspTxt                          'else use immediate 'display' text
        End If
      End If
'
' we now have everything. PlotX and PlotY are at the current X and Y coordinates.
' LastDir is det to any default angle, and TxtData contains the text to display
'
      With frmVisualCalc
        With .PicPlot
          .CurrentX = PlotX         'set start location
          .CurrentY = PlotY
          .ForeColor = PlotColor    'set color to draw text with
        End With
        Call PrintTextAtAngle(.PicPlot, TxtData, CSng(LastDir) * 57.29578!) 'plot text in degrees
        With .PicPlot
          Select Case Code
            Case iPrint 'Print
              PlotXDef = PlotXDef + CLng(CDbl(LineHeight) * Cos(LastDir))
              PlotYDef = CLng(.CurrentY + CSng(CDbl(LineHeight) * Sin(LastDir)))
              PlotX = PlotXDef
              PlotY = PlotYDef
            Case Else   'Print;
              PlotX = .CurrentX
              PlotY = .CurrentY
          End Select
        End With
      End With
    
' >> ---------------------------------------------------------------------------
    Case 226  ' >>
      Do
        i = GetInstruction(1)                   'get next code
        IncInstrPtr
        If i < 1 Or i > 9 Then                  'in range?
          Errstr = "Parameter is out of range (1-9)"
          Exit Do
        End If
        For Idx = 1 To i
          If DisplayReg = 0# Then Exit For  'if null, then nothing to do
          DisplayReg = DisplayReg / 2#
        Next Idx
        Exit Do
      Loop
      DisplayText = False
    
' y^ ---------------------------------------------------------------------------
    Case 227  ' y^
      Call Pend(iPower)
    
' X² ---------------------------------------------------------------------------
    Case 228  ' X²
      On Error Resume Next
      TV = DisplayReg * DisplayReg
      Call CheckError
      If Not ErrorFlag Then DisplayReg = TV
    
' Pi ---------------------------------------------------------------------------
    Case 229  ' Pi
      DisplayReg = vPi
      DisplayText = False
      
' Rnd --------------------------------------------------------------------------
    Case 230  ' Rnd
      DisplayReg = Rnd()
      DisplayText = False
    
' Mil --------------------------------------------------------------------------
    Case 231  ' Mil
      AngleType = TypMil
      Call UpdateStatus
    
' Pvt --------------------------------------------------------------------------
    Case 232  ' Pvt  'ignore
    
' Const ------------------------------------------------------------------------
    Case 233  ' Const
      Call CheckForLabel(InstrPtr, LabelWidth)  'skip over constant label
      IncInstrPtr                               'point to '{'
      InstrPtr = FindEblock(InstrPtr)           'set instruction pointer to end of block
    
' Struct -----------------------------------------------------------------------
    Case 234  ' Struct
      Call CheckForLabel(InstrPtr, LabelWidth)  'skip over structure label
      IncInstrPtr                               'point to '{'
      InstrPtr = FindEblock(InstrPtr)           'set instruction pointer to end of block
    
'-------------------------------------------------------------------------------
    Case 235  ' NxLbl  'ignore
    Case 236  ' PvLbl  'ignore
    
' Line -------------------------------------------------------------------------
    Case 237  ' Line
      IncInstrPtr                         'point to data
      Select Case GetInstruction(0)
        Case iLparen                      'define Line(x,y)-(x,y)[,z]
          If Not GetPlotXY(Xs, Ys, Errstr) Then Exit Sub
          If GetInstruction(1) = iMinus Then
            IncInstrPtr                   'point to '-'
            IncInstrPtr                   'point to '('
          Else
            Errstr = "Expected '-'"
            Exit Sub
          End If
        Case iMinus                       'define Line-(x,y)[,z]
          IncInstrPtr
          Xs = PlotX                      'default start
          Ys = PlotY
      End Select
      If Not GetPlotXY(Xe, Ye, Errstr) Then Exit Sub
      If GetInstruction(1) = iComma Then  '[,z]
        IncInstrPtr
        IncInstrPtr
        i = GetInstruction(0)             'get 0,1,2
      Else
        i = 0
      End If
      
      If Xs > Xe Then                     'make sure Xs < Xe
        Idx = Xs
        Xs = Xe
        Xe = Idx
      End If
      If Ys > Ye Then                     'make sure Ys < Ye
        Idx = Ys
        Ys = Ye
        Ye = Idx
      End If
      
      Select Case i
        Case 1                              'draw box
          Call DrawLine(Xs, Ys, Xe, Ys, PlotColor)
          Call DrawLine(Xe, Ys, Xe, Ye, PlotColor)
          Call DrawLine(Xe, Ye, Xs, Ye, PlotColor)
          Call DrawLine(Xs, Ye, Xs, Ys, PlotColor)
        Case 2                              'draw filled box
          Call DrawLine(Xs, Ys, Xe, Ys, PlotColor)
          Call DrawLine(Xe, Ys, Xe, Ye, PlotColor)
          Call DrawLine(Xe, Ye, Xs, Ye, PlotColor)
          Call DrawLine(Xs, Ye, Xs, Ys, PlotColor)
          Call PaintIt(Xs + 1, Ys + 1, PlotColor)
        Case Else                           'draw line
          Call DrawLine(Xs, Ys, Xe, Ye, PlotColor)
      End Select
    
' [ ----------------------------------------------------------------------------
    Case 238  ' [  'ignored
    
' ] ----------------------------------------------------------------------------
    Case 239  ' ]  'ignored
    
' ClrVar -----------------------------------------------------------------------
    Case 240  ' ClrVar
      If GetInstruction(1) = iAll Then
        IncInstrPtr
        Call ClearAllVariables                        'clear all variables
      Else
        Set Vptr = CheckVbl()                         'else only specified variable
        If Not Vptr Is Nothing Then
          Call ClearEmAll(Vptr)                       'clear it if it exists
        End If
      End If
      DisplayText = False
    
' SzOf -------------------------------------------------------------------------
    Case 241  ' SzOf
      Idx = -1                                        'init size value as failed
      Bol = GetInstruction(1) = i1X                   'see if next instruction is ABS
      If Bol Then IncInstrPtr                         'bump pointer if so
      Select Case GetInstruction(1)
        Case iIND                                     'indirection?
          Set Vptr = CheckIndVar()                    'get indirected pointer
          If Not Vptr Is Nothing Then
            Idx = SzOfElmnt(Vptr, Bol)                'grab size of target variable to Idx
          End If
        Case Is < 10                                  'numeric value (absolute variable number)
          Call CheckForNumber(Iptr - 1, 2, 99)        'grab variable number to TstData
          Idx = SzOfElmnt(Variables(CInt(TstData)).Vdata, Bol) 'grab size of variable to Idx
        Case Is < 128                                 'is label?
          Call CheckForLabel(Iptr - 1, LabelWidth)    'get label
          j = FindVblMatch(TxtData)                   'find a match in the variables list
          If CBool(j) Then                            'was a variable, so...
            InstrPtr = Iptr                           'reset instruction pointer
            Set Vptr = CheckVbl()                     'get variable
            Idx = SzOfElmnt(Vptr, Bol)                'grab size to Idx
          Else
            JL = FindLbl(TxtData, TypStruct)          'not variable. Structure?
            If CBool(JL) Then                         'yes
              If CBool(ActivePgm) Then
                Idx = ModStPl(ModLbls(JL).LblValue).StSiz 'get size of module's structure
              Else
                Idx = StructPl(Lbls(JL).LblValue).StSiz   'get size of structure
              End If
            Else
              Errstr = "Cannot find Variable or Structure referenced"
            End If
          End If
        Case iNvar
          Idx = 8
        Case iIvar
          Idx = 4
        Case iCvar
          Idx = 1
        Case iTvar
          Idx = DisplayWidth
      End Select
      If Idx <> -1 Then
        DisplayReg = CDbl(Idx) 'get size of specified item
      End If
      DisplayText = False

' Def --------------------------------------------------------------------------
    Case 242  ' Def
      Call CheckForLabel(InstrPtr, LabelWidth)        'get label to TxtData
      i = FindDefMatch(TxtData)                       'already defined?
      If i = 0 Then                                   'no
        i = FindDefMatch(vbNullString)                'a blank has been opened?
        If i = 0 Then                                 'no
          DefCnt = DefCnt + 1                         'so make one
          If DefCnt > DefSize Then
            DefSize = DefSize + DefInc                'bump pool size
            ReDim Preserve DefName(DefSize)           'bump pool
          End If
          i = DefCnt                                  'set location to stuff
        End If
        DefName(i) = TxtData                          'stuff name
      End If
    
' IfDef ------------------------------------------------------------------------
    Case 243  ' IfDef
      Call CheckForLabel(InstrPtr, LabelWidth)        'get label to TxtData
      i = FindDefMatch(TxtData)                       'already defined?
      DefDef = DefDef + 1                             'indicate we are processing Def blocks
      If DefDef > DefTrueSz Then
        DefTrueSz = DefTrueSz + 8
        ReDim Preserve DefTrue(DefTrueSz)
      End If
      DefTrue(DefDef) = CBool(i)                      'set True if defined
      
' Edef -------------------------------------------------------------------------
    Case 244  ' Edef
      If CBool(DefDef) Then
        DefDef = DefDef - 1                           'turn off Def processing
        If DefDef < 0 Then DefDef = 0
      Else
        Errstr = "Edef instruction, but no IfDef or !Def"
      End If
      
'----------------------
' EXTENDED FUNCTIONS: these are applied to a Compressed program
'----------------------
' STO IND ----------------------------------------------------------------------
    Case 245  ' STO IND
      ErrorFlag = False
      Set Vptr = CheckStruct()
      If ErrorFlag Then Exit Sub
      If Vptr Is Nothing Then Set Vptr = GrabInd()
      Call STO_Run(Vptr)
      
' RCL IND ----------------------------------------------------------------------
    Case 246  ' RCL IND
      ErrorFlag = False
      Set Vptr = CheckStruct()
      If ErrorFlag Then Exit Sub
      If Vptr Is Nothing Then Set Vptr = GrabInd()
      Call RCL_run(Vptr)
    
' EXC IND ----------------------------------------------------------------------
    Case 247  ' EXC IND
      ErrorFlag = False
      Set Vptr = CheckStruct()
      If ErrorFlag Then Exit Sub
      If Vptr Is Nothing Then Set Vptr = GrabInd()
      Call EXC_Run(Vptr)
    
' SUM IND ----------------------------------------------------------------------
    Case 248  ' SUM IND
      ErrorFlag = False
      Set Vptr = CheckStruct()
      If ErrorFlag Then Exit Sub
      If Vptr Is Nothing Then Set Vptr = GrabInd()
      Call SUM_Run(Vptr)
    
' MUL IND ----------------------------------------------------------------------
    Case 249  ' MUL IND
      ErrorFlag = False
      Set Vptr = CheckStruct()
      If ErrorFlag Then Exit Sub
      If Vptr Is Nothing Then Set Vptr = GrabInd()
      Call MUL_Run(Vptr)
    
' SUB IND ----------------------------------------------------------------------
    Case 250  ' SUB IND
      ErrorFlag = False
      Set Vptr = CheckStruct()
      If ErrorFlag Then Exit Sub
      If Vptr Is Nothing Then Set Vptr = GrabInd()
      Call SUB_Run(Vptr)
    
' DIV IND ----------------------------------------------------------------------
    Case 251  ' DIV IND
      ErrorFlag = False
      Set Vptr = CheckStruct()
      If ErrorFlag Then Exit Sub
      If Vptr Is Nothing Then Set Vptr = GrabInd()
      Call DIV_Run(Vptr)
    
' GTO IND ----------------------------------------------------------------------
      Case 252
      Call GTO_IND(CheckVbl())
      
' OP IND -----------------------------------------------------------------------
      Case 254: S = "OP IND"
        Set Vptr = CheckVbl()                     'get storage object
        If Vptr Is Nothing Then Exit Sub
        TstData = Val(ExtractValue(Vptr))         'get value there to temp variable
        If TstData < 0# Or TstData > CDbl(MaxOps) Then
          Errstr = "Invalid Operation number"
          Exit Sub
        End If
        Call ResetPndAll                          'OP must do this here
        Call ProcessOP(CInt(TstData))             'now invoke operation

' Fix IND ----------------------------------------------------------------------
      Case 255: S = "FIX IND"
        Set Vptr = CheckVbl()                   'point to desired storage object
        If Vptr Is Nothing Then Exit Sub        'error
        TV = Fix(Val(ExtractValue(Vptr)))       'get value there
        If TV < 0# Or TV > 9# Then              'in range?
          Errstr = "Parameter is out of range (0-9)"
          Exit Sub
        End If
        DspFmtFix = CInt(TV)                    'get value
        DspFmt = "0." & String$(DspFmtFix, "0") 'set format
        ScientifEE = DspFmt & "E+00"

' Pgm IND ----------------------------------------------------------------------
      Case 256
        Set Vptr = CheckVbl()                       'get variable object storage
        If Vptr Is Nothing Then Exit Sub            'error
        TstData = Val(ExtractValue(Vptr))           'get value there to temp variable
        If TstData = 0# Then                        'if selecting program 0 (User program)
          ActivePgm = 0                             'set it
          ModPrep = 0
          Exit Sub
        End If
      
        If TstData < 0# Or TstData > CDbl(ModCnt) Then  'check value range
          Errstr = "Invalid program number for module " & CStr(ModName)
          If ModName = 0 Then
            Errstr = Errstr & " (no module loaded)"
          Else
            Errstr = Errstr & " (1-" & CStr(ModCnt) & ")"
          End If
          Exit Sub
        End If
        
        ModPrep = CInt(TstData)   'get program number (activated when Call or Ukey cmd is invoked, next)

'---------------------------------------------------------------------
              '--2nd keys---------------------------------------------
'---------------------------------------------------------------------

' MDL ----------------------------------------------------------------
    Case 257  ' MDL 'ignored (for now)
      Select Case GetInstruction(1)
        Case iLoad  'Load
          IncInstrPtr
          Call CheckForNumber(InstrPtr, 4, 9999)
          DisplayReg = TstData
          Call Load_MDL
        Case iRCL
          DisplayReg = CDbl(ModCnt)
          DisplayText = False
          Call DisplayLine
        Case iLbl
          DisplayReg = CDbl(ModName)
          If CBool(Len(Trim$(ModLbl))) Then
            DisplayMsg Trim$(ModLbl)
          Else
            DisplayMsg "<LNoName>"
          End If
      End Select
' CMM ----------------------------------------------------------------
    Case 258  ' CMM
      Erase ModMem                    'erase module memory
      Erase ModMap
      ModSize = 0                     'make null size
      ModCnt = 0                      'no modules
      ModName = 0                     'mo module name
      ModLocked = False               'no locking
      ModLblCnt = 0                   'no labels
      ModStCnt = 0                    'no structures
      Erase ModLbls, ModLblMap, ModStPl, ModStMap, ModMem, ModMap
      frmVisualCalc.sbrImmediate.Panels("MDL").ToolTipText = "Currently loaded Module"
      If CBool(ActivePgm) Then
        ActivePgm = 0                 'reset active program to 00
        InstrErr = 0
        Call Reset_Support
        Call UpdateStatus
      End If

' CMs ----------------------------------------------------------------
    Case 261  ' CMs
      Call CMs_Support

' CP -----------------------------------------------------------------
    Case 262  ' CP
      TestReg = 0#    'CP only clears Test Register while in RUN mode
      DisplayText = False

'---------------------------------------------------------------------
    Case iList ' List   'ignored
    Case 264  ' BST    'ignored
    Case 265  ' DEL    'ignored
    Case 266  ' Paste  'ignored
'---------------------------------------------------------------------
    Case iUSR ' USR
      If GetInstruction(1) = iIND Then            'if USR IND xx...
        Set Vptr = CheckVbl()                     'get storage object
        If Vptr Is Nothing Then Exit Sub
        TstData = Val(ExtractValue(Vptr))         'get value there to temp variable
        If TstData < 0# Or TstData > CDbl(MaxUSR) Then
          Errstr = "Invalid Operation number"
          Exit Sub
        End If
      Else
        Call CheckForNumber(Iptr - 1, 2, MaxUSR)  'get value to TstData
      End If
      Call ResetPndAll                            'USR must do this here
      Call ProcessUSR(CInt(TstData))              'now invoke operation
    
' RtoP ---------------------------------------------------------------
    Case 268  ' RtoP
      TV = Sqr(DisplayReg * DisplayReg + TestReg * TestReg) 'get distance
      If TestReg = 0# Then                                  'if x is 0
        DisplayReg = 0#                                     'then angle is 0
      Else
        DisplayReg = RadToAng(Atn(DisplayReg / TestReg))    'else get angle
      End If
      TestReg = TV
      DisplayText = False
    
' Push ---------------------------------------------------------------
    Case 269  ' Push
      PushIdx = PushIdx + 1
      If PushIdx > PushSize Then
        PushSize = PushSize + PushLimit
        ReDim Preserve PushValues(PushSize)
      End If
      PushValues(PushIdx) = DisplayReg
      Call UpdateStatus
      DisplayText = False
    
' Pop ----------------------------------------------------------------
    Case 270  ' Pop
      If CBool(PushIdx) Then
        DisplayReg = PushValues(PushIdx)
        PushIdx = PushIdx - 1
        Call UpdateStatus
      Else
        Errstr = "Push Stack was empty"
      End If
      DisplayText = False

' StkEx --------------------------------------------------------------
    Case 271  ' StkEx
      If CBool(PushIdx) Then
        TV = DisplayReg
        DisplayReg = PushValues(PushIdx)
        PushValues(PushIdx) = TV
      Else
        Errstr = "Push Stack was empty"
      End If
      DisplayText = False

' SUB ----------------------------------------------------------------
    Case 272  ' SUB
      ErrorFlag = False
      Set Vptr = CheckStruct()
      If ErrorFlag Then Exit Sub
      If Vptr Is Nothing Then Set Vptr = CheckIndVar()
      Call SUB_Run(Vptr)

' DIV ----------------------------------------------------------------
    Case 273  ' DIV
      ErrorFlag = False
      Set Vptr = CheckStruct()
      If ErrorFlag Then Exit Sub
      If Vptr Is Nothing Then Set Vptr = CheckIndVar()
      Call DIV_Run(Vptr)

' < ------------------------------------------------------------------
    Case 274  ' <
      Call Pend(iLT)
      
' > ------------------------------------------------------------------
    Case 275  ' >
      Call Pend(iGT)

' Skey ---------------------------------------------------------------
    Case 276  ' Skey
      Call CheckForText(Iptr - 1, DisplayWidth)
      DspTxt = TxtData                    'move to DspTxt for processing
      Call HkeySkeySupport
      For i = 1 To 26
        If Kyz(i) Then Hidden(i) = False  'mark key as not hidden
      Next i
      Call RedoAlphaPad
      DisplayText = False

' eX -----------------------------------------------------------------
    Case 277  ' eX
      DisplayReg = Exp(DisplayReg)
      DisplayText = False

' E- -----------------------------------------------------------------
    Case 278  ' E-
      Call StatSUMsub

' StDev --------------------------------------------------------------
    Case 279  ' StDev
      Call StatStdDev

' Varnc --------------------------------------------------------------
    Case 280  ' Varnc
      Call StatVarnc

' Yint ---------------------------------------------------------------
    Case 281  ' Yint
      Call Yintercept

' LogX ---------------------------------------------------------------
    Case 286  ' LogX
      Call Pend(iLogX)
    
' Var ----------------------------------------------------------------
    Case 287  ' Var
      Call Var_IND(CheckIndVar())
    
' == -----------------------------------------------------------------
    Case 288  ' ==
      Call Pend(iEQ)

' && -----------------------------------------------------------------
    Case 289  ' &&
      Call Pend(iAnd)
    
' RFlg ---------------------------------------------------------------
    Case 290  ' RFlg
      i = GetInstruction(1)                     'get next code
      Select Case i
        Case 0 To 9                             '0-9
          flags(i) = False                      'reset flag
          IncInstrPtr                           'bump instruction pointer
        Case iIND                               'indirection?
          Set Vptr = CheckVbl()                 'point to desired storage object
          Call RFlg_IND(Vptr)
      End Select
    
' !Flg ---------------------------------------------------------------
    Case 291  ' !Flg
      i = GetInstruction(1)                     'get next code
      Select Case i
        Case 0 To 9                             '0-9
          Bol = Not flags(i)                    'set result
          IncInstrPtr                           'point to '{'-1
          Call BuildBraceStk(InstrPtr, -1, -1, Bol)
        Case iIND                               'indirection?
          Set Vptr = CheckVbl()                 'point to desired storage object
          Call NotFlg_IND(Vptr)
      End Select
    
' X!=T ---------------------------------------------------------------
    Case 292  ' X!=T
      Bol = DisplayReg <> TestReg           'set result
      Call BuildBraceStk(InstrPtr + 1, -1, -1, Bol)
      
' X<=T ---------------------------------------------------------------
    Case 293  ' X<=T
      Bol = DisplayReg <= TestReg           'set result
      Call BuildBraceStk(InstrPtr + 1, -1, -1, Bol)
    
' X<T ----------------------------------------------------------------
    Case 294  ' X<T
      Bol = DisplayReg < TestReg            'set result
      Call BuildBraceStk(InstrPtr + 1, -1, -1, Bol)
    
' NOP ----------------------------------------------------------------
    Case 295  ' NOP
    
' ; ------------------------------------------------------------------
    Case iSemiC     ' ;  'ignored
    Case iSemiColon  ' ;  'for FOR statments
      With BracePool(BraceIdx)
        If ForInit Then             'if we are running initialization code in FOR
          Call Pend(iRparen)        'terminate process code
          If .LpCond = -1 Then      'if no condition, assume it is TRUE
            InstrPtr = .LpStart     'and point to start of loop
            DisplayReg = .LpDspReg  'recover Display Register data
          Else
            .LpDspReg = DisplayReg  'save Display Register
            InstrPtr = .LpCond      'point to conditional
            Call Pend(iLparen)      'begin conditional check
          End If
          ForInit = False           'turn off flag
        Else                        'running into ';' from Conditional code in FOR
          Call Pend(iRparen)        'terminate conditional code
          .LpTrue = CBool(DisplayReg) 'set truth of conditional
          DisplayReg = .LpDspReg    'recover display register
          If .LpTrue Then           'if condition is true
            InstrPtr = .LpStart     'we need to transfer control to LpStart
          Else
            InstrPtr = .LpTerm      'else we are terminating
          End If
        End If
      End With

' Log ----------------------------------------------------------------
    Case 297  ' Log
      DisplayReg = Log(DisplayReg) / Log(10#)        'Common logarithm
      DisplayText = False
    
' 10^ ----------------------------------------------------------------
    Case 298  ' 10^
      DisplayReg = Exp(DisplayReg * Log(10#))        '10 to power of DisplayReg
      DisplayText = False
    
' /= -----------------------------------------------------------------
    Case 299  ' /=
      Call Pend(iDivEq)
      
' Fmt ----------------------------------------------------------------
    Case 300  ' Fmt
      Select Case GetInstruction(1)
        Case iIND
          Set Vptr = CheckVbl()                               'yes, grab variable
          If Not Vptr Is Nothing Then                         'error?
            If Variables(Vptr.VarRoot).VarType = vString Then 'no, but is it text?
              DspFmt = CStr(ExtractValue(Vptr))               'yes, so grab its data
            Else
              Errstr = "Indirection error"
            End If
          End If
        Case Else
          Call CheckForText(Iptr, DisplayWidth)               'grab current text
          DspFmt = TxtData                                    'grab format
      End Select
      DisplayText = False

' != -----------------------------------------------------------------
    Case 301  ' !=
      Call Pend(iNEQ)

' || -----------------------------------------------------------------
    Case 302  ' ||
      Call Pend(iOr)

' Frac ---------------------------------------------------------------
    Case 303  ' Frac
      DisplayReg = DisplayReg - Fix(DisplayReg)
      DisplayText = False

' Sgn ----------------------------------------------------------------
    Case 304  ' Sgn
      DisplayReg = CDbl(Sgn(DisplayReg))
      DisplayText = False

' !Fix ---------------------------------------------------------------
    Case 305  ' !Fix
      DspFmtFix = -1              'disable fixed-decimal places
      DspFmt = DefDspFmt          'set default format
      ScientifEE = DefScientific  'reset default scientif mode
      DisplayText = False

' DD.ddd -------------------------------------------------------------
    Case 306  ' D.ddd
      vDeg = Fix(DisplayReg)                            'get DDD
      DisplayReg = (DisplayReg - vDeg) * 3600#          'get seconds
      vMin = Fix(DisplayReg / 60#)                      'get MM
      vSec = (DisplayReg - vMin * 60#)                  'get SS.dddd
      DisplayReg = vDeg + vMin / 100# + vSec / 10000#   'get dd.ddddd
      DisplayText = False

' !EE ----------------------------------------------------------------
    Case 307  ' !EE
      If EEMode Then                                    'if EE mode was on...
        EEMode = False                                  'turn it off
        EngMode = False                                 'and Eng Mode
        Call UpdateStatus                               'reflect it on the status
        DisplayReg = CDbl(Format(DisplayReg, ScientifEE)) 'eliminate any rounding errors
      End If
      DisplayText = False

' Call ---------------------------------------------------------------
    Case 308  ' Call
      If GetInstruction(1) = iIND Then                  'indirection used?
        Set Vptr = CheckVbl()                           'yes, grab variable
        If Vptr Is Nothing Then Exit Sub                'error
        TV = Fix(Val(ExtractValue(Vptr)))               'no, so grab value stored there
        OldPgm = ActivePgm                              'save active program
        If CBool(ModPrep) And ModPrep <> ActivePgm Then
          ActivePgm = ModPrep                           'set ActivePgm to new Pgm
          If RunMode Then
            MRunMode = MRunMode + 1                     'temp pgm invoke if true
          End If
        End If
        ModPrep = 0                                     'disable flag
        i = GetInstrCnt(ModPrep)                        'get instruction count
        If TV < 0# Or TV > CDbl(i) Then                 'in range?
          Errstr = "Indirection out of range"           'no
          Exit Sub
        Else
          Call PushCall(0, CInt(TV), OldPgm)            'push return location
        End If
      Else  'label was provided
        Call CheckForLabel(InstrPtr, LabelWidth)        'grab called label
        OldPgm = ActivePgm                              'save active program
        If CBool(ModPrep) And ModPrep <> ActivePgm Then
          ActivePgm = ModPrep                           'set ActivePgm to new Pgm
          If RunMode Then
            MRunMode = MRunMode + 1                     'temp pgm invoke if true
          End If
        End If
        ModPrep = 0                                     'disable flag
        
        JL = FindLblMatch(TxtData)                      'get Lbls() index
        If CBool(ActivePgm) Then                        'if possible new is user program
          Pool = ModLbls                                'else use module's
        Else
          Pool = Lbls                                   'use user pool of labels
        End If
        
        With Pool(JL)                                   'using appropriate labels pool...
          Select Case .LblTyp
            Case TypKey, TypSbr                         'allow only Sbr and Ukey
              Call PushCall(JL, .lblAddr, OldPgm)       'push location
            Case Else
              Errstr = "Cannot invoke this label"
              ActivePgm = OldPgm                        'reset activepgm, if different
              Exit Sub
          End Select
        End With
      End If
      
      Call Run
' if after running, the Pmt command is activated, we will turn on the TXT mode (TextEntry),
' let the user type in a response, and by pressing TXT or '=' (ENTER), the program will continue
' (see KybdMain).
      If CBool(MRunMode) Then MRunMode = MRunMode - 1
      If PmtFlag Then                     'if user prompting turned on...
        LastTypedInstr = iTXT             'set TXT command
        Call ActiveKeypad                 'activate it (TXT or [=] (ENTER)) will start R/S cmd
      End If

' Trim ---------------------------------------------------------------
    Case 309  ' Trim
    Set Vptr = Trim_Run()
    If Not Vptr Is Nothing Then
      Vptr.VarStr = Trim$(Vptr.VarStr)
    End If

' LTrim --------------------------------------------------------------
    Case 310  ' LTrim
    Set Vptr = Trim_Run()
    If Not Vptr Is Nothing Then
      Vptr.VarStr = LTrim$(Vptr.VarStr)
    End If

' RTrim --------------------------------------------------------------
    Case 311  ' RTrim
    Set Vptr = Trim_Run()
    If Not Vptr Is Nothing Then
      Vptr.VarStr = RTrim$(Vptr.VarStr)
    End If
    
' *= -----------------------------------------------------------------
    Case 312  ' *=
      Call Pend(iMulEq)
    
' >= -----------------------------------------------------------------
    Case 314  ' >=
      Call Pend(iGE)

' ! ------------------------------------------------------------------
    Case 315  ' !
      Call Pend(iNot)
    
' Open ---------------------------------------------------------------
    Case 316  ' Open
      If Len(StorePath) = 0 Then
        Errstr = "No storage path yet defined"
        Exit Sub
      End If
'
' first grab filename to S
'
      If GetInstruction(1) = iIND Then              'we can specify a text variable
        Set Vptr = CheckVbl()
        If Vptr Is Nothing Then Exit Sub
        With Variables(Vptr.VarRoot)
          If .VarType <> vString Then
            Errstr = "Variable must be a Text type"
            Exit Sub
          End If
          S = CStr(ExtractValue(Vptr))              'grab filename
        End With
      Else                                          'else assume user-suppied filename
        Call CheckForLabel(InstrPtr, LabelWidth)
        S = TxtData
      End If
'
' next, check the I/O type
'
      InstrPtr = InstrPtr + 4               'point past 'For' and R|W|A|B and 'As'
      T = UCase$(Chr$(GetInstruction(-2)))  'get R|W|A|B code
'
' next check for I/O port
'
      i = GetInstruction(0)
      Select Case i
        Case 1 To 9                         'allow ports 1-9
        Case iIND                           'allow specifying an indirect variable
          Set Vptr = CheckVbl()
          If Vptr Is Nothing Then Exit Sub
          TV = Val(ExtractValue(Vptr))
          If TV < 1# Or TV > 9# Then
            Errstr = "Parameter is out of range (1-9)"
            Exit Sub
          End If
          i = CInt(TV)
        Case Else
          Errstr = "Invalid parameter"
          Exit Sub
      End Select
'
' now check for Len token (Required by Block mode)
'
      If GetInstruction(1) = iLen Then
        IncInstrPtr
        If GetInstruction(1) = iIND Then  'allow indrect variable
          Set Vptr = CheckVbl()
          If Vptr Is Nothing Then Exit Sub
          TV = Val(ExtractValue(Vptr))
          If TV < 1# Or T > 32767# Then
            Errstr = "Parameter is out of range (1-32767)"
            Exit Sub
          End If
          iX = CLng(TV)
        Else
          Call CheckForNumber(InstrPtr, 5, 32767)
          iX = CLng(TstData)
        End If
      End If
'
' error if a length specified, but File I/O is not Random (B)
'
      If CBool(j) And T <> "B" Then
        Errstr = "Cannot specify a LEN parameter except for the [B]lock I/O mode"
        Exit Sub
      ElseIf Not CBool(j) And T = "B" Then
        Errstr = "You must specify a LEN parameter for the [B]lock I/O mode"
        Exit Sub
      End If
'
' perform Open operation
'
      S = StorePath & "\DATA\" & S            'set full filepath
      
      On Error Resume Next
      Select Case T
        Case "R"
          Set Tstrm(i) = Fso.OpenTextFile(S, ForReading, False)
        Case "W"
          Set Tstrm(i) = Fso.OpenTextFile(S, ForWriting, True)
        Case "A"
          Set Tstrm(i) = Fso.OpenTextFile(S, ForAppending, True)
        Case "B"
          Set Tstrm(i) = Nothing  'ensure object is null
          Open S For Random Access Read Write As #i Len = iX
      End Select
      Call CheckError
      On Error GoTo 0
      
      If Not ErrorFlag Then 'if no errors
        With Files(i)
          .FileNum = i      'mark file as open
          .FileLen = j      'save length (for use in Block mode)
          .FileRec = 1      'used by block to track default location
        End With
      End If
      
' Close --------------------------------------------------------------
    Case 317  ' Close
      i = GetInstruction(1)
      Select Case i
        '---------
        Case 1 To 9 'close specified file
          With Files(i)
            If CBool(.FileNum) Then 'if file is actually open
              If Tstrm(i) Is Nothing Then
                Close #i          'is block, so close it
              Else
                Tstrm(i).Close    'close it
                Set Tstrm(i) = Nothing
              End If
              .FileNum = 0        'mark as closed
            End If
          End With
          Exit Sub                'all done
        '---------
        Case iAll                 'fall through
        '---------
        Case 0
          Errstr = "Parameter is out of range (1-9)"
          Exit Sub
      End Select
'
' is ALL or nothing, so close all open files
'
      Call CloseAll
    
' Read ---------------------------------------------------------------
    Case 318  ' Read
      Call FileIO(iRead, Errstr)
      
' Write --------------------------------------------------------------
    Case 319  ' Write
    Call FileIO(iWrite, Errstr)
    
' Swap ---------------------------------------------------------------
    Case 320  ' Swap
      Set Vptr = CheckIndVar()                          'get first value
      If Not Vptr Is Nothing Then                       'if OK so far
        If GetInstruction(1) = iComma Then              'expected comma?
          IncInstrPtr                                   'bump to comma
          Set sPtr = CheckIndVar()                      'get 2nd var
          If Not sPtr Is Nothing Then                   'if 2nd OK
            Tmp = ExtractValue(Vptr)                    'get first value
            Call StuffValue(Vptr, ExtractValue(sPtr))   'stuff 2nd to 1st
            Call StuffValue(sPtr, Tmp)                  'stuff 1st to 2nd
          End If
        End If
      End If
    
' GTO ----------------------------------------------------------------
    Case 321  ' GTO
    If GetInstruction(1) = iIND Then                  'indirection used?
      Set Vptr = CheckVbl()                           'yes, grab variable
      Call GTO_IND(Vptr)                              'set instruction pointer
    Else
      Call CheckForLabel(InstrPtr, LabelWidth)        'grab called label
      If Len(TxtData) = 1 Then
        Select Case UCase$(TxtData)
          Case "A" To "Z"
            JL = Asc(UCase$(TxtData)) - 64            'get key offset
            If CBool(ActivePgm) Then                  'module?
              JL = JL + ModLblMap(ActivePgm - 1)
              If ModLbls(JL).LblDat = 0 Then JL = 0   'if not defined in module
            Else
              If Lbls(JL).LblDat = 0 Then JL = 0      'if not defined in pgm
            End If
          Case Else
            JL = 0
        End Select
      Else
        JL = FindLblMatch(TxtData)                    'search for matching name
      End If
      
      
      If CBool(ActivePgm) Then
        Pool = ModLbls
      Else
        Pool = Lbls
      End If
      
      With Pool(JL)
        Select Case .LblTyp
          Case TypKey, TypSbr, TypLbl                 'allow only Sbr and Ukey and Lbl
            InstrPtr = .LblDat                        'set new pointer to its data adress-1
          Case Else
            Errstr = "Cannot Go To this Label"
        End Select
      End With
    End If
    
' LOF ----------------------------------------------------------------
    Case 322  ' LOF
      DisplayText = False
      If GetInstruction(1) = iIND Then                'indirection?
        Set Vptr = CheckVbl()                         'get pointer to variable
        If Vptr Is Nothing Then Exit Sub
        TstData = Fix(Val(ExtractValue(Vptr)))        'get value
        If TstData < 1# Or TV > 9# Then
          Errstr = "Expected value 1-9"
          Exit Sub
        End If
      Else
        Call CheckForNumber(InstrPtr, 1, 9)           'allow 0-9
      End If
      On Error Resume Next
      DisplayReg = CDbl(LOF(CInt(TstData)))
      Call CheckError
      On Error GoTo 0

' Get ----------------------------------------------------------------
    Case 323  ' Get
    Call FileIO(iGet, Errstr)
    
' Put ----------------------------------------------------------------
    Case 324  ' Put
    Call FileIO(iPut, Errstr)

' -= -----------------------------------------------------------------
    Case 325  ' -=
      Call Pend(iSubEq)
    
' SysBeep ------------------------------------------------------------
    Case 326  ' sysBP
      Call CheckForNumber(InstrPtr, 1, 4)
      Select Case CInt(TstData)
        Case 1
          Idx = beepSystemAsterisk
        Case 2
          Idx = beepSystemExclamation
        Case 3
          Idx = beepSystemHand
        Case 4
          Idx = beepSystemQuestion
        Case Else
          Idx = beepSystemDefault
      End Select
      Call MsgBeep(Idx)
    
' <= -----------------------------------------------------------------
    Case 327  ' <=
      Call Pend(iLE)
    
' Nor ----------------------------------------------------------------
    Case 328  ' Nor
      Call Pend(iNor)
    
' Incr ---------------------------------------------------------------
    Case 329  ' Inc
      Call Incr_IND(CheckIndVar())
    
' Decr ---------------------------------------------------------------
    Case 330  ' Dec
      Call Decr_IND(CheckIndVar())
      
' Dsz ----------------------------------------------------------------
    Case 331  ' Dsz
      Call Dsz_IND(CheckIndVar())
    
' Dsnz ---------------------------------------------------------------
    Case 332  ' Dsnz
      Call Dsnz_IND(CheckIndVar())
      
' All ----------------------------------------------------------------
    Case 333  ' All  'ignored
    
' Rtn ----------------------------------------------------------------
    Case 334  ' Rtn
      If CBool(SbrInvkIdx) Then       'if something is on the stack
        With SbrInvkStk(SbrInvkIdx)
          ActivePgm = .Pgm            'reset the invoking program number
          InstrPtr = .PgmInst         'set the return address-1
          BraceIdx = .PgmBrcIdx       'clear brace index to what it was before (quick clean stack)
        End With
        SbrInvkIdx = SbrInvkIdx - 1   'back off the stack
      Else
        RunMode = False               'disable Run mode
        ModPrep = 0
        Call ResetPndAll              'reset data
        StopMode = True               'set stop mode
      End If
    
' LSet ---------------------------------------------------------------
    Case 335  ' LSet
      Set Vptr = CheckIndVar()        'check for variable and for indirection
      If Vptr Is Nothing Then Exit Sub
      With Variables(Vptr.VarRoot)
        If .VarType <> vString Then   'text type?
          Errstr = "Variable must be a Text type"
        Else
          i = .VdataLen               'grab fixed width
          If i = 0 Then               'if not fixed width
            Errstr = "Variable is not fixed length"
          Else
            S = Trim$(Vptr.VarStr)    'grab string, trimmed up
            If Len(S) < i Then        'if the string is less than the fixed width
              Vptr.VarStr = S & String$(i - Len(S), 32)
            End If
          End If
        End If
      End With
    
' RSet ---------------------------------------------------------------
    Case 336  ' RSet
      Set Vptr = CheckIndVar()        'check for variable and for indirection
      If Vptr Is Nothing Then Exit Sub
      With Variables(Vptr.VarRoot)
        If .VarType <> vString Then   'text type?
          Errstr = "Variable must be a Text type"
        Else
          i = .VdataLen               'grab fixed width
          If i = 0 Then               'if not fixed width
            Errstr = "Variable is not fixed length"
          Else
            S = Trim$(Vptr.VarStr)    'grab string, trimmed up
            If Len(S) < i Then        'if the string is less than the fixed width
              Vptr.VarStr = String$(i - Len(S), 32) & S
            End If
          End If
        End If
      End With
    
' Printf -------------------------------------------------------------
    Case 337  ' Printf
      Set Vptr = CheckIndVar()                                  'grab variable
      If Not Vptr Is Nothing Then                               'if valid
        If Variables(Vptr.VarRoot).VarType = vString Then       'string?
          DspTxt = Format(DisplayReg, CStr(ExtractValue(Vptr))) 'yes, get format
          DisplayText = True                                    'indicate text display
          Call ForceDisplay                                     'display it
          Call NewLine                                          'advance to next line
          DisplayText = False
        Else
          Errstr = "Variable must be a Text type"
        End If
      End If
    
' += -----------------------------------------------------------------
    Case 338  ' +=
      Call Pend(iAddEq)
    
' RGB ----------------------------------------------------------------
    Case 339  ' RGB
      i = RGB_Support(iLparen, "(", Errstr) 'get RED
      If i = -1 Then Exit Sub
      j = RGB_Support(iComma, ",", Errstr)  'get Green
      If j = -1 Then Exit Sub
      K = RGB_Support(iComma, ",", Errstr)  'get Blue
      If K = -1 Then Exit Sub
      If GetInstruction(1) <> iRparen Then
        Errstr = "Expected ')'"
        Exit Sub
      End If
      DisplayReg = CDbl(RGB(i, j, K))
      DisplayText = False
      
' \ ------------------------------------------------------------------
    Case 341  ' \
      Call Pend(iBkSlsh)
    
' As -----------------------------------------------------------------
    Case 342  ' As   'ignored

' DBG ----------------------------------------------------------------
    Case 344  ' DBG
      Select Case GetInstruction(1)
        Case iOpen                          'check for DBG Open
          IncInstrPtr                       'skip new instruction
          If Not CBool(ActivePgm) Then      'if Pgm 00...
            TraceFlag = True                'enable tracing
            Tron = True
            frmVisualCalc.mnuFileTron.Checked = True
            Call UpdateStatus
          End If
          Exit Sub
        Case iClose                         'check for DBG Close
          IncInstrPtr                       'skip new instruction
          If Not CBool(ActivePgm) Then      'if Pgm 00...
            TraceFlag = False               'disable tracing
            Tron = False
            frmVisualCalc.mnuFileTron.Checked = False
            Call UpdateStatus
          End If
          Exit Sub
      End Select
      
      MRunMode = 0              'disable Module Run mode
      RunMode = False           'disable Run mode
      ModPrep = 0
      Call ResetPndAll          'reset data
      DoDebug = (ActivePgm = 0)
      DisplayText = False

' Gfree --------------------------------------------------------------
    Case 345  ' Gfree
    DisplayReg = 0#
    For Idx = 1 To 9
      If Files(Idx).FileNum = 0 Then
        DisplayReg = CDbl(Idx)
      End If
    Next Idx
    DisplayText = False

' Len ----------------------------------------------------------------
    Case 346  ' Len  'Ignored

' Stop ---------------------------------------------------------------
    Case 347  ' Stop
      RunMode = False           'disable Run mode
      ModPrep = 0
      Call ResetPndAll          'reset data
      StopMode = True

' With ---------------------------------------------------------------
    Case 348  ' With 'Ignored

' ',' ----------------------------------------------------------------
    Case iComma  ' ,  'ignored

' Val ----------------------------------------------------------------
    Case 350  ' Val
      DisplayReg = Val(DspTxt)                  'derive value from displayed data
      DisplayText = False

' Adv ----------------------------------------------------------------
    Case 351  ' Adv
      Call ForceDisplay                         'first force the display of current data
      If IsPending Then                         'anything pending?
        With frmVisualCalc.lstDisplay
          If PendOpn(PendIdx) = iMinus Then     '-ADV?
            If .ListIndex > 0 Then              'if we can back up
              Call SelectOnly(.ListIndex - 1)   'else select previous line
            End If
            PendIdx = PendIdx - 1               'back off pending index
            Exit Sub
          ElseIf PendOpn(PendIdx) = iAdd Then   '+ADV?
            If .ListIndex = .ListCount - 1 Then 'if we are at the end of the list
              .AddItem vbNullString             'force a new line
              Call SelectOnly(.ListCount - 1)   'select that line
            Else
              Call SelectOnly(.ListIndex + 1)   'else select next line
            End If
            PendIdx = PendIdx - 1               'back off pending index
            Exit Sub                            'avoid falling into NewLine
          End If
        End With
      End If
      Call NewLine                              'default advance
    
' << -----------------------------------------------------------------
    Case 354  ' <<
      Do
        i = GetInstruction(1)                   'get next code
        IncInstrPtr
        If i < 1 Or i > 9 Then                  'in range?
          Errstr = "Parameter is out of range (1-9)"
          Exit Do
        End If
        If DisplayReg = 0# Then Exit Do         'if null, then nothing to do
        On Error Resume Next
        For Idx = 1 To i
          DisplayReg = DisplayReg * 2#
          Call CheckError
          If ErrorFlag Then Exit For
        Next Idx
        Exit Do
      Loop
      DisplayText = False
    
' Root ---------------------------------------------------------------
    Case 355  ' Root
      Call Pend(iRoot)
      DisplayText = False
    
' Sqrt ---------------------------------------------------------------
    Case 356  ' Sqrt
      On Error Resume Next
      DisplayReg = Sqr(DisplayReg)  'get square root
      Call CheckError               'check for error
      On Error GoTo 0               'reset error processing
      DisplayText = False
      
' e ------------------------------------------------------------------
    Case 357  ' e
      DisplayReg = vE
      DisplayText = False

'---------------------------------------------------------------------
    Case 358  ' Rnd#
      DisplayReg = Rnd(DisplayReg)
      Randomize
      DisplayText = False

' Until --------------------------------------------------------------
'    Case 359  ' Until 'Will not be encountered. Handled by DO-processing

'---------------------------------------------------------------------
    Case 360  ' Pub  'ignored

' Enum ---------------------------------------------------------------
    Case 361  ' Enum
      If GetInstruction(1) < 10 Then                  'initial value for enumerations?
        Call CheckForNumber(InstrPtr, LabelWidth, 0)  'skip over
      End If
      IncInstrPtr                                     'point to '{'
      InstrPtr = FindEblock(InstrPtr)                 'set instruction pointer to end of block
      
'---------------------------------------------------------------------
    Case 362  ' AdrOf
      Call CheckForLabel(InstrPtr, LabelWidth)          'get label to TxtData
      If Len(TxtData) = 1 Then                          'if len=1, then A-Z user-key
        i = -1                                          'force not found
      Else
        i = FindVblMatch(TxtData)                       'defined as a variable?
      End If
      If i = -1 Then                                    'no
        If Len(TxtData) = 1 Then                        'if len=1, then A-Z user-key
          JL = Asc(UCase$(TxtData)) - 64                'get 1-26 index value
        Else
          JL = FindLblMatch(TxtData)                    'find a match
        End If
        If CBool(JL) Then                               'found a match?
          If CBool(ActivePgm) Then
            Pool = ModLbls
          Else
            Pool = Lbls
          End If
          
          With Pool(JL)                                 'yes, see if it is the proper type
            Select Case .LblTyp
              Case TypSbr, TypKey
                i = .lblAddr                            'address of definition block
              Case Else
                Errstr = "Cannot obtain AdrOf this item"
                Exit Sub
            End Select
          End With
        Else
          Errstr = "Parameter error"
          Exit Sub
        End If
      End If
      DisplayReg = CDbl(i)                              'get ultimate address (var# if variable)
      DisplayText = False

'---------------------------------------------------------------------
    Case 363  ' Pcmp  'ignored
    Case 364  ' Comp  'ignored
    
'---------------------------------------------------------------------
    Case 365  ' Circle
    IncInstrPtr                         'point to data
    If Not GetPlotXY(Xs, Ys, Errstr) Then Exit Sub
    If Not Circle_Support(Radius, Errstr) Then Exit Sub
    If Not Circle_Support(ArcStart, Errstr) Then Exit Sub
    If Not Circle_Support(ArcEnd, Errstr) Then Exit Sub
    If Not Circle_Support(Aspect, Errstr) Then Exit Sub
    If GetInstruction(1) = iComma Then  '[,z]
      IncInstrPtr
      IncInstrPtr
      i = GetInstruction(0)             'get 0,1
    Else
      i = 0
    End If
'
' if ArcStart and ArcEnd are 0, then set for cull circle,
' otherwise, convert ArcStart and ArcEnd to Radians
'
    ArcStart = CSng(AngToRad(CDbl(ArcStart)))
    ArcEnd = CSng(AngToRad(CDbl(ArcEnd)))
    If ArcStart >= vPi2 Then ArcStart = ArcStart - vPi2
    If ArcEnd >= vPi2 Then ArcEnd = ArcEnd - vPi2
    If ArcStart = ArcEnd Then ArcStart = ArcStart - 0.0001
'
' if Aspeci not supplied, use default of 1.0
'
    If Aspect = 0! Then Aspect = 1!
'
' draw circle or arc
'
    frmVisualCalc.PicPlot.Circle (CSng(Xs + PlotXOfst), CSng(Ys + PlotYOfst)), Radius, PlotColor, ArcStart, ArcEnd, Aspect
'
' flood fill if requested
'
    If CBool(i) Then Call PaintIt(Xs, Ys, PlotColor)

'---------------------------------------------------------------------
    Case 366  ' Split
      Call Split_Support(Errstr)
      
'---------------------------------------------------------------------
    Case 367  ' Join
      Call Join_Support(Errstr)

'---------------------------------------------------------------------
    Case 368  ' ReDim
      Set Vptr = CheckSngVbl()                            'get variable (do not check dims)
      If Not Vptr Is Nothing Then                         'defined?
        Idx = CLng(Vptr.VarRoot)                          'grab variable number
        If CheckRunDim2(iX, iY) Then                      'if dimension dims found...
          Call BuildMDAry(CLng(Vptr.VarRoot), iX, iY, True) 'process them, and preserve old data
        End If
      End If
      
'---------------------------------------------------------------------
    Case 369  ' Mid
      IncInstrPtr                         'skip '(' known to be here
      Set Vptr = CheckIndVar()
      If Vptr Is Nothing Then Exit Sub
      If Variables(Vptr.VarRoot).VarType <> vString Then
        Errstr = "Variable must be a Text type"
        Exit Sub
      End If
      S = CStr(ExtractValue(Vptr))        'get string there
'
'now grab index into string
'
      IncInstrPtr                         'skip comma known to be here
      If GetInstruction(1) = iIND Then    'are reference to a variable?
        Set Vptr = CheckVbl()             'get variable pointed
        If Vptr Is Nothing Then Exit Sub
        TstData = Val(ExtractValue(Vptr)) 'get value there
      Else
        Call CheckForNumber(InstrPtr, 2, 99)
      End If
      If TstData < 1# Or TstData > CDbl(DisplayWidth) Then
        Errstr = "Index for MID is out of range"
      Else
        i = CInt(TstData)                 'keep index
      End If
'
'now grab length of data to take from index
'
      IncInstrPtr                         'skip comma known to be here
      If GetInstruction(1) = iIND Then    'are reference to a variable?
        Set Vptr = CheckVbl()             'get variable pointed
        If Vptr Is Nothing Then Exit Sub
        TstData = Val(ExtractValue(Vptr)) 'get value there
      Else
        Call CheckForNumber(InstrPtr, 2, 99)
      End If
      If TstData < 1# Or TstData > CDbl(DisplayWidth) Then
        Errstr = "Index for MID is out of range"
      Else
        j = CInt(TstData)                 'keep length
      End If
      IncInstrPtr                         'skip ')' known to be here
'
' now obtain a result
'
      DspTxt = Mid$(S, i, j)              'aquire data
      DisplayText = True                  'we can display text

' Udef ---------------------------------------------------------------
    Case 370  ' Udef
      Call CheckForLabel(InstrPtr, LabelWidth)        'get label to TxtData
      i = FindDefMatch(TxtData)                       'already defined?
      If CBool(i) Then                                'if definition exists
        DefName(i) = vbNullString                     'remove it
      End If

' !Def ---------------------------------------------------------------
    Case 371  ' !Def
      Call CheckForLabel(InstrPtr, LabelWidth)        'get label to TxtData
      i = FindDefMatch(TxtData)                       'already defined?
      DefDef = DefDef + 1                             'indicate we are processing Def blocks
      If DefDef > DefTrueSz Then
        DefTrueSz = DefTrueSz + 8
        ReDim Preserve DefTrue(DefTrueSz)
      End If
      DefTrue(DefDef) = Not CBool(i)                  'set Not True if defined

' Delse --------------------------------------------------------------
    Case 372  ' Delse
      If CBool(DefDef) Then
        DefTrue(DefDef) = Not DefTrue(DefDef)         'flip def flag
      Else
        Errstr = "Delse instruction, but no IfDef or !Def"
      End If

'---------------------------------------------------------------------
' EXTENDED FUNCTIONS: these are applied to a Compressed program
'---------------------------------------------------------------------

' StFlg IND --------------------------------------------------------------------
    Case 373: S = "StFlg IND"
      Set Vptr = CheckVbl()                 'point to desired storage object
      Call StFlg_IND(Vptr)
      
' RFlg IND ---------------------------------------------------------------------
    Case 374
      Set Vptr = CheckVbl()                 'point to desired storage object
      Call RFlg_IND(Vptr)

' IfFlg IND --------------------------------------------------------------------
    Case 375
      Set Vptr = CheckVbl()                 'point to desired storage object
      Call IfFlg_IND(Vptr)

' !Flg IND ---------------------------------------------------------------------
    Case 379
      Set Vptr = CheckVbl()                 'point to desired storage object
      Call NotFlg_IND(Vptr)

' Dsz IND ----------------------------------------------------------------------
    Case 388
      Call Dsz_IND(GrabInd())
      
' Dsnz IND ---------------------------------------------------------------------
    Case 391
      Call Dsnz_IND(GrabInd())

' Incr IND ---------------------------------------------------------------------
    Case 394
      Call Incr_IND(GrabInd())

' Decr IND ---------------------------------------------------------------------
    Case 395
      Call Decr_IND(GrabInd())

' Asin -------------------------------------------------------------------------
    Case 400  ' Asin
      X = DisplayReg
      On Error Resume Next
      DisplayReg = RadToAng(Atn(X / Sqr(-X * X + 1#)))
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False

' Acos -------------------------------------------------------------------------
    Case 401  ' Acos
      X = DisplayReg
      On Error Resume Next
      DisplayReg = RadToAng(Atn(-X / Sqr(-X * X + 1#)) + vRA)
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False

' Atan -------------------------------------------------------------------------
    Case 402  ' Atan
      On Error Resume Next
      DisplayReg = RadToAng(Atn(DisplayReg))
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False

' SinH -------------------------------------------------------------------------
    Case 403  ' SinH
      X = DisplayReg
      On Error Resume Next
      DisplayReg = (Exp(X) - Exp(-X)) / 2#
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False
    
' CosH -------------------------------------------------------------------------
    Case 404  ' CosH
      X = DisplayReg
      On Error Resume Next
      DisplayReg = (Exp(X) + Exp(-X)) / 2#
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False
    
' TanH -------------------------------------------------------------------------
    Case 405  ' TanH
      X = DisplayReg
      On Error Resume Next
      DisplayReg = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False
      
' ArcSinH ----------------------------------------------------------------------
    Case 406  ' ArcSinH
      X = DisplayReg
      On Error Resume Next
      DisplayReg = Log(X + Sqr(X * X + 1#))
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
      DisplayText = False
    
' ArcCosH ----------------------------------------------------------------------
    Case 407  ' ArcCosH
      X = DisplayReg
      On Error Resume Next
      DisplayReg = Log(X + Sqr(X * X - 1#))
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False
    
' ArcTanH ----------------------------------------------------------------------
    Case 408  ' ArcTanH
      X = DisplayReg
      On Error Resume Next
      DisplayReg = Log((1# + X) / (1# - X)) / 2#
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False
    
' Asec -------------------------------------------------------------------------
    Case 409  ' Asec      'Acos(1/x)
      On Error Resume Next
      X = 1# / DisplayReg
      DisplayReg = RadToAng(Atn(-X / Sqr(-X * X + 1#)) + vRA)
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False
    
' Acsc -------------------------------------------------------------------------
    Case 410  ' Acsc      'Asin(1/x)
      On Error Resume Next
      X = 1# / DisplayReg
      DisplayReg = RadToAng(Atn(X / Sqr(-X * X + 1#)))
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False
    
' Acot -------------------------------------------------------------------------
    Case 411  ' Acot      'Atan(1/x)
      On Error Resume Next
      DisplayReg = RadToAng(Atn(1# / DisplayReg))
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False
    
' SecH -------------------------------------------------------------------------
    Case 412  ' SecH      '1/CosH(x)
      X = DisplayReg
      On Error Resume Next
      DisplayReg = 1# / ((Exp(X) + Exp(-X)) / 2#)
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False
    
' CscH -------------------------------------------------------------------------
    Case 413  ' CscH      '1/SinH(x)
      X = DisplayReg
      On Error Resume Next
      DisplayReg = 1# / ((Exp(X) - Exp(-X)) / 2#)
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False
    
' CotH -------------------------------------------------------------------------
    Case 414  ' CotH      '1/TanH(x)
      X = DisplayReg
      On Error Resume Next
      DisplayReg = 1# / ((Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X)))
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False
    
' ArcSecH ----------------------------------------------------------------------
    Case 415  ' ArcSecH   '1/ArcCosH(x)
      X = DisplayReg
      On Error Resume Next
      DisplayReg = 1# / Log(X + Sqr(X * X - 1#))
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False
    
' ArcCscH ----------------------------------------------------------------------
    Case 416  ' ArcCscH   '1/ArcSinH(x)
      X = DisplayReg
      On Error Resume Next
      DisplayReg = 1# / Log(X + Sqr(X * X + 1#))
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False
    
' ArcCotH ----------------------------------------------------------------------
    Case 417  ' ArcCotH   '1/ArcTanH(x)
      X = DisplayReg
      On Error Resume Next
      DisplayReg = 1# / (Log((1# + X) / (1# - X)) / 2#)
      Call CheckError               'check for error
      On Error GoTo 0
      DisplayText = False

' Var IND ----------------------------------------------------------------------
    Case 433
    Call Var_IND(GrabInd())
    
'-------------------------------------------------------------------------------
' special Compressr functions
'-------------------------------------------------------------------------------
    Case 486  'Special Else for Case Statement 'ignore
    Case 509  'CnstNxt  'the next instruction is an index into LclConsts()
    Case 510  'DblNxt   'the next to instructions consitute a Double value
    Case 511  'IntNxt   'the following instruction is an integer value
' User Keys A-Z ----------------------------------------------------------------
    Case Is > 900 '<A>-<Z> command keys
      JL = Code - 900                                 'get Lbls() index
      If CBool(ActivePgm) Then
        Idx = ModLbls(JL + ModLblMap(ActivePgm - 1)).LblDat
      Else
        Idx = Lbls(JL).LblDat
      End If
      If Not CBool(Idx) Then
        ForcError "Selected user-defined key is not defined"
        Exit Sub
      End If
      OldPgm = ActivePgm                              'save active program
      If CBool(ModPrep) And ModPrep <> ActivePgm Then
        ActivePgm = ModPrep                           'set ActivePgm to new Pgm
        If RunMode Then
          MRunMode = MRunMode + 1                     'temp pgm invoke if true
        End If
      End If
      ModPrep = 0                                     'disable flag
      
      'push return location
      If CBool(ActivePgm) Then                        'if possible new is user program
        Call PushCall(JL, ModLbls(JL + ModLblMap(ActivePgm - 1)).lblAddr, OldPgm)
      Else
        Call PushCall(JL, Lbls(JL).lblAddr, OldPgm)
      End If
      
      Call Run
' if after running, the Pmt command is activated, we will turn on the TXT mode (TextEntry),
' let the user type in a response, and by pressing TXT or '=' (ENTER), the program will continue
' (see KybdMain).
      If CBool(MRunMode) Then MRunMode = MRunMode - 1
      If PmtFlag Then                     'if user prompting turned on...
        LastTypedInstr = iTXT             'set TXT command
        Call ActiveKeypad                 'activate it (TXT or [=] (ENTER)) will start R/S cmd
      End If
'-------------------------------------------------------------------------------
  End Select
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

