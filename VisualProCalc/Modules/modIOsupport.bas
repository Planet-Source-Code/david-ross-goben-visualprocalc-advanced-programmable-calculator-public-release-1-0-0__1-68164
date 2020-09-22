Attribute VB_Name = "modIOsupport"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hWnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long

Private Const SW_NORMAL = 1

'*******************************************************************************
' Subroutine Name   : Clear_Screen
' Purpose           : Clear display Screen
'*******************************************************************************
Public Sub Clear_Screen()
  With frmVisualCalc
    With .lstDisplay
      .Clear                          'clear list
      .AddItem vbNullString           'add a blank line to the display
      Call SelectOnly(0)              'select only this item
    End With
  End With
  DisplayText = False
  StoreList = False
  ModuleList = False
End Sub

'*******************************************************************************
' Function Name     : CvtEng
' Purpose           : Convert Display register to Engineering Notation
'*******************************************************************************
Public Function CvtEng() As String
  Dim S As String, T As String
  Dim i As Integer, j As Integer, K As Integer
  
  S = Format(DisplayReg, ScientifEE)                          'get basic scientific format
  T = Right$(S, 3)                                            'grab exponent and exp sign
  S = Left$(S, Len(S) - 4)                                    'strip expoment and 'E'
  S = Left$(S, 1) & Mid$(S, 3)                                'strip decimal
  i = CInt(Right$(T, 2))                                      'get positive exponent value
  j = (i Mod 3) + 1                                           'get whole decimal index
  If Left$(T, 1) = "-" Then                                   'negative exponent?
    Select Case i
      Case 1
        i = 5                                                 'do not allow 0 exp
        j = 3
        S = S & "00"
      Case 2
        i = 4                                                 'do not allow 0 exp
        j = 2
        S = S & "0"
      Case 3                                                  'nothing to do
      Case Else
        S = String$(3 - j, "0") & S                           'allow prepended zeros
    End Select
    S = Trim$(CStr(CInt(Left$(S, j))) & "." & Mid$(S, j + 1)) 'set mantissa
  ElseIf i < 3 Then
    i = i + 3
    S = String$(4 - j, "0") & S                               'allow prepended zeros
    S = Trim$(CStr(CInt(Left$(S, 1))) & "." & Mid$(S, 2))     'set mantissa
  Else
    S = Trim$(CStr(CInt(Left$(S, j))) & "." & Mid$(S, j + 1)) 'set mantissa
  End If
  T = "E" & Left$(T, 1) & Format(i - j + 1, "00")             'compute new exp value
  If Right$(S, 1) = "." Then S = S & "0"                      'allow at least 1 digit after '.'
  CvtEng = S & T                                              'final Eng value
End Function

'*******************************************************************************
' Function Name     : CvtTyp
' Purpose           : Convert Display Register to selected number base
'*******************************************************************************
Public Function CvtTyp(ByVal Typ As BaseTypes) As String
  Dim S As String, T As String, Nib As String
  Dim i As Integer, j As Integer, K As Integer
  
  Select Case Typ
    '--------------------------------------------
    Case TypDec                           'DECIMAL
      If EEMode Then
        If EngMode Then                   'Engineering mode?
          S = CvtEng()                    'yes, get Eng notation
        Else
          S = Format(DisplayReg, ScientifEE) 'EE mode active
        End If
      Else
        If CBool(Len(DspFmt)) Then
          T = DspFmt                      'employ user-defined format
        Else
          T = DefDspFmt                   'else use the default
        End If
        S = Trim$(Format(DisplayReg, T))  'format display value (trim out spaces)
        If Val(S) = 0# And CBool(DisplayReg) Then
          S = Format(DisplayReg, ScientifEE)
        ElseIf Len(S) > DisplayWidth Then
         S = Format(DisplayReg, ScientifEE)
        End If
      End If
    '--------------------------------------------
    Case TypHex                           'HEXADECIMAL
      On Error Resume Next
      S = Hex$(Fix(DisplayReg))           'grab non-fractional data
      Call CheckError
      On Error GoTo 0
    '--------------------------------------------
    Case TypOct                           'OCTAL
      On Error Resume Next
      S = Oct$(Fix(DisplayReg))           'grab non-fractional data
      Call CheckError
      On Error GoTo 0
    '--------------------------------------------
    Case TypBin                           'BINARY
      On Error Resume Next
      T = Hex$(Fix(DisplayReg))           'grab non-fractional data as a HEX string
      Call CheckError
      On Error GoTo 0
      If ErrorFlag Then Exit Function
      For K = Len(T) To 1 Step -1         'process nibbles from the right
        i = CInt("&h" & Mid$(T, K, 1))    'grab a nibble from the temp data
        Nib = vbNullString                'init 4-bit result
        For j = 1 To 4                    'process each of the 4 bits of the nibble
          Nib = Chr$(CByte(i And 1) + 48) & Nib 'add to nibble string (revered)
          i = i \ 2                       'remove the bit from the master nibble
        Next j
        S = Nib & S                       'prepend nibble to master string
      Next K
      j = Len(S)                          'now trim off leading zeros
      For i = 1 To j
        If Mid$(S, i, 1) = "1" Then Exit For 'found a 1 digit, so done
      Next i
      If i > j Then i = j                 'if result is 0, then, keep only digit
      S = Mid$(S, i)                      'trim any left-padded zeros
  '--------------------------------------------
  End Select
  CvtTyp = S                              'return type
End Function

'*******************************************************************************
' Function Name     : DisplaySetup
' Purpose           : Set up the appearance of the Display Register
'*******************************************************************************
Public Function DisplaySetup() As String
  Dim S As String
  
  S = CvtTyp(BaseType)                    'convert DisplayReg to current base type
  If Len(S) > DisplayWidth Then S = Format(DisplayReg, ScientifEE)
  DisplaySetup = S                        'return display data
End Function

'*******************************************************************************
' Subroutine Name   : ForceDisplay
' Purpose           : Force the display of the active data on the display line.
'                   : This is used by the DisplayLine subroutine, and in the
'                   : RUN mode when flashing something on the display.
'*******************************************************************************
Public Sub ForceDisplay()
  If Not DisplayText Then       'if we should display DspplayReg data...
    DspTxt = DisplaySetup()     'get formatted value of DisplayReg to DspTxt
  End If
  With frmVisualCalc.lstDisplay 'display on current line
    If Len(DspTxt) > DisplayWidth Then
      .List(.ListIndex) = DspTxt
    Else
      .List(.ListIndex) = String$(DisplayWidth - Len(DspTxt), 32) & DspTxt
    End If
  End With
  DoEvents
End Sub

'*******************************************************************************
' Subroutine Name   : DisplayLine
' Purpose           : Display the immediate line data
'*******************************************************************************
Public Sub DisplayLine()
  If CBool(CharLimit) Then Exit Sub   'ignore the following if user typing data
  'if run mode, do not bother to display, but ensure DspTxt is properly set
  If RunMode Or LrnMode Or CBool(MRunMode) Then
    If Not DisplayText Then           'if we should display DisplayReg data...
      DspTxt = DisplaySetup()         'get formatted value of DisplayReg to DspTxt
    End If
    Exit Sub
  End If
  Call ForceDisplay                   'display the active data
End Sub

'*******************************************************************************
' Subroutine Name   : DisplayAccum
' Purpose           : Display the immediate line data
'*******************************************************************************
Public Sub DisplayAccum()
  Dim S As String
  
  DisplayReg = GetDspValue()      'get value to reg
  If RunMode Or LrnMode Or CBool(MRunMode) Then Exit Sub
  If CharLimit = 0 Then
    S = ValueSgn & ValueAccum     'get display data
    With frmVisualCalc.lstDisplay 'display on current line
      .List(.ListIndex) = String$(DisplayWidth - Len(S), 32) & S
    End With
  End If
End Sub

'*******************************************************************************
' Function Name     : GetDspValue
' Purpose           : Return decimal value of display
'*******************************************************************************
Private Function GetDspValue() As Double
  Dim Idx As Long
  Dim IV As Double
  
  If Len(ValueAccum) = 0 Then ValueAccum = "0"  'ensure something there
  
  Select Case BaseType
    Case TypBin                         'BINARY
      GetDspValue = 0#                  'init result
      For Idx = 1 To Len(ValueAccum)    'translate series of 1's and zeros to decimal
        GetDspValue = GetDspValue * 2# + CDbl(Asc(Mid$(ValueAccum, Idx, 1)) - 48)
      Next Idx
      
    Case TypOct                         'OCTAL
      GetDspValue = CDbl("&O" & ValueAccum)
    
    Case TypHex                         'HEXADECIMAL
      GetDspValue = CDbl("&H" & ValueAccum)
  
    Case TypDec                         'DECIMAL
      If Left$(ValueAccum, 1) = "." Then ValueAccum = "0" & ValueAccum
      GetDspValue = CDbl(ValueSgn & ValueAccum) 'apply sign and convert
  End Select
End Function

'*******************************************************************************
' Subroutine Name   : NewLine
' Purpose           : Add a line to the display, and show Immediate value
'*******************************************************************************
Public Sub NewLine()
  With frmVisualCalc.lstDisplay
    If .ListIndex = .ListCount - 1 Then       'if we are at the end of the list
      .AddItem vbNullString                   'force a new line force a new line
      Call SelectOnly(.ListCount - 1)         'select that line
    Else
      Call SelectOnly(.ListIndex + 1)         'else select next line
    End If
  End With
  DoEvents
  Call DisplayLine
End Sub

'*******************************************************************************
' Subroutine Name   : PanelsNVN
' Purpose           : Set visibility of panels
'*******************************************************************************
Public Sub PanelsVNV(Vis As Boolean)
  With frmVisualCalc.sbrImmediate
    If .Panels("InstrCnt").Visible <> Vis Then
      .Panels("InstrCnt").Visible = Vis
      .Panels("InstrPtr").Visible = Vis
      .Panels("Pgm").Visible = Vis
      .Panels("MDL").Visible = Vis
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : UpdateStatus
' Purpose           : Update Status panels
'*******************************************************************************
Public Sub UpdateStatus()
  Dim S As String
  Dim i As Integer
  
  With frmVisualCalc
    With .sbrImmediate
      If RunMode Then                                                   'Run Mode?
        Call PanelsVNV(False)                                           'yes, so hide panels
      Else
        Call PanelsVNV(True)                                            'else show panels
        frmVisualCalc.mnuWINListMDL.Enabled = CBool(ModName)
        '
        ' module number
        '
        If CBool(ModName) Then                                          'module loaded?
          S = "MDL" & Format(ModName, "0000")                           'set up display
        Else
          S = "No MDL"                                                  'else no module
        End If
        If .Panels("MDL").Text <> S Then .Panels("MDL").Text = S        'show active Module
        '
        ' program number
        '
        If Not CBool(ActivePgm) And CBool(PgmName) Then                 'if pgm loaded from file
          S = "Loaded Pgm" & Format(PgmName, "00")                      'show active program
        ElseIf CBool(ActivePgm) Then                                    'if module program active
          S = "Pgm " & Format(ActivePgm, "00")
        ElseIf CBool(InstrCnt) Then                                     'if user program present
          S = "Pgm 00"
        Else
          S = "No Pgm"                                                  'no program
        End If
        If .Panels("Pgm").Text <> S Then .Panels("Pgm").Text = S        'show active program
        '
        ' program steps
        '
        S = "Steps: " & CStr(InstrCnt)
        If .Panels("InstrCnt").Text <> S Then .Panels("InstrCnt").Text = S 'show instruction counter
        '
        ' program step #
        '
        S = "Step #: " & CStr(InstrPtr)
        If .Panels("InstrPtr").Text <> S Then .Panels("InstrPtr").Text = S 'show instruction pointer
        '
        ' format style
        '
        S = "Style: " & CStr(LRNstyle)
        If .Panels("Style").Text <> S Then .Panels("Style").Text = S    'show Style
        '
        ' trace mode
        '
        S = "Trace: " & Format(TraceFlag, "On/Off")
        If .Panels("Tron").Text <> S Then .Panels("Tron").Text = S      'show Trace Mode
      End If
    End With
    '
    'figure number base
    '
    Select Case BaseType
      Case TypHex
        S = "Hex"
      Case TypDec
        S = "Dec"
      Case TypOct
        S = "Oct"
      Case TypBin
        S = "Bin"
    End Select
    If .lblHDOB.Caption <> S Then .lblHDOB.Caption = S
    '
    'figure angle mode
    '
    Select Case AngleType
      Case TypDeg
        S = "Deg"
      Case TypRad
        S = "Rad"
      Case TypGrad
        S = "Grd"
      Case Else
        S = "Mil"
    End Select
    If .lblDRGM.Caption <> S Then .lblDRGM.Caption = S
    '
    ' figure Notation format
    '
    i = CInt(EEMode) And 1
    If .lblEE.BackStyle <> i Then .lblEE.BackStyle = i
    If EngMode Then
      S = "Eng"
    Else
      S = "EE"
    End If
    If .lblEE.Caption <> S Then .lblEE.Caption = S
    '
    ' highlight backgrounds as needed
    '
    i = CInt(LrnMode) And 1
    If .lblLRN.BackStyle <> i Then .lblLRN.BackStyle = i
    i = CInt(INSmode) And 1
    If .lblINS.BackStyle <> i Then .lblINS.BackStyle = i
    i = CInt(TextEntry) And 1
    If .lblTxt.BackStyle <> i Then .lblTxt.BackStyle = i
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : CheckValue
' Purpose           : Add a numeric to the entry
'*******************************************************************************
Public Sub CheckValue(Digit As String)
  Dim S As String
  Dim Idx As Integer
  
  If AllowExp Then
    With frmVisualCalc.lstDisplay
      ValueAccum = Trim$(.List(.ListIndex))           'rotate characters after "E+" or "E-"
      If Left$(ValueAccum, 1) = "-" Then ValueAccum = Mid$(ValueAccum, 2)
      ValueAccum = Format(CDbl(Left$(ValueAccum, Len(ValueAccum) - 2) & _
                          Right$(ValueAccum, 1) & Digit), ScientifEE)
    End With
  ElseIf CBool(CharCount) And CBool(CharLimit) Then   'if processing values for pending ops...
    ValueAccum = ValueAccum & Digit                   'apply digit
    ValueTyped = True                                 'something was typed
    S = vbNullString                                  'init ValueAccum prepend data
    If CBool(PndIdx) Then                             'command pending?
      For Idx = 1 To PndIdx                           'yes, scan all
      S = S & GetInst(PndStk(Idx)) & " "              'grab command and accumulate
      Next Idx
    End If
    SetTip S & ValueAccum                             'set dat to tip field in status bar
  Else
    If Digit = "0" And _
       (ValueAccum = "0" Or ValueAccum = "0.") And _
       ValueDec = False Then Exit Sub                 'ignore prepending zeros
    ValueAccum = ValueAccum & Digit                   'add a char (or at least 1 zero)
    ValueTyped = True                                 'something was typed
    If CBool(Len(ValueAccum)) And _
       GetDspValue() <> 0# And _
       Left$(ValueAccum, 1) = "0" And _
       ValueDec = False Then
      ValueAccum = Mid$(ValueAccum, 2)
    End If
  End If
  Call DisplayAccum                                   'display final result
End Sub

'*******************************************************************************
' Subroutine Name   : ResetValueAccum
' Purpose           : Reset accumulator
'*******************************************************************************
Public Sub ResetValueAccum()
  ValueAccum = vbNullString       'clear accumulator
  ValueSgn = vbNullString
  ValueDec = False
End Sub

'*******************************************************************************
' Subroutine Name   : ResetAccumulator
' Purpose           : Reset accumulator. If do decimal, add it
'*******************************************************************************
Public Sub ResetAccumulator()
  Dim S As String
  
  Call ResetValueAccum            'reset accumulator data
  If LrnMode Or RunMode Or CBool(MRunMode) Then Exit Sub
  If BaseType = TypDec And Not TextEntry Then
    With frmVisualCalc.lstDisplay
      If DisplayHasText Then
        S = CStr(DisplayReg)
        DisplayHasText = False
      Else
        S = LTrim$(.List(.ListIndex)) 'get data there
        If Len(S) = 0 Then S = "0"
      End If
'''      If IsNumeric(S) Then
'''        If InStr(1, S, ".") = 0 Then  'decimal in display?
'''          S = S & "."
'''          .List(.ListIndex) = String$(DisplayWidth - Len(S), 32) & S 'display on current line
'''        End If
'''      End If
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Reset_Support
' Purpose           : Support Reset command
'*******************************************************************************
Public Sub Reset_Support()
  If CBool(InstrErr) Then 'set instruction pointer to 0 or error line
    InstrPtr = InstrErr - 1
  Else
    InstrPtr = 0
  End If
  InstrErr = 0            'reset instruction error
  SbrInvkIdx = 0          'clear subroutine stack
  BraceIdx = 0            'and brace stack
  Call ResetPndAll        'and oending operation index
  StopMode = False        'disable Stop mode
  INSmode = False         'disable insert mode
  ErrorFlag = False       'turn off error flags
  flags(7) = False
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : CMs_Support
' Purpose           : Clear and rest variables
'*******************************************************************************
Public Sub CMs_Support()
  Dim Idx As Long
  
  If LrnMode Then Exit Sub              'do nothing if LRN mode active
'
' reset variables 00-MaxVar
'
  For Idx = 0 To MaxVar + 1             'allow 1 more for Structure I/O
    With Variables(Idx)
      .VarType = vNumber                'default type is numeric
      .VdataLen = 0                     'no fixed length string
      .VName = Space(CLng(LabelWidth))  'no user-defined variable name
      Set .Vdata = Nothing              'clear any defined classes assigned here
      Set .Vdata = New clsVarSto        'init new variable storage
      .Vdata.VarRoot = Idx              'give it a root variable number reference
      .VuDef = False                    'not yet user-defined
      .Vaddr = -1                       'default definition address (system defined)
    End With
  Next Idx
'
' reset statistical registers
'
  Call Op13
'
' reset default date and time format
'
  DateFmt = DefDtFmt
  TimeFmt = DefTmFmt
'
' reset current pending variable
'
  CurrentVar = -1                       'default current register is not defined
  DspTxt = vbNullString
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : CE_Support
' Purpose           : Reset display register
'*******************************************************************************
Public Sub CE_Support()
  DisplayText = False
  DspTxt = vbNullString
  If IsNumbers Then
    DisplayReg = 0#                 'set immediate value to 0.
    IsNumbers = False               'turn off numeric entry
  End If
  
  flags(7) = False                  'turn off error flag indicator
  ErrorPause = False                'ensure error pause off
  ErrorFlag = False                 'turn off error flag
  With frmVisualCalc
    .txtError(0).Visible = False    'hide error report fields
    .txtError(1).Visible = False
    If .rtbInfo.Visible Then
      .rtbInfo.Visible = False
      .rtbInfo.Text = vbNullString
      .cmdUp.Enabled = True
      .cmdDn.Enabled = True
      .cmdPgUp.Enabled = True
      .cmdBackspace.Enabled = True
      .cmdPgDn.Enabled = True
      .cmdTop.Enabled = True
      .cmdBtm.Enabled = True
      .mnuFile.Enabled = True
      .mnuWindow.Enabled = True
      .mnuHelp.Enabled = True
    End If
  End With
  
  CharCount = 0                     'reset character count
  CharLimit = 0
  Call ResetAccumulator             'null accumulator
  AllowExp = False                  'disable allowing exponent entry
  REMmode = 0                       'ensure remarks mode is off
  If Not DspLocked Then
    Call DisplayLine                'display null value if we were entering data
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : CLR_Support
' Purpose           : Clear pending operations
'*******************************************************************************
Public Sub CLR_Support()
  
  If LrnMode Then Exit Sub        'do nothing if LRN mode active
  
  SbrInvkIdx = 0                  'no pending call; ' Reset subroutine call stack
  BraceIdx = 0                    'no pending brace; Reset brace level stack
  PushIdx = 0                     'remove items from the Push Stack
  DisplayReg = 0#                 'immediate display value is 0; reset display and temp registers
  ErrorFlag = False               'reset error flags
  flags(7) = False
  LastTypedInstr = 128            'force alphanumeric assumption
  PmtFlag = False                 'turn off prompt flag
  REMmode = 0                     'turn off Rem Mode
  PendIdx = 0                     'reset pending operations
  Call ResetPndAll
  EEMode = False                  'reset the EE mode
  EngMode = False                 'reset the Engineering mode
'
' clear display and display reset line
'
  frmVisualCalc.PicPlot.Visible = False
  PlotTrigger = False
  Call Clear_Screen               'reset display screen
  Call CE_Support
  Call UpdateStatus
End Sub

'*******************************************************************************
' Subroutine Name   : Cmn_CP_Support
' Purpose           : Routines also common to Preprocess
'*******************************************************************************
Public Sub Cmn_CP_Support()
  Call ResetFmt                       'reset formatted text buffer
  Call Reset_Support                  'reset a bunch of things
  Call Clear_Screen                   'reset display screen
  Call CMs_Support                    'reset variables (calls Op13, CE_Support)
  Call Op23                           'reset ALL flags
  Call Op00                           'clear all text fields
  
  DisplayReg = 0#                     'immediate display value is 0.0
  TestReg = 0#                        'also test reg
  PushIdx = 0                         'reset Push/Pop stack
  DspLocked = False                   'turn off display locked flag
  Call DspBackground                  'update background of display list
  InstrErr = 0                        'clear last encountered error location
  StopMode = False                    'turn off stop mode
  AllowExp = False                    'turn off EE expoent entry mode
  DefDef = 0                          'turn off DEF processing
  DefTrueSz = 8                       'reinit DefTrue pool
  ReDim DefTrue(8)
  ErrorFlag = False                   'reset error flags
'
' Init structure Pool
'
  Erase StructPl                      'init structure pool
  StructCnt = 0                       'no items in the pool
'
' init brace level
'
  BraceSize = BraceDepth              'init depth (increase by BraceDepth also)
  ReDim BracePool(BraceDepth)
  BraceIdx = 0
'
' user-defined label storage
'
  Call RenewLabels
'
' Pending operation stack
'
  PendSize = SizePend                 'init depth
  ReDim PendValue(SizePend)           'pending value pool
  ReDim PendOpn(SizePend)             'pending operation pool
  ReDim PendHir(SizePend)             'priority level
  PendIdx = 0                         'no pending operations
  Call ResetPndAll
'
' Push stack
'
  PushSize = PushLimit                'initial size
  ReDim PushValues(PushLimit)
  PushIdx = 0                         'no pushes
'
' DEF pool
'
  DefSize = SizeDef                   'initial depth
  ReDim DefName(SizeDef)              'defined names
  DefCnt = 0                          'no definitions defined
'
' Program Invoke Pool
'
  SbrInvkSize = 16                    'initial depth
  ReDim SbrInvkStk(16)                'stack data
  SbrInvkIdx = 0                      'Index into pool
'
' set up keyboard display
'
  DisplayHasText = False                    'no text in display
  AllowSpace = False                        'do not allow spacebar
  Call frmVisualCalc.checkTextEntry(False)  'turn off text entry mode
  Call CLR_Support
'
' reset Plot data
'
  PlotTrigger = False
  Call PlotClr
  Call ResetListSupport
  Randomize
End Sub

'*******************************************************************************
' Subroutine Name   : CP_Support
' Purpose           : Remove program code (this incorporates CLR, CMs, and CE)
'*******************************************************************************
Public Sub CP_Support()
  Dim Idx As Integer

  If LrnMode Or RunMode Then Exit Sub 'do nothing if LRN mode active
  
  PgmName = 0                         'reset loaded program name
  Preprocessd = False                 'remove any compilation flags
  Compressd = False
'
' set numerous defaults
'
  InstrSize = SizeInst                'set number of instructions initially allowed (expandable +100 inc)
  ReDim Instructions(SizeInst)        'set aside space
  InstrPtr = 0                        'buttom of list
  InstrCnt = 0                        'no instructions
  IsDirty = False                     'no instructions have been learned by the calculator
  
  Call Cmn_CP_Support
  frmVisualCalc.mnuWinASCII.Enabled = False  'view ASCII
End Sub

'*******************************************************************************
' Subroutine Name   : Del_Support
' Purpose           : Support Delete File
'*******************************************************************************
Public Sub Del_Support()
  Dim S As String, Path As String
  Dim Idx As Integer
  
  With frmVisualCalc.lstDisplay
    S = Trim$(.List(.ListIndex))
    If Right$(S, 1) = ">" Then                            'strip any titled added
      Idx = InStr(1, S, "<")
      If CBool(Idx) Then S = RTrim$(Left$(S, Idx - 1))
    End If
    Select Case UCase$(Right$(S, 4))
      Case ".TXT", ".PGM", ".MDL"
        For Idx = .ListIndex - 1 To 0 Step -1             'find folder, if so
          Path = .List(Idx)
          If Left$(Path, 1) <> " " Then Exit For          'if no space at head, then folder
        Next Idx
        If Left$(Path, 1) = " " Then Exit Sub             'did not find anything
        Path = StorePath & "\" & RTrim$(Path) & "\" & S   'build path to text file
        If Fso.FileExists(Path) Then
          Call CmdNotActive
          If CenterMsgBoxOnForm(frmVisualCalc, "Delete selected file: " & S & "?", _
             vbOKCancel Or vbQuestion Or vbDefaultButton2, "Confirm Delete") = vbOK Then
            On Error Resume Next
            Call Fso.DeleteFile(Path)
            If Not CBool(Err.Number) Then
              Idx = .ListIndex
              .RemoveItem Idx
              If Idx = .ListCount Then Idx = Idx - 1
              If Idx < 0 Then Idx = 0
              SelectOnly Idx
            End If
            On Error GoTo 0
          End If
        End If
      Case Else
        CmdNotActive
    End Select
  End With
  DoEvents
End Sub

'*******************************************************************************
' Subroutine Name   : RenewLabels
' Purpose           : Used to restart label definitions. A-Z (1-26) are always first
'*******************************************************************************
Public Sub RenewLabels()
  Dim Idx As Long
  
  LblSize = LblDepth                      'initial depth (Increases by 10)
  ReDim Lbls(LblDepth)                    'structure array storage
'
'init special spacebar button (actually, this uses User-key space, but its beyond user control
'
  With Lbls(Idx)
    .lblName = "Space"                    'Set spacebar
    .lblCmt = "Space Bar"
    .LblTyp = TypKey                      'user-defined
    .LblScope = Pub                       'initially Public Scope
  End With
'
' set up user-defined keys
'
  For Idx = 1 To 26                       'A to Z
    Hidden(Idx) = False                   'unhide it
    With Lbls(Idx)
      .lblName = Chr$(64 + Idx)           'set initial name of A-Z
      .lblCmt = "User-Defined Key '" & Chr$(64 + Idx) & "'" 'comment
      .LblTyp = TypKey                    'user-defined
      .LblScope = Pub                     'initialy Public Scope
      .lblUdef = False                    'not yet user-defined
    End With
  Next Idx
  LblCnt = 27                             '27 user-definable items defined (Spacebar is special)
End Sub

'*******************************************************************************
' Function Name     : FindLbl
' Purpose           : Find a label of the defined type
'*******************************************************************************
Public Function FindLbl(Txt As String, Typ As LblTypes) As Long
  Dim Idx As Long
  Dim Test As String * LabelWidth
  
  Test = Txt                          'make sure we are upper case
'-----------------------------------------------
  If CBool(ActivePgm) Then           'if module program
    For Idx = ModLblMap(ActivePgm - 1) + 1 To ModLblMap(ActivePgm) - 1
      With ModLbls(Idx)
        If .LblTyp = Typ Then
          If StrComp(.lblName, Test, vbTextCompare) = 0 Then
            FindLbl = Idx             'return the index
            Exit Function
          End If
        End If
      End With
    Next Idx
'-----------------------------------------------
  Else                                'else scan through user pgm
    For Idx = 1 To LblCnt
      With Lbls(Idx)
        If .LblTyp = Typ Then
          If StrComp(.lblName, Test, vbTextCompare) = 0 Then
            FindLbl = Idx             'return the index
            Exit Function
          End If
        End If
      End With
    Next Idx
  End If
'-----------------------------------------------
  FindLbl = 0                         'else indicate it was not found
End Function

'*******************************************************************************
' Function Name     : GetDisplayText
' Purpose           : Return the contents of the display in a string
'*******************************************************************************
Public Function GetDisplayText() As String
  Dim Idx As Integer
  Dim S As String
  
  S = vbNullString                  'init result
  With frmVisualCalc.lstDisplay
    If DspLocked Then               'if locked display, we can trim it up
      For Idx = 0 To .ListCount - 1 'append all lines
        S = S & RTrim$(.List(Idx)) & vbCrLf 'grab a line
      Next Idx                      'do all of them
    Else
      For Idx = 0 To .ListCount - 1 'append all lines
        S = S & .List(Idx) & vbCrLf 'grab a line
      Next Idx                      'do all of them
    End If
  End With
  GetDisplayText = S                'return result
End Function

'*******************************************************************************
' Function Name     : DisplayMsg
' Purpose           : Display a Message on the console
'*******************************************************************************
Public Function DisplayMsg(Msg As String)
  Dim S As String
  Dim Dt As Boolean
  
  S = DspTxt                            'save display text data
  Dt = DisplayText                      'save DisplayText state
  With frmVisualCalc.lstDisplay         'report we loaded OK
    Call NewLine                        'bump display line
    .List(.ListIndex - 1) = Msg         'display message
  End With
  DisplayText = Dt                      'recover state and data
  DspTxt = S
End Function

'*******************************************************************************
' Subroutine Name   : CmdNotActive
' Purpose           : Issue s system beep if a command is not supported in
'                   : the current mode
'*******************************************************************************
Public Sub CmdNotActive()
  MsgBeep beepSystemAsterisk
End Sub

'*******************************************************************************
' Function Name     : GetInstrStr
' Purpose           : Get instruction text, surrounded by spaces
'*******************************************************************************
Public Function GetInstrStr(ByVal Key) As String
  GetInstrStr = " " & GetInst(Key) & " "
End Function

'*******************************************************************************
' Subroutine Name   : DspBackground
' Purpose           : Adjust color of background, depending on state
'*******************************************************************************
Public Sub DspBackground()
  Dim FC As Long
  With frmVisualCalc
    .imgLocked(0).Visible = DspLocked And Not LrnMode  'visibility of Lock icon
    .imgLocked(1).Visible = Not .imgLocked(0).Visible
    If .mnuWinGreenScreen.Checked Then
      FC = vbGreen
    Else
      FC = vbBlack
    End If
    With .lstDisplay
      .ForeColor = FC
      If LrnMode Then
        .BackColor = RGB(216, 228, 248)
        .ForeColor = vbBlack
      ElseIf Not DspLocked Then                     'if LRN mode or not locked...
        .BackColor = BackClr                        'display normal background
      ElseIf DspLocked Then                         'locked?
        If BackClr = vbBlack Then                   'if greenscreen
          .BackColor = &H404040                     'dark background
        Else
          .BackColor = &HE0E0E0                     'light background
        End If
      Else
        .BackColor = BackClr                        'not locked and not Lrn mode
      End If
    End With
  End With
End Sub

'*******************************************************************************
' Function Name     : ShellPath
' Purpose           : Shell out to OS to executea program
'*******************************************************************************
Public Function ShellPath(hWnd As Long, Cmd As String, Path As String) As Long
  Dim i As Long
  
  i = ShellExecute(hWnd, Cmd, Path, "", "", SW_NORMAL)
  If i = 0 Then i = 1
  If i > 32 Then i = 0
  ShellPath = i
End Function

'*******************************************************************************
' Subroutine Name   : RepointIdx
' Purpose           : Repoint select line is program display
'*******************************************************************************
Public Sub RepointIdx()
  Dim Idx As Integer, i As Integer
  Dim Map() As Integer, CntMax As Integer
  
  Select Case LRNstyle
    Case 0          'style 0 (Raw)
      i = InstrPtr
    Case 3          'Command groups format (debug style)
      Map = InstMap3
      CntMax = InstCnt3
    Case Else       'full formatted
      Map = InstMap
      CntMax = InstCnt
  End Select
  
  If CBool(LRNstyle) Then
    i = 0                               'init select line
    For Idx = 0 To CntMax - 1
      If Map(Idx) = InstrPtr Then       'map same as instruction pointer
        i = Idx
        Exit For
      End If
      If Map(Idx) > InstrPtr Then       'if we went past current point...
        If Idx = 0 Then Exit For        'cannot back up, so i=0
        i = Idx - 1                     'else back up one line
        Exit For
      End If
    Next Idx
  End If
  Idx = i - DisplayHeight / 2         'compute top index
  If Idx < 0 Then Idx = 0             'adjust for rollover
  SelectOnly i                        'select active line
  frmVisualCalc.lstDisplay.TopIndex = Idx 'set top index to center selection
End Sub

'*******************************************************************************
' Subroutine Name   : ListVdata
' Purpose           : List contents of all variables
'*******************************************************************************
Public Sub ListVdata(ByVal All As Boolean)
  Dim Idx As Long, X As Long, Y As Long, cnt As Long, Sz As Long
  Dim S As String, SS As String, Ary() As String
  Dim Vptr As clsVarSto, Xptr As clsVarSto, Yptr As clsVarSto
  Dim HoldDT As Boolean
  
  HoldDT = DisplayText                                  'hold current flag state
'
' disable any locking
'
  If DspLocked Then
    DspLocked = False
    Call DspBackground
  End If
  
  Call Clear_Screen                                     'clean up display
  frmVisualCalc.lstDisplay.List(0) = "Building list..."
  Screen.MousePointer = vbHourglass
  DoEvents
  LockControlRepaint frmVisualCalc.lstDisplay
'  Call Clear_Screen                                     'clean up display
  
  Sz = 128
  cnt = -1
  ReDim Ary(128)
  
  For Idx = 0 To 99                                     'process all 100 variables
    With Variables(Idx)
      Set Vptr = .Vdata                                 'get pointer to object
      SS = ListVdataB(Vptr, All)                        'display its contents
      If CBool(Len(SS)) Then                            'if something to display...
        cnt = cnt + 1                                   'bump index
        If cnt > Sz Then                                'adjust array size as needed
          Sz = Sz + 128
          ReDim Preserve Ary(Sz)
        End If
        Ary(cnt) = SS                                   'stuff displayable data
      End If
      If Not Vptr.LnkNext Is Nothing Then               'has an X dim?
        For X = 0 To Vptr.GetMaxDim                     'yes, so process X dim as well
          S = "  [" & CStr(X) & "]"                     'build references
          Set Xptr = Vptr.PntToLnk(X)                   'point to each X item
          If Not Xptr.LnkChild Is Nothing Then          'does it have a Y dim?
            Set Yptr = Xptr.LnkChild                    'yes, get pointer to it
            For Y = 0 To Yptr.GetMaxDim                 'process Y Dim
              SS = S & "[" & CStr(Y) & "]"              'build reference
              SS = ListVdataB(Yptr.PntToLnk(Y), All, SS) 'report contents of [X][Y]
             If CBool(Len(SS)) Then                    'if something to display...
                cnt = cnt + 1                          'bump index
                If cnt > Sz Then                       'adjust array size as needed
                  Sz = Sz + 128
                  ReDim Preserve Ary(Sz)
                End If
                Ary(cnt) = SS                          'stuff displayable data
              End If
            Next Y
          Else
            SS = ListVdataB(Xptr, All, S)               'else report contents of [X]
            If CBool(Len(SS)) Then                      'if something to display...
               cnt = cnt + 1                            'bump index
               If cnt > Sz Then                         'adjust array size as needed
                 Sz = Sz + 128
                ReDim Preserve Ary(Sz)
               End If
               Ary(cnt) = SS                            'stuff displayable data
             End If
          End If
        Next X
      End If
    End With
  Next Idx                                              'do all variables
  
  If cnt = -1 Then                                      'if nothjing to display...
    DisplayMsg "No variable data to display"            'report it
    CmdNotActive                                        'and sound off
  Else
    With frmVisualCalc.lstDisplay
      .Clear                                            'else clear display
      For Idx = 0 To cnt
        .AddItem Ary(Idx)                               'add all items in list
      Next Idx
      .AddItem vbNullString                             'add a blank line to end
      SelectOnly .ListCount - 1                         'select it
    End With
  End If
  UnlockControlRepaint frmVisualCalc.lstDisplay         'unlock the control
  Screen.MousePointer = vbNormal                        'reset pointer
  DisplayText = HoldDT                                  'reset display state
  DisplayLine                                           'display command line
End Sub

'*******************************************************************************
' Subroutine Name   : ListVdataB
' Purpose           : Process variable object contents
'*******************************************************************************
Private Function ListVdataB(Vptr As clsVarSto, ByVal All As Boolean, Optional AryData As String = vbNullString) As String
  Dim Idx As Long
  Dim S As String, Nm As String
  Dim Skip As Boolean
  
  With Variables(Vptr.VarRoot)
    S = CStr(ExtractValue(Vptr))                    'get object contents as string
    Skip = False                                    'turn off skip flag
    If Not All Then                                 'much check for nulls of not ALL
      If .VarType = vString Then                    'text?
        If CBool(Len(Trim$(S))) Then                'does it have actual data?
          S = """" & S & """"                       'yes, so encote it
        Else
          Skip = True                               'else we will skip it
        End If
      Else
        Skip = S = "0"                              'skip it if contents are "0"
      End If
    ElseIf .VarType = vString Then                  'will display all, so is data text?
      S = """" & S & """"                           'yes, so enquote it
    End If
    
    If Skip Then                                    'if we will skip display
      ListVdataB = vbNullString
    Else                                            'else we will display
      If CBool(Len(AryData)) Then                   'then if array references present...
        S = AryData & ": " & S                      'build message with references
      Else                                          'otherwise, simply report variable and contents
        Nm = Trim$(Variables(Vptr.VarRoot).VName)   'check for user-defined name
        If CBool(Len(Nm)) Then                      'has name?
          S = "Var '" & Nm & "' (" & Format(Vptr.VarRoot, "00") & "): " & S 'yes
        Else
          S = "Var (" & Format(Vptr.VarRoot, "00") & "): " & S              'else just number
        End If
      End If
      ListVdataB = S                                'report object data to Display List
    End If
  End With
End Function

'*******************************************************************************
' Subroutine Name   : PlayClick
' Purpose           : Play Resource Click
'*******************************************************************************
Public Sub PlayClick()
  PlayWavResource "Click", , , True
End Sub

'*******************************************************************************
' Subroutine Name   : BreakUpVal
' Purpose           : Add color value to Learn Mode Instruction list
'*******************************************************************************
Public Sub BreakUpVal(ByVal Value As Long)
  Dim S As String
  Dim Idx As Integer
  
  S = CStr(Value)                                   'text to process
  For Idx = 1 To Len(S)                             'handle character at a time
    Call AddInstruction(Asc(Mid$(S, Idx, 1)) - 48)  'convert "0"-"9" to 0-9
  Next Idx
End Sub

'*******************************************************************************
' Function Name     : AddTitle
' Purpose           : Apply app title to program
'*******************************************************************************
Public Function AddTitle(Text As String) As String
  Dim i As Integer
  
  i = InStr(1, Text, "MyApp")
  AddTitle = Left$(Text, i - 1) & App.Title & Mid$(Text, i + 5)
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

