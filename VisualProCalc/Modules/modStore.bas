Attribute VB_Name = "modStore"
Option Explicit
'
'------------------------------------------------------
' General Data Storage
'------------------------------------------------------
Public Fso As FileSystemObject    'file I/O support
'-----
Public colHelpBack As Collection  'collection for going backward in help
Public colFindList As Collection  'collection of search list items
Public FindListIdx As Integer     'where are we in colFindList
Public FindListLen As Integer     'length of selection text
'-----
Public NotePadPath As String      'path to Notepad.exe application
Public WordPadPath As String      'path to Wordpad.exe application
'-----
Public vPi As Double              'store value of Pi (Atn(1)*4)
Public vPi2 As Single             'store value of Pi * 2 (Atn(1)*8)
Public vRA As Double              'store 90-degree right angle (Atn(1)*2)
Public vE As Double               'store value of e  (exp(1))
'-----
Public LineHeight As Single       'height of a line in pixels
Public BackClr As Long            'store Display List background color
Public Hidden(26) As Boolean      'true if any Alpha keys are hidden
Public Variables(MaxVar + 1) As Variable 'Variables 0-MaxVar (allow extra for Structure ops)
Public flags(9) As Boolean        'flags 0-9
'------
Public StoreList As Boolean       'true when directory storage list has been displayed
Public ModuleList As Boolean      'true when Module Directory list has been displayed
Public TypeMatic As Boolean       'True if Typmatic enabled
Public TypeMatTxt As String       'type ahead storage
'-----
Public frmHelpLoaded As Boolean   'True if frmHelp is loaded
Public frmCDLoaded As Boolean     'True if FrmCoDisplay is loaded
Public frmSrchLoaded As Boolean   'True when frmSearch is loaded
'-----
Public IsDirty As Boolean         'TRUE if any code has been LEARNED and not saved
Public Cancel As Boolean          'general cancel flag
Public DspLocked As Boolean       'If true, user cannot type data
Public DspPgmList As Boolean      'if Pgm List displayed in locked window
Public HaveVHelp As Boolean       'True if we have a VCHelp file
Public SkipChg As Boolean         'True if we should skip selection processing in Help File
'-----
Public IgnoreClick As Boolean     'used by listbox resets during LRN Mode
Public PmtFlag As Boolean         'true if we will stop for text input
'-----
Public AngleType As AngleTypes    'Current Angle Type
Public BaseType As BaseTypes      'Current Numeric Base
'-----
Public StorePath As String        'current Data Storage path
Public DateFmt As String          'storage for date format string
Public TimeFmt As String          'storage for time format string
Public DspFmt As String           'default numeric display (EE, Fixed, etc)
Public DspFmtFix As Integer       'Fix format # decimal places (-1 is none)
Public ScientifEE As String       'string storing current scientific notation mode
Public AllowExp As Boolean        'true if Exponent allowed to be entered
'-----
Public ErrorFlag As Boolean       'Error Condition exists
Public ErrorPause As Boolean      'force pause when error encountered
Public tmrWait As Integer         'timer 1/2 second pause value (0-3 = 1/2 sec to 2 seconds)
'-----
Public TestReg As Double          '(T) register (Test)
Public DisplayReg As Double       '(X) register (Display)
Public TxtPltChr As String        'Text Plot character
Public DspTxt As String           'store string typed
Public Upcase As Boolean          'True if Uppercase only (used by Ukey to get Uppercase key letter)
Public HaveTxt As Boolean         'true if text data, as opposed to digits, have been typed
Public DisplayText As Boolean     'True if Text should be displayed from DspTxt, rather than DisplayReg
Public AllowSpace As Boolean      'True if the user can type a space
Public DisplayHasText As Boolean  'True if display contains text data
'-----
Public ValueAccum As String       'value accumulator
Public ValueDec As Boolean        'true if decimal used
Public ValueSgn As String         'blank or [-]
Public ValueTyped As Boolean      'True of something was typed
'-----
Public CurrentVar As Integer      'Current root variable number being used (00-99)
Public CurrentVarTyp As Vtypes    'Current variable type (Number, Text, Integer, Char)
Public CurrentVarObj As clsVarSto 'pointer to actual variable object
'---------------------------------
' Plot support
'---------------------------------
Public PlotX As Long              'Active Plot X position
Public PlotY As Long              'Active Plot Y position
Public PlotXDef As Long           'Default X position (for Line returns)
Public PlotYDef As Long           'Default Y position (For line increments)
Public PlotColor As Long          'drawing color
Public LastDir As Double          'text direction 0-7 (0,45,90,135,189,225,270,315)
Public PlotTrigger As Boolean     'True if we want events to occurr
Public PlotTriggerSbr As Integer  'program step of subroutine definition for plot triggering
Public LastPlotX As Long          'last X cursor position over Plot
Public LastPlotY As Long          'last Y cursor position over Plot
'---------------------------------
' File I/O Pool
'---------------------------------
Public Files(9) As FileBufs       'File I/O buffers
Public Tstrm(9) As TextStream     'text streams for non-block (binary) data
'---------------------------------
' Program Processing/Compression flags
'---------------------------------
Public AutoPprc As Boolean        'true when we should try to automatically Preprocess program loads
Public Preprocessd As Boolean     'true if program pre-processed OK
Public Preprocessing As Boolean   'True if preprocessing
Public Compressd As Boolean       'true if fully Compressed
Public Compressing As Boolean     'True if Compressing
Public HaveMain As Long           'index to active Pgm's Main() routine, if it exists
'---------------------------------
' Modes
'---------------------------------
Public LrnMode As Boolean         'LEARN mode active if TRUE
Public RunMode As Boolean         'If program code is running (some commands act differently)
Public MRunMode As Integer        'Non-zero if Module PGM invoked
Public HMRunMode As Integer       'hold module run mode for Pmt command
Public ModPrep As Integer         '<>0 if preparing to invoke Module program (Pgm command invoked)
Public SSTmode As Boolean         'True if Single Step mode was invoked
Public StopMode As Boolean        'Stop mode active, not just R/S, which can continue with R/S keypress
Public INSmode As Boolean         'TRUE if in INSERT mode
Public REMmode As Integer         'typing REM, 0=off, 1=[Rem ], 2=['], 2=Fmt data
Public LRNstyle As Integer        'LRN mode display format (0=basic, 1=formatted, 2=formatted w/line nums)
Public EEMode As Boolean          'Scientific Notation Mode
Public EngMode As Boolean         'Engineering mode active
Public Tron As Boolean            'Trace On (use in Run-time)
Public TraceFlag As Boolean       'actual Trace flag
'---------------------------------
' LRN Mode Supprt
'---------------------------------
Public LrnDsp As String           'save contents of display to a string
Public LrnSav As Integer          'index to line shown
Public LrnTop As Integer          'top line in display
'---------------------------------
' keyboard operation modification
'---------------------------------
Public Key2nd As Boolean          'true when the 2nd key is down
Public KeyShift As Boolean        'True when Shify Key active
Public KeyShf As Boolean          'True when ACTUAL shift key is pressed
Public TextEntry As Boolean       'true when Text entry key is active and enabled
Public VarLbl As Boolean          'if text data can identify a variable # or name
Public RS_Pressed As Boolean      'flag indicating R/S was pressed during Run
Public Query_Pressed As Boolean   'The Query button was pressed
'---------------------------------
' PUSH Pool
'---------------------------------
Public PushSize As Integer        'size of pool
Public PushValues() As Double     'values pushed onto the stack
Public PushIdx As Integer         'number of items pushed to stack
'---------------------------------
' INSTRUCTION Pool
'---------------------------------
Public InstrSize As Integer       'size of pool
Public Instructions() As Integer  'instruction pool (Variable Size. Initial is 1024, Inc by 128)
Public InstrCnt As Integer        'numer of instructions defined
Public InstrPtr As Integer        'pointer into instruction pool
Public InstrErr As Integer        'index to last encountered error
Public LastTypedInstr As Integer  'last typed instruction
'---------------------------------
' Formatted Listing Pool
'---------------------------------
Public InstCnt As Integer         'count of formatted instruction lines
Public InstFmt() As String        'Formatted instruction list. Ie, "RCL IND TEST"
Public InstMap() As Integer       'map of formatted instruction lines, where each line begins
Public InstCnt3 As Integer        'Debug-style count of formatted instruction lines
Public InstFmt3() As String       'Debug-style Formatted instruction list. Ie, "RCL IND TEST"
Public InstMap3() As Integer      'Debug-style map of formatted instruction lines, where each line begins
'---------------------------------
' Lrn Mode Formatted pool
'---------------------------------
Public FmtSize As Integer
Public FmtLst() As String
Public FmtMap() As Integer
Public FmtIdx As Integer
Public FmtCnt As Integer
'---------------------------------
' LABEL Pool
'---------------------------------
Public LblSize As Long            'size of pool (init to 50, inc by 20)
Public Lbls() As Labels           'label, structure, constant, and subroutine names
Public LblCnt As Long             'number of acual labels defined
'---------------------------------
' STRUCT Pool
'---------------------------------
Public StructPl() As StructPool   'pool of structure definitions
Public StructCnt As Integer       'number of items in the pool
'---------------------------------
' DEF Pool
'---------------------------------
Public DefSize As Integer         'size of pool
Public DefName() As String        'def name entry (initial = SizeDef, inc by DefInc)
Public DefCnt As Integer          'number of definitions defined
'---------------------------------
' PENDING OPERATION Pool
'---------------------------------
Public PendSize As Integer        'size of pool
Public PendValue() As Double      'pending match ops (init to 20, inc by 10)
Public PendOpn() As Integer       'pending operation
Public PendHir() As Integer       'pending priority level
Public PendIdx As Integer         'pointer into data
'---------------------------------
' Subroutine Invoke Pool
'---------------------------------
Public SbrInvkSize As Integer     'size of pool
Public SbrInvkStk() As SbrInvk    'pool
Public SbrInvkIdx As Integer      'subroutine stack location
'---------------------------------
' PENDING FUNCTION Pool
'---------------------------------
Public PndStk(5) As Integer       'pending commands (normally, no more than 1 or 2)
Public PndIdx As Integer          'pointer into data
Public IsNumbers As Boolean       'True if numeric key pressed
Public CharCount As Integer       'alpha-numeric character count
Public CharLimit As Integer       'limits on entry counts (1 or 2 or 3)
Public PndImmed As Double         'storage for DisplayReg value when pending op encountered
Public PndPrev As Double          'storage for previous pending value (used by Swap)
'---------------------------------
' TEMP Pool for Cut/Copy/Paste
'---------------------------------
Public ListText() As String       'storage for Cut/Copy/Past
Public ListCnt As Long            'number of items in ListText
'---------------------------------
' BRACING Pool
'---------------------------------
Public BraceSize As Integer       'size of brace pool
Public BracePool() As TrkLoop     'tracking pool for braced data
Public BraceIdx As Integer        'brace level (set by Run)
'---------------------------------
' Compressr References
'---------------------------------
Public TxtData As String          'constant data
Public TstData As Double          'double conversion of TxtData
'
Public InstTxt As String          'instruction string
Public Code As Integer            'current start instruction code
Public PrvCode As Integer         'previous start instruction coode
Public HaveDels As Boolean        'if any lines were merged (Style 1)
'-----
Public SbrDefFlg As Boolean       'Subroutine being defined (Sbr or Ukey)
'-----
Public UtlDefFlg As Boolean       'Until being defined
'-----
Public ForDefflg As Boolean       'true if we are defining the start of a For loop
Public ForIdx As Integer          'Depth of For blocks
'-----
Public WhiDefFlg As Boolean       'true if we are defining an expression for While...Loop
Public DoWhiDefFlg As Boolean     'true if we are defining an expression for Do..While()
Public WhiIdx As Integer          'depth of While blocks
'-----
Public DoIdx As Integer           'depth of Do blocks
'-----
Public SelDefFlg As Boolean       'true if we are defining a Select Expression
Public SelIdx As Integer          'depth of Select Blocks
'-----
Public CaseDefFlg As Boolean      'true if we are defining a Case statement
Public CaseIdx As Integer         'Depth of Case blocks
'-----
Public IfDefFlg As Boolean        'true if we are defining an If statement
Public IfIdx As Integer           'depth of If blocks
'-----
Public StDefFlg As Boolean        'struct is being defined
'-----
Public EnDefFlg As Boolean        'Enum is being defined
Public EnumIdx As Long            'value of enum data
'-----
Public ConDefFlg As Boolean       'True if a  constant is being created
'-----
Public DefDef As Integer          'True if processing a Def statement (disabled by Edef)
Public DefTrueSz As Integer       'size of DefTrue pool
Public DefTrue() As Boolean       'True if check is valid (used also by Delse)
'-----
Public ForInit As Boolean         'True when we are initializing a For loop (for encountering ';')
Public DoDebug As Boolean         'True if we want to go into the Debug Mode
'---------------------------------
' MODULE Pools
'---------------------------------
Public ActivePgm As Integer       'active program number (non-zero indicates Module Pgm)
Public PgmName As Integer         'current active program number loaded (user pgrams; but treated like 0)
Public ModLcl() As Integer        'compilation pool for local program (Pgm 00)
'-----
Public ModName As Integer         'Current Module name (number) 1-9999
Public ModLbl As String * DisplayWidth 'Module text name
Public ModLocked As Boolean       'true if the module is locked (cannot Op 32 programs)
Public ModCnt As Integer          'number of module programs defined
Public ModDirty As Boolean        'True if Module has been updated
'-----
Public ModLblCnt As Integer       'number labels defined for ModLbls()
Public ModLbls() As Labels        'list of module program's lables
Public ModLblMap() As Long        'index into ModLbls for each program
'-----
Public ModStCnt As Integer       'number of items in the pool
Public ModStPl() As StructPool    'pool of module structure definitions
Public ModStMap() As Long         'map to each Module's individual ModStPl() array
'-----
Public ModSize As Long            'length of entire module's program data in integers
Public ModMem() As Integer        'variable length Module memory storage
Public ModMap() As Long           'list of offsets for the end+1 of each module program (beg. of next)
'------------------------------------------------------
' End of definitions
'------------------------------------------------------

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************
