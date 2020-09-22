Attribute VB_Name = "modKbdPending"
Option Explicit

Private Hyp As Boolean    'use by angle functions for Hyperbolic modifier
Private Arc As Boolean    'used by angle functions for Arc modifier
Private KeyBase As Long   'keybase register for crunching up to 3 commands into 1

Public Kyz(26) As Boolean 'local storage for Hkey and Skey processing

'*******************************************************************************
' Subroutine Name   : ResetPnd
' Purpose           : Reset pending command that can have data following
'*******************************************************************************
Public Sub ResetPnd()
  IsNumbers = False               'set up for alpha-numeric
  CharCount = 0
  CharLimit = 0
End Sub

'*******************************************************************************
' Subroutine Name   : ResetPndAll
' Purpose           : Reset pending command that can have data following
'                   : and reset PndIdx to 0
'*******************************************************************************
Public Sub ResetPndAll()
  Call ResetPnd
  PndIdx = 0
End Sub

'*******************************************************************************
' Subroutine Name   : CheckAngles
' Purpose           : Process angle commands
'*******************************************************************************
Private Sub CheckAngles()
  Hyp = False           'turn off modifier flags
  Arc = False
'
' init key base, in case no pending operations
'
  KeyBase = 128
'
' if 3 pending operations...
'
  If PndIdx = 3 Then
    Select Case PndStk(2)
      Case iHyp  'Hyp
        Hyp = True
      Case iArc  'Arc
        Arc = True
      Case Else
        Call ResetPnd
        Exit Sub
    End Select
  End If
'
' if more than 1 pending operation...
' (this will merge if 3 pending ops)
'
  If PndIdx > 1 Then
    Select Case PndStk(1)
      Case iHyp  'Hyp
        If Hyp Then
          ForcError "Illegal command usage"
          Exit Sub
        End If
        Hyp = True
      Case iArc  'Arc
        If Arc Then
          ForcError "Illegal command usage"
          Exit Sub
        End If
        Arc = True
      Case Else
          ForcError "Illegal command usage"
        Exit Sub
    End Select
  End If
'
' nothing more to do if Hyp or Arc not set
'
  KeyBase = PndStk(PndIdx)  'get Sin, Cos, Tan command
  If Not Hyp And Not Arc Then Exit Sub
'
' init offset bases to Arc-only
'
  If KeyBase < 158 Then     'Sin, Cos, Tan
    KeyBase = KeyBase + 245 ' - 155 + 400
  Else                      'Sec, Csc, Cot
    KeyBase = KeyBase + 126 ' - 283 + 409
  End If
'
' check for presence of Hyp and Arc (we know not that at lease one is set)
'
  If Hyp And Arc Then
    KeyBase = KeyBase + 6                   'set to Arc+Hyp base
  ElseIf Hyp Then
    KeyBase = KeyBase + 3                   'set to Hyp-only base
  End If
  LastTypedInstr = KeyBase
  Call ResetPndAll
End Sub

'*******************************************************************************
' Subroutine Name   : HkeySkeySupport
' Purpose           : Support Hkey and Skey commands
'*******************************************************************************
Public Sub HkeySkeySupport()
  Dim S As String, C As String
  Dim Idx As Long, FmK As Long, ToK As Long, i As Long
  Dim Dash As Boolean
  
  If CBool(Len(DspTxt)) Then              'if key buffer contains data
    FmK = 0                               'init registers
    ToK = 0
    Dash = False
    For Idx = 1 To 26                     'flush array
      Kyz(Idx) = False
    Next Idx
    S = UCase$(DspTxt & ",")              'force final ops...
    
    For Idx = 1 To Len(S)                 'parse contents
      C = Mid$(S, Idx, 1)                 'get a character
      Select Case C
        Case "A" To "Z"
          If FmK = 0 Then                 'assume A, or A-, or A-B
            FmK = Asc(C) - 64
          ElseIf Dash Then                'if dash already processed
            If ToK = 0 Then
              ToK = Asc(C) - 64           'assume A-B, -B
              For i = FmK To ToK          'process A-B...
                Kyz(i) = True             'tag one way or the other
              Next i
              FmK = 0                         'init new list
              ToK = 0
              Dash = False
            Else
              ForcError "Bad data format"
            End If
          Else
            ForcError "Bad data format"
          End If
        Case "-"
          If Dash Then                    'if dash already defined
            ForcError "Bad data format"
          Else
            Dash = True                   'tag dash
            If FmK = 0 Then               'if fmk=0, then assume -B
              FmK = 1
            End If
          End If
        Case ","                          'assume separator, such as A,B,C-F
          If CBool(FmK) Then
            If Dash Then
              ToK = 26
            Else
              ToK = FmK
            End If
            For i = FmK To ToK
              Kyz(i) = True
            Next i
          End If
          FmK = 0                         'init new list
          ToK = 0
          Dash = False
      End Select
    Next Idx
    DspTxt = vbNullString                   'now clear prompt list
    DisplayReg = PndImmed                   'reset display to prior value
    Call ResetPndAll                        'reset any pending
  End If
End Sub

'*******************************************************************************
' Function Name     : CheckList
' Purpose           : Get a list of user-defined types
'*******************************************************************************
Public Function CheckList(ByVal Key As Integer) As Boolean
  Dim Typ As LblTypes
  Dim Idx As Integer, cnt As Integer, Sz As Integer, i As Integer, j As Integer
  Dim S As String, Ary() As String, T As String, U As String
  Dim Pidx As Long
  Dim Obj As clsVarSto
  
  Select Case Abs(Key)
    Case iMDL                                             'Module?
      If Key > 0 Then                                     'not forced
        If PndIdx <> 1 Then Exit Function                 'if nothing pending
        If PndStk(1) <> iList Then Exit Function          'if pending, but not List
        Call ResetPndAll                                  'else clear pend stack
      End If
      CheckList = True                                    'force true result
      DspPgmList = False                                  'indicate not a program list
      If Not CBool(ModName) Then
        ForcError "No Library Module loaded"              'if module not loaded
        Exit Function
      End If
'
' disable any locking
'
      If DspLocked Then
        DspLocked = False
        Call DspBackground
      End If
'
' prepare screen
'
      Clear_Screen
      DisplayMsg "Program list for Module " & CStr(ModName)
      DisplayMsg CStr(ModCnt) & " out of 99 possible programs exist"
      DisplayMsg String$(DisplayWidth, "-")               'display header
      For Idx = 0 To ModCnt - 1                           'process each program
        Pidx = ModMap(Idx)                                'get the start of a module
        S = vbNullString                                  'init title text
        If ModMem(Pidx) = iRem2 Then                      '"'" to start?
          For i = 1 To DisplayWidth                       'yes, so gather name
            Pidx = Pidx + 1                               'point to a character
            j = ModMem(Pidx)                              'grab it
            Select Case j
              Case Is < 10                                '0-9?
                Exit For                                  'done
              Case Is < 128                               'ASCII?
                S = S & Chr$(j)                           'append it if so
              Case Else
                Exit For                                  'other instructions
            End Select
          Next i
          T = "Pgm" & Format(Idx + 1, "00") & ": "        'init program data
          If CBool(Len(S)) Then                           'we have a title
            T = T & Trim$(S)                              'so add the title
          Else
            T = T & "NO TITLE"                            'else no title
          End If
          DisplayMsg T                                    'display result
        End If
      Next Idx                                            'display full contenst
      ModuleList = True
    
    Case iSbr, iUkey, iLbl, iStruct, iEnum, iConst, iVar, iPgm 'allowed types
      If Key > 0 Then
        If PndIdx <> 1 Then Exit Function                 'if nothing pending
        If PndStk(1) <> iList Then Exit Function          'if pending, but not List
        Call ResetPndAll                                  'else clear pend stack
      End If
      CheckList = True                                    'force true result
      DspPgmList = False                                  'indicate not a program list
      If CBool(ActivePgm) Then
        ForcError "Not enabled for Module Programs"       'if not Pgm00
        Exit Function
      End If
      If InstrCnt = 0 Then
        ForcError "No program present to list"            'if no program code
        Exit Function
      End If
      If Not Preprocessd Then                             'if not Preprocessed
        Call Preprocess                                   'Preprocess it
        If Not Preprocessd Then Exit Function             'if still not Preprocessed
      End If
      
      Select Case Abs(Key)
        Case iSbr
          Typ = TypSbr                                    'set type to search for
        Case iUkey
          Typ = TypKey
        Case iLbl
          Typ = TypLbl
        Case iStruct
          Typ = TypStruct
        Case iEnum
          Typ = TypEnum
        Case iConst
          Typ = TypConst
        Case iVar
          Typ = -1
        Case iPgm
          Call frmVisualCalc.mnuWinASCII_Click            'list program
          LastTypedInstr = 128                            'ignore 2nd key code
          Exit Function
      End Select
      
      T = GetInst(Abs(Key))                               'get instruction
      Sz = 128
      ReDim Ary(128)                                      'init list heading
      cnt = 0
      Ary(0) = "Defined items of type '" & T & "':"
'
' now build list
'
      If Typ = -1 Then                                    'Variables...
        For Idx = 0 To MaxVar                             'scan each variable
          With Variables(Idx)
            If .VuDef Then                                'if user-defined...
              cnt = cnt + 1                               'count 1
              If cnt > Sz Then
                Sz = Sz + 128
                ReDim Preserve Ary(Sz)                    'set aside space to store it
              End If
              Select Case .VarType
                Case vNumber
                  S = GetInstrStr(iNvar)
                Case vString
                  S = GetInstrStr(iTvar)
                Case vInteger
                  S = GetInstrStr(iIvar)
                Case vChar
                  S = GetInstrStr(iCvar)
              End Select
              If CBool(Len(Trim$(.VName))) Then
                S = S & "(" & Format(Idx, "00") & ") " & RTrim$(.VName) 'init definition
              Else
                S = S & Format(Idx, "00")                 'init definition
              End If
              Set Obj = .Vdata.LnkNext                    'point to base of X dim data
              If Not Obj Is Nothing Then                  'multi-dimensional?
                S = S & "[" & CStr(Obj.GetMaxDim) & "]"   'get X dimension
                Set Obj = Obj.LnkChild                    'does the X-dim have children (Y dims)?
                If Not Obj Is Nothing Then
                  S = S & "[" & CStr(Obj.GetMaxDim) & "]" 'get Y dimension if so
                End If
              End If
              Ary(cnt) = S & " @ " & CStr(.Vaddr)         'apply address
            End If
          End With
        Next Idx
      Else                                                'all other types
        For Idx = 0 To LblCnt - 1
          With Lbls(Idx)
            If .lblUdef And Typ = .LblTyp Then            'if user-defined of desired type...
              cnt = cnt + 1                               'count 1
              If cnt > Sz Then
                Sz = Sz + 128
                ReDim Preserve Ary(Sz)                    'set aside space for it
              End If
              If .LblScope = Pub Then                     'set scope
                S = GetInstrStr(iPub)
              Else
                S = GetInstrStr(iPvt)
              End If
              If Idx < 27 Then
                U = " " & Chr$(Idx + 64)
              Else
                U = vbNullString
              End If
              Ary(cnt) = S & T & U & " " & RTrim$(.lblName) & " @ " & CStr(.lblAddr)
            End If
          End With
        Next Idx
      End If
'
' now display results
'
      If cnt = 0 Then                                     'if none found...
        ForcError "No user-defined items found of type '" & T & "'"
      Else
        Call Clear_Screen                                 'else clear the screen
        LockControlRepaint frmVisualCalc.lstDisplay       'lock up
        With frmVisualCalc.lstDisplay
          .RemoveItem 0                                   'remove line added by Clear_Screen
          For Idx = 0 To cnt
            .AddItem Ary(Idx)                             'list items
          Next Idx
          Call SelectOnly(0)                              'set top line
        End With
        UnlockControlRepaint frmVisualCalc.lstDisplay     'lock up
        LastTypedInstr = 128                              'ignore 2nd key code
        DspLocked = True                                  'lock keyboard
        Call DspBackground
      End If
  End Select
End Function

'*******************************************************************************
' Subroutine Name   : PushPendKey
' Purpose           : Push a pending command onto the pending stack
'*******************************************************************************
Public Sub PushPendKey()
  If PndIdx = 5 Then
    ForcError "Sequential pending operation stack is too deep"
    ResetPndAll
  Else
    PndIdx = PndIdx + 1             'bump pending index
    PndStk(PndIdx) = LastTypedInstr 'save pending operation
    Call ResetPnd                   'reset pending info
    Call ResetValueAccum            'reset accumulator data
    PndPrev = PndImmed              'store previous pending value
    PndImmed = DisplayReg           'save Display Register
    DisplayReg = 0#                 'init Display Register to 0
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : CheckPnd
' Purpose           : Check for pending commands and data. Used by keyboard input
'*******************************************************************************
Public Sub CheckPnd(ByVal Key As Integer)
  Dim TV As Double
  Dim Idx As Long, i As Long, JL As Long
  Dim j As Integer, K As Integer
  Dim Vtyp As Vtypes
  Dim VV As Variant
  Dim S As String, Ary() As String, T As String
  Dim ts As TextStream
  Dim Bol As Boolean
  Dim Pool() As Labels
  
  Call ResetValueAccum            'reset accumulator data
  If CBool(CharCount) And CBool(CharLimit) And CBool(Key) And CBool(PndIdx) Then
    Call CheckPnd(0)              'flush out pending numeric input
  End If
  
  If CheckList(Key) Then Exit Sub 'Get a list of user-defined types, if required
  
  Select Case Key
    '-----------------------------
    Case iGTO, iCall
      If Not CBool(ActivePgm) Then
        If Not Preprocessd Then   'we need to Preprocess?
          Preprocess              'then do it
          LastTypedInstr = Key    'recover key
          If Not Preprocessd Then 'if error
            ResetPndAll
            Exit Sub
          End If
        End If
      End If
      
      PushPendKey                 'push current operation
      DspTxt = vbNullString       'initialize pending text
      CharCount = 0               'init character counte
      CharLimit = LabelWidth      'set character limit to width of 1 field
      AllowSpace = False          'do not allow typing of a space
      With frmVisualCalc
        Call .checkTextEntry(True) 'set up keyboard
        With .lstDisplay
          .List(.ListIndex) = vbNullString 'init display line
        End With
      End With
      SetTip vbNullString         'clear tip field
      DisplayReg = PndImmed       'reset display register
    '-----------------------------
    Case iHyp, iArc               'Hyp, Arc
      PushPendKey                 'push current operation
      DisplayReg = PndImmed       'reset display register
    '-----------------------------
    Case iMDL, iList              'MDL, List
      PushPendKey                 'push current operation
      DisplayReg = PndImmed       'reset display register
    '-----------------------------
    Case iFix                     'Fix
      PushPendKey                 'push current operation
      If PndIdx = 2 And PndStk(1) = iMDL Then
        DisplayReg = PndImmed     'reset display register (Allow MDL FIX (save))
      Else
        CharLimit = 1             'else set max characters we can enter
      End If
    '-----------------------------
    Case 155, 156, 157, 283, 284, 285 'Sin, Cos, Tan, Sec, Csc, Cot
      If CBool(PndIdx) Then
        Select Case PndStk(PndIdx)
          Case iArc, iHyp
           PushPendKey                 'push current operation
            DisplayReg = PndImmed       'reset display register
            Call CheckPnd(0)            'flush out pending input
        End Select
      End If
    '-----------------------------
    Case iLbl
      If PndIdx = 1 And PndStk(1) = iMDL Then
        Call ResetPndAll          'flush pends ops (we no longer need them)
        LastTypedInstr = 128
        DisplayReg = CDbl(ModName)
        If CBool(Len(Trim$(ModLbl))) Then
          DisplayMsg Trim$(ModLbl)
        Else
          DisplayMsg "<LNoName>"
        End If
      End If
    '-----------------------------
    Case iSUM, iSUB, iRCL
      If PndIdx = 1 And PndStk(1) = iMDL Then
        Call ResetPndAll          'flush pends ops (we no longer need them)
        LastTypedInstr = 128
        Select Case Key
          Case iSUM 'add current program to module
            Call Sum_MDL
          Case iSUB 'remove last pgm from Module
            If CBool(ModCnt) Then                           'if something we can do
              If ModLocked Then                             'if locked, report error
                ForcError "Cannot Subtract. Module is locked"
                Exit Sub
              End If
              If CenterMsgBoxOnForm(frmVisualCalc, _
                 "Are you sure you wish to subtract Pgm " & _
                 Format(ModCnt, "00") & _
                 " from the active Module?", _
                 vbYesNo Or vbQuestion, _
                 "Confirm Module Pgm Delete") = vbNo Then Exit Sub
              ModCnt = ModCnt - 1                           'drop 1 from count
              If CBool(ActivePgm) Then                      'reset program to Pgm 00
                ActivePgm = 0                               'force Pgm 00
                InstrPtr = 0                                'reset instruction pointer
              End If
              If Not CBool(ModCnt) Then                     'if Module erased
                SaveSetting App.Title, "Settings", "LoadedMDL", "0"
                If CBool(ModName) Then                      'if module has a name
                  If CBool(Len(StorePath)) Then             'if poath exists...
                    S = AddSlash(StorePath) & "MDL\MDL" & Format(ModName, "0000") & ".mdl"
                    If Fso.FileExists(S) Then               'if module exists as as file...
                      If CenterMsgBoxOnForm(frmVisualCalc, _
                         "Module Empty. Erase from Data Storage?", _
                         vbYesNo Or vbQuestion, _
                         "Delete Module") = vbYes Then
                        On Error Resume Next
                        Fso.DeleteFile S                    'delete module
                        On Error GoTo 0
                      End If
                    End If
                  End If
                End If
                ModName = 0
              End If
              Call UpdateStatus                             'update status bar
            Else
              ForcError "NO programs to subtract from Module " & Format(ModName, "0000")
            End If
          Case iRCL 'Recall # programs in module
            DisplayReg = CDbl(ModCnt)
            DisplayText = False
            Call DisplayLine
        End Select
      Else
        PushPendKey               'push current operation
        CharLimit = 2             'set max characters we can enter
      End If
    '-----------------------------
    'STO,EXC,MUL,DIV,IND,INCR,DECR,VAR,
    'TRIM,LTRIM,RTRIM,OP,Pgm,Nvar,Tvar,Ivar,Cvar, ClrVar, USR
    Case iSTO, iEXC, iMUL, iDIV, iIND, iIncr, iDecr, iVar, iTrim, iLTrim, _
         iRTrim, iOP, iPgm, iNvar, iTvar, iIvar, iCvar, iClrVar, iUSR
      PushPendKey                 'push current operation
      CharLimit = 2               'set max characters we can enter
    '-----------------------------
    'Load
    Case iLoad
      PushPendKey                 'push current operation
      CharLimit = 2               'init default character count to 2
      If PndIdx = 2 Then          '0-99
        If PndStk(1) = iMDL Then   'if MDL Load
          CharLimit = 4           '0-9999
        End If
      End If
    '-----------------------------
    'Save
    Case iSave
      If PndIdx = 1 And PndStk(1) = iMDL Then
        Call ResetPndAll            'reset pending ops
        LastTypedInstr = 128
        If Not CBool(ModName) Then
          ForcError "No code to save in the Module space"
        ElseIf ModLocked Then       'already locked?
          ForcError "Module cannot be saved. It is LOCKED"
        Else
          Call Save_MDL             'write module
        End If
        Exit Sub
      End If
      ' check for MDL FIX SAVE
      If PndIdx = 2 And PndStk(1) = iMDL And PndStk(2) = iFix Then
        Call ResetPndAll
        LastTypedInstr = 128
        If Not CBool(ModName) Then
          ForcError "No code to save in the Module space"
        ElseIf ModLocked Then       'already locked?
          ForcError "Module cannot be saved. It is LOCKED"
        Else
          ModLocked = True          'force lock on
          Call Save_MDL             'write module
        End If
        Exit Sub
      End If
'
' not Module, so assume Pgm
'
      PushPendKey                 'push current operation
      CharLimit = 2               'init default character count to 2
    '-----------------------------
    'Lapp, ASCII
    Case iLapp, iASCII
      PushPendKey                 'push current operation
      CharLimit = 2               'init default character count to 2
    '-----------------------------
    'StFlg, RFlg, SysBP, >>, <<, Style
    Case iStFlg, iRFlg, 326, 226, 354, iStyle
      PushPendKey                 'push current operation
      CharLimit = 1               'set max characters we can enter
    '-----------------------------
    Case 148, 276                 'Hkey, Skey
      PushPendKey                 'push current operation
      DspTxt = vbNullString       'initialize pending text
      CharCount = 0               'init character counte
      CharLimit = DisplayWidth    'set character limit to width of 1 field
      AllowSpace = False          'do not allow typing of a space
      With frmVisualCalc
        Call .checkTextEntry(True) 'set up keyboard
        With .lstDisplay
          .List(.ListIndex) = vbNullString 'init display line
        End With
      End With
      SetTip vbNullString         'clear tip field
    '-----------------------------
    Case 320                      'Swap
      PushPendKey                 'push current operation
      DisplayReg = PndImmed       'reset display register
      CharLimit = 2               'set max characters we can enter
    '-----------------------------
    Case 349                      '[,]
      If PndIdx = 1 Then
        If PndStk(1) = 320 Then   'if previous was Swap
          PushPendKey             'push current operation
          DisplayReg = PndPrev    'reset display register
          CharLimit = 2           'set max characters we can enter
        End If
      End If
    '-----------------------------
    Case 204  'Pmt
      PushPendKey                 'push current operation
      DspTxt = vbNullString       'initialize pending text
      CharCount = 0               'init character counte
      CharLimit = DisplayWidth    'set character limit to width of 1 field
      AllowSpace = True           'allow typing of a space
      With frmVisualCalc
        Call .checkTextEntry(True) 'set up keyboard
        With .lstDisplay
          .List(.ListIndex) = vbNullString 'init display line
        End With
      End With
      SetTip vbNullString         'clear tip field
    '-----------------------------
    Case 333  'All
      If PndIdx = 1 Then
        If PndStk(1) = 240 Then   'ClrVar (ClrVar All)?
          Call ClearAllVariables  'clear all defined variables
          Call ResetPndAll
        End If
      End If
    
    '-----------------------------
    ' process pending operation(s)...
    '-----------------------------
    Case 0
      Select Case PndIdx
        '-----------------------------
        ' 1 operation pending...
        '-----------------------------
        Case 1                    'if 1 pending....
          Select Case PndStk(1)
            '-----------------------------
            Case 320  'Swap
              Exit Sub            'we want to continue pending, because [,] and 2nd vaue to come...
            '-----------------------------
            Case iComma  '[,]
              ForcError "Comma [,] is improperly used"
            '-----------------------------
            Case 240  'ClrVar
              Call ClearVariable(DisplayReg)  'clear an individual variable
              DisplayReg = PndImmed           'reset value to previous
            '-----------------------------
            Case 226  '>>
              If DisplayReg < 1# Or DisplayReg > 9# Then
                ForcError "Invalid Shift value (1-9)"
              Else
                i = CLng(DisplayReg)                'get shift count
                DisplayReg = PndImmed               'get value to operate on
                For Idx = 1 To i
                  If DisplayReg = 0# Then Exit For  'if null, then nothing to do
                  DisplayReg = DisplayReg / 2#
                Next Idx
              End If
            '-----------------------------
            Case 354  '<<
              If DisplayReg < 1# Or DisplayReg > 9# Then
                ForcError "Invalid Shift value (1-9)"
              Else
                i = CLng(DisplayReg)                'get shift count
                DisplayReg = PndImmed               'get value to operate on
                If DisplayReg <> 0# Then
                  On Error Resume Next
                  For Idx = 1 To i
                    DisplayReg = DisplayReg * 2#    'shift left 1 bit
                    Call CheckError
                    If ErrorFlag Then Exit For
                  Next Idx
                  On Error GoTo 0
                End If
                If ErrorFlag Then
                  DisplayReg = PndImmed             'if an error, reset value
                End If
              End If
            '-----------------------------
            Case iGTO   'GTO
              If CBool(CharCount) Then      'if a response was entered
                If CBool(ActivePgm) Then
                  ForcError "Command invalid for Module programs"
                Else
                  RunMode = False           'ensure RUN is off
                  MRunMode = 0              'and Module run
                  ModPrep = 0
                  If Len(DspTxt) = 1 Then
                    Select Case UCase$(DspTxt)
                      Case "A" To "Z"
                        JL = Asc(UCase$(DspTxt)) - 64           'get A-Z offset
                        If CBool(ActivePgm) Then                'if module
                          JL = JL + ModLblMap(ActivePgm - 1)
                          If ModLbls(JL).LblDat = 0 Then JL = 0 'if not defined in module
                        Else
                          If Lbls(JL).LblDat = 0 Then JL = 0    'if not defined in pgm
                        End If
                      Case Else
                        JL = 0
                    End Select
                  Else
                    JL = FindLblMatch(DspTxt) 'search for matching name
                  End If
                  If JL = 0 Then
                    ForcError "Selected Label does not exist"
                  Else
                    If CBool(ActivePgm) Then
                      InstrPtr = ModLbls(JL).lblAddr
                    Else
                      InstrPtr = Lbls(JL).lblAddr
                    End If
                    StopMode = False        'disable stop mode if we changed instruction
                    BraceIdx = 0            'reset braceing index us user key hit
                    SbrInvkIdx = 0          'reset subr stack if user-key hit
                    Call UpdateStatus       'update status
                  End If
                End If
              End If
              DisplayReg = PndImmed         'reset display register
            '-----------------------------
            Case iCall  'CALL
              If CBool(CharCount) Then            'if a response was entered
                RunMode = False                   'ensure RUN is off
                MRunMode = 0                      'and Module run
                ModPrep = 0
                JL = FindLblMatch(DspTxt)         'search for matching name
                If JL = 0 Then
                  ForcError "Selected Label does not exist"
                Else
                  If CBool(ActivePgm) Then
                    Pool = ModLbls
                  Else
                    Pool = Lbls
                  End If
                  
                  With Pool(JL)
                    Select Case .LblTyp
                      Case TypSbr, TypKey         'allow only Ukey and Sbr
                        BraceIdx = 0              'reset braceing index us user key hit
                        SbrInvkIdx = 0            'reset subr stack if user-key hit
                        Call PushCall(JL, .lblAddr, ActivePgm, True) 'set new instr. ptr
                        Call Run                  'run code
                        DisplayReg = PndImmed     'reset display register
                        If PmtFlag Then           'if user prompting turned on...
                          LastTypedInstr = iTXT   'set TXT command
                          Call ActiveKeypad       'activate it (TXT or [=] (ENTER)) will start R/S cmd
                        Else
                          Call DisplayLine        'else terminating run...
                        End If
                      Case Else
                        ForcError "Invalid parameter"
                    End Select
                  End With
                End If
              End If
            '-----------------------------
            Case 141  'STO
              If Variables(CLng(DisplayReg)).VarType = vString Then
                Call SetVarValue(DisplayReg, CVar(DspTxt))        'assume data stored in DspTxt
              Else
                Call SetVarValue(DisplayReg, CVar(PndImmed))
              End If
              DisplayReg = PndImmed
            '-----------------------------
            Case 142 'RCL
              If Variables(CLng(DisplayReg)).VarType = vString Then
                DspTxt = CStr(GetVarValue(DisplayReg))
                With frmVisualCalc.lstDisplay
                  .List(.ListIndex) = String$(DisplayWidth - Len(DspTxt), 32) & DspTxt
                End With
                Call ResetPndAll              'reset pending commands
                Exit Sub                      'all done
              Else
                DisplayReg = CDbl(GetVarValue(DisplayReg))
              End If
            '-----------------------------
            Case 143  'EXC
              If Variables(CLng(DisplayReg)).VarType = vString Then
                S = CStr(GetVarValue(DisplayReg))
                Call SetVarValue(DisplayReg, CVar(DspTxt))  'stuff immediate value to var
                DspTxt = S
                DisplayText = True
                DisplayReg = PndImmed
              Else
                TV = DisplayReg                       'variable to exchange with
                DisplayReg = CDbl(GetVarValue(TV))    'set display register to stored value
                Call SetVarValue(TV, CVar(PndImmed))  'stuff immediate value to var
              End If
            '-----------------------------
            Case 144  'SUM
              i = CLng(DisplayReg)
              If Variables(i).VarType = vString Then  'if text, merge variable with DspData
                With Variables(i).Vdata
                  .VarStr = .VarStr & DspTxt
                End With
              Else
                On Error Resume Next
                TV = CDbl(GetVarValue(DisplayReg)) + PndImmed 'add immediate value to var value
                Call CheckError
                On Error GoTo 0
                If Not ErrorFlag Then
                  Call SetVarValue(DisplayReg, CVar(TV))  'stuff result value to var
                  DisplayReg = PndImmed                   'reset display to prior value
                End If
              End If
            '-----------------------------
            Case 145  'MUL
              If Variables(CLng(DisplayReg)).VarType = vString Then
                ForcError "This operation cannot be performed with Text variables"
              Else
                On Error Resume Next
                TV = CDbl(GetVarValue(DisplayReg)) * PndImmed 'mult immediate value to var value
                Call CheckError
                On Error GoTo 0
                If Not ErrorFlag Then
                  Call SetVarValue(DisplayReg, CVar(TV))  'stuff result value to var
                  DisplayReg = PndImmed                   'reset display to prior value
                End If
              End If
            '-----------------------------
            Case 272  'SUB
              If Variables(CLng(DisplayReg)).VarType = vString Then
                ForcError "This operation cannot be performed with Text variables"
              Else
                On Error Resume Next
                TV = CDbl(GetVarValue(DisplayReg)) - PndImmed 'sub immediate value from var value
                Call CheckError
                On Error GoTo 0
                If Not ErrorFlag Then
                  Call SetVarValue(DisplayReg, CVar(TV))  'stuff result value to var
                  DisplayReg = PndImmed                   'reset display to prior value
                End If
              End If
            '-----------------------------
            Case 273  'DIV
              If Variables(CLng(DisplayReg)).VarType = vString Then
                ForcError "This operation cannot be performed with Text variables"
              Else
                TV = DisplayReg                         'variable to exchange with
                On Error Resume Next
                TV = CDbl(GetVarValue(DisplayReg)) / PndImmed 'div var value by immediate value
                Call CheckError
                On Error GoTo 0
                If Not ErrorFlag Then
                  Call SetVarValue(DisplayReg, CVar(TV))  'stuff result value to var
                  DisplayReg = PndImmed                   'reset display to prior value
                End If
              End If
            '-----------------------------
            Case 212  'Nvar
              With Variables(CLng(DisplayReg))
                .VarType = vNumber
                .VuDef = True
              End With
            
            Case 225  'Tvar
              With Variables(CLng(DisplayReg))
                .VarType = vString
                .VuDef = True
              End With
              
            Case 340  'Ivar
              With Variables(CLng(DisplayReg))
                .VarType = vInteger
                .VuDef = True
              End With
              
            Case 353  'Cvar
              With Variables(CLng(DisplayReg))
                .VarType = vChar
                .VuDef = True
              End With
              
            '-----------------------------
            Case 329  'INCR
              If Variables(CLng(DisplayReg)).VarType = vString Then
                ForcError "This operation cannot be performed with Text variables"
              Else
                TV = CDbl(GetVarValue(DisplayReg)) + 1# 'add 1 to var value
                Call SetVarValue(DisplayReg, CVar(TV))  'stuff result value to var
                DisplayReg = PndImmed                   'reset display to prior value
              End If
            '-----------------------------
            Case 330  'DECR
              If Variables(CLng(DisplayReg)).VarType = vString Then
                ForcError "This operation cannot be performed with Text variables"
              Else
                TV = CDbl(GetVarValue(DisplayReg)) - 1# 'subtract one from var value
                Call SetVarValue(DisplayReg, CVar(TV))  'stuff result value to var
                DisplayReg = PndImmed                   'reset display to prior value
              End If
            '-----------------------------
            Case 287  'VAR
              CurrentVar = CInt(DisplayReg)
              With Variables(CurrentVar)
                CurrentVarTyp = .VarType
                Set CurrentVarObj = .Vdata            'keyboard mode uses ONLY base variables
              End With
              DisplayReg = PndImmed                   'reset display to prior value
            '-----------------------------
            Case 309  'Trim
              With Variables(CInt(DisplayReg))
                If .VarType <> vString Then
                  ForcError "Variable must be a Text type"
                ElseIf CBool(.VdataLen) Then
                  ForcError "A fixed-length string cannot be trimmed"
                Else
                  .Vdata.VarStr = Trim$(.Vdata.VarStr)
                End If
              End With
            '-----------------------------
            Case 310  'LTrim
              With Variables(CInt(DisplayReg))
                If .VarType <> vString Then
                  ForcError "Variable must be a Text type"
                ElseIf CBool(.VdataLen) Then
                  ForcError "A fixed-length string cannot be trimmed"
                Else
                  .Vdata.VarStr = LTrim$(.Vdata.VarStr)
                End If
              End With
            '-----------------------------
            Case 311  'RTrim
              With Variables(CInt(DisplayReg))
                If .VarType <> vString Then
                  ForcError "Variable must be a Text type"
                ElseIf CBool(.VdataLen) Then
                  ForcError "A fixed-length string cannot be trimmed"
                Else
                  .Vdata.VarStr = RTrim$(.Vdata.VarStr)
                End If
              End With
            '-----------------------------
            Case 204  ' Pmt
              If CBool(CharCount) Then  'if a prompt was entered
                Call NewLine            'advance to new line for user response
                RunMode = False         'disable mode
                Call ResetPndAll        'reset pending commands
                Call UpdateStatus       'update status
                LastTypedInstr = iTXT   'set TXT command
                Call ActiveKeypad       'activate it (TXT or [=] (ENTER)) will start R/S cmd
                PmtFlag = True          'enable prompt flag
                Exit Sub                'do not turn this stuff off
              End If
            '-----------------------------
            Case iLoad, iSave, iLapp, iASCII ' Load, Save, Lapp, ASCII
              If Len(StorePath) = 0 Then
                ForcError "No storage path yet defined"
                Call ResetPndAll
                Exit Sub
              End If
              'try using loaded pgm name, if none supplied as is SAVE or ASCII
              If PndStk(1) = iSave Or PndStk(1) = iASCII Then
                If DisplayReg = 0# Then DisplayReg = CDbl(PgmName)
              End If
              If DisplayReg = 0# Then
                ForcError "New Program Name must be 01-99"
              Else
                T = Format(DisplayReg, "00")
                S = RemoveSlash(StorePath) & "\PGM\Pgm" & T
                j = FreeFile(0)
                Select Case PndStk(1)
                  '-----------------------------
                  Case iLoad  ' Load  Load Binary
                    RunMode = False                         'turn off run modes
                    MRunMode = 0
                    ModPrep = 0
                    If Not Fso.FileExists(S & ".pgm") Then
                      ForcError " Pgm" & T & ".pgm" & vbCrLf & _
                                " Program does not exist in path:" & vbCrLf & _
                                " " & StorePath & "\PGM"
                      Call ResetPndAll
                      Call DisplayLine
                      Exit Sub
                    Else
                      If IsDirty And CBool(PgmName) And PgmName <> CInt(T) Then
                        CmdNotActive
                        If CenterMsgBoxOnForm(frmVisualCalc, _
                           "The current program was loaded as Pgm" & Format(PgmName, "00") & ".pgm," & _
                           " which is presently modified," & vbCrLf & _
                           "and you are wanting to load Pgm" & T & ".pgm. Do you wish to continue?", _
                           vbYesNo Or vbDefaultButton2 Or vbQuestion, _
                           "Overwrite Unsaved Program?") = vbNo Then
                          Call ResetPndAll
                          DisplayReg = PndImmed               'reset display register
                          Call DisplayLine
                          Exit Sub
                        End If
                      ElseIf IsDirty And PgmName = 0 Then
                        If CenterMsgBoxOnForm(frmVisualCalc, _
                           "The current program code is presently modified and unsaved," & vbCrLf & _
                           "and you are wanting to load Pgm" & T & ".pgm. Do you wish to continue?", _
                           vbYesNo Or vbDefaultButton2 Or vbQuestion, _
                           "Overwrite Unsaved Program?") = vbNo Then
                          Call ResetPndAll
                          DisplayReg = PndImmed               'reset display register
                          Call DisplayLine
                          Exit Sub
                        End If
                      End If
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
                        Call ResetBracing                     'reset any special bracing
                        frmVisualCalc.mnuWinASCII.Enabled = True
                        Call UpdateStatus
                        PgmName = CInt(T)
                        IsDirty = False                       'indicate not dirty
                        Call UpdateStatus
                        If AutoPprc Then
                          Call Preprocess
                          If Preprocessd Then
                            DisplayMsg "Loaded Binary Pgm" & T & ".pgm OK"
                          End If
                        Else
                          Call ResetListSupport
                          DisplayMsg "Loaded Binary Pgm" & T & ".pgm OK"
                        End If
                      End If
                      Close #j                                'close file
                    End If
                    DisplayReg = PndImmed                   'reset display to prior value
                  
                  '-----------------------------
                  Case iSave  ' Save Binary code
                    If InstrCnt = 0 Then                    'if nothing to do
                      ForcError "No LEARNED code exists"
                    Else
                      On Error Resume Next
                      If Fso.FileExists(S & ".pgm") Then
                        If PgmName <> CInt(T) And CBool(PgmName) Then
                          CmdNotActive
                          If CenterMsgBoxOnForm(frmVisualCalc, _
                             "This program was loaded as Pgm" & Format(PgmName, "00") & ".pgm," & vbCrLf & _
                             "but you are saving it as Pgm" & T & ".pgm," & vbCrLf & _
                             "which already exists. Continue?", _
                             vbYesNo Or vbQuestion Or vbDefaultButton2, _
                             "Over-Write Program") = vbNo Then
                            Call ResetPndAll
                            DisplayReg = PndImmed               'reset display register
                            Call DisplayLine
                            Exit Sub
                          End If
                        ElseIf ActivePgm = 0 And PgmName = 0 Then
                          If CenterMsgBoxOnForm(frmVisualCalc, _
                             "Overwirte Pgm" & T & ".pgm with this fresh program?", _
                             vbYesNo Or vbQuestion Or vbDefaultButton2, _
                             "Over-Write Program") = vbNo Then
                            Call ResetPndAll
                            DisplayReg = PndImmed               'reset display register
                            Call DisplayLine
                            Exit Sub
                          End If
                        End If
                        Fso.DeleteFile S & ".pgm"
                      End If
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
                    RunMode = False                         'turn off run modes
                    MRunMode = 0
                    ModPrep = 0
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
                        IsDirty = True                      'indicate pool is now dirty
                        Preprocessd = False
                        Compressd = False
                      End If
                      Close #j
                      If AutoPprc Then
                        Call Preprocess
                      Else
                        Call ResetListSupport
                      End If
                    Else
                      ForcError " Pgm" & T & ".pgm" & vbCrLf & _
                                " Program does not exist in path:" & vbCrLf & _
                                " " & StorePath & "\PGM"
                    End If
                    DisplayReg = PndImmed                   'reset display to prior value
                  '-----------------------------
                  Case iASCII  ' ASCII 'save to ASCII file
                    On Error Resume Next
                    If Fso.FileExists(S & ".txt") Then Fso.DeleteFile (S & ".txt")
                    Set ts = Fso.OpenTextFile(S & ".txt", ForWriting, True)
                    Call CheckError                   'check for errors
                    On Error GoTo 0
                    If Not ErrorFlag Then             'if OK
                      Ary = BuildInstrArray()         'get array list
                      If IsDimmed(Ary) Then
                        ts.Write Join(Ary, vbCrLf)      'write data
                        DisplayMsg "Saved ASCII Pgm" & T & ".txt OK"
                      End If
                    End If
                    ts.Close                          'close file
                    DisplayReg = PndImmed             'reset display to prior value
                  '-----------------------------
                End Select
              End If
            '-----------------------------
            'Sin, Cos, Tan, Sec, Csc, Cot
            Case 155, 156, 157, 283, 284, 285
              Call CheckAngles
              Exit Sub
            '-----------------------------
            Case 162  'StFlg
              flags(DisplayReg) = True
              DisplayReg = PndImmed                   'reset display to prior value
              SetTip vbNullString
            '-----------------------------
            Case 290  'RFlg
              flags(DisplayReg) = False
              DisplayReg = PndImmed                   'reset display to prior value
              SetTip vbNullString
            '-----------------------------
            Case iStyle  'Style
              If DisplayReg > 3# Then
                ForcError "Style parameter is 0-3 (Off/On/On with line numbers)"
              Else
                LRNstyle = CInt(DisplayReg)           'set new style
                SaveSetting App.Title, "Settings", "Style", CStr(LRNstyle)
                Call UpdateStatus
                DisplayReg = PndImmed                 'reset display to prior value
                If Not LrnMode Then
                  If DspLocked Then
                    Call frmVisualCalc.mnuWinASCII_Click  'list program
                    Call ResetPndAll
                    Exit Sub
                  ElseIf AutoPprc Then
                    If Not Preprocessd Then               'if not at least Preprocessed...
                      Call Preprocess                     'then Preprocess
                    End If
                  End If
                End If
              End If
              SetTip vbNullString
            '-----------------------------
            Case 135  'OP
              TV = Fix(DisplayReg)
              DisplayReg = PndImmed                   'reset display to prior value
              If TV < 0# Or TV > CDbl(MaxOps) Then
                ForcError "Op Code is not defined"
              Else
                Call ResetPndAll                      'OP must do this here, regardless of later invoke
                Call ProcessOP(CLng(TV))
              End If
            '-----------------------------
            Case iUSR  'USR
              TV = Fix(DisplayReg)
              DisplayReg = PndImmed                   'reset display to prior value
              If TV < 0# Or TV > CDbl(MaxUSR) Then
                ForcError "USR Op Code is not defined"
              Else
                Call ResetPndAll                      'USR must do this here, regardless of later invoke
                Call ProcessUSR(CLng(TV))
              End If
            '-----------------------------
            Case 130  'Pgm
              TV = Fix(DisplayReg)                      'get program number
              DisplayReg = PndImmed                     'reset display to prior value
              If TV < 0# Or TV > CDbl(ModCnt) Then      'if greater than max module number
                ForcError "Pgm " & Format(TV, "00") & " is not defined in the active module"
              Else
                i = CInt(TV)                            'get desired module program
                If i <> ActivePgm Then                  'if we are not in the same program...
                  ActivePgm = i                         'set active program
                  InstrErr = 0                          'reset instruction pointer
                  Call Reset_Support
                  Call RedoAlphaPad
                  Call UpdateStatus
                  If CBool(ActivePgm) Then              'if module program
                    i = ModLblMap(ActivePgm - 1)        'get lower bounds
                    JL = ModLblMap(ActivePgm) - 1       'get upper bounds
                    Pool = ModLbls                      'get pool to search
                  Else                                  'if pgm 00
                    i = 0                               'get lower bounds
                    JL = LblCnt                         'get upper bounds
                    Pool = Lbls                         'get pool to search
                  End If
                  For Idx = i To JL                     'scan selected pool
                    With Pool(Idx)
                      If .LblTyp = TypSbr Then          'found a Sbr?
                        If StrComp(Trim$(.lblName), "MAIN", vbTextCompare) = 0 Then
                          InstrPtr = .LblDat + 1        'was main(), so set to code block + 1
                          Call Run                      'run Main() subr
                          If PmtFlag Then               'if user prompting turned on...
                            LastTypedInstr = iTXT       'set TXT command
                            Call ActiveKeypad           'activate it (TXT or [=] (ENTER)) will start R/S cmd
                          Else
                            Call DisplayLine            'else terminating run...
                          End If
                          Exit For
                        End If
                      End If
                    End With
                  Next Idx
                End If
              End If
            '-----------------------------
            Case iFix  'Fix
              TV = Fix(DisplayReg)                      'get format
              If TV < 0# Or TV > 9# Then
                ForcError "Parameter is out of range (0-9)"
              Else
                DspFmtFix = CLng(TV)                    'save decimal count
                DisplayReg = PndImmed                   'reset display to prior value
                DspFmt = "0." & String$(DspFmtFix, "0") 'set format
                ScientifEE = DspFmt & "E+00"
              End If
              SetTip vbNullString
            '-----------------------------
            Case 148  'Hkey
              Call HkeySkeySupport                    'process list
              For Idx = 1 To 26
                If Kyz(Idx) Then Hidden(Idx) = True   'mark key as hidden
              Next Idx
              Call RedoAlphaPad
            '-----------------------------
            Case 276  'Skey
              Call HkeySkeySupport                    'process list
              For Idx = 1 To 26
                If Kyz(Idx) Then Hidden(Idx) = False  'mark key as visible
              Next Idx
              Call RedoAlphaPad
            '-----------------------------
            Case 326  'SysBP
              Dim SB As BeepType
              TV = Fix(DisplayReg)
              DisplayReg = PndImmed                   'reset display to prior value
              If TV < 0# Or TV > 4# Then
                ForcError "Parameter is out of range (0-4)"
              Else
                Select Case CInt(TV)
                  Case 1
                    SB = beepSystemAsterisk
                  Case 2
                    SB = beepSystemExclamation
                  Case 3
                    SB = beepSystemHand
                  Case 4
                    SB = beepSystemQuestion
                  Case Else
                    SB = beepSystemDefault
                End Select
                Call MsgBeep(SB)
              End If
              SetTip vbNullString
            '-----------------------------
            Case 226  '>>
              TV = Fix(DisplayReg)        'get shift amount
              If TV < 1# Or TV > 9# Then  'verify range ok
                ForcError "Parameter is out of range (0-9)"
              Else
                i = CLng(TV)              'get range
                TV = Fix(PndImmed)        'get value to change
                For Idx = 1 To i
                   TV = Fix(TV / 2#)      'shift right once
                Next Idx
                DisplayReg = TV           'result to display
              End If
              SetTip vbNullString
            '-----------------------------
            Case 354  '<<
              TV = Fix(DisplayReg)        'get shift amount
              If TV < 1# Or TV > 9# Then  'verify range ok
                ForcError "Parameter is out of range (0-9)"
              Else
                i = CLng(TV)              'get range
                TV = Fix(PndImmed)        'get value to change
                For Idx = 1 To i
                   TV = TV * 2#           'shift right once
                Next Idx
                DisplayReg = TV           'result to display
              End If
              SetTip vbNullString
            '-----------------------------
            Case Else
              ForcError "Command is not supported in chained operations"
          End Select
          Call ResetPndAll
          Call DisplayLine
        '-----------------------------
        ' 2 operations pending...
        '-----------------------------
        Case 2
          Select Case PndStk(2)
            Case iIND  'IND
              Select Case PndStk(1)
                'STO, RCL, EXC, SUM, MUL, SUB, DIV, INC, DEC, Fmt
                Case 141, 142, 143, 144, 145, 272, 273, 329, 330, iFmt
                  TV = Fix(DisplayReg)
                  If TV < 0# Or TV > DMaxVar Then
                    ForcError "Variable number is out of range"
                    Exit Sub
                  End If
                  If PndStk(1) <> iFmt Then
                    DisplayReg = CDbl(GetVarValue(TV))
                    PndIdx = 1            'set index as 1 opn
                    Call CheckPnd(0)      'recurse to it
                    Exit Sub              'further processing done by invoke
                  ElseIf Variables(CLng(TV)).VarType = vString Then
                    DspTxt = CStr(GetVarValue(TV))
                    DspFmt = DspTxt                 'employ user-supplied format
                    DspFmtFix = -1                  'disabled fixed format
                    DisplayReg = PndImmed           'reset display to prior value
                    CharLimit = 0                   'allow DisplayLine to work
                    Call DisplayLine
                  Else
                    ForcError "Text Variable expected"
                  End If
              End Select
            '-----------------------------
            Case iLoad, iSave, iLapp, iASCII 'Load, Save, Lapp, ASCII
              If PndStk(1) = 130 Then   'Pgm (Pgm Load, Pgm Save, Pgm Lapp, Pgm ASCII)
                PndStk(1) = PndStk(2)   'set Load,Save,Lapp,ASCII to base stack posn
                PndIdx = 1              'set index as 1 opn
                Call CheckPnd(0)        'recurse to it
                Exit Sub                'further processing done by invoke
              ElseIf PndStk(1) = iMDL Then 'MDL (MDL Load, MDL Save)
                Select Case PndStk(2)
                  Case iLoad            'Load
                    Call Load_MDL
                  Case iSave, iLapp, iASCII 'Save,Lapp,ASCII
                    ForcError "MDL does not support this Instruction"
                End Select
              End If
            '-----------------------------
            Case 155, 156, 157, 283, 284, 285 'Sin, Cos, Tan, Sec, Csc, Cot
              Call CheckAngles
            '-----------------------------
            Case iComma  '[,]
              Select Case PndStk(1)
                Case 320 'Swap
                  DisplayReg = Fix(DisplayReg)
                  PndImmed = Fix(PndImmed)
                  If DisplayReg < 0# Or DisplayReg > DMaxVar Or PndImmed < 0# Or PndImmed > DMaxVar Then
                    ForcError "A swap variable is out of range"
                  Else
                    VV = GetVarValue(DisplayReg)
                    Call SetVarValue(DisplayReg, GetVarValue(PndImmed))
                    Call SetVarValue(PndImmed, VV)
                    DisplayReg = PndPrev
                  End If
                Case Else
                ForcError "Comma [,] is improperly used"
              End Select
            Case Else
              ForcError "Command is not supported in chained operations"
          End Select
          Call ResetPndAll
        '-----------------------------
        ' 3 operations pending...
        '-----------------------------
        Case 3
          KeyBase = 0
          Select Case PndStk(3)
            '-----------------------------
            Case 155, 156, 157, 283, 284, 285 'Sin, Cos, Tan, Sec, Csc, Cot
              Call CheckAngles
            If KeyBase = 0 Then
              ForcError "Illegal command use"
            End If
            '-----------------------------
          End Select
          Call ResetPndAll
        '-----------------------------
        ' too many operations pending
        '-----------------------------
        Case Else
          ForcError "Command is not supported in chained operations"
          Call ResetPndAll
      End Select
  End Select
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

