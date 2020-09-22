Attribute VB_Name = "modActivekybd"
Option Explicit

'*******************************************************************************
' Subroutine Name   : ActiveKeypad
' Purpose           : Handle Hot Keyboard
'*******************************************************************************
Public Sub ActiveKeypad()
  Dim S As String
  Dim TV As Double, vDeg As Double, vMin As Double, vSec As Double
  Dim X As Double, Y As Double
  Dim Idx As Long
  Dim Iptr As Integer
  
  If AllowExp Then                  'if Allow Exponent enabled...
    If LastTypedInstr <> 222 Then   'and not +/-
      AllowExp = LastTypedInstr < 10  'keep it only if the typed data are digits
    End If
  End If
  
  Tron = False
  Select Case LastTypedInstr
'-------------------------------------------------------------------------------
              '--Base keys--
'-------------------------------------------------------------------------------
    Case 0 To 9
      Call CheckValue(Chr$(LastTypedInstr + 48))
'-------------------------------------------------------------------------------
    Case Is < 129 'ASCII text
      Exit Sub    'ASCII is handled elsewhere
'-------------------------------------------------------------------------------
'    Case 129  ' LRN   'handled by Main keyboard routing (which invokes us here)
'    Case 130  ' Pgm   'Handled by pending process
'    Case 131  ' Load  'handled by pending process
'    Case 132  ' Save  'handled by pending process
'-------------------------------------------------------------------------------
    Case 133  ' CE
      Call CE_Support
'-------------------------------------------------------------------------------
    Case 134  ' CLR
      Call CLR_Support
'-------------------------------------------------------------------------------
'    Case 135  ' OP  'Handled by pending process
'-------------------------------------------------------------------------------
    Case 136  ' SST
      If StopMode Then                      'prevent continuance if STOP cmd set flag
        CmdNotActive
        Exit Sub
      End If
'
' ensure program exists, and if Pgm 00, ensure that it is Preprocessed
'
      If Not CBool(ActivePgm) And CBool(InstrCnt) Then
        If Not Preprocessd Then             'make sure Preprocessed
          Call Preprocess
          If Not Preprocessd Then Exit Sub  'error
        End If
        SSTmode = True                      'set single-step mode
        Call Run
' if after running, the Pmt command is activated, we will turn on the TXT mode (TextEntry),
' let the user type in a response, and by pressing TXT or '=' (ENTER), the program will continue
' (see KybdMain).
        If PmtFlag Then                     'if user prompting turned on...
          LastTypedInstr = iTXT             'set TXT command
          Call ActiveKeypad                 'activate it (TXT or [=] (ENTER)) will start R/S cmd
        Else
          Call DisplayLine                  'else terminating run...
        End If
      End If
'-------------------------------------------------------------------------------
    Case 137  ' INS
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 138  ' Cut   'active only in Learn Mode
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 139  ' Copy  'active only in Learn Mode
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 140  ' PtoR
      DisplayReg = AngToRad(DisplayReg)       'get angle in Radians
      TV = TestReg * Cos(DisplayReg)          'get x
      DisplayReg = TestReg * Sin(DisplayReg)  'get y
      TestReg = TV                            'set x
      Call DisplayLine
'-------------------------------------------------------------------------------
'    Case 141  ' STO 'handled by pending process
'    Case 142  ' RCL 'handled by pending process
'    Case 143  ' EXC 'handled by pending process
'    Case 144  ' SUM 'handled by pending process
'    Case 145  ' MUL 'handled by pending process
'    Case 146  ' IND 'handled by pending process
'-------------------------------------------------------------------------------
    Case 147  ' Reset 'Reset program counter to 0
      InstrErr = 0
      Call Reset_Support
      Call UpdateStatus
      Call DisplayLine
'-------------------------------------------------------------------------------
'    Case 148  'Hkey 'handled by pending process
'-------------------------------------------------------------------------------
    Case 149  ' lnX   'Natural Logarithm
      DisplayReg = Log(DisplayReg)
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 150  ' E+    'Statistical SUM
      Call StatSUMadd
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 151  ' Mean  'Statistical Mean
      Call StatMean
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 152  ' X!    'display factorial of value 0<= x <=69
      If Factorial(DisplayReg) Then
        Call DisplayLine
      End If
'-------------------------------------------------------------------------------
    Case 153  ' X><T  'Swap Display Register and Test Register
      TV = DisplayReg
      DisplayReg = TestReg
      TestReg = TV
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case iHyp  ' Hyp
'-------------------------------------------------------------------------------
    Case 155  ' Sin
      On Error Resume Next
      DisplayReg = Sin(AngToRad(DisplayReg))
      Call CheckError
      On Error GoTo 0
      If Not ErrorFlag Then Call DisplayLine
'-------------------------------------------------------------------------------
    Case 156  ' Cos
      On Error Resume Next
      DisplayReg = Cos(AngToRad(DisplayReg))
      Call CheckError
      On Error GoTo 0
      If Not ErrorFlag Then Call DisplayLine
'-------------------------------------------------------------------------------
    Case 157  ' Tan
      On Error Resume Next
      DisplayReg = Tan(AngToRad(DisplayReg))
      Call CheckError
      On Error GoTo 0
      If Not ErrorFlag Then Call DisplayLine
'-------------------------------------------------------------------------------
    Case 158  ' 1/X
      On Error Resume Next
      DisplayReg = 1# / DisplayReg
      Call CheckError
      On Error GoTo 0
      If Not ErrorFlag Then Call DisplayLine
'-------------------------------------------------------------------------------
    Case 159  ' Txt
      If CBool(PndIdx) Then
        Call CheckPnd(0)                      'flush out pending numeric input
        If ErrorFlag Then Exit Sub            'if error encountered, then done
      End If
      
      DspTxt = vbNullString                   'initialize pending text
      CharCount = 0                           'init character counter
      CharLimit = DisplayWidth                'set character limit
      AllowSpace = True                       'allow typing of a space
      With frmVisualCalc
        Call .checkTextEntry(True)            'set up keyboard
        With .lstDisplay
          If .ListIndex <> .ListCount - 1 Then 'if current line is not last line...
            DspTxt = Trim$(.List(.ListIndex)) 'grab the line data
            CharCount = Len(DspTxt)           'establish character count
            .List(.ListIndex) = String$(DisplayWidth - Len(DspTxt), 32) & DspTxt  'flush right
            DisplayHasText = True             'text is present
          Else
           .List(.ListIndex) = vbNullString   'init current line
          End If
        End With
      End With
      SetTip vbNullString                     'clear tip field
'-------------------------------------------------------------------------------
    Case 160  ' Hex
      If BaseType <> TypHex Then              'ignore if alread type
        If BaseType = TypDec Then DisplayReg = Fix(DisplayReg)  'remove decimal
        BaseType = TypHex                     'set type
        Call UpdateStatus
        Call EnableNums
        Call DisplayLine
      End If
'-------------------------------------------------------------------------------
    Case 161  ' &
      Call Pend(iAndB)
'-------------------------------------------------------------------------------
'    Case 162  ' StFlg 'handled by pending process
'-------------------------------------------------------------------------------
    Case 163  ' IfFlg 'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 164  ' X==T  'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 165  ' X>=T  'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 166  ' X>T   'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 167  ' Dfn
      If BaseType = TypHex Then
        Call CheckValue("B")
      End If
'-------------------------------------------------------------------------------
    Case iColon  ' : 'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 169  ' (
      Call Pend(iLparen)
'-------------------------------------------------------------------------------
    Case 170  ' )
      Call Pend(iRparen)
'-------------------------------------------------------------------------------
    Case 171  ' /
      Call Pend(iDVD)
'-------------------------------------------------------------------------------
'    Case iStyle  ' Style 'Handled by pending process
'-------------------------------------------------------------------------------
    Case 173  ' Dec
      If BaseType <> TypDec Then              'ignore if already type
        BaseType = TypDec                     'set type
        Call UpdateStatus
        Call EnableNums
        Call DisplayLine
      End If
'-------------------------------------------------------------------------------
    Case 174  ' |
      Call Pend(iOrB)
'-------------------------------------------------------------------------------
    Case 175  ' Int
      DisplayReg = Fix(DisplayReg)
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 176  ' Abs
      DisplayReg = Abs(DisplayReg)
      Call DisplayLine
'-------------------------------------------------------------------------------
'    Case 177  ' Fix 'handled by pending process
'-------------------------------------------------------------------------------
    Case 178  ' D.MS  Convert DDD.MMSSdddd to DDD.ddddd
      vDeg = Fix(DisplayReg)                        'get DDD
      DisplayReg = (DisplayReg - vDeg) * 100#       'get MM.SSsss
      vMin = Fix(DisplayReg)                        'get MM
      vSec = (DisplayReg - vMin) * 100#             'get SS.dddd
      DisplayReg = vDeg + vMin / 60# + vSec / 3600# 'get dd.ddddd
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 179  ' EE
      If PndIdx = 1 Then
        If PndStk(1) = iFix Then
          DisplayReg = PndImmed
          Call ResetPndAll
          EngMode = True          'allow Engineering mode
        End If
      End If
      If Not EEMode Then
        EEMode = True             'update status if just turning on
        Call UpdateStatus
      End If
      DisplayText = False         'force number entry
      Call DisplayLine            'display contents of DisplayReg
      AllowExp = True             'enable adding Exponent
'-------------------------------------------------------------------------------
    Case 180  ' Sbr
      If BaseType = TypHex Then
        Call CheckValue("C")
      Else
        CmdNotActive
      End If
'-------------------------------------------------------------------------------
    Case 184  ' x
      Call Pend(iMult)
'-------------------------------------------------------------------------------
    Case iRem, iRem2 ' Rem and [']
      If CBool(PndIdx) Then
        Call CheckPnd(0)                      'flush out pending numeric input
        If ErrorFlag Then Exit Sub            'if error encountered, then done
      End If
      
      DspTxt = vbNullString                   'initialize pending text
      CharCount = 0                           'init character counte
      If LastTypedInstr = iRem Then
        REMmode = 1 'use "Rem "
        CharLimit = DisplayWidth - 4          'set character limit to width less "Rem "
      Else
        REMmode = 2 'use "'"
        CharLimit = DisplayWidth - 1          'set character limit to width less "'"
      End If
      
      AllowSpace = True                       'allow typing of a space
      With frmVisualCalc
        Call .checkTextEntry(True)            'set up keyboard
        With .lstDisplay
          .List(.ListIndex) = vbNullString 'init display line
        End With
      End With
      SetTip vbNullString                     'clear tip field
'-------------------------------------------------------------------------------
    Case 186  ' Oct
      If BaseType <> TypOct Then                                'ignore if already type
        If BaseType = TypDec Then DisplayReg = Fix(DisplayReg)  'remove decimal
        BaseType = TypOct                                       'set type
        Call UpdateStatus                                       'display it in status
        Call EnableNums                                         'enable 0-7 only
        Call DisplayLine                                        'update display contents
      End If
'-------------------------------------------------------------------------------
    Case 187  ' ~
      Call Pend(iNotB)
'-------------------------------------------------------------------------------
    Case 188  ' Select  'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 189  ' Case    'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 190  ' {       'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 191  ' }       'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 192  ' Deg
      AngleType = TypDeg
      Call UpdateStatus
'-------------------------------------------------------------------------------
    Case 193  ' Lbl
      If BaseType = TypHex Then
        Call CheckValue("D")
      End If
'-------------------------------------------------------------------------------
    Case 197  ' -
      Call Pend(iMinus)
'-------------------------------------------------------------------------------
    Case 198  ' Beep
      Beep
'-------------------------------------------------------------------------------
    Case 199  ' Bin
      If BaseType <> TypBin Then                                'ignore if already type
        If BaseType = TypDec Then DisplayReg = Fix(DisplayReg)  'remove decimal
        BaseType = TypBin                                       'set type
        Call UpdateStatus                                       'display it in status
        Call EnableNums                                         'enable 0-7 only
        Call DisplayLine                                        'update display contents
      End If
'-------------------------------------------------------------------------------
    Case 200  ' ^
      Call Pend(iXorB)
'-------------------------------------------------------------------------------
    Case 201  ' For   'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 202  ' Do    'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 203  ' While 'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
'    Case 204  ' Pmt 'handled by pending process
'-------------------------------------------------------------------------------
    Case 205  ' Rad
      AngleType = TypRad
      Call UpdateStatus
'-------------------------------------------------------------------------------
    Case 206  ' UKey
      If BaseType = TypHex Then
        Call CheckValue("E")
      End If
'-------------------------------------------------------------------------------
    Case 210  ' +
      Call Pend(iAdd)
'-------------------------------------------------------------------------------
    Case 211  ' Plot
      With frmVisualCalc.PicPlot
        .Visible = Not .Visible
      End With
      PlayWavResource "Swoosh"
'-------------------------------------------------------------------------------
'    Case 212  ' Nvar    'handled by pending operations
'-------------------------------------------------------------------------------
    Case 213  ' %
      Call Pend(iMod)
'-------------------------------------------------------------------------------
    Case 214  ' If    'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 215  ' Else  'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 216  ' Cont  'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 217  ' Break 'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 218  ' Grad
      AngleType = TypGrad
      Call UpdateStatus
'-------------------------------------------------------------------------------
    Case 219  ' R/S   'start and stop execution
      If BaseType = TypHex Then
        Call CheckValue("F")
      Else
        PmtFlag = False
        If Not StopMode Then                  'if Stop instruction not active...
          If Not (Preprocessd Or CBool(ActivePgm)) Then
            Call Preprocess                   'process the code if Pgm00 and not Preprocessed
            If Not Preprocessd Then Exit Sub  'error
          End If
          With frmVisualCalc.lstDisplay       'erase current line
            .List(.ListIndex) = vbNullString
          End With
          Call Run                            'run program
' if after running, the Pmt command is activated, we will turn on the TXT mode (TextEntry),
' let the user type in a response, and by pressing TXT or '=' (ENTER), the program will continue
' (see KybdMain).
          If PmtFlag Then                     'if user prompting turned on...
            LastTypedInstr = iTXT             'set TXT command
            Call ActiveKeypad                 'activate it (TXT or [=] (ENTER)) will start R/S cmd
          Else
            Call DisplayLine                  'else terminating run...
          End If
'---------------
        End If
      End If
'-------------------------------------------------------------------------------
    Case 221  ' .
      If Not ValueDec Then CheckValue (".")
      ValueDec = True
      ValueTyped = True                                               'something was typed
'-------------------------------------------------------------------------------
    Case 222  ' +/-
      If CBool(Len(ValueAccum)) Then                                  'if accumulator contains data...
        If Val(ValueAccum) <> 0# Then                                 'and it is not zero...
          If AllowExp Then                                            'EE mode and EXP mod allowed
            With frmVisualCalc.lstDisplay
              ValueAccum = Trim$(.List(.ListIndex))                   'grab line of data
              If Left$(ValueAccum, 1) = "-" Then ValueAccum = Mid$(ValueAccum, 2)
              S = Mid$(ValueAccum, Len(ValueAccum) - 2, 1)            'grab sign
              If S = "+" Then                                         'flip Exponent sign
                S = "-"
              Else
                S = "+"
              End If
              Mid$(ValueAccum, Len(ValueAccum) - 2, 1) = S            'save update
            End With
          Else
            If ValueSgn = vbNullString Then                           'if positive
              ValueSgn = "-"                                          'make negative
            Else
              ValueSgn = vbNullString                                 'else we are now positive
            End If
          End If
          ValueTyped = True                                           'something was typed
          Call DisplayAccum                                           'display accumulator value
        End If
      Else  'we are flipping a value that cannot otherwise be numerically added to...
        With frmVisualCalc.lstDisplay
          S = LTrim$(.List(.ListIndex))                               'get current line data
          If Val(S) <> 0# Then                                        'if value is not 0...
            If Left$(S, 1) = "-" Then                                 'if sign is present
              S = Mid$(S, 2)                                          'remove it
            Else
              S = "-" & S                                             'else add it
            End If
            DisplayReg = Val(S)                                       'update immediate register
            .List(.ListIndex) = String$(DisplayWidth - Len(S), 32) & S 'display on current line
          End If
        End With
      End If
'-------------------------------------------------------------------------------
    Case 223  ' =
      AllowExp = False          'terminating calculations, so do not allow altering exponent
      Call Pend(iEqual)
'-------------------------------------------------------------------------------
    Case 224  ' Print 'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
'    Case 225  ' Tvar  'handled by pending operations
'    Case 226  ' >>  'handled by pending operations
'-------------------------------------------------------------------------------
    Case 227  ' y^
      Call Pend(iPower)
'-------------------------------------------------------------------------------
    Case 228  ' X²
      On Error Resume Next
      TV = DisplayReg * DisplayReg
      Call CheckError
      If ErrorFlag Then Exit Sub
      DisplayReg = TV
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 229  ' Pi
      DisplayReg = vPi
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 230  ' Rnd
      DisplayReg = Rnd()
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 231  ' Mil
      AngleType = TypMil
      Call UpdateStatus
'-------------------------------------------------------------------------------
    Case 232  ' Pvt     'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 233  ' Const   'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 234  ' Struct  'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 235  ' NxtLbl  'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 236  ' PrvLbl  'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 237  ' Line    'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 238  ' [       'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 239  ' ]       'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
'    Case 240  ' ClrVar  'handled by pre-processor
'-------------------------------------------------------------------------------
    Case 241  ' SzOf    'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 242  ' Def     'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 243  ' IfDef   'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 244  ' Edef    'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
              '--2nd keys---------------------------------------------
'-------------------------------------------------------------------------------
'    Case 257  ' MDL   'handled by pending process
'-------------------------------------------------------------------------------
    Case 258  ' CMM
      Erase ModMem                    'erase module memory
      Erase ModMap
      Erase ModLblMap
      Erase ModStPl
      Erase ModLbls
      frmVisualCalc.sbrImmediate.Panels("MDL").ToolTipText = "Currently loaded Module"
      
      ModSize = 0                     'make null size
      ModCnt = 0                      'no modules
      ModLblCnt = 0
      ModStCnt = 0
      ModLocked = False
      ModName = 0                     'reset module name
      SaveSetting App.Title, "Settings", "LoadedMDL", "0"
      DisplayReg = 0#
      
      If CBool(ActivePgm) Then
        ActivePgm = 0                 'reset active program to 00
        InstrErr = 0
        Call Reset_Support
        Call UpdateStatus
        Call RedoAlphaPad
      Else
        Call UpdateStatus
      End If
      DisplayMsg "Module memory space initialized"
'-------------------------------------------------------------------------------
'    Case 259  ' Lapp  'handled by pending process
'    Case 260  ' ASCII 'handled by pending process
'-------------------------------------------------------------------------------
    Case 261  ' CMs
      Call CMs_Support
'-------------------------------------------------------------------------------
    Case 262  ' CP
      Call CP_Support
'-------------------------------------------------------------------------------
'    Case iList' List    'handled by pre-processor
'-------------------------------------------------------------------------------
    Case 264  ' BST
      Call Backstep
'-------------------------------------------------------------------------------
    Case 265  ' DEL      'not supported in Non-LRN mode
      Call Del_Support
'-------------------------------------------------------------------------------
'    Case 266  ' Paste   'handled by preprocessor
'-------------------------------------------------------------------------------
    Case iUSR  ' USR     'Handled by pending process
'-------------------------------------------------------------------------------
    Case 268  ' RtoP
      TV = Sqr(DisplayReg * DisplayReg + TestReg * TestReg) 'get distance
      If TestReg = 0# Then                                  'if x is 0
        DisplayReg = 0#                                     'then angle is 0
      Else
        DisplayReg = RadToAng(Atn(DisplayReg / TestReg))    'else get angle
      End If
      TestReg = TV
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 269  ' Push
      PushIdx = PushIdx + 1
      If PushIdx > PushSize Then
        PushSize = PushSize + PushLimit
        ReDim Preserve PushValues(PushSize)
      End If
      PushValues(PushIdx) = DisplayReg
      Call UpdateStatus
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 270  ' Pop
      If CBool(PushIdx) Then
        DisplayReg = PushValues(PushIdx)
        PushIdx = PushIdx - 1
        Call UpdateStatus
        Call DisplayLine
      Else
        ForcError "Push Stack was empty"
      End If
'-------------------------------------------------------------------------------
    Case 271  ' StkEx
      If CBool(PushIdx) Then
        TV = DisplayReg
        DisplayReg = PushValues(PushIdx)
        PushValues(PushIdx) = TV
        Call DisplayLine
      Else
        ForcError "Push Stack was empty"
      End If
'-------------------------------------------------------------------------------
'    Case 272  ' SUB 'handled by muti-command processor
'    Case 273  ' DIV 'handled by muti-command processor
'-------------------------------------------------------------------------------
    Case 274  ' < 'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 275  ' > 'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
'    Case 276  ' Skey  'handled by pre-processor
'-------------------------------------------------------------------------------
    Case 277  ' eX
      DisplayReg = Exp(DisplayReg)
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 278  ' E-
      Call StatSUMsub
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 279  ' StDev
      Call StatStdDev
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 280  ' Varnc
      Call StatVarnc
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 281  ' Yint
      Call Yintercept
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case iArc  ' Arc 'handled by multi-command processor
      If BaseType = TypHex Then
        Call CheckValue("A")
      End If
'-------------------------------------------------------------------------------
    Case 283  ' Sec
      On Error Resume Next
      DisplayReg = 1# / Cos(AngToRad(DisplayReg))
      Call CheckError
      On Error GoTo 0
      If Not ErrorFlag Then Call DisplayLine
'-------------------------------------------------------------------------------
    Case 284  ' Csc (Cosec)
      On Error Resume Next
      DisplayReg = 1# / Sin(AngToRad(DisplayReg))
      Call CheckError
      On Error GoTo 0
      If Not ErrorFlag Then Call DisplayLine
'-------------------------------------------------------------------------------
    Case 285  ' Cot (Cotan)
      On Error Resume Next
      DisplayReg = 1# / Tan(AngToRad(DisplayReg))
      Call CheckError
      On Error GoTo 0
      If Not ErrorFlag Then Call DisplayLine
'-------------------------------------------------------------------------------
    Case 286  ' LogX
      Call Pend(iLogX)
'-------------------------------------------------------------------------------
'    Case 287  ' Var   'handled by pending process
'-------------------------------------------------------------------------------
    Case 288  ' ==
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 289  ' &&
      CmdNotActive
'-------------------------------------------------------------------------------
'    Case 290  ' RFlg  'Handled by pending process
'-------------------------------------------------------------------------------
    Case 291  ' !Flg  'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 292  ' X!=T  'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 293  ' X<=T  'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 294  ' X<T   'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 295  ' Dfn   'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case iSemiC  ' ;     'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 297  ' Log
      On Error Resume Next
      DisplayReg = Log(DisplayReg) / Log(10#)        'Common logarithm
      Call CheckError
      On Error GoTo 0
      If Not ErrorFlag Then Call DisplayLine
'-------------------------------------------------------------------------------
    Case 298  ' 10^
      On Error Resume Next
      DisplayReg = Exp(DisplayReg * Log(10#))        '10 to power of DisplayReg
      Call CheckError
      On Error GoTo 0
      If Not ErrorFlag Then Call DisplayLine
'-------------------------------------------------------------------------------
    Case 299  ' /=
      Call Pend(iDivEq)
'-------------------------------------------------------------------------------
    Case 300  ' Fmt
      If CBool(PndIdx) Then
        Call CheckPnd(0)                      'flush out pending numeric input
        If ErrorFlag Then Exit Sub            'if error encountered, then done
      End If
      
      DspTxt = vbNullString                   'initialize pending text
      CharCount = 0                           'init character counte
      CharLimit = DisplayWidth                'set character limit to width
      REMmode = 3                             'borrow REMmode flag
      AllowSpace = True                       'allow typing of a space
      With frmVisualCalc
        Call .checkTextEntry(True)            'set up keyboard
        With .lstDisplay
          .List(.ListIndex) = vbNullString 'init display line
        End With
      End With
      SetTip vbNullString                     'clear tip field
'-------------------------------------------------------------------------------
    Case 301  ' !=   'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 302  ' ||   'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 303  ' Frac
      DisplayReg = DisplayReg - Fix(DisplayReg)
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 304  ' Sgn
      DisplayReg = CDbl(Sgn(DisplayReg))
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 305  ' !Fix
      DspFmtFix = -1              'disable fixed-decimal places
      DspFmt = DefDspFmt          'set default format
      ScientifEE = DefScientific  'reset default scientif mode
      Call DisplayLine            'update display
'-------------------------------------------------------------------------------
    Case 306  ' D.ddd Convert DDD.dddd to DDD.MMSSddd
      vDeg = Fix(DisplayReg)                            'get DDD
      DisplayReg = (DisplayReg - vDeg) * 3600#          'get seconds
      vMin = Fix(DisplayReg / 60#)                      'get MM
      vSec = (DisplayReg - vMin * 60#)                  'get SS.dddd
      DisplayReg = vDeg + vMin / 100# + vSec / 10000#   'get dd.ddddd
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 307  ' !EE
      If EEMode Then                                    'if EE mode was on...
        EEMode = False                                  'turn it off
        EngMode = False                                 'and engineering mode
        Call UpdateStatus                               'reflect it on the status
        With frmVisualCalc.lstDisplay
          DisplayReg = Val(Trim$(.List(.ListIndex)))  'correct any round-off errors
        End With
        Call DisplayLine                                'display new data
      End If
      AllowExp = False                                  'ensure exponent addition is disabled
'-------------------------------------------------------------------------------
'    Case 308  ' Call  'handled by pending process
'    Case 309  ' Trim  'handled by pending process
'    Case 310  ' LTrim 'handled by pending process
'    Case 311  ' RTrim 'handled by pending process
'-------------------------------------------------------------------------------
    Case 312  ' *=
      Call Pend(iMulEq)
'-------------------------------------------------------------------------------
'    Case iRem2  ' [']     'handled by preprocessor
'-------------------------------------------------------------------------------
    Case 314  ' >=    'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 315  ' !     'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 316  ' Open  'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 317  ' Close 'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 318  ' Read  'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 319  ' Write 'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
'    Case 320  ' Swap  'handled by pending process
'    Case 321  ' GTO   'handled by pending process
'-------------------------------------------------------------------------------
    Case 322  ' LOF    'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 323  ' Get     'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 324  ' Put     'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 325  ' -=
      Call Pend(iSubEq)
'-------------------------------------------------------------------------------
'    Case 326  ' sysBP 'handled by key processor
'-------------------------------------------------------------------------------
    Case 327  ' <=      'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 328  ' Nor     'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
'    Case 329  ' Inc 'handled by pend processor
'    Case 330  ' Dec 'handled by pend processor
'-------------------------------------------------------------------------------
    Case 331  ' Dsz     'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 332  ' Dsnz    'ignored
      CmdNotActive
'    Case 333  ' All  'handled by pending processor
'-------------------------------------------------------------------------------
    Case 334  ' Rtn     'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 335  ' LSet    'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 336  ' RSet    'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 337  ' Printf  'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 338  ' +=
      Call Pend(iAddEq)
'-------------------------------------------------------------------------------
    Case 339  ' RGB   'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
'    Case 340  ' Ivar    'handled by pending operations
'-------------------------------------------------------------------------------
    Case 341  ' \
      Call Pend(iBkSlsh)
'-------------------------------------------------------------------------------
    Case 342  ' As      'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 343  ' ElseIf  'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 344  ' DBG     'toggle trace mode
      TraceFlag = Not TraceFlag
      frmVisualCalc.mnuFileTron.Checked = TraceFlag
      Call UpdateStatus
'-------------------------------------------------------------------------------
    Case 345  ' Gfree
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 346  ' Len     'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 347  ' STOP    'Do cont continue execution with R/S
      If RunMode Then                       'if currently running...
        StopMode = True                     'turn on Stop mode (prevents R/S from working...)
        RunMode = False                     'disable mode
        Call ResetPnd                       'reset data
        Call UpdateStatus
        Call DisplayLine                    'show display register
      End If
'-------------------------------------------------------------------------------
    Case 348  ' With 'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
'    Case 349  ' , 'handled by preprocessor
'-------------------------------------------------------------------------------
    Case 350  ' Val
      DisplayReg = Val(DspTxt)                  'derive value from displayed data
      CharLimit = 0                             'turn off character limit so Display line will work
      Call DisplayLine                          'display updated value
'-------------------------------------------------------------------------------
    Case 351  ' Adv
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
'-------------------------------------------------------------------------------
    Case 352  ' Print;  'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
'    Case 353  ' Cvar    'handled by pending operations
'    Case 354  ' <<  'handled by pending operations
'-------------------------------------------------------------------------------
    Case 355  ' Root
      Call Pend(iRoot)
'-------------------------------------------------------------------------------
    Case 356  ' Sqrt
      On Error Resume Next
      DisplayReg = Sqr(DisplayReg)  'get square root
      Call CheckError               'check for error
      On Error GoTo 0               'reset error processing
      Call DisplayLine              'display data
'-------------------------------------------------------------------------------
    Case 357  ' e
      DisplayReg = vE
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 358  ' Rnd#
      DisplayReg = Rnd(DisplayReg)
      Randomize
      Call DisplayLine              'display data
'-------------------------------------------------------------------------------
    Case 359  ' Until   'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 360  ' Pub     'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 361  ' Enum    'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 362  ' AdrOf   'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 363  ' Pcmp
      If InstrCnt = 0 Then
        ForcError "No program to process"
      Else
        Preprocessd = False                   'make sure flag is turned off
        Compressd = False
        Call Preprocess                       'run the pre-processor
      End If
'-------------------------------------------------------------------------------
    Case 364  ' Comp
      If InstrCnt = 0 Then
        ForcError "No program to Compress"
      Else
        Compressd = False                      'make sure flag is turned off
        Call Compress                          'then Compress it
      End If
'-------------------------------------------------------------------------------
    Case 365  ' Circle
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 366  ' Split   'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 367  ' Join    'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 368  ' ReDim   'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 369  ' Mid     'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 370  ' Udef    'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 371  ' !Def    'ignored
      CmdNotActive
'-------------------------------------------------------------------------------
    Case 372  ' Delse   'ignored
      CmdNotActive

'-------------------------------------------------------------------------------
' EXTENDED FUNCTIONS
'-------------------------------------------------------------------------------
    Case 400  ' Asin
      X = DisplayReg
      On Error Resume Next
      DisplayReg = RadToAng(Atn(X / Sqr(-X * X + 1#)))
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 401  ' Acos
      X = DisplayReg
      On Error Resume Next
      DisplayReg = RadToAng(Atn(-X / Sqr(-X * X + 1#)) + vRA)
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 402  ' Atan
      On Error Resume Next
      DisplayReg = RadToAng(Atn(DisplayReg))
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 403  ' SinH
      X = DisplayReg
      On Error Resume Next
      DisplayReg = (Exp(X) - Exp(-X)) / 2#
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 404  ' CosH
      X = DisplayReg
      On Error Resume Next
      DisplayReg = (Exp(X) + Exp(-X)) / 2#
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 405  ' TanH
      X = DisplayReg
      On Error Resume Next
      DisplayReg = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 406  ' ArcSinH
      X = DisplayReg
      On Error Resume Next
      DisplayReg = Log(X + Sqr(X * X + 1#))
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 407  ' ArcCosH
      X = DisplayReg
      On Error Resume Next
      DisplayReg = Log(X + Sqr(X * X - 1#))
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 408  ' ArcTanH
      X = DisplayReg
      On Error Resume Next
      DisplayReg = Log((1# + X) / (1# - X)) / 2#
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 409  ' Asec      'Acos(1/x)
      On Error Resume Next
      X = 1# / DisplayReg
      DisplayReg = RadToAng(Atn(-X / Sqr(-X * X + 1#)) + vRA)
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 410  ' Acsc      'Asin(1/x)
      On Error Resume Next
      X = 1# / DisplayReg
      DisplayReg = RadToAng(Atn(X / Sqr(-X * X + 1#)))
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 411  ' Acot      'Atan(1/x)
      On Error Resume Next
      DisplayReg = RadToAng(Atn(1# / DisplayReg))
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 412  ' SecH      '1/CosH(x)
      X = DisplayReg
      On Error Resume Next
      DisplayReg = 1# / ((Exp(X) + Exp(-X)) / 2#)
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 413  ' CscH      '1/SinH(x)
      X = DisplayReg
      On Error Resume Next
      DisplayReg = 1# / ((Exp(X) - Exp(-X)) / 2#)
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 414  ' CotH      '1/TanH(x)
      X = DisplayReg
      On Error Resume Next
      DisplayReg = 1# / ((Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X)))
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 415  ' ArcSecH   '1/ArcCosH(x)
      X = DisplayReg
      On Error Resume Next
      DisplayReg = 1# / Log(X + Sqr(X * X - 1#))
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 416  ' ArcCscH   '1/ArcSinH(x)
      X = DisplayReg
      On Error Resume Next
      DisplayReg = 1# / Log(X + Sqr(X * X + 1#))
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
    Case 417  ' ArcCotH   '1/ArcTanH(x)
      X = DisplayReg
      On Error Resume Next
      DisplayReg = 1# / (Log((1# + X) / (1# - X)) / 2#)
      Call CheckError               'check for error
      On Error GoTo 0
      Call DisplayLine
'-------------------------------------------------------------------------------
  End Select
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

