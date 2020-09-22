Attribute VB_Name = "modToggleKeypad"
Option Explicit

'*******************************************************************************
' Subroutine Name   : ShowKeypad
' Purpose           : Toggle 2nd key displays on main control keypad
'*******************************************************************************
Public Sub ShowKeypad()
  Dim Idx As Integer
  
  With frmVisualCalc
    If Not Key2nd Then                'if 2nd key IS NOT pressed
      .cmdKeyPad(1).Caption = "LRN"
      .cmdKeyPad(2).Caption = "Pgm"
      .cmdKeyPad(3).Caption = "Load"
      .cmdKeyPad(4).Caption = "Save"
      .cmdKeyPad(5).Caption = "CE"
      .cmdKeyPad(6).Caption = "CLR"
      .cmdKeyPad(7).Caption = "OP"
      .cmdKeyPad(8).Caption = "SST"
      .cmdKeyPad(9).Caption = "INS"
      .cmdKeyPad(10).Caption = "Cut"
      .cmdKeyPad(11).Caption = "Copy"
      .cmdKeyPad(12).Caption = "PtoR"
      .cmdKeyPad(13).Caption = "STO"
      .cmdKeyPad(14).Caption = "RCL"
      .cmdKeyPad(15).Caption = "EXC"
      .cmdKeyPad(16).Caption = "SUM"
      .cmdKeyPad(17).Caption = "MUL"
      .cmdKeyPad(18).Caption = "IND"
      .cmdKeyPad(19).Caption = "Reset"
      .cmdKeyPad(20).Caption = "Hkey"
      .cmdKeyPad(21).Caption = "lnX"
      .cmdKeyPad(22).Caption = "å+" 'E+
      .cmdKeyPad(23).Caption = "Mean"
      .cmdKeyPad(24).Caption = "X!"
      .cmdKeyPad(25).Caption = "X><T"
      
      If BaseType = TypHex And Not RunMode And Not CBool(MRunMode) Then
        .cmdKeyPad(26).Caption = "A"
        .cmdKeyPad(26).Enabled = Not TextEntry
      Else
        .cmdKeyPad(26).Caption = "Arc"
      End If
      
      .cmdKeyPad(27).Caption = "Sin"
      .cmdKeyPad(28).Caption = "Cos"
      .cmdKeyPad(29).Caption = "Tan"
      .cmdKeyPad(30).Caption = "1/x"
      .cmdKeyPad(31).Caption = "Txt"
      .cmdKeyPad(32).Caption = "Hex"
      .cmdKeyPad(33).Caption = "&&" '&
      .cmdKeyPad(34).Caption = "StFlg"
      .cmdKeyPad(35).Caption = "IfFlg"
      .cmdKeyPad(36).Caption = "X==T"
      .cmdKeyPad(37).Caption = "X>=T"
      .cmdKeyPad(38).Caption = "X>T"
      
      If BaseType = TypHex And Not RunMode And Not CBool(MRunMode) Then
        .cmdKeyPad(39).Caption = "B"
        .cmdKeyPad(39).Enabled = Not TextEntry
      Else
        .cmdKeyPad(39).Caption = "Dfn"
      End If
      
      .cmdKeyPad(40).Caption = ";"
      .cmdKeyPad(41).Caption = "("
      .cmdKeyPad(42).Caption = ")"
      .cmdKeyPad(43).Caption = "÷"
      .cmdKeyPad(44).Caption = "Style"
      .cmdKeyPad(45).Caption = "Dec"
      .cmdKeyPad(46).Caption = "|"
      .cmdKeyPad(47).Caption = "Int"
      .cmdKeyPad(48).Caption = "Abs"
      .cmdKeyPad(49).Caption = "Fix"
      .cmdKeyPad(50).Caption = "D.MS"
      .cmdKeyPad(51).Caption = "EE"
      
      If BaseType = TypHex And Not RunMode And Not CBool(MRunMode) Then
        .cmdKeyPad(52).Caption = "C"
        .cmdKeyPad(52).Enabled = Not TextEntry
      Else
        .cmdKeyPad(52).Caption = "Sbr"
      End If
      
      .cmdKeyPad(53).Caption = "7"
      .cmdKeyPad(54).Caption = "8"
      .cmdKeyPad(55).Caption = "9"
      .cmdKeyPad(56).Caption = "x"
      .cmdKeyPad(57).Caption = "'"
      .cmdKeyPad(58).Caption = "Oct"
      .cmdKeyPad(59).Caption = "~"
      .cmdKeyPad(60).Caption = "Select"
      .cmdKeyPad(61).Caption = "Case"
      .cmdKeyPad(62).Caption = "{"
      .cmdKeyPad(63).Caption = "}"
      .cmdKeyPad(64).Caption = "Deg"
      
      If BaseType = TypHex And Not RunMode And Not CBool(MRunMode) Then
        .cmdKeyPad(65).Caption = "D"
        .cmdKeyPad(65).Enabled = Not TextEntry
      Else
        .cmdKeyPad(65).Caption = "Lbl"
      End If
      
      .cmdKeyPad(66).Caption = "4"
      .cmdKeyPad(67).Caption = "5"
      .cmdKeyPad(68).Caption = "6"
      .cmdKeyPad(69).Caption = "¾"  ' -
      .cmdKeyPad(70).Caption = "Beep"
      .cmdKeyPad(71).Caption = "Bin"
      .cmdKeyPad(72).Caption = "^"
      .cmdKeyPad(73).Caption = "For"
      .cmdKeyPad(74).Caption = "Do"
      .cmdKeyPad(75).Caption = "While"
      .cmdKeyPad(76).Caption = "Pmt"
      .cmdKeyPad(77).Caption = "Rad"
      
      If BaseType = TypHex And Not RunMode And Not CBool(MRunMode) Then
        .cmdKeyPad(78).Caption = "E"
        .cmdKeyPad(78).Enabled = Not TextEntry
      Else
        .cmdKeyPad(78).Caption = "Ukey"
      End If
      
      .cmdKeyPad(79).Caption = "1"
      .cmdKeyPad(80).Caption = "2"
      .cmdKeyPad(81).Caption = "3"
      .cmdKeyPad(82).Caption = "+"
      .cmdKeyPad(83).Caption = "Plot"
      .cmdKeyPad(84).Caption = "Nvar"
      .cmdKeyPad(85).Caption = "%"
      .cmdKeyPad(86).Caption = "If"
      .cmdKeyPad(87).Caption = "Else"
      .cmdKeyPad(88).Caption = "Cont"
      .cmdKeyPad(89).Caption = "Break"
      .cmdKeyPad(90).Caption = "Grad"
      
      If BaseType = TypHex And Not RunMode And Not CBool(MRunMode) Then
        .cmdKeyPad(91).Caption = "F"
        .cmdKeyPad(91).Enabled = Not TextEntry
      Else
        .cmdKeyPad(91).Caption = "R/S"
      End If
      
      .cmdKeyPad(92).Caption = "0"
      .cmdKeyPad(93).Caption = "."
      .cmdKeyPad(94).Caption = "+/-"
      .cmdKeyPad(95).Caption = "="
      .cmdKeyPad(96).Caption = "Print"
      .cmdKeyPad(97).Caption = "Tvar"
      .cmdKeyPad(98).Caption = ">>"
      .cmdKeyPad(99).Caption = "y^"
      .cmdKeyPad(100).Caption = "X²"  'x2
      .cmdKeyPad(101).Caption = "p"   'Pi
      .cmdKeyPad(102).Caption = "Rnd"
      .cmdKeyPad(103).Caption = "Mil"
      .cmdKeyPad(104).Caption = "Pvt"
      .cmdKeyPad(105).Caption = "Const"
      .cmdKeyPad(106).Caption = "Struct"
      .cmdKeyPad(107).Caption = "NxLbl"
      .cmdKeyPad(108).Caption = "PvLbl"
      .cmdKeyPad(109).Caption = "Line"
      .cmdKeyPad(110).Caption = "["
      .cmdKeyPad(111).Caption = "]"
      .cmdKeyPad(112).Caption = "ClrVar"
      .cmdKeyPad(113).Caption = "SzOf"
      .cmdKeyPad(114).Caption = "Def"
      .cmdKeyPad(115).Caption = "IfDef"
      .cmdKeyPad(116).Caption = "Edef"
      
      'Handle 0-9 keys
      If Not .cmdKeyPad(1).Enabled Then
        .cmdKeyPad(53).Enabled = True  '7
        .cmdKeyPad(54).Enabled = True  '8
        .cmdKeyPad(55).Enabled = True  '9
        .cmdKeyPad(66).Enabled = True  '4
        .cmdKeyPad(67).Enabled = True  '5
        .cmdKeyPad(68).Enabled = True  '6
        .cmdKeyPad(79).Enabled = True  '1
        .cmdKeyPad(80).Enabled = True  '2
        .cmdKeyPad(81).Enabled = True  '3
        .cmdKeyPad(92).Enabled = True  '0
      End If
      
      'handle alpha/user-defined keys
      If TextEntry Then     'if text entry mode...
        Call ResetAlphaPad  'Reset keyboard pad to A-Z values for text entry
      Else
        Call RedoAlphaPad   'Reset AlphaPad to user-defined keys
      End If
'---------------------------------------------------------------------
    Else                      'process 2nd key text assignments
'---------------------------------------------------------------------
      For Idx = 53 To 94                                      'range containing numeric keypad
        Select Case Idx
          Case 53, 54, 55, 66, 67, 68, 79, 80, 81, 92         'ignore numeric keypad keys for disablement
            .cmdKeyPad(Idx).Enabled = True                    'ensure keys are enabled
        End Select
      Next Idx
      .cmdKeyPad(1).Caption = "MDL"
      .cmdKeyPad(2).Caption = "CMM"
      .cmdKeyPad(3).Caption = "Lapp"
      .cmdKeyPad(4).Caption = "ASCII"
      .cmdKeyPad(5).Caption = "CMs"
      .cmdKeyPad(6).Caption = "CP"
      .cmdKeyPad(7).Caption = "USR"
      .cmdKeyPad(8).Caption = "BST"
      .cmdKeyPad(9).Caption = "DEL"
      .cmdKeyPad(10).Caption = "Paste"
      .cmdKeyPad(11).Caption = "List"
      .cmdKeyPad(12).Caption = "RtoP"
      .cmdKeyPad(13).Caption = "Push"
      .cmdKeyPad(14).Caption = "Pop"
      .cmdKeyPad(15).Caption = "StkEx"
      .cmdKeyPad(16).Caption = "SUB"
      .cmdKeyPad(17).Caption = "DIV"
      .cmdKeyPad(18).Caption = "<"
      .cmdKeyPad(19).Caption = ">"
      .cmdKeyPad(20).Caption = "Skey"
      .cmdKeyPad(21).Caption = "eX"
      .cmdKeyPad(22).Caption = "å-" 'E-
      .cmdKeyPad(23).Caption = "StDev"
      .cmdKeyPad(24).Caption = "Varnc"
      .cmdKeyPad(25).Caption = "Yint"
      
      .cmdKeyPad(26).Caption = "Hyp"
      If BaseType <> TypDec Then
        .cmdKeyPad(26).Enabled = False
      End If
      
      .cmdKeyPad(27).Caption = "Sec"
      .cmdKeyPad(28).Caption = "Csc"
      .cmdKeyPad(29).Caption = "Cot"
      .cmdKeyPad(30).Caption = "LogX"
      .cmdKeyPad(31).Caption = "Var"
      
      .cmdKeyPad(32).Caption = "=="
      If BaseType <> TypDec Then
        .cmdKeyPad(32).Enabled = False
      End If
      
      .cmdKeyPad(33).Caption = "&&&&" '&&
      If BaseType <> TypDec Then
        .cmdKeyPad(33).Enabled = False
      End If
      
      .cmdKeyPad(34).Caption = "RFlg"
      .cmdKeyPad(35).Caption = "!Flg"
      .cmdKeyPad(36).Caption = "X!=T"
      .cmdKeyPad(37).Caption = "X<=T"
      .cmdKeyPad(38).Caption = "X<T"
      
      .cmdKeyPad(39).Caption = "NOP"
      If BaseType <> TypDec Then
        .cmdKeyPad(39).Enabled = False
      End If
      
      .cmdKeyPad(40).Caption = ":"
      
      .cmdKeyPad(41).Caption = "Log"
      If BaseType <> TypDec Then
        .cmdKeyPad(41).Enabled = False
      End If
      
      .cmdKeyPad(42).Caption = "10^"
      If BaseType <> TypDec Then
        .cmdKeyPad(42).Enabled = False
      End If
      
      .cmdKeyPad(43).Caption = "÷="
      If BaseType <> TypDec Then
        .cmdKeyPad(43).Enabled = False
      End If
      
      .cmdKeyPad(44).Caption = "Fmt"
      
      .cmdKeyPad(45).Caption = "!="
      If BaseType <> TypDec Then
        .cmdKeyPad(45).Enabled = False
      End If
      
      .cmdKeyPad(46).Caption = "||"
      If BaseType <> TypDec Then
        .cmdKeyPad(46).Enabled = False
      End If
      
      .cmdKeyPad(47).Caption = "Frac"
      .cmdKeyPad(48).Caption = "Sgn"
      .cmdKeyPad(49).Caption = "!Fix"
      .cmdKeyPad(50).Caption = "D.ddd"
      .cmdKeyPad(51).Caption = "!EE"
      
      .cmdKeyPad(52).Caption = "Call"
      If BaseType <> TypDec Then
        .cmdKeyPad(52).Enabled = False
      End If
      
      .cmdKeyPad(53).Caption = "Trim"
      .cmdKeyPad(54).Caption = "LTrim"
      .cmdKeyPad(55).Caption = "RTrim"
      
      .cmdKeyPad(56).Caption = "x="
      If BaseType <> TypDec Then
        .cmdKeyPad(56).Enabled = False
      End If
      
      .cmdKeyPad(57).Caption = "Rem"
      
      .cmdKeyPad(58).Caption = ">="
      If BaseType <> TypDec Then
        .cmdKeyPad(58).Enabled = False
      End If
      
      .cmdKeyPad(59).Caption = "!"
      If BaseType <> TypDec Then
        .cmdKeyPad(59).Enabled = False
      End If
      
      .cmdKeyPad(60).Caption = "Open"
      .cmdKeyPad(61).Caption = "Close"
      .cmdKeyPad(62).Caption = "Read"
      .cmdKeyPad(63).Caption = "Write"
      .cmdKeyPad(64).Caption = "Swap"
      
      .cmdKeyPad(65).Caption = "GoTo"
      If BaseType <> TypDec Then
        .cmdKeyPad(65).Enabled = False
      End If
      
      .cmdKeyPad(66).Caption = "LOF"
      .cmdKeyPad(67).Caption = "Get"
      .cmdKeyPad(68).Caption = "Put"
      
      .cmdKeyPad(69).Caption = "-="
      If BaseType <> TypDec Then
        .cmdKeyPad(69).Enabled = False
      End If
      
      .cmdKeyPad(70).Caption = "sysBP"
      
      .cmdKeyPad(71).Caption = "<="
      If BaseType <> TypDec Then
        .cmdKeyPad(71).Enabled = False
      End If
      
      .cmdKeyPad(72).Caption = "Nor"
      If BaseType <> TypDec Then
        .cmdKeyPad(72).Enabled = False
      End If
      
      .cmdKeyPad(73).Caption = "Incr"
      .cmdKeyPad(74).Caption = "Decr"
      .cmdKeyPad(75).Caption = "Dsz"
      .cmdKeyPad(76).Caption = "Dsnz"
      .cmdKeyPad(77).Caption = "All"
      
      .cmdKeyPad(78).Caption = "Rtn"
      If BaseType <> TypDec Then
        .cmdKeyPad(78).Enabled = False
      End If
      
      .cmdKeyPad(79).Caption = "LSet"
      .cmdKeyPad(80).Caption = "RSet"
      .cmdKeyPad(81).Caption = "Printf"
      
      .cmdKeyPad(82).Caption = "+="
      If BaseType <> TypDec Then
        .cmdKeyPad(82).Enabled = False
      End If
      
      .cmdKeyPad(83).Caption = "RGB"
      .cmdKeyPad(84).Caption = "Ivar"
      
      .cmdKeyPad(85).Caption = "\"
      If BaseType <> TypDec Then
        .cmdKeyPad(85).Enabled = False
      End If
      
      .cmdKeyPad(86).Caption = "As"
      .cmdKeyPad(87).Caption = "ElseIf"
      .cmdKeyPad(88).Caption = "DBG"
      .cmdKeyPad(89).Caption = "Gfree"
      .cmdKeyPad(90).Caption = "Len"
      
      .cmdKeyPad(91).Caption = "Stop"
      If BaseType <> TypDec Then
        .cmdKeyPad(91).Enabled = False
      End If
      
      .cmdKeyPad(92).Caption = "With"
      
      .cmdKeyPad(93).Caption = ","
      If BaseType <> TypDec Then
        .cmdKeyPad(93).Enabled = False
      End If
      
      .cmdKeyPad(94).Caption = "Val"
      If BaseType <> TypDec Then
        .cmdKeyPad(94).Enabled = False
      End If
      
      .cmdKeyPad(95).Caption = "Adv"
      .cmdKeyPad(96).Caption = "Print;"
      .cmdKeyPad(97).Caption = "Cvar"
      .cmdKeyPad(98).Caption = "<<"
      .cmdKeyPad(99).Caption = "Root"
      .cmdKeyPad(100).Caption = "Sqrt"
      .cmdKeyPad(101).Caption = "e"
      .cmdKeyPad(102).Caption = "Rnd#"
      .cmdKeyPad(103).Caption = "Until"
      .cmdKeyPad(104).Caption = "Pub"
      .cmdKeyPad(105).Caption = "Enum"
      .cmdKeyPad(106).Caption = "AdrOf"
      .cmdKeyPad(107).Caption = "Pcmp"
      .cmdKeyPad(108).Caption = "Comp"
      .cmdKeyPad(109).Caption = "Circle"
      .cmdKeyPad(110).Caption = "Split"
      .cmdKeyPad(111).Caption = "Join"
      .cmdKeyPad(112).Caption = "ReDim"
      .cmdKeyPad(113).Caption = "Mid"
      .cmdKeyPad(114).Caption = "Udef"
      .cmdKeyPad(115).Caption = "!Def"
      .cmdKeyPad(116).Caption = "Delse"
      
      'Handle 0-9 keys, and disable if required
      If Not .cmdKeyPad(1).Enabled Then
        .cmdKeyPad(53).Enabled = False '7
        .cmdKeyPad(54).Enabled = False '8
        .cmdKeyPad(55).Enabled = False '9
        .cmdKeyPad(66).Enabled = False '4
        .cmdKeyPad(67).Enabled = False '5
        .cmdKeyPad(68).Enabled = False '6
        .cmdKeyPad(79).Enabled = False '1
        .cmdKeyPad(80).Enabled = False '2
        .cmdKeyPad(81).Enabled = False '3
        .cmdKeyPad(92).Enabled = False '0
      End If
      
      'handle alpha keys
      If TextEntry Then       'if text entry mode...
        Call ResetAlphaPad    'reset keys to alpha format for 2nd keys
      Else
        Call RedoAlphaPad     'else force A-Z and user-defined labels
      End If
    End If
  End With
  Call EnableNums           'enable digits on the numeric keypad, based on type
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

