Attribute VB_Name = "modTipCheck"
Option Explicit

Public TipAsText As Boolean 'true if TipText will be the only Destination
Public TipText As String    'hold tip text

'*******************************************************************************
' Subroutine Name   : SetTip
' Purpose           : Set data on the status bar
'*******************************************************************************
Public Sub SetTip(Txt As String)
  If TipText <> Txt Then
    TipText = Txt
    If Not TipAsText Then frmVisualCalc.sbrImmediate.Panels(1).Text = Txt
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : TipCheck
' Purpose           : Display Tip for a button
'*******************************************************************************
Public Sub TipCheck(ByVal Index As Integer, Optional AsText As Boolean = False)
  If RunMode Then Exit Sub                  'do nothing if in RUN mode
  TipAsText = AsText
  If Not Key2nd Then
    Select Case Index
      Case 1  'LRN
        SetTip "Toggle program LEARN mode"
      Case 2  'PGM
        SetTip "Program"
      Case 3  'Load
        SetTip "Load Module/Program"
      Case 4  'Save
        SetTip "Save Module/Program"
      Case 5  'CE
        SetTip "Clear typed text and values, or recover from errors"
      Case 6  'CLR
        SetTip "Clear pending operations, display, and value"
      Case 7  'OP
        SetTip "Special display, statistical, linear analysis commands"
      Case 8  'SST
        SetTip "Single-Step"
      Case 9  'INS
        SetTip "Toggles Insert and Overtype in LRN Mode"
      Case 10 'CUT
        SetTip "Cut selected line(s)"
      Case 11 'Copy
        SetTip "Copy selected line(s)"
      Case 12 'PtoR
        SetTip "Polar to Rectangular conversion"
      Case 13 'STO
        SetTip "Store the display register value into a variable"
      Case 14 'RCL
        SetTip "Recall a variable value to the display register"
      Case 15 'EXC
        SetTip "Exchange the display register value with that of a variable"
      Case 16 'SUM
        SetTip "Add the display register value to a variable"
      Case 17 'MUL
        SetTip "Multiply the display register value to a variable"
      Case 18 'IND
        SetTip "Indirection modifier, such as STO IND xx"
      Case 19 'Reset
        SetTip "Reset the call stack, brace stack, and program instruction pointer"
      Case 20 'Hkey
        SetTip "Hide a list of user-defined keys"
      Case 21 'Lnx
        SetTip "Natural Logarithm"
      Case 22 'E+
        SetTip "Statistics SUM"
      Case 23 'Mean
        SetTip "Calculate the statistics MEAN"
      Case 24 'x!
        SetTip "Factorial"
      Case 25 'X><T
        SetTip "Exchange the Display Register (X) with the Test Register (T)"
      Case 26 'Hyp
        If BaseType = TypHex And Not RunMode And Not CBool(MRunMode) Then
          SetTip "Hexadecimal A"
        Else
          SetTip "Arc (Inverse) modifier for angle keys"
        End If
      Case 27 'Sin
        SetTip "Get Sine of an angle"
      Case 28 'Cos
        SetTip "Get Cosine of an angle"
      Case 29 'Tan
        SetTip "Get Tangent of an angle"
      Case 30 '1/x
        SetTip "Get Reciprocal of display register"
      Case 31 'Txt
        SetTip "Begin typing Text (also "" key), or to terminate text input (also ENTER & other keypad keys)"
      Case 32 'Hex
        SetTip "Select Hexadecimal number base (16; 0-9, A-F)"
      Case 33 '&
        SetTip "Binary AND"
      Case 34 'StFlg
        SetTip "Set flag (0-9)"
      Case 35 'IfFlg
        SetTip "If Flag (0-9) is set, perform conditional block"
      Case 36 'X==T
        SetTip "If display Register equals the Test Register, perform conditional block"
      Case 37 'X>=T
        SetTip "If display Register is greater or equal to the Test Register, perform conditional block"
      Case 38 'X>T
        SetTip "If display Register is greater than the Test Register, perform conditional block"
      Case 39 'Dfn
        If BaseType = TypHex And Not RunMode And Not CBool(MRunMode) Then
          SetTip "Hexadecimal B"
        Else
          SetTip "Optional Define declaration"
        End If
      Case 40 ';
        SetTip "Statement separator"
      Case 41 '(
        SetTip "Opening math evaluation priority modifier"
      Case 42 ')
        SetTip "Closing math evaluation priority modifier"
      Case 43 '/
        SetTip "Division"
      Case 44 'Style
        SetTip "Program Listing format option (0-3)"
      Case 45 'Dec
        SetTip "Select Decimal number base (10; 0-9: default)"
      Case 46 '|
        SetTip "Binary OR"
      Case 47 'Int
        SetTip "Truncate decimal portion"
      Case 48 'Abs
        SetTip "Acquire Absolute positive value"
      Case 49 'Fix
        SetTip "Define decimal place display (0-9), or with EE, enable Engineering Notation"
      Case 50 'D.MS
        SetTip "Convert DDD.MMSSss to Decimal format DDD.ddd"
      Case 51 'EE
        SetTip "Enter Exponent display format (Scientific Notation)"
      Case 52 'Sbr
        If BaseType = TypHex And Not RunMode And Not CBool(MRunMode) Then
          SetTip "Hexadecimal C"
        Else
          SetTip "Define Subroutine"
        End If
      Case 53 '7
        SetTip "Digit 7"
      Case 54 '8
        SetTip "Digit 8"
      Case 55 '9
        SetTip "Digit 9"
      Case 56 '*
        SetTip "Multiplication"
      Case 57 '[']
        SetTip "Enter a program remark"
      Case 58 'Oct
        SetTip "Select Octal number base (8; 0-7)"
      Case 59 '~
        SetTip "Binary Not (invert bits)"
      Case 60 'Select
        SetTip "Begin a case selection block"
      Case 61 'Case
        SetTip "Case selector"
      Case 62 '{
        SetTip "Opening brace for a code block"
      Case 63 '}
        SetTip "Closing brace for a code block"
      Case 64 'Deg
        SetTip "Set Degrees mode (1/2 circle = 180; default)"
      Case 65 'Lbl
        If BaseType = TypHex And Not RunMode And Not CBool(MRunMode) Then
          SetTip "Hexadecimal D"
        Else
          SetTip "Define Label"
        End If
      Case 66 '4
        SetTip "Digit 4"
      Case 67 '5
        SetTip "Digit 5"
      Case 68 '6
        SetTip "Digit 6"
      Case 69 '-
          SetTip "Subtraction"
      Case 70 'Beep
        SetTip "Issue default Beep"
      Case 71 'Bin
        SetTip "Select Binary number base (2; 0-1)"
      Case 72 '^
        SetTip "Binary XOR"
      Case 73 'For
        SetTip "Begin a FOR loop"
      Case 74 'Do
        SetTip "Begin a DO, DO-WHILE, or DO-UNTIL loop"
      Case 75 'While
        SetTip "Being a WHILE loop"
      Case 76 'Pmt
        SetTip "Pause the program to allow the user to type a text response"
      Case 77 'Rad
        SetTip "Set Radians mode (1/2 circle = pi)"
      Case 78 'Key
        If BaseType = TypHex And Not RunMode And Not CBool(MRunMode) Then
          SetTip "Hexadecimal E"
        Else
          SetTip "Define User-Key"
        End If
      Case 79 '1
        SetTip "Digit 1"
      Case 80 '2
        SetTip "Digit 2"
      Case 81 '3
        SetTip "Digit 3"
      Case 82 '+
          SetTip "Addition"
      Case 83 'Plot
        SetTip "Plot lines, graphs, or charts"
      Case 84 'NVar
        SetTip "Declare a variable as numeric (default)"
      Case 85 '%
        SetTip "Modulo divisor"
      Case 86 'If
        SetTip "Begin a conditional block"
      Case 87 'Else
        SetTip "Block to execute when the IF condition is false"
      Case 88 'Cont
        SetTip "Continue execution at the top of the loop"
      Case 89 'Break
        SetTip "Break out of a loop"
      Case 90 'Grad
        SetTip "Set Grads mode (1/2 circle = 200)"
      Case 91 'R/S
        If BaseType = TypHex And Not RunMode And Not CBool(MRunMode) Then
          SetTip "Hexadecimal F"
        Else
          SetTip "Run/Stop"
        End If
      Case 92 '0
        SetTip "Digit 0"
      Case 93 '.
        SetTip "Decimal point (Member Operator from Structures)"
      Case 94 '+/-
        SetTip "Change sign"
      Case 95 '=
        SetTip "Finalize math operation, or terminate text input"
      Case 96 'Print
        SetTip "Print text on a plot"
      Case 97 'TVar
        SetTip "Declare a variable for TEXT storage"
      Case 98 '>>
        SetTip "Shift a value right x bits"
      Case 99 'y^
        SetTip "Raise the display register to a specified power"
      Case 100  'x2
        SetTip "Square the display register"
      Case 101  'pi
        SetTip "Set the display register to pi"
      Case 102  'Rnd
        SetTip "Return the next random number in a sequence"
      Case 103  'Mil
        SetTip "Set Mil mode (1/2 circle = 3200)"
      Case 104  'Pvt
        SetTip "Declare a Label or Subroutine Private to this program"
      Case 105  'Const
        SetTip "Declare a Constant"
      Case 106  'Const
        SetTip "Declare a Structure"
      Case 107  'NxLbl
        SetTip "Scan to next user-defined object definition"
      Case 108  'PvLbl
        SetTip "Scan to previous user-defined object definition"
      Case 109  'Line
        SetTip "Draw a line on the plot display"
      Case 110  '[
        SetTip "Begin an array dimension declaration"
      Case 111  ']
        SetTip "End an array dimension declaration"
      Case 112  'ClrVar
        SetTip "Nullify a variable and any arrays attached to it"
      Case 113  'SzOf
        SetTip "return the size of a string, a variable, or the number of elements in an array"
      Case 114  'Def
        SetTip "Define a conditional Compressr directive"
      Case 115  'IfDef
        SetTip "If a Compressr constant is defined, Compress some tasks"
      Case 116  'Edef
        SetTip "End a Compressr directive block"
    End Select
  Else
    Select Case Index
      Case 1    'MDL
        SetTip "Module"
      Case 2    'CMM
        SetTip "Clear Module Memory"
      Case 3    'Lapp
        SetTip "Load an application, or append current to the active module"
      Case 4    'ASCII
        SetTip "Save the program is ASCII format (useful for printing"
      Case 5    'CMs
        SetTip "Clear memory registers"
      Case 6    'CP
        SetTip "Clear active program memory (clear only Test-Register during RUN)"
      Case 7    'USR
        SetTip "Invoke USER DEFINED OPERATION (these are Ops defined by VB programmer in ModUSR)"
      Case 8    'BST
        SetTip "Back step instruction pointer"
      Case 9    'DEL
        SetTip "Delete selected program lines"
      Case 10   'Paste
        SetTip "Paste lines into program list"
      Case 11   'List
        SetTip "Used to list various internal lists"
      Case 12   'RtoP
        SetTip "Rectangular to Polar conversion"
      Case 13   'Push
        SetTip "Push the display register value onto a stack"
      Case 14   'Pop
        SetTip "Pop a value from a stack into the Display Register"
      Case 15   'StkEx
        SetTip "Swap the Display register with the most recent value on a stack"
      Case 16   'Sub
        SetTip "Subtract the display register from a variable"
      Case 17   'Div
        SetTip "Divide a variable by the display register"
      Case 18   '[<]
        SetTip "If left expression is less than right expression, perform conditional block"
      Case 19   '[>]
        SetTip "If left expression is greater than right expression, perform conditional block"
      Case 20   'Skey
        SetTip "Show selected user-defined keys"
      Case 21 'eX
        SetTip "Natural Antilog"
      Case 22 'E-
        SetTip "Subtract from a Statistical Sum"
      Case 23 'StDev
        SetTip "Compute the Standard Deviation of a Statistical Sum"
      Case 24 'Varnc
        SetTip "Compute Variance with N-1 weight"
      Case 25 'Yint
        SetTip "Calculate y-intercept (b) and slope (m)"
      Case 26 'Arc
        SetTip "Hyperbolic modifier for angle keys"
      Case 27 'Sec
        SetTip "Acquire the Secant of an angle"
      Case 28 'Csc
        SetTip "Acquire the Cosecant of an angle"
      Case 29 'Cot
        SetTip "Acquire the Cotangent of an angle"
      Case 30 'LogX
        SetTip "Acquire the Log of the left value to the specified base on the right"
      Case 31 'Var
        SetTip "Set current variable, and allow entry of a variable name"
      Case 32 '==
        SetTip "If left and right expressions are equal, perform conditional block"
      Case 33 '&&
        SetTip "Logical AND"
      Case 34 'RFlg
        SetTip "Reset flag (0-9)"
      Case 35 '!Flg
        SetTip "If flag (0-9) is not set, perform conditional block"
      Case 36 'X!=T
        SetTip "If display Register does not equal the Test Register, perform conditional block"
      Case 37 'X<=T
        SetTip "If display Register is less than or equal to the Test Register, perform conditional block"
      Case 38 'X<T
        SetTip "If display Register is less than the Test Register, perform conditional block"
      Case 39 'NOP
        SetTip "No Operation code placeholder (useful for white space separators)"
      Case 40 ':
        SetTip "Label terminator"
      Case 41 'Log
        SetTip "Common Logarithm"
      Case 42 '10^
        SetTip "Common Antilog"
      Case 43 '/=
        SetTip "Divide the Display Register into the current variable"
      Case 44 'Fmt
        SetTip "Define a numeric format string"
      Case 45 '!=
        SetTip "If left and right expressions are not equal, perform conditional block"
      Case 46 '||
        SetTip "Logical OR"
      Case 47 'Frac
        SetTip "Keep only fractional part of the Display Register"
      Case 48 'Sgn
        SetTip "Get the Signum of the display value"
      Case 49 '!Fix
        SetTip "Cancel Fixed Decimal length format"
      Case 50  'D.ddd
        SetTip "Convert DDD.ddd to DMS format DDD.MMSSss"
      Case 51 '!EE
        SetTip "Turn off Engineering and Scientific Notation"
      Case 52 'Call
        SetTip "Call a subroutine"
      Case 53 'Trim
        SetTip "Trim blanks from the left and right of a string variable"
      Case 54 'LTrim
        SetTip "Trim blanks from the left of a string variable"
      Case 55 'RTrim
        SetTip "Trim blanks from the right of a string variable"
      Case 56 '*=
        SetTip "Multiply the Display Register to the current variable"
      Case 57 'Rem
        SetTip "Enter a program remark"
      Case 58 '>=
        SetTip "If left expression is greater or equal to right expression, perform conditional block"
      Case 59 '[!]
        SetTip "Logical NOT"
      Case 60 'Open
        SetTip "Open a file for reading or writing"
      Case 61 'Close
        SetTip "Close one or all files"
      Case 62 'Read
        SetTip "Read from a file"
      Case 63 'Write
        SetTip "Write to a file"
      Case 64 'Swap
        SetTip "Swap the values of two variables"
      Case 65 'Gto
        SetTip "Go To a Lbl, Sbr, or Ukey location"
      Case 66 'LOF
        SetTip "Return the length of an open file"
      Case 67 'Get
        SetTip "Read a fix-sized record from a file"
      Case 68 'Put
        SetTip "Write a fixed sized record to a file"
      Case 69 '-=
        SetTip "Subtract the Display Register from the current variable"
      Case 70 'SysBP
        SetTip "Issue a system tone (0-4)"
      Case 71 '<=
        SetTip "If left expression is less than or equal to right expression, perform conditional block"
      Case 72 'Nor
        SetTip "Logical NOR (Not OR)"
      Case 73 'Incr
        SetTip "Increment the specified variable value"
      Case 74 'Decr
        SetTip "Decrement the specified variable value"
      Case 75 'Dsz
        SetTip "Decrement the specified variable and skip instruction block if zero"
      Case 76 'Dsnz
        SetTip "Decrement the specified variable and skip instruction block if not zero"
      Case 77 'All
        SetTip "General instruction used by some commands to select all allowed parameters"
      Case 78 'Rtn
        SetTip "Return from a subroutine and continue execution"
      Case 79 'LSet
        SetTip "Pad text on the right of a fixed-length string"
      Case 80 'RSet
        SetTip "Pad text on the left of a fixed-length string"
      Case 81 'Printf
        SetTip "Format string to the internal string storage register"
      Case 82 '+=
        SetTip "Add the Display Register to the current variable"
      Case 83 'RGB
        SetTip "Derive color value from RGB value (R,G,B)"
        Case 84 'Ivar
          SetTip "Declare an Integer variable"
      Case 85 '\
        SetTip "Integer division"
      Case 86 'As
        SetTip "Assign a file being opened as a specified file port (1-10)"
      Case 87 'ElseIf
        SetTip "Special type of Else statement, merged with another If statement"
      Case 88 'DBG
        If LrnMode Then
          SetTip "Enter Debug Mode. Enable/Disable Trace Mode with Open/Close parameter"
        Else
          SetTip "Toggle the Program Trace Mode"
        End If
      Case 89 'Gfree
        SetTip "Returns a free file I/O port number"
      Case 90 'Len
        SetTip "Set a string to a specified length"
      Case 91 'Stop
        SetTip "Stop program execution, without R/S continuation"
      Case 92 'With
        SetTip "Define variable/Structure to use WITH File I/O commands"
      Case 93 '[,]
        SetTip "List item separator"
      Case 94 'Val
        SetTip "Convert a formatted text value to a value in the Display Register"
      Case 95 'Adv
        SetTip "Advance the display one line"
      Case 96 'Print;
        SetTip "Print data to the display, but do not advance the line"
      Case 97 'Cvar
        SetTip "Declare a 1-character type variable"
      Case 98 '<<
        SetTip "Shift a value left x bits"
      Case 99 'Root
        SetTip "get right expression root of the left expression"
      Case 100  'Sqrt
        SetTip "Get the square root of the Display Register"
      Case 101  'e
        SetTip "Return the value of epsilon"
      Case 102  'Rnd#
        SetTip "Set a seed value for the random number generator"
      Case 103  'Until
        SetTip "Perform a Do Loop UNTIL a condition is met"
      Case 104  'Pub
        SetTip "Declare a Label or Subroutine Public to outside programs"
      Case 105  'Enum
        SetTip "Declare an enumeration list"
      Case 106  'AdrOf"
        SetTip "Derive the instruction address of a named item"
      Case 107  'Pcmp
        SetTip "Preprocess program (performed automatically by running and Comp). Use also 'F5'"
      Case 108  'Comp
        SetTip "Compress a robust user program to a module-storable state"
      Case 109  'Circle
        SetTip "Draw a Circle on the plot display"
      Case 110  'Split
        SetTip "Split a delimited stream of data to an array"
      Case 111  'Join
        SetTip "Joins an array into a data stream"
      Case 112  'Redim
        SetTip "Redimension an array non-destructively"
      Case 113  'Mid
        SetTip "Set or extract sub strings from a string"
      Case 114  'UDef
        SetTip "Undefine a conditional Compressr directive"
      Case 115  '!Def
        SetTip "If a Compressr constant is not defined, Compress some tasks"
      Case 116  'Delse
        SetTip "If conditional check (IfDef or !Def) is false, Compress this block"
    End Select
  End If
  TipAsText = False
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

