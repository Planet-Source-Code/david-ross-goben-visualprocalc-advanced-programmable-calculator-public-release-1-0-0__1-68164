Attribute VB_Name = "modOP"
Option Explicit

Public Const MaxOps As Long = 91  'max operations supported (so far)

'*******************************************************************************
' OP Codes 00 - 91
'----- Text processing ---------------------------------------------------------
' 00  Initialize Display Fields and set Text-Plot character to "*". There are 6 display fields,
'     with 6 characters to a field, numbered 1 to 6, from the left of the display to right.
' 01  Assign text to Field 1 of the Display Fields, Right Justified (6 character field).
' 02  Assign text to Field 2 of the Display Fields, Right Justified (6 character field).
' 03  Assign text to Field 3 of the Display Fields, Right Justified (6 character field).
' 04  Assign text to Field 4 of the Display Fields, Right Justified (6 character field).
' 05  Assign text to Field 5 of the Display Fields, Right Justified (6 character field).
' 06  Assign text to Field 6 of the Display Fields, Right Justified (6 character field).
' 07  Store the entire selected Display Line of text into the Display Fields, flush left.
' 08  Store the entire selected Display Line of text into the Display Fields, flush right.
' 09  Store the entire selected Display Line of text into the Display Fields, Centered.
' 10  Display all 6 Display Fields (36 characters) on the selected Display Line and advance
'     the Select Line.
' 11  Display all 6 Display Fields on the selected Display Line, merged with the contents
'     of the Display Register, then advance the Select Line.
'----- Module ID ---------------------------------------------------------------
' 12  Set the Display Register to the Module Number.
'----- Statistical and Linear Analysis -----------------------------------------
' 13  Initialize the Statistical and Linear Regression registers.
' 14  Return SUM y to the Display Register.
' 15  Return SUM y² to the Display Register.
' 16  Return N (Number of entries) to the Display Register.
' 17  Return SUM x to the Display Register.
' 18  Return SUM x² to the Display Register.
' 19  Return SUM xy to the Display Register.
' 20  Calculate Correlation Coefficient to the Display Register.
' 21  Calculate linear estimate y' against entered x' to the Display Register.
' 22  Calculate linear estimate x' against entered y' to the Display Register.
'----- Reset Flags -------------------------------------------------------------
' 23  Reset general Boolean Flags 0-9.
'----- Error checking ----------------------------------------------------------
' 24  Set Flag 7 if no error condition exists.
' 25  Set Flag 7 if an error condition exists.
'----- Text-Plot ---------------------------------------------------------------
' 26  Take the first character of TEXT input and make it the Text-Plot character (default = *).
' 27  Set the Text-Plot character to columns 0-35 of the Display Fields, from the right.
' 28  Set the Text-Plot character to columns 0-35 of the Display Fields, from the left.
' 29  Position the Display Line selector to the line specified by the Display Register (0-22).
' 30  Set Display Register to the selected Display Line number (0-22).
' 31  Initialize the display area for 36 x 23 Text Plotting.
' 32  Text Plot text data on row x, column y, in format: x,y:Text.
'----- Display Fields and VAR --------------------------------------------------
' 33  Apply the contents of the default variable (set by VAR) to all the Display Fields, flush left.
' 34  Apply the contents of the default variable (set by VAR) to all the Display Fields, flush right.
' 35  Extract the combined contents of the 6 Display Fields to the default variable (set by VAR).
'----- Program Data ------------------------------------------------------------
' 36  Download current module program into active memory (Learn Memory).
' 37  Set the Display register to the active program number.
'----- Display Flash -----------------------------------------------------------
' 38  Flash the Display Register value on the Selected Display Line for 1/2 second.
' 39  Flash all 6 Display Fields on the Selected Display Line for 1/2 second.
' 40  Flash all 6 Display Fields, merged with the Display Register value for 1/2 second.
' 41  Set the timer pause for OP codes 34-36 in 1/2 second intervals (1-4 = .5 to 2 seconds).
'----- Base Conversion ---------------------------------------------------------
' 42  Convert the Display Register value to Decimal Format.
' 43  Convert the Display Register value to Scientific Notation Format.
' 44  Convert the Display Register value to Hexadecimal Format.
' 45  Convert the Display Register value to Octal Format.
' 46  Convert the Display Register value to Binary Format.
' 47  Convert the Display Register value to Engineering Notation Format.
'----- Date Operations ---------------------------------------------------------
' 48  Set the Date Display format to the format defined on the selected Display Line.
'  Sample Date Display formats are:
'
'  Format Text          Displayed Text
'  ------------------------------------------------
'  m/d/yy               2/15/07
'  mm/dd/yy             02/15/07
'  dddd, mmmm dd, yyyy  Thursday, February 15, 2007
'  d-mmm                15-Feb
'  mmmm-yy              February-2007
'  Long Date            Thursday, February 15, 2007
'  Medium Date          15-Feb-07
'  Short Date           2/15/2007
'
' 49  Get the current date (number of days since Dec 30, 1899) and time in seconds Since midnight.
'     The format is "days.time", where the fractional time is the time in seconds / (60x60x24).
'     Time can be manually calculated to Hours ÷ 24 + minutes ÷ 1440 + seconds ÷ 86400.
' 50  Display date in current date format.
' 51  Set the Display Register to the Month from the Date value in the Display Register.
' 52  Set the Display Register to the Day of Month from the Date value in the Display Register.
' 53  Set the Display Register to the Year from the Date value in the Display Register.
' 54  From the Date value in the Display Register, set the Display Register to 1 if the date is
'     within a Leap Year, and to 0 if it is not.
' 55  Set the Display Register to the Day of Week from the Date value in the Display Register
'     (1=Sunday, 2 = Monday,...7=Saturday).
' 56  Set the Display Register to the Day of Year from the Date value in the Display Register.
' 57  Set the Display Register to the Week of Year from the Date value in the Display Register (1-53).
' 58  Set the Display Register to the Number of Days in the Current Month from the Date value in the
'     Display Register.
' 59  From the Date value in the Display Register, set the Display Register to 1 if the Date falls
'     on a Weekend, otherwise set it to 0.
' 60  Add (or subtract if negative) the number of months in the Display Register to the Date value
'     in the Test Register.
' 61  Add (or subtract if negative) the number of years in the Display Register to the Date value
'     in the Test Register.
' 62  Get Full Month name from a Date value in the Display Register.
'----- Time Operations ---------------------------------------------------------
' 63  Set Time Display format.
'  Sample Time Display formats are:
'
'  Format Text        Displayed Text
'  ---------------------------------
'  hh:mm:ss             14:34:28
'  hh:mm AM/PM          02:34 PM
'  h:mm:ss a/p          2:34:28 p
'  Long Time            02:34:28 PM
'  Medium Time          02:34 PM
'  Short Time           13:34
' 64  Display the time from a Date value in the Display Register in the current time format.
' 65  Get Hours from a Date value in the Display Register.
' 66  Get Minutes from a Date value in the Display Register.
' 67  Get Seconds from a Date value in the Display Register.
' 68  Add (or subtract if negative), the number of Hours in the Display Register to a Date value
'     in the Test Register.
' 69  Add (or subtract if negative), the number of Minutes in the Display Register to a Date value
'     in the Test Register.
' 70  Add (or subtract if negative), the number of Seconds in the Display Register to a Date value
'     in the Test Register.
'----- Graphical Operations ----------------------------------------------------
' 71  Set plot point size from a value in the Display Register: 8-24; default is 10.
' 72  Sets drawing color from a value in the Display Register: default is black (0); use [RGB]
'     to obtain color values from tri-color values.
' 73  Set the Display Register to the Text line height in pixels.
' 74  Set Text line height in pixels from a value in the Display Register: 0-32; this is reset
'     when the point size set by OP 71.
' 75  Sets the drawing mode from a value in the Display Register: 0=Copy (default),
'     1= OR, 2=AND, 3=XOR, 4=NXOR, 5=NOT, 6=Solid line, 7=Dash Line, 8=Dash-Dot Line,
'     9=Dash-Dot-Dot line. -1 to -4 defines a line draw width of 1 (default) to 4 respectively.
' 76  Set the Display Register to the last X position the mouse cursor was over (used by click event).
' 77  Set the Display Register to the last Y position the mouse cursor was over (used by click event).
'----- Misc Operations ---------------------------------------------------------
' 78  Convert the Display Register value to a percentage.
'     If there are no pending calculations, the Display Register is simply divided by 100.
'     If there is a pending +, apply a % increase (i.e., 2500 + 25%: "2500 + 15 Op 78"
'       yields 2875). This is internally evaluated as: ((100+25)÷100x2500). Note the no need for '='.
'     If there is a pending -, apply a % discount (i.e., 3500 discounted 25%:
'       "3500 - 25 Op 78" yields 2625). This is internally evaluated as: ((100-25)÷100x3500).
'       Note the no need for '='.
'     If there is a pending x, compute a % increase (i.e., What is the percentage increase from 300
'       if we add 500: "300 x 500 Op 78" yields 160). This is internally evaluated as:
'       ((300+500)÷500x100). Note the no need for '='.
'     If there is a pending ÷, compute a % of a value (i.e., to find what % 80 is of 125:
'       "125 ÷ 80 Op 78" yields 64). This is internally evaluated as: (80÷125x100).
'       Note the no need for '='. If you want to compute % difference from 80 to 125,
'       use "(100-125÷80 OP 78)", which yields 36.
' 79  Compute Permutations. Compute permuting 'r' different members among 'n' members,
'     where n>=r: nPr = n!÷(n-r)!.
'     Entry Method: "n [X><T] r OP 79".
' 80  Compute Combinations. Compute how many combinations of 'r' members can be obtained
'     when there are 'n' total members, where n>=r: nCr = n!÷(r!(n-r)!).
'     Entry Method: "n [X><T] r OP 80".
' 81  Convert floating point decimal value to fraction. Hence, .625 will result
'     in a text value of "5/8", 14.1875 results in "14·3/16", and -2.002 becomes "-2·1/500".
' 82  Convert the angle in the Display Register from the current Angle Mode to Degrees.
' 83  Convert the angle in the Display Register from the current Angle Mode to Radians.
' 84  Convert the angle in the Display Register from the current Angle Mode to Grads.
' 85  Convert the angle in the Display Register from the current Angle Mode to Mils.
' 86  Set Display Register to the length of the text on the selected Display Line.
'----- Physical Constants ------------------------------------------------------
' 87  Provided an index number (1-22) in the Display Register, a Physical constant value
'     will be placed in the Display Register. To use these values normally includes
'     additional calculations. For example, index 21, Faraday Constant, is a value of
'     9.648670e+07, the value is described as "9.648670e+07 C k mole-¹" (the -¹ is
'     actually an inverse indicator), meaning that this constant is used to complete
'     these calculations. A simple rule of thumb is, if you do not understand them,
'     then you do not need to use them.
'
'     Index Symbol Description:Value
'     ----- ------ ------------
'       0     F    Faraday Constant:         9.648670E+07 C k mole-¹
'       1     c    Speed of Light:           2.9979250E+08 m/sec-¹
'       2     e    Electron Charge:          1.6021917E-10 C
'       3     N    Avogado Number:           6.022169E+26 k mole-¹
'       4     eV   Electron Volt:            1.602E-19 J
'       5     me   Electron Rest Mass:       9.109558E-31 kg
'       6     Mp   Proton Rest Mass:         1.672614E-27 kg
'       7     Mn   Neutron Rest Mass:        1.674920E-27 kg
'       8     amu  Atomic Mass Unit:         1.660531E-27 kg
'       9     e/me Electron Charge to Mass ratio: 1.7588028E+11 C kg-¹
'      10     h    Planck Constant:          6.626196E-34 J-sec
'      11     Roo  RydBerg Constant:         1.09737312E+07 m-¹
'      12     Ro   Gas Constant:             8.31434E+03 J-k mole-¹ K-¹
'      13     k    Boltzmann Constant:       1.380622E-23 JK-¹
'      14     G    Gravitational Constant:   6.6732E-11 N-m²kg-²
'      15     µb   Bohr Magaton:             9.274096E-24 JT-¹
'      16     µe   Electron Magnetic Moment: 9.284851E-24 JT-¹
'      17     µp   Proton Magnetic Moment:   1.4106203E-24 JT-¹
'      18     lc   Compton Wavelength of the Electron: 2.4263096E-26 m
'      19     lc.p Compton Wavelength of the Proton:   1.3214409E-15 m
'      20     lc.n Compton Wavelength of the Neutron:  1.3196217E-15 m
'      21    o     Stefan-Boltzmann Constant 5.56704E-08 W/m2-K4
'      22    ao    Bohr Radius:  5.5292E-11 m
'----- Special Date Finder -----------------------------------------------------
' 88  Date Finder. This is a special date finder routine that returns a date in
'     VisualCalc date format (see OP 49), as the number of days since Dec 30, 1899.
'     The provided text format is: Year[,Month[,[Day[,Week[,Weekday]]]]]
'       Special Ranges: Week (1-5), WeekDay (1-7; 1=Sunday)
'       Keep Day=0 if you will specify a Week number.
'       Allowed Input Samples:
'         Return date for Jan 1,1987              : 1987
'         Return date for March 1,1987            : 1987,3
'         Return date for July 4,1987             : 1987,7,4
'         Return date for June,1987, Week 2       : 1987,6,0,2
'         Return date for Monday, May,1987, Week 3: 1987,5,0,3,2
'     This function can compute fixed holidays, such as Labor Day; the first
'     Monday in September (year,9,0,1,2).
'
'     NOTE: Weeks (1-5) begin on Sunday. Week 5 is special, as it will most-
'           often exceed the given month. It will in effect cause a search
'           backward for a legal week. This is important for finding such
'           holidays as Memorial Day, which is the last Monday in May. If May 1
'           falls on Sunday, then there are 5 Sundays in May, and the 5th week's
'           Monday falls on May 30. However, if May falls on any other day of
'           the week, there are only 4 Sundays in May. Using week 5 will ensure
'           that the last week is used, as it will back up a week until it finds
'           a legal Sunday for that month (year,5,0,5,2).
'
'     Some Noteable Holidays:
'       New Year's Day: January 1 (year).
'       Birthday of Martin Luther King, Jr.: the third Monday in January (year,1,0,3,2).
'       Washington 's Birthday: the third Monday in February (year,2,0,3,2).
'       Memorial Day, the last Monday in May (year,5,0,5,2).
'       Independence Day: July 4 (year,7,4).
'       Labor Day, the first Monday in September (year,9,0,1,2).
'       Columbus Day: the second Monday in October (year,10,0,2,2).
'       Veterans Day: November 11 (year,11,11).
'       Thanksgiving Day: the fourth Thursday in November (year,11,0,4,5).
'       Christmas Day: December 25 (year,12,25).
'----- Conversion Factors ------------------------------------------------------
' 89  Provide conversion factors for translating one unit of measure to another.
'     The result is multiplied by the value to convert to obtain the value of the
'     desired unit of measure.
'
'INDEX   FROM unit     TO unit        Result to MULTIPLY by
'----------------------------------------------------------
'0       acres         square feet    43560
'1       acres         square miles   0.0015625
'2       angstroms     centimeters    0.00000001
'3       angstroms     inches         2.540e+08
'4       angstroms     meters         0.00000000010
'5       ast unit      kilometers     1.495e+08
'6       ast unit      miles          9.289499323948143e+07
'7       board ft      cubic feet     0.0833333333333333
'8       bushels       cubic cm       35239.070
'9       cord ft       cords          0.1250
'10      coft ft       cubic feet     16
'11      centimeters   angstroms      1.0e+8
'12      centimeters   feet           0.0328083989501312
'13      centimeters   inches         0.3937007874015748
'14      centimeters   kilometers     0.000010
'15      centimeters   meters         0.010
'16      centimeters   miles          0.0000062137119224
'17      centimeters   yards          0.0109361329833771
'18      cords         cord feet      8
'10      cubic cm      cubic inches   0.0610237440947323
'20      cubic cm      cubic metes    0.0000010
'21      cubic cm      cubic yards    0.0000013079506193
'22      cubic feet    board feet     12
'23      cubic feet    cord feet      0.06250
'24      cubic feet    cubic inches   1728
'25      cubic feet    cubic meters   0.0283168465920000
'26      cubic feet    cubic yards    0.0370370370370370
'27      cubic inches  cubic cm       16.3870640
'28      cubic inches  cubic feet     0.0005787037037037
'29      cubic inches  cubic yards    0.0000214334705075
'30      cubic meters  cubic cm       1000000.0
'31      cubic meters  cubic feet     35.3146667214885903
'32      cubic meters  cubic yards    1.3079506193143922
'33      cubic yards   cubic meters   764554.85798400
'34      cubic yards   cubic feet     27
'35      cubic yards   cubic inches   46656
'36      cubic yards   cubic meters   0.7645548579840000
'37      ft per sec    miles per hr   0.6818181818181818
'38      feet          centimeters    30.480
'39      feet          inches         12
'40      feet          kilometers     0.00030480
'41      feet          meters         0.30480
'42      feet          miles          0.0001893939393939
'43      feet          rods           0.0606060606060606
'44      feet          yards          0.3333333333333333
'45      inches        centimeters    2.540
'46      inches        feet           0.0833333333333333
'47      inches        metes          0.02540
'48      inches        miles          0.0000157828282828
'49      inches        yards          0.0277777777777778
'50      kilometers    ast units      1.4950e+08
'51      kilometers    centimeters    100000.0
'52      kilometers    feet           3280.8398950131233600
'53      kilometers    meters         1000.0
'54      kilometers    miles          0.6213711922373340
'55      kilometers    rods           198.8387815159468700
'56      miles         ast units      0.0000000107648428
'57      miles         centimeters    160934.40
'58      miles         feet           5280
'59      miles         inches          63360
'60      miles         kilometers      1.6093440
'61      miles         meters          1609.3440
'62      miles         rods            320
'63      miles         yards           1760.0
'64      miles per hr  ft per sec      1.4666666666666667
'65      meters        angstroms       1.0e+10
'66      meters        centimetes      100
'67      meters        feet            3.2808398950131234
'68      meters        inches          39.3700787401574803
'69      meters        kilometers      0.0010
'70      meters        miles           0.0006213711922373
'71      meters        rods            0.1988387815159469
'72      meters        yards           1.0936132983377078
'73      rods          feet            16.50
'74      rods          kilometers      0.00502920
'75      rods          meters          5.029200
'76      rods          miles           0.0031250
'77      rods          yards           5.50
'78      square ft     sq in           144
'79      square ft     sq miles        0.0000229568411387
'80      square in     sq ft           0.0069444444444444
'81      square in     sq mi           0.0000000002490977
'82      square mi     acres           640
'83      square mi     sq ft           2.787840e+07
'84      square mi     sq in           4.01448960e+09
'85      yards         centimeters     91.440
'86      yards         feet            3
'87      yards         inches          36
'88      yards         meters          0.91440
'89      yards         miles           0.0005681818181818
'90      yards         rods            0.1818181818181818
'----- Binomial Coefficient ----------------------------------------------------
' 90 Compute Binomial Coefficient.
'    Binomial Coefficient = n! / (j!(n-j)!
'    N is in test register, J is in the Display Register
'----- Atan2 Function ----------------------------------------------------------
' 91  Provided the Y value in the Test Register, and the X value in the
'     Display register, return the angle from the X axis at point (y,x).
'*******************************************************************************

'*******************************************************************************
' Private valriables used for display fields
'*******************************************************************************
Private Fields(FieldCnt) As String 'display fields)

'*******************************************************************************
' Subroutine Name   : ProcessOP
' Purpose           : Invoke selected OP code
'*******************************************************************************
Public Sub ProcessOP(ByVal OP As Integer)
  Select Case OP
    Case 0
      Call Op00
    Case 1
      Call Op01
    Case 2
      Call Op02
    Case 3
      Call Op03
    Case 4
      Call Op04
    Case 5
      Call Op05
    Case 6
      Call Op06
    Case 7
      Call Op07
    Case 8
      Call Op08
    Case 9
      Call Op09
    Case 10
      Call Op10
    Case 11
      Call Op11
    Case 12
      Call Op12
    Case 13
      Call Op13
    Case 14
      Call Op14
    Case 15
      Call Op15
    Case 16
      Call Op16
    Case 17
      Call Op17
    Case 18
      Call Op18
    Case 19
      Call Op19
    Case 20
      Call Op20
    Case 21
      Call Op21
    Case 22
      Call Op22
    Case 23
      Call Op23
    Case 24
      Call Op24
    Case 25
      Call Op25
    Case 26
      Call Op26
    Case 27
      Call Op27
    Case 28
      Call Op28
    Case 29
      Call Op29
    Case 30
      Call Op30
    Case 31
      Call Op31
    Case 32
      Call Op32
    Case 33
      Call Op33
    Case 34
      Call Op34
    Case 35
      Call Op35
    Case 36
      Call Op36
    Case 37
      Call Op37
    Case 38
      Call Op38
    Case 39
      Call Op39
    Case 40
      Call Op40
    Case 41
      Call Op41
    Case 42
      Call Op42
    Case 43
      Call Op43
    Case 44
      Call Op44
    Case 45
      Call Op45
    Case 46
      Call Op46
    Case 47
      Call Op47
    Case 48
      Call Op48
    Case 49
      Call Op49
    Case 50
      Call Op50
    Case 51
      Call Op51
    Case 52
      Call Op52
    Case 53
      Call Op53
    Case 54
      Call Op54
    Case 55
      Call Op55
    Case 56
      Call Op56
    Case 57
      Call Op57
    Case 58
      Call Op58
    Case 59
      Call Op59
    Case 60
      Call Op60
    Case 61
      Call Op61
    Case 62
      Call Op62
    Case 63
      Call Op63
    Case 64
      Call Op64
    Case 65
      Call Op65
    Case 66
      Call Op66
    Case 67
      Call Op67
    Case 68
      Call Op68
    Case 69
      Call Op69
    Case 70
      Call Op70
    Case 71
      Call Op71
    Case 72
      Call Op72
    Case 73
      Call Op73
    Case 74
      Call Op74
    Case 75
      Call Op75
    Case 76
      Call Op76
    Case 77
      Call Op77
    Case 78
      Call Op78
    Case 79
      Call Op79
    Case 80
      Call Op80
    Case 81
      Call Op81
    Case 82
      Call Op82
    Case 83
      Call Op83
    Case 84
      Call Op84
    Case 85
      Call Op85
    Case 86
      Call Op86
    Case 87
      Call Op87
    Case 88
      Call Op88
    Case 89
      Call Op89
    Case 90
      Call Op90
    Case 91
      Call Op91
  End Select
'
' if timer is enabled, then pause
'
  With frmVisualCalc.tmrPause
    If .Enabled Then
      Do While .Enabled
        DoEvents
      Loop
    End If
  End With
End Sub

'*******************************************************************************
' Function Name     : FormatString
' Purpose           : Format text to a single field
'*******************************************************************************
Private Function FormatString(Txt As String) As String
  Dim S As String
  
  S = Txt                                     'get text to format
  If Len(S) > FieldWidth Then
    S = Left$(S, FieldWidth)                  'truncate if too long
  ElseIf Len(S) < FieldWidth Then
    S = String$(FieldWidth - Len(S), 32) & S  'pad left if too short
  End If
  FormatString = S
End Function

'*******************************************************************************
' Function Name     : OpMerge
' Purpose           : Merge all fields into 1 string
'*******************************************************************************
Private Function OpMerge() As String
  Dim S As String
  
  S = String$(DisplayWidth, 32)
  Mid$(S, 1, FieldWidth) = Fields(1)
  Mid$(S, FieldWidth + 1, FieldWidth) = Fields(2)
  Mid$(S, FieldWidth * 2 + 1, FieldWidth) = Fields(3)
  Mid$(S, FieldWidth * 3 + 1, FieldWidth) = Fields(4)
  Mid$(S, FieldWidth * 4 + 1, FieldWidth) = Fields(5)
  Mid$(S, FieldWidth * 5 + 1, FieldWidth) = Fields(6)
  OpMerge = S
End Function

'*******************************************************************************
' Subroutine Name   : Op00
' Purpose           : Initialize Field Registers and set Text-Plot character to "*"
'*******************************************************************************
Public Sub Op00()
  Dim S As String
  
  S = String$(FieldWidth, 32) 'init fields to blank
  Fields(1) = S
  Fields(2) = S
  Fields(3) = S
  Fields(4) = S
  Fields(5) = S
  Fields(6) = S

  TxtPltChr = "*"             'set default text-plot character
End Sub

'*******************************************************************************
' Routines to assign field text
'*******************************************************************************
Private Sub Op01()
  Fields(1) = FormatString(DspTxt)
End Sub

Private Sub Op02()
  Fields(2) = FormatString(DspTxt)
End Sub

Private Sub Op03()
  Fields(3) = FormatString(DspTxt)
End Sub

Private Sub Op04()
  Fields(4) = FormatString(DspTxt)
End Sub

Private Sub Op05()
  Fields(5) = FormatString(DspTxt)
End Sub

Private Sub Op06()
  Fields(6) = FormatString(DspTxt)
End Sub

'*******************************************************************************
' Subroutine Name   : Op07
' Purpose           : Store full text line into fields, flush left
'*******************************************************************************
Private Sub Op07()
  Dim S As String
  Dim Idx As Long
  
  If RunMode Then
    If DisplayText Then
      S = DspTxt
    Else
      S = DisplaySetup
    End If
  Else
    With frmVisualCalc.lstDisplay
      S = .List(.ListIndex)                                   'grab selection line
    End With
  End If
  If Len(S) < DisplayWidth Then                             'pad right
    S = S & String$(DisplayWidth - Len(S), 32)
  End If
  For Idx = 1 To FieldCnt                                   'apply to fields
    Fields(Idx) = Mid$(S, Idx * FieldWidth - 5, FieldWidth)
  Next Idx
End Sub

'*******************************************************************************
' Subroutine Name   : Op08
' Purpose           : Store full text line into fields, flush right
'*******************************************************************************
Private Sub Op08()
  Dim S As String
  Dim Idx As Long
  
  If RunMode Then
    If DisplayText Then
      S = DspTxt
    Else
      S = DisplaySetup
    End If
  Else
    With frmVisualCalc.lstDisplay
      S = .List(.ListIndex)                                   'grab selection line
    End With
  End If
  If Len(S) < DisplayWidth Then                             'pad left
    S = String$(DisplayWidth - Len(S), 32) & S
  End If
  For Idx = 1 To FieldCnt                                   'apply to fields
    Fields(Idx) = Mid$(S, Idx * FieldWidth - 5, FieldWidth)
  Next Idx
End Sub

'*******************************************************************************
' Subroutine Name   : Op09
' Purpose           : Store full text line into fields, centered
'*******************************************************************************
Private Sub Op09()
  Dim S As String
  Dim Idx As Long
  
  If RunMode Then
    If DisplayText Then
      S = DspTxt
    Else
      S = DisplaySetup
    End If
  Else
    With frmVisualCalc.lstDisplay
      S = .List(.ListIndex)                                   'grab selection line
    End With
  End If
  If Len(S) < DisplayWidth Then                             'pad right
    S = String$(DisplayWidth \ 2 - Len(S) \ 2, 32) & S        'center the text
    S = S & String$(DisplayWidth - Len(S), 32)
  End If
  For Idx = 1 To FieldCnt                                   'apply to fields
    Fields(Idx) = Mid$(S, Idx * FieldWidth - 5, FieldWidth)
  Next Idx
End Sub

'*******************************************************************************
' Subroutine Name   : Op10
' Purpose           : Display Op data only
'*******************************************************************************
Private Sub Op10()
  DspTxt = OpMerge()              'get merged data
  With frmVisualCalc.lstDisplay
    .List(.ListIndex) = DspTxt    'stuff data to current line
  End With
  If RunMode Then
    DoEvents                      'let screen update
    DisplayText = True
  Else
    Call NewLine                  'advance to a new line if not RUN mode
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Op11
' Purpose           : Merge Op data and display value
'*******************************************************************************
Private Sub Op11()
  Dim S As String, T As String
  
  If DisplayText Then
    T = DspTxt
  Else
    T = DisplaySetup                          'get value that is/will be normally displayed
  End If
  S = OpMerge()                             'get merged text fields
  DspTxt = Left$(S, DisplayWidth - Len(T)) & T 'overwrite right side of text with display data
  With frmVisualCalc.lstDisplay
    .List(.ListIndex) = DspTxt              'display it
  End With
  If RunMode Then
    DoEvents                                'let screen update
    DisplayText = True
  Else
    Call NewLine                            'advance to a new line if not RUN mode
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Op12
' Purpose           : Set the Display Register to the Module Number
'*******************************************************************************
Private Sub Op12()
  DisplayReg = ModName
  DisplayText = False
End Sub

'*******************************************************************************
' NOTE: OP13 thru OP22 are in modStatLinear.bas
'*******************************************************************************

'*******************************************************************************
' Subroutine Name   : Op23
' Purpose           : Clear Flags 0-9
'*******************************************************************************
Public Sub Op23()
  Dim i As Integer
  
  For i = 0 To 9
    flags(i) = False  'clear all flags
  Next i
End Sub

'*******************************************************************************
' Subroutine Name   : Op24
' Purpose           : Set Flag 7 if no error condition exists
'*******************************************************************************
Private Sub Op24()
  flags(7) = Not CBool(InstrErr)
End Sub

'*******************************************************************************
' Subroutine Name   : Op25
' Purpose           : Set Flag 7 if error condition exists
'*******************************************************************************
Private Sub Op25()
  flags(7) = CBool(InstrErr)
End Sub

'*******************************************************************************
' Subroutine Name   : Op26
' Purpose           : Take first character of TXT input and make it the
'                   : Text-Plot character (default [*])
'*******************************************************************************
Private Sub Op26()
  If CBool(Len(DspTxt)) Then
    TxtPltChr = Left$(DspTxt, 1)  'use user-defined character
  Else
    ForcError "No text to derive plot character from"
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Op27
' Purpose           : Text-Plot character to columns 0-(DisplayWidth-1),
'                   : from the right, from the 0-35 value in the display register
'*******************************************************************************
Private Sub Op27()
  Dim TV As Double
  Dim Idx As Long
  
  TV = Fix(DisplayReg)                          'get column
  If TV < 0# Or TV >= CDbl(DisplayWidth) Then   'check for out of range
    ForcError "The display register value range is 0-" & CStr(DisplayWidth - 1)
  Else
    If Len(TxtPltChr) = 0 Then TxtPltChr = "*"  'ensure we have a plot character
    Idx = CLng(TV)                              'get position
    Mid$(Fields(6 - Idx \ FieldWidth), FieldWidth - Idx Mod FieldWidth, 1) = TxtPltChr
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Op28
' Purpose           : Text-Plot character to columns 0-(DisplayWidth-1),
'                   : from the left, from the 0-35 value in the display register
'*******************************************************************************
Private Sub Op28()
  Dim TV As Double
  Dim Idx As Long
  
  TV = Fix(DisplayReg)                          'get column
  If TV < 0# Or TV >= CDbl(DisplayWidth) Then   'check for out of range
    ForcError "The display register value is greater than 35"
  Else
    If Len(TxtPltChr) = 0 Then TxtPltChr = "*"  'ensure we have a plot character
    Idx = CLng(TV)                              'get position
    Mid$(Fields(Idx \ FieldWidth + 1), Idx Mod FieldWidth + 1, 1) = TxtPltChr
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Op29
' Purpose           : Position selection to a line on the display (0-27)
'*******************************************************************************
Private Sub Op29()
  Dim TV As Double
  Dim Idx As Integer, IV As Integer, dst As Integer, TI As Integer
  
  TV = Fix(DisplayReg)                                  'get the desired line
  If TV < 1# Or TV > CDbl(DisplayHeight - 1) Then       'in range?
    ForcError "Op38 range is 1-" & CStr(DisplayHeight - 1) 'nope
  Else
    IV = CInt(TV)                                       'save copy of line
    With frmVisualCalc.lstDisplay
      TI = .TopIndex                                    'save top index
      dst = TI + IV - 1                                 'set destination
      If dst > .ListCount - 1 Then                      'greater than current top?
        For Idx = .ListCount - 1 To dst - 1             'generate new lines if so...
          .AddItem vbNullString
        Next Idx
      End If
      .TopIndex = TI                                    'reset top index
      Call SelectOnly(dst)                              'select desired line
    End With
    DoEvents
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Op30
' Purpose           : Set Display register to the display line number (0-27)
'*******************************************************************************
Private Sub Op30()
  With frmVisualCalc.lstDisplay
    DisplayReg = CDbl(.ListIndex - .TopIndex)
  End With
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op31
' Purpose           : Initialize display for 36 x 28 text plot
'*******************************************************************************
Private Sub Op31()
  Dim Idx As Integer
  Dim S As String
  
  Call Clear_Screen                       'erase screen data
  S = String$(DisplayWidth, 32)           'set null space
  With frmVisualCalc.lstDisplay
    For Idx = 1 To DisplayHeight - 1      'full with blanks
      .AddItem S
    Next Idx
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : Op32
' Purpose           : Plot data on row x, column y, in format x,y:data
'*******************************************************************************
Private Sub Op32()
  Dim X As Integer, Y As Integer
  Dim S As String, Ds As String, Xs As String, Ys As String
  Dim Oops As Boolean
  Dim TV As Double
  
  With frmVisualCalc.lstDisplay
    Do
      If .ListCount < DisplayHeight Then Call Op31          'init screen if neccessary
      If Len(DspTxt) = 0 Then Exit Do                       'if nothing to process
      X = InStr(1, DspTxt, ",")                             'find comma
      Y = InStr(X + 1, DspTxt, ":")                         'find colon
      If X = 0 Or Y = 0 Then Exit Do
      Xs = Trim$(Left$(DspTxt, X - 1))                      'get x
      If Len(Xs) = 0 Then Exit Do
      Ys = Trim$(Mid$(DspTxt, X + 1, Y - X - 1))            'get y
      If Len(Ys) = 0 Then Exit Do
      Ds = Mid$(DspTxt, Y + 1)                              'get data
      If Len(Ds) = 0 Then Exit Do
      TV = Fix(Val(Xs))                                     'see if X is in width range
      If TV < 1# Or TV > CDbl(DisplayWidth) Then Exit Do
      X = CInt(TV)
      TV = Fix(Val(Ys))                                     'see if Y is in height range
      If TV < 1# Or TV > CDbl(DisplayHeight) Then Exit Do
      Y = CInt(TV)
      S = .List(X)                                          'get line to work on
      If Len(S) < DisplayWidth Then S = String$(DisplayWidth - Len(S), 32) & S
      Mid$(S, Y, Len(Ds)) = Ds                              'stuff data
      If Len(S) > DisplayWidth Then S = Left$(S, DisplayWidth)
      .List(X) = S                                          'display it
      DoEvents                                              'let paints catch up
      Exit Sub                                              'all is ok
    Loop
    ForcError "This operation requires text format: x,y:TextData"     'error
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : Op33
' Purpose           : Apply default variable to all the fields, flush left
'*******************************************************************************
Private Sub Op33()
  Dim S As String
  Dim Idx As Long
  
  If CurrentVar = -1 Then                                     'default varaible defined?
    ForcError "Current default variable is not defined"       'no, so error
  Else
    S = CStr(ExtractValue(CurrentVarObj))                     'else get text data
    If Len(S) < DisplayWidth Then                             'make display width
      S = S & String$(DisplayWidth - Len(S), 32)              'pad right
    End If
    For Idx = 1 To FieldCnt                                   'apply to fields
      Fields(Idx) = Mid$(S, (Idx - 1) * FieldWidth + 1, FieldWidth)
    Next Idx
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Op34
' Purpose           : Apply default variable to all the fields, flush right
'*******************************************************************************
Private Sub Op34()
  Dim S As String
  Dim Idx As Long
  
  If CurrentVar = -1 Then                                     'default varaible defined?
    ForcError "Current default variable is not defined"       'no, so error
  Else
    S = CStr(ExtractValue(CurrentVarObj))                     'else get text data
    If Len(S) < DisplayWidth Then                             'make display width
      S = String$(DisplayWidth - Len(S), 32) & S              'pad left
    End If
    For Idx = 1 To FieldCnt                                   'apply to fields
      Fields(Idx) = Mid$(S, (Idx - 1) * FieldWidth + 1, FieldWidth)
    Next Idx
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Op35
' Purpose           : Extract fields to default variable
'*******************************************************************************
Private Sub Op35()
  Dim S As String
  
  If CurrentVar = -1 Then                                     'default varaible defined?
    ForcError "Current default variable is not defined"       'no, so error
  ElseIf CurrentVarTyp <> vString Then
    ForcError "Operation requires a default string variable"
  Else
    S = OpMerge
    Call StuffValue(CurrentVarObj, CVar(S))                   'assign data to variable
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Op36
' Purpose           : Download current module program into active memory
'*******************************************************************************
Private Sub Op36()
  Dim Idx As Long
  Dim HldPgm As Integer, Iptr As Integer
  
  If Not CBool(ActivePgm + ModPrep) Then
    ForcError "You must first use the Pgm command to activate the Module Program to download"
    Exit Sub
  End If
  If ModLocked Then
    ForcError "Cannot download the currently selected program. The Module is LOCKED"
    Exit Sub
  End If
  '
  ' init pgm 00 space
  '
  HldPgm = ActivePgm                  'save Pgm #, incase Run Mode
  Iptr = InstrPtr                     'and save instruction pointer
  Call CP_Support
  '
  ' copy instructions
  '
  If CBool(ModPrep) Then ActivePgm = ModPrep
    
  InstrCnt = GetInstrCnt()            'get number of instructions
  Do While InstrCnt > InstrSize       'make sure instruction buffer is big enough
    InstrSize = InstrSize + InstrInc  'increment by offset increment
  Loop
  ReDim Instructions(InstrSize)       'ensure size of buffer matches and initialized
  InstrPtr = 0                        'now download instructions
  For Idx = ModMap(ActivePgm - 1) To ModMap(ActivePgm) - 1
    Instructions(InstrPtr) = ModMem(Idx)
    InstrPtr = InstrPtr + 1
  Next Idx
  InstrPtr = 0                        'init to start
  '
  ' copy labels
  '
  LblCnt = ModLblMap(ActivePgm) - ModLblMap(ActivePgm - 1)
  Do While LblCnt > LblSize
    LblSize = LblSize + DefInc        'yes, so bump pool by increment
  Loop
  ReDim Preserve Lbls(LblSize)
  LblCnt = 0
  For Idx = ModLblMap(ActivePgm - 1) + 1 To ModLblMap(ActivePgm) - 1
    Lbls(LblCnt) = ModLbls(Idx)
    If Lbls(LblCnt).LblTyp = TypStruct Then
      Lbls(LblCnt).LblValue = Lbls(LblCnt).LblValue - ModStMap(ActivePgm - 1)
    End If
    LblCnt = LblCnt + 1
  Next Idx
  '
  ' copy structures
  '
  Idx = ModStMap(ActivePgm) - ModStMap(ActivePgm - 1)
  If CBool(Idx) Then
    ReDim StructPl(Idx)                  'set aside space for new struct
    StructCnt = 0
    For Idx = ModStMap(ActivePgm - 1) To ModStMap(ActivePgm) - 1
      StructCnt = StructCnt + 1
      CloneStruct StructPl(StructCnt), ModStPl(Idx)
    Next Idx
  End If
  '
  ' set pgm 00 as active, or reset pgm if invoked
  '
  PgmName = ActivePgm         'set loaded program ID
  If RunMode And CBool(ModPrep) And CBool(HldPgm) Then  'if pgm invoked
    ActivePgm = HldPgm        'reset active program
    InstrPtr = Iptr           'reset instruction pointer
    ModPrep = 0               'reset pgm invoke flag
    UpdateStatus              'update status changes
    Exit Sub
  End If
  ActivePgm = 0               'force pgm 00
  If AutoPprc Then Call Preprocess  'preprocess as needed
  frmVisualCalc.mnuWinASCII.Enabled = Preprocessd
  DisplayMsg "Sucessfully downloaded Pgm " & Format(PgmName, "00")
  ModPrep = 0                 'reset pgm invoke flag
  UpdateStatus                'update status changes
  RunMode = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op37
' Purpose           : Set the Display Register to the active program number
'*******************************************************************************
Private Sub Op37()
  DisplayReg = CDbl(ActivePgm)
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op38
' Purpose           : Flash the Display Register value on the display for 1/2 second
'*******************************************************************************
Private Sub Op38()
  Call ForceDisplay                     'display the active data on the current line
  frmVisualCalc.tmrPause.Enabled = True
  DoEvents
End Sub

'*******************************************************************************
' Subroutine Name   : Op39
' Purpose           : Display Op data only on the display for 1/2 second
'*******************************************************************************
Private Sub Op39()
  DspTxt = OpMerge()
  With frmVisualCalc.lstDisplay
    .List(.ListIndex) = DspTxt
  End With
  frmVisualCalc.tmrPause.Enabled = True
  DoEvents
  DisplayText = True
End Sub

'*******************************************************************************
' Subroutine Name   : Op40
' Purpose           : Merge Op data and display value on the display for 1/2 second
'*******************************************************************************
Private Sub Op40()
  Dim S As String, T As String
  
  T = DisplaySetup                          'get value that is/will be normally displayed
  S = OpMerge()                             'get merged text fields
  DspTxt = Left$(S, DisplayWidth - Len(T)) & T 'overwrite right side of text with display data
  With frmVisualCalc.lstDisplay
    .List(.ListIndex) = DspTxt              'display it
  End With
  frmVisualCalc.tmrPause.Enabled = True
  DoEvents
  DisplayText = True
End Sub

'*******************************************************************************
' Subroutine Name   : op41
' Purpose           : Increase pause timer to 1/2 second intervals
'*******************************************************************************
Private Sub Op41()
  Dim TV As Double
  
  TV = Fix(DisplayReg)
  If TV < 0# Or TV > 4# Then
    ForcError "Op37 range is 1-4 (1[.5 sec], 2[1.0 sec], 3[1.5 sec], and 4[2.0 sec]), 0 resets"
  Else
    tmrWait = CInt(TV)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Op42
' Purpose           : Decimal Display Register to internal string storage
'*******************************************************************************
Private Sub Op42()
  DspTxt = CvtTyp(TypDec)
  DisplayText = True
End Sub

'*******************************************************************************
' Subroutine Name   : Op43
' Purpose           : EE Scientific Notation Display Register to internal string storage
'*******************************************************************************
Private Sub Op43()
  DspTxt = Format(DisplayReg, ScientifEE)
  DisplayText = True
End Sub

'*******************************************************************************
' Subroutine Name   : Op44
' Purpose           : Hex Display Register to internal string storage
'*******************************************************************************
Private Sub Op44()
  DspTxt = CvtTyp(TypHex)
  DisplayText = True
End Sub

'*******************************************************************************
' Subroutine Name   : Op45
' Purpose           : Octal Display Register to internal string storage
'*******************************************************************************
Private Sub Op45()
  DspTxt = CvtTyp(TypOct)
  DisplayText = True
End Sub

'*******************************************************************************
' Subroutine Name   : Op46
' Purpose           : Binary Display Register to internal string storage
'*******************************************************************************
Private Sub Op46()
  DspTxt = CvtTyp(TypBin)
  DisplayText = True
End Sub

'*******************************************************************************
' Subroutine Name   : Op47
' Purpose           : Engineering Notation Display Register to internal string storage
'*******************************************************************************
Private Sub Op47()
  DspTxt = CvtEng
  DisplayText = True
End Sub

'*******************************************************************************
' Subroutine Name   : Op48
' Purpose           : Set Date Display format
'*******************************************************************************
Private Sub Op48()
  DateFmt = DspTxt
  DisplayText = True
End Sub

'*******************************************************************************
' Subroutine Name   : Op49
' Purpose           : Get the current date (number of days since Dec 30, 1899)
'                   : and time in seconds from midnight. the format is "days.time",
'                   : where the fractional time is the time in seconds / (60x60x24)
'*******************************************************************************
Private Sub Op49()
  DisplayReg = CDbl(Now)
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op50
' Purpose           : Display the date in current date format
'*******************************************************************************
Private Sub Op50()
  DspTxt = Format(CDate(DisplayReg), DateFmt)
  DisplayText = True
  If Not RunMode Then
    Call DisplayLine
    DisplayText = False
    Call NewLine                  'advance to a new line if not RUN mode
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Op51
' Purpose           : Get the Month from a Date value
'*******************************************************************************
Private Sub Op51()
  DisplayReg = CDbl(Month(CDate(DisplayReg)))
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op52
' Purpose           : Get Day of Month from a Date value
'*******************************************************************************
Private Sub Op52()
  DisplayReg = CDbl(Day(CDate(DisplayReg)))
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op53
' Purpose           : Get Year from a Date value
'*******************************************************************************
Private Sub Op53()
  DisplayReg = CDbl(Year(CDate(DisplayReg)))
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op54
' Purpose           : From a date value, Return 1 if a leapyear, 0 if not
'*******************************************************************************
Public Sub Op54()
  Dim Bol As Boolean
  
  Bol = IsLeapYear(CDate(DisplayReg))
  If Bol Then
    DisplayReg = 1#
  Else
    DisplayReg = 0#
  End If
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op55
' Purpose           : Return the Day Of Week from a date value (1=Sunday, 2 = Monday, etc)
'*******************************************************************************
Private Sub Op55()
  On Error Resume Next
  DisplayReg = Weekday(CDate(DisplayReg))
  Call CheckError
  On Error GoTo 0
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op56
' Purpose           : Get day of year for a date value
'*******************************************************************************
Private Sub Op56()
  Dim i As Integer, K As Integer, Ly As Integer
  Dim Dt As Date
  
  Dt = CDate(DisplayReg)
  Ly = 0                      'compute leapyear for February
  If IsLeapYear(Dt) Then Ly = 1
  K = CInt(Day(Dt))           'init date to current day of month
  
  For i = 1 To CInt(Month(Dt)) - 1 'loop through each preceding month from current
    Select Case i
      Case 1, 3, 5, 7, 8, 10  'Jan, Mar, May, Jul, Aug, Oct
        K = K + 31
      Case 2                  'Feb
        K = K + 28 + Ly
      Case Else               'Apr, Jun, Sep, Nov
        K = K + 30
    End Select
  Next i
  DisplayReg = CDbl(K)        'return day of year
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op57
' Purpose           : Get Week of Year from a date value
'*******************************************************************************
Private Sub Op57()
  Call Op56                                 'set displayreg to day of year
  DisplayReg = Fix(DisplayReg / 7#) + 1#    'compute week number, from 1
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op58
' Purpose           : Get number of days in the current month
'*******************************************************************************
Private Sub Op58()
  DisplayReg = CDbl(GetLastDayOfMonth(CDate(DisplayReg)))
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op59
' Purpose           : From a date value, Return 1 if a weekend, 0 if not
'*******************************************************************************
Private Sub Op59()
  If IsWeekend(CDate(DisplayReg)) Then
    DisplayReg = 1#
  Else
    DisplayReg = 0#
  End If
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op60
' Purpose           : Add (or subtract if negative) the number of months in the
'                   : Display Register to a date value in the Test Register
'*******************************************************************************
Private Sub Op60()
  On Error Resume Next
  DisplayReg = CDbl(AddMonths(CDate(TestReg), CInt(DisplayReg)))
  Call CheckError
  On Error GoTo 0
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op61
' Purpose           : Add (or subtract if negative), the number of years in the
'                   : Display Register to a date value in the Test Register
'*******************************************************************************
Private Sub Op61()
  On Error Resume Next
  DisplayReg = CDbl(AddYears(CDate(TestReg), CInt(DisplayReg)))
  Call CheckError
  On Error GoTo 0
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op62
' Purpose           : Get Full Month name from a date value
'*******************************************************************************
Private Sub Op62()
  DspTxt = Format(CDate(DisplayReg), "mmmm")
  If Not RunMode Then
    With frmVisualCalc.lstDisplay
      .List(.ListIndex) = DspTxt
    End With
    Call NewLine   'advance to a new line
  End If
  DisplayText = True
End Sub

'*******************************************************************************
' Subroutine Name   : Op63
' Purpose           : Set Time Display format
'*******************************************************************************
Private Sub Op63()
  TimeFmt = DspTxt
  DisplayText = True
End Sub

'*******************************************************************************
' Subroutine Name   : Op64
' Purpose           : display the time from a date value in the current time format to the display text
'*******************************************************************************
Private Sub Op64()
  DspTxt = Format(CDate(DisplayReg), TimeFmt)
  DisplayText = True
  If Not RunMode Then
    Call DisplayLine
    DisplayText = False
    Call NewLine   'advance to a new line
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Op65
' Purpose           : Get Hours from a date value
'*******************************************************************************
Private Sub Op65()
  DisplayReg = CDbl(Hour(CDate(DisplayReg)))
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op66
' Purpose           : Get Minutes from a date value
'*******************************************************************************
Private Sub Op66()
  On Error Resume Next
  DisplayReg = CDbl(Minute(CDate(DisplayReg)))
  Call CheckError
  On Error GoTo 0
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op67
' Purpose           : Get Seconds from a date value
'*******************************************************************************
Private Sub Op67()
  On Error Resume Next
  DisplayReg = CDbl(Second(CDate(DisplayReg)))
  Call CheckError
  On Error GoTo 0
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op68
' Purpose           : Add (or subtract if negative), the number of Hours in the
'                   : T-reg to a date value in the display
'*******************************************************************************
Private Sub Op68()
  Dim Dt As Date
  
  Dt = CDate(DisplayReg)
  On Error Resume Next
  DisplayReg = CDbl(TimeSerial(Hour(Dt) + CInt(TestReg), Minute(Dt), Second(Dt)))
  Call CheckError
  On Error GoTo 0
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op69
' Purpose           : Add (or subtract if negative), the number of Minutes in the
'                   : T-reg to a date value in the display
'*******************************************************************************
Private Sub Op69()
  Dim Dt As Date
  
  Dt = CDate(DisplayReg)
  On Error Resume Next
  DisplayReg = CDbl(TimeSerial(Hour(Dt), Minute(Dt) + CInt(TestReg), Second(Dt)))
  Call CheckError
  On Error GoTo 0
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op70
' Purpose           : Add (or subtract if negative), the number of Seconds in the
'                   : T-reg to a date value in the display
'*******************************************************************************
Private Sub Op70()
  Dim Dt As Date
  
  Dt = CDate(DisplayReg)
  On Error Resume Next
  DisplayReg = CDbl(TimeSerial(Hour(Dt), Minute(Dt), Second(Dt)) + CInt(TestReg))
  Call CheckError
  On Error GoTo 0
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op71
' Purpose           : Set plot point size (8-24; default is 10)
'*******************************************************************************
Private Sub Op71()
  Dim TV As Double
  
  TV = Abs(Fix(DisplayReg))
  If TV < 8 Or TV > 24 Then
    ForcError "Point size range is 8 through 24"
  Else
    frmVisualCalc.lblChkSize.FontSize = CLng(TV)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Op72
' Purpose           : Sets drawing color value (default is black;
'                   : use RGB to obtain color values)
'*******************************************************************************
Private Sub Op72()
  Dim TV As Double
  Dim TI As Long
  
  TV = Fix(DisplayReg)
  On Error Resume Next
  TI = CLng(TV)
  Call CheckError
  On Error GoTo 0
  If ErrorFlag Then Exit Sub
  frmVisualCalc.PicPlot.ForeColor = TI
  PlotColor = TI
End Sub

'*******************************************************************************
' Subroutine Name   : Op73
' Purpose           : Get Text line height in pixels
'*******************************************************************************
Private Sub Op73()
  DisplayReg = CDbl(LineHeight)
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op74
' Purpose           : Set Text line height in pixels (0-32; reset when point size set)
'*******************************************************************************
Private Sub Op74()
  Dim TV As Double
  
  TV = Fix(DisplayReg)
  If TV < 0 Or TV > 32 Then
    ForcError "Allowed range is 0-32"
  Else
    LineHeight = CLng(TV)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Op75
' Purpose           : Sets the drawing mode (0=Copy, 1=OR,
'                   : 2=AND, 3=XOR, 4=NXOR, 5=NOT, 6=Solid Line,
'                   : 7=Dash Line, 8=Dash-Dot Line, 9=Dash-Dot-Dot Line.
'                   : States 0 (Copy) and 6 (Solid line) are default).
'                   : -1 thru -4 sets the draw width to 1 to 4, respectively.
'*******************************************************************************
Private Sub Op75()
  Dim TV As Double
  
  TV = Fix(DisplayReg)
  If TV < -4 Or TV > 9 Then
    ForcError "Allowed range is 0-9"
  Else
    With frmVisualCalc.PicPlot
      Select Case CInt(TV)
        Case Is < 0 'draw width 1-4
          .DrawWidth = Abs(CInt(TV))
        Case 0  'Copy
          .DrawMode = vbCopyPen
        Case 1  'OR
          .DrawMode = vbMergePen
        Case 2  'AND
          .DrawMode = vbMaskPen
        Case 3  'XOR
          .DrawMode = vbXorPen
        Case 4  'NXOR
          .DrawMode = vbNotXorPen
        Case 5  'NOT
          .DrawMode = vbNotCopyPen
        Case 6  'Solid Line
          .DrawStyle = 1
        Case 7  'Dash Line
          .DrawStyle = 2
        Case 8  'Dash-Dot Line
          .DrawStyle = 3
        Case 9  'Dash-Dot-Dot Line
          .DrawStyle = 4
      End Select
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Op76
' Purpose           : Get last X positions cursor was over (provided by click event)
'*******************************************************************************
Private Sub Op76()
  DisplayReg = CDbl(LastPlotX - PlotXOfst)
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op77
' Purpose           : Get last Y positions cursor was over (provided by click event)
'*******************************************************************************
Private Sub Op77()
  DisplayReg = CDbl(LastPlotY - PlotYOfst)
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op78
' Purpose           : Get percentage value of display register
'*******************************************************************************
Private Sub Op78()
  Select Case PendIdx
    Case 0
      DisplayReg = DisplayReg / 100#
    Case Else
      If DisplayReg = 0# Or PendValue(PendIdx) = 0# Then
        DisplayReg = 0#
      Else
        Select Case PendOpn(PendIdx)
          Case iAdd     'add-on percentage
            DisplayReg = (100# + DisplayReg) / 100# * PendValue(PendIdx)
          Case iMinus   'subtract percentage
            DisplayReg = (100# - DisplayReg) / 100# * PendValue(PendIdx)
          Case iMult    'Compute % increase
            DisplayReg = (PendValue(PendIdx) + DisplayReg) / DisplayReg * 100#
          Case iDVD     'compute % of
            DisplayReg = DisplayReg / PendValue(PendIdx) * 100#
          Case Else
            ForcError "Unrecognized percentage operator"
            Exit Sub
        End Select
      End If
      PendIdx = PendIdx - 1
  End Select
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op79
' Purpose           : Compute Permutations. Compute permuting 'r' different
'                   : members among 'n' members, where n>=r: nPr = n!÷(n-r)!
'                   : Entry Method: "n [X><T] r OP 79"
'*******************************************************************************
Private Sub Op79()
  Dim TN As Double, Tnr As Double
  
  If TestReg < DisplayReg Then
    ForcError "'n' parameter must be equal to or greater than 'r'"
    Exit Sub
  ElseIf TestReg <= 0# Then
    ForcError "'n' parameter cannot equal to or less than '0'"
    Exit Sub
  ElseIf TestReg > 69# Then
    ForcError "'n' cannot exceed 69"
    Exit Sub
  End If
  
  TN = TestReg                        'get n
  Tnr = TN - DisplayReg               'get n-r
  If Not Factorial(TN) Then Exit Sub  'get n!
  If Not Factorial(Tnr) Then Exit Sub 'get (n-r)!
  DisplayReg = TN / Tnr               'compute n!/(n-r)!
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op80
' Purpose           : Compute Combinations. Compute how many groups of 'r' members
'                   : can be obtained when there are 'n' total members,
'                   : where n>=r: nCr = n!÷(r!x(n-r)!)
'                   : Entry Method: "n [X><T] r OP 80"
'*******************************************************************************
Private Sub Op80()
  Dim TN As Double, Tr As Double, Tnr As Double
  
  If TestReg < DisplayReg Then
    ForcError "'n' parameter must be equal to or greater than 'r'"
    Exit Sub
  ElseIf TestReg <= 0# Then
    ForcError "'n' parameter cannot equal to or less than '0'"
    Exit Sub
  ElseIf TestReg > 69# Then
    ForcError "'n' cannot exceed 69"
    Exit Sub
  End If
  TN = TestReg                        'get n
  Tr = DisplayReg                     'get r
  Tnr = TN - Tr                       'get n-r
  If Not Factorial(TN) Then Exit Sub  'get n!
  If Not Factorial(Tr) Then Exit Sub  'get r!
  If Not Factorial(Tnr) Then Exit Sub 'get (n-r)!
  DisplayReg = TN / (Tr * Tnr)        'compute n!/r!(n-r)!
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op81
' Purpose           : Convert floating point decimal value to fraction
'*******************************************************************************
Private Sub Op81()
  DspTxt = DecToFraction(DisplayReg)
  DisplayText = True
End Sub

'*******************************************************************************
' Subroutine Name   : Op82
' Purpose           : Convert the value in the Display Register from
'                   : the current Angle Type to Degrees.
'*******************************************************************************
Private Sub Op82()
  On Error Resume Next
  Select Case AngleType
    Case TypDeg
    Case TypRad
      DisplayReg = DisplayReg * 180# / vPi
    Case TypGrad
      DisplayReg = DisplayReg * 0.9
    Case TypMil
      DisplayReg = DisplayReg * 0.05625
  End Select
  Call CheckError
  On Error GoTo 0
End Sub

'*******************************************************************************
' Subroutine Name   : Op83
' Purpose           : Convert the value in the Display Register from
'                   : the current Angle Type to Radians.
'*******************************************************************************
Private Sub Op83()
  On Error Resume Next
  Select Case AngleType
    Case TypDeg
      DisplayReg = DisplayReg * vPi / 180#
    Case TypRad
    Case TypGrad
      DisplayReg = DisplayReg * vPi / 200#
    Case TypMil
      DisplayReg = DisplayReg * vPi / 3200#
  End Select
  Call CheckError
  On Error GoTo 0
End Sub

'*******************************************************************************
' Subroutine Name   : Op84
' Purpose           : Convert the value in the Display Register from
'                   : the current Angle Type to Grads.
'*******************************************************************************
Private Sub Op84()
  On Error Resume Next
  Select Case AngleType
    Case TypDeg
      DisplayReg = DisplayReg / 0.9
    Case TypRad
      DisplayReg = DisplayReg * 200# / vPi
    Case TypGrad
    Case TypMil
      DisplayReg = DisplayReg / 16#
  End Select
  Call CheckError
  On Error GoTo 0
End Sub

'*******************************************************************************
' Subroutine Name   : Op85
' Purpose           : Convert the value in the Display Register from
'                   : the current Angle Type to Mils.
'*******************************************************************************
Private Sub Op85()
  On Error Resume Next
  Select Case AngleType
    Case TypDeg
      DisplayReg = DisplayReg / 0.05625
    Case TypRad
      DisplayReg = DisplayReg * 3200# / vPi
    Case TypGrad
      DisplayReg = DisplayReg * 16
    Case TypMil
  End Select
  Call CheckError
  On Error GoTo 0
End Sub

'*******************************************************************************
' Subroutine Name   : Op86
' Purpose           : Set display register to length of internal string storage
'*******************************************************************************
Private Sub Op86()
  DisplayReg = CDbl(Len(DspTxt))
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op87
' Purpose           : Set display register to a physical constant
'*******************************************************************************
Private Sub Op87()
  Dim i As Long
  
  On Error Resume Next
  i = CLng(DisplayReg)        'number can become integer?
  Call CheckError             'check for overflow
  On Error GoTo 0             'turn off error trap
  If ErrorFlag Then Exit Sub  'exit if error
  
  Select Case i
    Case 0  'F    Fraday Constant:          9.648670E+07 C k mole-¹
      DisplayReg = 96486700#
    Case 1  'c    Speed of Light:           2.9979250E+08 m/sec-¹
      DisplayReg = 299792500#
    Case 2  'e    Electron Charge:          1.6021917E-10 C
      DisplayReg = 1.6021917E-10
    Case 3  'N    Avogado Number:           6.022169E+26 k mole-¹
      DisplayReg = 6.022169E+26
    Case 4  '4     eV   Electron Volt:      1.602E-19 J
      DisplayReg = 1.602E-19
    Case 5  'me   Electron Rest Mass:       9.109558E-31 kg
      DisplayReg = 9.109558E-31
    Case 6  'Mp   Proton Rest Mass:         1.672614E-27 kg
      DisplayReg = 1.672614E-27
    Case 7  'Mn   Neutron Rest Mass:        1.674920E-27 kg
      DisplayReg = 1.67492E-27
    Case 8  'amu  Atomic Mass Unit:         1.660531E-27 kg
      DisplayReg = 1.660531E-27
    Case 9  'e/me Electron Charge to Mass ratio: 1.7588028E+11 C kg-¹
      DisplayReg = 175880280000#
    Case 10 'h    Planck Constant:          6.626196E-34 J-sec
      DisplayReg = 6.626196E-34
    Case 11 'Roo  RydBerg Constant:         1.09737312E+07 m-¹
      DisplayReg = 10973731.2
    Case 12 'Ro   Gas Constant:             8.31434E+03 J-k mole-¹ K-¹
      DisplayReg = 8314.34
    Case 13 'k    Boltzmann Constant:       1.380622E-23 JK-¹
      DisplayReg = 1.380622E-23
    Case 14 'G    Gravitational Constant:   6.6732E-11 N-m²kg-²
      DisplayReg = 0.000000000066732
    Case 15 'µb   Bohr Magaton:             9.274096E-24 JT-¹
      DisplayReg = 9.274096E-24
    Case 16 'µe   Electron Magnetic Moment: 9.284851E-24 JT-¹
      DisplayReg = 9.284851E-24
    Case 17 'µp   Proton Magnetic Moment:   1.4106203E-24 JT-¹
      DisplayReg = 1.4106203E-24
    Case 18 'lc   Compton Wavelength of the Electron: 2.4263096E-26 m
      DisplayReg = 2.4263096E-26
    Case 19 'lc.p Compton Wavelength of the Proton:   1.3214409E-15 m
      DisplayReg = 1.3214409E-15
    Case 20 'lc.n Compton Wavelength of the Neutron:  1.3196217E-15 m
      DisplayReg = 1.3196217E-15
    Case 21 'o    Stefan-Boltzmann Constant: 5.56704E-08 W/m2-K4
      DisplayReg = 0.0000000556704
    Case 22 'ao    Bohr Radius:  5.291772108E-11 m
      DisplayReg = 5.291772108E-11
    Case Else
      ForcError "Pysical Constant value is out of range (0-22)"
  End Select
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op88
' Purpose           : Date Finder. This is a special date finder routine that
'                   : returns a date in VisualCalc date format (see OP 49), as
'                   : the number of days since Dec 30, 1899. The provided text
'                   : format is: Year[,Month[,[Day[,Week[,Weekday]]]]]
'                   : Special Ranges: Week (1-4), WeekDay (1-7; 1=Sunday)
'                   : Allowed Input Samples:
'                   :   Return date for Jan 1,1987              : 1987
'                   :   Return date for March 1,1987            : 1987,3
'                   :   Return date for July 4,1987             : 1987,7,4
'                   :   Return date for June,1987, Week 2       : 1987,6,0,2
'                   :   Return date for Monday, May,1987, Week 3: 1987,5,0,3,2
'*******************************************************************************
Public Sub Op88()
  Dim X As Integer, Y As Integer
  Dim S As String
  Dim yr As Integer, Mn As Integer, Dy As Integer, wk As Integer, WkDy As Integer
  
  S = Trim$(DspTxt)                                       'get data to process
  DisplayText = False
  Do
    If Len(S) = 0 Then Exit Do                            'if nothing to process
'
' get year
'
    On Error Resume Next
    X = InStr(1, S, ",")                                  'find comma
    If Not CBool(X) Then                                  'just year
      DisplayReg = CDbl(DateSerial(CInt(S), 1, 1))        'get Jan 1
      If CBool(Err.Number) Then Exit Do
      Exit Sub                                            'return result
    End If
    yr = Left$(S, X - 1)
    S = LTrim$(Mid$(S, X + 1))
'
' get month
'
    X = InStr(1, S, ",")                                  'find comma
    If Not CBool(X) Then
      DisplayReg = CDbl(DateSerial(yr, CInt(S), 1))       'compute Mn/01/yr
      If CBool(Err.Number) Then Exit Do
      Exit Sub                                            'return date
    End If
    Mn = Left$(S, X - 1)                                  'grab motnth
    If Mn < 1 Or Mn > 12 Then Exit Do                     'out of bounds
    S = LTrim$(Mid$(S, X + 1))
'
' get day
'
    X = InStr(1, S, ",")                                  'find comma
    If Not CBool(X) Then
      Dy = CInt(S)                                        'grab day
      If CBool(Err.Number) Then Exit Do
      If Dy < 0 Or Dy > 31 Then Exit Do                   'initial check
      If CBool(Dy) Then                                   'if day specified
        If Not IsDate(CStr(Mn) & "/" & CStr(Dy) & "/" & CStr(yr)) Then Exit Do
        DisplayReg = CDbl(DateSerial(yr, Mn, Dy))         'get date
        If CBool(Err.Number) Then Exit Do
        Exit Sub                                          'return date
      End If
    End If
    Dy = Left$(S, X - 1)                                  'get day
    S = LTrim$(Mid$(S, X + 1))
    If CBool(Dy) Then
      ForcError "Day must be 0 when you will also specify a Week"
      Exit Sub
    End If
    Dy = 1                                                'assume 1st of month
'
' find Sunday to start first full week
'
    Do While Weekday(DateSerial(yr, Mn, Dy)) <> 1         'while not Sumday (1)
      Dy = Dy + 1                                         'bump a day
    Loop
'
' get week
'
    X = InStr(1, S, ",")                                  'find comma
    If Not CBool(X) Then
      wk = CInt(S)                                        'grab week #
      If CBool(Err.Number) Then Exit Do
      S = vbNullString
    Else
      wk = Left$(S, X - 1)                                'grab week #
      S = LTrim$(Mid$(S, X + 1))
    End If
    If wk < 1 Or wk > 5 Then Exit Do                      'error if not 1-5
    Dy = Dy - 7                                           'init offset
'
' point to Sunday of Desired week
'
    Do While CBool(wk)
      Dy = Dy + 7
      wk = wk - 1
    Loop
'
' if we possibly went beyond month with week 5, then back up to legal Sunday
'
    If Not IsDate(CStr(Mn) & "/" & CStr(Dy) & "/" & CStr(yr)) Then
      Dy = Dy - 7
    End If
'
' get weekday
'
    If CBool(Len(S)) Then
      WkDy = CInt(S)                                      'grab weekdat
      If CBool(Err.Number) Then Exit Do
      If WkDy < 1 Or WkDy > 7 Then Exit Do                'in range?
      Do While Weekday(DateSerial(yr, Mn, Dy)) <> WkDy    'point to weekday
        Dy = Dy + 1
      Loop
    End If
'
' return final date
'
    DisplayReg = CDbl(DateSerial(yr, Mn, Dy))
    Exit Sub
  Loop
  ForcError "Required format: yr[,mn[,[dy[,wk[,wkdy]]]]]"     'error
End Sub

'*******************************************************************************
' Subroutine Name   : OP89
' Purpose           : 'Provide conversion factors for translating one unit of
'                   : measure to another. The result is multiplied by the value to
'                   : convert to obtain the value of the desired unit of measure.
'*******************************************************************************
Private Sub Op89()
  Dim i As Long
  
  On Error Resume Next
  i = CLng(DisplayReg)        'number can become integer?
  Call CheckError             'check for overflow
  On Error GoTo 0             'turn off error trap
  If ErrorFlag Then Exit Sub  'exit if error
  
  Select Case i
    'Convert acres to square feet
    Case 0: DisplayReg = 43560
    'Convert acres to square miles
    Case 1: DisplayReg = 0.0015625
    'Convert angstroms to centimeters
    Case 2: DisplayReg = 0.00000001
    'Convert angstroms to inches
    Case 3: DisplayReg = 254000000#
    'Convert angstroms to meters
    Case 4: DisplayReg = 0.0000000001
    'Convert ast unit to kilometers
    Case 5: DisplayReg = 149500000#
    'Convert ast unit to miles
    Case 6: DisplayReg = 92894993.2394814
    'Convert board ft to cubic feet
    Case 7: DisplayReg = 8.33333333333333E-02
    'Convert bushels to cubic cm
    Case 8: DisplayReg = 35239.07
    'Convert cord ft to cords
    Case 9: DisplayReg = 0.125
    'Convert coft ft to cubic feet
    Case 10: DisplayReg = 16
    'Convert centimeters to angstroms
    Case 11: DisplayReg = 100000000#
    'Convert centimeters to feet
    Case 12: DisplayReg = 3.28083989501312E-02
    'Convert centimeters to inches
    Case 13: DisplayReg = 0.393700787401575
    'Convert centimeters to kilometers
    Case 14: DisplayReg = 0.00001
    'Convert centimeters to meters
    Case 15: DisplayReg = 0.01
    'Convert centimeters to miles
    Case 16: DisplayReg = 6.2137119224E-06
    'Convert centimeters to yards
    Case 17: DisplayReg = 1.09361329833771E-02
    'Convert cords to cord feet
    Case 18: DisplayReg = 8
    'Convert cubic cm to cubic inches
    Case 10: DisplayReg = 6.10237440947323E-02
    'Convert cubic cm to cubic metes
    Case 20: DisplayReg = 0.000001
    'Convert cubic cm to cubic yards
    Case 21: DisplayReg = 1.3079506193E-06
    'Convert cubic feet to board feet
    Case 22: DisplayReg = 12
    'Convert cubic feet to cord feet
    Case 23: DisplayReg = 0.0625
    'Convert cubic feet to cubic inches
    Case 24: DisplayReg = 1728
    'Convert cubic feet to cubic meters
    Case 25: DisplayReg = 0.028316846592
    'Convert cubic feet to cubic yards
    Case 26: DisplayReg = 0.037037037037037
    'Convert cubic inches to cubic cm
    Case 27: DisplayReg = 16.387064
    'Convert cubic inches to cubic feet
    Case 28: DisplayReg = 5.787037037037E-04
    'Convert cubic inches to cubic yards
    Case 29: DisplayReg = 2.14334705075E-05
    'Convert cubic meters to cubic cm
    Case 30: DisplayReg = 1000000#
    'Convert cubic meters to cubic feet
    Case 31: DisplayReg = 35.3146667214886
    'Convert cubic meters to cubic yards
    Case 32: DisplayReg = 1.30795061931439
    'Convert cubic yards to cubic meters
    Case 33: DisplayReg = 764554.857984
    'Convert cubic yards to cubic feet
    Case 34: DisplayReg = 27
    'Convert cubic yards to cubic inches
    Case 35: DisplayReg = 46656
    'Convert cubic yards to cubic meters
    Case 36: DisplayReg = 0.764554857984
    'Convert ft per sec to miles per hr
    Case 37: DisplayReg = 0.681818181818182
    'Convert feet to centimeters
    Case 38: DisplayReg = 30.48
    'Convert feet to inches
    Case 39: DisplayReg = 12
    'Convert feet to kilometers
    Case 40: DisplayReg = 0.0003048
    'Convert feet to meters
    Case 41: DisplayReg = 0.3048
    'Convert feet to miles
    Case 42: DisplayReg = 1.893939393939E-04
    'Convert feet to rods
    Case 43: DisplayReg = 6.06060606060606E-02
    'Convert feet to yards
    Case 44: DisplayReg = 0.333333333333333
    'Convert inches to centimeters
    Case 45: DisplayReg = 2.54
    'Convert inches to feet
    Case 46: DisplayReg = 8.33333333333333E-02
    'Convert inches to metes
    Case 47: DisplayReg = 0.0254
    'Convert inches to miles
    Case 48: DisplayReg = 1.57828282828E-05
    'Convert inches to yards
    Case 49: DisplayReg = 2.77777777777778E-02
    'Convert kilometers to ast units
    Case 50: DisplayReg = 149500000#
    'Convert kilometers to centimeters
    Case 51: DisplayReg = 100000#
    'Convert kilometers to feet
    Case 52: DisplayReg = 3280.83989501312
    'Convert kilometers to meters
    Case 53: DisplayReg = 1000#
    'Convert kilometers to miles
    Case 54: DisplayReg = 0.621371192237334
    'Convert kilometers to rods
    Case 55: DisplayReg = 198.838781515947
    'Convert miles to ast units
    Case 56: DisplayReg = 1.07648428E-08
    'Convert miles to centimeters
    Case 57: DisplayReg = 160934.4
    'Convert miles to feet
    Case 58: DisplayReg = 5280
    'Convert miles to inches
    Case 59: DisplayReg = 63360
    'Convert miles to kilometers
    Case 60: DisplayReg = 1.609344
    'Convert miles to meters
    Case 61: DisplayReg = 1609.344
    'Convert miles to rods
    Case 62: DisplayReg = 320
    'Convert miles to yards
    Case 63: DisplayReg = 1760#
    'Convert miles per hr to ft per sec
    Case 64: DisplayReg = 1.46666666666667
    'Convert meters to angstroms
    Case 65: DisplayReg = 10000000000#
    'Convert meters to centimetes
    Case 66: DisplayReg = 100
    'Convert meters to feet
    Case 67: DisplayReg = 3.28083989501312
    'Convert meters to inches
    Case 68: DisplayReg = 39.3700787401575
    'Convert meters to kilometers
    Case 69: DisplayReg = 0.001
    'Convert meters to miles
    Case 70: DisplayReg = 6.213711922373E-04
    'Convert meters to rods
    Case 71: DisplayReg = 0.198838781515947
    'Convert meters to yards
    Case 72: DisplayReg = 1.09361329833771
    'Convert rods to feet
    Case 73: DisplayReg = 16.5
    'Convert rods to kilometers
    Case 74: DisplayReg = 0.0050292
    'Convert rods to meters
    Case 75: DisplayReg = 5.0292
    'Convert rods to miles
    Case 76: DisplayReg = 0.003125
    'Convert rods to yards
    Case 77: DisplayReg = 5.5
    'Convert square ft to sq in
    Case 78: DisplayReg = 144
    'Convert square ft to sq miles
    Case 79: DisplayReg = 2.29568411387E-05
    'Convert square in to sq ft
    Case 80: DisplayReg = 6.9444444444444E-03
    'Convert square in to sq mi
    Case 81: DisplayReg = 2.490977E-10
    'Convert square mi to acres
    Case 82: DisplayReg = 640
    'Convert square mi to sq ft
    Case 83: DisplayReg = 27878400#
    'Convert square mi to sq in
    Case 84: DisplayReg = 4014489600#
    'Convert yards to centimeters
    Case 85: DisplayReg = 91.44
    'Convert yards to feet
    Case 86: DisplayReg = 3
    'Convert yards to inches
    Case 87: DisplayReg = 36
    'Convert yards to meters
    Case 88: DisplayReg = 0.9144
    'Convert yards to miles
    Case 89: DisplayReg = 5.681818181818E-04
    'Convert yards to rods
    Case 90: DisplayReg = 0.181818181818182
    Case Else
      ForcError "Conversion factor index is out of range (0-90)"
  End Select
End Sub

'*******************************************************************************
' Subroutine Name   : Op90
' Purpose           : Compute Binomial Coefficient.
'                   : Binomial Coefficient = n! / (j!(n-j)!
'                   : N is in test register, J is in the Display Register
'*******************************************************************************
Private Sub Op90()
  Dim Idx As Long, TJ As Long, TN As Long, i As Long
  Dim TV As Double

  On Error Resume Next
  TJ = CLng(DisplayReg)                     'get j
  TN = CLng(TestReg)                        'get n
  
  TV = 1#
  For Idx = 0 To TJ - 1                     'compute Binomial Coefficient
    TV = TV * CDbl((TN - Idx)) / CDbl((TJ - Idx))
    If CBool(Err.Number) Then Exit For      'if error
  Next Idx
  
  Call CheckError                           'check for error
  On Error GoTo 0                           'disentagle from error trapping
  If Not ErrorFlag Then                     'if no error
    DisplayReg = TV                         'assign result to Display Register
    DisplayText = False                     'ensure it will be displayed
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Op91
' Purpose           : Provided the Y value in the Test Register and
'                   : the X value in the Display register, return the
'                   : angle from the X axis at point (y,x).
'*******************************************************************************
Private Sub Op91()
  DisplayReg = RadToAng(ATan2(TestReg, DisplayReg))
  DisplayText = False
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************
