Attribute VB_Name = "modStatLinear"
Option Explicit
'----- Statistical and Linear Analysis -----------------------------------------
' Op13  Initialize the Statistical and Linear Regression registers.
' Op14  Return SUM y to the Display Register.
' Op15  Return SUM y² to the Display Register.
' Op16  Return N (Number of entries) to the Display Register.
' Op17  Return SUM x to the Display Register.
' Op18  Return SUM x² to the Display Register.
' Op19  Return SUM xy to the Display Register.
' Op20  Calculate Correlation Coefficient to the Display Register.
' Op21  Calculate linear estimate y' against entered x' to the Display Register.
' Op22  Calculate linear estimate x' against entered y' to the Display Register.
' StatSUMadd: E+
' StatSUMsub: E-
' StatMean  : Compute statistical Mean
' StatVarnc : Compute Variance using N weighting
'           : Derive Standard Deviation for N weighting using: StDev X²
' StatStdDev: Compute Standard Deviation using N-1 weighting
'           : Derive Variance for N-1 weighting using: StDev Sqrt
' CalcSlope : Calculate slope (m) and return the value
' Yintercept: Calculate y-intercept (b) and slope (m)

'*******************************************************************************
' Private statistical variables
'*******************************************************************************
Private m_Ey As Double            'SUM y
Private m_Ey2 As Double           'SUM y²
Private m_N As Double             'N
Private m_Ex As Double            'SUM x
Private m_Ex2 As Double           'SUM x²
Private m_Exy As Double           'SUM x*y

'*******************************************************************************
' Subroutine Name   : Op13
' Purpose           : Init Statistical Mode
'*******************************************************************************
Public Sub Op13()
  If LrnMode Then Exit Sub
  
  m_Ey = 0#                       'erase registers
  m_Ey2 = 0#
  m_N = 0#
  m_Ex = 0#
  m_Ex2 = 0#
  m_Exy = 0#
  
  If EEMode Then
    EEMode = False                'turn off Enter Exponent display
    EngMode = False               'disable Eng mode
    Call UpdateStatus             'reflect on display
  End If
  PendIdx = 0                     'reset pending operations
  PndIdx = 0
  DisplayText = False
End Sub

'*******************************************************************************
' return statistical register values
'*******************************************************************************
Public Sub Op14()
  DisplayReg = m_Ey   'return SUM y
  DisplayText = False
End Sub

Public Sub Op15()
  DisplayReg = m_Ey2  'return SUM y Squared
  DisplayText = False
End Sub

Public Sub Op16()
  DisplayReg = m_N    'return number of items summed
  DisplayText = False
End Sub

Public Sub Op17()
  DisplayReg = m_Ex   'return SUM x
  DisplayText = False
End Sub

Public Sub Op18()
  DisplayReg = m_Ex2  'return SUM x Squared
  DisplayText = False
End Sub

Public Sub Op19()
  DisplayReg = m_Exy  'return x * y
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op20
' Purpose           : Calculate Correlation Coefficient (R)
'*******************************************************************************
Public Sub Op20()
  Dim SDx As Double, SDy As Double
  
  SDx = Sqr((m_Ex2 - m_Ex * m_Ex / m_N) / (m_N - 1))  'StdDevX
  SDy = Sqr((m_Ey2 - m_Ey * m_Ey / m_N) / (m_N - 1))  'StdDevY
  DisplayReg = CalcSlope() * SDx / SDy                'R = (m StdDevX) / StdDevY
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op21
' Purpose           : Calculate linear estimate of y'
'*******************************************************************************
Public Sub Op21()
  Dim m As Double
  
  m = CalcSlope
  DisplayReg = m * DisplayReg + (m_Ey - m * m_Ex) / m_N 'mx + b
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : Op22
' Purpose           : Calculate linear estimate of x'
'*******************************************************************************
Public Sub Op22()
  Dim b As Double
  
  b = (m_Ey - CalcSlope() * m_Ex) / m_N   'b = (Ey - m Ex) / N
  DisplayReg = (DisplayReg - b) / m_N     '(y - b) / m
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : StatSUMadd
' Purpose           : E+
'*******************************************************************************
Public Sub StatSUMadd()
  m_Ey = m_Ey + DisplayReg                  'perform general summation of y
  m_Ey2 = m_Ey2 + DisplayReg * DisplayReg   'sum y2 accumulator
  m_Ex = m_Ex + TestReg                     'perform general summation of x (optional)
  m_Ex2 = m_Ex2 + TestReg * TestReg         'sum x2 accumulator
  m_Exy = m_Exy + DisplayReg * TestReg      'Sum x*y
  m_N = m_N + 1#                            'Bump number of items entered (N)
  DisplayReg = m_N                          'display the entry number
  TestReg = 0#                              'reset Test Register
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : StatSUMsub
' Purpose           : E-
'*******************************************************************************
Public Sub StatSUMsub()
  If m_N = 0# Then
    ForcError "No items exist in the SUM list"
    Exit Sub
  End If
  m_Ey = m_Ey - DisplayReg
  m_Ey2 = m_Ey2 - DisplayReg * DisplayReg
  m_Ex = m_Ex - TestReg
  m_Ex2 = m_Ex2 - TestReg * TestReg
  m_Exy = m_Exy - DisplayReg * TestReg
  m_N = m_N - 1#
  DisplayReg = m_N
  TestReg = 0#
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : StatMean
' Purpose           : Compute statistical Mean
'*******************************************************************************
Public Sub StatMean()
  TestReg = m_Ex / m_N    'average X value to Test Register
  DisplayReg = m_Ey / m_N 'average Y value to the Display Register
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : StatVarnc (Statistical Variance)
' Purpose           : Compute Variance using N weighting
'                   : Derive Standard Deviation for N weighting using: StDev X²
'*******************************************************************************
Public Sub StatVarnc()
  Dim TV As Double
  
  TV = m_Ex / m_N                     'mean x
  TestReg = m_Ex2 / m_N - TV * TV     'SUMx2/n - (mean x)2
  TV = m_Ey / m_N                     'mean y
  DisplayReg = m_Ey2 / m_N - TV * TV  'SUMy2/n - (mean y)2
  DisplayText = False
End Sub

'*******************************************************************************
' Subroutine Name   : StatStdDev
' Purpose           : Compute Standard Deviation using N-1 weighting
'                   : Derive Variance for N-1 weighting using: StDev Sqrt
'*******************************************************************************
Public Sub StatStdDev()
  TestReg = Sqr((m_Ex2 - m_Ex * m_Ex / m_N) / (m_N - 1))
  DisplayReg = Sqr((m_Ey2 - m_Ey * m_Ey / m_N) / (m_N - 1))
  DisplayText = False
End Sub

'*******************************************************************************
' Function Name     : CalcSlope
' Purpose           : Calculate slope (m) and return the value
'*******************************************************************************
Private Function CalcSlope() As Double
  CalcSlope = (m_Exy - m_Ex * m_Ey / m_N) / (m_Ex2 - m_Ex * m_Ex / m_N)
End Function

'*******************************************************************************
' Subroutine Name   : Yintercept
' Purpose           : Calculate y-intercept (b) and slope (m)
'*******************************************************************************
Public Sub Yintercept()
  TestReg = CalcSlope()                       'm
  DisplayReg = (m_Ey - TestReg * m_Ex) / m_N  'b = (Ey - m Ex) / N
  DisplayText = False
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

