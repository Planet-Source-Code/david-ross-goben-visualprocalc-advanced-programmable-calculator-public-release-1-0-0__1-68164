Attribute VB_Name = "modPlot_Support"
Option Explicit

'-------------------------------------------------------------------------------
' API support for this module
'-------------------------------------------------------------------------------
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Long) As Long
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Const FLOODFILLBORDER = 0
Private Const FLOODFILLSURFACE = 1

'*******************************************************************************
' Subroutine Name   : PrintReset
' Purpose           : Reset Plot's Print functions
'*******************************************************************************
Public Sub PrintReset()
  With frmVisualCalc.lblChkSize
    .FontSize = 10                'set 10-point height (10*20/15=13 pixels)
    LineHeight = .Height \ Screen.TwipsPerPixelY 'default height for 10-point in pixels
  End With
  PlotX = PlotXDef
  PlotY = PlotYDef
  LastDir = 0#
End Sub

'*******************************************************************************
' Subroutine Name   : PlotClr
' Purpose           : Reset Plot's Screen
'*******************************************************************************
Public Sub PlotClr()
  With frmVisualCalc.PicPlot  'erase plot field
    .BackColor = vbWhite
    .ForeColor = vbBlack
    .Cls
    .DrawMode = vbCopyPen
    .ScaleMode = vbPixels
    .DrawStyle = vbSolid
    .DrawWidth = 1
  End With
  PlotColor = vbBlack         'default draw color
  PlotXDef = PlotXOfst        'offset from left
  PlotYDef = PlotYOfst        'offset from top
  PlotX = PlotXOfst
  PlotY = PlotYOfst
  LastPlotX = 0&              'less offset
  LastPlotY = 0&
  Call PrintReset
End Sub

'*******************************************************************************
' Function Name     : RGB_Support
' Purpose           : Provide run-time support for the RGB token
'*******************************************************************************
Public Function RGB_Support(ByVal Code As Integer, Txt As String, Errstr As String) As Integer
  Dim Vptr As clsVarSto
  
  If GetInstruction(1) <> Code Then             'check expected token
    Errstr = "Expected '" & Txt & "'"           'provide error
    RGB_Support = -1                            'error code
  Else
    IncInstrPtr                                 'point past '(' or ","
    If GetInstruction(1) = iIND Then
      Set Vptr = CheckVbl()                     'get variable object storage
      If Vptr Is Nothing Then Exit Function     'error
      TstData = Val(ExtractValue(Vptr))         'get value there to temp variable
    Else
      Call CheckForNumber(InstrPtr, 3, 255)
    End If
    On Error Resume Next
    RGB_Support = CInt(TstData)                 'return value
    Call CheckError
    On Error GoTo 0
    If ErrorFlag Then RGB_Support = -1          'indicate error
  End If
End Function

'*******************************************************************************
' Function Name     : Circle_Support
' Purpose           : Provide data support for Circle
'*******************************************************************************
Public Function Circle_Support(Value As Single, Errstr As String) As Boolean
  Dim Vptr As clsVarSto
  
  If GetInstruction(1) <> iComma Then
    Errstr = "Expected ','"
    Exit Function
  End If
  
  IncInstrPtr
  If GetInstruction(1) = iIND Then            'using a variable?
    Set Vptr = CheckVbl()                     'get variable object storage
    If Vptr Is Nothing Then Exit Function     'error
    TstData = Val(ExtractValue(Vptr))         'get value there to temp variable
  Else
    Call CheckForValue(InstrPtr)
  End If
  On Error Resume Next
  Value = CSng(TstData)                       'get value
  Call CheckError
  If ErrorFlag Then Exit Function
  On Error GoTo 0
  Circle_Support = True
End Function

'*******************************************************************************
' Function Name     : GetPlotValue
' Purpose           : Support geting a Plot Value, and getting Print data
'*******************************************************************************
Public Function GetPlotValue(Prm As String, ByVal HighRange As Long, _
                             Result As Long, Errstr As String) As Boolean
  Dim Vptr As clsVarSto
  
  'IncInstrPtr                                 'skip '(' or ','
  If GetInstruction(1) = iIND Then            'using a variable?
    Set Vptr = CheckVbl()                     'get variable object storage
    If Vptr Is Nothing Then Exit Function     'error
    TstData = Val(ExtractValue(Vptr))         'get value there to temp variable
  Else
    Call CheckForNumber(InstrPtr, 3, 0)
  End If
  If TstData < 0# Or TstData > CDbl(HighRange) Then
    Errstr = Prm & " plot value out of range (0-" & CStr(HighRange) & ")"
    Exit Function
  End If
  Result = CLng(TstData)                           'get value
  GetPlotValue = True
End Function

'*******************************************************************************
' Function Name     : GetPlotXY
' Purpose           : Get Plot's X and Y values. Within parens. Allow vars
'*******************************************************************************
Public Function GetPlotXY(X As Long, Y As Long, Errstr As String) As Boolean
  Dim Vptr As clsVarSto
  
  If GetInstruction(0) <> iLparen Then        'error if not "("
    Errstr = "Expected '('"
    Exit Function
  End If
  
  If Not GetPlotValue("X", PlotWidth, X, Errstr) Then Exit Function  'get X to TstData
  
  If GetInstruction(1) = iComma Then          'check for comma separator
    IncInstrPtr
  Else
    Errstr = "Expected ','"
    Exit Function
  End If
  
  If Not GetPlotValue("Y", PlotHeight, Y, Errstr) Then Exit Function    'get Y to TstData
  
  If GetInstruction(1) = iEparen Then         'check for ')'
    IncInstrPtr
  Else
    Errstr = "Expected ')'"
    Exit Function
  End If
  GetPlotXY = True                            'success
End Function

'*******************************************************************************
' Function Name     : GetPointClr
' Purpose           : Return the color at the specified point
'*******************************************************************************
Public Function GetPointClr(ByVal X As Long, Y As Long) As Long
  GetPointClr = GetPixel(frmVisualCalc.PicPlot.hDC, X + PlotXOfst, Y + PlotYOfst)
End Function

'*******************************************************************************
' Subroutine Name   : SetPoint
' Purpose           : Set a single point
'*******************************************************************************
Public Sub SetPoint(ByVal X As Long, Y As Long, Clr As Long)
  Call SetPixel(frmVisualCalc.PicPlot.hDC, X + PlotXOfst, Y + PlotYOfst, Clr) 'then do new color
  PlotX = X
  PlotY = Y
End Sub

'*******************************************************************************
' Subroutine Name   : MoveTo
' Purpose           : Set draw position
'*******************************************************************************
Public Sub MoveTo(ByVal X As Long, ByVal Y As Long)
  Call MoveToEx(frmVisualCalc.PicPlot.hDC, X + PlotXOfst, Y + PlotYOfst, 0&)
  PlotX = X
  PlotY = Y
End Sub

'*******************************************************************************
' Subroutine Name   : DrawLine
' Purpose           : Draw a line
'*******************************************************************************
Public Sub DrawLine(ByVal Xs As Long, ByVal Ys As Long, _
                    ByVal Xe As Long, ByVal Ye As Long, ByVal Clr As Long)
  MoveTo Xs, Ys
  frmVisualCalc.PicPlot.ForeColor = PlotColor
  Call LineTo(frmVisualCalc.PicPlot.hDC, Xe + PlotXOfst, Ye + PlotYOfst)
  PlotX = Xe
  PlotY = Ye
End Sub

'*******************************************************************************
' Subroutine Name   : PaintIt
' Purpose           : Perform flood fill
'*******************************************************************************
Public Sub PaintIt(ByVal X As Long, ByVal Y As Long, ByVal nClr As Long, Optional Brdr As Long = -1)
  Dim iBdr As Long
  
  If Brdr = -1 Then                                       'if user did not supply a border color...
    iBdr = nClr                                           'then use new color as border color
  Else
    iBdr = Brdr                                           'else user-specified
  End If
  
  frmVisualCalc.PicPlot.FillColor = nClr
  Call ExtFloodFill(frmVisualCalc.PicPlot.hDC, X + PlotXOfst, Y + PlotYOfst, iBdr, FLOODFILLBORDER)
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

