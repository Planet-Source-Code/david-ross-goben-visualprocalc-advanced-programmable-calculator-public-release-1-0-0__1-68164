Attribute VB_Name = "modChooseColor"
Option Explicit
'*******************************************************************************
' API stuff
'*******************************************************************************
Public Type RGBQUAD
  rgbBlue     As Byte
  rgbGreen    As Byte
  rgbRed      As Byte
  rgbReserved As Byte
End Type

Private Type CHOOSECOLORSTRUCT
  lStructSize     As Long
  hwndOwner       As Long
  hInstance       As Long
  rgbResult       As Long
  lpCustColors    As Long
  flags           As Long
  lCustData       As Long
  lpfnHook        As Long
  lpTemplateName  As String
End Type

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLORSTRUCT) As Long

Private Const CC_RGBINIT = &H1
Private Const CC_FULLOPEN = &H2

'*******************************************************************************
' Function Name     : GetColor
' Purpose           : Get color from ChooseColor API Dialog Box
'*******************************************************************************
Public Function GetColor(frm As Form, Clr As Long) As Boolean
  Dim cc As CHOOSECOLORSTRUCT
  Dim dwCustomColors(15) As Long
  Dim cnt As Long
'
' init some gray shades
'
  For cnt = 240 To 15 Step -15
    dwCustomColors((cnt \ 15) - 1) = RGB(cnt, cnt, cnt)
  Next cnt
'
' invoke dialog
'
  With cc
    .hwndOwner = frm.hWnd
    .rgbResult = PlotColor
    .flags = CC_RGBINIT Or CC_FULLOPEN
    .lpCustColors = VarPtr(dwCustomColors(0))
    .lStructSize = Len(cc)
    If CBool(ChooseColor(cc)) Then
      Clr = .rgbResult                'color changed
      GetColor = True
    End If
  End With
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

