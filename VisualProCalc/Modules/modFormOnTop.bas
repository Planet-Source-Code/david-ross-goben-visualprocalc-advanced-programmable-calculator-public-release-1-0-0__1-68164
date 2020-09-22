Attribute VB_Name = "modFormOnTop"
Option Explicit
'~modFormOnTop.bas;
'Maintain form location
'*************************************************
' modFormOnTop: Maintain form location
'
' The following functions are provided:
'
' StayOnTop(): Place the specified form on top of display stack.
' NotOnTop():  No longer place specified form on top of the display stack.
'              Use this when the form will STILL be displayed,
'              but not on top, otherwise this is not needed, such as
'              during a form.Hide.
'*************************************************

'*************************************************
' API Interface
'*************************************************
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

'*************************************************
' Place the specified form on top of display stack
'*************************************************
Public Sub StayOnTop(frm As Form)
  Call SetWindowPos(frm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW)
End Sub

'*************************************************
' No longer place specified form on top of the display stack
'
' Use this when the form will STILL be displayed,
' but not on top, otherwise this is not needed, such as
' during a form.Hide.
'*************************************************
Public Sub NotOnTop(frm As Form)
  Call SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW)
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

