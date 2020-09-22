Attribute VB_Name = "modCtrPtrOnBtn"
Option Explicit
'~modCtrPtrOnBtn.bas;
'Center the mouse on a specified control.
'*********************************************************
' modCtrPtrOnBtn: The CtrPtrOnBtn() subroutine centers the
'                 mouse cursor on the specified control.
'EXAMPLE:
'  CtrPtrOnBtn cmdOK, MyForm ' move to OK button
'*********************************************************

'*********************************************************
' API calls, types, and constants
'*********************************************************
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal Dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

Private Const MOUSEEVENTF_MOVE = &H1        '  mouse move
Private Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move

'*********************************************************
' MoveMouse(): track the mouse cursor to the specified control.
'              Auto-click the control if desired. Optionally set a delay.
'*********************************************************
Public Sub CtrPtrOnBtn(myBtn As Control, ParenthWnd As Long, myFrm As Form)
  Dim pt As POINTAPI
  Dim dl As Long
  Dim DestX As Long, DestY As Long, curx As Long, cury As Long
  Dim distx As Long, disty As Long
  Dim screenx As Long, screeny As Long
  Dim finished As Boolean
  Dim ptsperx As Long, ptspery As Long
'
' set form as top-most
'
  myFrm.ZOrder 0
  DoEvents
'
' Get screen coordinates in pixels at center of button
'
  
  With myBtn
    pt.X = (.Left + .Width \ 2) \ Screen.TwipsPerPixelX
    pt.Y = (.Top + .Height \ 2) \ Screen.TwipsPerPixelY
  End With
  dl& = ClientToScreen(ParenthWnd, pt)      'get coordinate based on button
  screenx = GetSystemMetrics(SM_CXSCREEN)   'get screen size x and y
  screeny = GetSystemMetrics(SM_CYSCREEN)
  
  DestX = pt.X * &HFFFF& / screenx
  DestY = pt.Y * &HFFFF& / screeny
'
' About how many mouse points per pixel
'
  ptsperx = &HFFFF& / screenx * 4&
  ptspery = &HFFFF& / screeny * 4&
'
' Now move mouse
'
  Do
    dl = GetCursorPos(pt)                   'get current mouse position
    curx = pt.X * &HFFFF& / screenx         'compute x location
    cury = pt.Y * &HFFFF& / screeny         'compute y location
    distx = DestX - curx                    'get distance to move x
    disty = DestY - cury                    'get distance to move y
'
' Move closer
'
    curx = curx + Sgn(distx) * ptsperx * 2  'move in that direction
    cury = cury + Sgn(disty) * ptspery * 2
    If Abs(distx) < 2 * ptsperx Then curx = DestX 'do not exceed limits
    If Abs(disty) < 2 * ptspery Then cury = DestY
    If curx& = DestX And cury = DestY Then finished = True
    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, curx, cury, 0, 0
    Sleep 25                                'wait a little for effect
  Loop While Not finished
End Sub


