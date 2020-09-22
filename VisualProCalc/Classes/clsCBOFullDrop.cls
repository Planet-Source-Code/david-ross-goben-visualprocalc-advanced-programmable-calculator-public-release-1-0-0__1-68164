VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCBOFullDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'~clsCBOFullDrop.cls;ModHookclsCBOFullDrop.bas;
'allows ComboBox dropdowns to display its ENTIRE list
'*********************************************************************
' clsCBOFullDrop: This class module allows ComboBox dropdowns to display its ENTIRE
'                 list, not just a partial. Declare one of these object for each
'                 ComboBox you with to affect.
#If False Then
'USAGE:
' In form header:
  Private cboDropHandler As clsCBOFullDrop
' In Form_Load() event:
  Set cboDropHandler = New clsCBOFullDrop
  cboDropHandler.hWnd = myComboBox.hWnd
' In Form_Unload() event:
  Set cboDropHandler = Nothing
#End If
'
' NOTE: This Class uses the module ModHookclsCBOFullDrop.bas
'*********************************************************************

'-------------------------------------------------------------------------------
' API Stuff
'-------------------------------------------------------------------------------
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Const declarations
Private Const LB_GETCOUNT = &H18B
Private Const LB_GETITEMHEIGHT = &H1A1
Private Const SM_CYBORDER = 6
Private Const WM_LBUTTONUP = &H202

' The WM_CTLCOLORLISTBOX message is sent to the parent window of a list box before
' the system draws it. By reacting to this message, we can define the actual
' screen footprint we want to use, instead of the limited default.
Private Const WM_CTLCOLORLISTBOX = &H134

'-------------------------------------------------------------------------------
' Private Member variable to store Hwnd value. Save for unhooking
'-------------------------------------------------------------------------------
Private m_hWnd As Long        'local copy og hWnd for combobox
Private m_oldTop As Long      'keep track of list location within WndProc
Private m_oldHeight As Long

'*******************************************************************************
' Subroutine Name   : UnHkPrv
' Purpose           : Unhook any previous hooking definition
'*******************************************************************************
Private Sub UnHkPrv()
  If CBool(m_hWnd) Then         'if association exists...
    Call UnhookCBOhWnd(m_hWnd)  'unhook handle
    m_hWnd = 0                  'remove tag to prevent redundancy
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Class_Terminate
' Purpose           : When terminating, release anything set
'*******************************************************************************
Private Sub Class_Terminate()
  Call UnHkPrv                    'remove any previous hook
End Sub

'*******************************************************************************
' Let Name          : hWnd
' Purpose           : Assign a new combobox hWnd. If one already set, unhook it first
'*******************************************************************************
Public Property Let hWnd(ByVal NewhWnd As Long)
  Call UnHkPrv                    'remove any previous hook
  If IsWindow(NewhWnd) Then       'hook only if hWnd is for a Window object
    m_hWnd = NewhWnd              'save hwnd for later unhook
    Call HookCBOhWnd(m_hWnd, Me)  'pass hook the combobox handle and clsCBOFullDrop object
  End If
End Property

'*******************************************************************************
' Function Name     : MyCBOWndProc
' Purpose           : WndProc function
'*******************************************************************************
Friend Function MyCBOWndProc(hWnd As Long, Msg As Long, wp As Long, lp As Long) As Long
  Dim iHeight As Long   'single item height in pixels
  Dim nCount As Long    '# items in list to display
  Dim nHeight As Long   'target height
  Dim nWidth As Long    'combo box width
  Dim ScHeight As Long  'screen height in pixels
  Dim rLst As RECT      'full size dimensions
  Dim rCbo As RECT      'combo box position

  Select Case Msg
    Case WM_LBUTTONUP   'when left button up (click), reset dims
      m_oldHeight = 0
      m_oldTop = 0
    
    'when preparing to paint, size the droplist to show as much as possible
    Case WM_CTLCOLORLISTBOX
      iHeight = SendMessage(lp, LB_GETITEMHEIGHT, 0, ByVal 0&)  'get single-item height
      nCount = SendMessage(lp, LB_GETCOUNT, 0, ByVal 0&)        'get # if items in list
      Call GetWindowRect(lp, rLst)                              'get initial combo box drop list size
      nHeight = (nCount * iHeight) + (GetSystemMetrics(SM_CYBORDER) * 2)
      Call GetWindowRect(hWnd, rCbo)                            'get cbo posn
      ScHeight = (Screen.Height \ Screen.TwipsPerPixelY)        'get screen height in pixels
      nWidth = rLst.Right - rLst.Left                           'get width of box
      If rCbo.Bottom + nHeight > ScHeight Then                  'list below screen?
        If rCbo.Top - nHeight < 0 Then                          'can move above box?
          nHeight = ((ScHeight - rCbo.Bottom + iHeight - 1) \ iHeight) * iHeight 'no, so shorten
        Else
          rLst.Top = rCbo.Top - nHeight                         'flip to above box if we can
        End If
        If m_oldHeight <> nHeight Or m_oldTop <> rLst.Top Then  'update only if different
          Call MoveWindow(lp, rLst.Left, rLst.Top, nWidth, nHeight, 1&) 'size the window
        End If
      Else      'nothing existing, no screen bleed-over, so just do it
        Call MoveWindow(lp, rLst.Left, rLst.Top, nWidth, nHeight, 1&) 'size the window
      End If
      m_oldHeight = nHeight                                     'save dims
      m_oldTop = rLst.Top
  End Select
  MyCBOWndProc = InvokeCBOWndProc(hWnd, Msg, wp, lp)            'invoke regular proc
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

