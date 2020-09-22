Attribute VB_Name = "modActivatePrv"
Option Explicit
'~modActivatePrv.bas;
'Reactivate an application
'*******************************************************************************
' modActivatePrv - The ActivatePrv() function will reactivate an application.
'                  This is useful if you wish to run only a single instance of an
'                  application that can run multiple instances if you simply shell
'                  out to it. The function first looks to see if a window with a
'                  specific caption, such as "Calculator", already exists as a
'                  running application. If it is minimized (iconic), then it is
'                  restored and given focus. Regardless, it is brought to the top
'                  of the display stack and given focus.
'*******************************************************************************

'------------------------------
'API Stuff
'------------------------------
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Private AppTtl As String
Private AppHwnd As Long

Private Const SW_SHOW = 5
Private Const SW_RESTORE = 9

'*******************************************************************************
' Subroutine Name   : ActivateApp
' Purpose           : Start/reactivate an application
'*******************************************************************************
Public Function ActivatePrv(frm As Form) As Boolean
  Dim OldTitle As String
  Dim Hnd As Long
  
  OldTitle = frm.Caption                    'save form caption
  frm.Caption = "unwanted instance"         'garble it
  Hnd = FindActiveWindow(OldTitle)          'App already loaded?
  frm.Caption = OldTitle                    'reset form name
  If CBool(Hnd) Then                        'app alreadedy loaded...
    Call SetForegroundWindow(Hnd)           'set it as the foreground window
    If CBool(IsIconic(Hnd)) Then            'is it iconic?
      ShowWindow Hnd, SW_RESTORE            'yes, so restore it and make it active
    Else                                    'move to top, though do not keep topmost
      ShowWindow Hnd, SW_SHOW               'not iconic, so move to display top
    End If
  End If
  ActivatePrv = CBool(Hnd)                  'successful if Hnd <> 0
End Function

'*******************************************************************************
' Function Name     : FindActiveWindow
' Purpose           : Return the handle if the window caption is found
'*******************************************************************************
Public Function FindActiveWindow(AppTitle As String) As Long
  AppTtl = AppTitle                                       'set text string to search for
  AppHwnd = 0                                             'init handle to zippo
  Call EnumWindows(AddressOf Callback_EnumWindows, 0&)    'find an active window
  FindActiveWindow = AppHwnd                              'return found handle (0 if not found)
End Function

'*******************************************************************************
' Function Name     : Callback_EnumWindows
' Purpose           : Callback function to support finding an active window
'*******************************************************************************
Private Function Callback_EnumWindows(ByVal hWnd As Long, ByVal lpData As Long) As Long
  Dim WinCaption As String
  Dim Clen As Long
  
  If Not CBool(GetParent(hWnd)) Then                      'process only if parent is desktop
    WinCaption = Space$(256)                              'pick up caption
    Clen = GetWindowText(hWnd, WinCaption, 256&)          'length to Slen, data to S
'
' because FindWindow() does not seem to operate correctly on Win98, and windows can return with
' blank captions, we must check to see if the caption length is non-zero, the window is visible
' (Win98 seems to keep tossed windows in a cache in case they are called again)
'
    If CBool(Clen) Then                           'contains text?
      If CBool(IsWindowVisible(hWnd)) Then        'is form visible?
        If Left$(WinCaption, Clen) = AppTtl Then  'captions match?
          AppHwnd = hWnd                          'grab found handle
          Callback_EnumWindows = 0                'Cancel enumeration
          Exit Function
        End If
      End If
    End If
  End If
  Callback_EnumWindows = 1      ' Continue enumeration
End Function

