Attribute VB_Name = "ModHookclsCBOFullDrop"
Option Explicit
'~ModHookclsCBOFullDrop.bas;clsCBOFullDrop.cls;
'Module used exclusively by clsCBOFullDrop.cls
'-------------------------------------------------------------------------------
' API Stuff
'-------------------------------------------------------------------------------
Private Const GWL_WNDPROC As Long = -4&
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
' Property names used to stash info within window props. Multiple combo boxes can use these names without
' clashing because they are each further associated by unique hWnd values.
'
Private Const NewCBOWndProc As String = "NewCBOWndProc"
Private Const OldCBOWndProc As String = "OldCBOWndProc"

'*******************************************************************************
' Function Name     : MyCBOHook
' Purpose           : Invoke hooked procedure
'*******************************************************************************
Public Function MyCBOHook(ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
  Dim Obj As clsCBOFullDrop                         'local pointer to a clsCBOFullDrop object
  Dim lpObjPtr As Long                              'pointer to actual object procedure
  
  lpObjPtr = GetProp(hWnd, NewCBOWndProc)           'get prodedure to invoke
  CopyMemory Obj, lpObjPtr, 4                       'copy to Obj pointer
  MyCBOHook = Obj.MyCBOWndProc(hWnd, Msg, wp, lp)   'invoke actual WndProc for object
  CopyMemory Obj, Nothing, 4                        'clear local Obj hook
End Function

'*******************************************************************************
' Subroutine Name   : UnhookCBOhWnd
' Purpose           : Unhook a procedure
'*******************************************************************************
Public Sub UnhookCBOhWnd(hWnd As Long)
  Dim lpWndProc As Long
    
  lpWndProc = GetProp(hWnd, OldCBOWndProc)          'get procedure invoked before our hook inserted
  If CBool(lpWndProc) Then                          'if defined...
   Call SetWindowLong(hWnd, GWL_WNDPROC, lpWndProc) 'plug back into chain, replacing current link
  End If
  Call RemoveProp(hWnd, NewCBOWndProc)              'then remove properties...
  Call RemoveProp(hWnd, OldCBOWndProc)
End Sub

'*******************************************************************************
' Subroutine Name   : HookCBOhWnd
' Purpose           : Hook a procedure, provided Combobox handle an interface class
'*******************************************************************************
Public Sub HookCBOhWnd(hWnd As Long, WndProcObj As clsCBOFullDrop)
  If CBool(GetProp(hWnd, OldCBOWndProc)) Then                         'if already hooked
    Call UnhookCBOhWnd(hWnd)                                          'first unhook previous hook
  End If
  Call SetProp(hWnd, NewCBOWndProc, ObjPtr(WndProcObj))               'save proc to invoke
  Call SetProp(hWnd, OldCBOWndProc, GetWindowLong(hWnd, GWL_WNDPROC)) 'save current proc hooking before
  Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf MyCBOHook)          'insert our new hook
End Sub

'*******************************************************************************
' Function Name     : InvokeCBOWndProc
' Purpose           : Invoke old procedure
'*******************************************************************************
Public Function InvokeCBOWndProc(hWnd As Long, Msg As Long, wp As Long, lp As Long) As Long
   InvokeCBOWndProc = CallWindowProc(GetProp(hWnd, OldCBOWndProc), hWnd, Msg, wp, lp)
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

