Attribute VB_Name = "modFindMatch"
Option Explicit
'~modFindMatch.bas;
'Find text match in a listbox, combobox, or textbox control
'********************************************************************************
' modFindMatch:
' Find text match in a listbox, combobox, or textbox control.
'
' The functions return -1 if no match was found. Note that search are
' not case sensestive.
'
' Note that the difference between FindMatch() and FindExactMatch() is
' that FindMatch search from the start of each list item and checks for
' an entry beginning with the specified string, and FindExactMatch()
' checkes for a match with ALL text in a line in a listbox or combobox.
'
' Search for text in a textbox are identical between the two functions,
' and will return the starting offset, from 1, of the match (the text in
' a textbox is considered one long string), and will search for a
' non-case-sensetive match of the full test string.
'
' This module provides the following fuctions:
'
' FindMatch():      Find general text match in listbox, combobox, or textbox, starting at beginning
' FindExactMatch(): Find exact text match in listbox, combobox, or textbox, starting at beginning
'
'EXAMPLE:
'  Dim Result As Long
'  Result = FindMatch(ListBox1, "Jeff Goldblum")
'  Debug.Print Result
'
' NOTE: List entries in listboxes and comboboxes start with line 0 (zero).
'
'********************************************************************************

'********************************************************************************
' API Interface
'********************************************************************************
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Const LB_FINDSTRING = &H18F
Private Const CB_FINDSTRING = &H14C
Private Const LB_FINDSTRINGEXACT = &H1A2
Private Const CB_FINDSTRINGEXACT = &H158

'********************************************************************************
' FindMatch(): Find general text match in listbox, combobox, or textbox, starting at beginning
' Using the optional StartIndex will allow you to start from the last match index.  However,
' notice that when no additional matches are found, the last-found match is returned.  Hence,
' if you use this option, keep track of the last match, so that when a repeat index or
' a LOWER value is returned, this flags that no more matches were found.
'********************************************************************************
Public Function FindMatch(ctl As Control, Text As String, Optional StartIndex As Long = -1) As Long
  Dim Result As Long
  Dim Index As Long
  
  Index = StartIndex
  Result = -1                     'init result to failure
  If TypeOf ctl Is ListBox Then
    Result = SendMessageByString(ctl.hwnd, LB_FINDSTRING, Index, Text)
  End If
  If TypeOf ctl Is ComboBox Then
    Result = SendMessageByString(ctl.hwnd, CB_FINDSTRING, Index, Text)
  End If
  If TypeOf ctl Is TextBox Then
    Result = InStr(LCase$(ctl.Text), LCase$(Text))
  End If
  If Result < 0 Then             'make sure errors match value
    FindMatch = -1
  Else
    FindMatch = Result
  End If
  
End Function

'********************************************************************************
' FindExactMatch(): Find exact text match in listbox, combobox, or textbox, starting at beginning
' Using the optional StartIndex will allow you to start from the last match index.  However,
' notice that when no additional matches are found, the last-found match is returned.  Hence,
' if you use this option, keep track of the last match, so that when a repeat index is
' returned, this flags that no more matches were found.
'********************************************************************************
Public Function FindExactMatch(ctl As Control, Text As String, Optional StartIndex As Long = -1) As Long
  Dim Result As Long
  Dim Index As Long
  
  Index = StartIndex
  If TypeOf ctl Is ListBox Then
    Result = SendMessageByString(ctl.hwnd, LB_FINDSTRINGEXACT, Index, Text)
  ElseIf TypeOf ctl Is ComboBox Then
    Result = SendMessageByString(ctl.hwnd, CB_FINDSTRINGEXACT, Index, Text)
  ElseIf TypeOf ctl Is TextBox Then
    Result = InStr(LCase$(Text), LCase$(Text))
  Else
    Result = -1                  'init result to failure
  End If
  If Result < 0 Then             'make sure errors match value
    FindExactMatch = -1
  Else
    FindExactMatch = Result
  End If
  
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

