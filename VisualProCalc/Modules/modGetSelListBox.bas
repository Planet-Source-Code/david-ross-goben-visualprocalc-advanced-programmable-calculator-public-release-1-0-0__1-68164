Attribute VB_Name = "modGetSelListBox"
Option Explicit
'~modGetSelListBox.bas;
'return a variant array containing indexes to all items selected. Also deselect/select all
'*******************************************************************************
' modGetSelListBox:  The GetSelListBox() function return a Long array that contains
'                    the indexes into the listbox of all items selected.
' DeSelAllListBox(): Deselect/select all items in the ListBox.
'EXAMPLE:
'Private Sub Command1_Click()
'  Dim Mylist() As Long, X As Long
'
'  Mylist = GetSelListBox(List1)              'check for selections in listbox
'  If Not IsNull(Mylist) Then                 'if anything selected
'    For X = LBound(Mylist) To UBound(Mylist) 'safe ranging...
'      Debug.Print List1.List(Mylist(X))      'display selected item
'      List1.List(Mylist(X)).Selected = False 'reset selected item
'    Next X                                   'do all
'  End If
'End Sub
'*******************************************************************************

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_GETSELITEMS As Long = &H191

'*******************************************************************************
' GetSelListBox(): Return a Long array that contains the indexes into the listbox
'                  if any items selected.
'*******************************************************************************
Public Function GetSelListBox(LstBx As ListBox) As Long()
Dim ItemIndexes() As Long, iNumItems As Long

iNumItems = LstBx.SelCount
If CBool(iNumItems) Then
  ReDim ItemIndexes(iNumItems - 1)
  SendMessage LstBx.hwnd, LB_GETSELITEMS, iNumItems, ItemIndexes(0)
  GetSelListBox = ItemIndexes
End If
End Function

'*******************************************************************************
' DeSelAllListBox(): Unselect all selected items in the ListBox
'*******************************************************************************
Public Sub DeSelAllListBox(LstBx As ListBox, Optional SelectAll As Boolean = False)
  Dim Idx As Long, cnt As Long, Lst() As Long
  
  With LstBx
    If CBool(.SelCount) Then
      Lst = GetSelListBox(LstBx)          'get list of selected items
      For Idx = 0 To UBound(Lst)
        .Selected(CInt(Lst(Idx))) = False 'reset items in list
      Next Idx
    End If
  End With
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

