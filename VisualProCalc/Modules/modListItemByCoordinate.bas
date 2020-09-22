Attribute VB_Name = "modListItemByCoordinate"
Option Explicit
'~modListItemByCoordinate.bas;
'Get a ListIndex in a ListBox by coodinates. Useful on Right-Clicks
'*********************************************************************************
' modListItemByCoordinate - The ListItemByCoordinate() function returns the ListIndex
'                           of an item in a ListBox. Although this is normally possible
'                           with a left-click, where you can obtain it by the obvious
'                           ListBox.ListIndex property, the item is not selected on a
'                           right-pick or mouse-move, where you might want to display
'                           information on the item in a ToolTip.
'EXAMPLE:
'  Private Sub ListBox1_MouseDown(Button As Insert_Project_Name, Shift As Integer, _
'                                 X As Single, Y As Single)
'    If Button = 2 Then 'Right Click
'      ListBox1.ToolTopText = "Right-Clicked on " & _
'                  ListBox1.List(ListItemByCoordinate(ListBox1, X, Y))
'    End If
'  End Sub
'
'NOTE: This function will also work with Drive Boxes, Dir Boxes, and File Boxes.
'*********************************************************************************
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_GETITEMHEIGHT = &H1A1

Public Function ListItemByCoordinate(ListBox As Object, X As Single, Y As Single) As Long
 Dim Idx As Long
 
 Idx = Y \ (SendMessage(ListBox.hWnd, LB_GETITEMHEIGHT, 0&, 0&) * Screen.TwipsPerPixelY) + ListBox.TopIndex
 If Idx >= ListBox.ListCount Then Idx = -1
 ListItemByCoordinate = Idx
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

