VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find All Instances of Text in Help"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Recent"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Erase entire recent selection list"
      Top             =   660
      Width           =   1275
   End
   Begin VB.PictureBox PicBack 
      Height          =   615
      Left            =   300
      ScaleHeight     =   555
      ScaleWidth      =   735
      TabIndex        =   5
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Find"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      ToolTipText     =   "Find matches"
      Top             =   660
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      ToolTipText     =   "Close Find dialog"
      Top             =   660
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1620
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   180
      Width           =   4275
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   795
      Left            =   3780
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1402
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmFind.frx":0000
   End
   Begin VB.Label lblDEL 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete single entry by selecting it, then hitting the  Delete key."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1380
      TabIndex        =   8
      Top             =   1260
      Width           =   2355
   End
   Begin VB.Label lblFinding 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Found:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1620
      TabIndex        =   2
      Top             =   720
      Width           =   600
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text To Search for:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1410
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private myShadow As clsShadow 'form shadow class
Private cboDropHandler As clscboFullDrop
Private colFind As Collection

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Load and prepare form
'*******************************************************************************
Private Sub Form_Load()
  Dim Idx As Integer, cnt As Integer
  Dim S As String
  
  frmVisualCalc.Enabled = False
  Me.Icon = frmVisualCalc.Icon                        'copy icon from main form
  Me.PicBack.Picture = frmVisualCalc.PicBack.Picture  'image to use for backgrpund
  InitTileFormBackground Me.PicBack                   'init tiling for background
  Me.lblFinding.Visible = False
  With Me.lblDEL
    .Top = Me.lblFinding.Top - 90
    .Left = Me.lblFinding.Left
    .Visible = True
  End With
  Me.Height = 1665
  Set colFind = New Collection
'
' get list of items for combobox
'
  cnt = CInt(GetSetting(App.Title, "Settings", "FindCnt", "0"))
  For Idx = 1 To cnt
    S = Trim$(GetSetting(App.Title, "Settings", "Find" & CStr(Idx), vbNullString))
    If CBool(Len(S)) Then
      On Error Resume Next
      colFind.Add S, S
      If Not CBool(Err.Number) Then Me.Combo1.AddItem S
      On Error GoTo 0
    End If
  Next Idx
  Me.cmdClear.Enabled = CBool(Me.Combo1.ListCount)
'
' set up the last-sought item
'
  S = Trim$(GetSetting(App.Title, "Settings", "LastFind", vbNullString))
  With Me.Combo1
    If CBool(Len(S)) Then
        .Text = vbNullString
        For Idx = 0 To .ListCount - 1
          If StrComp(S, .List(Idx), vbTextCompare) = 0 Then
            .ListIndex = Idx
            Me.cmdOK.Enabled = CBool(Len(.Text))
            Exit For
          End If
        Next Idx
        If Idx = .ListCount Then .Text = S
    Else
      .Text = S
    End If
  End With
'
' set handler for combobox
' Disable the following 2 lines if you will be debugging it.
'
  Set cboDropHandler = New clscboFullDrop
  cboDropHandler.hWnd = Combo1.hWnd
'
' apply a shadow to form
' Disable the following block of code if you will be debugging it.
'
  Set myShadow = New clsShadow
  With myShadow
    If .Shadow(Me) Then
      .Depth = 10
      .Transparency = 128
    Else
      Set myShadow = Nothing
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Paint
' Purpose           : Refresh background
'*******************************************************************************
Private Sub Form_Paint()
  PaintTileFormBackground Me, Me.PicBack
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Unload form, save find list
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Dim cnt As Integer
  Dim S As String
  
  SaveSetting App.Title, "Settings", "LastFind", Me.Combo1.Text
  With colFind
    cnt = .Count
    If cnt > 25 Then cnt = 25   'save only the last 25 items
    SaveSetting App.Title, "Settings", "FindCnt", CStr(cnt)
    Do While CBool(cnt)
      SaveSetting App.Title, "Settings", "Find" & CStr(cnt), .Item(.Count)
      cnt = cnt - 1
      .Remove .Count
    Loop
  End With
'
' clear resources
'
  Set colFind = Nothing
  Set myShadow = Nothing
  Set cboDropHandler = Nothing
  frmVisualCalc.Enabled = True
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCancel_Click
' Purpose           : Cancel Find operation
'*******************************************************************************
Private Sub cmdCancel_Click()
  Call PlayClick
  Unload Me
  frmVisualCalc.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : Combo1_GotFocus
' Purpose           : Highlight full line when we have focus
'*******************************************************************************
Private Sub Combo1_GotFocus()
  With Me.Combo1
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : Combo1_Change
' Purpose           : Enable Find button when text present
'*******************************************************************************
Private Sub Combo1_Change()
  Call PlayClick
  Me.cmdOK.Enabled = CBool(Len(Trim$(Me.Combo1.Text)))
End Sub

'echo above
Private Sub Combo1_Click()
  Call PlayClick
  Me.cmdOK.Enabled = CBool(Len(Trim$(Me.Combo1.Text)))
End Sub

'*******************************************************************************
' Subroutine Name   : Combo1_KeyDown
' Purpose           : Delete an entry
'*******************************************************************************
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim S As String
  Dim Idx As Integer
  
  If KeyCode = 46 Then                        'DEL?
    With Me.Combo1
      Idx = .ListIndex                        'yes, get selected index
      If Idx = -1 Then Exit Sub               'cannot do anything
      S = .List(Idx)                          'else get text of selection (removed from Combo box)
      .RemoveItem Idx                         'remove from combo list
      If Idx = .ListCount Then Idx = Idx - 1  'fix up index
      If Idx = -1 Then Exit Sub               'if nothing there
      .ListIndex = Idx                        'else set new selection
    End With
    With colFind
      For Idx = 1 To .Count                   'now find is collection
        If StrComp(.Item(Idx), S, vbTextCompare) = 0 Then
          .Remove Idx                         'found it, so remove it
          Exit For
        End If
      Next Idx
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : cmdOK_Click
' Purpose           : Search for all matches
'*******************************************************************************
Private Sub cmdOK_Click()
  Dim Txt As String, S As String, Pth As String
  Dim Idx As Long, cnt As Long
  
  Call PlayClick
  S = Me.Combo1.Text
  If CBool(InStr(1, S, "@")) Then
    CenterMsgBoxOnForm Me, "Cannot include '@' symbol in search.", vbOKOnly Or vbExclamation, "Illegal Character"
    Exit Sub
  End If

  Me.cmdOK.Enabled = False
  Me.cmdCancel.Enabled = False              'disable buttons
  Me.Combo1.Enabled = False
  Me.cmdClear.Enabled = False
  
  Screen.MousePointer = vbHourglass         'show that we are busy
  DoEvents
'
' load help file and gather topics
'
  Pth = AddSlash(App.Path) & "VPCHelp.rtf"  'path to help file
  With Me.RichTextBox1
    .LoadFile Pth, rtfRTF                   'load help file
    Txt = .Text
  End With
'
' init find list
'
  Me.lblDEL.Visible = False
  With Me.lblFinding
    .Caption = "Found: 0"
    .Visible = True
    .Refresh
  End With
'
' get data to find, init accumulator
'
  S = Me.Combo1.Text
  FindListLen = Len(S)
  Call ClearFindList
'
' add item to list
'
  On Error Resume Next
  colFind.Add S, UCase$(S)
  If Not CBool(Err.Number) Then Me.Combo1.AddItem S
  On Error GoTo 0
'
' search for all matches
'
  Idx = InStr(1, Txt, "@@")                   'skip to first topic
  Idx = InStr(Idx, Txt, S, vbTextCompare)     'find first match
  Do While CBool(Idx)
    colFindList.Add CStr(Idx)
    cnt = cnt + 1
    With Me.lblFinding
      .Caption = "Found: " & CStr(cnt)
      .Refresh
    End With
    Idx = InStr(Idx + FindListLen, Txt, S, vbTextCompare)
  Loop
'
' see if we found anything
'
  If colFindList.Count = 0 Then
    CenterMsgBoxOnForm Me, "Did not find any matches.", vbOKOnly Or vbInformation, "Nothing Found"
    Me.cmdOK.Enabled = True
    Me.cmdCancel.Enabled = True
    Me.Combo1.Enabled = True
    Me.cmdClear.Enabled = CBool(Me.Combo1.ListCount)
    Me.lblFinding.Visible = False
    Me.lblDEL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
'
' process matches
'
  FindListIdx = 1         'init to first match
  If frmHelpLoaded Then   'remove Previous option of Help file is already open
    With colHelpBack      'clear out any previous tpoic pages stored
      Do While .Count
        .Remove 1
      Loop
    End With
    frmHelp.Toolbar1.Buttons("Back").Enabled = False  'disable in case set
  End If
  Call ShowFindListItem   'find items
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdClear_Click
' Purpose           : Clear recent list
'*******************************************************************************
Private Sub cmdClear_Click()
  If CenterMsgBoxOnForm(Me, "Verify erase recent list.", vbYesNo Or vbQuestion, "Verify") = vbNo Then Exit Sub
  SaveSetting App.Title, "Settings", "FindCnt", "0" 'mark nothing to load
  Me.Combo1.Clear                                   'clear user-visible list
  Me.cmdClear.Enabled = False                       'disable button
  With colFind                                      'clear hidden list
    Do While CBool(.Count)
      .Remove 1
    Loop
  End With
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

