VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search for Help Topic"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7065
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Key &Location"
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
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "When a keypad key is selected, show location on keypad"
      Top             =   120
      Width           =   1635
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   1260
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   660
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4515
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1755
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3096
      _Version        =   393217
      TextRTF         =   $"frmSearch.frx":0000
   End
   Begin VB.PictureBox PicBack 
      Height          =   615
      Left            =   360
      ScaleHeight     =   555
      ScaleWidth      =   735
      TabIndex        =   6
      Top             =   1080
      Width           =   795
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
      Left            =   6060
      TabIndex        =   5
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Display"
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
      Left            =   5100
      TabIndex        =   4
      Top             =   4800
      Width           =   855
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E1FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   540
      Width           =   6795
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Type a topic, the text to find, or select from the main or drop-down list the item you want to find help for."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   9
      Top             =   4740
      Width           =   4815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Topic:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   495
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private myShadow As clsShadow 'form shadow class
Private cboDropHandler As clsCBOFullDrop

Private colSrch As Collection 'collection of help selections
Private NoUpdt As Boolean     'updating flag
Private HoldCode As Long      'hold selection code

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Load the search lists
'*******************************************************************************
Private Sub Form_Load()
  Dim Pth As String, Txt As String, S As String, T As String
  Dim Idx As Long, Idy As Long
'
' disable search options in linked forms
'
  With frmVisualCalc
    .mnuHelpSrch.Enabled = False
    .mnuHelpFind.Enabled = False
  End With
  If frmHelpLoaded Then
    frmHelp.Toolbar1.Buttons("Search").Enabled = False
  End If
'
' init business
'
  Screen.MousePointer = vbHourglass
  DoEvents
  Me.Icon = frmVisualCalc.Icon                        'copy icon from main form
  Me.PicBack.Picture = frmVisualCalc.PicBack.Picture  'image to use for backgrpund
  InitTileFormBackground Me.PicBack                   'init tiling for background
  Set colSrch = New Collection
  NewQuery = 1000
  NewQueryOFST = 0
'
' load help file and gather topics
'
  Pth = AddSlash(App.Path) & "VPCHelp.rtf"  'path to help file
  With Me.RichTextBox1
    .LoadFile Pth, rtfRTF                       'load help file
    Txt = .Text
  End With
  
  Idx = -1                                      'init index to base-2
  NoUpdt = True
  Do
    Idx = InStr(Idx + 2, Txt, "@@")             'find a key entry '@@num@'
    Idy = InStr(Idx, Txt, vbCr)                 'find line terminator
    S = Trim$(Mid$(Txt, Idx, Idy - Idx))        'get line
    Idy = InStr(3, S, "@")                      'find trailing '@'
    If Left$(S, Idy) = "@@1000@" Then Exit Do   'terminator
    T = Cleanit(Trim$(Mid$(S, Idy + 1)))        'get just line title to T
    List1.AddItem T                             'add data
    List2.AddItem T & String$(100 - Len(T), 32) & Left$(S, Idy)
  Loop
  List1.ListIndex = -1                          'init pointers
  List1.TopIndex = 0
'
' get search list
'
  Idx = CLng(GetSetting(App.Title, "Settings", "SrchCnt", "0"))
  For Idy = 0 To Idx - 1
    S = GetSetting$(App.Title, "Settings", "Srch" & Format(Idy, "00"))
    colSrch.Add S, S
    Combo1.AddItem S
  Next Idy
'
' get last search
'
  S = GetSetting(App.Title, "Settings", "LastSrch", vbNullString)
  NoUpdt = False
  If CBool(Len(S)) Then
    With Me.Combo1
      For Idx = 0 To .ListCount - 1
        If S = .List(Idx) Then
          .ListIndex = Idx
          Exit For
        End If
      Next Idx
    End With
  End If
  Me.cmdOK.Enabled = CBool(Len(Trim$(Me.Combo1.Text)))  'enable OK if data in search text box
  If Me.cmdOK.Enabled Then
    HoldCode = ImpCmd(Trim$(Me.Combo1.Text))
    Me.cmdShow.Enabled = HoldCode <> 128
  End If
'
' set handler for combobox
' Disable the following 2 lines if you will be debugging it.
'
  Set cboDropHandler = New clsCBOFullDrop
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
'
' reset display cursor and indicate form loaded
'
  Screen.MousePointer = vbDefault
  frmSrchLoaded = True
End Sub

'*******************************************************************************
' Function Name     : Cleanit
' Purpose           : Strip brackets from help topic entries
'*******************************************************************************
Private Function Cleanit(Txt As String) As String
  Dim S As String
  Dim i As Integer
  
  S = Txt
  If InStr(1, S, "'") = 0 Then
    Do While InStr(1, S, "[")
      i = InStr(1, S, "[")
      S = Left$(S, i - 1) & Mid$(S, i + 1)
    Loop
    
    Do While InStr(1, S, "]")
      i = InStr(1, S, "]")
      S = Left$(S, i - 1) & Mid$(S, i + 1)
    Loop
  End If
  
  Cleanit = S
End Function

'*******************************************************************************
' Subroutine Name   : Form_Paint
' Purpose           : Refresh background
'*******************************************************************************
Private Sub Form_Paint()
  PaintTileFormBackground Me, Me.PicBack
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Save any changes to the Search window
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Dim Idx As Integer
'
' save search list (up to 25 previous entries)
'
  If CBool(Me.Combo1.ListCount) Then
    Idx = 0
    With colSrch
      Do While CBool(.Count)
        If Idx < 25 Then
          SaveSetting App.Title, "Settings", "Srch" & Format(Idx, "00"), .Item(.Count)
          Idx = Idx + 1
        End If
        .Remove .Count
      Loop
      SaveSetting App.Title, "Settings", "SrchCnt", CStr(Idx)
    End With
  End If
'
' clear resources
'
  frmSrchLoaded = False
  Set myShadow = Nothing
  Set cboDropHandler = Nothing
  Set colSrch = Nothing
'
' enable tags
'
  With frmVisualCalc
    .mnuHelpSrch.Enabled = True
    .mnuHelpFind.Enabled = True
  End With
  If frmHelpLoaded Then
    frmHelp.Toolbar1.Buttons("Search").Enabled = True
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Combo1_Change
' Purpose           : When the combo list text is changed
'*******************************************************************************
Private Sub Combo1_Change()
  Dim S As String, T As String
  Dim Idx As Long
  Dim HaveMatch As Boolean
  
  If NoUpdt Then Exit Sub                 'ignore if blocked
  S = Me.Combo1.Text                      'get text to process
  Me.cmdOK.Enabled = CBool(Len(S))        'enable OK key as required
  If Not Me.cmdOK.Enabled Then Exit Sub   'nothing to do if nothing there
  
  HaveMatch = False                       'init no match
  NoUpdt = True                           'block updates
  With Me.Combo1
    If CBool(.ListCount) Then             'if Combo box has items in its list
      Idx = FindMatch(Me.Combo1, S)       'find a match
      If Idx <> -1 Then                   'found one?
        .Text = .List(Idx)                'yes, apply partials
        .SelStart = Len(S)                'highlight data beyond user-typed characters
        .SelLength = Len(.Text) - Len(S)
        HaveMatch = True                  'indicate we have a match
      End If
    End If
    
    Idx = FindMatch(Me.List1, S)          'find a match in the main list
    If Idx <> -1 Then                     'found one?
      Me.List1.ListIndex = Idx            'yes, select it in the main list
      If Not HaveMatch Then               'did we find a match previously?
        .Text = Me.List1.List(Idx)        'no, so apply partial as above
        .SelStart = Len(S)
        .SelLength = Len(.Text) - Len(S)
      End If
    End If
  End With
  If Me.cmdOK.Enabled Then
    HoldCode = ImpCmd(Trim$(Me.Combo1.Text))
    Me.cmdShow.Enabled = HoldCode <> 128
  End If
  NoUpdt = False                          'turn off blocker
End Sub

'*******************************************************************************
' Subroutine Name   : Combo1_Click
' Purpose           : Selected something from the combo list
'*******************************************************************************
Private Sub Combo1_Click()
  Dim S As String
  Dim Idx As Long
  
  If NoUpdt Then Exit Sub                 'blocked processing
  PlayClick 'play resource click
  With Me.Combo1
    S = .Text                             'get user indexed selection text
    Idx = FindExactMatch(Me.List1, S)     'exact match in master list?
    If Idx <> -1 Then                     'found a match?
      NoUpdt = True                       'yes, block updates
      Me.List1.ListIndex = Idx            'select item in master list
      NoUpdt = False                      'stop blocking updates
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : Combo1_GotFocus
' Purpose           : Highlight combo text when it gets focus
'*******************************************************************************
Private Sub Combo1_GotFocus()
  With Me.Combo1
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cmdOK_Click
' Purpose           : Accept Selection
'*******************************************************************************
Private Sub cmdOK_Click()
  Dim S As String, Txt As String, T As String
  Dim Idx As Long, Idy As Long, Idz As Long
  
  PlayClick                                     'play resource click
  S = Me.Combo1.Text                            'get selection text
  If CBool(InStr(1, S, "@")) Then
    CenterMsgBoxOnForm Me, "Cannot include '@' symbol in search.", vbOKOnly Or vbExclamation, "Illegal Character"
    Exit Sub
  End If
  Idx = FindExactMatch(Me.List1, S)             'find an exact match in the master list
  NewQueryOFST = 0                              'init to top of topic
  If Idx = -1 Then                              'did not find one
    Txt = Me.RichTextBox1.Text                  'so we will search for a match in the main text
    Idz = InStr(1, Txt, S, vbTextCompare)       'found a match?
    If CBool(Idz) Then
      Idx = InStrRev(Txt, "@@", Idz)            'yes, find leading @@
      If CBool(Idx) Then                        'found one?
        Idy = InStr(Idx + 2, Txt, "@")          'find ending @
        T = Mid$(Txt, Idx + 2, Idy - Idx - 2)   'grab code there
        NewQuery = CInt(T)                      'save as new query data
        NewQueryOFST = Idz - Idy - 1            'set topic offset to item
        NewQueryLen = Len(S)
      End If
    End If
    If NewQuery = 1000 Then                     'error of NewQuery still 1000
      CenterMsgBoxOnForm Me, "Cannot find text:" & vbCrLf & "'" & S & "'", vbOKOnly Or vbExclamation, "Text Not Found"
      Exit Sub
    End If
  Else
    T = Me.List2.List(Idx)                      'found an exact match for selected
    Idx = InStrRev(T, "@@")                     'grab index
    T = Mid$(T, Idx + 2)
    NewQuery = CInt(Left$(T, Len(T) - 1))       'topic number
  End If
  
  On Error Resume Next
  colSrch.Add S, S                              'add to collection
  If Not CBool(Err.Number) Then
    NoUpdt = True                               'new item, so we will add it to the combo list
    With Me.Combo1
      .AddItem S
      .ListIndex = .NewIndex
    End With
    NoUpdt = False
  End If
  SaveSetting App.Title, "Settings", "LastSrch", Me.Combo1.Text
  Unload Me                                     'now return to invoker
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCancel_Click
' Purpose           : Close form
'*******************************************************************************
Private Sub cmdCancel_Click()
  PlayClick                             'play resource click
  Unload Me
  frmVisualCalc.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : List1_Click
' Purpose           : When an item in the list is clicked
'*******************************************************************************
Private Sub List1_Click()
  If NoUpdt Then Exit Sub               'ignore if blocked
  PlayClick                             'play resource click
  
  With Me.List1
    Me.Combo1.Text = .List(.ListIndex)  'get selection in master list
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : List1_DblClick
' Purpose           : Force Accept Selection
'*******************************************************************************
Private Sub List1_DblClick()
  Me.cmdOK.Value = True
End Sub

'*******************************************************************************
' Subroutine Name   : cmdShow_Click
' Purpose           : Show keypad key
'*******************************************************************************
Private Sub cmdShow_Click()
  Dim Set2nd As Boolean
  Dim i As Integer
  Dim S As String
  
  On Error Resume Next
  S = Me.Combo1.Text                            'get selection text
  colSrch.Add S, S                              'add to collection
  If Not CBool(Err.Number) Then
    NoUpdt = True                               'new item, so we will add it to the combo list
    With Me.Combo1
      .AddItem S
      .ListIndex = .NewIndex
    End With
    NoUpdt = False
  End If
  On Error GoTo 0
  SaveSetting App.Title, "Settings", "LastSrch", Me.Combo1.Text
  
  Me.Hide
  DoEvents
  With frmVisualCalc
    Select Case HoldCode
      Case Is < 10
        Select Case HoldCode  'set 0-9 keys to actual offsets on keypad
          Case 0
            i = 92
          Case 1
            i = 79
          Case 2
            i = 80
          Case 3
            i = 81
          Case 4
            i = 66
          Case 5
            i = 67
          Case 6
            i = 68
          Case 7
            i = 53
          Case 8
            i = 54
          Case 9
            i = 55
        End Select
        CtrPtrOnBtn .cmdKeyPad(i), .PicKeys.hWnd, frmVisualCalc
      Case Is < 128   'ascii keys
        CmdNotActive
      Case Is < 256   'primary keys
        i = HoldCode - 128
        .chk2nd.Value = vbUnchecked
        CtrPtrOnBtn .cmdKeyPad(i), .PicKeys.hWnd, frmVisualCalc
      Case Is > 900   'user keys
        i = HoldCode - 900
        CtrPtrOnBtn .cmdUsrA(i), .PicKeys.hWnd, frmVisualCalc
      Case Is < 512   '2nd keys
        i = HoldCode - 256
        .chk2nd.Value = vbChecked
        CtrPtrOnBtn .cmdKeyPad(i), .PicKeys.hWnd, frmVisualCalc
    End Select
  End With
  Unload Me
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************
