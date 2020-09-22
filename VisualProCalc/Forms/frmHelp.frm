VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   BackColor       =   &H00E1FFFF&
   Caption         =   "Topic Help"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   7545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   7545
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4320
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.Timer tmrWordpad 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   4680
      Top             =   3420
   End
   Begin VB.Timer tmrOnTop 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4680
      Top             =   4020
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4140
      Top             =   4020
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   5445
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   12488
            Text            =   "Press X button to close the Help display. Double click word to search for reference"
            TextSave        =   "Press X button to close the Help display. Double click word to search for reference"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   3420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":2024
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":28FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":31D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":362A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":3F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":47DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   1376
      ButtonWidth     =   1376
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Previous"
            Key             =   "Back"
            Object.ToolTipText     =   "Go back to previous help topic"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "WordPad"
            Key             =   "WordPad"
            Object.ToolTipText     =   "Launch Wordpad with help contents"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search..."
            Key             =   "Search"
            Object.ToolTipText     =   "Search for text in help file (Shift+F1)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find..."
            Key             =   "Find"
            Object.ToolTipText     =   "FInall all text matches"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "Exit"
            Object.ToolTipText     =   "Close the Topic Help window"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Prev"
            Key             =   "Prev"
            Object.ToolTipText     =   "Previous match in Find List"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "List"
            Key             =   "List"
            Object.ToolTipText     =   "Select a Find Index from a list"
            ImageIndex      =   7
            Style           =   5
            Object.Width           =   400
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Next"
            Key             =   "Next"
            Object.ToolTipText     =   "Next match in Find List"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbInfo 
      Height          =   4275
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   7541
      _Version        =   393217
      BackColor       =   14811135
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmHelp.frx":50B8
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblWidth"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4500
      TabIndex        =   4
      Top             =   2280
      Width           =   795
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000015&
      X1              =   60
      X2              =   6360
      Y1              =   5340
      Y2              =   5340
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      X1              =   0
      X2              =   6480
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Menu mnuPopUpHelp 
      Caption         =   "mnuPopUpHelp"
      Begin VB.Menu mnuPopUpSelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuPopUpCopy 
         Caption         =   "&Copy selection"
      End
      Begin VB.Menu mnuPopUpSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupPrevious 
         Caption         =   "Previous Topic"
      End
      Begin VB.Menu mnuPopUpWordPad 
         Caption         =   "Open topic in WordPad"
      End
      Begin VB.Menu mnuPopUpSearch 
         Caption         =   "&Search for selection"
      End
      Begin VB.Menu mnuPopupFind 
         Caption         =   "&Find all text matches"
      End
      Begin VB.Menu mnuPopUpPrev 
         Caption         =   "Prevous match in List"
      End
      Begin VB.Menu mnuPopUpList 
         Caption         =   "Select match from a list"
      End
      Begin VB.Menu mnuPopUpNext 
         Caption         =   "Next match in list"
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'****************************************************
' Constants, API calls
'****************************************************
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" _
       (ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long
Private Const EM_LINESCROLL As Long = &HB6
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_LBUTTON = &H1
'
' local stuff
'
Private Resize As Boolean     'flag True when resizing being processed
Private myShadow As clsShadow 'form shadow class
Private DblClkRtb As Boolean  'flag used to record DblClick in rtbInfo

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Match the main form
'*******************************************************************************
Private Sub Form_Load()
  Dim Lf As Long, Wd As Long, Tp As Long, HT As Long
  
  Me.Icon = frmVisualCalc.Icon
  Me.cmdCancel.Left = -1440
  Me.mnuPopUpHelp.Visible = False
  Call GetScreenWorkArea(Lf, Wd, Tp, HT)
  Me.Width = CLng(GetSetting(App.Title, "Settings", "HlpWidth", CStr(Wd \ 2)))
  Me.Height = CLng(GetSetting(App.Title, "Settings", "HlpHeight", CStr(HT)))
  Me.Left = CLng(GetSetting(App.Title, "Settings", "HlpLeft", CStr(Wd \ 2)))
  Me.Top = CLng(GetSetting(App.Title, "Settings", "HlpTop", CStr(Tp)))
  Me.WindowState = CLng(GetSetting(App.Title, "Settings", "HlpWinState", CStr(vbNormal)))
  Me.Toolbar1.Buttons("WordPad").Enabled = CBool(Len(WordPadPath))
  Me.Toolbar1.Buttons("Back").Enabled = CBool(colHelpBack.Count > 1)
  
  With colFindList
    Me.Toolbar1.Buttons("Prev").Enabled = .Count > 2 And FindListIdx > 1
    Me.Toolbar1.Buttons("Next").Enabled = .Count > 2 And FindListIdx < .Count
  End With
'
' hook sizing subclass
' Disable the following line if you will be debugging it.
'
  Call HookWin(Me.hWnd, m_HlphWnd)
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
' indicate form loaded
'
  frmHelpLoaded = True
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Save position/size settings
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Me.rtbInfo.Text = vbNullString
  If Me.WindowState = vbNormal Then
    SaveSetting App.Title, "Settings", "HlpLeft", CStr(Me.Left)
    SaveSetting App.Title, "Settings", "HlpTop", CStr(Me.Top)
    SaveSetting App.Title, "Settings", "HlpWidth", CStr(Me.Width)
    SaveSetting App.Title, "Settings", "HlpHeight", CStr(Me.Height)
  End If
  SaveSetting App.Title, "Settings", "HlpWinState", CStr(Me.WindowState)
  Set myShadow = Nothing
  If CBool(m_HlphWnd) Then Call UnhookWin(Me.hWnd, m_HlphWnd)
  frmHelpLoaded = False
  frmVisualCalc.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Resize the form
'*******************************************************************************
Private Sub Form_Resize()
  If Resize Then Exit Sub
  Select Case Me.WindowState
    Case vbMinimized
      Exit Sub
    Case vbNormal
      If Not CBool(m_HlphWnd) Then
        If GetKeyState(VK_LBUTTON) < 0 Then           'if left mouse button down
          With Me.tmrWait                             'let timer handle fix
            .Enabled = False                          'disable timer
            DoEvents                                  'let screen catch up
            .Enabled = True                           're-enable timer (also resets it)
          End With
          Exit Sub
        End If
      End If
      If Me.Width < 6000 Or Me.Height < 6000 Then 'if too small, smooth reset with a timer
        With Me.tmrWait
          .Enabled = False                        'turn timer off, if on
          DoEvents                                'let screen catch up
          .Enabled = True                         'restart and reset timer
        End With
        Exit Sub
      End If
  End Select
  With Line1
    .Y1 = Me.Toolbar1.Height
    .Y2 = .Y1
    .X1 = 0
    .X2 = Me.ScaleWidth
  End With
  
  With Line2
    .Y1 = Me.ScaleHeight - Me.StatusBar1.Height - 15
    .Y2 = .Y1
    .X1 = 0
    .X2 = Me.ScaleWidth
  End With
  
  With Me.rtbInfo
    .Left = 120
    .Top = Me.Toolbar1.Height + .Left
    .Width = Me.ScaleWidth - .Left      'resize RTF form to scale
    .Height = Me.ScaleHeight - Me.StatusBar1.Height - Me.Toolbar1.Height - .Left * 2
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopUpList_Click
' Purpose           : Select from list of matches
'*******************************************************************************
Private Sub mnuPopUpList_Click()
  Dim SS As String
  Dim Idx As Long
  
  SS = Trim$(InputBox("Enter desired index (1-" & CStr(colFindList.Count) & "):", "Enter Find Index", CStr(FindListIdx)))
  If CBool(Len(SS)) Then
    If IsNumeric(SS) Then
      On Error Resume Next
      Idx = CLng(SS)
      If CBool(Err.Number) Then Exit Sub
      On Error GoTo 0
      If Idx < 1 Or Idx > colFindList.Count Then Exit Sub
      FindListIdx = Idx
      Call ShowFindListItem
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopUpNext_Click
' Purpose           : Go to next list match
'*******************************************************************************
Private Sub mnuPopUpNext_Click()
  FindListIdx = FindListIdx + 1
  Call ShowFindListItem
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopUpPrev_Click
' Purpose           : Go to previous list match
'*******************************************************************************
Private Sub mnuPopUpPrev_Click()
  FindListIdx = FindListIdx - 1
  Call ShowFindListItem
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopupPrevious_Click
' Purpose           : Go back to previous topic
'*******************************************************************************
Private Sub mnuPopupPrevious_Click()
  Dim Idx As Long, Idy As Long, Idz As Long
  Dim SS As String
  Dim wHandle As Long
    
  With colHelpBack
    If .Count < 2 Then
      Me.Toolbar1.Buttons("Back").Enabled = False
      CmdNotActive
      Exit Sub                            'noting to go back to
    End If
    .Remove .Count                        'remove current
    SS = .Item(.Count)
    Idx = InStr(1, SS, ";")
    If CBool(Idx) Then
      Idz = CLng(Mid$(SS, Idx + 1))       'get previous selstart
      Idx = CLng(Left$(SS, Idx - 1))      'get previous item
    Else
      Idz = 0                             'set to start of file
      Idx = CLng(SS)                      'get previous item
    End If
  End With
  Screen.MousePointer = vbHourglass
  DoEvents
  LockControlRepaint Me.rtbInfo
  With Me.rtbInfo
    .LoadFile AddSlash(App.Path) & "VPCHelp.rtf", rtfRTF
    Idy = InStr(Idx + 1, .Text, "@@") - 4 'find next block
    If CBool(Idy) Then                    'found it?
      .SelStart = Idy - 1                 'yes, so strip next data off
      .SelLength = Len(.Text) - Idy
      .SelText = vbNullString
      .SelStart = 0                       'strip loading text off
      .SelLength = Idx
      .SelText = vbNullString
      If CBool(Idz) Then .SelStart = Len(.Text)
      .SelStart = Idz                     'set cursor to top of form
    End If
    Me.Toolbar1.Buttons("Back").Enabled = CBool(colHelpBack.Count > 1)
    UnlockControlRepaint Me.rtbInfo
    .Refresh                              'refresh screen
  End With
  Screen.MousePointer = vbDefault
  DoEvents
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopUpWordPad_Click
' Purpose           : Open topic in WordPad
'*******************************************************************************
Private Sub mnuPopUpWordPad_Click()
  Dim SS As String
  
  With Me.tmrWordpad
    .Enabled = False
    On Error Resume Next
    SS = AddSlash(App.Path) & "WP.tmp"
    Me.rtbInfo.SaveFile SS, rtfRTF
    DoEvents
    Shell WordPadPath & " """ & SS & """", vbNormalFocus
    .Enabled = True
    Do While .Enabled
      DoEvents
    Loop
    Kill SS
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : rtbInfo_DblClick
' Purpose           : Itext double-clicked, so flag it, and let selection catch up
'*******************************************************************************
Private Sub rtbInfo_DblClick()
  DblClkRtb = True
End Sub

'*******************************************************************************
' Subroutine Name   : rtbInfo_SelChange
' Purpose           : If selection changed and dbl-click forced it, do it only then
'*******************************************************************************
Private Sub rtbInfo_SelChange()
  If DblClkRtb Then
    mnuPopupSearch_Click  'perform search if Dbl-Click on RtbInfo
  End If
  DblClkRtb = False
End Sub

'*******************************************************************************
' Subroutine Name   : rtbInfo_MouseDown
' Purpose           : Bring up Help submenu
'*******************************************************************************
Private Sub rtbInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    Me.mnuPopUpCopy.Enabled = CBool(Me.rtbInfo.SelLength)
    Me.mnuPopUpSearch.Enabled = Me.mnuPopUpCopy.Enabled
    Me.mnuPopupFind.Enabled = Me.mnuPopUpCopy.Enabled
    
    Me.mnuPopupPrevious.Enabled = Me.Toolbar1.Buttons("Back").Enabled
    Me.mnuPopUpPrev.Enabled = Me.Toolbar1.Buttons("Prev").Enabled
    Me.mnuPopUpList.Enabled = Me.Toolbar1.Buttons("List").Enabled
    Me.mnuPopUpNext.Enabled = Me.Toolbar1.Buttons("Next").Enabled
    
    PopupMenu Me.mnuPopUpHelp, vbPopupMenuRightButton
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopupCopy_Click
' Purpose           : Copy selection in Help window
'*******************************************************************************
Private Sub mnuPopupCopy_Click()
  Screen.MousePointer = vbHourglass
  DoEvents
  Clipboard.Clear
  With Me.rtbInfo
    Clipboard.SetText .SelText, vbCFText
    Clipboard.SetText .SelRTF, vbCFRTF
  End With
  Screen.MousePointer = vbDefault
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopupSelectAll_Click
' Purpose           : Select full contents of Help window
'*******************************************************************************
Private Sub mnuPopUpSelectAll_Click()
  With Me.rtbInfo
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopupFind_Click
' Purpose           : Find all text matches
'*******************************************************************************
Private Sub mnuPopupFind_Click()
  Load frmFind
  With frmFind
    .Combo1.Text = Me.rtbInfo.SelText
    .Show vbModeless, Me
    .cmdOK.Value = True
  End With
End Sub

Private Sub rtbInfo_Click()
  If Not Selecting Then PlayClick 'play resource click
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopupSearch
' Purpose           : User chose text to search for
'*******************************************************************************
Private Sub mnuPopupSearch_Click()
  Dim S As String, C As String, Pth As String, Txt As String, S2 As String, SS As String
  Dim Idx As Long, Idy As Long, Idz As Long
  Dim FndMatch As Boolean
  
  If Not HaveVHelp Then Exit Sub                    'if help file not found, just exit
  If SkipChg Then Exit Sub                          'ignore if skipping updates
  
  Screen.MousePointer = vbHourglass
  DoEvents
  Selecting = True
  FndMatch = False
  Do
    With Me.rtbInfo
      SS = ";" & CStr(.SelStart)                    'save select point
      Txt = .Text                                   'grab text user is looking at
      S = Trim$(.SelText)                           'grab selected text, strip surrounding spaces
      Select Case Len(S)
        Case 0                                      'nothing grabbed
          FndMatch = True
          Exit Do
        Case Is > 30                                'too long
          FndMatch = True
          Exit Do
        Case 1                                      'if single character, broaden scope
          S = vbNullString
          Idx = InStrRev(Txt, "[", .SelStart + 1)
          If CBool(Idx) Then
            Idy = InStr(.SelStart + 1, Txt, "]")
            If CBool(Idy) Then
              S = Mid$(Txt, Idx, Idy - Idx + 1)
              If Len(S) > 20 Then S = vbNullString
            End If
          End If
      End Select
    
      Pth = AddSlash(App.Path) & "VPCHelp.rtf"  'path to help file
      If Left$(S, 1) <> "[" Then
        S2 = vbNullString
        Idx = InStrRev(Txt, "[", .SelStart + 1)
        If CBool(Idx) Then
          Idy = InStr(.SelStart + 1, Txt, "]")
          If CBool(Idy) Then
            S2 = Mid$(Txt, Idx, Idy - Idx + 1)
            If Len(S2) > 30 Then S2 = vbNullString
          End If
        End If
        If CBool(Len(S2)) Then S = S2
      End If
    End With
    
    If Left$(S, 1) = "[" Then                       'if seeking [cmd], prepend "@"
      S = "@" & S
      If Right$(S, 1) <> "]" Then S = S & "]"       'ensure it is closed
      S2 = "@'" & Mid$(S, 3, Len(S) - 3) & "'"
      Select Case S2
        Case "@'0'", "@'1'", "@'2'", "@'3'", "@'4'", "@'5'", "@'6'", "@'7'", "@'8'", "@'9'"
          S = "@'0'"
        Case "@')'"
          S = "@'('"
        Case "@']'"
          S = "@'['"
        Case "@'}'"
          S = "@'{'"
        Case "@'If'", "@'Else'", "@'ElseIf'"
          S = "@'If'"
        Case "@'('", "@'['", "@'{'"
          S = S2
      End Select
      If S = "@[']" Then
        S = "@[Rem]"
      End If
    Else
      If IsNumeric(S) Then                          'treat numeric as a command #
        S = "@@" & S & "@"
        C = vbNullString
      Else
        C = S
        With Me.rtbInfo
          Idx = InStrRev(Txt, "[", .SelStart + 1)
          If CBool(Idx) Then
            Idy = InStr(.SelStart + 1, Txt, "]")
            If CBool(Idy) Then
              S = Mid$(Txt, Idx, Idy - Idx + 1)
              If Len(S) > 30 Then S = vbNullString
            End If
          End If
        End With
        If S = vbNullString Then
          S = C
        Else
          S = Mid$(S, 2, Len(S) - 2)
        End If
        Idx = ImpCmd(S)                             'else try to see if S was a Command token
        If Idx = 128 Then                           'not?
          S = "@[" & S & "]"                        'hm, try finding it anyway
        Else
          Select Case S
            Case "If", "Else", "ElseIf"
              S = "@'If'"
            Case Else
              S = "@@" & CStr(Idx) & "@"            'else use found command #
          End Select
        End If
      End If
    End If
    
    SkipChg = True
    Call ClearFindList
    LockControlRepaint Me.rtbInfo
    
    With Me.rtbInfo
      On Error Resume Next
      .LoadFile Pth, rtfRTF                         'load help file
      If CBool(Err.Number) Then Exit Do
      On Error GoTo 0
      Txt = .Text                                   'grab text
      Idx = InStr(34, Txt, S)                       'find first instance of data
      If CBool(Idx) Then
        If Left$(S, 2) = "@@" Then
          Idx = InStr(Idx, Txt, "@[")
        ElseIf Left$(S, 1) = "@" Then
          Idx = InStr(Idx, Txt, Left$(S, 2))
        Else
          Idx = InStr(Idx, Txt, "@[")
        End If
      ElseIf CBool(Len(C)) Then                   'if we did NOT find anything, try C's code
        Idx = InStr(34, Txt, C)
        If CBool(Idx) Then Idx = InStrRev(Txt, "@@", Idx)
      End If
      If Not CBool(Idx) Then Exit Do
      With colHelpBack
        If CBool(.Count) Then                 'if something in buffer
          S = .Item(.Count)
          Idz = InStr(1, S, ";")
          If CBool(Idz) Then S = Left$(S, Idz - 1)
          If CLng(S) <> Idx Then              'if not same as previous
            .Remove .Count                    'remove old
            .Add S & SS                       'add new
            .Add CStr(Idx)                    'then add index
          End If
        Else
          .Add CStr(Idx)                      'add anyway if nothing in buffer
        End If
        Me.Toolbar1.Buttons("Back").Enabled = CBool(.Count)
      End With
      Idy = InStr(Idx + 1, .Text, "@@") - 4 'find next block
      If CBool(Idy) Then                    'found it?
        FndMatch = True
        .SelStart = Idy - 1                 'yes, so strip next data off
        .SelLength = Len(.Text) - Idy
        .SelText = vbNullString
        .SelStart = 0                       'strip loading text off
        .SelLength = Idx
        .SelText = vbNullString
        .SelStart = 0                       'set cursor to top of form
      End If
    End With
    Exit Do
  Loop
  
  SkipChg = False
  If FndMatch Then
    UnlockControlRepaint Me.rtbInfo
    DoEvents
    Me.rtbInfo.Refresh                        'refresh screen
    Selecting = False
    Screen.MousePointer = vbDefault
    DoEvents
  Else
    Query -15
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : tmrOnTop_Timer
' Purpose           : Set form no longer on top
'*******************************************************************************
Private Sub tmrOnTop_Timer()
  Me.tmrOnTop.Enabled = False
  NotOnTop Me
End Sub

'*******************************************************************************
' Subroutine Name   : tmrWait_Timer
' Purpose           : Check if window size gets too small
'*******************************************************************************
Private Sub tmrWait_Timer()
  If Not CBool(m_HlphWnd) Then
    If GetKeyState(VK_LBUTTON) < 0 Then Exit Sub
  End If
  Me.tmrWait.Enabled = False                      'turn off timer
  If Me.WindowState = vbMinimized Then Exit Sub   'do nothing if minimized
  Resize = True
  If Me.Width < 6000 Then Me.Width = 6000         'resize to minimum dims
  If Me.Height < 6000 Then Me.Height = 6000
  Resize = False
  Call Form_Resize
End Sub

Private Sub tmrWordpad_Timer()
  Me.tmrWordpad.Enabled = False
End Sub

'*******************************************************************************
' Subroutine Name   : Toolbar1_ButtonClick
' Purpose           : Toolbar buttons clicked
'*******************************************************************************
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  
  PlayClick                                 'play resource click
  Select Case Button.Key
    Case "Exit"                             'if stop, then simply unload
      Unload Me
      Exit Sub
  
    Case "Back"                             'if Back, then try to back up index
      Call mnuPopupPrevious_Click
  
    Case "WordPad"                            'if wordpad
      Call mnuPopUpWordPad_Click
    
    Case "Search"
      frmSearch.Show vbModal, Me
      If NewQuery <> 1000 Then
        Call Query(9997)
      End If
      
    Case "Find"
      frmFind.Show vbModal, Me
    
    Case "Prev"
      FindListIdx = FindListIdx - 1
      Call ShowFindListItem
    
    Case "List"
      Call mnuPopUpList_Click
    
    Case "Next"
      FindListIdx = FindListIdx + 1
      Call ShowFindListItem
    
  End Select
  
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCancel_Click
' Purpose           : Close form via ESC key
'*******************************************************************************
Private Sub cmdCancel_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : Form_KeyDown
' Purpose           : Check for special keys
'*******************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Debug.Print KeyCode
  If KeyCode = 112 And Shift = vbShiftMask Then 'Shift+F1
    frmSearch.Show vbModal, Me
    If NewQuery <> 1000 Then
      Call Query(9997)
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Toolbar1_ButtonMenuClick
' Purpose           : Handle Button Menu selections
'*******************************************************************************
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  FindListIdx = ButtonMenu.Index
  Call ShowFindListItem
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************
