VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCoDisplay 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Co-Display Formatted Source in Learn Mode"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2700
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCoDisplay.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCoDisplay.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCoDisplay.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCoDisplay.frx":1A8E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   1376
      ButtonWidth     =   1561
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Syncronize"
            Key             =   "Refresh"
            Object.ToolTipText     =   "Syncromize Co-Display with Learn Mode code"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Auto Sync"
            Key             =   "AutoUpdate"
            Object.ToolTipText     =   "Push to auto-syncromize Co-Display with Learn Mode code"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy All"
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy list to the clipboard"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "Stop"
            Object.ToolTipText     =   "Close Co-Display WIndow"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4350
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   5715
            Text            =   "Select Line to select in actual Source"
            TextSave        =   "Select Line to select in actual Source"
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
   Begin VB.ListBox lstSrc 
      BackColor       =   &H00F8E4D8&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      ItemData        =   "frmCoDisplay.frx":1EE0
      Left            =   0
      List            =   "frmCoDisplay.frx":1EE2
      TabIndex        =   0
      Top             =   780
      Width           =   2295
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2700
      Top             =   1980
   End
   Begin VB.Label lblSizing 
      AutoSize        =   -1  'True
      Caption         =   "lblSizing"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   420
      TabIndex        =   2
      Top             =   2700
      Width           =   1080
   End
End
Attribute VB_Name = "frmCoDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_LBUTTON = &H1
Private Resize As Boolean       'flag True when resizing being processed

Private myShadow As clsShadow   'form shadow class

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Match the main form
'*******************************************************************************
Private Sub Form_Load()
  Dim Lf As Long, Wd As Long, Tp As Long, HT As Long
  Dim Idx As Integer, SelLn As Integer, i As Integer
  Dim S As String
  
  Me.lblSizing.Top = -1440
  Me.Icon = frmVisualCalc.Icon
  Call GetScreenWorkArea(Lf, Wd, Tp, HT)
  Me.Width = CLng(GetSetting(App.Title, "Settings", "CDWidth", CStr(Wd \ 4)))
  Me.Height = CLng(GetSetting(App.Title, "Settings", "CDHeight", CStr(HT)))
  Me.Left = CLng(GetSetting(App.Title, "Settings", "CDLeft", CStr(CInt(Wd * 0.75))))
  Me.Top = CLng(GetSetting(App.Title, "Settings", "CDTop", CStr(Tp)))
  Me.Toolbar1.Buttons("AutoUpdate").Value = CInt(GetSetting(App.Title, "Settings", "CDUpdate", "1"))
  Me.Toolbar1.Buttons("Refresh").Enabled = Not CBool(Me.Toolbar1.Buttons("AutoUpdate").Value)
'
' load the listbox with the formatted output
'
  With Me.lstSrc
    .Clear
    SelLn = 0                                       'hold select line
    i = 0                                           'index valid lines
    For Idx = 0 To FmtCnt - 1
      If FmtMap(Idx) <> -1 Then                     'if this line has not been merged out...
        .AddItem FmtLst(Idx)                        'add to listbox
        If FmtMap(Idx) <= InstrPtr Then SelLn = i   'if within range, set as indexed line
        i = i + 1                                   'bump acceptable line counter
      End If
    Next Idx
    On Error Resume Next
    .ListIndex = SelLn                              'set select line
    Idx = SelLn - .Height \ frmVisualCalc.lblChkSize.Height \ 2 'ensure we are not at top
    If Idx < 0 Then Idx = 0                         'we may have to...
    .TopIndex = Idx                                 'set top line
  End With
'
' hook sizing subclass
' Disable the following line if you will be debugging it.
'
  Call HookWin(Me.hWnd, m_CDhWnd)                   'hook sizing monitor
'
' apply a shadow to form
' Disable the following block of code if you will be debugging it.
'
  Set myShadow = New clsShadow                      'set up shadow
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
  frmCDLoaded = True                                'indicate form loaded
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Resize the form
'*******************************************************************************
Private Sub Form_Resize()
  If Resize Then Exit Sub
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
  If Me.Width < 5000 Or Me.Height < 5000 Then 'if too small, smooth reset with a timer
    With Me.tmrWait
      .Enabled = False                        'turn timer off, if on
      DoEvents                                'let screen catch up
      .Enabled = True                         'restart and reset timer
    End With
    Exit Sub
  End If
  With Me.lstSrc
    .Left = 0
    .Top = Me.Toolbar1.Height
    .Width = Me.ScaleWidth                     'resize RTF form to scale
    .Height = Me.ScaleHeight - Me.StatusBar1.Height - Me.Toolbar1.Height
  End With
  On Error Resume Next
  frmVisualCalc.lstDisplay.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Save position/size settings
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  SaveSetting App.Title, "Settings", "CDLeft", CStr(Me.Left)
  SaveSetting App.Title, "Settings", "CDTop", CStr(Me.Top)
  SaveSetting App.Title, "Settings", "CDWidth", CStr(Me.Width)
  SaveSetting App.Title, "Settings", "CDHeight", CStr(Me.Height)
  Set myShadow = Nothing
  If CBool(m_CDhWnd) Then Call UnhookWin(Me.hWnd, m_CDhWnd)
  frmCDLoaded = False
  frmVisualCalc.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : lstSrc_Click
' Purpose           : Select a line
'*******************************************************************************
Private Sub lstSrc_Click()
  Dim i As Integer, Idx As Integer
  
  If Not frmCDLoaded Then Exit Sub    'nothing to do if form not loaded
  If ResetPnt Then Exit Sub           'ignore if resetting a location
  PlayClick                           'play resource click
  i = Me.lstSrc.ListIndex + 1         'get current point
  If i = 0 Then Exit Sub              'not pointing to anywhere
  For Idx = 0 To FmtCnt - 1           'find valid lines
    If FmtMap(Idx) <> -1 Then
      i = i - 1                       'count off a line
      If i = 0 Then Exit For          'found enought valid lines
    End If
  Next Idx
  If Idx > FmtCnt Then Exit Sub       'ignore if we passed through whole list
  InstrPtr = FmtMap(Idx)              'set instruction location to start of line
  SelectOnly InstrPtr                 'select that line in the program listing
  With frmVisualCalc.lstDisplay       'now try to center the selected line
    i = .ListIndex - (DisplayHeight \ 2)
    If i < 0 Then i = 0
    .TopIndex = i
  End With
  frmVisualCalc.lstDisplay.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : lstSrc_MouseMove
' Purpose           : Display data that is too long as a tooltip
'*******************************************************************************
Private Sub lstSrc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim S As String
  
  With Me.lblSizing
    S = Me.lstSrc.List(ListItemByCoordinate(Me.lstSrc, X, Y))
    .Caption = S
    If .Width < Me.lstSrc.Width - 240 Then S = vbNullString
  End With
  With Me.lstSrc
    If .ToolTipText <> S Then .ToolTipText = S
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : lstSrc_Scroll
' Purpose           : Scroll list
'*******************************************************************************
Private Sub lstSrc_Scroll()
  frmVisualCalc.lstDisplay.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : tmrWait_Timer
' Purpose           : Check if window size gets too small
'*******************************************************************************
Private Sub tmrWait_Timer()
  If Not CBool(m_CDhWnd) Then
    If GetKeyState(VK_LBUTTON) < 0 Then Exit Sub
  End If
  Me.tmrWait.Enabled = False                      'turn off timer
  Resize = True
  If Me.Width < 5000 Then Me.Width = 5000         'resize to minimum dims
  If Me.Height < 5000 Then Me.Height = 5000
  Resize = False
  Call Form_Resize
End Sub

'*******************************************************************************
' Subroutine Name   : RepointIndex
' Purpose           : Repoint the display index to match the Program Step
'*******************************************************************************
Public Sub RepointIndex()
  Dim Idx As Integer, i As Integer
  
  i = 0
  For Idx = 0 To FmtCnt - 1
    If FmtMap(Idx) <> -1 Then
      If InstrPtr < FmtMap(Idx) Then
        i = i - 1
        If i = -1 Then i = 0
        ResetPnt = True
        Me.lstSrc.ListIndex = i
        i = i - Me.lstSrc.Height \ frmVisualCalc.lblChkSize.Height \ 2
        If i < 0 Then i = 0
        Me.lstSrc.TopIndex = i
        ResetPnt = False
        Exit For
      End If
      i = i + 1
    End If
  Next Idx
End Sub

'*******************************************************************************
' Subroutine Name   : Toolbar1_ButtonClick
' Purpose           : Handle click on buttons in the toolbar
'*******************************************************************************
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim S As String
  Dim Idx As Long
  
  PlayClick                         'play resource click
  Select Case Button.Key
    Case "Refresh"                  'simply refresh the display
      SetUpCoDisplay
    
    Case "AutoUpdate"               'toggle auto-update
      SaveSetting App.Title, "Settings", "CDUpdate", CStr(Button.Value)
      Me.Toolbar1.Buttons("Refresh").Enabled = Not CBool(Me.Toolbar1.Buttons("AutoUpdate").Value)
      If CBool(Button.Value) Then   'and update now if turned on
        SetUpCoDisplay
      End If
      
    Case "Copy"                     'copy listing to the clipboard
      S = vbNullString              'accumulator
      With Me.lstSrc
        For Idx = 0 To .ListCount - 1
          S = S & .List(Idx) & vbCrLf 'build listing as a string of text
        Next Idx
      End With
      With Clipboard
        .Clear                      'erase any other data
        .SetText S, vbCFText        'set text as data in clipboard
      End With
    
    Case "Stop"                     'unload the form if the Exit (Stop) button selected
      Unload Me
  End Select
  frmVisualCalc.lstDisplay.SetFocus
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

