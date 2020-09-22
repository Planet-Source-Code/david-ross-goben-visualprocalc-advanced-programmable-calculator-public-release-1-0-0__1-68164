VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   5760
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7245
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975.653
   ScaleMode       =   0  'User
   ScaleWidth      =   6803.431
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "frmAbout.frx":0000
      Top             =   3720
      Width           =   7095
   End
   Begin VB.PictureBox PicBack 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   735
      TabIndex        =   6
      Top             =   2340
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
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
      Height          =   345
      Left            =   5625
      TabIndex        =   0
      Top             =   2385
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "Send email..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5640
      TabIndex        =   1
      ToolTipText     =   "Send an email to David at DavidGoben@yahoo.com"
      Top             =   2835
      Width           =   1245
   End
   Begin VB.Label lblHistory 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Release History"
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
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   60
      Picture         =   "frmAbout.frx":0006
      Top             =   120
      Width           =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   1239.548
      X2              =   6606.23
      Y1              =   1563.343
      Y2              =   1563.343
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":1BF0
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1110
      Index           =   0
      Left            =   1260
      TabIndex        =   2
      Top             =   1140
      Width           =   5745
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   1295.892
      X2              =   6606.23
      Y1              =   1573.696
      Y2              =   1573.696
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
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
      Left            =   6360
      TabIndex        =   5
      Top             =   180
      Width           =   660
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":1D4C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   885
      Left            =   1320
      TabIndex        =   3
      Top             =   2400
      Width           =   4050
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   0
      Left            =   1260
      TabIndex        =   4
      Top             =   120
      Width           =   900
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   1239.548
      X2              =   6592.145
      Y1              =   496.957
      Y2              =   496.957
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   1239.548
      X2              =   6592.145
      Y1              =   496.957
      Y2              =   496.957
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Personal Programmable Calculator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   1
      Left            =   1260
      TabIndex        =   7
      Top             =   765
      Width           =   5490
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private myShadow As clsShadow 'form shadow class

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Initialize About Box
'*******************************************************************************
Private Sub Form_Load()
  Dim Pth As String
  Dim ts As TextStream
  
  Me.Icon = frmVisualCalc.Icon                        'copy icon from main form
  Me.Caption = "About " & App.Title                   'give us a title
  lblVersion.Caption = "Version " & GetAppVersion()   'build version
  If CBool(App.Major) Then
    Me.lblTitle(0).Caption = App.Title
  Else
    Me.lblTitle(0).Caption = App.Title & " (Beta)"
  End If
  Me.lblDescription(0).Caption = AddTitle(Me.lblDescription(0).Caption)
  Pth = AddSlash(App.Path) & "ReleaseHistory.txt"
  If CBool(Len(Dir(Pth))) Then
    Set ts = Fso.OpenTextFile(Pth, ForReading, False)
    Me.Text1.Text = ts.ReadAll
    ts.Close
    Me.Text1.SelStart = 0
  Else
    Me.Height = 4000
  End If
  
  Me.PicBack.Picture = frmVisualCalc.PicBack.Picture  'image to use for backgrpund
  InitTileFormBackground Me.PicBack                   'init tiling for background
'
' apply a shadow to form
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
' Purpose           : Unloading. Remove allocated resources
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Set myShadow = Nothing
End Sub

'*******************************************************************************
' Subroutine Name   : cmdOK_Click
' Purpose           : Close form
'*******************************************************************************
Private Sub cmdOK_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdSysInfo_Click
' Purpose           : Send author email
'*******************************************************************************
Private Sub cmdSysInfo_Click()
  SendEMail Me.hWnd, "DavidGoben@yahoo.com", App.Title
End Sub

