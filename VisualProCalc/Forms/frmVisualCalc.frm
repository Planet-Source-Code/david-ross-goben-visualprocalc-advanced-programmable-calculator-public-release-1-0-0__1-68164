VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmVisualCalc 
   Caption         =   "VisualProCalc"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   14775
   ClipControls    =   0   'False
   ForeColor       =   &H80000016&
   Icon            =   "frmVisualCalc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmVisualCalc.frx":1CFA
   ScaleHeight     =   6570
   ScaleWidth      =   14775
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4800
      Top             =   5820
   End
   Begin VB.PictureBox PicScroll 
      Height          =   5895
      Left            =   60
      ScaleHeight     =   5835
      ScaleWidth      =   465
      TabIndex        =   135
      Top             =   60
      Width           =   525
      Begin VB.CommandButton cmdBackspace 
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   21.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         Picture         =   "frmVisualCalc.frx":1E4C
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   2715
         Width           =   435
      End
      Begin VB.CommandButton cmdBtm 
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   15.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         Picture         =   "frmVisualCalc.frx":2156
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   4285
         Width           =   435
      End
      Begin VB.CommandButton cmdTop 
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   15.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         Picture         =   "frmVisualCalc.frx":2460
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   1145
         Width           =   435
      End
      Begin VB.CommandButton cmdPgDn 
         Caption         =   "Pg Dn"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   3500
         Width           =   435
      End
      Begin VB.CommandButton cmdPgUp 
         Caption         =   "Pg Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   1930
         Width           =   435
      End
      Begin VB.CommandButton cmdDn 
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   15.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         Picture         =   "frmVisualCalc.frx":276A
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   5070
         Width           =   435
      End
      Begin VB.CommandButton cmdUp 
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   15.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         Picture         =   "frmVisualCalc.frx":2A74
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   360
         Width           =   435
      End
      Begin VB.Image imgLocked 
         Height          =   240
         Index           =   1
         Left            =   120
         Picture         =   "frmVisualCalc.frx":2D7E
         Top             =   60
         Width           =   240
      End
      Begin VB.Image imgLocked 
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "frmVisualCalc.frx":2EC8
         Top             =   60
         Width           =   240
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000015&
         FillColor       =   &H80000013&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   30
         Top             =   15
         Width           =   375
      End
   End
   Begin VB.TextBox txtFocus 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   130
      TabStop         =   0   'False
      Text            =   "txtFocus"
      Top             =   1860
      Width           =   735
   End
   Begin VB.Timer tmrPause 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   960
      Top             =   1380
   End
   Begin VB.PictureBox PicBack 
      Height          =   1470
      Left            =   2700
      Picture         =   "frmVisualCalc.frx":3012
      ScaleHeight     =   1410
      ScaleWidth      =   1470
      TabIndex        =   142
      Top             =   2340
      Width           =   1530
   End
   Begin VB.PictureBox PicKeys 
      Height          =   5895
      Left            =   5355
      ScaleHeight     =   5835
      ScaleWidth      =   9315
      TabIndex        =   2
      Top             =   60
      Width           =   9375
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey Z"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   26
         Left            =   5360
         Style           =   1  'Graphical
         TabIndex        =   172
         Top             =   1080
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey Y"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   25
         Left            =   4020
         Style           =   1  'Graphical
         TabIndex        =   171
         Top             =   1080
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   24
         Left            =   2680
         Style           =   1  'Graphical
         TabIndex        =   170
         Top             =   1080
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey W"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   23
         Left            =   1340
         Style           =   1  'Graphical
         TabIndex        =   169
         Top             =   1080
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   22
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   168
         Top             =   1080
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey U"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   21
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   167
         Top             =   720
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey T"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   20
         Left            =   6700
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   720
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   19
         Left            =   5360
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   720
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   18
         Left            =   4020
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   720
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey Q"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   17
         Left            =   2680
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   720
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   16
         Left            =   1340
         Style           =   1  'Graphical
         TabIndex        =   162
         Top             =   720
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey O"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   15
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   720
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey N"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   14
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   360
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey M"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   13
         Left            =   6700
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   360
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey L"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   12
         Left            =   5360
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   360
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey K"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   11
         Left            =   4020
         Style           =   1  'Graphical
         TabIndex        =   157
         Top             =   360
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey J"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   10
         Left            =   2680
         Style           =   1  'Graphical
         TabIndex        =   156
         Top             =   360
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey I"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   9
         Left            =   1340
         Style           =   1  'Graphical
         TabIndex        =   155
         Top             =   360
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey H"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   8
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   360
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey G"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   7
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   0
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey F"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   6
         Left            =   6700
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   0
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   5
         Left            =   5360
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   0
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey D"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   4
         Left            =   4020
         Style           =   1  'Graphical
         TabIndex        =   150
         Top             =   0
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   3
         Left            =   2680
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   0
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   2
         Left            =   1340
         Style           =   1  'Graphical
         TabIndex        =   148
         Top             =   0
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Ukey A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   1
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   0
         Width           =   1275
      End
      Begin VB.CheckBox cmdUsrA 
         BackColor       =   &H00606060&
         Caption         =   "Space"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   0
         Left            =   6705
         Style           =   1  'Graphical
         TabIndex        =   146
         Top             =   1080
         Width           =   1275
      End
      Begin VB.CheckBox chkShift 
         BackColor       =   &H00606060&
         Caption         =   "Shift"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   8340
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdBtnHelp 
         BackColor       =   &H8000000D&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   124
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox chk2nd 
         BackColor       =   &H00FFC000&
         Caption         =   "2nd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00DF967A&
         Caption         =   "LRN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   780
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H80000005&
         Caption         =   "Pgm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1500
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "Load"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2220
         TabIndex        =   120
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2940
         TabIndex        =   119
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H0000C0FF&
         Caption         =   "CE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H004040FF&
         Caption         =   "CLR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "OP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   5100
         TabIndex        =   116
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "SST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   5820
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "INS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   6540
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Cut"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   7260
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   7980
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "PtoR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   8700
         TabIndex        =   111
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "STO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   60
         TabIndex        =   110
         Top             =   2100
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "RCL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   780
         TabIndex        =   109
         Top             =   2100
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "EXC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   1500
         TabIndex        =   108
         Top             =   2100
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "SUM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   2220
         TabIndex        =   107
         Top             =   2100
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "MUL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   2940
         TabIndex        =   106
         Top             =   2100
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "IND"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   2100
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   2100
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "Hkey"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   5100
         TabIndex        =   103
         Top             =   2100
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "lnX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   5820
         TabIndex        =   102
         Top             =   2100
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "å+"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   6540
         TabIndex        =   101
         Top             =   2100
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "Mean"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   7260
         TabIndex        =   100
         Top             =   2100
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "X!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   7980
         TabIndex        =   99
         Top             =   2100
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "X><T"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   8700
         TabIndex        =   98
         Top             =   2100
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "Arc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   60
         TabIndex        =   97
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "Sin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   780
         TabIndex        =   96
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "Cos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   1500
         TabIndex        =   95
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "Tan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   2220
         TabIndex        =   94
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "1/X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   30
         Left            =   2940
         TabIndex        =   93
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00DF967A&
         Caption         =   "Txt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   31
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "Hex"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   32
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "&&"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   33
         Left            =   5100
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "StFlg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   34
         Left            =   5820
         TabIndex        =   89
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "IfFlg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   35
         Left            =   6540
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "X==T"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   36
         Left            =   7260
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "X>=T"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   37
         Left            =   7980
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "X>T"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   38
         Left            =   8700
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Dfn"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   39
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   3060
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   ";"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   40
         Left            =   780
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   3060
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "("
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   41
         Left            =   1500
         TabIndex        =   82
         Top             =   3060
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   ")"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   42
         Left            =   2220
         TabIndex        =   81
         Top             =   3060
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H0000C0FF&
         Caption         =   "÷"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   43
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   3060
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "Style"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   44
         Left            =   3660
         TabIndex        =   79
         Top             =   3060
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "Dec"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   45
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   3060
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "|"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   46
         Left            =   5100
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   3060
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "Int"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   47
         Left            =   5820
         TabIndex        =   76
         Top             =   3060
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "Abs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   48
         Left            =   6540
         TabIndex        =   75
         Top             =   3060
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "Fix"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   49
         Left            =   7260
         TabIndex        =   74
         Top             =   3060
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "D.MS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   50
         Left            =   7980
         TabIndex        =   73
         Top             =   3060
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "EE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   51
         Left            =   8700
         TabIndex        =   72
         Top             =   3060
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Sbr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   52
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   3540
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   53
         Left            =   780
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   3540
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   54
         Left            =   1500
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   3540
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   55
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   3540
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H0000C0FF&
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   56
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   3540
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "'"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   57
         Left            =   3660
         TabIndex        =   66
         Top             =   3540
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "Oct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   58
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   3540
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   59
         Left            =   5100
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   3540
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   60
         Left            =   5820
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   3540
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Case"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   61
         Left            =   6540
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   3540
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "{"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   62
         Left            =   7260
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   3540
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "}"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   63
         Left            =   7980
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   3540
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "Deg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   64
         Left            =   8700
         TabIndex        =   59
         Top             =   3540
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Lbl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   65
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   66
         Left            =   780
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   67
         Left            =   1500
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   68
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H0000C0FF&
         Caption         =   "¾"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   69
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "Beep"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   70
         Left            =   3660
         TabIndex        =   53
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "Bin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   71
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "^"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   72
         Left            =   5100
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "For"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   73
         Left            =   5820
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Do"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   74
         Left            =   6540
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "While"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   75
         Left            =   7260
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Pmt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   76
         Left            =   7980
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "Rad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   77
         Left            =   8700
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "UKey"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   78
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4500
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   79
         Left            =   780
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   4500
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   80
         Left            =   1500
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   4500
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   81
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4500
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H0000C0FF&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   82
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4500
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Plot"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   83
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   4500
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "Nvar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   84
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4500
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   85
         Left            =   5100
         TabIndex        =   38
         Top             =   4500
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "If"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   86
         Left            =   5820
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   4500
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Else"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   87
         Left            =   6540
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   4500
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Cont"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   88
         Left            =   7260
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   4500
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Break"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   89
         Left            =   7980
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   4500
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "Grad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   90
         Left            =   8700
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4500
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H0000FF00&
         Caption         =   "R/S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   91
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   4980
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   92
         Left            =   780
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   4980
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   93
         Left            =   1500
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   4980
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "+/-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   94
         Left            =   2220
         TabIndex        =   29
         Top             =   4980
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H0000C0FF&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   95
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4980
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   96
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   4980
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "Tvar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   97
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4980
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   98
         Left            =   5100
         TabIndex        =   25
         Top             =   4980
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "y^"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   99
         Left            =   5820
         TabIndex        =   24
         Top             =   4980
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "X²"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   100
         Left            =   6540
         TabIndex        =   23
         Top             =   4980
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "p"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   101
         Left            =   7260
         TabIndex        =   22
         Top             =   4980
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "Rnd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   102
         Left            =   7980
         TabIndex        =   21
         Top             =   4980
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "Mil"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   103
         Left            =   8700
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4980
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Pvt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   104
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5460
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Const"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   105
         Left            =   780
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   5460
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Struct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   106
         Left            =   1500
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5460
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "PvLbl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   108
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5460
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Line"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   109
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5460
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "["
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   110
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5460
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   111
         Left            =   5100
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5460
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "ClrVar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   112
         Left            =   5820
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5460
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "SzOf"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   113
         Left            =   6540
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5460
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Def"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   114
         Left            =   7260
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5460
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "IfDef"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   115
         Left            =   7980
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5460
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Edef"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   116
         Left            =   8700
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5460
         Width           =   615
      End
      Begin VB.CommandButton cmdKeyPad 
         BackColor       =   &H00E1FFFF&
         Caption         =   "NxLbl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   107
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5460
         Width           =   615
      End
      Begin VB.Line Line24 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   8640
         X2              =   8640
         Y1              =   4920
         Y2              =   5400
      End
      Begin VB.Line Line20 
         BorderColor     =   &H8000000D&
         X1              =   3600
         X2              =   3600
         Y1              =   1980
         Y2              =   2520
      End
      Begin VB.Line Line19 
         BorderColor     =   &H8000000D&
         X1              =   4320
         X2              =   4320
         Y1              =   1980
         Y2              =   2520
      End
      Begin VB.Line Line18 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   3600
         X2              =   3600
         Y1              =   1980
         Y2              =   1440
      End
      Begin VB.Line Line11 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   8640
         X2              =   9300
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   2160
         X2              =   5760
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   3600
         X2              =   9300
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         Height          =   495
         Left            =   5760
         Top             =   1500
         Width           =   2895
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   720
         X2              =   0
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line23 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   3600
         X2              =   0
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line22 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   720
         X2              =   720
         Y1              =   5400
         Y2              =   3480
      End
      Begin VB.Line Line21 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   5760
         X2              =   9360
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line17 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   1440
         X2              =   720
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line16 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   1440
         X2              =   1440
         Y1              =   3000
         Y2              =   3480
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         Height          =   1455
         Left            =   5760
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000015&
         BorderWidth     =   3
         X1              =   0
         X2              =   9360
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   3600
         X2              =   3600
         Y1              =   5880
         Y2              =   2520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   5760
         X2              =   5760
         Y1              =   5400
         Y2              =   1980
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   4320
         X2              =   9360
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   7200
         X2              =   7200
         Y1              =   5400
         Y2              =   5880
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   3600
         X2              =   0
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         Height          =   495
         Left            =   0
         Top             =   1500
         Width           =   2175
      End
      Begin VB.Line Line12 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   5040
         X2              =   5040
         Y1              =   1440
         Y2              =   5400
      End
      Begin VB.Line Line13 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   4320
         X2              =   4320
         Y1              =   2520
         Y2              =   5820
      End
      Begin VB.Line Line14 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   3600
         X2              =   5040
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line15 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   2160
         X2              =   2160
         Y1              =   5400
         Y2              =   5880
      End
   End
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   645
      ScaleHeight     =   5895
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   60
      Width           =   4635
      Begin RichTextLib.RichTextBox rtbSearch 
         Height          =   1275
         Left            =   2040
         TabIndex        =   174
         Top             =   3960
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   2249
         _Version        =   393217
         BackColor       =   16777215
         TextRTF         =   $"frmVisualCalc.frx":36DF
      End
      Begin RichTextLib.RichTextBox rtbInfo 
         Height          =   1275
         Left            =   240
         TabIndex        =   173
         Top             =   3960
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2249
         _Version        =   393217
         BackColor       =   14811135
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmVisualCalc.frx":3767
      End
      Begin VB.TextBox txtError 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   1
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   128
         Text            =   "frmVisualCalc.frx":37ED
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtError 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   585
         Index           =   0
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   127
         Text            =   "ERROR"
         Top             =   600
         Width           =   3315
      End
      Begin VB.PictureBox PicPlot 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H80000005&
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1485
         Left            =   300
         ScaleHeight     =   97
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   126
         Top             =   2280
         Width           =   1545
      End
      Begin VB.ListBox lstDisplay 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   5520
         ItemData        =   "frmVisualCalc.frx":3801
         Left            =   0
         List            =   "frmVisualCalc.frx":3803
         MultiSelect     =   2  'Extended
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   4635
      End
      Begin VB.Label lblTxt 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TXT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2250
         TabIndex        =   134
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblInstr 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Instr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1440
         TabIndex        =   133
         Top             =   0
         Width           =   570
      End
      Begin VB.Label lblCode 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   660
         TabIndex        =   132
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblLoc 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Step"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   0
         TabIndex        =   131
         Top             =   0
         Width           =   525
      End
      Begin VB.Label lblINS 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008800&
         Height          =   300
         Left            =   2652
         TabIndex        =   129
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblLRN 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LRN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3054
         TabIndex        =   125
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblDRGM 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Deg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3858
         TabIndex        =   6
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblHDOB 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dec"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4260
         TabIndex        =   5
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblEE 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3456
         TabIndex        =   4
         Top             =   0
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar sbrImmediate 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6255
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14208
            Text            =   "Tip Window"
            TextSave        =   "Tip Window"
            Key             =   "Tips"
            Object.ToolTipText     =   "Quick tips and Quick Listings display"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1296
            MinWidth        =   706
            Text            =   "No MDL"
            TextSave        =   "No MDL"
            Key             =   "MDL"
            Object.ToolTipText     =   "Currently loaded Module"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1323
            MinWidth        =   706
            Text            =   "No Pgm"
            TextSave        =   "No Pgm"
            Key             =   "Pgm"
            Object.ToolTipText     =   "Currently active program"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1376
            MinWidth        =   706
            Text            =   "Steps: 0"
            TextSave        =   "Steps: 0"
            Key             =   "InstrCnt"
            Object.ToolTipText     =   "Number of program steps recorded"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1455
            MinWidth        =   706
            Text            =   "Step #: 0"
            TextSave        =   "Step #: 0"
            Key             =   "InstrPtr"
            Object.ToolTipText     =   "Current program step index"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1244
            MinWidth        =   706
            Text            =   "Style: 0"
            TextSave        =   "Style: 0"
            Key             =   "Style"
            Object.ToolTipText     =   "Click to cycle through Style options"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1640
            MinWidth        =   706
            Text            =   "Trace: Off"
            TextSave        =   "Trace: Off"
            Key             =   "Tron"
            Object.ToolTipText     =   "Click to toggle trace mode on and off"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1429
            MinWidth        =   706
            TextSave        =   "4/3/2007"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1349
            MinWidth        =   706
            TextSave        =   "8:06 AM"
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
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      Caption         =   "lblWidth"
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
      Left            =   1920
      TabIndex        =   175
      Top             =   6000
      Width           =   960
   End
   Begin VB.Label lblChkSize 
      AutoSize        =   -1  'True
      Caption         =   "lblChkSize"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   145
      Top             =   6000
      Width           =   1200
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuFileCopy 
         Caption         =   "Copy &entire display list to Clipboard"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileCopySel 
         Caption         =   "Copy &selection to Clipboard"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFilePaste 
         Caption         =   "Paste clipboard to the display"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNotepad 
         Caption         =   "Launch &Notepad application"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileWordPad 
         Caption         =   "Launch &WordPad application"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileListDir 
         Caption         =   "List &directory of Data Storage Location"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "&Import Selected ASCII file from directory list"
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileImportClipBrd 
         Caption         =   "Impo&rt ASCII program from Clipboard"
      End
      Begin VB.Menu mnuFileImportSegment 
         Caption         =   "Import ASCII program SE&GMENT from clipboard"
      End
      Begin VB.Menu mnuFIleSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrcmp 
         Caption         =   "Automatically preprocess programs"
      End
      Begin VB.Menu mnuFileReloadMDL 
         Caption         =   "Automatically reload last-active Module on start-up"
      End
      Begin VB.Menu mnuFileReloadPgm 
         Caption         =   "Automatically reload last-active Program file on start-up"
      End
      Begin VB.Menu mnuFileSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileTron 
         Caption         =   "Toggle program Trace Mode"
      End
      Begin VB.Menu mnufileTypeMatic 
         Caption         =   "Toggle TypeMatic Keyboard"
      End
      Begin VB.Menu mnuFileTglCoDisplay 
         Caption         =   "Toggle Co-Display of Formatted Source in Learn Mode"
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&View"
      Begin VB.Menu mnuWinGreenScreen 
         Caption         =   "Toggle &Green Screen"
      End
      Begin VB.Menu mnuWinRight 
         Caption         =   "Display on &Right side"
      End
      Begin VB.Menu mnuKeypad 
         Caption         =   "&Keypad display level"
         Begin VB.Menu mnuKeypadBasic 
            Caption         =   "&Basic keypad layout"
         End
         Begin VB.Menu mnuKeypadAdvanced 
            Caption         =   "&Advanced keypad layout"
         End
         Begin VB.Menu mnuKeypadFull 
            Caption         =   "&Programmer's keypad layout (full)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuWinSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWinListVdata 
         Caption         =   "List &Variable contents"
      End
      Begin VB.Menu mnuWinListVdataNZ 
         Caption         =   "List &non-null Variable contents"
      End
      Begin VB.Menu mnuWinSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWINListMDL 
         Caption         =   "List &Module program directory"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuWinSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWinASCII 
         Caption         =   "List Pgm, formatted to Style"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuWinVar 
         Caption         =   "List V&ariable (Xvar)  declarations"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuWinLbl 
         Caption         =   "List &Labels (Lbl)"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuWinUkey 
         Caption         =   "List &User-defined keys (Ukey)"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuWinSbr 
         Caption         =   "List &Subroutines (Sbr)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuWinConst 
         Caption         =   "List C&onstants (Const)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuWinStruct 
         Caption         =   "List S&tructures (Struct)"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpIntro 
         Caption         =   "Introduction to MyApp..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpGetStarted 
         Caption         =   "Getting Started with MyApp..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuHelpCodes 
         Caption         =   "Program codes used by MyApp..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "Sorted Index of Command keys..."
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuHelpTrig 
         Caption         =   "Trigonometric Functions..."
      End
      Begin VB.Menu mnuHelpHistory 
         Caption         =   "A Brief History of Programmable Calculators..."
      End
      Begin VB.Menu mnuHelpSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpSrch 
         Caption         =   "&Search for Topic..."
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuHelpFind 
         Caption         =   "Find all text matches..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuHelpNext 
         Caption         =   "Go to Next match..."
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuHelpPrev 
         Caption         =   "Go to Previous match..."
         Shortcut        =   +^{F3}
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpSepHlp 
         Caption         =   "Show &help in a separate window"
      End
      Begin VB.Menu mnuHelpSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About MyApp..."
      End
   End
   Begin VB.Menu mnuMainSep 
      Caption         =   "                                                                                   "
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuMemStk 
      Caption         =   "Define Base Data Storage  Location"
   End
   Begin VB.Menu mnuPopUpHelp 
      Caption         =   "mnuPopUpHelp"
      Begin VB.Menu mnuPopupSelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuPopupSearch 
         Caption         =   "Search"
      End
      Begin VB.Menu mnuPopupFind 
         Caption         =   "Find all text matches"
      End
      Begin VB.Menu mnuPopupCopy 
         Caption         =   "&Copy selection"
      End
   End
End
Attribute VB_Name = "frmVisualCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************************
' API stuff
'*******************************************************************************
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_LBUTTON = &H1
'
' local storage
'
Private OverForm As Boolean   'used when mouse moves over non-buttons
Private LastPlotXt As Single  'use to track X and Y positions over the plot window
Private LastPlotYt As Single
Private Resize As Boolean     'flag True when resizing being processed
Private NoQuery As Boolean    'used by 2nd key and Ctrl
Private DblClkRtb As Boolean  'flag used to record DblClick in rtbInfo

Dim myShadow As clsShadow     'form Shadow class reference

'*******************************************************************************
' Subroutine Name   : Form_Initialize
' Purpose           : Add XP button support
'*******************************************************************************
Private Sub Form_Initialize()
  Call FormInitialize
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Set keys as required for startup
'*******************************************************************************
Private Sub Form_Load()
  Dim Idx As Integer, i As Integer, Wid As Integer, Hit As Integer, Bts As Integer
  Dim Lf As Long, Wd As Long, Tp As Long, HT As Long  'screen size storage
  Dim S As String, T As String, Cmd As String, CmdAry() As String
  Dim IsRdy As Boolean          'if Data Storage Location valid
  Dim DoIntro As Boolean        'flag TRUE if intro help should be displayed
  Dim Drv As Drive              'drive reference
  Dim LoadedMDL As Boolean      'True if user loaded a module on the command line
  Dim LoadedPGM As Boolean      'True if user loaded a program on the command line
'
' determine form caption and beta status
'
  If CBool(App.Major) Then      'if not Beta code
    Me.Caption = App.Title & " " & GetAppVersion & ": " & AppTitle & " "
  Else                          'beta
    Me.Caption = App.Title & " (Beta) " & GetAppVersion & ": " & AppTitle & " "
  End If
'
' check for previous instance, and activate it
'
  If App.PrevInstance Then      'previous instance exists?
    If ActivatePrv(Me) Then     'activate previous instance if found
      Unload Me                 'unload current instance
      Exit Sub                  'leave and let previous instance assume control
    End If
  End If
'
' we will be running this instance, so set it up
'
  Set Fso = New FileSystemObject    'enable file I/O
  Set colHelpBack = New Collection  'set aside new collection for help list
  Set colFindList = New Collection  'and find list
  
  S = GetCurrentDisplaySize()       'get display width and height, color, color bits, etc.
  Call GetXYBits(S, Wid, Hit, Bts)  'extract width and height from function response
  If Wid < 1024 Or Hit < 768 Then   'must be at least 1024 x 768 pixels
    S = "Current Resulution is: " & S & vbCrLf
    If Wid < 1024 Then S = S & "Screen width must be at least 1024." & vbCrLf
    If Hit < 768 Then S = S & "Screen height must be at least 768." & vbCrLf
    S = S & vbCrLf & "Correct this resultion problem and try again."
    CenterMsgBoxOnForm Me, S, vbOKOnly Or vbExclamation, "Screen Resolution Too Small"
    IsDirty = False                 'prevent dirty data prompts from kicking in
    ModDirty = False
    Unload Me
    Exit Sub
  End If
  
  IsDirty = False               'pgm space not dirty (not altered)
  ModDirty = False              'module space not dirty
  Cmd = Trim$(Command$)         'get command line
  
  Me.Height = WinMinH           'used in IDE to ensure proper height (VB6 sometimes bugs out
                                'when using sizable window with full win button disabled)
  Me.PicScroll.Left = 60        'init control placement on form
  Me.picDisplay.Left = Me.PicScroll.Width + 120
  Me.PicKeys.Left = Me.ScaleWidth - Me.PicKeys.Width - 60
  Me.mnuPopUpHelp.Visible = False
  Me.mnuHelpIntro.Caption = AddTitle(Me.mnuHelpIntro.Caption)
  Me.mnuHelpGetStarted.Caption = AddTitle(Me.mnuHelpGetStarted.Caption)
  Me.mnuHelpCodes.Caption = AddTitle(Me.mnuHelpCodes.Caption) 'add app title to message
  Me.mnuHelpAbout.Caption = AddTitle(Me.mnuHelpAbout.Caption)
  
  BackClr = vbWhite             'init displaylist background color
  Me.mnuFileExit.Caption = "E&xit" & vbTab & "Alt+F4"
  ModName = 0                   'turn off module name (number)
  ModLbl = " "                  'blank label
  PgmName = 0                   'and Pgm name
  LrnMode = False               'turn off LRN mode
  BaseType = TypDec             'default numberic base to 10
  AngleType = TypDeg            'angles are in degrees
  DspFmt = DefDspFmt            'remove any special display formatting
  DspFmtFix = -1                'indicate no fixed decimal places
  TraceFlag = False             'ensure tracing is off
  Call CP_Support               'clear enverything out, reinitialize program
  vRA = Atn(1#) * 2#            'define value of right angle (90-degrees) in radians
  vPi = Atn(1#) * 4#            'define value of PI (1/2 circle or 180 degrees)
  vPi2 = CSng(vPi * 2#)         'single precision radians for a circle
  vE = Exp(1#)                  'define value of e (epsilon)
  ScientifEE = DefScientific    'reset default scientif mode
  'get single line height in twips (used in Plot mode)
  LineHeight = CSng(Me.lblChkSize.Height \ Screen.TwipsPerPixelY)
'
' init plot window
'
  With Me.PicPlot
    .Visible = False              'hide plot window
    .AutoSize = False             'disable frame autosizing
    .AutoRedraw = True            'enable for display persistence
    .DrawMode = vbCopyPen         'over-write palette
    .ScaleMode = vbPixels         'pixel mode
    .DrawStyle = vbSolid          'so special drawing features
    .DrawWidth = 1                '1 pixel wide pen
    .FillStyle = 0                'Solid
    .Left = 0                     'flush left in container
    .Top = Me.lstDisplay.Top      'match to main Display list
    .Width = Me.lstDisplay.Width
    .Height = Me.picDisplay.Height - .Top - 30
  End With
'
' init error report display
'
  With Me.txtError(0)
    .Visible = False
    .Left = 8
    .Width = Me.picDisplay.Width - 16
    .Top = Me.lstDisplay.Top
    .BorderStyle = 0
  End With
  
  With Me.txtError(1)
    .Visible = False
    .Left = 8
    .Width = Me.picDisplay.Width - 16
    .BorderStyle = 0
    .Top = Me.lstDisplay.Top + Me.txtError(0).Height
    .Height = Me.lstDisplay.Height - Me.txtError(0).Height
  End With
'
' Set up Help data display field
'
  With Me.rtbInfo
    .Left = 0
    .Top = Me.lstDisplay.Top
    .Width = Me.picDisplay.Width - 120
    .Height = Me.picDisplay.Height - .Top - 60
    .Visible = False
  End With
'
' Set up background search data
'
  With Me.rtbSearch
    .Left = 0
    .Top = 0
    .Width = Me.ScaleWidth
    .Height = Me.ScaleHeight
    .Visible = False
  End With
  
  Me.lblLoc.BackStyle = 0   'disable learn mode headers
  Me.lblCode.BackStyle = 0
  Me.lblInstr.BackStyle = 0
  
  SetTip vbNullString               'init button tips field
  InitTileFormBackground Me.PicBack 'init tiling for background
  Me.txtFocus.Left = Me.Left - 1440 'hide locked focus field
  Me.lblChkSize.Top = -2880         'hide sizing fields
  Me.lblWidth.Top = -1440
'
' see if entry exists. If not, set DoIntro flag to True
'
  DoIntro = Not CBool(GetSetting(App.Title, "Settings", "Width", "0"))
'
' Get flag for showing help in a separate window
'
  LastQuery = 9999                  'disable "previous" query
  Me.mnuHelpSepHlp.Checked = CBool(GetSetting(App.Title, "Settings", "SepHelp", "1"))
'
' check TypMatic keyboard
'
  With Me.mnufileTypeMatic
    TypeMatic = CBool(GetSetting(App.Title, "Settings", "TypeMatic", "0"))
    .Checked = TypeMatic
  End With
'
' check co-display option
'
  Me.mnuFileTglCoDisplay.Checked = CBool(GetSetting(App.Title, "Settings", "CoDisplay", "1"))
'
' get auto Preprocess option
'
  AutoPprc = CBool(GetSetting(App.Title, "Settings", "AutoPprc", "1"))
  Me.mnuFilePrcmp.Checked = AutoPprc
'
' set Formatted output option (0-3)
'
  LRNstyle = CInt(GetSetting(App.Title, "Settings", "Style", "1"))
'
' center the form on screen if the user has not previously changed the location
'
  Call GetScreenWorkArea(Lf, Wd, Tp, HT)
  Resize = True     'delay and form resizing
  Me.Width = CLng(GetSetting(App.Title, "Settings", "Width", CStr(WinMinW)))
  Me.Height = CLng(GetSetting(App.Title, "Settings", "Height", CStr(WinMinH)))
  Me.Left = CLng(GetSetting(App.Title, "Settings", "Left", CStr((Wd - Me.Width) \ 2 + Lf)))
  Me.Top = CLng(GetSetting(App.Title, "Settings", "Top", CStr((HT - Me.Height) \ 2 + Tp)))
  Me.WindowState = CLng(GetSetting(App.Title, "Settings", "WinState", CStr(vbNormal)))
  Resize = False    'stop blocking resize processing
  Call Form_Resize  'force resize processing
'
' adjust form further by checking the display side option, and the screen color option
'
  If CBool(GetSetting(App.Title, "Settings", "WinRight", "0")) Then Call mnuWinRight_Click
  If CBool(GetSetting(App.Title, "Settings", "GreenScreen", "0")) Then Call mnuWinGreenScreen_Click
'
' set auto-load states
'
  Me.mnuFileReloadMDL.Checked = CBool(GetSetting(App.Title, "Settings", "ReloadMDL", "1"))
  Me.mnuFileReloadPgm.Checked = CBool(GetSetting(App.Title, "Settings", "ReloadPgm", "0"))
'
' get the default data storage path
'
  StorePath = GetSetting(App.Title, "Settings", "StorePath", vbNullString)
  If CBool(Len(StorePath)) Then
    StorePath = RemoveSlash(StorePath)
    If Not Fso.FolderExists(StorePath) Then StorePath = vbNullString
  End If
  If Not CBool(Len(StorePath)) Then
    StorePath = AddSlash(App.Path) & "Samples"
    If Not Fso.FolderExists(StorePath) Then StorePath = vbNullString
  End If
  
  IsRdy = False                                             'init Storage path as not found
  If CBool(Len(StorePath)) Then
    If Left$(StorePath, 2) = "\\" Then                      'network path?
      Idx = InStr(3, StorePath, "\")                        'yes, so extract net drive
      If CBool(Idx) Then                                    'if we found network drive root
        Idx = InStr(Idx + 1, StorePath, "\")                'find end drive drive parm
        If CBool(Idx) Then Idx = Idx - 1                    'do not include '\'
      End If
    Else
      Idx = InStr(1, StorePath, ":")                        'else extract Drive data
    End If
    If CBool(Idx) Then
      S = Left$(StorePath, Idx)                             'grab drive
      On Error Resume Next
      Set Drv = Fso.GetDrive(S)                             'get drive device
      If Not CBool(Err.Number) Then                         'if all OK
        On Error GoTo 0
        Do
          If Fso.DriveExists(Drv) Then                      'if drive exists...
            If Drv.IsReady Then                             'if drive ready...
              If Fso.FolderExists(StorePath) Then           'if folder exists...
                If Fso.FolderExists(StorePath & "\MDL") And _
                   Fso.FolderExists(StorePath & "\PGM") And _
                   Fso.FolderExists(StorePath & "\DATA") Then
                  IsRdy = True                              'already initialized
                  Exit Do
                End If
              End If
            End If
          End If
          If Not IsRdy Then                                 'if Ready flag not yet set
            If CenterMsgBoxOnForm(Me, "User-specified Data Storage path: " & StorePath & vbCrLf & vbCrLf & _
                  "Warning, current default Data Storage location appears to be invalid." & vbCrLf & vbCrLf & _
                  "Please prepare it and select Retry, or select Cancel to ignore it...", _
                  vbRetryCancel Or vbExclamation, "Data Storage Error") = vbCancel Then Exit Do
          End If
        Loop
      End If
    End If
    On Error GoTo 0
    With Me.mnuMemStk
      If Not IsRdy Then
        StorePath = vbNullString
        .Caption = "Define Base Data Storage Location"      'storage path is not ready
      Else
        .Caption = "Base Data Storage Location: " & StorePath
        Call SaveSetting(App.Title, "Settings", "StorePath", StorePath)
      End If
    End With
  End If
'
' Aquire Notepad path
'
  S = GetWindowsDir()
  NotePadPath = AddSlash(S) & "Notepad.exe"
  If Len(Dir$(NotePadPath)) = 0 Then
    NotePadPath = AddSlash(GetSystemDir()) & "Notepad.exe"
  End If
  Me.mnuFileNotepad.Enabled = CBool(Len(Dir$(NotePadPath)))
'
' Aquire WordPad path
'
  i = InStrRev(S, "\")
  S = Left$(S, i) & "Program Files\"
  WordPadPath = S & "Accessories"
  If CBool(Len(Dir$(WordPadPath, vbDirectory))) Then
    WordPadPath = WordPadPath & "\WordPad.exe"
    If Not CBool(Len(Dir$(WordPadPath))) Then WordPadPath = vbNullString
  Else
    WordPadPath = vbNullString
  End If
  If Not CBool(Len(WordPadPath)) Then
    WordPadPath = S & "Windows NT\Accessories"
    If CBool(Len(Dir$(WordPadPath, vbDirectory))) Then
      WordPadPath = WordPadPath & "\WordPad.exe"
      If Not CBool(Len(Dir$(WordPadPath))) Then WordPadPath = vbNullString
    End If
  End If
  Me.mnuFileWordPad.Enabled = CBool(Len(WordPadPath))
'
' Check VPCHelp.rtf path
'
  HaveVHelp = Fso.FileExists(AddSlash(App.Path) & "VPCHelp.rtf")
  Me.cmdBtnHelp.Enabled = HaveVHelp
  Me.mnuHelpGetStarted.Enabled = HaveVHelp
  Me.mnuHelpIntro.Enabled = HaveVHelp
  Me.mnuHelpCodes.Enabled = HaveVHelp
  Me.mnuHelpSrch.Enabled = HaveVHelp
  Me.mnuHelpFind.Enabled = HaveVHelp
  Me.mnuHelpSepHlp.Enabled = HaveVHelp
  If DoIntro Then DoIntro = HaveVHelp
'
' init help items
'
  Call ClearFindList
'
' enable view user-last selected (default is full)
'
  Select Case CInt(GetSetting(App.Title, "Settings", "KeyLayout", "2"))
    Case 0
      Call mnuKeypadBasic_Click
    Case 1
      Call mnuKeypadAdvanced_Click
    Case Else
      Call mnuKeypadFull_Click
  End Select
'
' command list entered
'
  LoadedMDL = False
  LoadedPGM = False
  If CBool(Len(Cmd)) Then
    If CBool(InStr(1, Cmd, ",")) Then
      CmdAry = Split(Cmd, ",")
    Else
      CmdAry = Split(Cmd, " ")
    End If
    
    ErrorFlag = False
    For Idx = 0 To UBound(CmdAry)
      S = Trim$(CmdAry(Idx))
      If CBool(Len(S)) Then
        Select Case UCase$(Right$(S, 4))
          Case ".PGM"
            S = Mid$(S, 4)                'strip "PGM" from left
            S = Left$(S, Len(S) - 4)      'strip ".pgm" from right
            If IsNumeric(S) Then
              DisplayReg = Val(S)         'set pgm # to main register
              PndIdx = 1                  'force LOAD onto pending stack
              PndStk(1) = iLoad
              Call CheckPnd(0)            'load program
              LoadedPGM = True
            End If
          Case ".MDL"
            S = Mid$(S, 4)                'strip "MDL" from left
            S = Left$(S, Len(S) - 4)      'strip ".mdl" from right
            If IsNumeric(S) Then
              DisplayReg = Val(S)         'set module # to main register
              PndIdx = 2                  'force LOAD onto pending stack
              PndStk(1) = iMDL            'set MDL Load on pending stack
              PndStk(2) = iLoad
              Call CheckPnd(0)            'load program
              LoadedMDL = True
            End If
        End Select
      End If
      If ErrorFlag Or Not IsNumeric(S) Then Exit For
    Next Idx
  End If
'
' see if we should reload last-saved module
'
  If Me.mnuFileReloadMDL.Checked And Not LoadedMDL Then
    i = 0
    Idx = CInt(GetSetting(App.Title, "Settings", "LoadedMDL", "0"))
    If CBool(Idx) Then
      If CBool(Len(StorePath)) Then
        S = AddSlash(StorePath) & "MDL"
        If Fso.FolderExists(S) Then
          T = "MDL" & Format(Idx, "0000") & ".mdl"
          S = AddSlash(S) & T
          If Fso.FileExists(S) Then i = Idx
        End If
      End If
    End If
    
    If CBool(i) Then
      DisplayReg = CDbl(i)
      Call Load_MDL
    Else
      SaveSetting App.Title, "Settings", "LoadedMDL", "0"
    End If
  End If
'
' see if we should reload last-saved program
'
  If Me.mnuFileReloadPgm.Checked And Not LoadedPGM Then
    i = 0
    Idx = CInt(GetSetting(App.Title, "Settings", "LoadedPgm", "0"))
    If CBool(Idx) Then
      If CBool(Len(StorePath)) Then
        S = AddSlash(StorePath) & "PGM"
        If Fso.FolderExists(S) Then
          T = "PGM" & Format(Idx, "00") & ".pgm"
          S = AddSlash(S) & T
          If Fso.FileExists(S) Then i = Idx
        End If
      End If
    End If
    
    If CBool(i) Then
      DisplayReg = CDbl(i)
      PndIdx = 1                    'force LOAD onto pending stack
      PndStk(1) = iLoad
      Call CheckPnd(0)              'load program
    Else
      SaveSetting App.Title, "Settings", "LoadedPgm", "0"
    End If
  End If
'
' see if intro should be shown
'
  If DoIntro Then
    Me.Show
    DoEvents
    Query -13
  End If
'
' hook window for sizing control
' Disable the following line if you will be debugging form.
'
  Call HookWin(Me.hWnd, m_VChWnd)
'
' apply a shadow to form
' Disable the following block of code if you will be debugging form.
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
' Function Name     : ExitApp
' Purpose           : Handle exiting the program. Check for LRN data
'*******************************************************************************
Private Function ExitApp() As Integer
  ExitApp = 0                 'init to success
  Me.PicPlot.Visible = False
  If ModDirty Then
    If CenterMsgBoxOnForm(Me, "The Module has unsaved data." & vbCrLf & vbCrLf & _
                              "Go ahead and exit?", _
                              vbYesNo Or vbExclamation Or vbDefaultButton2, _
                              "Usaved Module Data") = vbNo Then ExitApp = 1
  End If
  If ExitApp = 0 Then
    If IsDirty Then
      If CenterMsgBoxOnForm(Me, "The program buffer has unsaved data." & vbCrLf & vbCrLf & _
                                "Go ahead and exit?", _
                                vbYesNo Or vbExclamation Or vbDefaultButton2, _
                                "Usaved Program Code") = vbNo Then ExitApp = 1
    End If
  End If
'
' save currently loaded Module and Pgm
'
  SaveSetting App.Title, "Settings", "LoadedMDL", CStr(ModName)
  SaveSetting App.Title, "Settings", "LoadedPgm", CStr(PgmName)
End Function

'*******************************************************************************
' Subroutine Name   : Form_MouseMove
' Purpose           : Clear tips if cursor over form
'*******************************************************************************
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If CBool(CharLimit) Then Exit Sub
  OverForm = True
  SetTip vbNullString
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Paint
' Purpose           : Paint the background tiling
'*******************************************************************************
Private Sub Form_Paint()
  PaintTileFormBackground Me, Me.PicBack
End Sub

'*******************************************************************************
' Subroutine Name   : Form_QueryUnload
' Purpose           : Unloading Form
'*******************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then  'if we are using the X button...
    Cancel = ExitApp()                    'cancel exit if user chooses to
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Handle resizing the main form
'*******************************************************************************
Private Sub Form_Resize()
  If Resize Then Exit Sub                           'if handling resizing...
  Select Case Me.WindowState
    Case vbMinimized                                'minimized?
      Exit Sub                                      'do nothing
    Case vbNormal
      If GetKeyState(VK_LBUTTON) < 0 Then           'if left mouse button down
        With Me.tmrWait                             'let timer handle fix
          .Enabled = False                          'disable timer
          DoEvents                                  'let screen catch up
          .Enabled = True                           're-enable timer (also resets it)
        End With
        Exit Sub
      End If
      If Me.Width < WinMinW Or Me.Height < WinMinH Then 'if sizing goes too small...
        With Me.tmrWait                             'let timer handle fix
          .Enabled = False                          'disable timer
          DoEvents                                  'let screen catch up
          .Enabled = True                           're-enable timer (also resets it)
        End With
        Exit Sub
      End If
    Case vbMaximized
      If Me.Width < WinMinW Then Me.Width = WinMinW   'resize to minimum dims
      If Me.Height < WinMinH Then Me.Height = WinMinH
  End Select
  
  With Me.picDisplay
    .Width = Me.ScaleWidth - Me.PicKeys.Width - Me.PicScroll.Width - 120
    .Height = Me.ScaleHeight - Me.sbrImmediate.Height - 60
    Me.lstDisplay.Width = .Width - 120
    Me.lstDisplay.Height = .Height - Me.lstDisplay.Top
    Me.PicScroll.Height = .Height - 60
    Me.txtError(0).Width = .Width - 16
    Me.txtError(1).Width = .Width - 16
    Me.txtError(1).Height = .Height - Me.txtError(0).Height
    Me.lblHDOB.Left = .Width - Me.lblHDOB.Width - 120
  End With
  
  Me.lblDRGM.Left = Me.lblHDOB.Left - Me.lblDRGM.Width - 30
  Me.lblEE.Left = Me.lblDRGM.Left - Me.lblEE.Width - 30
  Me.lblLRN.Left = Me.lblEE.Left - Me.lblLRN.Width - 30
  Me.lblINS.Left = Me.lblLRN.Left - Me.lblINS.Width - 30
  Me.lblTxt.Left = Me.lblINS.Left - Me.lblTxt.Width - 30
  
  With Me.rtbInfo
    .Width = Me.picDisplay.Width - 120
    .Height = Me.picDisplay.Height - .Top - 60
  End With
  With Me.rtbSearch
    .Width = Me.ScaleWidth
    .Height = Me.ScaleHeight
  End With
  
  If mnuWinRight.Checked Then                     'to right...
    Me.PicKeys.Left = 60
    Me.PicScroll.Left = Me.ScaleWidth - Me.PicScroll.Width - 60
    Me.picDisplay.Left = Me.PicKeys.Width + 120
  Else                                            'to left...
    Me.PicScroll.Left = 60
    Me.picDisplay.Left = Me.PicScroll.Width + 120
    Me.PicKeys.Left = Me.ScaleWidth - Me.PicKeys.Width - 60
  End If
  DoEvents
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Unloading Form
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Dim Idx As Long
'
' clear out all variable class objects
'
  For Idx = 0 To MaxVar
    Set Variables(Idx).Vdata = Nothing
  Next Idx
  Set Fso = Nothing
  
  Set colHelpBack = Nothing
  Set colFindList = Nothing
  '
  ' save window location if not minimized
  '
  If Me.WindowState <> vbMinimized Then
    If Me.WindowState = vbNormal Then
      SaveSetting App.Title, "Settings", "Left", CStr(Me.Left)
      SaveSetting App.Title, "Settings", "Top", CStr(Me.Top)
      SaveSetting App.Title, "Settings", "Width", CStr(Me.Width)
      SaveSetting App.Title, "Settings", "Height", CStr(Me.Height)
    End If
    SaveSetting App.Title, "Settings", "WinState", CStr(Me.WindowState)
  End If
'
' ensure secondary forms unloaded
'
  Unload frmHelp
  Unload frmAbout
  Unload frmCoDisplay
  Unload frmSearch
  Unload frmFind
'
' remove shadow object
'
  Set myShadow = Nothing
'
' reset form sizing subclasser
'
  If CBool(m_VChWnd) Then Call UnhookWin(Me.hWnd, m_VChWnd)
End Sub

'*******************************************************************************
' Subroutine Name   : CheckTextEntry
' Purpose           : Process TEXT ENTRY button
'*******************************************************************************
Public Sub checkTextEntry(ByVal TxtEnabled As Boolean)
  If Me.chkShift.Enabled And Not TxtEnabled Then
    Me.chkShift.Value = vbUnchecked
  End If
  Me.chkShift.Enabled = TxtEnabled
  TextEntry = TxtEnabled                            'set TRUE if we want text entry
  Me.cmdUsrA(0).Enabled = AllowSpace And TxtEnabled 'enable space if text entry and process allows it
  VarLbl = False                                    'turn off variable labeling option (if on)
  Upcase = False
  HaveTxt = False                                   'no text data typed
  
  If Not TextEntry Then                             'if not in text entry mode
    Call RedoAlphaPad                               'set user-defined keys
  Else
    Call ResetAlphaPad                              'if in text entry mode, set alpha keys
  End If
  Call UpdateStatus                                 'reflect text entry mode in status
End Sub

'*******************************************************************************
' Subroutine Name   : chk2nd_Click
' Purpose           : Toggle 2nd key
'*******************************************************************************
Public Sub chk2nd_Click()
  Static IgnoreIt As Boolean   'used by 2nd key
  
  If IgnoreIt Then Exit Sub
  If Not NoQuery Then PlayClick         'play resource click
  If RunMode Then
    IgnoreIt = True
    If Key2nd Then
      Me.chk2nd.Value = vbChecked       'ensure key is down
    Else
      Me.chk2nd.Value = vbUnchecked     'ensure key is up
    End If
    DoEvents                            'let screen catch up
    IgnoreIt = False
    Exit Sub
  Else
    If Query_Pressed And Not NoQuery Then
      IgnoreIt = True
      If Key2nd Then
        Me.chk2nd.Value = vbChecked       'ensure key is down
      Else
        Me.chk2nd.Value = vbUnchecked     'ensure key is up
      End If
      IgnoreIt = False
      Call Query(-102)
      Exit Sub
    End If
  End If
  Key2nd = Me.chk2nd.Value = vbChecked    'set TRUE if 2nd key pressed
  LastTypedInstr = 128
  Call ResetAccumulator                   'reset accumulator data
  Call ShowKeypad                         'display new keyfaces
  If CBool(CharCount) And CBool(CharLimit) Then Exit Sub
  SetTip vbNullString                     'nullify status line if we are not keying something in
End Sub

'*******************************************************************************
' Subroutine Name   : cmdKeyPad_Click
' Purpose           : Keypad click. Note that codes are +128 (2nds are +256).
'*******************************************************************************
Private Sub cmdKeyPad_Click(Index As Integer)
  PlayClick                   'play resource click
  If RunMode Then
    If Index = 91 Then        'if R/S key
      RS_Pressed = True       'indicate R/S pressed
    End If
  Else
    If Query_Pressed Then     'if we are wanting help with this key
      Call Query(Index)
    Else
      Call MainKeyPad(Index)  'else process key
    End If
  End If
  Me.txtFocus.SetFocus        'set "dafault" focus
End Sub

'*******************************************************************************
' Subroutine Name   : Form_KeyDown
' Purpose           : Set 2nd Key if SHIFT key pressed
'*******************************************************************************
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim Idx As Integer, Idy As Long
  Dim S As String, Path As String
  
'  Debug.Print KeyCode
'  Debug.Print Shift
  If KeyCode = 116 And Shift = 0 Then                     'F5
    If LrnMode Or RunMode Then
      CmdNotActive
      Exit Sub
    End If
    Preprocessd = False
    Call Preprocess
    Exit Sub
  End If
  
  If KeyCode = 16 And Shift = vbShiftMask Then            'SHIFT
    KeyShf = True
    With Me.chkShift
      If .Enabled Then
        If .Value = vbUnchecked Then .Value = vbChecked
      End If
    End With
  ElseIf KeyCode = 17 And Shift = vbCtrlMask Then         'CTRL
    If Me.chk2nd.Value = vbUnchecked Then
      NoQuery = True    'prevent some unecessary processing
      Me.chk2nd.Value = vbChecked
      NoQuery = False   're-enable processing
    End If
    Exit Sub
  ElseIf KeyCode = 86 And Shift = vbCtrlMask Then         'CTRL-V
    If RunMode Or LrnMode Or CBool(MRunMode) Then Exit Sub
    S = Clipboard.GetText(vbCFText)                       'grab data from clipboard
    If CBool(Len(S)) Then                                 'if data exists...
      Idx = InStr(1, S, vbCr)                             'check for CR
      If CBool(Idx) Then S = Left$(S, Idx - 1)            'strip
      If CBool(Len(S)) Then                               'still data?
        DspTxt = S                                        'set to DspTxt
        DisplayText = Not IsNumeric(S)                    'see if numeric
        If Not DisplayText Then DisplayReg = Val(S)       'if not, set to DisplayReg
        Call DisplayLine                                  'but display the line
      End If
    End If
  ElseIf Shift = 0 Then                                   'do following only if no SHIFT or CTRL or ALT
    With Me.lstDisplay
      Select Case KeyCode
        Case 33 'PgUp
          Idx = .ListIndex - DisplayHeight
          Idy = Idx - DisplayHeight \ 2
          If Idx < 0 Then Idx = 0
          If Idy < 0 Then Idy = 0
          Call SelectOnly(Idx)              'select only this item
          .TopIndex = Idy
          
        Case 34 'PgDn
          Idx = .ListIndex + DisplayHeight
          If Idx >= .ListCount Then Idx = .ListCount - 1
          Idy = Idx - DisplayHeight \ 2
          If Idy < 0 Then Idy = 0
          Call SelectOnly(Idx)              'select only this item
          .TopIndex = Idy
        
        Case 35 'End
          Idx = .ListCount - 1
          Idy = Idx - DisplayHeight \ 2
          If Idy < 0 Then Idy = 0
          Call SelectOnly(Idx)              'select only this item
          .TopIndex = Idy
        
        Case 36 'Home
          Call SelectOnly(0)                'select only this item
          .TopIndex = 0
        
        Case 37 'Left Arrow
          Me.cmdBackspace.Value = True
          
        Case 38 'Up Arrow
          Idx = .ListIndex - 1
          If Idx < 0 Then Idx = 0
          Idy = Idx - DisplayHeight \ 2
          If Idy < 0 Then Idy = 0
          Call SelectOnly(Idx)              'select only this item
          .TopIndex = Idy
          
        Case 39, 40 'Down Arrow
          Idx = .ListIndex + 1
          If Idx = .ListCount Then Idx = Idx - 1
          Idy = Idx - DisplayHeight \ 2
          If Idy < 0 Then Idy = 0
          Call SelectOnly(Idx)              'select only this item
          .TopIndex = Idy
        
        Case 45 'Insert
          If LrnMode Then
            INSmode = Not INSmode
            Call UpdateStatus
          End If
        
        Case 46 'Delete
          If LrnMode Then
            Call DeleteInstruction
          Else
            Call Del_Support
          End If
      End Select
    End With
  End If

End Sub

'*******************************************************************************
' Subroutine Name   : Form_KeyPress
' Purpose           : Handle keyboard keys on keypad, if text entry enabled
'*******************************************************************************
Private Sub Form_KeyPress(KeyAscii As Integer)
  Dim C As String, S As String, cc As String
  Dim i As Integer
  Dim Hld2nd As Boolean

  If RunMode Then
    If KeyAscii = 27 Then     'if ESC pressed in Run Mode, allow it to act as R/S key
      RS_Pressed = True       'indicate R/S pressed
    End If
    Exit Sub
  End If
'
' check for escape key
'
  PlayClick                 'play resource click
  If KeyAscii = 27 Then     'ESC pressed and not run mode?
    If Query_Pressed Then   'Query active?
      Query_Pressed = False 'turn it off if so
      Me.MousePointer = 0   'reset cursor
      DoEvents              'and let screen refresh
    ElseIf LrnMode Then
      SetTip vbNullString
    ElseIf DspLocked Then
      DspLocked = False     'turn off program display (we will emulate keyboard CE function)
      Call DspBackground
      Call CLR_Support
    Else
      Call CE_Support       'and reset things CE normally does
      PendIdx = 0           'reset pending operations
      Call ResetPndAll
    End If
    Exit Sub                'exit anyway on ESC
  End If
'
' Tie ENTER to '=' key
'
  If KeyAscii = 13 Then
    If TypeMatic And BaseType <> TypHex Then            'can use typmatic keyboard?
      If CBool(Len(TypeMatTxt)) Then                    'yes, anything in buffer?
        i = ImpCmd(TypeMatTxt)                          'yes, legal command?
        If i = 128 Then
          CmdNotActive                                  'no
        Else
          Hld2nd = Key2nd                               'yes, so save 2nd key state
          If i > 255 Then
            i = i - 256                                 'set to 2nd key offset
            Key2nd = True                               'force 2nd key
            Me.cmdKeyPad(i).Value = True                'activate its keypad button
          Else
            i = i - 128                                 'not 2nd key offset
            Key2nd = False                              'force 2nd key off
            Me.cmdKeyPad(i).Value = True                'activate its keypad button
          End If
          Key2nd = Hld2nd                               'reset 2nd key actual state
          TypeMatTxt = vbNullString                     'reset TypeMatic buffer
          SetTip vbNullString
          Exit Sub
        End If
      End If
    Else
      TypeMatTxt = vbNullString
    End If
  End If
  
  If KeyAscii = 13 And Not Key2nd Then
    If BaseType <> TypHex And PendIdx = 0 And CharCount = 0 And (CBool(InstrCnt) Or CBool(ActivePgm)) Then
      Me.cmdKeyPad(91).Value = True     'treat ENTER like R/S
    Else
      Me.cmdKeyPad(95).Value = True     'else like '='
    End If
    Exit Sub
  End If
'
' tie backspace to Backsp key
'
  If KeyAscii = 8 Then
    Me.cmdBackspace.Value = True        'backspace text entry
    Exit Sub
  End If
'
' support other keyboard ties...
'
  C = Chr$(KeyAscii)                                'get user key
  
  If TextEntry Then                                 'if TEXT ENTRY MODE
    If KeyAscii > 31 Then                           'Allow all non-Control Characters
      If Len(DspTxt) < CharLimit Then               'less than limit?
        If LrnMode Then
          Select Case C
            Case "0" To "9"                         'this will be handled in main processor
            Case Else
              LastTypedInstr = KeyAscii             'set key command
              AddInstruction LastTypedInstr         'add to learn mode\
          End Select
        Else
          Select Case C
            Case "0" To "9"                         'this will be handled in main processor
            Case Else
              DspTxt = DspTxt & Chr$(KeyAscii)      'else use key caption for text character
              CharCount = Len(DspTxt)               'establish character count
              If Not RunMode Then                   'show updates if not in RUN mode
                Me.lstDisplay.List(Me.lstDisplay.ListIndex) = String$(DisplayWidth - Len(DspTxt), 32) & DspTxt
              End If
              DisplayHasText = True
          End Select
        End If
      Else
        CmdNotActive                                'issue beep if at text entry limit
      End If
    End If
    If KeyAscii < 48 Or KeyAscii > 57 Then Exit Sub 'exit regardless, if Text Entry mode
  ElseIf KeyAscii = 34 And Not RunMode Then         'allow ["] to init text entry
    LastTypedInstr = iTXT                           'force TXT mode
    If LrnMode Then
      Call LrnKeypad                                'for learn mode
    Else
      Call ActiveKeypad                             'or active keyboard
    End If
    Exit Sub
  End If
  
  Hld2nd = Key2nd                                   'hold current state of 2nd key (we may temp change it)
  cc = UCase$(C)
  Select Case cc                                    'check character typed in non-text-entry mode
    Case "A" To "Z", "<", ">", "!", " "       'allowed for typmatic
    Case "/"
      If TypeMatic And CBool(Len(TypeMatTxt)) Then
        If UCase$(TypeMatTxt) = "R" Then                  'Possible R/S?
          TypeMatTxt = TypeMatTxt & C                     'yes, assume so
          SetTip "TypeMatic: " & TypeMatTxt               'display TypeMatic status
          Exit Sub
        End If
      End If
    Case Else
      If TypeMatic And CBool(Len(TypeMatTxt)) Then
        Select Case cc
          Case "="
            Select Case UCase$(Right$(TypeMatTxt, 1))
              Case "X", "<", ">", "=", "!"
                TypeMatTxt = TypeMatTxt & C               'bind typematic with '=' key
                SetTip "TypeMatic: " & TypeMatTxt         'display TypeMatic status
                Exit Sub
            End Select
          Case ";"
            If StrComp(TypeMatTxt, "PRINT", vbTextCompare) = 0 Then
              TypeMatTxt = TypeMatTxt & cc
              cc = vbNullString
            End If
        End Select
        i = ImpCmd(TypeMatTxt)
        TypeMatTxt = vbNullString
        SetTip vbNullString
        If i = 128 Then
          CmdNotActive
        Else
          Hld2nd = Key2nd
          If i > 255 Then
            i = i - 256
            Key2nd = True                               'force 2nd key
            Me.cmdKeyPad(i).Value = True                'activate its keypad button
          Else
            i = i - 128
            Key2nd = False                              'force 2nd key off
            Me.cmdKeyPad(i).Value = True                'activate its keypad button
          End If
          Key2nd = Hld2nd                               'reset 2nd key state
        End If
      End If
  End Select
  
  If cc = vbNullString Then Exit Sub
  Select Case cc                                    'check character typed in non-text-entry mode
    Case ","                                        'comma?
      Key2nd = True                                 'force 2nd key
      Me.cmdKeyPad(93).Value = True                 'activate its keypad button
      Key2nd = Hld2nd                               'reset 2nd key state
      Exit Sub
    Case ";"                                        'Semicolon?
      Key2nd = False                                'force not 2nd key
      Me.cmdKeyPad(40).Value = True                 'activate its keypad button
      Key2nd = Hld2nd                               'reset 2nd key state
      Exit Sub
    Case ":"                                        'colon?
      Key2nd = True                                 'force 2nd key
      Me.cmdKeyPad(40).Value = True                 'activate its keypad button
      Key2nd = Hld2nd                               'reset 2nd key state
      Exit Sub
    Case "0" To "9"
    Case Else
      If TypeMatic And BaseType <> TypHex Then      'TypeMatic keyboard and not Hex mode?
        If Not CBool(Len(TypeMatTxt)) And CBool(InStr(1, ".+-*=/()[]{}", cc)) Then
          TypeMatTxt = vbNullString
        ElseIf CBool(Len(TypeMatTxt)) And CBool(InStr(1, ".+-*=()[]{}", cc)) Then
          TypeMatTxt = vbNullString
        ElseIf C = " " Then                         'force TypeMatic processing?
          If CBool(Len(TypeMatTxt)) Then            'yes, anything to process?
            i = ImpCmd(TypeMatTxt)                  'yes, is legal command?
            TypeMatTxt = vbNullString
            SetTip vbNullString
            If i = 128 Then
              CmdNotActive                          'no
            Else
              Hld2nd = Key2nd                       'else save 2nd key state
              If i > 255 Then
                i = i - 256                         'set 2nd key offset
                Key2nd = True                       'force 2nd key
                Me.cmdKeyPad(i).Value = True        'activate its keypad button
              Else
                i = i - 128                         'set not 2nd key offset
                Key2nd = False                      'force 2nd key off
                Me.cmdKeyPad(i).Value = True        'activate its keypad button
              End If
              Key2nd = Hld2nd                       'reset 2nd key actual state
            End If
          Else
            CmdNotActive                            'nothing, so beep
            SetTip vbNullString
          End If
        Else
          TypeMatTxt = TypeMatTxt & C               'add key to TypeMatic buffer
          SetTip "TypeMatic: " & TypeMatTxt         'display TypeMatic status
          Exit Sub
        End If
      End If
  End Select
'
' do numbers check if 2nd key not pressed
'
  If Not Key2nd Then                    'if 2nd key not down
    S = "1234567890.+-*/=()[]{}"        'allowable keys
    If BaseType = TypHex And Not RunMode And Not CBool(MRunMode) Then
      S = "ABCDEF1234567890+-*/=()[]{}" 'allowable keys for HEX
    End If
    
    i = 0
    If CBool(InStr(1, S, cc)) Then      'found a key?
      Select Case cc                    'map key to keyboard
        Case "0"
          i = 92                        'set key code for 0
        Case "1"
          i = 79
        Case "2"
          i = 80
        Case "3"
          i = 81
        Case "4"
          i = 66
        Case "5"
          i = 67
        Case "6"
          i = 68
        Case "7"
          i = 53
        Case "8"
          i = 54
        Case "9"
          i = 55
        Case "."
          i = 93
        Case "+"
          i = 82
        Case "-"
          i = 69
        Case "*"
          i = 56
        Case "/"
          i = 43
        Case "="
          i = 95
        Case "("
          i = 41
        Case ")"
          i = 42
        
        Case "A"
          i = 26
        Case "B"
          i = 39
        Case "C"
          i = 52
        Case "D"
          i = 65
        Case "E"
          i = 78
        Case "F"
          i = 91
        Case "["
          i = 110
        Case "]"
          i = 111
        Case "{"
          i = 62
        Case "}"
          i = 63
      End Select
      If CBool(i) Then
        Me.cmdKeyPad(i).Value = True      'invoke key
      End If
      Exit Sub
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Form_KeyUp
' Purpose           : Reset 2nd Key if CTRL key pressed
'*******************************************************************************
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 16 And Shift = 0 Then      'SHIFT key
    KeyShf = False
    With Me.chkShift
      If .Enabled Then
        If .Value = vbChecked Then .Value = vbUnchecked
      End If
    End With
  ElseIf KeyCode = 17 And Shift = 0 Then  'CNTRL key
    If Me.chk2nd.Value = vbChecked Then
      NoQuery = True
      Me.chk2nd.Value = vbUnchecked
      NoQuery = False
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : imgLocked_Click
' Purpose           : Show help for lock/unlock flag
'*******************************************************************************
Private Sub imgLocked_Click(Index As Integer)
  If RunMode Then Exit Sub  'ignore if run Mode
  If Query_Pressed Then
    Call Query(-117)        'display help for key if Query active
    Exit Sub
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : lblCode_Click
' Purpose           : Show help for CODE indicator
'*******************************************************************************
Private Sub lblCode_Click()
  If Query_Pressed Then
    Call Query(-5)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : lblDRGM_Click
' Purpose           : Show help for Angle Mode indicator
'*******************************************************************************
Private Sub lblDRGM_Click()
  If Query_Pressed Then
    Call Query(-11)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : lblEE_Click
' Purpose           : Show help for EE indicator
'*******************************************************************************
Private Sub lblEE_Click()
  If Query_Pressed Then
    Call Query(-10)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : lblHDOB_Click
' Purpose           : Show help for Number Base indicator
'*******************************************************************************
Private Sub lblHDOB_Click()
  If Query_Pressed Then
    Call Query(-12)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : lblINS_Click
' Purpose           : Show help for INS indicator
'*******************************************************************************
Private Sub lblINS_Click()
  If Query_Pressed Then
    Call Query(-8)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : lblInstr_Click
' Purpose           : Show help for INSTR. indicator
'*******************************************************************************
Private Sub lblInstr_Click()
  If Query_Pressed Then
    Call Query(-6)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : lblLoc_Click
' Purpose           : Show help for STEP indicator
'*******************************************************************************
Private Sub lblLoc_Click()
  If Query_Pressed Then
    Call Query(-4)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : lblLRN_Click
' Purpose           : Show help for LRN indicator
'*******************************************************************************
Private Sub lblLRN_Click()
  If Query_Pressed Then
    Call Query(-9)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : lbltxt_Click
' Purpose           : Show help for TXT indicator
'*******************************************************************************
Private Sub lbltxt_Click()
  If Query_Pressed Then
    Call Query(-7)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : lstDisplay_Click
' Purpose           : Reset instruction pointer when in LRN Mode
'*******************************************************************************
Private Sub lstDisplay_Click()
  Dim S As String, T As String
  Dim i As Integer, j As Integer
  
  If RunMode Then Exit Sub                          'ignore if run Mode
  If Not Selecting Then
    PlayClick                                       'play resource click
  End If
  If LrnMode Then
    If IgnoreClick Then Exit Sub                    'ignore select when resetting pointer in pgm
    If KeyShf Or Key2nd Then Exit Sub               'shift or control pressed
    SelectOnly Me.lstDisplay.ListIndex              'select instruction
    If frmCDLoaded Then
      Call frmCoDisplay.RepointIndex
    End If
    StopMode = False                                'changed instruction
    AllowSpace = False                              'do not allow typing of a space
    Call Me.checkTextEntry(False)                   'set up keyboard
    Call UpdateStatus
    Exit Sub
  ElseIf DspLocked Then
    S = Me.lstDisplay.List(Me.lstDisplay.ListIndex) 'get displayed line
    If Not CBool(Len(S)) Then
      SetTip vbNullString
      Exit Sub
    End If
    j = InStr(1, S, " @ ")                          'see if it contains possible location
    If CBool(j) Then
      i = InStr(1, S, GetInstrStr(iSbr))            'find additional checks
      If i = 0 Then i = InStr(1, S, GetInstrStr(iLbl))
      If i = 0 Then i = InStr(1, S, GetInstrStr(iUkey))
      If i = 0 Then i = InStr(1, S, GetInstrStr(iStruct))
      If i = 0 Then i = InStr(1, S, GetInstrStr(iEnum))
      If i = 0 Then i = InStr(1, S, GetInstrStr(iConst))
      If i = 0 Then i = InStr(1, S, GetInstrStr(iNvar))
      If i = 0 Then i = InStr(1, S, GetInstrStr(iTvar))
      If i = 0 Then i = InStr(1, S, GetInstrStr(iIvar))
      If i = 0 Then i = InStr(1, S, GetInstrStr(iCvar))
      If CBool(i) Then
        S = Trim$(Mid$(S, j + 3))                   'grab possible address
        On Error Resume Next
        i = CInt(Fix(Val(S)))                       'try to derive the address
        If Err.Number = 0 Then                      'if all is well
          If i < InstrCnt Then                      'if within instruction range
            InstrPtr = i                            'set instruction pointer
            Call UpdateStatus                       'and update status
          End If
        End If
        On Error GoTo 0
      End If
    Else                                            'go to LRN mode at selected point
      Select Case LRNstyle
        Case 0                                      'raw format
          InstrPtr = Me.lstDisplay.ListIndex        'listindex is also program instruction pointer
          Call UpdateStatus
        Case Else
          If Preprocessd Then
            If LRNstyle = 3 Then                    'grab index address from appropriate map array
              InstrPtr = InstMap3(Me.lstDisplay.ListIndex)  'debug-style mapping
            Else
              InstrPtr = InstMap(Me.lstDisplay.ListIndex)   'standard formatted style mapping
            End If
            Call UpdateStatus                       'echk instruction pointer address to status bar
          End If
      End Select
    End If
  End If
  SetTip vbNullString
End Sub

'*******************************************************************************
' Subroutine Name   : lstDisplay_DblClick
' Purpose           : If DspLocked flag set, go to selected line
'*******************************************************************************
Private Sub lstDisplay_DblClick()
  Dim HldKey As Boolean
  Dim S As String, Path As String
  Dim Idx As Integer
  
  If LrnMode Then Exit Sub
  If DspLocked Then     'if DspLocked, then open LRN Mode (we will be at selected line)
    HldKey = Key2nd     'save state of actual key state
    Key2nd = False      'assume off
    Call MainKeyPad(1)  'select LRN
    Key2nd = HldKey     'reset to actual state
    Exit Sub
  End If
'
' see if we have a list to process
'
  If Not StoreList And Not ModuleList Then Exit Sub
'
' see if user double-clicked a directory listing entry
'
  With Me.lstDisplay
    S = UCase$(Trim$(.List(.ListIndex)))                  'get data on line
    
    If Left$(S, 3) = "PGM" And Mid$(S, 6, 1) = ":" Then   'possible listing for module
      Idx = CInt(Mid$(S, 4, 2))                           'grab program number
      If Idx > 0 And Idx <= ModCnt Then                   'in range?
        Call Clear_Screen
        Call DisplayLine
        DisplayReg = CDbl(Idx)                            'set to DisplayReg
        PndIdx = 1                                        'force [Pgm] [nn]
        PndStk(1) = iPgm
        Call CheckPnd(0)
        DisplayReg = 0                                      'reset display register
      End If
      Exit Sub                                            'all done
    End If
    
    If Right$(S, 1) = ">" Then                            'strip possible comments
      Idx = InStr(1, S, "<")
      If CBool(Idx) Then S = RTrim$(Left$(S, Idx - 1))
    End If
    
    Select Case Right$(S, 4)                              'check 4 right-most characters
      Case ".TXT"
        For Idx = .ListIndex - 1 To 0 Step -1             'find folder, if so
          Path = .List(Idx)
          If Left$(Path, 1) <> " " Then Exit For          'if no space at head, then folder
        Next Idx
        If Left$(Path, 1) = " " Then Exit Sub             'did not find anything
        Path = StorePath & "\" & RTrim$(Path) & "\" & S   'build path to text file
        Call ShellPath(Me.hWnd, "open", Path)             'opon it with user's default text editor
      
      Case ".PGM"
        S = Mid$(S, 4)                                  'strip "PGM" from left
        S = Left$(S, Len(S) - 4)                        'strip ".pgm" from right
        DisplayReg = CDbl(S)                            'set pgm # to main register
        PndIdx = 1                                      'force LOAD onto pending stack
        PndStk(1) = iLoad
        Call CheckPnd(0)                                'load program
      
      Case ".MDL"
        .Clear
        .AddItem " "
        .ListIndex = 0
        S = Mid$(S, 4)                                  'strip "MDL" from left
        S = Left$(S, Len(S) - 4)                        'strip ".mdl" from right
        DisplayReg = CDbl(S)                            'set module # to main register
        PndIdx = 2                                      'force LOAD onto pending stack
        PndStk(1) = iMDL                                'set MDL Load on pending stack
        PndStk(2) = iLoad
        Call CheckPnd(0)                                'load program
    End Select
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : lstDisplay_GotFocus
' Purpose           : Prevent double scrolling by keeping focus off of the listbox
'*******************************************************************************
Private Sub lstDisplay_GotFocus()
  Me.txtFocus.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : lstDisplay_MouseMove
' Purpose           : Support displaying tooltips for selected instructions
'*******************************************************************************
Public Sub lstDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim S As String
  Dim i As Long
  Dim HldKey2nd As Boolean
  
  If Not LrnMode Then                     'if not learn mode
    If CBool(CharLimit) Then Exit Sub     'and not entering data
    S = Me.lstDisplay.List(ListItemByCoordinate(Me.lstDisplay, X, Y))
    Me.lblWidth.Caption = S
    If Me.lblWidth.Width <= Me.lstDisplay.Width Then 'if text width <= displayable width
      If DspLocked Then                   'if display locked
        S = "Click row sets Pgm Step. Double-Click opens LRN mode at that step"
      Else
        If StoreList Then                 'if display list
          S = "Double-click file to open/load it, DEL to delete with confirmation"
        ElseIf ModuleList Then            'if displaying module program list
          S = "Double-click Pgm line activates that program"
        Else
          S = vbNullString                'not locked or list, so show no tips
        End If
      End If
    End If
    SetTip S
  Else
    With Me.lstDisplay
      S = .List(ListItemByCoordinate(Me.lstDisplay, X, Y))
      'Debug.Print S
      If Len(S) < 10 Then
        S = vbNullString
      Else
        i = CInt(Mid$(S, 7, 3))
        Select Case i
          'convert special end braces to normal
          Case iBCBrace, iSCBrace, iICBrace, iDCBrace, iWCBrace, _
               iFCBrace, iCCBrace, iDWBrace, iDUBrace, iEIBrace, _
               iENBrace, iSTBrace, iSIBrace, iCNBrace
            i = iRCbrace
          'conver special parens to normal
          Case iUparen, iIparen, iWparen, iFparen, iCparen, iEparen, iSparen, iDWparen
            i = iRparen
          'convert special Else for Case back to normal
          Case iCaseElse
            i = iElse
          'convert ';' in FOR statements
          Case iSemiColon
            i = iSemiC
        End Select
        HldKey2nd = Key2nd
        Select Case i
          Case Is < 10
            S = "Digit '" & Chr$(i + 48) & "'"
          Case Is < 128
            S = "Text character '" & Chr$(i) & "'"
          Case Is < 256
            Key2nd = False
            i = i - 128
            Call TipCheck(i, True)
            S = TipText
          Case Is > 900
            S = "User-Defined key '" & Chr$(i - 836) & "'"
          Case Else
            Key2nd = True
            i = i - 256
            Call TipCheck(i, True)
            S = TipText
        End Select
        Key2nd = HldKey2nd
      End If
      If .ToolTipText <> S Then
        .ToolTipText = S
      End If
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileCopySel_Click
' Purpose           : Copy selected items to the clipboard
'*******************************************************************************
Private Sub mnuFileCopySel_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-18)
      Exit Sub
    End If
  End If
  
  Dim Mylist() As Long
  Dim Idx As Integer
  Dim S As String
  
  S = vbNullString                        'init destination buffer
  If Me.rtbInfo.Visible Then
    With Clipboard
      .Clear
      .SetText Me.rtbInfo.SelRTF, vbCFRTF
      .SetText Me.rtbInfo.SelText, vbCFText
    End With
  Else
    Mylist = GetSelListBox(Me.lstDisplay)   'check for selections in listbox
    With Me.lstDisplay
      If .SelCount = 1 Then
        S = Trim$(.List(Mylist(0)))
      Else
        For Idx = 0 To UBound(Mylist)         'process each selected item in the listbox
          S = S & .List(Mylist(Idx)) & vbCrLf 'add selected itm
        Next Idx
      End If
    End With

    With Clipboard
      .Clear                              'clear clipboard
      .SetText S, vbCFText                'save text to it
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileImport_Click
' Purpose           : Import ASCII file from selection
'*******************************************************************************
Private Sub mnuFileImport_Click()
  Dim S As String, Path As String
  Dim Idx As Integer
  
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-23)
      Exit Sub
    End If
  End If
  
  With Me.lstDisplay
    S = UCase$(Trim$(.List(.ListIndex)))
    If Right$(S, 1) = ">" Then                        'strip possible comments
      Idx = InStr(1, S, "<")
      If CBool(Idx) Then S = RTrim$(Left$(S, Idx - 1))
    End If
    If Right$(S, 4) <> ".TXT" Then
      ForcError "Selected entry is not an ASCII text file"
      Exit Sub
    End If
    For Idx = .ListIndex - 1 To 0 Step -1           'find folder, if so
      Path = .List(Idx)
      If Left$(Path, 1) <> " " Then Exit For        'if no space at head, then folder
    Next Idx
  End With
  If Left$(Path, 1) = " " Then
    ForcError "Invalid selection. Be sure to do a directory listing"
    Exit Sub                                        'did not find anything
  End If
  Path = StorePath & "\" & RTrim$(Path) & "\" & S   'build path to text file
  If Fso.FileExists(Path) Then
    Call ImportFile(Path, False, False)
  Else
    ForcError "Selected file cannot be location from the Data Storage Location"
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileImportClipBrd_Click
' Purpose           : Import ASCII file from clipboard
'*******************************************************************************
Private Sub mnuFileImportClipBrd_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-24)
      Exit Sub
    End If
  End If
  
  Call ImportFile(vbNullString, True, False)
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileImportSegment_Click
' Purpose           : Import ASCII file SEGMENT from clipboard
'*******************************************************************************
Private Sub mnuFileImportSegment_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-25)
      Exit Sub
    End If
  End If
  
  Call ImportFile(vbNullString, True, True)
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileListDir_Click
' Purpose           : List directory of storage folder
'*******************************************************************************
Private Sub mnuFileListDir_Click()
  Dim Fld As Folder
  Dim Fil As File
  Dim Fn As Integer
  Dim IntV As Integer, Idx As Integer
  Dim Cmt As String, Lbl As String * DisplayWidth
  Dim ts As TextStream
  
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-22)
      Exit Sub
    End If
  End If
  
  If Not CBool(Len(StorePath)) Then
    ForcError "No storage path yet defined"
    Exit Sub
  End If
'
' disable any locking
'
  If DspLocked Then
    DspLocked = False
    Call DspBackground
  End If
'
' init the screen
'
  Call Clear_Screen
'
' list each folder and file
'
  With Me.lstDisplay
    For Each Fld In Fso.GetFolder(StorePath).SubFolders
      .AddItem Fld.Name                       'display a folder name
      For Each Fil In Fld.Files
        Cmt = vbNullString                    'init comment
        Select Case Right$(UCase$(Fil.Name), 4)
          Case ".PGM"
            On Error Resume Next
            Fn = FreeFile(0)                    'we will try opening the pgm file
            Open AddSlash(StorePath) & AddSlash(Fld.Name) & Fil.Name For Binary Access Read As #Fn
            If Not CBool(Err.Number) Then
              Get #Fn, , IntV                   'get first instruction
                Select Case IntV
                  Case iRem, iRem2              'remark?
                    For Idx = 1 To DisplayWidth 'gather text
                      Get #Fn, , IntV
                      Select Case IntV
                        Case Is < 10            'digit 0-9
                          Exit For
                        Case Is < 128           'ascii
                          Cmt = Cmt & Chr$(IntV)
                        Case Else               'instructions
                          Exit For
                      End Select
                    Next Idx
                    Cmt = Trim$(Cmt)
                    If CBool(Len(Cmt)) Then Cmt = " <" & Cmt & ">"
                End Select
            End If
            Close #Fn
            On Error GoTo 0
          Case ".MDL"
            On Error Resume Next
            Fn = FreeFile(0)                    'we will try opening the pgm file
            Open AddSlash(StorePath) & AddSlash(Fld.Name) & Fil.Name For Binary Access Read As #Fn
            If Not CBool(Err.Number) Then
              For Idx = 1 To 5
                Get #Fn, , IntV                 'get some prelim data
              Next Idx
              Get #Fn, , Lbl
              If CBool(Len(Lbl)) Then Cmt = " <" & Trim$(Lbl) & ">"
            End If
            Close #Fn
            On Error GoTo 0
          Case ".TXT"
          Set ts = Fso.OpenTextFile(AddSlash(StorePath) & AddSlash(Fld.Name) & Fil.Name, ForReading, False)
          Cmt = ts.ReadLine
          ts.Close
          If Left$(Cmt, 1) = "'" Then
            Cmt = " <" & Mid$(Cmt, 2) & ">"
          ElseIf StrComp(Left$(Cmt, 4), "REM ", vbTextCompare) = 0 Then
            Cmt = " <" & Mid$(Cmt, 5) & ">"
          Else
            Cmt = vbNullString
          End If
        End Select
        .AddItem "  " & Fil.Name & Cmt        'list its file contents
      Next Fil
    Next Fld
  End With
  
  SelectOnly 0                                'select the top-most line
  CharLimit = 0                               'force display of data
  DisplayReg = 0#
  DisplayText = False
  StoreList = True
  Call DisplayLine                            'display data line
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileNotepad_Click
' Purpose           : Launch the Notepad text editor application
'*******************************************************************************
Private Sub mnuFileNotepad_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-20)
      Exit Sub
    End If
  End If
  
  Shell NotePadPath, vbNormalFocus
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileTglCoDisplay_Click
' Purpose           : Toggle Co-Display of formatted Src in Learn Mode
'*******************************************************************************
Private Sub mnuFileTglCoDisplay_Click()
  Dim Bol As Boolean
  
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-41)
      Exit Sub
    End If
  End If
  
  With Me.mnuFileTglCoDisplay
    Bol = Not .Checked
    .Checked = Bol
    SaveSetting App.Title, "Settings", "CoDisplay", CStr(Bol)
  End With
  
  If Bol Then
    If LrnMode Then
      Call InitCoDisplay
    End If
  Else
    If frmCDLoaded Then
      Unload frmCoDisplay
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnufileTypeMatic_Click
' Purpose           : Toggle TypeMatic Keyboard
'*******************************************************************************
Private Sub mnufileTypeMatic_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-30)
      Exit Sub
    End If
  End If
  
  With Me.mnufileTypeMatic
    TypeMatic = Not .Checked
    .Checked = TypeMatic
    SaveSetting App.Title, "Settings", "TypeMatic", CStr(TypeMatic)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileWordpad_Click
' Purpose           : Launch the Wordpad text editor application
'*******************************************************************************
Private Sub mnuFileWordpad_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-21)
      Exit Sub
    End If
  End If
  
  Shell WordPadPath, vbNormalFocus
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFilePaste_Click
' Purpose           : Paste the clipboard to the display
'*******************************************************************************
Private Sub mnuFilePaste_Click()
  Dim S As String
  Dim Idx As Long
  
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-19)
      Exit Sub
    End If
  End If
  
  If RunMode Or CBool(MRunMode) Then Exit Sub
  
  If LrnMode Then
    Call PasteInstruction                               'perform instruction paste
    Exit Sub
  End If
  
  S = Clipboard.GetText(vbCFText)                       'grab data from clipboard
  If CBool(Len(S)) Then                                 'if data exists...
    Idx = InStr(1, S, vbCr)                             'check for CR
    If CBool(Idx) Then S = Left$(S, Idx - 1)            'strip
    If CBool(Len(S)) Then                               'still data?
      DspTxt = S                                        'set to DspTxt
      DisplayText = Not IsNumeric(S)                    'see if numeric
      If Not DisplayText Then DisplayReg = Val(S)       'if not, set to DisplayReg
      Call DisplayLine                                  'but display the line
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFilePrcmp_Click
' Purpose           : Flag auto Preprocess
'*******************************************************************************
Private Sub mnuFilePrcmp_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-26)
      Exit Sub
    End If
  End If
  
  AutoPprc = Not Me.mnuFilePrcmp.Checked
  Me.mnuFilePrcmp.Checked = AutoPprc
  Call SaveSetting(App.Title, "Settings", "AutoPprc", CStr(AutoPprc))
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileReloadMDL_Click
' Purpose           : Change setting for auto-reload of Module
'*******************************************************************************
Private Sub mnuFileReloadMDL_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-27)
      Exit Sub
    End If
  End If
  
  Me.mnuFileReloadMDL.Checked = Not Me.mnuFileReloadMDL.Checked
  SaveSetting App.Title, "Settings", "ReloadMDL", CStr(Me.mnuFileReloadMDL.Checked)
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileReloadPgm_Click
' Purpose           : Change setting for auto-reload of Pgm
'*******************************************************************************
Private Sub mnuFileReloadPgm_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-28)
      Exit Sub
    End If
  End If
  
  Me.mnuFileReloadPgm.Checked = Not Me.mnuFileReloadPgm.Checked
  SaveSetting App.Title, "Settings", "ReloadPgm", CStr(Me.mnuFileReloadPgm.Checked)
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFile_Click
' Purpose           : Set up Import options When File menu activated
'*******************************************************************************
Private Sub mnuFile_Click()
  Dim S As String, Path As String
  Dim Idx As Integer
  
  Path = " "                                          'init to failure
  With Me.lstDisplay
    S = UCase$(Trim$(.List(.ListIndex)))              'grav select line
    If Right$(S, 1) = ">" Then                        'strip possible comments
      Idx = InStr(1, S, "<")
      If CBool(Idx) Then S = RTrim$(Left$(S, Idx - 1))
    End If
    If Right$(S, 4) = ".TXT" Then                     'text file?
      For Idx = .ListIndex - 1 To 0 Step -1           'find folder, if so
        Path = .List(Idx)
        If Left$(Path, 1) <> " " Then Exit For        'if no space at head, then folder
      Next Idx
    End If
  End With
  Me.mnuFileImport.Enabled = Left$(Path, 1) <> " " And Not LrnMode    'enable option if it seems valid
  Me.mnuFileImportSegment.Enabled = CBool(Len(Clipboard.GetText(vbCFText))) And Not LrnMode
  Me.mnuFileImportClipBrd.Enabled = Me.mnuFileImportSegment.Enabled And Not LrnMode
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileTron_Click
' Purpose           : Toggle program Trace Mode
'*******************************************************************************
Private Sub mnuFileTron_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-29)
      Exit Sub
    End If
  End If
  
  TraceFlag = Not TraceFlag
  Me.mnuFileTron.Checked = TraceFlag
  Call UpdateStatus
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHelpCodes_Click
' Purpose           : View Program Codes used by VisualCalc
'*******************************************************************************
Private Sub mnuHelpCodes_Click()
  Call Query(-14)
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHelpFind_Click
' Purpose           : Find Items
'*******************************************************************************
Private Sub mnuHelpFind_Click()
  If Query_Pressed Then
    If RunMode Then
      Me.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-45)
      Exit Sub
    End If
  End If
  frmFind.Show vbModeless, Me
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHelpGetStarted_Click
' Purpose           : Show Getting Started help
'*******************************************************************************
Private Sub mnuHelpGetStarted_Click()
  Call Query(0)
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHelpTrig_Click
' Purpose           : Show help for Trig functions
'*******************************************************************************
Private Sub mnuHelpTrig_Click()
  Call Query(-43)
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHelpHistory_Click
' Purpose           : Brief history of programmable calculators
'*******************************************************************************
Private Sub mnuHelpHistory_Click()
  Call Query(-42)
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHelpIndex_Click
' Purpose           : Sorted Index of Commands
'*******************************************************************************
Private Sub mnuHelpIndex_Click()
  Call Query(-15)
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHelpIntro_Click
' Purpose           : Get Introductory Help
'*******************************************************************************
Private Sub mnuHelpIntro_Click()
  Call Query(-13)
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHelpPrev_Click
' Purpose           : Show Previous Help Item match
'*******************************************************************************
Private Sub mnuHelpPrev_Click()
  If Query_Pressed Then
    If RunMode Then
      Me.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-45)
      Exit Sub
    End If
  End If
  
  FindListIdx = FindListIdx - 1
  Call ShowFindListItem
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHelpNext_Click
' Purpose           : Show Next Help Item match
'*******************************************************************************
Private Sub mnuHelpNext_Click()
  If Query_Pressed Then
    If RunMode Then
      Me.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-45)
      Exit Sub
    End If
  End If
  
  FindListIdx = FindListIdx + 1
  Call ShowFindListItem
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHelpSepHlp_Click
' Purpose           : Toggle showing help in a separate window
'*******************************************************************************
Private Sub mnuHelpSepHlp_Click()
  Dim Bol As Boolean
  
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-16)
      Exit Sub
    End If
  End If
  
  Bol = Not Me.mnuHelpSepHlp.Checked                      'flip flag state
  Me.mnuHelpSepHlp.Checked = Bol                          'store new state
  SaveSetting App.Title, "Settings", "SepHelp", CStr(Bol)
  If Bol Then                                             'if enabled
    If Me.rtbInfo.Visible Then                            'if main win has help
      Call CE_Support                                     'clear it
      If LastQuery <> 9999 Then
        Call Query(9998)
      ElseIf colFindList.Count > 0 Then
        ShowFindListItem
      End If
    End If
  Else
    If frmHelpLoaded Then                                 'check only if FrmHelp loaded
      If frmHelp.Visible Then                             'else if help win up
        Unload frmHelp                                    'then remove it
        If Not LrnMode Then
          If LastQuery <> 9999 Then
            Call Query(9998)
          ElseIf colFindList.Count > 0 Then
            ShowFindListItem
          End If
        End If
      End If
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHelpSrch_Click
' Purpose           : Activate Search
'*******************************************************************************
Private Sub mnuHelpSrch_Click()
  If Query_Pressed Then
    If RunMode Then
      Me.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-44)
      Exit Sub
    End If
  End If
  
  frmSearch.Show vbModal, Me
  If NewQuery <> 1000 Then
    Call Query(9997)
  End If
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
' Subroutine Name   : mnuWinASCII_Click
' Purpose           : Display Program listing to the display
'*******************************************************************************
Public Sub mnuWinASCII_Click()
  Dim Ary() As String
  Dim Idx As Integer, i As Integer
  
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-38)
      Exit Sub
    End If
  End If
  
  If InstrCnt = 0 Then                'if no instructions to process
    Call CenterMsgBoxOnForm(Me, "No instructions to Compress.", vbOKOnly Or vbExclamation, "No Pgm Code")
    Exit Sub
  ElseIf Not Preprocessd Then         'if not Preprocessed...
    Call Preprocess                   'proprocess it
    If Not Preprocessd Then Exit Sub  'exit if errors
  End If
  
  Ary = BuildInstrArray()             'get array list of the program
  Call Clear_Screen                   'clear the screen (add an initial, top line
  With frmVisualCalc
    LockControlRepaint .lstDisplay    'loc repainst for faster processomg
    With .lstDisplay
      .Clear
      i = UBound(Ary)
      Do While Not CBool(Len(Trim$(Ary(i))))
        i = i - 1
      Loop
      For Idx = 0 To i                'load the array to the display list
        .AddItem Ary(Idx)
      Next Idx
    End With
    Call RepointIdx                   'point to appropriate line in list
    UnlockControlRepaint .lstDisplay  'allow updates
    DspLocked = True
    Call DspBackground
    DspPgmList = True
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileCopy_Click
' Purpose           : Copy display to clipboard
'*******************************************************************************
Private Sub mnuFileCopy_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-17)
      Exit Sub
    End If
  End If
  
  If Me.rtbInfo.Visible Then
    With Clipboard
      .Clear
      .SetText Me.rtbInfo.TextRTF, vbCFRTF
      .SetText Me.rtbInfo.Text, vbCFText
    End With
  Else
    With Clipboard
      .Clear                              'clear clipboard
      .SetText GetDisplayText(), vbCFText 'save text to it
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileExit_Click
' Purpose           : Exit program
'*******************************************************************************
Private Sub mnuFileExit_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-31)
      Exit Sub
    End If
  End If
  
  If ExitApp() = 0 Then Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHelpAbout_Click
' Purpose           : Display About Box
'*******************************************************************************
Private Sub mnuHelpAbout_Click()
  frmVisualCalc.MousePointer = 0
  Query_Pressed = False             'turn off query mode
  frmAbout.Show vbModal, Me
End Sub

'*******************************************************************************
' Subroutine Name   : mnuWinGreenScreen_Click
' Purpose           : Roggle display colors
'*******************************************************************************
Private Sub mnuWinGreenScreen_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-32)
      Exit Sub
    End If
  End If
  
  Me.mnuWinGreenScreen.Checked = Not Me.mnuWinGreenScreen.Checked
  Call SaveSetting(App.Title, "Settings", "GreenScreen", CStr(Me.mnuWinGreenScreen.Checked))
  
  If Me.mnuWinGreenScreen.Checked Then
    BackClr = vbBlack
    Me.lstDisplay.ForeColor = vbGreen
    Me.txtError(0).BackColor = vbBlack
    Me.txtError(1).BackColor = vbBlack
    Me.txtError(1).ForeColor = vbGreen
  Else
    BackClr = vbWhite
    Me.lstDisplay.ForeColor = vbBlack
    Me.txtError(0).BackColor = vbWhite
    Me.txtError(1).BackColor = vbWhite
    Me.txtError(1).ForeColor = vbBlack
  End If
  Call DspBackground
End Sub

'*******************************************************************************
' Subroutine Name   : mnuWinListVdata_Click
' Purpose           : List Variable contents
'*******************************************************************************
Private Sub mnuWinListVdata_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-35)
      Exit Sub
    End If
  End If
  
  Call ListVdata(True)
End Sub

'*******************************************************************************
' Subroutine Name   : mnuWinListVdataNZ_Click
' Purpose           : List non-zero Variable contents
'*******************************************************************************
Private Sub mnuWinListVdataNZ_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-36)
      Exit Sub
    End If
  End If
  
  Call ListVdata(False)
End Sub

'*******************************************************************************
' Subroutine Name   : mnuWinRight_Click
' Purpose           : Toggle placing the display on the right side of the calculator
'*******************************************************************************
Private Sub mnuWinRight_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-33)
      Exit Sub
    End If
  End If
  
  mnuWinRight.Checked = Not mnuWinRight.Checked   'toggle option
  Call SaveSetting(App.Title, "Settings", "WinRight", CStr(mnuWinRight.Checked))
  
  If mnuWinRight.Checked Then                     'to right...
    Me.PicKeys.Left = 60
    Me.PicScroll.Left = Me.ScaleWidth - Me.PicScroll.Width - 60
'    Me.picDisplay.Left = Me.PicScroll.Left - Me.picDisplay.Width - 60
    Me.picDisplay.Left = Me.PicKeys.Width + 120
  Else                                            'to left...
    Me.PicScroll.Left = 60
    Me.picDisplay.Left = Me.PicScroll.Width + 120
    Me.PicKeys.Left = Me.ScaleWidth - Me.PicKeys.Width - 60
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : picDisplay_Paint
' Purpose           : Tile background
'*******************************************************************************
Private Sub picDisplay_Paint()
  PaintTilePicBackground Me.picDisplay, Me.PicBack
End Sub

'*******************************************************************************
' Subroutine Name   : PicKeys_Paint
' Purpose           : Tile background
'*******************************************************************************
Private Sub PicKeys_Paint()
  PaintTilePicBackground Me.PicKeys, Me.PicBack
End Sub

'*******************************************************************************
' Purpose           : Clear tips if cursor over picture background
'*******************************************************************************
Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If CBool(CharLimit) Then Exit Sub
  OverForm = True
  SetTip vbNullString
End Sub

Private Sub PicKeys_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If CBool(CharLimit) Then Exit Sub
  OverForm = True
  SetTip vbNullString
End Sub

Private Sub PicPlot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  LastPlotXt = X            'save X and Y positions
  LastPlotYt = Y
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  OverForm = True
  SetTip vbNullString
End Sub

Private Sub PicScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If CBool(CharLimit) Then Exit Sub
  OverForm = True
  SetTip vbNullString
End Sub

'*******************************************************************************
' Subroutine Name   : chk2nd_MouseMove
' Purpose           : tip for 2nd key
'*******************************************************************************
Private Sub chk2nd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Toggle 2nd mode functions (also keyboard Ctrl key)"
End Sub

'*******************************************************************************
' Subroutine Name   : cmdBackspace_MouseMove
' Purpose           : Tip for backspace
'*******************************************************************************
Private Sub cmdBackspace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  If Me.rtbInfo.Visible Then
    SetTip "Go back to a previously displayed help window"
  Else
    SetTip "Backspace character or digit entry (Also keyboard Backspace key)"
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : chkShift_MouseMove
' Purpose           : Show tip for Shift key
'*******************************************************************************
Private Sub chkShift_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Shift for keyboard uppercase characters (Also keyboard Shift key)"
End Sub

'*******************************************************************************
' Subroutine Name   : cmdKeyPad_MouseMove
' Purpose           : Show tips for keyboard commands
'*******************************************************************************
Private Sub cmdKeyPad_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static LastMove As Integer
  Static Last2nd As Integer
  
  If RunMode Then Exit Sub  'ignore if run Mode
  If OverForm Then                                          'if was previously off buttons
    OverForm = False                                        'then turn flag off
  Else
    If Index = LastMove And Last2nd = Key2nd Then Exit Sub  'stay less busy with this little test
  End If
  
  LastMove = Index                                          'set new last move
  Last2nd = Key2nd
  
  If CharLimit = 0 Then Call TipCheck(Index)
End Sub

'*******************************************************************************
' Subroutine Name   : cmdUsrA_MouseMove
' Purpose           : User-defined and text entry key tips
'*******************************************************************************
Private Sub cmdUsrA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static LastMove As Integer
  Static Last2nd As Integer
  
  If RunMode Then Exit Sub  'ignore if run Mode
  If OverForm Then                                          'if was previously off buttons
    OverForm = False                                        'then turn flag off
  Else
    If Index = LastMove And Last2nd = Key2nd Then Exit Sub  'stay less busy with this little test
  End If
  
  LastMove = Index                                          'set new last move
  Last2nd = Key2nd
  
  If CharLimit = 0 Then                                     'if something not pending
    If TextEntry Then                                       'if text entry, use generic tag
      SetTip "Alpha-mode Entry key"
    Else
      If CBool(ActivePgm) Then
        SetTip ModLbls(Index + ModLblMap(ActivePgm - 1)).lblCmt 'else apply user comments
      Else
        SetTip Lbls(Index).lblCmt                           'else apply user comments
      End If
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : PicPlot_Click
' Purpose           : User clicked on a location on the plot field
'*******************************************************************************
Private Sub PicPlot_Click()
  If RunMode Then Exit Sub      'do nothing more if Run Mode active
  PlayClick                     'play resource click
  LastPlotX = CLng(LastPlotXt)  'recover X and Y values (from MouseMove)
  LastPlotY = CLng(LastPlotYt)
  If Me.PicPlot.Visible Then    'if plot screen visuble
    If PlotTrigger Then         'if the user defined a trigger subroutine
      InstrPtr = PlotTriggerSbr
      Call Run
      If PmtFlag Then           'if user prompting turned on...
        LastTypedInstr = iTXT   'set TXT command
        Call ActiveKeypad       'activate it (TXT or [=] (ENTER)) will start R/S cmd
      Else
        Call DisplayLine        'else terminating run...
      End If
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : cmdBackspace_Click
' Purpose           : Support backspace key
'*******************************************************************************
Private Sub cmdBackspace_Click()
  Dim Idx As Long, Idy As Long, Idz As Integer
  Dim SS As String
  
  If RunMode Then Exit Sub  'ignore if run Mode
  PlayClick                 'play resource click
  If Query_Pressed Then
    Call Query(-113)        'display help for key if Query active
    Exit Sub
  End If
  
  If CBool(Len(TypeMatTxt)) Then
    TypeMatTxt = vbNullString
    SetTip vbNullString
    Exit Sub
  End If
  
  If Me.rtbInfo.Visible Then                'if Help displayed...
    With colHelpBack
      If .Count < 2 Then
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
      frmVisualCalc.cmdBackspace.Enabled = CBool(.Count > 1)
    End With
    With Me.rtbSearch
      .LoadFile AddSlash(App.Path) & "VPCHelp.rtf", rtfRTF
      Idy = InStr(Idx + 1, .Text, "@@") - 4 'find next block
      If CBool(Idy) Then                    'found it?
        .SelStart = Idy - 1                 'yes, so strip next data off
        .SelLength = Len(.Text) - Idy
        .SelText = vbNullString
        .SelStart = 0                       'strip loading text off
        .SelLength = Idx
        .SelText = vbNullString
        .SelStart = 0                       'set cursor to top of form
      End If
    End With
    With Me.rtbInfo
      .TextRTF = Me.rtbSearch.TextRTF       'now copy data to help display form
      Me.rtbSearch.Text = vbNullString      'erase local
      If CBool(Idz) Then .SelStart = Len(.Text)
      .SelStart = Idz                       'set start of help topic
      .Refresh                              'refresh screen
      DoEvents
    End With
  End If
  
  If TextEntry Then
    If CBool(Len(DspTxt)) Then                'if something to remove...
      DspTxt = Left$(DspTxt, Len(DspTxt) - 1) 'strip last character from string
      CharCount = Len(DspTxt)                 'set character count
      If LrnMode Then                         'if we are in the LRN mode...
        Call LrnBST                           'back up pointer
        Call DeleteInstruction                'remove entry from list
      Else  'we are in active Text Entry...
        Me.lstDisplay.List(Me.lstDisplay.ListIndex) = String$(DisplayWidth - Len(DspTxt), 32) & DspTxt
        DisplayHasText = True
      End If
    Else
      TextEntry = False
    End If
  End If
  If Not TextEntry Then                       'allow backup with delete if in LRN mode
    If LrnMode Then                           'if we are in the LRN mode...
      Call LrnBST                             'back up pointer
      Call DeleteInstruction                  'remove entry from list
    Else
      If ValueTyped And CBool(Len(ValueAccum)) Then
        AllowExp = False                                        'ensure exp edit is deleted
        If Len(ValueAccum) = 1 Then                             'if text length was 1 digit
          ValueAccum = "0"                                      'keep zero
          ValueTyped = False                                    'indicate nothing typed
        Else
          If Right$(ValueAccum, 1) = "." Then ValueDec = False  'if removing decimal place
          Idx = InStr(1, ValueAccum, "E")                       'check for exponent presence
          If CBool(Idx) Then                                    'if it was found...
            ValueAccum = CStr(Val(Left$(ValueAccum, Idx - 1)))  'strip exponent
          Else
            ValueAccum = Left$(ValueAccum, Len(ValueAccum) - 1) 'else just strip last digit (or [.])
          End If
        End If
        Call DisplayAccum                                       'display result
      End If
    End If
  End If
  On Error Resume Next
  Me.txtFocus.SetFocus               'set "dafault" focus
End Sub

'*******************************************************************************
' Subroutine Name   : chkShift_Click
' Purpose           : Activate/Deactivate Shift key
'*******************************************************************************
Private Sub chkShift_Click()
  Static Ignore As Boolean
  
  If RunMode Then Exit Sub                  'ignore if run Mode
  PlayClick                                 'play resource click
  If Query_Pressed Or Me.rtbInfo.Visible Then
    Ignore = True
    If KeyShift Then
      Me.chkShift.Value = vbChecked
    Else
      Me.chkShift.Value = vbUnchecked
    End If
    Ignore = False
    If Query_Pressed Then
      Call Query(-100)                      'ASCII keys, User Keys, Space
    Else
      CmdNotActive
    End If
    Exit Sub
  End If
  KeyShift = Me.chkShift.Value = vbChecked  'set TRUE if shift key pressed
  LastTypedInstr = 128
  Call ResetAccumulator                     'reset accumulator data
  Call ResetAlphaPad                        'reset pad for case
  If CBool(CharCount) And CBool(CharLimit) Then Exit Sub
  SetTip vbNullString
End Sub

'*******************************************************************************
' Subroutine Name   : cmdUsrA_Click
' Purpose           : Support keyboard entry
'*******************************************************************************
Private Sub cmdUsrA_Click(Index As Integer)
  Static Ignore As Boolean
  Dim Idx As Integer, OldPgm As Integer
  Dim LngV As Long
  Dim S As String
  Dim Pool() As Labels
  
  If Ignore Then Exit Sub                 'if we are already in this button
  PlayClick                               'play resource click
  Ignore = True                           'prevent following cmd from repeating
  Me.cmdUsrA(Index).Value = vbUnchecked   'uncheck the check button
  Ignore = False                          'now reset flag
  DoEvents                                'let screen catch up
  Me.cmdUsrA(Index).Refresh
'
' check if user hit Query key
'
  If Query_Pressed Then
    If TextEntry Then
      Call Query(-1)                      'text key help
    Else
      Call Query(-2)                      'user key help
    End If
    Exit Sub
  End If
  
  If Me.rtbInfo.Visible Then
    CmdNotActive
    Exit Sub
  End If
  Me.txtFocus.SetFocus                    'hide focus
  If ErrorPause Or DspLocked Or RunMode Then  'ignore if run mode, error flag, or display locked
    CmdNotActive
    Exit Sub                              'ignore if display locked
  End If
  
  If TextEntry Then                                             'if text entry mode...
    If Len(DspTxt) < CharLimit Then                             'less than limit?
      HaveTxt = True
      If Index = 0 And Not Key2nd Then                          'support space
        S = " "
      Else
        S = Me.cmdUsrA(Index).Caption                           'else use key caption for text character
      End If
      If LrnMode Then
        LastTypedInstr = Asc(S)                                 'set keycommand
        AddInstruction LastTypedInstr                           'add to learn mode
      Else
        DspTxt = DspTxt & S
        CharCount = Len(DspTxt)                                 'establish character count
        If Not RunMode Then                                     'show updates if not in RUN mode
          Me.lstDisplay.List(Me.lstDisplay.ListIndex) = String$(DisplayWidth - Len(DspTxt), 32) & DspTxt
        End If
        DisplayHasText = True
      End If
    End If
  Else
    If LrnMode Then                                             'if learn mode, allow entering user buttons
      Call AddInstruction(Index + 900)                          'add command
    Else
      BraceIdx = 0                                              'reset braceing index us user key hit
      SbrInvkIdx = 0                                            'reset subr stack if user-key hit
      If CBool(ActivePgm) Then
        Pool = ModLbls
        LngV = CLng(Index) + ModLblMap(ActivePgm - 1)
      Else
        If CBool(InstrCnt) Then
          If Not Preprocessd Then                               'if not Preprocessed
            Call Preprocess                                     'Preprocess it
            If Not Preprocessd Then Exit Sub                    'if still not Preprocessed
          End If
        End If
        Pool = Lbls
        LngV = CLng(Index)
      End If
      With Pool(LngV)                                           'point to user-defined key data
        Idx = .lblAddr                                          'get address of definition
        If .LblDat = 0 Then                                     'if no address...
          If CBool(ActivePgm) Or Preprocessd Then
            CmdNotActive
          Else
            ForcError "User-Defined Key '" & Chr$(Index + 64) & "' is not defined in program"
          End If
        Else
          InstrPtr = .LblDat + 1                                'point past '{'
          ModPrep = 0                                           'disable flag
          Call Run                                              'run program
' if after running, the Pmt command is activated, we will turn on the TXT mode (TextEntry),
' let the user type in a response, and by pressing TXT or '=' (ENTER), the program will continue
' (see KybdMain).
          If CBool(MRunMode) Then MRunMode = MRunMode - 1
          If PmtFlag Then                     'if user prompting turned on...
            LastTypedInstr = iTXT             'set TXT command
            Call ActiveKeypad                 'activate it (TXT or [=] (ENTER)) will start R/S cmd
          Else
            Call DisplayLine                  'else terminating run...
          End If
'---------------
        End If
      End With
    End If
  End If
End Sub

'*******************************************************************************
' Tips for various indicators
'*******************************************************************************
Private Sub lblLoc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Program Step in LRN mode"
End Sub

Private Sub lblCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Instruction code in LRN mode"
End Sub

Private Sub lblInstr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Keypad command for instruction code"
End Sub

Private Sub lbltxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Text Entry Mode active indicator"
End Sub

Private Sub lblINS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Insert Mode active indicator"
End Sub

Private Sub lblLRN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "LRN Mode active indicator"
End Sub

Private Sub lblEE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  If Me.lblEE.Caption = "Eng" Then
    SetTip "Engineering Format active indicator"
  Else
    SetTip "Enter Exponent active indicator"
  End If
End Sub

Private Sub lblDRGM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Angles mode"
End Sub

Private Sub lblHDOB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Number Base"
End Sub

Private Sub rtbInfo_Click()
  If RunMode Then Exit Sub        'ignore if run Mode
  If Not Selecting Then PlayClick 'play resource click
End Sub

Private Sub sbrImmediate_PanelClick(ByVal Panel As MSComctlLib.Panel)
  Dim TV As Double
  Select Case Panel.Key
    Case "Tron"     'Trace mode
      TraceFlag = Not TraceFlag
      Me.mnuFileTron.Checked = TraceFlag
      Call UpdateStatus
    
    Case "Style"    'Style mode
      LRNstyle = LRNstyle + 1
      If LRNstyle > 3 Then LRNstyle = 0
      SaveSetting App.Title, "Settings", "Style", CStr(LRNstyle)
      Call UpdateStatus
      TV = DisplayReg                           'save displayreg value
      If Not LrnMode Then
        If DspLocked Then
          Call frmVisualCalc.mnuWinASCII_Click  'list program
          Call ResetPndAll
          Exit Sub
        ElseIf AutoPprc Then
          If Not Preprocessd Then               'if not at least Preprocessed...
            Call Preprocess                     'then Preprocess
          End If
        End If
      End If
      DisplayReg = TV                           'reset display to prior value
  End Select
End Sub

'*******************************************************************************
' Subroutine Name   : tmrWait_Timer
' Purpose           : Check if window size gets too small
'*******************************************************************************
Private Sub tmrWait_Timer()
  If GetKeyState(VK_LBUTTON) < 0 Then
    Exit Sub                                      'if mouse down, then ignore for now
  End If
  Me.tmrWait.Enabled = False                      'turn off timer
  If Me.WindowState = vbMinimized Then Exit Sub   'do nothing if minimized
  Resize = True
  If Me.Width < WinMinW Then Me.Width = WinMinW   'resize to minimum dims
  If Me.Height < WinMinH Then Me.Height = WinMinH
  Resize = False
  Call Form_Resize                                'process resizing
End Sub

Private Sub txtError_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  OverForm = True
  SetTip vbNullString
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
  If SkipChg Then Exit Sub                          'skip processing selection changes
  
  Screen.MousePointer = vbHourglass
  Selecting = True
  FndMatch = False
  Do
    With Me.rtbInfo
      SS = ";" & CStr(.SelStart)                    'save start address
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
              If Len(S) > 30 Then S = vbNullString
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
    
    With Me.rtbSearch
      On Error Resume Next
      .LoadFile Pth, rtfRTF                         'load help file
      If CBool(Err.Number) Then Exit Do
      On Error GoTo 0
      Txt = .Text                                   'grab text
      Idx = InStr(1, Txt, S)                        'find first instance of data
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
      If Left$(S, 2) = "@@" Then
        Idx = InStr(Idx, Txt, "@[")
      ElseIf Left$(S, 1) = "@" Then
        Idx = InStr(Idx, Txt, Left$(S, 2))
      Else
        Idx = InStr(Idx, Txt, "@[")
      End If
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
          .Add CStr(Idx)                      'add anyway if noting in buffer
        End If
        frmVisualCalc.cmdBackspace.Enabled = CBool(.Count > 1)
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
    
    With Me.rtbInfo
      .TextRTF = Me.rtbSearch.TextRTF       'now copy data to help display form
      Me.rtbSearch.Text = vbNullString      'erase local
      .SelStart = 0                         'set start of help topic
      .Refresh                              'refresh screen
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
' Subroutine Name   : tmrPause_Timer
' Purpose           : Support Pause Mode
'*******************************************************************************
Private Sub tmrPause_Timer()
  Static LocalTimer As Long
  
  If LocalTimer = 0 Then                      'if local timer is set to 0
    If tmrWait = 0 Then                       'and pause value is 0
      LocalTimer = 1                          'default to 1/2 second
    Else
      LocalTimer = tmrWait                    'set set limit
    End If
  End If
  
  LocalTimer = LocalTimer - 1                 'back count off
  If LocalTimer = 0 Then                      'if we timed out...
    Me.tmrPause.Enabled = False               'turn off timer
    If RunMode Then                           'if we are in the RUN mode...
      With Me.lstDisplay
        .List(.ListIndex) = vbNullString      'rehide the display
      End With
      DoEvents
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : PicScroll_Paint
' Purpose           : Tile background
'*******************************************************************************
Private Sub PicScroll_Paint()
  PaintTilePicBackground Me.PicScroll, Me.PicBack
End Sub

'*******************************************************************************
' Support scrolling buttons
'*******************************************************************************
Private Sub cmdPgUp_Click()
  If RunMode Then Exit Sub  'ignore if run Mode
  PlayClick                 'play resource click
  If Query_Pressed Then
    Call Query(-112)        'display help for key if Query active
    Exit Sub
  End If
  Call Form_KeyDown(33, 0)
  Me.txtFocus.SetFocus      'set "dafault" focus
End Sub

Private Sub cmdPgDn_Click()
  If RunMode Then Exit Sub  'ignore if run Mode
  PlayClick                 'play resource click
  If Query_Pressed Then
    Call Query(-114)        'display help for key if Query active
    Exit Sub
  End If
  Call Form_KeyDown(34, 0)
  Me.txtFocus.SetFocus      'set "dafault" focus
End Sub

Private Sub cmdBtm_Click()
  If RunMode Then Exit Sub  'ignore if run Mode
  PlayClick 'play resource click
  If Query_Pressed Then
    Call Query(-115)        'display help for key if Query active
    Exit Sub
  End If
  Call Form_KeyDown(35, 0)
  Me.txtFocus.SetFocus      'set "dafault" focus
End Sub

Private Sub cmdTop_Click()
  If RunMode Then Exit Sub  'ignore if run Mode
  PlayClick 'play resource click
  If Query_Pressed Then
    Call Query(-111)        'display help for key if Query active
    Exit Sub
  End If
  Call Form_KeyDown(36, 0)
  Me.txtFocus.SetFocus      'set "dafault" focus
End Sub

Private Sub cmdUp_Click()
  If RunMode Then Exit Sub  'ignore if run Mode
  PlayClick                 'play resource click
  If Query_Pressed Then
    Call Query(-110)        'display help for key if Query active
    Exit Sub
  End If
  Call Form_KeyDown(38, 0)
  Me.txtFocus.SetFocus      'set "dafault" focus
End Sub

Private Sub cmdDn_Click()
  If RunMode Then Exit Sub  'ignore if run Mode
  PlayClick                 'play resource click
  If Query_Pressed Then
    Call Query(-116)        'display help for key if Query active
    Exit Sub
  End If
  Call Form_KeyDown(39, 0)
  Me.txtFocus.SetFocus      'set "dafault" focus
End Sub

'*******************************************************************************
' Tips for buttons
'*******************************************************************************
Private Sub cmdBtnHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Click this button, and then click another button for additional information"
End Sub

Private Sub cmdBtm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Move to the bottom of the display list"
End Sub

Private Sub rtbInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Press CE to clear Help display. Double-click word to search for reference"
End Sub

Private Sub cmdDn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Move selection line in display down"
End Sub

Private Sub cmdPgDn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Page Down in the display list"
End Sub

Private Sub cmdPgUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Page Up in the display list"
End Sub

Private Sub cmdTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Move to the top of the display list"
End Sub

Private Sub cmdUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  SetTip "Move selection line in display up"
End Sub

Private Sub imgLocked_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode Then Exit Sub  'ignore if run Mode
  If CBool(CharLimit) Then Exit Sub
  If Index = 0 Then
    SetTip "The Display List is locked. Use the CE key to clear and unlock it"
  Else
    SetTip "The Display List is unlocked and available for editing"
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuKeypadBasic_Click
' Purpose           : View Basic keypad (Basic Programmer's Calc)
'*******************************************************************************
Private Sub mnuKeypadBasic_Click()
  Dim Idx As Integer
  
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-34)
      Exit Sub
    End If
  End If
  
  PanelsVNV False                         'hide programming panels
  Me.mnuKeypadBasic.Checked = True        'set checks as needed
  Me.mnuKeypadAdvanced.Checked = False
  Me.mnuKeypadFull.Checked = False
  Me.mnuMemStk.Visible = False
  SaveSetting App.Title, "Settings", "KeyLayout", "0"
  
  TextEntry = False                     'turn off text entry mode
  
  For Idx = 0 To 26                     'hide User Keys
    Me.cmdUsrA(Idx).Visible = False
  Next Idx
  
  For Idx = 1 To MaxKeys                'enable/disable and show/hide keypads as needed
    Select Case Idx
      Case 41 To 43, 53 To 56, 66 To 69, 79 To 82, 92 To 95, _
           32, 33, 45, 46, 58, 59, 71, 72, 85, 98, 5, 6
        Me.cmdKeyPad(Idx).Visible = True
      Case 26, 39, 52, 65, 78, 91
        Me.cmdKeyPad(Idx).Visible = True
        If BaseType <> TypHex Or Key2nd Then
          Me.cmdKeyPad(Idx).Enabled = False
        End If
      Case Else
        Me.cmdKeyPad(Idx).Visible = False
    End Select
  Next Idx
  
  Me.lblLoc.Visible = False            'hide various programming status fields
  Me.lblCode.Visible = False
  Me.lblInstr.Visible = False
  Me.lblTxt.Visible = False
  Me.lblINS.Visible = False
  Me.lblLRN.Visible = False
  Me.lblEE.Visible = False
  Me.lblDRGM.Visible = False
  Me.chkShift.Value = vbUnchecked
  Me.chkShift.Visible = False
  With Me.sbrImmediate
    .Panels("MDL").Visible = False
    .Panels("Pgm").Visible = False
    .Panels("InstrCnt").Visible = False
    .Panels("InstrPtr").Visible = False
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : mnuKeypadAdvanced_Click
' Purpose           : View Advanced keypad (Scientific Programmer's Calc)
'*******************************************************************************
Private Sub mnuKeypadAdvanced_Click()
  Dim Idx As Integer
  
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-34)
      Exit Sub
    End If
  End If
  
  PanelsVNV False                         'hide programming panels
  Me.mnuKeypadBasic.Checked = False       'set checks as needed
  Me.mnuKeypadAdvanced.Checked = True
  Me.mnuKeypadFull.Checked = False
  Me.mnuMemStk.Visible = False
  SaveSetting App.Title, "Settings", "KeyLayout", "1"
  
  TextEntry = False                     'turn off text entry
  
  For Idx = 0 To 26                     'hide User Keys
    Me.cmdUsrA(Idx).Visible = False
  Next Idx
  
  For Idx = 1 To MaxKeys                'enable/disable and show/hide keypads as needed
    Select Case Idx
      Case 12 To 18, 21 To 25, 27 To 30, 32, 33, 41 To 43, 45 To 51, _
           53 To 56, 58, 59, 64, _
           66 To 69, 71, 72, 77, _
           79 To 82, 85, 90, _
           92 To 95, 98 To 103, 5 To 7
        Me.cmdKeyPad(Idx).Visible = True
        If Key2nd Then
          Me.cmdKeyPad(77).Enabled = False
          Me.cmdKeyPad(90).Enabled = False
          Me.cmdKeyPad(103).Enabled = False
        End If
      Case 26 'Hyp/Arc
        Me.cmdKeyPad(Idx).Visible = True
        If BaseType <> TypDec And BaseType <> TypHex Then
          Me.cmdKeyPad(Idx).Enabled = False
        ElseIf BaseType = TypHex Then
          If Key2nd Then
            Me.cmdKeyPad(Idx).Enabled = False
          Else
            Me.cmdKeyPad(Idx).Enabled = True
          End If
        Else
          Me.cmdKeyPad(Idx).Enabled = True
          If Key2nd Then
            Me.cmdKeyPad(Idx).Caption = "Hyp"
          Else
            Me.cmdKeyPad(Idx).Caption = "Arc"
          End If
        End If
        
      Case 39, 52, 65, 78, 91
        Me.cmdKeyPad(Idx).Visible = True
        If BaseType <> TypHex Or Key2nd Then
          Me.cmdKeyPad(Idx).Enabled = False
        End If
      Case Else
        Me.cmdKeyPad(Idx).Visible = False
    End Select
  Next Idx
  Me.lblLoc.Visible = False               'hide/display data status fields as needed
  Me.lblCode.Visible = False
  Me.lblInstr.Visible = False
  Me.lblTxt.Visible = False
  Me.lblINS.Visible = False
  Me.lblLRN.Visible = False
  Me.lblEE.Visible = True
  Me.lblDRGM.Visible = True
  Me.chkShift.Value = vbUnchecked
  Me.chkShift.Visible = False
  With Me.sbrImmediate
    .Panels("MDL").Visible = False
    .Panels("Pgm").Visible = False
    .Panels("InstrCnt").Visible = False
    .Panels("InstrPtr").Visible = False
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : mnuKeypadFull_Click
' Purpose           : View Programmer's Full Keypad
'*******************************************************************************
Private Sub mnuKeypadFull_Click()
  Dim Idx As Integer
  
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-34)
      Exit Sub
    End If
  End If
  
  PanelsVNV True                              'enable status fields in status bar
  Me.mnuKeypadBasic.Checked = False           'set checks as needed
  Me.mnuKeypadAdvanced.Checked = False
  Me.mnuKeypadFull.Checked = True
  Me.mnuMemStk.Visible = True
  SaveSetting App.Title, "Settings", "KeyLayout", "2"
  
  TextEntry = False                         'turn off text entry mode
  
  For Idx = 0 To 26                         'show User Keys as specified
    Me.cmdUsrA(Idx).Visible = Hidden(Idx)
  Next Idx
  
  For Idx = 1 To MaxKeys                    'enable/show all keypads
    Me.cmdKeyPad(Idx).Visible = True
    Me.cmdKeyPad(Idx).Enabled = True
  Next Idx
  
  Me.lblLoc.Visible = True                  'display all programming status fields
  Me.lblCode.Visible = True
  Me.lblInstr.Visible = True
  Me.lblTxt.Visible = True
  Me.lblINS.Visible = True
  Me.lblLRN.Visible = True
  Me.lblEE.Visible = True
  Me.lblDRGM.Visible = True
  Me.chkShift.Visible = True
  With Me.sbrImmediate
    .Panels("MDL").Visible = True
    .Panels("Pgm").Visible = True
    .Panels("InstrCnt").Visible = True
    .Panels("InstrPtr").Visible = True
  End With
  Call Me.chk2nd_Click                      'fix up display
End Sub

'*******************************************************************************
' Subroutine Name   : mnuMemStk_Click
' Purpose           : Change/Create a Data Storage location
'*******************************************************************************
Private Sub mnuMemStk_Click()
  Dim Fn As String
  Dim IsRdy As Boolean
  
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-40)
      Exit Sub
    End If
  End If
  
  If CBool(Len(StorePath)) Then
    Fn = Fso.GetFolder(StorePath)
  Else
    Fn = CurDir$
  End If
  Fn = DirBrowser(Me.hWnd, ViewDirsOnly, "Select Base Data Storage Location", Fn)
  If Len(Fn) = 0 Then Exit Sub
'
' if the path is a file, we cannot use it
'
  If Fso.FileExists(Fn) Then
    CmdNotActive
    CenterMsgBoxOnForm Me, Fn & " is a file, and cannot be used as a Data Storage location.", vbOKOnly Or vbExclamation, "Data Storage Error"
    Exit Sub
  End If
'
' if the path is a drive or netdrive, we cannot use it
'
  If Right$(Fn, 1) = "\" Or Right$(Fn, 1) = "/" Then
    CmdNotActive
    CenterMsgBoxOnForm Me, "Cannot specify a root location on a drive", vbOKOnly Or vbExclamation, "Cannot Specify Root"
    Exit Sub
  End If
'
' determine if the folers exists, and if so, if it is already initialized
'
  If Not Fso.FolderExists(Fn) Then
    On Error Resume Next
    Fso.CreateFolder Fn
    If CBool(Err.Number) Then
      CmdNotActive
      CenterMsgBoxOnForm Me, "Cannot create " & Fn & " Data Storage location", vbOKOnly Or vbExclamation, "Cannot Create Location"
      Exit Sub
    End If
    On Error GoTo 0
    IsRdy = False
  Else
    If Fso.FolderExists(Fn & "\MDL") And _
       Fso.FolderExists(Fn & "\PGM") And _
       Fso.FolderExists(Fn & "\DATA") Then
      IsRdy = True  'already initialized
    End If
  End If
'
' if the location is not initialized, then initialize it by adding the required paths
'
  If Not IsRdy Then
    If Fso.GetFolder(Fn).Attributes = ReadOnly Then
      CmdNotActive
      CenterMsgBoxOnForm Me, "Cannot initialize " & Fn & "." & vbCrLf & "It is Read-Only.", _
                             vbExclamation Or vbOKOnly, "Location is Read-Only"
      Exit Sub
    End If
    On Error Resume Next
    If Not Fso.FolderExists(Fn & "\MDL") Then
      Fso.CreateFolder (Fn & "\MDL")
      If CBool(Err.Number) Then
        CmdNotActive
        CenterMsgBoxOnForm Me, "Cannot create Required MDL folder", _
                               vbExclamation Or vbOKOnly, "Stick Initialization Failure"
        Exit Sub
      End If
    End If
    If Not Fso.FolderExists(Fn & "\PGM") Then
      Fso.CreateFolder (Fn & "\PGM")
      If CBool(Err.Number) Then
        CmdNotActive
        CenterMsgBoxOnForm Me, "Cannot create Required PGM folder", _
                               vbExclamation Or vbOKOnly, "Stick Initialization Failure"
        Exit Sub
      End If
    End If
    If Not Fso.FolderExists(Fn & "\DATA") Then
      Fso.CreateFolder (Fn & "\DATA")
      If CBool(Err.Number) Then
        CmdNotActive
        CenterMsgBoxOnForm Me, "Cannot create Required DATA folder", _
                               vbExclamation Or vbOKOnly, "Stick Initialization Failure"
        Exit Sub
      End If
    End If
'
' we have initialize the Data Storage location
'
  End If
  On Error GoTo 0
  
  StorePath = Fn    'set new path
  SaveSetting App.Title, "Settings", "StorePath", Fn
  With Me.mnuMemStk
    .Caption = "Base Data Storage Location: " & StorePath
  End With
  Me.mnuFileListDir.Enabled = True
End Sub

'*******************************************************************************
' Process listing data
'*******************************************************************************
Private Sub mnuWINListMDL_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-37)
      Exit Sub
    End If
  End If
  
  CheckList -iMDL
End Sub

Private Sub mnuWinVar_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-39)
      Exit Sub
    End If
  End If
  
  CheckList -iVar
End Sub

Private Sub mnuWinLbl_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-39)
      Exit Sub
    End If
  End If
  
  CheckList -iLbl
End Sub

Private Sub mnuWinUkey_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-39)
      Exit Sub
    End If
  End If
  
  CheckList -iUkey
End Sub

Private Sub mnuWinSbr_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-39)
      Exit Sub
    End If
  End If
  
  CheckList -iSbr
End Sub

Private Sub mnuWinConst_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-39)
      Exit Sub
    End If
  End If
  
  CheckList -iConst
End Sub

Private Sub mnuWinStruct_Click()
  If Query_Pressed Then
    If RunMode Then
      frmVisualCalc.MousePointer = 0
      Query_Pressed = False
    Else
      Call Query(-39)
      Exit Sub
    End If
  End If
  
  CheckList -iStruct
End Sub

'*******************************************************************************
' Subroutine Name   : cmdBtnHelp_Click
' Purpose           : Process Queries
'*******************************************************************************
Private Sub cmdBtnHelp_Click()
  If RunMode Then Exit Sub
  PlayClick                 'play resource click
  If DspLocked Then
    CmdNotActive
    Exit Sub
  End If
  If Query_Pressed Then
    Call Query(-101)        'Query key
    Exit Sub
  Else
    Query_Pressed = True    'enable query
    Me.MousePointer = 99    'set special mouse pointer
    DoEvents                'update screen
  End If
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************
