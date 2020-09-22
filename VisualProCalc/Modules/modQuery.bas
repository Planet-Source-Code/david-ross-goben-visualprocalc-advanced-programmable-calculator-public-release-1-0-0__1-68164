Attribute VB_Name = "modQuery"
Option Explicit

Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" _
       (ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long
Private Const EM_LINESCROLL As Long = &HB6

Public TopicOfst As Long
Public LastQuery As Integer 'previous query index
Public NewQuery As Integer
Public NewQueryOFST As Long
Public NewQueryLen As Integer

'*******************************************************************************
' Subroutine Name   : Query
' Purpose           : Process VCHelp Queries
'*******************************************************************************
Public Sub Query(ByVal Index As Integer)
  Dim T As String
  Dim Pth As String, S As String
  Dim Idx As Long, Idy As Long
  Dim RTB As RichTextBox
  Dim IndexOfst As Long
  
  Selecting = True
  With frmVisualCalc
    .mnuFile.Enabled = .mnuHelpSepHlp.Checked
    .mnuWindow.Enabled = .mnuFile.Enabled
  End With
  frmVisualCalc.MousePointer = 0
  Screen.MousePointer = vbHourglass
  DoEvents
  Do
    If Not frmHelpLoaded And Not frmVisualCalc.rtbInfo.Visible And Index <> 9998 Then
      With colHelpBack                  'reinit backstepping in help
        Do While CBool(.Count)
          .Remove 1
        Loop
      End With
    End If
    
    Query_Pressed = False             'turn off query mode
  
    DoEvents                          'let screen catch up
    IndexOfst = 0                     'start of tpoic
    If Index = 9998 Then              'we want to redisplay last query
      Index = LastQuery
    ElseIf Index = 9997 Then          'if new query is requested
      Index = NewQuery
      IndexOfst = NewQueryOFST
    Else
      If Index > 0 Then
        If Key2nd Then
          Index = Index + 256         'mark 2nd key opcodes as 256+
        Else
          Index = Index + 128         'mark normal opcodes as 128+ (128 is ignored)
        End If
      End If
    End If
    Select Case Index
      '0-9 digits
      Case 220, 207, 208, 209, 194, 195, 196, 181, 182, 183
        Index = 181
      Case 169, 170   '( and )
        Index = 169
      Case iRem, iRem2 'Rem and '
        Index = iRem2
      Case 238, 239   '[ and ]
        Index = 238
      Case 190, 191   '{ and }
        Index = 190
      Case 188, 189   'Select and Case
        Index = 188
      Case 214, 215, 343 'If and Else and ElseIf
        Index = 214
    End Select
    
    If BaseType = TypHex Then
      Select Case Index
        Case iHyp
          Index = 10
        Case iDfn
          Index = 11
        Case iSbr
          Index = 12
        Case iLbl
          Index = 13
        Case iUkey
          Index = 14
        Case iRunStop
          Index = 15
      End Select
    End If
   
    Pth = AddSlash(App.Path) & "VPCHelp.rtf" 'default path to help file
    If Not Fso.FileExists(Pth) Then Exit Do
'
' allow faster processing by keeping rtb native to target
'
    With frmVisualCalc
      If .mnuHelpSepHlp.Checked Then
        If Not frmHelpLoaded Then Load frmHelp
        Set RTB = frmHelp.rtbInfo
      Else
        Set RTB = .rtbInfo
      End If
    
      If RTB.Visible Then
        Call LockControlRepaint(RTB)
      ElseIf Not .mnuHelpSepHlp.Checked Then
        .cmdUp.Enabled = False
        .cmdDn.Enabled = False
        .cmdPgUp.Enabled = False
        .cmdBackspace.Enabled = False
        .cmdPgDn.Enabled = False
        .cmdTop.Enabled = False
        .cmdBtm.Enabled = False
      End If
      
      With RTB
        On Error Resume Next
        .LoadFile Pth, rtfRTF                   'load help file
        If CBool(Err.Number) Then Exit Do
        On Error GoTo 0
        T = "@@" & CStr(Index) & "@"            'text to search for
        
        Idx = InStr(1, .Text, T)                'find search text
        If CBool(Idx) Then                      'found it?
          Idx = Idx + Len(T) - 1                'yes, so point to last character
          With colHelpBack
            If CBool(.Count) Then
              S = .Item(.Count)
              Idy = InStr(1, S, ";")
              If CBool(Idy) Then S = Left$(S, Idy - 1)
              If CLng(S) <> Idx Then  'if last item <> new item...
                .Add CStr(Idx)                    'save location on internal stack
              End If
            Else
              .Add CStr(Idx)                      'save location on internal stack
            End If
          End With
          Idy = InStr(Idx + 1, .Text, "@@") - 4 'find next block
          If CBool(Idy) Then                    'found it?
            SkipChg = True
            .SelStart = Idy - 1                 'yes, so strip next data off
            .SelLength = Len(.Text) - Idy
            .SelText = vbNullString
            .SelStart = 0                       'strip loading text off
            .SelLength = Idx
            .SelText = vbNullString
            If CBool(IndexOfst) Then
              .SelStart = Len(.Text)            'bittom of text
              .SelStart = IndexOfst             'point to match line
              .SelLength = NewQueryLen
              .SelColor = vbRed
              .SelStart = IndexOfst
              Call SendMessageByNum(.hWnd, EM_LINESCROLL, 0&, -2&)  'scroll up 2 lines
            Else
              .SelStart = 0                     'set cursor to top of form
            End If
            SkipChg = False
            If RTB.Visible Then Call UnlockControlRepaint(RTB)
            If frmVisualCalc.mnuHelpSepHlp.Checked Then 'if help on separate form
              With frmHelp
                If Not .Visible Then                    'if form not yet visisble
                  .Show                                 'show it
                  DoEvents
                End If
                .Toolbar1.Buttons("Back").Enabled = CBool(colHelpBack.Count > 1)
                If .WindowState = vbMinimized Then      'if minimized, then bring it up
                  .WindowState = vbNormal
                End If
                StayOnTop frmHelp                       'place form on top
                DoEvents
                .tmrOnTop.Enabled = True
              End With
            Else
              frmVisualCalc.cmdBackspace.Enabled = CBool(colHelpBack.Count > 1)
              With frmVisualCalc.rtbInfo
                If Not .Visible Then .Visible = True    'otherwise, simply show main form help
              End With
            End If
            LastQuery = Index
            DoEvents
          End If
        End If
      End With
    End With
    Exit Do
  Loop
  Call ClearFindList
  Screen.MousePointer = vbDefault
  DoEvents
  Selecting = False
End Sub

'*******************************************************************************
' Subroutine Name   : ClearFindList
' Purpose           : Erase contents of Find List collection
'*******************************************************************************
Public Sub ClearFindList()
  With colFindList
    Do While CBool(.Count)
      .Remove 1
    Loop
    FindListIdx = 0
  End With
  With frmVisualCalc
    .mnuHelpPrev.Enabled = False
    .mnuHelpNext.Enabled = False
    If frmHelpLoaded Then
      With frmHelp
        .Caption = "Topic Help"
        .Toolbar1.Buttons("Prev").Enabled = False
        .Toolbar1.Buttons("Next").Enabled = False
        .Toolbar1.Buttons("List").Enabled = False
      End With
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : ShowFindListItem
' Purpose           : Show an item in the find list
'*******************************************************************************
Public Sub ShowFindListItem()
  Dim Bol As Boolean
  Dim Idx As Long, Idy As Long, Idz As Long, Idw As Long
  Dim Pth As String, Txt As String, S As String, TTL As String
  Dim Topic As Integer, SecFnd As Integer
  Dim RTB As RichTextBox
  Dim LoadHlp As Boolean
  
  Screen.MousePointer = vbHourglass           'indicate that we are busy
  DoEvents
'
' keep flag of help file information
'
  If frmVisualCalc.mnuHelpSepHlp.Checked Then 'showing separate form?
    Bol = frmHelpLoaded                       'yes, is it already loaded
    If Not Bol Then
      Load frmHelp                            'load help form if not loaded
    End If
    Set RTB = frmHelp.rtbInfo
  Else
    With frmVisualCalc                        'we are displaying help on main form
      Set RTB = .rtbInfo                      'set to local RTB
      Bol = .rtbInfo.Visible                  'flag if visible or not
    End With
  End If
  
  If Bol Then                                 'is RTB displayed?
    LockControlRepaint RTB                    'lock RTB updates if already visible
  Else
    TopicOfst = 0
  End If
  
  Idx = CLng(colFindList(FindListIdx))        'get index of data
  LoadHlp = True
  If CBool(TopicOfst) Then
    If Idx >= TopicOfst And Idx < TopicOfst + Len(RTB.Text) And FindListIdx <> 1 Then
      LoadHlp = False
    End If
  End If
  LastQuery = 9999                            'reset regular query processing
'
' load help file and gather topics
'
  SkipChg = True
  With RTB
    If LoadHlp Then
      Pth = AddSlash(App.Path) & "VPCHelp.rtf" 'path to help file
      .LoadFile Pth, rtfRTF                   'load help file
      Txt = .Text                             'grab text
      Idy = InStrRev(Txt, "@@", Idx)          'find start of data
      If Idy = 0 Then
        Idy = 1
        Idz = 1
      Else
        Idz = InStr(Idy + 2, Txt, "@") + 1    'find end of topic data
        Idy = Idx - Idz
      End If
      Idx = InStr(Idx + 1, Txt, "@@") - 4     'find end of topic
      If Idx <= 0 Then Idx = Len(Txt)
      .SelStart = Idx - 1
      .SelLength = Len(.Text) - Idx
      .SelText = vbNullString
      .SelStart = 0                           'strip loading text off
      TopicOfst = Idz - 1
      .SelLength = TopicOfst
      .SelText = vbNullString
    Else
      Idy = Idx - TopicOfst - 1               'point to position in current topic
    End If
    .SelStart = Idy                           'set cursor to top of form
    .SelLength = FindListLen
'
' now find all instances of data on the page
'
    S = .SelText                              'get copy of text we are searching for
    TTL = "Find Text: " & S
    Txt = .Text                               'get copy of text
    Idx = InStr(1, Txt, S, vbTextCompare)     'find matches
    SecFnd = 0
    Do While CBool(Idx)
      SecFnd = SecFnd + 1
      .SelStart = Idx - 1
      .SelLength = FindListLen
      .SelColor = vbRed
      .SelBold = True
      Idx = InStr(Idx + FindListLen, Txt, S, vbTextCompare)
    Loop
'
' now report to primary target
'
    If Idx < 0 Then Idx = 0                   'do not split below bounds
    .SelStart = Len(.Text)                    'point to end, to adjust scroll
    .SelStart = Idy                           'set cursor to target
    .SelLength = FindListLen
    .SelLength = 0
    Call SendMessageByNum(.hWnd, EM_LINESCROLL, 0&, -2&)  'scroll up 2 lines
    DoEvents
  End With
  
  If Bol Then                                 'if form was already loaded
    UnlockControlRepaint RTB
  ElseIf frmVisualCalc.mnuHelpSepHlp.Checked Then 'should the help form be displayed
    frmHelp.Show                              'else show the help form
  Else
    RTB.Visible = True                        'is local, so show the help form
    frmVisualCalc.mnuFile.Enabled = False
    frmVisualCalc.mnuWindow.Enabled = False
  End If
  With frmVisualCalc
    .mnuHelpPrev.Enabled = FindListIdx > 1
    .mnuHelpNext.Enabled = FindListIdx < colFindList.Count
    If frmHelpLoaded Then
      With frmHelp
        .Caption = "Topic Help: (List Item " & CStr(FindListIdx) & " of " & _
                   CStr(colFindList.Count) & "; " & CStr(SecFnd) & " found in this topic)"
        .Toolbar1.Buttons("Prev").Enabled = FindListIdx > 1
        .Toolbar1.Buttons("Next").Enabled = FindListIdx < colFindList.Count
        With .Toolbar1.Buttons("List")
          .Enabled = CBool(colFindList.Count)
          With .ButtonMenus
            .Clear
            For Idx = 1 To colFindList.Count
              .Add , "K" & CStr(Idx), CStr(Idx)
            Next Idx
          End With
        End With
      End With
    End If
  End With
  Screen.MousePointer = vbDefault             'no longer busy
  DoEvents
  SkipChg = False
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

