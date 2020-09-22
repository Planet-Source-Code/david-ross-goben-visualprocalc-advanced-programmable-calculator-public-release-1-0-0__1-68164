Attribute VB_Name = "modKbdSupport"
Option Explicit

Public Selecting As Boolean

'*******************************************************************************
' Subroutine Name   : ResetAlphaPad
' Purpose           : Reset keyboard pad to A-Z values
'*******************************************************************************
Public Sub ResetAlphaPad()
  Dim i As Integer
  Dim S As String
  Dim Bol As Boolean
  
  Bol = frmVisualCalc.chk2nd.Value = vbUnchecked  'set flag if 2nd key up
  If Not TextEntry Then Bol = True                'force flag if not text entry mode
  If Bol Then                                     'flag set?
    If KeyShift Then                              'shifted?
      S = AlphaKeys                               'yes, so uppercase keys (A-Z)
    Else
      S = LCase$(AlphaKeys)                       'set lowercase key (a-z)
    End If
  Else
    S = AltKeys                                   '2nd keys (special characters)
  End If
  '
  ' handle spacebar seperately
  '
  With frmVisualCalc.cmdUsrA(0)
    If Bol Then
      .Caption = "SPACE"
    Else
      .Caption = "!"
    End If
    .Visible = True
  End With
  
  For i = 1 To 26
    With frmVisualCalc.cmdUsrA(i)
      If i = 5 Then       '[E] and [&] are a special case
        If Bol Then
          If KeyShift Then
            .Caption = "E"
          Else
            .Caption = "e"
          End If
        Else
          .Caption = "&&"
        End If
      Else
        .Caption = Mid$(S, i + 1, 1)
      End If
      .Visible = True     'force visible
    End With
  Next i
End Sub

'*******************************************************************************
' Subroutine Name   : RedoAlphaPad
' Purpose           : Reset AlphaPad to user-defined key labels and visibility
'*******************************************************************************
Public Sub RedoAlphaPad()
  Dim i As Long
  Dim Actv As Integer
'
' reset display buttons to Tag contents if user-redefined
'
  With frmVisualCalc
    If .mnuKeypadFull.Checked Then                        'if full keyboard...
      For i = 0 To 26                                     'process SPACE to "Z"
        With .cmdUsrA(i)
          If CBool(ActivePgm) Then                        'if module program
            Actv = ActivePgm                              'save pgm ID, in case not invoked
            If CBool(SbrInvkIdx) Then Actv = SbrInvkStk(1).Pgm 'was invoked, so get root pgm
            .Caption = RTrim$(ModLbls(i + ModLblMap(Actv - 1)).lblName) '(0-26 + pgm base offset)
          Else
            .Caption = RTrim$(Lbls(i).lblName)            'set user-defined name from user program
          End If
          .Visible = Not Hidden(i)                        'reset visibility
        End With
      Next i
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : EnableNums
' Purpose           : Enable digits on the numeric keypad, based on type.
'                   : This is used to support Bin, Oct, Dec, and Hex key displays
'*******************************************************************************
Public Sub EnableNums()
  Dim Bol As Boolean
  Dim Idx As Long
      
  If RunMode Or CBool(MRunMode) Then Exit Sub
  If Not TextEntry Then
    With frmVisualCalc
      If .mnuKeypadFull.Checked Then
        If Key2nd Then
          .cmdKeyPad(26).Caption = "Hyp"    'reset key names
          .cmdKeyPad(39).Caption = "NOP"
          .cmdKeyPad(52).Caption = "Call"
          .cmdKeyPad(65).Caption = "Gto"
          .cmdKeyPad(78).Caption = "Rtn"
          .cmdKeyPad(91).Caption = "Stop"
        Else
          .cmdKeyPad(26).Caption = "Arc"    'reset key names
          .cmdKeyPad(39).Caption = "Dfn"
          .cmdKeyPad(52).Caption = "Sbr"
          .cmdKeyPad(65).Caption = "Lbl"
          .cmdKeyPad(78).Caption = "Ukey"
          .cmdKeyPad(91).Caption = "R/S"
        End If
      End If
      
      If BaseType = TypDec Then         'enable all keys if Decimal mode, except in BASIC and ADVANCED keypads
        For Idx = 1 To MaxKeys
          Select Case Idx
            Case 5, 6, 41, 42, 43, 56, 69, 82, 32, 33, 45, 46, 58, 59, 71, 72, 18, 77, 90, 103
              If .mnuKeypadBasic.Checked Then         'BASIC
                If Key2nd Then
                  .cmdKeyPad(Idx).Enabled = False
                Else
                  .cmdKeyPad(Idx).Enabled = True
                End If
              ElseIf .mnuKeypadAdvanced.Checked Then  'ADVANCED
                If Key2nd Then
                  If Idx = 41 Or Idx = 42 Or Idx = 5 Then      'Log10 & 10^ & CMs
                    .cmdKeyPad(Idx).Enabled = True
                  Else
                    .cmdKeyPad(Idx).Enabled = False
                  End If
                Else
                  .cmdKeyPad(Idx).Enabled = True
                End If
              Else                                    'FULL
                .cmdKeyPad(Idx).Enabled = True
              End If
            Case 95, 98 '=/ADV, >>/<<
              .cmdKeyPad(Idx).Enabled = True
            Case 26     'Hyp/Arc
              If .mnuKeypadBasic.Checked Then
                .cmdKeyPad(Idx).Enabled = False
              Else
                If Key2nd Then
                  .cmdKeyPad(Idx).Caption = "Hyp"
                  .cmdKeyPad(Idx).Enabled = True
                Else
                  .cmdKeyPad(Idx).Caption = "Arc"
                  .cmdKeyPad(Idx).Enabled = True
                End If
              End If
            Case 39, 52, 65, 78, 91 'A-F keys for Hex
              If Not .mnuKeypadFull.Checked Then
                .cmdKeyPad(Idx).Enabled = False
              Else
                .cmdKeyPad(Idx).Enabled = True
              End If
            Case 92, 79, 53, 54, 55, 66, 67, 68, 80, 81, 93, 94 '0-9,'., +/-
              If .mnuKeypadFull.Checked Then
                .cmdKeyPad(Idx).Enabled = True
              ElseIf Key2nd Then
                .cmdKeyPad(Idx).Enabled = False
              Else
                .cmdKeyPad(Idx).Enabled = True
              End If
            Case Else
              .cmdKeyPad(Idx).Enabled = True
          End Select
        Next Idx
      Else
        For Idx = 1 To MaxKeys
          Select Case Idx             'enable non-decimal command keys
            Case 41, 42, 43, 56, 69, 82, 32, 33, 45, 46, 58, 59, 71, 72, 85, 7
              If Key2nd Then
                .cmdKeyPad(Idx).Enabled = False
              Else
                .cmdKeyPad(Idx).Enabled = True
              End If
            Case 95, 98 '=/ADV, >>/<<
              .cmdKeyPad(Idx).Enabled = True
            Case 26     'Hyp/Arc
              If .mnuKeypadBasic.Checked Then
                If Key2nd Or BaseType <> TypHex Then
                  .cmdKeyPad(Idx).Enabled = False
                End If
              ElseIf .mnuKeypadAdvanced.Checked Then
                If Key2nd Then
                  .cmdKeyPad(Idx).Enabled = False
                ElseIf BaseType <> TypHex Then
                  .cmdKeyPad(Idx).Enabled = False
                End If
              Else
                If Key2nd Then
                  .cmdKeyPad(Idx).Enabled = False
                ElseIf BaseType <> TypHex Then
                  .cmdKeyPad(Idx).Enabled = False
                End If
              End If
            Case 5, 6 'CE/CMM, CLR/CP
              .cmdKeyPad(Idx).Enabled = True
            Case Else                 'disable all others
              .cmdKeyPad(Idx).Enabled = False
          End Select
        Next Idx
'
' now enable only numeric keys base allows
'
        Select Case BaseType
          Case TypBin
            For Idx = 53 To 94
              If Key2nd Then
                .cmdKeyPad(Idx).Enabled = False
              Else
                Select Case Idx
                  Case 92, 79                         '0-1 enabled
                    .cmdKeyPad(Idx).Enabled = True
                  Case 53, 54, 55, 66, 67, 68, 80, 81, 93, 94 '2-9, ., +/-
                    .cmdKeyPad(Idx).Enabled = False
                End Select
              End If
            Next Idx
            
          Case TypOct
            For Idx = 53 To 94
              If Key2nd Then
                .cmdKeyPad(Idx).Enabled = False
              Else
                Select Case Idx
                  Case 92, 79, 80, 81, 66, 67, 68, 53 '0-7 enabled
                    .cmdKeyPad(Idx).Enabled = True
                  Case 54, 55, 93, 94                 '8-9, ., +/-
                    .cmdKeyPad(Idx).Enabled = False
                End Select
              End If
            Next Idx
          
          Case TypHex
            For Idx = 53 To 94
              If Key2nd Then
                .cmdKeyPad(Idx).Enabled = False
              Else
                Select Case Idx
                  Case 92, 79, 53, 54, 55, 66, 67, 68, 80, 81 '0-9
                    .cmdKeyPad(Idx).Enabled = True
                  Case 93, 94  '., +/-
                    .cmdKeyPad(Idx).Enabled = False
                End Select
              End If
            Next Idx
            If Not Key2nd Then
              .cmdKeyPad(26).Caption = "A"  'allow A-F for Hex (10-15)
              .cmdKeyPad(39).Caption = "B"
              .cmdKeyPad(52).Caption = "C"
              .cmdKeyPad(65).Caption = "D"
              .cmdKeyPad(78).Caption = "E"
              .cmdKeyPad(91).Caption = "F"
              If Not TextEntry Then
                .cmdKeyPad(26).Enabled = True 'Hyp
                .cmdKeyPad(39).Enabled = True 'Dfn
                .cmdKeyPad(52).Enabled = True 'Sbr
                .cmdKeyPad(65).Enabled = True 'Lbl
                .cmdKeyPad(78).Enabled = True 'Ukey
                .cmdKeyPad(91).Enabled = True 'R/S
              End If
            End If
        End Select
      End If
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : SelectOnly
' Purpose           : Mark selection of only 1 line in display
'*******************************************************************************
Public Sub SelectOnly(Itm As Integer)
  Dim Lst() As Long
  Dim Idx As Integer, i As Integer
  
  Selecting = True
  Do
    With frmVisualCalc
      With .lstDisplay
        If .SelCount = 1 Then             '1 line already selected?
          If .ListIndex = Itm Then        'and target is current
            If .Selected(.ListIndex) Then 'and if it is already selected...
              Exit Do                      'then nothing to do
            End If
          End If
        End If
      End With
      
      IgnoreClick = True                'prevent listbox click event from resetting instr. pointer
      Call DeSelAllListBox(.lstDisplay) 'deselect all selected items
      With .lstDisplay
        .ListIndex = Itm                'mark listindex
        .Selected(Itm) = True           'mark selection
      End With
    End With
    IgnoreClick = False
    Exit Do
  Loop
  
  If LrnMode Then             'if LRN mode, get tip for selected instruction
    InstrPtr = Itm            'set instruction line
    Call UpdateStatus         'display status changes
    Call BuildPgmLine         'display code horizontally in status bar
  End If
  Selecting = False
End Sub

'*******************************************************************************
' Subroutine Name   : Sum_MDL
' Purpose           : Add active pgm to Module
'*******************************************************************************
Public Sub Sum_MDL()
  Dim Idx As Long, ModOffset As Long
  Dim S As String, Lbl As String
  '
  ' check for no code to save
  '
  If InstrCnt = 0 Then
    ForcError "No program code to add to module"
    Exit Sub
  End If
  '
  ' if program has changed, prompt the user
  '
  If IsDirty Then
    If CenterMsgBoxOnForm(frmVisualCalc, _
        "The program buffer has unsaved data." & vbCrLf & _
        "You may want to save it first." & vbCrLf & vbCrLf & _
        "Go ahead and add to module?", _
        vbYesNo Or vbQuestion Or vbDefaultButton2, _
        "Usaved Program Code") = vbNo Then Exit Sub
  End If
  '
  ' if no module name, prompt for one
  '
  If ModName = 0 Then
    S = Trim$(InputBox("Enter Module number (1-9999) to name this new Module:", "Enter Module ID", vbNullString))
    If Len(S) = 0 Then Exit Sub
    If Not IsNumeric(S) Then
      ForcError "Invalid Module Number. Must be digits 1-9999"
      Exit Sub
    End If
    
    Lbl = Trim$(InputBox("Enter Short Description of Module number to identify this new Module:", "Enter Module Description", vbNullString))
    If Len(Lbl) = 0 Then Exit Sub
  End If
  '
  ' Compress program, if not Compressed already
  '
  If Not Compressd Then Call Compress 'try Compressing if not
  If Not Compressd Then Exit Sub     'if still not, then error
  If ModName = 0 Then
    ModName = CInt(S)               'else save new module name
    ModLocked = False               'init not locked
    ModLbl = Lbl
  End If
  '
  ' ensure base set (should be, but better to be safe)
  '
  ReDim Preserve ModMap(ModCnt + 1)
  ReDim Preserve ModLblMap(ModCnt + 1)
  ReDim Preserve ModStMap(ModCnt + 1)
  ModMap(ModCnt) = ModSize
  ModLblMap(ModCnt) = ModLblCnt
  ModStMap(ModCnt) = ModStCnt
  ModOffset = ModStCnt                            'save starting base
  ModCnt = ModCnt + 1
  '
  ' add program code to module
  '
  Idx = ModSize + InstrCnt
  ReDim Preserve ModMem(Idx)
  For Idx = 0 To InstrCnt - 1
    ModMem(ModSize) = Instructions(Idx)
    ModSize = ModSize + 1
  Next Idx
  '
  ' add program labels
  '
  Idx = ModLblCnt + LblCnt                        'set new size
  ReDim Preserve ModLbls(Idx)                     'resize for data
  For Idx = 0 To LblCnt - 1                       'copy labels
    ModLbls(ModLblCnt) = Lbls(Idx)                'copy label
    If ModLbls(ModLblCnt).LblTyp = TypStruct Then 'update index if Struct
      ModLbls(ModLblCnt).LblValue = Lbls(Idx).LblValue + ModOffset
    End If
    ModLblCnt = ModLblCnt + 1
  Next Idx
  '
  ' add program structures
  '
  If CBool(StructCnt) Then
    Idx = ModStCnt + StructCnt
    ReDim Preserve ModStPl(Idx)
    For Idx = 1 To StructCnt
      CloneStruct ModStPl(ModStCnt), StructPl(Idx)
      ModStCnt = ModStCnt + 1
    Next Idx
  End If
  '
  ' mark upper bounds
  '
  ModMap(ModCnt) = ModSize
  ModLblMap(ModCnt) = ModLblCnt
  ModStMap(ModCnt) = ModStCnt
  '
  ' now clean up
  '
  Call CP_Support         'remove pgm code
  ActivePgm = ModCnt      'activate new library program
  Call RedoAlphaPad       'update keyboard
  DisplayReg = ModCnt     'display new module number
  DisplayText = False
  ModDirty = True
  Call UpdateStatus
  DisplayMsg "Module save program as Pgm" & Format(ModCnt, "00")
End Sub

'*******************************************************************************
' Subroutine Name   : Save_MDL
' Purpose           : Save Module to File
'*******************************************************************************
Public Sub Save_MDL()
  Dim Path As String, Nm As String
  Dim Idx As Long, Idy As Long
  Dim Fn As Integer
  '
  ' check storage path
  '
  If Len(StorePath) = 0 Then
    ForcError "No storage path defined"
    Exit Sub
  End If
  '
  ' build path to module
  '
  Nm = "MDL" & Format(ModName, "0000") & ".mdl"
  Path = RemoveSlash(StorePath) & "\MDL\" & Nm
  '
  ' see if we can delete it
  '
  If Fso.FileExists(Path) Then
    On Error Resume Next
    Fso.DeleteFile Path
    If CBool(Err.Number) Then
      ForcError "Cannot update " & Nm & ". Storage Path Read-Only?"
      Exit Sub
    End If
  End If
    
  Fn = FreeFile(0)
  Open Path For Binary Access Write As #Fn
  If CBool(Err.Number) Then
    ForcError "Cannot open for writing " & Nm & ". Storage Path Read-Only?"
    Exit Sub
  End If
  On Error GoTo 0
  '
  ' write header
  '
  Put #Fn, , ModCnt             '# modules
  Put #Fn, , ModSize            'total size of ModMem
  Put #Fn, , ModLblCnt          'total number of labels
  Put #Fn, , ModStCnt           'total number of structures
  Put #Fn, , ModLbl             'save description of module
  
  Put #Fn, , ModMap             'save memory maps
  Put #Fn, , ModLblMap
  If CBool(ModStCnt) Then
    Put #Fn, , ModStMap
  End If
  '
  ' write instruction tables
  '
  Put #Fn, , ModMem
  '
  ' write moduled locked flag
  '
  Put #Fn, , ModLocked          'saved locked flag
  '
  ' write labels
  '
  For Idx = 0 To ModLblCnt - 1
    With ModLbls(Idx)
      Put #Fn, , .lblAddr
      Put #Fn, , .lblCmt
      Put #Fn, , .LblDat
      Put #Fn, , .LblEnd
      Put #Fn, , .lblName
      Put #Fn, , .LblScope
      Put #Fn, , .LblTyp
      Put #Fn, , .lblUdef
      Put #Fn, , .LblValue
    End With
  Next Idx
  '
  ' write structures
  '
  If CBool(ModStCnt) Then
    For Idx = 0 To ModStCnt - 1
      With ModStPl(Idx)
        Put #Fn, , .StSiz
        Put #Fn, , .StItmCnt
        '
        ' write structure items for each structure
        '
        For Idy = 0 To .StItmCnt - 1
          With .StItems(Idy)
            Put #Fn, , .siLen
            Put #Fn, , .SiName
            Put #Fn, , .siOfst
            Put #Fn, , .siType
          End With
        Next Idy
      End With
    Next Idx
  End If
  Close #Fn
  '
  ' clean house
  '
  ModDirty = False              'module no longer dirty
  LastTypedInstr = 128
  Call Clear_Screen
  If ModLocked Then
    DisplayMsg "Saved Module " & Nm & " OK, and LOCKED"
  Else
    DisplayMsg "Saved Module " & Nm & " OK"
  End If
  SaveSetting App.Title, "Settings", "LoadedMDL", CStr(ModName)
End Sub

'*******************************************************************************
' Subroutine Name   : Load_MDL
' Purpose           : Load Module from File
'*******************************************************************************
Public Sub Load_MDL()
  Dim Path As String, Nm As String
  Dim Lbl As String * DisplayWidth, S As String
  Dim Idx As Long, Idy As Long
  Dim Fn As Integer
  '
  ' check storage path
  '
  If Len(StorePath) = 0 Then
    ForcError "No storage path defined"
    Exit Sub
  End If
  '
  ' build path to module
  '
  Nm = "MDL" & Format(DisplayReg, "0000") & ".mdl"
  Path = RemoveSlash(StorePath) & "\MDL\" & Nm
  '
  ' see if we can load it
  '
  If Not Fso.FileExists(Path) Then
    ForcError "Cannot find " & Nm & " in current storage path: " & StorePath
    Exit Sub
  End If
  
  Fn = FreeFile(0)
  Open Path For Binary Access Read As #Fn
  If CBool(Err.Number) Then
    ForcError "Cannot open for writing " & Nm & ". Storage Path Read-Only?"
    Exit Sub
  End If
  On Error GoTo 0
  '
  ' write header
  '
  Get #Fn, , ModCnt             '# modules
  Get #Fn, , ModSize            'total size of ModMem
  Get #Fn, , ModLblCnt          'total number of labels
  Get #Fn, , ModStCnt           'total number of structures
  Get #Fn, , Lbl                'descriptive label for module
  
  ReDim ModMap(ModCnt)          'dimension buffers to receive data
  ReDim ModLblMap(ModCnt)       'allow 1 more due to uppbounds checks
  ReDim ModStMap(ModCnt)
  ReDim ModMem(ModSize)
  ReDim ModLbls(ModLblCnt - 1)  'label definitions
  If CBool(ModStCnt) Then
    ReDim ModStPl(ModStCnt - 1) 'structure definitions
  End If
  
  Get #Fn, , ModMap             'load memory map for program starts
  Get #Fn, , ModLblMap          'load memory map for start of Program label definitions
  If CBool(ModStCnt) Then
    Get #Fn, , ModStMap           'load memory map for structure defs for each program in module
  End If
  '
  ' read instruction tables
  '
  Get #Fn, , ModMem
  '
  ' read moduled locked flag
  '
  Get #Fn, , ModLocked
  '
  ' read labels
  '
  For Idx = 0 To ModLblCnt - 1
    With ModLbls(Idx)
      Get #Fn, , .lblAddr             'get definition address
      Get #Fn, , .lblCmt              'get comment for Ukeys, data for Const
      Get #Fn, , .LblDat              'address of item's block data
      Get #Fn, , .LblEnd              'address of end of block
      Get #Fn, , .lblName             'item's name
      Get #Fn, , .LblScope            'item's public/private scope
      Get #Fn, , .LblTyp              'item's type
      Get #Fn, , .lblUdef             'set if item user-defined
      Get #Fn, , .LblValue            'Enum value, or Structure Index
    End With
  Next Idx
  '
  ' write structures
  '
  If CBool(ModStCnt) Then
    For Idx = 0 To ModStCnt - 1
      With ModStPl(Idx)
        Get #Fn, , .StSiz             'get byte size of structure
        Get #Fn, , .StItmCnt          'get # of Structure Items
        .StBuf = String$(.StSiz, 0)   'make buffer size of data
        ReDim .StItems(.StItmCnt - 1) 'make room for Structure Items
        '
        ' read structure items for each structure
        '
        For Idy = 0 To .StItmCnt - 1
          With .StItems(Idy)
            Get #Fn, , .siLen
            Get #Fn, , .SiName
            Get #Fn, , .siOfst
            Get #Fn, , .siType
          End With
        Next Idy
      End With
    Next Idx
  End If
  Close #Fn             'close file
  '
  ' do house cleaning
  '
  ModDirty = False      'module no longer dirty
  Idx = ActivePgm
  If CBool(ActivePgm) Then
    ActivePgm = 0
    RunMode = False       'turn off run modes
    MRunMode = 0
    ModPrep = 0
    BraceIdx = 0          'reset braceing index us user key hit
    SbrInvkIdx = 0        'reset subr stack if user-key hit
  End If
  ModName = CInt(DisplayReg)
  ModLbl = Lbl
  DisplayReg = 0#
  Call UpdateStatus       'update statusbar messages
  If CBool(Idx) Or Not RunMode Then
    Call RedoAlphaPad     'reset keyboard titles
  End If
  Call ResetPndAll
  LastTypedInstr = 128
  DisplayText = False
  SaveSetting App.Title, "Settings", "LoadedMDL", CStr(ModName)
  S = Trim$(Lbl)          'grab module label
  If CBool(Len(S)) Then   'if label exists
    frmVisualCalc.sbrImmediate.Panels("MDL").ToolTipText = "Currently loaded Module: " & S
  End If
  If Not RunMode Then
    DisplayMsg "Loaded Module " & Nm & " OK"
    If CBool(Len(S)) Then DisplayMsg "<" & S & ">"
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : CloneStruct
' Purpose           : Clone structure
'*******************************************************************************
Public Sub CloneStruct(dst As StructPool, Src As StructPool)
  Dim Idx As Integer
  
  dst.StBuf = Src.StBuf
  dst.StItmCnt = Src.StItmCnt
  dst.StSiz = Src.StSiz
  ReDim dst.StItems(dst.StItmCnt)
  For Idx = 0 To dst.StItmCnt - 1
    dst.StItems(Idx).siLen = Src.StItems(Idx).siLen
    dst.StItems(Idx).SiName = Src.StItems(Idx).SiName
    dst.StItems(Idx).siOfst = Src.StItems(Idx).siOfst
    dst.StItems(Idx).siType = Src.StItems(Idx).siType
  Next Idx
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


