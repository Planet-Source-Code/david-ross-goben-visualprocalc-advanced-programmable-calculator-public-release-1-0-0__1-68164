Attribute VB_Name = "modLRNSupport"
Option Explicit

'*******************************************************************************
' Subroutine Name   : AddInstruction
' Purpose           : Add an instruction to the LRN mode
'*******************************************************************************
Public Sub AddInstruction(ByVal Inst As Integer)
  Dim S As String, C As String
  Dim Idx As Long
  
  Select Case Inst
    Case 128                                                'null code
      Exit Sub                                              'ignore it
    
    Case Is > 900                                           'user-defined key
      S = "<Ukey_" & Chr$(Inst - 836) & ">"                 'handle User-defined keys (A=900; 900-64=836)
  
    Case 0 To 9                                             'digits 0-9
      S = Chr$(Inst + 48)                                   'get plain ascii code
      If TextEntry Then                                     'if text entry, then count text
        DspTxt = DspTxt & S                                 'append new character to text
        CharCount = Len(DspTxt)                             'characters entered
        If VarLbl Then
          S = "[" & S & "]"                                 'get embraced ascii code if inside label
        End If
      End If
          
    Case Else
      If Inst < 128 Then
        If Upcase Then
          C = UCase$(Chr$(Inst))                            'handle ASCII characters
          Inst = Asc(C)                                     'reset instruction
          S = "[" & C & "]"                                 'get representation
        Else
          S = "[" & Chr$(Inst) & "]"                        'handle ASCII characters
        End If
        If TextEntry Then                                   'if text entry, count text
          DspTxt = DspTxt & Chr$(Inst)                      'append text
          CharCount = Len(DspTxt)                           'characters entered
        End If
      Else
        S = GetInst(Inst)                                   'handle key commands
        If TextEntry Then                                   'if text entry, turn it off
          Call ResetPnd                                     'reset pending operations
          AllowSpace = False                                'do not allow typing of a space
          Call frmVisualCalc.checkTextEntry(False)          'set up keyboard
        End If
      End If
  End Select
  
  S = Format(InstrPtr, "0000  ") & Format(Inst, "000   ") & S  'format line
  
  With frmVisualCalc
    LockControlRepaint .lstDisplay
    With .lstDisplay
      If INSmode And InstrPtr < InstrCnt Then               'if INS mode active...
        For Idx = InstrCnt - 1 To InstrPtr Step -1          'move instructions up
          .List(Idx + 1) = Format(Idx + 1, "0000") & Mid$(.List(Idx), 5)
          Instructions(Idx + 1) = Instructions(Idx)
        Next Idx
        InstrCnt = InstrCnt + 1                             'bump number of available instructions
        If InstrCnt > InstrSize Then                        'if it excedes the size of the pool
          InstrSize = InstrSize + InstrInc                  'bump pool size
          ReDim Preserve Instructions(InstrSize)            'bump pool
        End If
        Instructions(InstrCnt) = 0                          'nullify top location
        .AddItem Format(InstrCnt, "0000  \0\0\0")
      End If
      
      .List(InstrPtr) = S                                   'stuff new instruction
      Instructions(InstrPtr) = Inst                         'stuff instruction
      
      If InstrPtr = InstrCnt Then                           'add new instruction to top
        .AddItem Format(InstrPtr + 1, "0000  \0\0\0")
        InstrCnt = InstrCnt + 1                             'bump number of available instructions
        If InstrCnt > InstrSize Then                        'if it excedes the size of the pool
          InstrSize = InstrSize + InstrInc                  'bump pool size
          ReDim Preserve Instructions(InstrSize)            'bump pool
        End If
        Instructions(InstrCnt) = 0                          'nullify top location
      End If
      InstrPtr = InstrPtr + 1
      Call SelectOnly(InstrPtr)                             'select only this item
    End With
    UnlockControlRepaint .lstDisplay
    
    Preprocessd = False                                     'disable compilation check
    Compressd = False
    IsDirty = True                                          'indicate we are dirty
    
    Call UpdateStatus                                       'update status bar
  End With
  If frmCDLoaded Then
    With frmCoDisplay
      If .Toolbar1.Buttons("AutoUpdate").Value = tbrPressed Then
        SetUpCoDisplay
      End If
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : CutInstruction
' Purpose           : Cut instruction from program
'*******************************************************************************
Public Sub CutInstruction()
  Dim S As String
  Dim Idx As Integer, i As Integer, j As Integer
  Dim ListData() As Long
  
  With frmVisualCalc
    LockControlRepaint .lstDisplay
    With .lstDisplay
      If CBool(.SelCount) Then                              'anything to copy?
        Preprocessd = False                                 'disable compilation check
        Compressd = False
        ListCnt = .SelCount                                 'yes, save count
        ReDim ListText(ListCnt - 1) As String               'set storage space
        ListData = GetSelListBox(frmVisualCalc.lstDisplay)  'grab list of indexes
        j = UBound(ListData)                                'keep a copy of the array size
        
        For Idx = j To 0 Step -1                            'move backward through the list
          i = CInt(ListData(Idx))                           'keep copy of index
          Instructions(i) = 127                             'Del place-holder
          If i = .ListCount - 1 Then                        'if top of list...
            ListText(Idx) = vbNullString                    'then ignore it
          Else
            ListText(Idx) = .List(i)                        'save copy of line (in reverse order)
            .RemoveItem i                                   'now remove the item
          End If
        Next Idx
        
        For Idx = i To .ListCount - 1                       'now renumber everything forward
          .List(Idx) = Format(Idx, "0000") & Mid$(.List(Idx), 5)
        Next Idx
        
        j = i                                               'copy pointer to J
        For Idx = i To InstrCnt - 1                         'strip Del instructions
          Instructions(j) = Instructions(Idx)               'copy an instruction
          If Instructions(j) <> 127 Then j = j + 1          'if J points to 127, do not inc
        Next Idx
        Instructions(j) = 0                                 'null top-most
        InstrCnt = j                                        'set number of instructions
        Call SelectOnly(i)                                  'select only this item
      End If
    End With
    IsDirty = CBool(InstrCnt)                               'indicate we are dirty if something left
    UnlockControlRepaint .lstDisplay
  End With
  Call UpdateStatus
  
  If frmCDLoaded Then
    With frmCoDisplay
      If .Toolbar1.Buttons("AutoUpdate").Value = tbrPressed Then
        SetUpCoDisplay
      End If
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : CopyInstruction
' Purpose           : Copy instruction from program
'*******************************************************************************
Public Sub CopyInstruction()
  Dim Idx As Integer, i As Integer
  Dim ListData() As Long
  
  With frmVisualCalc.lstDisplay
    If CBool(.SelCount) Then                              'anything to copy?
      ListCnt = .SelCount                                 'yes, save count
      ReDim ListText(ListCnt - 1) As String               'set storage space
      ListData = GetSelListBox(frmVisualCalc.lstDisplay)  'grab list of indexes
      For Idx = UBound(ListData) To 0 Step -1
        i = CInt(ListData(Idx))                           'keep copy of index
        If i = .ListCount - 1 Then                        'if top of list...
          ListText(Idx) = vbNullString                    'then ignore it
        Else
          ListText(Idx) = .List(i)                        'stuff instructions to list (reverse order)
        End If
      Next Idx
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : PasteInstruction
' Purpose           : Paste Instructions into program
'*******************************************************************************
Public Sub PasteInstruction()
  Dim S As String
  Dim Idx As Integer, i As Integer
  Dim ListData() As Long
  
  If CBool(ListCnt) Then                                  'if data to paste
    Preprocessd = False                                   'disable compilation check
    Compressd = False
    IsDirty = True                                        'indicate we are dirty
    With frmVisualCalc
      LockControlRepaint .lstDisplay                      'lock repaints
      With .lstDisplay
        i = .ListIndex                                    'get current line
        For Idx = UBound(ListText) To 0 Step -1           'process commands to insert
          If CBool(Len(ListText(Idx))) Then               'if command is valid
            .AddItem ListText(Idx), i                     'insert it
          End If
        Next Idx
        
        For Idx = i To .ListCount - 1                     'now renumber everything forward
          .List(Idx) = Format(Idx, "0000") & Mid$(.List(Idx), 5)
        Next Idx
        
        InstrCnt = .ListCount                             'add number of instructions to paste
        If InstrCnt > InstrSize Then                      'bump instruction pool if needed
          Do While InstrCnt > InstrSize
            InstrSize = InstrSize + 100                   'bump by 100s
          Loop
          ReDim Preserve Instructions(InstrSize)          'resize it
        End If
        
        For Idx = i To InstrCnt - 1                       'copy updated instruction list
          Instructions(Idx) = CInt(Mid$(.List(Idx), 7, 3))
        Next Idx
        
        Call SelectOnly(i)                                'select only this item
      End With
      UnlockControlRepaint .lstDisplay
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : DeleteInstruction
' Purpose           : Delete instruction from program
'*******************************************************************************
Public Sub DeleteInstruction()
  Call CutInstruction   'cut data
  ListCnt = 0           'then erase storage for cut data
  Erase ListText
End Sub

'*******************************************************************************
' Subroutine Name   : BuildInstrList
' Purpose           : Build the instruction list
'*******************************************************************************
Public Sub BuildInstrList()
  Dim Idx As Integer, Inst As Integer
  
  LockControlRepaint frmVisualCalc.lstDisplay                 'lock display refreshes
  '
  ' erase display, then format it
  '
  With frmVisualCalc.lstDisplay                               'we will update this window
    .Clear                                                    'clear list
'
'BASIC Display
'
    For Idx = 0 To InstrCnt - 1                           'now cycle through each instruction
      Inst = Instructions(Idx)                            'get instruction
      .AddItem Format(Idx, "0000  ") & Format(Inst, "000   ") & GetInst(Inst)
    Next Idx
    .AddItem Format(Idx, "0000  \0\0\0")
    Call SelectOnly(InstrPtr)                             'select only this item
'
' position the insert line to the middle
'
    Idx = .ListIndex - DisplayHeight / 2  'up 1/2 screen height
    If Idx < 0 Then Idx = 0               'if we are too low in the instruction list
    .TopIndex = Idx                       'set top display line
  End With
'
' unlock display
'
  UnlockControlRepaint frmVisualCalc.lstDisplay               'reset display
End Sub

'*******************************************************************************
' Function Name     : BuildInstrArray
' Purpose           : Build an array of the instruction list
'*******************************************************************************
Public Function BuildInstrArray() As String()
  Dim Idx As Integer, Inst As Integer
  Dim Ary() As String
  
  If InstrCnt = 0 Then                      'nothing to process?
    ForcError "No LEARNED code exists"
    Exit Function
  End If
  
  If CBool(LRNstyle) Then
    If Not Preprocessd Then                 'Preprocessed?
      Call Preprocess
      If Not Preprocessd Then Exit Function
    End If
  End If
  
  Select Case LRNstyle
    Case 0  'BASIC Display
      ReDim Ary(InstrCnt - 1)               'size the array
      For Idx = 0 To InstrCnt - 1           'now cycle through each instruction
        Ary(Idx) = "  " & GetInst(Instructions(Idx)) 'grab data for instruction
      Next Idx
      BuildInstrArray = Ary                 'return data array
    Case 1  'Formatted listing
      BuildInstrArray = InstFmt             'return formatted array
    Case 2  'Formatted listing with program steps
      ReDim Ary(InstCnt - 1)                'size the formatted array
      For Idx = 0 To InstCnt - 1
        Ary(Idx) = Format(InstMap(Idx), "0000  ") & InstFmt(Idx)
      Next Idx
      BuildInstrArray = Ary                 'return data array
    Case Else 'basic listing with text separation
      BuildInstrArray = InstFmt3            'return formatted array
  End Select
End Function

'*******************************************************************************
' Function Name     : GetInst
' Purpose           : Return text for instruction
'*******************************************************************************
Public Function GetInst(ByVal Inst As Integer) As String
  Dim S As String
  
  If Inst > 31 And Inst < 128 Then
    If Preprocessing Then
      S = Chr$(Inst)
    Else
      S = "[" & Chr$(Inst) & "]"
    End If
  Else
    Select Case Inst
      Case 128: S = vbNullString
      Case 0 To 9: S = Chr$(Inst + 48)
      Case 129: S = "LRN"
      Case 130: S = "Pgm"
      Case 131: S = "Load"
      Case 132: S = "Save"
      Case 133: S = "CE"
      Case 134: S = "CLR"
      Case 135: S = "OP"
      Case 136: S = "SST"
      Case 137: S = "INS"
      Case 138: S = "Cut"
      Case 139: S = "Copy"
      Case 140: S = "PtoR"
      Case 141: S = "STO"
      Case 142: S = "RCL"
      Case 143: S = "EXC"
      Case 144: S = "SUM"
      Case 145: S = "MUL"
      Case 146: S = "IND"
      Case 147: S = "Reset"
      Case 148: S = "Hkey"
      Case 149: S = "lnX"
      Case 150: S = "E+"
      Case 151: S = "Mean"
      Case 152: S = "X!"
      Case 153: S = "X><T"
      Case iHyp: S = "Hyp"
      Case 155: S = "Sin"
      Case 156: S = "Cos"
      Case 157: S = "Tan"
      Case 158: S = "1/X"
      Case 159: S = "Txt"
      Case 160: S = "Hex"
      Case 161: S = "&"
      Case 162: S = "StFlg"
      Case 163: S = "IfFlg"
      Case 164: S = "X==T"
      Case 165: S = "X>=T"
      Case 166: S = "X>T"
      Case 167: S = "Dfn"
      Case iColon: S = ":"
      Case 169: S = "("
      Case 170: S = ")"
      Case 171: S = "÷"
      Case iStyle: S = "Style"
      Case 173: S = "Dec"
      Case 174: S = "|"
      Case 175: S = "Int"
      Case 176: S = "Abs"
      Case 177: S = "Fix"
      Case 178: S = "D.MS"
      Case 179: S = "EE"
      Case 180: S = "Sbr"
      Case 184: S = "x"
      Case iRem: S = "Rem"
      Case 186: S = "Oct"
      Case 187: S = "~"
      Case 188: S = "Select"
      Case 189: S = "Case"
      Case 190: S = "{"
      Case 191: S = "}"
      Case 192: S = "Deg"
      Case 193: S = "Lbl"
      Case 197: S = "-"
      Case 198: S = "Beep"
      Case 199: S = "Bin"
      Case 200: S = "^"
      Case 201: S = "For"
      Case 202: S = "Do"
      Case 203: S = "While"
      Case 204: S = "Pmt"
      Case 205: S = "Rad"
      Case 206: S = "Ukey"
      Case 210: S = "+"
      Case 211: S = "Plot"
      Case 212: S = "Nvar"
      Case 213: S = "%"
      Case 214: S = "If"
      Case 215: S = "Else"
      Case 216: S = "Cont"
      Case 217: S = "Break"
      Case 218: S = "Grad"
      Case 219: S = "R/S"
      Case 221: S = "."
      Case 222: S = "+/-"
      Case 223: S = "="
      Case 224: S = "Print"
      Case 225: S = "Tvar"
      Case 226: S = ">>"
      Case 227: S = "y^"
      Case 228: S = "X²"
      Case 229: S = "Pi"
      Case 230: S = "Rnd"
      Case 231: S = "Mil"
      Case 232: S = "Pvt"
      Case 233: S = "Const"
      Case 234: S = "Struct"
      Case 235: S = "NxLbl"
      Case 236: S = "PvLbl"
      Case 237: S = "Line"
      Case 238: S = "["
      Case 239: S = "]"
      Case 240: S = "ClrVar"
      Case 241: S = "SzOf"
      Case 242: S = "Def"
      Case 243: S = "IfDef"
      Case 244: S = "Edef"
  '----------------------
  ' EXTENDED FUNCTIONS
  '----------------------
      Case 245: S = "STO IND"
      Case 246: S = "RCL IND"
      Case 247: S = "EXC IND"
      Case 248: S = "SUM IND"
      Case 249: S = "MUL IND"
      Case 250: S = "SUB IND"
      Case 251: S = "DIV IND"
      Case 252: S = "GTO IND"
      Case 254: S = "OP IND"
      Case 255: S = "FIX IND"
      Case 256: S = "PGM IND"
  '---------------------------------------------------------------------
                '--2nd keys---------------------------------------------
  '---------------------------------------------------------------------
      Case 257: S = "MDL"
      Case 258: S = "CMM"
      Case 259: S = "Lapp"
      Case 260: S = "ASCII"
      Case 261: S = "CMs"
      Case 262: S = "CP"
      Case iList: S = "List"
      Case 264: S = "BST"
      Case 265: S = "DEL"
      Case 266: S = "Paste"
      Case iUSR: S = "USR"
      Case 268: S = "RtoP"
      Case 269: S = "Push"
      Case 270: S = "Pop"
      Case 271: S = "StkEx"
      Case 272: S = "SUB"
      Case 273: S = "DIV"
      Case 274: S = "<"
      Case 275: S = ">"
      Case 276: S = "Skey"
      Case 277: S = "eX"
      Case 278: S = "E-"
      Case 279: S = "StDev"
      Case 280: S = "Varnc"
      Case 281: S = "Yint"
      Case iArc: S = "Arc"
      Case 283: S = "Sec"
      Case 284: S = "Csc"
      Case 285: S = "Cot"
      Case 286: S = "LogX"
      Case 287: S = "Var"
      Case 288: S = "=="
      Case 289: S = "&&"
      Case 290: S = "RFlg"
      Case 291: S = "!Flg"
      Case 292: S = "X!=T"
      Case 293: S = "X<=T"
      Case 294: S = "X<T"
      Case 295: S = " " ' NOP
      Case iSemiC: S = ";"
      Case 297: S = "Log"
      Case 298: S = "10^"
      Case 299: S = "÷="
      Case 300: S = "Fmt"
      Case 301: S = "!="
      Case 302: S = "||"
      Case 303: S = "Frac"
      Case 304: S = "Sgn"
      Case 305: S = "!Fix"
      Case 306: S = "D.ddd"
      Case 307: S = "!EE"
      Case 308: S = "Call"
      Case 309: S = "Trim"
      Case 310: S = "LTrim"
      Case 311: S = "RTrim"
      Case 312: S = "x="
      Case iRem2: S = "'"
      Case 314: S = ">="
      Case 315: S = "!"
      Case 316: S = "Open"
      Case 317: S = "Close"
      Case 318: S = "Read"
      Case 319: S = "Write"
      Case 320: S = "Swap"
      Case 321: S = "GTO"
      Case 322: S = "LOF"
      Case 323: S = "Get"
      Case 324: S = "Put"
      Case 325: S = "-="
      Case 326: S = "sysBP"
      Case 327: S = "<="
      Case 328: S = "Nor"
      Case 329: S = "Incr"
      Case 330: S = "Decr"
      Case 331: S = "Dsz"
      Case 332: S = "Dsnz"
      Case 333: S = "All"
      Case 334: S = "Rtn"
      Case 335: S = "LSet"
      Case 336: S = "RSet"
      Case 337: S = "Printf"
      Case 338: S = "+="
      Case 339: S = "RGB"
      Case 340: S = "Ivar"
      Case 341: S = "\"
      Case 342: S = "As"
      Case 343: S = "ElseIf"
      Case 344: S = "DBG"
      Case 345: S = "Gfree"
      Case 346: S = "Len"
      Case 347: S = "Stop"
      Case 348: S = "With"
      Case 349: S = ","
      Case 350: S = "Val"
      Case 351: S = "Adv"
      Case 352: S = "Print;"
      Case 353: S = "Cvar"
      Case 354: S = "<<"
      Case 355: S = "Root"
      Case 356: S = "Sqrt"
      Case 357: S = "e"
      Case 358: S = "Rnd#"
      Case 359: S = "Until"
      Case 360: S = "Pub"
      Case 361: S = "Enum"
      Case 362: S = "AdrOf"
      Case 363: S = "Pcmp"
      Case 364: S = "Comp"
      Case 365: S = "Circle"
      Case 366: S = "Split"
      Case 367: S = "Join"
      Case 368: S = "ReDim"
      Case 369: S = "Mid"
      Case 370: S = "Udef"
      Case 371: S = "!Def"
      Case 372: S = "Delse"
  '----------------------
  ' EXTENDED FUNCTIONS
  '----------------------
      Case 373: S = "StFlg IND"
      Case 374: S = "RFlg IND"
      Case 375: S = "IfFlg IND"
      Case 379: S = "!Flg IND"
      Case 388: S = "Dsz IND"
      Case 391: S = "Dsnz IND"
      Case 394: S = "Incr IND"
      Case 395: S = "Decr IND"
      Case 400: S = "Asin"
      Case 401: S = "Acos"
      Case 402: S = "Atan"
      Case 403: S = "SinH"
      Case 404: S = "CosH"
      Case 405: S = "TanH"
      Case 406: S = "ArcSinH"
      Case 407: S = "ArcCosH"
      Case 408: S = "ArcTanH"
      Case 409: S = "ASec"
      Case 410: S = "ACsc"
      Case 411: S = "Acot"
      Case 412: S = "SecH"
      Case 413: S = "CscH"
      Case 414: S = "CotH"
      Case 415: S = "ArcSecH"
      Case 416: S = "ArcCscH"
      Case 417: S = "ArcCotH"
      
      Case 433: S = "Var IND"
'
' special Compressr functions
'
      Case 485: S = ";"         'Special ';' for FOR statements
      Case 486: S = "Else"      'Special Else for Case
      Case 487 To 494: S = ")"  'special end parens
      Case 495 To 508: S = "}"  'special end braces
      Case 497: S = "CnstNxt"
      Case 498: S = "DblNxt"
      Case 499: S = "IntNxt"
      
      Case Is > 900
        S = "<Ukey_" & Chr$(Inst - 836) & ">"                 'handle User-defined keys
      
      Case Else: S = vbNullString
    End Select
  End If
  GetInst = S
End Function

'*******************************************************************************
' Subroutine Name   : Backstep
' Purpose           : Support the BST (Backstep) command
'*******************************************************************************
Public Sub Backstep()
  If InstrPtr = 0 Then Exit Sub                 'if we cannot backstep
  '
  ' now see if the new location is text or numeric
  '
  Do
    InstrPtr = InstrPtr - 1                       'back up instruction
    Select Case Instructions(InstrPtr)
      Case 146, 295, 221, 222, 179, Is < 129      'IND, null, NOP, 0 - 9, [.],
                                                  '[+/-], EE, or Text data
      Case iColon, iSemiC, iSemiColon, iLbrkt, iLparen, iLCbrace '[:],[;],[,(, or {
      Case iUparen, iIparen, iWparen, iFparen, iCparen, _
           iEparen, iSparen, iDWparen, iRparen
      Case iBCBrace, iSCBrace, iICBrace, iDCBrace, iWCBrace, _
           iFCBrace, iCCBrace, iDWBrace, iDUBrace, iEIBrace, _
           iENBrace, iSTBrace, iSIBrace, iCNBrace, iRCbrace
      Case iRbrkt, iRem, iRem2, iComma            '], REM, ['], or [,]
      Case Else
        Exit Do                                   'done if none of the above
    End Select
    If InstrPtr = 0 Then Exit Do                  'If at start, do not back up
  Loop
  Call UpdateStatus                               'update instruction pointer is display
End Sub

'*******************************************************************************
' Function Name     : LrnBST
' Purpose           : Employ LRN Mode BST command
'*******************************************************************************
Public Sub LrnBST()
  Dim Idx As Integer
  
  If InstrPtr = 0 Then Exit Sub           'if nothing to do, then exit
  InstrPtr = InstrPtr - 1                 'else back up 1 step
  Call SelectOnly(InstrPtr)               'select only this item
  Idx = InstrPtr - DisplayHeight / 2      'center line
  If Idx < 0 Then Idx = 0
  frmVisualCalc.lstDisplay.TopIndex = Idx
End Sub

'*******************************************************************************
' Subroutine Name   : NxtPrvLabel
' Purpose           : Support scrolling to next or previous instruction
'*******************************************************************************
Public Sub NxtPrvLabel(ByVal Forward As Boolean)
  Dim Idx As Integer, FmV As Integer, ToV As Integer, IncV As Integer
  
  Idx = InstrPtr                              'save instruction pointer
  If Not Preprocessd Then                     'if not Preprocessed...
    Call Preprocess                           'proprocess it
    If Not Preprocessd Then Exit Sub          'exit if errors
    LrnDsp = GetDisplayText()                 'save display data
    With frmVisualCalc
      With .lstDisplay
        LrnSav = .ListIndex                   'save selected line
        LrnTop = .TopIndex                    'save top displayed line
      End With
      Call BuildInstrList                     'build instruction list
      .lblLoc.BackStyle = 1                   'show learn mode headers
      .lblCode.BackStyle = 1
      .lblInstr.BackStyle = 1
      Call DspBackground
    End With
  End If
  InstrPtr = Idx                              'reset instruction pointer
  Call UpdateStatus
  
  If Forward Then                             'if going for NEXT label...
    FmV = InstrPtr + 1                        'init to preset+1
    If FmV >= InstrCnt Then Exit Sub          'too far, so exit
    ToV = InstrCnt - 1                        'to last step
    IncV = 1                                  'forward
  Else                                        'else going for PREVIOUS label...
    FmV = InstrPtr - 1                        'start at present-1
    If FmV < 0 Then Exit Sub                'too far, so exit
    ToV = 0                                   'go to start
    IncV = -1                                 'backward
  End If
  
  For Idx = FmV To ToV Step IncV              'scan range
    Select Case Instructions(Idx)             'check token
      Case iLbl, iSbr, iUkey, iConst, iStruct, iEnum, iNvar, iTvar, iIvar, iCvar
        Exit For                              'found it, so process
     End Select
  Next Idx
  
  If Idx = InstrCnt Or Idx = -1 Then Exit Sub 'if beyond data, then do nothing
  InstrPtr = Idx                              'set instruction pointyer
  
  Call SelectOnly(InstrPtr)                   'select line
  ToV = InstrPtr - DisplayHeight / 2          'center selection
  If ToV < 0 Then ToV = 0                     'went too low
  frmVisualCalc.lstDisplay.TopIndex = ToV     'set top of listbox
  Call UpdateStatus                           'reflect changes in the status bar
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

