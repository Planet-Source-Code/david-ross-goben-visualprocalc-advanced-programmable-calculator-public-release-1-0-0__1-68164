Attribute VB_Name = "modImport"
Option Explicit
'******************************************************************************
' Import a file
'******************************************************************************
Private iAry() As Integer 'temporary instruction list
Private iSize As Integer  'size of list
Private Iptr As Integer   'pointer into list
Private ErrorV As Boolean 'error flag

'*******************************************************************************
' Subroutine Name   : ImportFile
' Purpose           : Import a file. MAIN ROUTINE
'*******************************************************************************
Public Sub ImportFile(Filepath As String, ByVal ClipBrd As Boolean, ByVal PgmClip As Boolean)
  Dim Ary() As String, S As String, T As String
  Dim Idx As Long, i As Long, j As Long
  Dim Ary2() As Integer, HldIP As Integer
  Dim Hld2nd As Boolean
  
  Ary = ScrubFile(Filepath, ClipBrd)        'initialize file
  If ErrorV Then Exit Sub                   'errors
  iSize = SizeInst                          'init temp instruction pool
  ReDim iAry(SizeInst)
  Iptr = 0
  
  For Idx = 0 To UBound(Ary)
    S = Ary(Idx)                            'get an entry
    'Debug.Print S
    If CBool(Len(S)) Then                   'contains data?
      If Left$(S, 1) = """" Then            'yes, is it text?
        For i = 2 To Len(S) - 1             'yes, so apply ascii, less quotes
          AddInst CInt(Asc(Mid$(S, i, 1)))
        Next i
      Else
        Do                                  'proces non-text data
          If IsNumeric(Left$(S, 1)) And IsNumeric(Right$(S, 1)) Then
            Do While IsNumeric(Left$(S, 1))   'if numeric, send code 1 at a time
              AddInst CInt(Left$(S, 1))
              S = Mid$(S, 2)                  'strip spent digit
              If Len(S) = 0 Then Exit Do      'nothing more to do
            Loop
          End If
          For i = Len(S) To 1 Step -1       'if S contains data still...
            j = ImpCmd(Left$(S, i))         'try obtaining a command
            If j <> 128 Then                'found one?
              AddInst j                     'yes, so add it to list
              S = Mid$(S, i + 1)            'get remainder
              Exit For                      'done with loop
            End If
          Next i
          If Len(S) = 0 Then Exit Do        'anything more to do?
          If j = 128 Then
            ForcError "Import Error: " & S
            Exit Sub
          End If
        Loop
      End If
    End If
  Next Idx
'
' all imported, no clean up and provide it to the user
'
  If PgmClip Then                           'if importing a program clip
    Idx = InstrCnt + Iptr - 1               'set total size
    i = SizeInst                            'init buffer size
    Do While i < Idx
      i = i + InstrInc                      'bumpo by increment
    Loop
    ReDim Ary2(i)                           'size destination array
    For i = 0 To InstrPtr - 1
      Ary2(i) = Instructions(i)             'gram main pgm until InstrPtr-1
    Next i
    For Idx = 0 To Iptr - 1
      Ary2(i) = iAry(Idx)                   'add imported commands
      i = i + 1
    Next Idx
    For Idx = InstrPtr To InstrCnt - 1     'now add rest of main program
      Ary2(i) = Instructions(Idx)
      i = i + 1
    Next Idx
    Instructions = Ary2                     'transfer new instructions
    InstrCnt = i                            'set instruction count
    Erase Ary2
  Else
    Call CP_Support                         'erase main program space
    Instructions = iAry                     'transfer new instructions
    InstrCnt = Iptr                         'set instruction count
  End If
  Erase iAry                                'kill old buffer
  IsDirty = True
  Preprocessd = False
  Compressd = False
  frmVisualCalc.mnuWinASCII.Enabled = True  'allow viewing list
  Hld2nd = Key2nd                           'fake 2nd key off (if not)
  Key2nd = False
  HldIP = InstrPtr                          'save this, in case we are in LRN mode
  Call MainKeyPad(1)                        'toggle between learn and calc modes
  InstrPtr = HldIP                          'in case next takes us to LRN mode
  Call MainKeyPad(1)
  Key2nd = Hld2nd                           'reset 2nd key
End Sub

'*******************************************************************************
' Function Name     : ScrubFile
' Purpose           : Initialize file to import. Support routine
'*******************************************************************************
Private Function ScrubFile(Filepath As String, ByVal ClipBrd As Boolean) As String()
  Dim Ary() As String, Ary2() As String, S As String
  Dim C As String, R As String, Nx As String, Pv As String
  Dim Idx As Long, Qidx As Long, Pidx As Long, i As Long
  Dim ts As TextStream
  Dim Style2 As Boolean
'
' first just try to read the file
'
  If ClipBrd Then                           'if reading from clipboard
    ErrorV = True                           'init to error
    S = Clipboard.GetText(vbCFText)         'grab text
    If Len(S) = 0 Then Exit Function        'nothing
    Ary = Split(S, vbCrLf)                  'else build array
    If Not IsDimmed(Ary) Then Exit Function 'nothing
    ErrorV = False                          'else assume ok
  Else
    On Error Resume Next
    Set ts = Fso.OpenTextFile(Filepath, ForReading, False)  'try to open the file
    If Not CBool(Err.Number) Then Ary = Split(ts.ReadAll, vbCrLf)
    ts.Close
    ErrorV = CBool(Err.Number)
    If ErrorV Then Exit Function
    On Error GoTo 0
  End If
'
' test for Style 2 (line numbers precede code
'
  Style2 = Left$(Ary(0), 5) = "0000 "
'
' now initially parse it
'
  For Idx = 0 To UBound(Ary)
    S = Ary(Idx)                                  'grab a line
    If Style2 Then S = Mid$(S, 11)                'whack program step if Style 2
    S = Trim$(S)                                  'trim surrounding spaces
    If CBool(Len(S)) Then
      i = InStr(1, S, """""")                     'found doubled quotes?
      Do While CBool(i)
        S = Left$(S, i - 1) & Chr$(129) & Mid$(S, i + 2)  'strip and covert to internal token
        i = InStr(1, S, """""")
      Loop
      If Len(S) > 1 And Left$(S, 1) = "'" Then
        S = "' """ & Mid$(S, 2) & """"              'handle remarks
      ElseIf UCase$(Left$(S, 4)) = "REM " Then
        S = "Rem """ & Mid$(S, 5) & """"
      ElseIf Len(S) = 0 Then                        'allow for NOP line breaks
        S = "NOP"                                   'and mark blank lines
      ElseIf Len(S) = 3 Then                        'handle ASCII text for Style 0
        If Left$(S, 1) = "[" And Right$(S, 1) = "]" Then
          S = """" & Mid$(S, 2, 1) & """"
        End If
      End If
    End If
    Ary(Idx) = S                                  'stuff final result
  Next Idx
'
' now do the serious stuff. Break the file apart at spaces, rejoin text
'
  S = Join(Ary, " ")                  'replace CRLF with a space...
  Ary = Split(S, " ")                 'and then split everything again on spaces
  Qidx = 0
  For Idx = 0 To UBound(Ary)
    S = Ary(Idx)
    If CBool(Qidx) Then                     'joining text
      Ary(Qidx) = Ary(Qidx) & Chr$(31) & S  'pad with temp spaces (31)
      Ary(Idx) = vbNullString               'remove stripped line
    End If
    
    If CBool(Len(S)) Then
      If Left$(S, 1) = """" Then      'if left character is a quote
        If CBool(Qidx) Then           'if quotes on...
          Qidx = 0                    'then turn off
        Else
          Qidx = Idx                  'else turn on
        End If
      End If
      If Len(S) > 1 Then              'if multi-char, check for turning back off
        If Right$(S, 1) = """" Then Qidx = 0
        If Qidx Then                  'if Qidx still set...
          Pidx = InStr(2, S, """")    'but quote is embedded WITHIN text
          If CBool(Pidx) Then         'Found enbedded quote?
            If Ary(Qidx) = S Then     'if save is same as source...
              Ary(Qidx) = Left$(S, Pidx) & " " & Mid$(S, Pidx + 1) 'simply break text up
            Else                      'otherwise we will append quoted stuff, then link rest after space
              Ary(Qidx) = Ary(Qidx) & Chr$(31) & Left$(S, Pidx) & " " & Mid$(S, Pidx + 1)
            End If
            Qidx = 0                  'and turn off quote flag, since we found embedded quote
          End If
        End If
      End If
    End If
  Next Idx
'
' compress the buffer by strupping blank data
'
  Qidx = 0
  For Idx = 0 To UBound(Ary)
    If CBool(Len(Ary(Idx))) Then  'if data present...
      Ary(Qidx) = Ary(Idx)        'move to compressed space
      Qidx = Qidx + 1             'bump new index
    End If
  Next Idx
  For Idx = Qidx - 1 To 0 Step -1
    If Ary(Idx) = "NOP" Then
      Qidx = Qidx - 1
    Else
      Exit For
    End If
  Next Idx
  ReDim Preserve Ary(Qidx - 1)    'resize buffer
  S = Join(Ary, " ")              'combine the data
  Ary = Split(S, " ")             'then re-split it
'
' now convert temp text spaces back to real spaces
'
  For Idx = 0 To UBound(Ary)
    S = Ary(Idx)
    If Left$(S, 1) = Chr$(34) Then  'process only quoted text
      Qidx = InStr(1, S, Chr$(31))  'found temp space?
      Do While CBool(Qidx)
        Mid$(S, Qidx, 1) = " "      'replace with real space
        Qidx = InStr(Qidx + 1, S, Chr$(31))
      Loop                          'check all
      Ary(Idx) = S                  'stuff result
    End If
  Next Idx
  ScrubFile = Ary                   'return result
End Function

'*******************************************************************************
' Function Name     : ImpCmd
' Purpose           : Check imported command for proper syntax. Support routine
'*******************************************************************************
Public Function ImpCmd(Txt As String) As Integer
  Dim Idx As Integer, LnTxt As Integer, LnRslt As Integer
  Dim S As String
'
' handle Ukey data
'
  If UCase$(Left$(Txt, 6)) = "<UKEY_" Then
    ImpCmd = Asc(UCase$(Mid$(Txt, 7, 1))) - 64 + 900
    Exit Function
  End If
'
' handle all others
'
  LnTxt = Len(Txt)
  Select Case UCase$(Txt)
    Case Chr$(129)
      Idx = 34
    Case "SELECT": Idx = 188
    Case "STRUCT": Idx = 234
    Case "CLRVAR": Idx = 240
    Case "PRINTF": Idx = 337
    Case "ELSEIF": Idx = 343
    Case "PRINT;": Idx = 352
    Case "CIRCLE": Idx = 365
    Case "RESET": Idx = 147
    Case "STFLG": Idx = 162
    Case "IFFLG": Idx = 163
    Case "STYLE": Idx = 172
    Case "WHILE": Idx = 203
    Case "BREAK": Idx = 217
    Case "PRINT": Idx = 224
    Case "CONST": Idx = 233
    Case "NXLBL": Idx = 235
    Case "PVLBL": Idx = 236
    Case "IFDEF": Idx = 243
    Case "ASCII": Idx = 260
    Case "PASTE": Idx = 266
    Case "STKEX": Idx = 271
    Case "STDEV": Idx = 279
    Case "VARNC": Idx = 280
    Case "D.DDD": Idx = 306
    Case "LTRIM": Idx = 310
    Case "RTRIM": Idx = 311
    Case "CLOSE": Idx = 317
    Case "WRITE": Idx = 319
    Case "SYSBP": Idx = 326
    Case "GFREE": Idx = 345
    Case "UNTIL": Idx = 359
    Case "ADROF": Idx = 362
    Case "SPLIT": Idx = 366
    Case "REDIM": Idx = 368
    Case "DELSE": Idx = 372
    Case "PASTE": Idx = 266
    Case "LOAD": Idx = 131
    Case "SAVE": Idx = 132
    Case "COPY": Idx = 139
    Case "PTOR": Idx = 140
    Case "HKEY": Idx = 148
    Case "MEAN": Idx = 151
    Case "X><T": Idx = 153
    Case "X==T": Idx = 164
    Case "X>=T": Idx = 165
    Case "D.MS": Idx = 178
    Case "CASE": Idx = 189
    Case "BEEP": Idx = 198
    Case "UKEY": Idx = 206
    Case "PLOT": Idx = 211
    Case "NVAR": Idx = 212
    Case "ELSE": Idx = 215
    Case "CONT": Idx = 216
    Case "GRAD": Idx = 218
    Case "TVAR": Idx = 225
    Case "LINE": Idx = 237
    Case "SZOF": Idx = 241
    Case "EDEF": Idx = 244
    Case "LAPP": Idx = 259
    Case "LIST": Idx = iList
    Case "RTOP": Idx = 268
    Case "PUSH": Idx = 269
    Case "SKEY": Idx = 276
    Case "YINT": Idx = 281
    Case "LOGX": Idx = 286
    Case "RFLG": Idx = 290
    Case "!FLG": Idx = 291
    Case "X!=T": Idx = 292
    Case "X<=T": Idx = 293
    Case "FRAC": Idx = 303
    Case "!FIX": Idx = 305
    Case "CALL": Idx = 308
    Case "TRIM": Idx = 309
    Case "OPEN": Idx = 316
    Case "READ": Idx = 318
    Case "SWAP": Idx = 320
    Case "INCR": Idx = 329
    Case "DECR": Idx = 330
    Case "DSNZ": Idx = 332
    Case "LSET": Idx = 335
    Case "RSET": Idx = 336
    Case "IVAR": Idx = 340
    Case "STOP": Idx = 347
    Case "WITH": Idx = 348
    Case "CVAR": Idx = 353
    Case "ROOT": Idx = 355
    Case "SQRT": Idx = 356
    Case "RND#": Idx = 358
    Case "ENUM": Idx = 361
    Case "PCMP": Idx = 363
    Case "COMP": Idx = 364
    Case "JOIN": Idx = 367
    Case "UDEF": Idx = 370
    Case "!DEF": Idx = 371
    Case "COPY": Idx = 139
    Case "LIST": Idx = iList
    Case "1/X": Idx = 158
    Case "PGM": Idx = 130
    Case "CLR": Idx = 134
    Case "SST": Idx = 136
    Case "INS": Idx = 137
    Case "CUT": Idx = 138
    Case "STO": Idx = 141
    Case "RCL": Idx = 142
    Case "EXC": Idx = 143
    Case "SUM": Idx = 144
    Case "MUL": Idx = 145
    Case "IND": Idx = 146
    Case "LNX": Idx = 149
    Case "HYP": Idx = iHyp
    Case "SIN": Idx = 155
    Case "COS": Idx = 156
    Case "TAN": Idx = 157
    Case "TXT": Idx = 159
    Case "HEX": Idx = 160
    Case "X>T": Idx = 166
    Case "DFN": Idx = 167
    Case "DEC": Idx = 173
    Case "INT": Idx = 175
    Case "ABS": Idx = 176
    Case "FIX": Idx = 177
    Case "SBR": Idx = 180
    Case "REM": Idx = iRem
    Case "OCT": Idx = 186
    Case "DEG": Idx = 192
    Case "LBL": Idx = 193
    Case "BIN": Idx = 199
    Case "FOR": Idx = 201
    Case "PMT": Idx = 204
    Case "RAD": Idx = 205
    Case "R/S": Idx = 219
    Case "+/-": Idx = 222
    Case "RND": Idx = 230
    Case "MIL": Idx = 231
    Case "PVT": Idx = 232
    Case "DEF": Idx = 242
    Case "MDL": Idx = 257
    Case "CMM": Idx = 258
    Case "CMS": Idx = 261
    Case "BST": Idx = 264
    Case "DEL": Idx = 265
    Case "LBL": Idx = iUSR
    Case "POP": Idx = 270
    Case "SUB": Idx = 272
    Case "DIV": Idx = 273
    Case "ARC": Idx = iArc
    Case "SEC": Idx = 283
    Case "CSC": Idx = 284
    Case "COT": Idx = 285
    Case "VAR": Idx = 287
    Case "X<T": Idx = 294
    Case "NOP": Idx = 295
    Case "LOG": Idx = 297
    Case "10^": Idx = 298
    Case "FMT": Idx = 300
    Case "SGN": Idx = 304
    Case "!EE": Idx = 307
    Case "GTO": Idx = 321
    Case "LOF": Idx = 322
    Case "GET": Idx = 323
    Case "PUT": Idx = 324
    Case "NOR": Idx = 328
    Case "DSZ": Idx = 331
    Case "ALL": Idx = 333
    Case "RTN": Idx = 334
    Case "RGB": Idx = 339
    Case "DBG": Idx = 344
    Case "LEN": Idx = 346
    Case "VAL": Idx = 350
    Case "ADV": Idx = 351
    Case "PUB": Idx = 360
    Case "MID": Idx = 369
    Case "LRN": Idx = LRNKey
    Case "TXT": Idx = iTXT
    Case "SST": Idx = 136
    Case "INS": Idx = 137
    Case "CUT": Idx = 138
    Case "USR": Idx = iUSR
    Case "BST": Idx = 264
    Case "DEL": Idx = 265
    Case "NOP": Idx = iNOP
    Case "CE": Idx = 133
    Case "OP": Idx = 135
    Case "E+": Idx = 150
    Case "X!": Idx = 152
    Case "EE": Idx = 179
    Case "DO": Idx = 202
    Case "IF": Idx = 214
    Case ">>": Idx = 226
    Case "Y^": Idx = 227
    Case "X²": Idx = 228
    Case "PI": Idx = 229
    Case "CP": Idx = 262
    Case "EX": Idx = 277
    Case "E-": Idx = 278
    Case "==": Idx = 288
    Case "&&": Idx = 289
    Case "÷=": Idx = 299
    Case "!=": Idx = 301
    Case "||": Idx = 302
    Case "X=": Idx = 312
    Case ">=": Idx = 314
    Case "-=": Idx = 325
    Case "<=": Idx = 327
    Case "+=": Idx = 338
    Case "AS": Idx = 342
    Case "<<": Idx = 354
    Case "&": Idx = 161
    Case ":": Idx = iColon
    Case "(": Idx = 169
    Case ")": Idx = 170
    Case "÷": Idx = 171
    Case "|": Idx = 174
    Case "X": Idx = 184
    Case "~": Idx = 187
    Case "{": Idx = 190
    Case "}": Idx = 191
    Case "-": Idx = 197
    Case "^": Idx = 200
    Case "+": Idx = 210
    Case "%": Idx = 213
    Case ".": Idx = 221
    Case "=": Idx = 223
    Case "[": Idx = 238
    Case "]": Idx = 239
    Case "<": Idx = 274
    Case ">": Idx = 275
    Case ";": Idx = iSemiC
    Case "'": Idx = iRem2
    Case "!": Idx = 315
    Case "\": Idx = 341
    Case ",": Idx = 349
    Case "E": Idx = 357
    Case "0" To "9": Idx = Asc(Txt) - 48
    Case Else: Idx = 128
  End Select
'
' ensure wierd trick VB does with numbers did not sneak by
' (ie, "3)" will result in 3, which is of course incorrect)
'
  Select Case Idx
    Case 128, iNOP
    Case Else
      If Len(GetInst(Idx)) <> LnTxt Then Idx = 128
  End Select
  ImpCmd = Idx
End Function

'*******************************************************************************
' Subroutine Name   : AddInst
' Purpose           : Add an instruction to the program list. Support routine
'*******************************************************************************
Private Sub AddInst(ByVal Code As Integer)
  If Iptr > iSize Then          'ensure enough space to hold the instruction exists
    iSize = iSize + InstrInc
    ReDim Preserve iAry(iSize)
  End If
  If Code = 129 Then Code = 34  'convert special quote place-holders
  iAry(Iptr) = Code
  Iptr = Iptr + 1
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

