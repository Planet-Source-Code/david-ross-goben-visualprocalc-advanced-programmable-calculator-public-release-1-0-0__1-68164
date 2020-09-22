Attribute VB_Name = "ModCompress"
Option Explicit

'*******************************************************************************
' Subroutine Name   : Compress
' Purpose           : Tighten code further after Preprocessr.
'                   : Remove comments, and other optional tokens, such as 'Dfn'
'                   : Merge many commands, such as HYP ARC Sin, or STO IND.
'*******************************************************************************
Public Sub Compress()
  Dim HldIptr As Integer, i As Integer
  Dim NumData As String, C As String
  Dim LclIdx As Integer, StructEnd As Integer
  Dim HaveStruct As Boolean
  
  If Compressd Then Exit Sub           'if we are already Compressed, then we can run it
  Compressing = True
  Preprocessd = False                 'force Preprocess to set variable names
  Call Preprocess                     'invoke Compress pre-processor
  If Not Preprocessd Then
    Compressing = False
    Exit Sub                          'if that failed, then nothing more to do
  End If
  
  InstrPtr = 0                        'init to base
  ReDim ModLcl(InstrCnt)              'init size of local module data
  LclIdx = 0                          'init stuffing point
  HaveStruct = False                  'no structures defined yet
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'  Debug.Assert False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  Do While InstrPtr < InstrCnt
    If CBool(InstrPtr) Then           'if instruction pointer not 0, get previous code
      PrvCode = Instructions(InstrPtr - 1)
    Else
      PrvCode = -1                    'no previous code
    End If
        
    If HaveStruct Then
      Code = 512
    Else
      Code = Instructions(InstrPtr)     'grab current code
    End If
    Select Case Code
  '---------------------------------------------------------------------
      Case iRem, iRem2    'we will remove remarks
        If CBool(InstrPtr) Then                     'do not erase if rem at VERY START (use for pgm List)
          Do
            Select Case Instructions(InstrPtr + 1)  'is next ASCII?
              Case 10 To 127
                InstrPtr = InstrPtr + 1             'yes, so skip
              Case Else
                Exit Do                             'else we are at last character of remark
            End Select
          Loop
          Code = 128                                'do not copy command
        Else
          Code = iRem2                              'use "'"
        End If
  '---------------------------------------------------------------------
      Case iDfn, iNOP, iDBG
        Code = 128                                'remove optional keywords
  '---------------------------------------------------------------------
      Case iHyp
        InstrPtr = InstrPtr + 1
        Select Case Instructions(InstrPtr)
          Case iSin
            Code = 403  ' SinH
          Case iCos
            Code = 404  ' CosH
          Case iTan
            Code = 405  ' TanH
          Case iSec
            Code = 412  ' SecH
          Case iCsc
            Code = 413  ' CscH
          Case iCot
            Code = 414  ' CotH
          Case iArc
            InstrPtr = InstrPtr + 1
            Select Case Instructions(InstrPtr)
              Case iSin
                Code = 406 ' ArcSinH
              Case iCos
                Code = 407  ' ArcCosH
              Case iTan
                Code = 408  ' ArcTanH
              Case iSec
                Code = 415  ' ArcSecH
              Case iCsc
                Code = 416  ' ArcCscH
              Case iCot
                Code = 417  ' ArcCotH
            End Select
        End Select
      '---------------------------------------
      Case iArc
        InstrPtr = InstrPtr + 1
        Select Case Instructions(InstrPtr)
          Case iSin
            Code = 400  ' Asin
          Case iCos
            Code = 401  ' Acos
          Case iTan
            Code = 402  ' Atan
          Case iSec
            Code = 409  ' ASec
          Case iCsc
            Code = 410  ' ACsc
          Case iCot
            Code = 411  ' Acot
        End Select
      '---------------------------------------
      Case iStruct                            'define structure
        i = InstrPtr + 1                      'find opening brace
        Do While Instructions(i) <> iLCbrace
          i = i + 1
        Loop
        StructEnd = FindEblock(i)             'get end brace location
        HaveStruct = True                     'indicate structure definition
      '---------------------------------------
      Case iNvar, iTvar, iIvar, iCvar
        Call CheckForNumber(InstrPtr, 2, 99)            'check variable number
        NumData = CStr(TstData)                         'grab variable number as string
        If Instructions(InstrPtr + 1) = iLbl Then       'name supplied?
          Call CheckForLabel(InstrPtr + 1, LabelWidth)  'skip name of variable (we will not need it)
        End If
      '---------------------------------------
      Case iSTO, iRCL, iEXC, iSUM, iMUL, iSUB, iDIV, iVar, _
           iTrim, iRTrim, iLTrim, iLSet, iRSet, iMid, iReDim
        If Instructions(InstrPtr + 1) <> iIND Then
          NumData = ChkCompVbl()
          If Not CBool(Len(NumData)) Then Exit Sub
        End If
      '---------------------------------------
      Case iClrVar
        Select Case Instructions(InstrPtr + 1)
          Case iIND, iAll
          Case Else
            ModLcl(LclIdx) = Code
            LclIdx = LclIdx + 1
            NumData = ChkCompVbl()
            If Not CBool(NumData) Then Exit Sub
            Code = 128
        End Select
      '---------------------------------------
      Case iIND
        Select Case PrvCode
          Case iSTO
            ModLcl(LclIdx - 1) = 245  ' STO IND
            Code = 128                'prevent IND from being stuffed to ModLcl() array
          Case iRCL
            ModLcl(LclIdx - 1) = 246  ' RCL IND
            Code = 128
          Case iEXC
            ModLcl(LclIdx - 1) = 247  ' EXC IND
            Code = 128
          Case iSUM
            ModLcl(LclIdx - 1) = 248  ' SUM IND
            Code = 128
          Case iMUL
            ModLcl(LclIdx - 1) = 249  ' MUL IND
            Code = 128
          Case iSUB
            ModLcl(LclIdx - 1) = 250  ' SUB IND
            Code = 128
          Case iDIV
            ModLcl(LclIdx - 1) = 251  ' DIV IND
            Code = 128
          Case iGTO
            ModLcl(LclIdx - 1) = 252  ' GTO IND
            Code = 128
          Case iOP
            ModLcl(LclIdx - 1) = 254  ' OP  IND
            Code = 128
          Case iFix
            ModLcl(LclIdx - 1) = 255  ' FIX IND
            Code = 128
          Case iPgm
            ModLcl(LclIdx - 1) = 256  ' PGM IND
            Code = 128
          Case iStFlg
            ModLcl(LclIdx - 1) = 373  ' StFlg IND
            Code = 128
          Case iRFlg
            ModLcl(LclIdx - 1) = 374  ' RFlg IND
            Code = 128
          Case iIfFlg
            ModLcl(LclIdx - 1) = 375  ' IfFlg IND
            Code = 128
          Case inFlg
            ModLcl(LclIdx - 1) = 377  ' !Flg IND
            Code = 128
          Case iDsz
            ModLcl(LclIdx - 1) = 378  ' Dsz IND
            Code = 128
          Case iDsnz
            ModLcl(LclIdx - 1) = 379  ' Dsnz IND
            Code = 128
          Case iIncr
            ModLcl(LclIdx - 1) = 380  ' Incr IND
            Code = 128
          Case iDecr
            ModLcl(LclIdx - 1) = 381  ' Decr IND
            Code = 128
          Case iVar
            ModLcl(LclIdx - 1) = 433  ' Var IND
            Code = 128
        End Select
        InstrPtr = InstrPtr + 1             'point past IND
        NumData = ChkCompVbl()              'Grab Data
        If Not CBool(NumData) Then Exit Sub 'error
      Case iVar
        Select Case PrvCode
          Case iSTO, iRCL, iEXC, iSUM, iMUL, iSUB, iDIV, iIND, iWith
            Code = 128                      'ignore code
        End Select
    End Select
   '---------------------------------------
    If HaveStruct Then
      Code = Instructions(InstrPtr)     'grab current code
    End If
    
    If Code <> 128 Then                 'if we have a valid code
      ModLcl(LclIdx) = Code             'transfer it
      LclIdx = LclIdx + 1               'bump dest
    End If
    
    If CBool(Len(NumData)) Then         'if numeric data (variable number)
      For i = 1 To Len(NumData)         'convert each ASCII character to 0 to 9
        ModLcl(LclIdx) = Asc(Mid$(NumData, i, 1)) - 48
        LclIdx = LclIdx + 1
      Next i
      NumData = vbNullString            'clear accumulator
    End If
    
    InstrPtr = InstrPtr + 1             'bump the instruction pointer for next instruction
    
    If HaveStruct Then                  'if structure being defined
      HaveStruct = InstrPtr < StructEnd 'set flag based on index
    End If
  Loop
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  Compressing = False               'no longer Compressing
  InstrCnt = LclIdx                 'set new upper bounds
  LclIdx = LclIdx - 1               'set to upper bounds
  ReDim Instructions(LclIdx)        'redim Main program space
'
' copy new data to base
'
  For i = 0 To LclIdx
    Instructions(i) = ModLcl(i)
  Next i
  InstrPtr = 0                      're-init to base
  Erase ModLcl                      'remove temp array
  IsDirty = True                    'program has changed
'
' force Preprocess again, for formatted listing purposes
'
  Preprocessd = False               'force Preprocess to set variable names
  Call Preprocess                   'invoke Compress pre-processor
  If Not Preprocessd Then Exit Sub  'if that failed, then nothing more to do
'
' all is OK, so allow Compressed flag to be set
'
  Compressd = True                   'if we are here, then we Compressed OK
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

