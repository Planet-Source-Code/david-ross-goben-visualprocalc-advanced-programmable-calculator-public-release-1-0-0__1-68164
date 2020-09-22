Attribute VB_Name = "modInstrSupport"
Option Explicit

'*******************************************************************************
' Function Name     : GetInstrPtr
' Purpose           : Return Instruction Pointer for current Program
'*******************************************************************************
Public Function GetInstrPtr() As Integer
  GetInstrPtr = InstrPtr
End Function

'*******************************************************************************
' Function Name     : GetInstrCnt
' Purpose           : Get Instruction Count for current Program
'*******************************************************************************
Public Function GetInstrCnt(Optional Pgm As Integer = 0) As Integer
  If CBool(Pgm) Then                'did the user supply a program number?
    GetInstrCnt = ModMap(Pgm) - ModMap(Pgm - 1)
  ElseIf CBool(ActivePgm) Then
    GetInstrCnt = ModMap(ActivePgm) - ModMap(ActivePgm - 1)
  Else
    GetInstrCnt = InstrCnt          'If Pgm 0, simply return active instruction count
  End If
End Function

'*******************************************************************************
' Function Name     : SetInstrCnt
' Purpose           : Set Instruction Count for current Program (used only by Pgm00)
'*******************************************************************************
Public Sub SetInstrCnt(ByVal Inst As Integer)
  If Not CBool(ActivePgm) Then
    InstrCnt = Inst
  End If
End Sub

'*******************************************************************************
' Function Name     : GetInstruction
' Purpose           : Get Instruction for current Program as specified offset
'*******************************************************************************
Public Function GetInstruction(ByVal Offset As Integer) As Integer
  Dim ptr As Long
  
  ptr = InstrPtr + Offset                   'point to target location
  If ptr >= GetInstrCnt() Or ptr < 0 Then   'if we exceed memory
    GetInstruction = -1
  Else
    If CBool(ActivePgm) Then                                'if module pgm...
      GetInstruction = ModMem(ModMap(ActivePgm - 1) + ptr)  'ActivePgm-1 is offset to Base of ActivePgm
    Else
      GetInstruction = Instructions(ptr)                    'return user pgm instruction found
    End If
  End If
End Function

'*******************************************************************************
' Function Name     : GetInstructionAt
' Purpose           : Get Instruction for current Program
'*******************************************************************************
Public Function GetInstructionAt(ByVal Iptr As Integer) As Integer
  Dim ptr As Long
  
  If Iptr >= GetInstrCnt() Or Iptr < 0 Then 'if we exceed memory
    GetInstructionAt = -1
  Else
    If CBool(ActivePgm) Then
      GetInstructionAt = ModMem(ModMap(ActivePgm - 1) + Iptr)
    Else
      GetInstructionAt = Instructions(Iptr) 'return instruction found
    End If
  End If
End Function

'*******************************************************************************
' Subroutine Name   : IncInstrPtr
' Purpose           : Increment Instruction Pointer for current Program
'*******************************************************************************
Public Function IncInstrPtr() As Boolean
  InstrPtr = InstrPtr + 1           'bump program index
  If InstrPtr >= GetInstrCnt() Then 'if we exceeded size of program...
    InstrPtr = 0                    'loop around if at max
    IncInstrPtr = True              'indicate this looping occurred
  End If
End Function

'*******************************************************************************
' Function Name     : FindEPar
' Purpose           : Find end of the current () block, type sensetive
'*******************************************************************************
Public Function FindEPar(ByVal Ofst As Integer) As Integer
  Dim Iptr As Integer, ParCnt As Integer
  
  Iptr = Ofst                         'init ptr offset
  ParCnt = 1                          'we are inside a select block
  Do
    Iptr = Iptr + 1                   'bump index
    Select Case GetInstruction(Iptr)
      Case iLparen                    'new () block?
        ParCnt = ParCnt + 1           'yes, so found another (
      Case iRparen                    'else end paren for a () block?
        ParCnt = ParCnt - 1           'yes, so back off 1
        If ParCnt = 0 Then            'if exhausted
          FindEPar = InstrPtr + Iptr + 1 'point beyond ")"
          Exit Function
        End If
      Case iUparen, iIparen, iWparen, iFparen, iCparen, iEparen, iSparen, iDWparen
        ParCnt = ParCnt - 1           'yes, so back off 1
        If ParCnt = 0 Then            'if exhausted
          ForcError "The 'matching' ending parentheses is reserved for another instruction"
          FindEPar = -1               'indicate matching ")" not found
          Exit Function
        End If
      Case -1                         'exceeded memory
        FindEPar = -1                 'indicate ")" not found
        Exit Function
    End Select
  Loop
End Function

'*******************************************************************************
' Function Name     : FindEPar2
' Purpose           : Find end of the current () block, type insensetive
'*******************************************************************************
Public Function FindEPar2(ByVal Ofst As Integer) As Integer
  Dim Iptr As Integer, ParCnt As Integer
  
  Iptr = Ofst                         'init ptr offset
  ParCnt = 1                          'we are inside a select block
  Do
    Iptr = Iptr + 1                   'bump index
    Select Case GetInstruction(Iptr)
      Case iLparen                    'new () block?
        ParCnt = ParCnt + 1           'yes, so found another (
      Case iRparen, iUparen, iIparen, iWparen, _
           iFparen, iCparen, iEparen, iSparen, iDWparen
        ParCnt = ParCnt - 1           'yes, so back off 1
        If ParCnt = 0 Then            'if exhausted
          FindEPar2 = InstrPtr + Iptr + 1 'point beyond ")"
          Exit Function
        End If
      Case -1                         'exceeded memory
        FindEPar2 = -1                'indicate ")" not found
        Exit Function
    End Select
  Loop
End Function

'*******************************************************************************
' Function Name     : FindEbrace
' Purpose           : Find a corresponding end brace, type sensetive
'*******************************************************************************
Public Function FindEbrace() As Integer
  Dim Iptr As Integer, BrcLvl As Integer
  
  BrcLvl = 1                          'init level 1 (we are at an opening brace
  Iptr = 0                            'init ptr offset
  
  Do
    Iptr = Iptr + 1                   'bump index
    Select Case GetInstruction(Iptr)
      Case iLCbrace
        BrcLvl = BrcLvl + 1           'bump count higher
      Case iRCbrace                   'normal end brace
        BrcLvl = BrcLvl - 1           'else back off 1
        If BrcLvl = 0 Then            'if exhausted
          FindEbrace = InstrPtr + Iptr
          Exit Function
        End If
      Case iBCBrace, iSCBrace, iICBrace, iDCBrace, iWCBrace, _
           iFCBrace, iCCBrace, iDWBrace, iDUBrace, iEIBrace, _
           iENBrace, iSTBrace, iSIBrace, iCNBrace
        BrcLvl = BrcLvl - 1           'else back off 1
        If BrcLvl = 0 Then            'if exhausted
          ForcError "the 'matching' ending brace is reserved for another command"
          FindEbrace = -1             'indicate matching "}" not found
          Exit Function
        End If
      Case -1                         'exceeded memory
        FindEbrace = -1
        Exit Function
    End Select
  Loop
End Function

'*******************************************************************************
' Subroutine Name   : FindForInfo
' Purpose           : Find 3 parts in a FOR loop definition
'*******************************************************************************
Public Sub FindForInfo(ByRef Init As Integer, ByRef Incr As Integer, ByRef Cond As Integer)
  Dim Idx As Integer, Code As Integer
  Dim SemiCnt As Long
  
  Idx = 1                                 'point to "(" beyond FOR (we know "(" exists)
  Init = 1                                'assume Initialization data is after "(" for now
  Incr = -1                               'init incrementer/process and conditional to "do nothing"
  Cond = -1
  SemiCnt = 0                             'init to no semicolons found
  
  Do While GetInstruction(Idx) <> iFparen 'now search to end of FOR declaration
    Idx = Idx + 1                         'point to next instruction
    Code = GetInstruction(Idx)            'get code found there
    If Preprocessing And Code = iSemiC Then 'if Preprocessing, convert iSemiC to iSemiColon
      Code = iSemiColon
      Instructions(InstrPtr + Idx) = Code 'set new code
    End If
    If Code = iSemiColon Then             'found a semicolon?
      Select Case SemiCnt
        Case 0
          SemiCnt = 1                     'found 1
          Cond = Idx + 1                  'assign it to the Cond pointer
        Case 1
          SemiCnt = 2                     'found 2
          Incr = Idx + 1                  'assign process/Increment data pointer
        Case Else
          SemiCnt = 3                     'found too many
      End Select
    End If
  Loop
  If SemiCnt <> 2 Then Init = 0           'if we did not find 2, then we have an error
'
' now check to see if the Process or Incrementor point to "do nothing" code
'
  Idx = InstrPtr                          'point to FOR instruction
  If CBool(Init) Then                     'if Init is a valid value...
    If GetInstruction(Init) = iSemiColon Then
      Init = -1                           'do nothing if we see a semicolon
    Else
      Init = Idx                          'point to '('-1
    End If
    If GetInstruction(Cond) = iSemiColon Then
      Cond = -1                           'do nothing if we see a semicolon
    Else
      Cond = Idx + Cond - 1               'point to ';'
    End If
    If GetInstruction(Incr) = iFparen Then
      Incr = -1                           'do nothing if we see the end of the def
    Else
      Incr = Idx + Incr - 1               'point to ';'
    End If
  End If
End Sub

'*******************************************************************************
' Function Name     : FindEblock
' Purpose           : Find the end of a block. This subroutine does not care
'                   : about a certain type of end brace terminator, but simply
'                   : wants to find a corresponding end brace (used by Run-Time)
'*******************************************************************************
Public Function FindEblock(ByVal Iptr As Integer) As Integer
  Dim BrcLvl As Integer
  
  BrcLvl = 1
  Do
    Iptr = Iptr + 1                   'bump index
    Select Case GetInstructionAt(Iptr)
      Case iLCbrace
        BrcLvl = BrcLvl + 1           'bump count higher
      Case iBCBrace, iSCBrace, iICBrace, iDCBrace, iWCBrace, _
           iFCBrace, iCCBrace, iDWBrace, iDUBrace, iEIBrace, _
           iENBrace, iSTBrace, iSIBrace, iCNBrace, iRCbrace
        BrcLvl = BrcLvl - 1           'else back off 1
        If BrcLvl = 0 Then            'if exhausted
          FindEblock = Iptr
          Exit Function
        End If
      Case -1                         'exceeded memory (should never happen after precomp)
        FindEblock = -1
        Exit Function
    End Select
  Loop
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

