Attribute VB_Name = "modDefinitions"
Option Explicit

'-------------------------------------------------------------------------------
' Constant Definitions
'-------------------------------------------------------------------------------
Public Const AppTitle As String = "The Personal Programmable Calculator"

Public Const WinMinW As Long = 14910          'minimum main window width
Public Const WinMinH As Long = 7095           'minimum main window height

Public Const DisplayWidth As Integer = 36     '(6 fields x 6 character width)
Public Const DisplayHeight As Integer = 28    'number of lines in the display
Public Const FieldCnt As Integer = 6          'number of fields on the display
Public Const FieldWidth As Integer = 6        'single field width
Public Const LabelWidth As Integer = 12       'width of label, variable, and subroutine names

Public Const PlotWidth As Long = 300          'pixel width of plot field (Actual = 309)
Public Const PlotHeight As Long = 360         'pixel height of plot field (Actual = 368)
Public Const PlotXOfst As Long = 3            '3 for left offset to draw region (act. width is 309)
Public Const PlotYOfst As Long = 2            '2 for top offset to draw region (act. height is 368)

Public Const SizeInst As Integer = 1024       'number of instruction to start the pool with
Public Const InstrInc As Integer = 128        'incremental offset for bumping the size of the pool
Public Const MaxInstr As Integer = 9999       'the instruction list cannot exceed this amount
Public Const SizePend As Integer = 32         'pending operation size
Public Const PendInc As Integer = 16          'incremental offset
Public Const SizeDef As Integer = 32          'intial Def-inition size of 32 symbols
Public Const DefInc As Integer = 16           'incremental offset
Public Const MaxVar As Long = 99              'max base variable number (00-99)
Public Const DMaxVar As Double = 99#          'max base variable number (99)
Public Const IndSpc As Integer = 2            '2 white spaces per indent level
Public Const MaxKeys As Integer = 116         'max number of keypad keys (less user/text keys)
Public Const LblDepth As Integer = 32         'incremental depth for labels (increment by 32)
Public Const SbrDepth As Integer = 64         'initial depth for subroutine calls
Public Const ParenDepth As Integer = 64       'incremental depth for parentheses levels
Public Const PushLimit As Integer = 16        'max push levels
Public Const BraceDepth As Integer = 16       'initial depth for braced data tracking
'-------------------------------------------------------------------------------
' general constants
'-------------------------------------------------------------------------------
'note in the next line below that "@" is the SPACE place-holder
Public Const AlphaKeys As String = "@ABCDEFGHIJKLMNOPQRSTUVWXYZ"  'Keyboard A-Z, a-z keys
Public Const AltKeys As String = "!@#$%&*()-_+=\/<>,.?:;|[]""'"   'alt keys (shift or otherwise)
Public Const DefDspFmt As String = "0.##############"
Public Const DefScientific As String = "0.0#############E+00"
Public Const TxtPltDef As String = "*"          'default Text-Plot character
Public Const MaxPLong As Double = 2147483647#   'max positive value a long integer can hold
Public Const DefDtFmt As String = "short date"  'default date format
Public Const DefTmFmt As String = "HH:MM:SS"    'default time format
'-------------------------------------------------------------------------------
' bracing symbols
'-------------------------------------------------------------------------------
Public Const iDWparen As Integer = 487        ' ) special end paren for Do...While()
Public Const iEparen As Integer = 488         ' ) special end paren for normal (
Public Const iSparen As Integer = 489         ' ) special end paren for Select definitions
Public Const iUparen As Integer = 490         ' ) special end paren for Until definitions
Public Const iCparen As Integer = 491         ' ) special end paren for Case definitions
Public Const iIparen As Integer = 492         ' ) special end paren for If definitions
Public Const iWparen As Integer = 493         ' ) special end paren for While definitions
Public Const iFparen As Integer = 494         ' ) special end paren for For definitions

Public Const iLCbrace As Integer = 190        ' { normal open brace
Public Const iRCbrace As Integer = 191        ' } normal close brace
Public Const iEIBrace As Integer = 495        ' } special end brace for Enum item blocks
Public Const iENBrace As Integer = 496        ' } special end brace for Enum blocks
Public Const iSIBrace As Integer = 497        ' } special end brace for Struct item blocks
Public Const iSTBrace As Integer = 498        ' } special end brace for Struct blocks
Public Const iCNBrace As Integer = 499        ' } special end brace for Const blocks
Public Const iDWBrace As Integer = 500        ' } special end brace for Do-While blocks
Public Const iDUBrace As Integer = 501        ' } special end brace for Do-Until blocks
Public Const iICBrace As Integer = 502        ' } special end brace for If blocks
Public Const iDCBrace As Integer = 503        ' } special end brace for Do blocks
Public Const iWCBrace As Integer = 504        ' } special end brace for While blocks
Public Const iFCBrace As Integer = 505        ' } special end brace for For blocks
Public Const iSCBrace As Integer = 506        ' } special end brace for Select blocks
Public Const iBCBrace As Integer = 507        ' } special end brace for Sbr blocks
Public Const iCCBrace As Integer = 508        ' } special end brace for Case blocks
'-------------------------------------------------------------------------------
' the following are used for AOS math operations. ALl commands are assumed to be level 14,
' except as noted in the following:
'-------------------------------------------------------------------------------
Public Const iDivEq As Integer = 299   'level 14 'high   ' /=
Public Const iMulEq As Integer = 312   'level 14         ' x=
Public Const iSubEq As Integer = 325   'level 14         ' -=
Public Const iAddEq As Integer = 338   'level 14         ' +=  Plus all other functions
                                    '-----------------
Public Const iLparen As Integer = 169  'level 13         ' (
Public Const iRparen As Integer = 170  'level 13         ' )
Public Const iLbrkt As Integer = 238   'level 13         ' [
Public Const iRbrkt As Integer = 239   'level 13         ' ]
                                    '-----------------
Public Const iNot As Integer = 315     'level 12         ' !
Public Const iNotB As Integer = 315    'level 12         ' ~
                                    '-----------------
Public Const iPower As Integer = 227   'level 11         ' y^
Public Const iRoot As Integer = 335    'level 11         ' Root
Public Const iLogX As Integer = 286    'level 11         ' LogX
                                    '-----------------
Public Const iMult As Integer = 184    'level 10         ' *
Public Const iDVD As Integer = 171     'level 10         ' ÷
Public Const iBkSlsh As Integer = 341  'level 10         ' \
Public Const iMod As Integer = 213     'level 10         ' %
                                    '-----------------
Public Const iAdd As Integer = 210     'level 9          ' +
Public Const iMinus As Integer = 197   'level 9          ' -
                                    '-----------------
Public Const iShL As Integer = 354     'level 8          ' <<
Public Const iShR As Integer = 226     'level 8          ' >>
                                    '-----------------
Public Const iLT As Integer = 274      'level 7          ' <
Public Const iLE As Integer = 327      'level 7          ' <=
Public Const iGT As Integer = 275      'level 7          ' >
Public Const iGE As Integer = 314      'level 7          ' >=
                                    '-----------------
Public Const iNEQ As Integer = 301     'level 6          ' !=
Public Const iEQ As Integer = 288      'level 6          ' ==
                                    '-----------------
Public Const iAndB As Integer = 161    'level 5          ' &
                                    '-----------------
Public Const iXorB As Integer = 200    'level 4          ' ^
                                    '-----------------
Public Const iOrB As Integer = 174     'level 3          ' |
                                    '-----------------
Public Const iAnd As Integer = 289     'level 2          ' &&
                                    '-----------------
Public Const iOr As Integer = 302      'level 1          ' ||
Public Const iNor As Integer = 328     'level 1          ' Nor
                                    '-----------------
Public Const iEqual As Integer = 223   'level 0  'lowest ' =
'
' The above operations affect the math execution priority. When an operation is encountered, the
' stack is checked for a pending operation. If one is pending, and it has a equal or higher
' priority than the current, the pending operations are executed, doing lowest last.
'-------------------------------------------------------------------------------
'The following are key constants for important primary keys
'-------------------------------------------------------------------------------
Public Const CEKey As Integer = 133                   ' CE
Public Const iCLR As Integer = 134                    ' CLR
Public Const iCP As Integer = 262                     ' CP
Public Const EEKey As Integer = 179                   ' EE
Public Const LRNKey As Integer = 129                  ' LRN
Public Const iRunStop As Integer = 219                ' R/S
Public Const iIND As Integer = 146                    ' IND

Public Const iDfn As Integer = 167                    ' Dfn
Public Const iPvt As Integer = 232                    ' Pvt
Public Const iPub As Integer = 360                    ' Pub
Public Const iList As Integer = 267                   ' List

Public Const iNvar As Integer = 212                   ' Nvar
Public Const iTvar As Integer = 225                   ' Tvar
Public Const iIvar As Integer = 340                   ' Ivar
Public Const iCvar As Integer = 353                   ' Cvar

Public Const iMDL As Integer = 257                    ' MDL
Public Const iPgm As Integer = 130                    ' Pgm
Public Const iSbr As Integer = 180                    ' Sbr
Public Const iUkey As Integer = 206                   ' Ukey
Public Const iLbl As Integer = 193                    ' Lbl
Public Const iStruct As Integer = 234                 ' Struct
Public Const iEnum As Integer = 361                   ' Enum
Public Const iConst As Integer = 233                  ' Const
Public Const iVar As Integer = 287                    ' Var
Public Const iCall As Integer = 308                   ' Call
Public Const iGTO As Integer = 321                    ' GTO
Public Const iAdrOf As Integer = 362                  ' AdrOf
Public Const iFmt As Integer = 300                    ' Fmt
Public Const iTXT As Integer = 159                    ' TXT
Public Const iColon As Integer = 296                  ' [:] colon
Public Const iSemiC As Integer = 168                  ' [;] semi-colon
Public Const iSemiColon As Integer = 485              ' [;] special semi-colon for FOR statements
Public Const iComma As Integer = 349                  ' [,] comma
Public Const iDot As Integer = 221                    ' [.] decimal place
Public Const iPi As Integer = 229                     ' Pi
Public Const iEp As Integer = 357                     ' e
Public Const iNOP As Integer = 295                    ' NOP
Public Const iPop As Integer = 270                    ' Pop
Public Const iAll As Integer = 333                    ' All
Public Const iDBG As Integer = 344                    ' DBG
Public Const iGfree As Integer = 245                  ' Gfree
Public Const iPrint As Integer = 224                  ' Print
Public Const iPrintx As Integer = 352                 ' Print;
Public Const iAdv As Integer = 351                    ' Adv
Public Const iReset As Integer = 147                  ' Reset
Public Const iPlot As Integer = 211                   ' Plot
Public Const iDelse As Integer = 373                  ' Delse
Public Const iEDef As Integer = 244                   ' Edef
Public Const iRead As Integer = 318                   ' Read
Public Const iWrite As Integer = 319                  ' Write
Public Const iGet As Integer = 323                    ' Get
Public Const iPut As Integer = 324                    ' Put
Public Const iUSR As Integer = 263                    ' USR
Public Const iLen As Integer = 346                    ' Len
Public Const iAs As Integer = 342                     ' As
Public Const iWith As Integer = 348                   ' With
Public Const iIncr As Integer = 329                   ' Incr
Public Const iDecr As Integer = 330                   ' Decr
Public Const iSTO As Integer = 141                    ' STO
Public Const iRCL As Integer = 142                    ' RCL
Public Const iEXC As Integer = 143                    ' EXC
Public Const iSUM As Integer = 144                    ' SUM
Public Const iMUL As Integer = 145                    ' MUL
Public Const iSUB As Integer = 272                    ' SUB
Public Const iDIV As Integer = 273                    ' DIV
Public Const iClrVar As Integer = 240                 ' ClrVar
Public Const iTrim As Integer = 309                   ' Trim
Public Const iLTrim As Integer = 310                  ' LTrim
Public Const iRTrim As Integer = 311                  ' RTrim
Public Const iLSet As Integer = 335                   ' LSet
Public Const iRSet As Integer = 336                   ' RSet
Public Const iReDim As Integer = 368                  ' ReDim
Public Const iMid As Integer = 369                    ' MID
Public Const iOP As Integer = 135                     ' OP
Public Const iFix As Integer = 177                    ' Fix
Public Const iStFlg As Integer = 162                  ' StFlg
Public Const iIfFlg As Integer = 163                  ' IfFlg
Public Const iRFlg As Integer = 290                   ' RFlg
Public Const inFlg As Integer = 291                   ' !Flg
Public Const iDsz As Integer = 331                    ' Dsz
Public Const iDsnz As Integer = 332                   ' Dsnz
Public Const iSin As Integer = 155                    ' Sin
Public Const iCos As Integer = 156                    ' Cos
Public Const iTan As Integer = 157                    ' Tan
Public Const iSec As Integer = 283                    ' Sec
Public Const iCsc As Integer = 284                    ' Csc
Public Const iCot As Integer = 285                    ' Cot
Public Const iRem As Integer = 313                    ' Rem
Public Const iRem2 As Integer = 185                   ' '
Public Const iOpen As Integer = 316                   ' Open
Public Const iClose As Integer = 317                  ' Close
Public Const iLOF As Integer = 322                    ' LOF
Public Const iPmt As Integer = 204                    ' Pmt
Public Const iAbs As Integer = 176                    ' Abs
Public Const i1X As Integer = 158                     ' 1/X
Public Const iStyle As Integer = 172                  ' Style
Public Const iPrintf As Integer = 337                 ' Printf
Public Const iVal As Integer = 350                    ' Val
Public Const iHyp As Integer = 282                    ' Hyp
Public Const iArc As Integer = 154                    ' Arc
Public Const iLoad As Integer = 131                   ' Load
Public Const iSave As Integer = 132                   ' Save
Public Const iLapp As Integer = 259                   ' Lapp
Public Const iASCII As Integer = 260                  ' ASCII
'--- Looping
Public Const iFor As Integer = 201                    ' For
Public Const iDo As Integer = 202                     ' Do
Public Const iWhile As Integer = 203                  ' While
Public Const iUntil As Integer = 359                  ' Until
'--- Blocking
Public Const iIff As Integer = 214                    ' If
Public Const iElse As Integer = 215                   ' Else
Public Const iElseIf As Integer = 343                 ' ElseIf
'--- Selecting
Public Const iSelect As Integer = 188                 ' Select
Public Const iSelectT As Integer = 432                ' SelectT
Public Const iCase As Integer = 189                   ' Case
Public Const iCaseElse As Integer = 486               ' Else to Follow Case

'-------------------------------------------------------------------------------
' Enumeration Definitions
'-------------------------------------------------------------------------------

'Recognized varaible types
Public Enum Vtypes
  vNumber               'double precision floating point (128-bit storage)
  vInteger              'long integer (64-bit storage)
  vString               'string of text (variable size storage)
  vChar                 'byte (0-255) (8-bit storage)
End Enum

'Angle conversion factor
Public Enum AngleTypes
  TypDeg                'Degrees
  TypRad                'Radians
  TypGrad               'Grads
  TypMil                'Mil
End Enum

'Numeric Base
Public Enum BaseTypes
  TypDec                'Decimal
  TypHex                'Hexadecimal
  TypOct                'Octal
  TypBin                'Binary
End Enum

'Label types
Public Enum LblTypes
  TypLbl                'Label
  TypSbr                'Subroutine
  TypKey                'Keypad User-Defined key (A-Z)
  TypConst              'Constant
  TypEnum               'Enumeration
  TypStruct             'Structure
End Enum

'Scope of labels, subroutines
Public Enum Scopes
  Pvt   'Private  Visible only to Program (in a module)
  Pub   'Public   Visible to All programs in module and accessing user program
End Enum

'-------------------------------------------------------------------------------
' Structure Definitions
'-------------------------------------------------------------------------------
'
' The following type will be used to store the base (immediate) variable information, such
' as type (Number, Text, etc), program-defined name, optional fixed length for text, and a
' pointer to the primary variable, off of which 1- or 2-D arrays can be assigned. Hence,
' an array element ultimately refers back to this definition to determine base properties.
'
Public Type Variable            'base variable definitions
  VarType As Vtypes             'default type of variable
  VName As String * LabelWidth  'optional Name for Variable
  Vdata As clsVarSto            'Value storage (and multi-dimensional links)
  VdataLen As Integer           'Size of string type data for fixed-length strings
  VuDef As Boolean              'True when User defines a variable (used for listing purposes)
  Vaddr As Integer              'address where variable was user-defined
End Type
'
' This type is used to store user-defined keys, labels, subroutine names, constants, and structures.
'
Public Type Labels                  'label, subroutine, constant, key, and structure names
  LblTyp As LblTypes                'type of label (Lbl, Sbr, Ukey, Const, Enum, or Struct)
  lblName As String * LabelWidth    'Name for label (uppercase)
  lblCmt As String * DisplayWidth   'optional comment (used by user-defined keys and Const storage)
  LblScope As Scopes                'scope of label (Private or Public)
  lblAddr As Integer                'address of label definition
  LblDat As Integer                 'address of data following label
  LblEnd As Integer                 'address of end brace
  LblValue As Long                  'Number used by Enum, and as Index by Struct int StructPl()
  lblUdef As Boolean                'True when User defines a label (used for Ukey definitions)
End Type
'
' Structure items (used by StructPool, below)
'
Public Type StructItms
  SiName As String * LabelWidth     'Structure Item name
  siType As Vtypes                  'Structure Item type
  siLen As Integer                  'Structure Item length
  siOfst As Integer                 'offset within I/O block for data
End Type
'
' This type keeps track of structures
'
Public Type StructPool
  StSiz As Integer                  'full size of structure (accumulated Structure Items) in bytes
  StBuf As String                   'File I/O buffer (set to StSiz)
  StItmCnt As Integer               'number of items in the StItems() array
  StItems() As StructItms           'individual structure items
End Type
'
' this type (structure) is used to keep track of braced data, such as loops
'
Public Type TrkLoop
  LpStart As Integer    'location of the opening brace
  LpTerm As Integer     'location of the end brace (on do...while and do..until, this points
                        'to the termination of the while() or until() condition)
  LpProcess As Integer  'points to the locations of a process, used by For loops. -1 if not used
  LpCond As Integer     'points to the condition, if defined. If -1, (DO loop), assume TRUE condition
  LpSelect As Double    'Select value for testing with the conditional Case statement
  LpTrue As Boolean     'logical flag for use by If-type expressions (flipped by Else/ElseIf)
  LpLoop As Boolean     'True if Looping block, else False
  LpDspReg As Double    'Save DisplayReg during Conditionals and processes
End Type
'
' this type is used to provide file I/O streaming support
'
Public Type FileBufs
  FileNum As Integer    'file number assigned by system
  FileLen As Integer    'fixed block size (user for Block I/O)
  FileRec As Long       'record number currently at
  FileBufT As String    'I/O buffer for Text
End Type
'
' This type is for the call stack, to keep track of where the program should return to, and a
' reference to the object that it will return to. This can also be used to view a call stack.
'
Public Type SbrInvk
  Pgm As Integer        'program number invoke was called from
  SbrIdx As Long        'index into the pgm's Lbls() list
  PgmInst As Integer    'location to return to
  PgmBrcIdx As Integer  'brace index on entry
  PgmLiveInvk As Boolean 'true if invoked from keyboard
End Type

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

