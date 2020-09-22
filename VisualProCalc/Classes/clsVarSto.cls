VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVarSto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'This is a very simple class with little support for robust variable protection.
'This class is not meant to be a general-purpose user class, but is designed
'especially for this application, and should be used only within it. In this case,
' the protection is robust enough.
'-------------------------------------------------------------------------------
' Varaible Storage. These may be individual variables, or array elelments
'-------------------------------------------------------------------------------
'
Public VarNum As Double        'storage for Number when variable is TypNumber
Public VarInt As Long          'storage for Integer when variable is TypInteger
Public VarStr As String        'storage for String when variable is TypText
Public VarChar As Byte         'storage for Char when variable is TypChar
Public StPlIdx As Long         'if actually structure, is index into structure pool
Public StItmIdx As Long        'if actually structure, Structure Item Index

Public LnkNum As Long          'link number in a sequential array tangent
Attribute LnkNum.VB_VarMemberFlags = "200"
Attribute LnkNum.VB_VarUserMemId = 0
Attribute LnkNum.VB_VarProcData = ";Position"
Public VarRoot As Long         'root variable number (0-99) this object is associated with

Public LnkPrev As clsVarSto    'Previous item in this Sequential Array Chain, or origin node
Public LnkNext As clsVarSto    'Next Item in this Sequential Array Chain
Public LnkChild As clsVarSto   'child link (allows multi-dimensional arrays)

'*******************************************************************************
' Subroutine Name   : Init
' Purpose           : Reset Values
'*******************************************************************************
Public Sub Init()
  VarNum = 0#             'initialize all possible variable storage values
  VarInt = 0&
  VarStr = vbNullString
  VarChar = 0
  StPlIdx = 0             'reset any structure indexing (Structure Pool index)
  StItmIdx = 0            'Sequential index of item in structure
End Sub

'*******************************************************************************
' Subroutine Name   : Class_Initialize
' Purpose           : Initialize variable storage to default null
'*******************************************************************************
Private Sub Class_Initialize()
  Set LnkPrev = Nothing   'init links to nothing for now
  Set LnkNext = Nothing
  Set LnkChild = Nothing
  LnkNum = 0              'assume base of new array list (at least for now)
  Call Init               'reset potential variable data
End Sub

'*******************************************************************************
' Subroutine Name   : Class_Terminate
' Purpose           : Ensure previous variable in chain is no longer referencing us
'*******************************************************************************
Private Sub Class_Terminate()
  Set Me.LnkChild = Nothing
  Set Me.LnkNext = Nothing
End Sub

'*******************************************************************************
' Function Name     : AddNewVar
' Purpose           : Add a new variable at the end of the current chain
'*******************************************************************************
Public Function AddNewVar() As clsVarSto
  If Me.LnkNext Is Nothing Then             'if we are at the top of this list...
    Set AddNewVar = New clsVarSto           'return pointer to a brand new variable
    AddNewVar.LnkNum = Me.LnkNum + 1        'assign a link number to one greater than current
    AddNewVar.VarRoot = Me.VarRoot          'donate root variable number associated with
    Set Me.LnkNext = AddNewVar              'add new variable to end of this chain
    Set AddNewVar.LnkPrev = Me              'set object's previous sibling as previous link (this object)
  Else
    Set AddNewVar = Me.LnkNext.AddNewVar()  'else scan to end of chain, add, and return new variable
  End If
End Function

'*******************************************************************************
' Function Name     : PntToLast
' Purpose           : Return a pointer to the last link in the chain
'*******************************************************************************
Public Function PntToLast() As clsVarSto
  If Me.LnkNext Is Nothing Then           'if we are the last item in the chain...
    Set PntToLast = Me                    'return our pointer
  Else
    Set PntToLast = Me.LnkNext.PntToLast  'else recurse down and find it, and return that one
  End If
End Function

'*******************************************************************************
' Function Name     : GetMaxDim
' Purpose           : Return the max Dimension of a list
'*******************************************************************************
Public Function GetMaxDim() As Long
  GetMaxDim = Me.PntToLast.LnkNum 'get link number of last item in this sequential chain
End Function

'*******************************************************************************
' Function Name     : PntToLnk
' Purpose           : Return a pointer to the specified array element
'*******************************************************************************
Public Function PntToLnk(ByVal LnkNum As Long) As clsVarSto
  If LnkNum < 0 Then Exit Function
  
  If Me.LnkNum > LnkNum Then                          'if we are too high...
    PntToLnk = Me.LnkPrev.PntToLnk(LnkNum)            'traverse backward
  ElseIf Me.LnkNum = LnkNum Then                      'else if we are the sought link...
    Set PntToLnk = Me                                 'return our pointer
  Else
    If Me.LnkNext Is Nothing Then Exit Function       'if we are not, but the next does not exist
    Set PntToLnk = Me.LnkNext.PntToLnk(LnkNum)        'else continue down the chain, and return the last link
  End If
End Function

'*******************************************************************************
' Function Name     : BuildAry
' Purpose           : Add 1-D Array. Size includes offset 0, so Size=10 sets elements to 0-10
'                   : An array is built from an existing base variable. Also used by Me.AddAry().
'*******************************************************************************
Public Function BuildAry(Size As Long, Redimn As Boolean) As clsVarSto
  Dim Nvar As clsVarSto
  Dim Idx As Long, IdxBase As Long
'
' check for invalid sizing
'
  If Size <= 0 Then
    Set BuildAry = Nothing    'not valid, so return nothing
    Exit Function
  End If
'
' if current link is not the base chain link, then go to it
'
  If CBool(Me.LnkNum) Then                            'if we are not at the bottom of the chain...
    Set BuildAry = Me.LnkPrev.BuildAry(Size, Redimn)  'ensure we start at the bottom
    Exit Function
  End If
'
' we are at the base link, so begin our work
'
  IdxBase = 1                                         'init base index (1 beyond current (0))
'
' If we are redimensioning, we do not want to destroy the contents of elements that
' are not to be removed
'
  If Redimn Then                                      'if we are to REDIM...
    Nvar = Me.PntToLast                               'point to last link in chain
    If Nvar.LnkNum > Size Then                        'if we are reducing the array...
      Set Nvar = Me.PntToLnk(Size)                    'point to link last desired link...
      Set Nvar.LnkNext = Nothing                      'and remove any above it
      Set BuildAry = Me                               'and return a pointer to the base link
      Exit Function
    ElseIf Nvar.LnkNum = Size Then                    'if we are already set to desired size...
      Set BuildAry = Me                               'simply return a pointer to the base link
      Exit Function                                   'we are done
    End If
    IdxBase = Nvar.LnkNum + 1                         'set new base index for creation
  Else
    Set Me.LnkNext = Nothing                          'if not redim, then redo entire array
    Me.Init                                           'and reset storage elements
  End If
'
' now build links in chain...
'
  If CBool(Size) Then                                 'if size is for more than one element (0=Me)
    For Idx = IdxBase To Size                         'list of new elements to add
      Call Me.AddNewVar                               'add a new variable
    Next Idx                                          'build all
  End If
  Set BuildAry = Me                                   'return base of array (me)
End Function

'*******************************************************************************
' Function Name     : AddAry
' Purpose           : Build another new Dim from an exiting array element.
'*******************************************************************************
Public Function AddAry(ByVal Size As Long, ByVal Redimn As Boolean) As clsVarSto
  Dim Nvar As clsVarSto
  Dim Idx As Long
'
' check for invalid sizing
'
  If Size <= 0 Then
    Set AddAry = Nothing                  'not valid, so return nothing
    Exit Function
  End If
'
' we are the base link for this new dimension, so begin our work
'
  If Me.LnkChild Is Nothing Then          'if a child array is not yet defined...
    Set Me.LnkChild = New clsVarSto       'get a pointer to a brand new list tangent (its linknum is 0)
    Set Me.LnkChild.LnkPrev = Me          'set object's parent as the current varaible
    Me.LnkChild.VarRoot = Me.VarRoot      'give it the root variable we are associated with
  End If
  Set AddAry = Me.LnkChild.BuildAry(Size, Redimn) 'now build an array from it
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

