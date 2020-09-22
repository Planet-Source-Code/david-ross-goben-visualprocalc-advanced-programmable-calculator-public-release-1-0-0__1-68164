Attribute VB_Name = "modGetScreenWorkArea"
Option Explicit
'~modGetScreenWorkArea.bas;
'Obtains the screen working area, or the task bar location
'*******************************************************************************
' modGetScreenWorkArea - The GetScreenWorkArea() subroutine obtains the screen
'                        working area, which is the desktop space not occupied by
'                        the task bar. Normally the task bar is located at the
'                        bottom of the screen, but this can be move to any side.
'                        The results of this function can be used to size or
'                        position forms so that they do not cover up, or is not
'                        covered up by the task bar.
'                      * GetTaskBarPosition() retrieve the location of the windows
'                        task bar, return the Top, Left, Height, and Width of the
'                        task bar in Twips.
'
'EXAMPLE; Size MyForm to the working area of the screen
'  With MyForm
'    GetScreenWorkArea .Left, .Width, .Top, .Height
'  End With
'*******************************************************************************

'*******************************************************************************
' Types, constants, API calls
'*******************************************************************************
Private Declare Function SystemParametersInfoRect Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long
Private Declare Function AppBarMessage Lib "shell32.dll" Alias "SHAppBarMessage" (ByVal dwMessage As Long, pData As APPBARDATA) As Long

Private Const SPI_GETWORKAREA = 48
Private Const ABM_GETTASKBARPOS = &H5

Private Type RECT
  lLeft As Long
  lTop As Long
  lRight As Long
  lBottom As Long
End Type

Private Type APPBARDATA
  cbSize As Long
  hWnd As Long
  uCallbackMessage As Long
  uEdge As Long
  rc As RECT
  lParam As Long '  message specific
End Type

'*******************************************************************************
' GetScreenWorkArea(): Obtain the screen working area, which is
'                      the desktop space not occupied by the task
'                      bar. Normally the task bar is located at
'                      the bottom of the screen, but this can be
'                      move to any side. The results of this function
'                      can be used to size or position forms so that
'                      they do not cover up, or not covered up by the
'                      task bar.
'*******************************************************************************
Public Sub GetScreenWorkArea(ScrLeft As Long, ScrWidth As Long, _
                         ScrTop As Long, ScrHeight As Long)
  Dim rc As RECT
  
  Call SystemParametersInfoRect(SPI_GETWORKAREA, vbNull, rc, 0&)  'get work zrea
'
' now convert to twips
'
  With Screen
    ScrLeft = rc.lLeft * .TwipsPerPixelX                      'set left in twips
    ScrWidth = rc.lRight * .TwipsPerPixelX - ScrLeft          'set width in twips
    ScrTop = rc.lTop * .TwipsPerPixelY                        'set top in twips
    ScrHeight = rc.lBottom * .TwipsPerPixelY - ScrTop         'set height in twips
  End With
End Sub

'*******************************************************************************
' GetTaskBarPosition(): Retrieve the location of the windows task bar, return the
'                       Top, Left, Height, and Width of the task bar in Twips.
'*******************************************************************************
Public Sub GetTaskBarPosition(BarLeft As Long, BarWidth As Long, _
                         BarTop As Long, BarHeight As Long)
  Dim BarData As APPBARDATA
  
  BarData.cbSize = Len(BarData)
  BarData.hWnd = 0
  Call AppBarMessage(ABM_GETTASKBARPOS, BarData)
  With BarData.rc
    BarLeft = .lLeft * Screen.TwipsPerPixelX
    BarWidth = .lRight * Screen.TwipsPerPixelX - BarLeft
    BarTop = .lTop * Screen.TwipsPerPixelY
    BarHeight = .lBottom * Screen.TwipsPerPixelY - BarTop
  End With
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

