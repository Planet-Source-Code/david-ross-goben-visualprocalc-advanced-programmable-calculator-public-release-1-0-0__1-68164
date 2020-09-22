Attribute VB_Name = "modTileFormBackground"
Option Explicit
'~modTileFormBackground.bas;
'Support form background tiling
'*******************************************************************************
' modTileFormBackground - Support form background tiling. You should place a
'                         call to the InitTileFormBackground() subroutine in the
'                         Form_Load event, and a call to PaintTileFormBackground()
'                         in the form's Form_Paint event. Supply the name of the
'                         PictureBox that holds the image you want to paint to
'                         both subroutines, and the form name (or Me) to the
'                         PaintTileFormBackground() subroutine.
'
'                         The PaintTilePicBackground()  routine is used to also
'                         tile the image to another picturebox on the form, such as as
'                         picturebox background acting as a container for controls
'                         (This can also be used instead of PaintTileFormBackground()
'                         if all you want to do is tile to the picturebox container).
'
'                         NOTE: For best results, be sure the picturebox containing
'                               the tiling image is located on the same form.
'EXAMPLE:
'Private Sub Form_Load()
'  InitTileFormBackground Picture1       'init picturebox
'End Sub
'
'Private Sub Form_Paint()
'  PaintTileFormBackground Me, Picture1  'tile the picture onto the form
'End Sub
'
'Private Sub MyPictureBox_Paint()
'  PaintTilePicBackground Me.MyPictureBox, Picture1  'tile the picture onto the picturebox
'End Sub
'*******************************************************************************

'*******************************************************************************
' InitTileFormBackground: Initialize the picture box for use
'*******************************************************************************
Public Sub InitTileFormBackground(MyPicture As PictureBox)
  With MyPicture
    .AutoSize = True          'force picture's autosize on
    .BorderStyle = 0          'set borderstyle to None
    .Visible = False          'hide the picture
  End With
End Sub

'*******************************************************************************
' PaintTileFormBackground: Repaint tiling
'*******************************************************************************
Public Sub PaintTileFormBackground(MyForm As Form, MyPicture As PictureBox)
  Dim i As Long, J As Long
  
  With MyPicture
    For i = 0 To MyForm.ScaleWidth Step .Width      'draw across top
      For J = 0 To MyForm.ScaleHeight Step .Height  'draw across height
        MyForm.PaintPicture .Picture, i, J          'draw a frame
      Next J
    Next i
  End With
End Sub

'*******************************************************************************
' PaintTilePicBackground: Repaint tiling
'*******************************************************************************
Public Sub PaintTilePicBackground(Mypic As PictureBox, MyPicture As PictureBox)
  Dim i As Long, J As Long
  
  With MyPicture
    For i = 0 To Mypic.ScaleWidth Step .Width      'draw across top
      For J = 0 To Mypic.ScaleHeight Step .Height  'draw across height
        Mypic.PaintPicture .Picture, i, J          'draw a frame
      Next J
    Next i
  End With
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

