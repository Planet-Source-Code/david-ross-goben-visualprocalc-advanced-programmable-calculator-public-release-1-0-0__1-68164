Attribute VB_Name = "modPlayWaveFile"
Option Explicit
'~modPlayWavFile.bas;
'Play wav files
'*******************************************************************************
' modPlayWavFile: The PlayWavFile() function provides access to the mciExecute
'                 function in the standard WinMM.DLL system file. This allows
'                 you to play wave files without having to use the msmci.ocx control.
'
'                 The PlayWavResource() function allows you to play a wav file that
'                 is stored in a resource (RES) file. This is useful if you do not
'                 wish the user to be able to life the wav file from your application
'                 directory. NOTE: This function will not play in the VB development
'                 mode because the App.hInstance value is that for the VB development
'                 environment, and not the actual application. Running this command
'                 from the EXE works fine (kind of trying, I know, but worth the effort).
'
'                 The PlayWavStop() subroutine will stop any sounds that are playing.
'                 This is useful when you are playing a sound in a continuous loop.
'
'                 If you set the optional PlayAsync parameter to TRUE, then the sound
'                 will play while the application is running, otherwise the app will
'                 wait for the sound to finish playing before continuing. The
'                 optional paraneter NoWait, if set to TRUE, will exit the function
'                 without playing the sound if the sound driver is busy. The optional
'                 parameter Playloop, if set to TRUE, will cause the sound to play in
'                 a continuous loop
'EXAMPLE using a file:
'  Call PlayWavFile(ExpandEnvStrings("%windir%\media\whee.wav"))
'
'EXAMPLE using a resource file:
'''  Assume a file called OPERATNL.wav exists in \\Source\root\DavidG\Sounds, create a
'''  file called Test.rc, and add the following line:
'''Operatnl WAVE \\Source\Root\DavidG\Sounds\Operatnl.wav
'''  Compile this to Test.res using the DOS-level command:
'''RC test.rc
''' Alternatively, use the VB Resource Editor, add the wav file, then edit the properties
'''  of the resouce 101, and change the Type to "WAVE" and the ID to "Operatnl", including
'''  the quotes.
''' Now create a new project, add the Test.res file to the project (Project\Add File...),
'''  add a button to the form named Command1, and add the following code to the form:
'Private Sub Command1_Click()
'  Call PlayWavResource("Operatnl")
'End Sub
'''  Compile this code to an executable and run the executable.
'*******************************************************************************

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Const SND_FILENAME = &H20000
Private Const SND_RESOURCE = &H40004
Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOWAIT = &H2000

'*******************************************************************************
' PlayWavFile(): Play sound from a file
'*******************************************************************************
Public Function PlayWavFile(ByVal FileName As String, _
                            Optional PlayAsync As Boolean = False, _
                            Optional PlayLoop As Boolean = False, _
                            Optional NoWait As Boolean = False) As Integer
  Dim flags As Long
  
  If PlayAsync Then
    flags = SND_FILENAME Or SND_SYNC Or SND_NODEFAULT
  Else
    flags = SND_FILENAME Or SND_ASYNC Or SND_NODEFAULT
  End If
  If NoWait Then flags = flags Or SND_NOWAIT        'check for NoWait flag
  If PlayLoop Then flags = flags Or SND_LOOP        'check for continuous play
  
  PlayWavFile = CInt(PlaySound(FileName, 0&, flags))  'play file
End Function

'*******************************************************************************
' PlayWavResource(): Play sound from a resource file
'*******************************************************************************
Public Function PlayWavResource(ByVal SoundName As String, _
                            Optional PlayAsync As Boolean = False, _
                            Optional PlayLoop As Boolean = False, _
                            Optional NoWait As Boolean = False) As Integer
  Dim flags As Long
  
  If PlayAsync Then
    flags = SND_RESOURCE Or SND_SYNC Or SND_NODEFAULT
  Else
    flags = SND_RESOURCE Or SND_ASYNC Or SND_NODEFAULT
  End If
  If NoWait Then flags = flags Or SND_NOWAIT        'check for NoWait flag
  If PlayLoop Then flags = flags Or SND_LOOP        'check for continuous play
  PlayWavResource = CInt(PlaySound(SoundName, App.hInstance, flags))
End Function

'*******************************************************************************
' PlayWavStop(): stop playing any sounds that might be playing
'*******************************************************************************
Public Sub PlayWavStop()
  Call PlaySound(vbNullString, 0&, 0&)
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

