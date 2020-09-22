Attribute VB_Name = "modURLandEMail"
Option Explicit
'~modURLandEMail.bas;
'Perform internet hyperlink, or open email app
'*************************************************
' modURLandEMail: The following subroutines are supported:
'
' HyperJump(): Perform internet hyperlink (launch onto a web page)
' SendEMail(): send email to someone (opens the user's email server)
'              (For Hwnd, simply supply the Me.hWnd from your form)
'
'*************************************************
'-------------------------------------------------
' You may wish to ensure an interconnection first by using the
' CheckInternetConnect() function in modCheckInternetConnect
'-------------------------------------------------

'*************************************************
' API calls and declarations
'*************************************************
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOW = 5

'*************************************************
' HyperJump(): Perform internet hyperlink (launch onto a web page)
'*************************************************
Public Sub HyperJump(ByVal URLPath As String)
  Call ShellExecute(0&, vbNullString, URLPath, vbNullString, vbNullString, vbNormalFocus)
End Sub

'*************************************************
' SendEMail(): send email to someone (opens the user's email server)
' (Simply supply the Me.hWnd from your form)
'*************************************************
Public Sub SendEMail(hwnd As Long, EMailAddress As String, Subject As String, _
                     Optional cc As String = vbNullString, _
                     Optional BCC As String = vbNullString, _
                     Optional BODY As String = vbNullString, _
                     Optional ATTACH As String = vbNullString)
  Dim S As String
  
  S = "mailto:" & Trim$(EMailAddress) & "?Subject=" & Trim$(Subject)
  If CBool(Len(Trim$(cc))) Then S = S & "&Cc=" & Trim$(cc)
  If CBool(Len(Trim$(BCC))) Then S = S & "&Bcc=" & Trim$(BCC)
  If CBool(Len(Trim$(ATTACH))) Then S = S & "&Attach=" & Chr(34) & Trim$(ATTACH) & Chr(34)
  If CBool(Len(Trim$(BODY))) Then S = S & "&Body=" & Trim$(BODY)
  Call ShellExecute(hwnd, "open", S, vbNullString, vbNullString, SW_SHOW)
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

