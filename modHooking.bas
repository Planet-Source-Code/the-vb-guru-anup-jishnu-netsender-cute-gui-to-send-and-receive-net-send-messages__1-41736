Attribute VB_Name = "modHooking"
Option Explicit

'================================================================================
' © Copyright 2002  AnupJishnu@hotmail.com  All rights reserved.
'================================================================================
' Module Name:      modHooking
' Description:      This module contains fucntions for hooking the window system calls. _
                    It is this module which will actually trap the Net send messages _
                    which we receive.
' Created By:       Anup Jishnu
' Creation Date:    17/12/2002
'================================================================================
'

Public Const SW_HIDE = 0
Public Const GWL_WNDPROC = (-4)
Public Const RSH_DEREGISTER = 0
Public Const RSH_REGISTER = 1
Public Const HSHELL_WINDOWCREATED = 1
Public Const WM_SYSCOMMAND = &H112  'Used the system menu.
Private Const WM_CLOSE = &H10

Dim strData             As String
Dim strRcvdFrom         As String
Dim strRcvdBy           As String
Dim strRcvddate         As String
Dim strRcvdTime         As String

Public OldProc          As Long
Public uRegMsg             As Long
Dim hWndChild           As Long
Dim retHwnd             As Long

Public intDataLength    As Integer
Dim intStart            As Integer
Dim intEnd              As Integer

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Function RegisterShellHook Lib "shell32" Alias "#181" (ByVal hwnd As Long, ByVal nAction As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Function WndProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'************************************************************************
'This procedure is placed as the Window Procedure which will be called _
    by the system whenever there is any activity with regards to any _
    window in the system.
'lParam holds the Window Handle of window to which message is passed.
'We had registered a Window Message - uRegMsg, while starting/creating the Hook.
'************************************************************************
    
    If hwnd = frmSend.hwnd And wParam = NEW_MENUID + 1 And wMsg = WM_SYSCOMMAND Then
        'The user has clicked on the menu "About from frmSend window.
        'Show the about Box.
        frmAbout.bolShowButton = True
        frmAbout.Show vbModal
    End If
    
  If wMsg = uRegMsg Then
     Select Case wParam
            Case HSHELL_WINDOWCREATED
                
                If GetWndText(lParam) = "Messenger Service " Then
                    
                    ShowWindow lParam, SW_HIDE
                    
                    'Get Window Text.
                    strData = String(1000, vbNullChar)
                        
                    'Get Child Window handle with type Static (Label).
                    hWndChild = FindWindowEx(lParam, 0&, "Static", vbNullString)
                    
                    'Get the text of the Child Window (Label).
                    retHwnd = GetWindowText(hWndChild, strData, intDataLength)
                    
                    'Destroy the wondow.
                    'PostMessage lParam, WM_KEYDOWN, VK_RETURN, &H1&
                    PostMessage lParam, WM_CLOSE, 0&, &H1&
                    
                    'Now we have raw data. Format it to get dataa in our format.
                    If retHwnd > 0 Then
                        intStart = 0
                        intEnd = 0
                        strData = Left$(strData, retHwnd)
                        
                        intStart = 14
                        intEnd = InStr(intStart, strData, " ", vbTextCompare)
                        strRcvdFrom = Mid$(strData, intStart, intEnd - intStart)
                
                        intStart = intEnd + 4
                        intEnd = InStr(intStart, strData, " ", vbTextCompare)
                        strRcvdBy = Mid$(strData, intStart, intEnd - intStart)
                
                        intStart = intEnd + 4
                        intEnd = InStr(intStart, strData, " ", vbTextCompare)
                        strRcvddate = Format$(Mid$(strData, intStart, intEnd - intStart), "dd/MMM/yyyy")
                
                        intStart = intEnd + 1
                        intEnd = InStr(intStart, strData, Chr(13), vbTextCompare)
                        strRcvdTime = Mid$(strData, intStart, intEnd - intStart)
                
                        intStart = intEnd + 4
                        intEnd = InStr(intStart, strData, Chr(0), vbTextCompare)
                        'As the RTF box will not display the Enter Charecters properly, we have to replace them with "\par "
                        strData = Replace(Mid$(strData, intStart, Len(strData) - intStart + 1), Chr(13), "\par ", , , vbBinaryCompare)
                        
                        If LenB(frmSend.txtRecvd.Text) > 0 Then
                            'frmSend.txtRecvd.Text = frmSend.txtRecvd.Text & vbCrLf & strRcvddate & " " & strRcvdTime & "       [" & strRcvdFrom & "]"
                            strRTFString = strRTFString & " \par Received: \par " & strRcvddate & " " & strRcvdTime & "            \b[" & strRcvdFrom & "]\b0 \par \par " & strData & "\par \par " & MSG_END
                        Else
                            'frmSend.txtRecvd.Text = strRcvddate & " " & strRcvdTime & " [" & strRcvdFrom & "]" & vbCrLf & strData & vbCrLf & vbCrLf & MSG_END
                            strRTFString = "Received: \par " & strRcvddate & " " & strRcvdTime & "            \b[" & strRcvdFrom & "]\b0 \par \par " & strData & "\par \par " & MSG_END
                        End If
                        
                        'Display teh message received in the Message Box.
                        frmSend.txtRecvd.TextRTF = NS_RTF_HEADDER & strRTFString & NS_RTF_FOOTER
                        frmSend.txtRecvd.SelStart = Len(frmSend.txtRecvd.Text)
                        
                        frmSend.lblStatus.Caption = "Status: New Message Received from " & strRcvdFrom
                        DoEvents
                        
                        bolMsgIcon = True
                    End If
                End If
     End Select
     
  Else
    'Call teh original window Procedure function.
     WndProc = CallWindowProc(OldProc, hwnd, wMsg, wParam, lParam)
  End If
End Function

Private Function GetWndText(hwnd As Long) As String
  Dim k As Long, sName As String
  sName = Space$(128)
  k = GetWindowText(hwnd, sName, 128)
  If k > 0 Then sName = Left$(sName, k) Else sName = "No caption"
  GetWndText = sName
End Function

