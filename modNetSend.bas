Attribute VB_Name = "modNetSend"
Option Explicit

'================================================================================
' © Copyright 2002  AnupJishnu@hotmail.com  All rights reserved.
'================================================================================
' Module Name:      modNetSend
' Description:
' Created By:       Anup Jishnu
' Creation Date:    17/12/2002
'================================================================================
'


Public bolMsgIcon As Boolean

Public Declare Function NetWkstaGetInfo Lib "Netapi32" (ByVal sServerName$, _
                ByVal lLevel&, vBuffer As Any) As Long
'NetWkstaGetInfo is a fairly simple function that takes only three parameters. _
The first parameter is a Unicode string that contains the name of the server _
on which the function will be executed. A null string specifies that the function _
will run on the local system. The second parameter determines the level of _
information the function will return. NetWkstaGetInfo returns several different _
types of information depending on the value you specify in this parameter. _
However, the NetMessage utility uses only the most basic information, specified _
by passing the value 100. The third parameter is a pointer to a buffer that _
contains a data structure filled with the network information that corresponds _
to the information level specified by the second parameter.


Public Declare Function NetMessageBufferSend Lib "Netapi32" (ByVal sServerName$, ByVal sMsgName$, ByVal sFromName$, ByVal sMessageText$, ByVal lBufferLength&) As Long
'The NetMessageBufferSend function's first four parameters take Unicode strings. _
The first parameter contains the name of the server on which the function will be _
executed. The second parameter contains the message recipient's name. The third _
parameter contains the message sender's name. The fourth parameter contains the _
text of the message to be sent. The fifth parameter, a long integer, identifies _
the length of the message to be sent.

'Because the NetMessageBufferSend function is declared in a DLL, all the string _
parameters are declared with the ByVal keyword. Although ByVal ordinarily causes _
VB to pass a parameter's value, this keyword works differently for passing _
string parameters to an external DLL such as netapi32.dll. Using ByVal to pass _
strings to functions contained in external DLLs causes VB to convert the strings _
from VB string format into the C string format required by most DLLs, including _
netapi32.dll. The fifth parameter also uses the ByVal keyword, but because this _
parameter contains a numeric variable and not a string, ByVal works as you'd _
expect: It causes VB to pass the value of the numeric variable, rather than _
passing a pointer to the variable.


'WKSTA_INFO_100 Structure
Type WKSTA_INFO_100
    wki100_platform_id As Long
    wki100_computername As Long
    wki100_langroup As Long
    wki100_ver_major As Long
    wki100_ver_minor As Long
End Type

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


'=====================================================================
  'Used for Generating RTF Text within the Rich Text Box.
  'Public Const NS_RTF_HEADDER = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fswiss\fcharset0 Arial;}}" & vbCrLf & "\viewkind4\uc1\pard\f0\fs20 "
   Public Const MSG_END = "-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x"
   Public Const NS_RTF_HEADDER = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fswiss\fprq2\fcharset0 Microsoft Sans Serif;}{\f1\fswiss\fcharset0 Arial;}}" & vbCrLf & "{\colortbl ;\red102\green0\blue0;}" & vbCrLf & "\viewkind4\uc1\pard\cf1\f0\fs16 "
   Public Const NS_RTF_FOOTER = "\cf0\f1\fs20\par" & vbCrLf & "}"
   Public strRTFString As String
'========================================================================


'=====================================================================
  'Used To change the system Menu.
Global Const MF_STRING = &H0
Global Const MF_SEPARATOR = &H800
Public Const NEW_MENUID = &H200  'Decimal 512
Public Declare Function GetSystemMenu& Lib "user32" (ByVal hwnd&, ByVal bRevert&)
Public Declare Function AppendMenuByString& Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem$)
'=====================================================================






'GetLocalSystemName Function

Public Function fn_GetLocalSystemName() As String
'************************************************************************
'This function will return the Machine Name on which teh application is running.
'************************************************************************
    Dim lReturnCode As Long
    Dim bBuffer(512) As Byte
    Dim i As Integer
    Dim twkstaInfo100 As WKSTA_INFO_100, lwkstaInfo100 As Long
    Dim lwkstaInfo100StructPtr As Long
    Dim sLocalName As String
    lReturnCode = NetWkstaGetInfo("", 100, lwkstaInfo100)
    lwkstaInfo100StructPtr = lwkstaInfo100
    If lReturnCode = 0 Then
        CopyMemory twkstaInfo100, ByVal lwkstaInfo100StructPtr, Len(twkstaInfo100)
        CopyMemory bBuffer(0), ByVal twkstaInfo100.wki100_computername, 512
        
        ' Get every other byte from Unicode string
        i = 0
        Do While bBuffer(i) <> 0
            sLocalName = sLocalName & Chr$(bBuffer(i))
            i = i + 2
        Loop
        fn_GetLocalSystemName = sLocalName
    End If
End Function




