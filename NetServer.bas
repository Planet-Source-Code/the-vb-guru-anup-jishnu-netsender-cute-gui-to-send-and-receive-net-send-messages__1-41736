Attribute VB_Name = "NetServer"
Option Explicit

'================================================================================
' © Copyright 2002  AnupJishnu@hotmail.com  All rights reserved.
'================================================================================
' Module Name:      NetServer
' Description:      This module will fill the list of computers from all domains _
                    on the network into teh available computers list.
' Created By:       Anup Jishnu
' Creation Date:    17/12/2002
'================================================================================
'
                  
Public Declare Function NetServerEnum _
    Lib "Netapi32.dll" ( _
    vServername As Any, _
    ByVal lLevel As Long, _
    vBufptr As Any, _
    lPrefmaxlen As Long, _
    lEntriesRead As Long, _
    lTotalEntries As Long, _
    vServerType As Any, _
    ByVal sDomain As String, _
    vResumeHandle As Any) _
    As Long

Public Declare Sub RtlMoveMemory _
    Lib "Kernel32" ( _
    dest As Any, _
    vSrc As Any, _
    ByVal lSize&)

Public Declare Sub lstrcpyW _
    Lib "Kernel32" ( _
    vDest As Any, _
    ByVal sSrc As Any)
    

Public Declare Function NetApiBufferFree _
    Lib "Netapi32.dll" ( _
    ByVal lpBuffer As Long) _
    As Long

Public Type SERVER_INFO_101
    dw_platform_id As Long
    ptr_name As Long
    dw_ver_major As Long
    dw_ver_minor As Long
    dw_type As Long
    ptr_comment As Long
End Type



Public Const SV_TYPE_WORKSTATION = &H1
Public Const SV_TYPE_DOMAIN_ENUM = &H80000000

Public SERVERTYPE  As Long

'Public Function FillDomainTree(lType As Long, tvw As TreeView) As Boolean
Public Function fn_FillDomainTree(lType As Long) As Boolean
         
 Dim lReturn As Long
Dim Server_Info As Long
Dim lEntries As Long
Dim lTotal As Long
Dim lMax As Long
Dim vResume As Variant
Dim tServer_info_101 As SERVER_INFO_101
Dim sServer As String
Dim sDomain As String
Dim lServerInfo101StructPtr As Long
Dim X As Long, i As Long
Dim bBuffer(512) As Byte

    lReturn = NetServerEnum( _
        ByVal 0&, _
        101, _
        Server_Info, _
        lMax, _
        lEntries, _
        lTotal, _
        ByVal lType, _
        sDomain, _
        vResume)

    If lReturn <> 0 Then
        Exit Function
    End If

    X = 1
    lServerInfo101StructPtr = Server_Info

    Do While X <= lTotal
        DoEvents
        RtlMoveMemory _
            tServer_info_101, _
            ByVal lServerInfo101StructPtr, _
            Len(tServer_info_101)

        lstrcpyW bBuffer(0), _
            tServer_info_101.ptr_name

      
        i = 0
        Do While bBuffer(i) <> 0
            sServer = sServer & _
                Chr$(bBuffer(i))
            i = i + 2
            DoEvents
        Loop
        
       Call sb_AddDomainServers(SERVERTYPE, sServer)
        DoEvents
        X = X + 1
            sServer = vbNullString
        lServerInfo101StructPtr = _
            lServerInfo101StructPtr + _
            Len(tServer_info_101)

    Loop

    lReturn = NetApiBufferFree(Server_Info)
        
End Function

Private Sub sb_AddDomainServers(lType As Long, Parentkey As String)

Dim lReturn As Long
Dim Server_Info As Long
Dim lEntries As Long
Dim lTotal As Long
Dim lMax As Long
Dim vResume As Variant
Dim tServer_info_101 As SERVER_INFO_101
Dim sServer As String
Dim sDomain As String
Dim lServerInfo101StructPtr As Long
Dim X As Long, i As Long
Dim bBuffer(512) As Byte

sDomain = StrConv(Parentkey, vbUnicode)

    
    lReturn = NetServerEnum( _
        ByVal 0&, _
        101, _
        Server_Info, _
        lMax, _
        lEntries, _
        lTotal, _
        ByVal lType, _
        sDomain, _
        vResume)

    If lReturn <> 0 Then
        
        Exit Sub
    End If

    X = 1
    lServerInfo101StructPtr = Server_Info

    Do While X <= lTotal
        DoEvents
        RtlMoveMemory _
            tServer_info_101, _
            ByVal lServerInfo101StructPtr, _
            Len(tServer_info_101)

        lstrcpyW bBuffer(0), _
            tServer_info_101.ptr_name


        i = 0
        Do While bBuffer(i) <> 0
            sServer = sServer & _
                Chr$(bBuffer(i))
            i = i + 2
            DoEvents
        Loop
        frmSend.lstAvailableComputers.AddItem sServer
        
        DoEvents
        X = X + 1
            sServer = vbNullString
        lServerInfo101StructPtr = _
            lServerInfo101StructPtr + _
            Len(tServer_info_101)

    Loop

    lReturn = NetApiBufferFree(Server_Info)


End Sub
