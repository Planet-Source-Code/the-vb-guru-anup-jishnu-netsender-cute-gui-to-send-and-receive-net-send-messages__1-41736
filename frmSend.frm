VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSend 
   BackColor       =   &H0093BEE2&
   Caption         =   "NetSender v:1.1 - [Send Net Message]"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   Icon            =   "frmSend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRemoveAllComps 
      BackColor       =   &H00D6E7EF&
      Caption         =   "<<<"
      Height          =   375
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   435
   End
   Begin VB.CommandButton cmdClearAll 
      BackColor       =   &H00D6E7EF&
      Caption         =   "Clear &All"
      Height          =   405
      Left            =   8937
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5325
      Width           =   1170
   End
   Begin VB.CommandButton cmdClearRecvd 
      BackColor       =   &H00D6E7EF&
      Caption         =   "Clear Rec&vd Messages"
      Height          =   405
      Left            =   6774
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5325
      Width           =   1995
   End
   Begin VB.Frame fraRecvd 
      BackColor       =   &H0093BEE2&
      Caption         =   "Messages History"
      ForeColor       =   &H00000066&
      Height          =   2325
      Left            =   4590
      TabIndex        =   18
      Top             =   2880
      Width           =   6885
      Begin RichTextLib.RichTextBox txtRecvd 
         Height          =   2055
         Left            =   90
         TabIndex        =   7
         Top             =   195
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   3625
         _Version        =   393217
         BackColor       =   15726591
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmSend.frx":1272
      End
   End
   Begin VB.CommandButton cmdClearSent 
      BackColor       =   &H00D6E7EF&
      Caption         =   "&Clear Sent Messages"
      Height          =   405
      Left            =   4611
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5325
      Width           =   1995
   End
   Begin VB.Timer timeFlash 
      Interval        =   500
      Left            =   7530
      Top             =   435
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00D6E7EF&
      Caption         =   "&Refresh Available Computers List"
      Height          =   405
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5325
      Width           =   3030
   End
   Begin VB.Frame fraFrom 
      BackColor       =   &H0093BEE2&
      Caption         =   "From"
      ForeColor       =   &H00000066&
      Height          =   645
      Left            =   4590
      TabIndex        =   17
      Top             =   90
      Width           =   3180
      Begin VB.TextBox txtFrom 
         BackColor       =   &H00EFF7FF&
         ForeColor       =   &H00000066&
         Height          =   300
         Left            =   90
         TabIndex        =   5
         Top             =   225
         Width           =   2985
      End
   End
   Begin VB.CommandButton cmdRemoveTarget 
      BackColor       =   &H00D6E7EF&
      Caption         =   "&<<"
      Height          =   375
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2475
      Width           =   435
   End
   Begin VB.CommandButton cmdAddTarget 
      BackColor       =   &H00D6E7EF&
      Caption         =   "&>>"
      Height          =   375
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1725
      Width           =   435
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00D6E7EF&
      Caption         =   "&Quit"
      Height          =   405
      Left            =   10275
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5325
      Width           =   1170
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00D6E7EF&
      Caption         =   "&Send"
      Height          =   405
      Left            =   3273
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5325
      Width           =   1170
   End
   Begin VB.Frame fraComputerList 
      BackColor       =   &H0093BEE2&
      Caption         =   "Available Computers"
      ForeColor       =   &H00000066&
      Height          =   5100
      Left            =   105
      TabIndex        =   15
      Top             =   90
      Width           =   1935
      Begin VB.ListBox lstAvailableComputers 
         BackColor       =   &H00EFF7FF&
         ForeColor       =   &H00000066&
         Height          =   4740
         Left            =   75
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1755
      End
   End
   Begin VB.Frame fraMessage 
      BackColor       =   &H0093BEE2&
      Caption         =   "Send Message"
      ForeColor       =   &H00000066&
      Height          =   2145
      Left            =   4590
      TabIndex        =   14
      Top             =   735
      Width           =   6885
      Begin VB.TextBox txtMessage 
         BackColor       =   &H00EFF7FF&
         ForeColor       =   &H00000066&
         Height          =   1875
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   195
         Width           =   6675
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5010
      Top             =   5235
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSend.frx":12F4
            Key             =   "Mail"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSend.frx":2576
            Key             =   "Msg"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSend.frx":2890
            Key             =   "Default"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraTargetComputers 
      BackColor       =   &H0093BEE2&
      Caption         =   "Target Computers"
      ForeColor       =   &H00000066&
      Height          =   5100
      Left            =   2535
      TabIndex        =   16
      Top             =   90
      Width           =   1935
      Begin VB.ListBox lstTargetComputers 
         BackColor       =   &H00EFF7FF&
         ForeColor       =   &H00000066&
         Height          =   4740
         Left            =   75
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   1755
      End
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      ForeColor       =   &H00000066&
      Height          =   540
      Left            =   7845
      TabIndex        =   19
      Top             =   180
      Width           =   3570
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUp_Restore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuPopUp_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUp_About 
         Caption         =   "&About"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPopUp_Sep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPopUp_Exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' © Copyright 2002  AnupJishnu@hotmail.com  All rights reserved.
'================================================================================
' Module Name:      frmSend
' Description: Thanx to Levi for RTB Hyperlinks.
' Created By:       Anup Jishnu
' Creation Date:    17/12/2002
'================================================================================
'

Dim intLoopCounter As Integer
Dim intWidth As Integer, intHeight As Integer
Dim intCounter As Integer
Dim bolShowMainIcon As Boolean  'To decide if the Main Icon is shown or not.
Dim strAryMsg() As String   'This array will hold the Message broken into parts.
Private Const MSG_LEN = 895 - 15 'The 15 is deducted as we will add "Part 10 of 10" & vbcrlf.

Private Sub cmdAddTarget_Click()
    
    For intLoopCounter = 0 To lstAvailableComputers.ListCount - 1
        If intLoopCounter <= lstAvailableComputers.ListCount - 1 Then
            If lstAvailableComputers.Selected(intLoopCounter) = True Then
                lstTargetComputers.AddItem lstAvailableComputers.List(intLoopCounter)
                lstAvailableComputers.RemoveItem intLoopCounter
                intLoopCounter = intLoopCounter - 1
            End If
        End If
    Next

End Sub

Private Sub cmdClearAll_Click()
    cmdClearRecvd_Click
    cmdClearSent_Click
    cmdRemoveAllComps_Click
End Sub

Private Sub cmdClearRecvd_Click()
    txtRecvd.Text = vbNullString
    lblStatus.Caption = "Status: "
    strRTFString = vbNullString
End Sub

Private Sub cmdClearSent_Click()
    txtMessage.Text = vbNullString
    lblStatus.Caption = "Status: "
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Screen.MousePointer = vbHourglass
    lstAvailableComputers.Clear
    fn_FillDomainTree SV_TYPE_DOMAIN_ENUM
    Screen.MousePointer = Default
End Sub

Private Sub cmdRemoveAllComps_Click()
    For intLoopCounter = 0 To lstTargetComputers.ListCount - 1
        If intLoopCounter <= lstTargetComputers.ListCount - 1 Then
            lstAvailableComputers.AddItem lstTargetComputers.List(intLoopCounter)
            lstTargetComputers.RemoveItem intLoopCounter
            intLoopCounter = intLoopCounter - 1
        End If
    Next
End Sub

Private Sub cmdRemoveTarget_Click()
    
    For intLoopCounter = 0 To lstTargetComputers.ListCount - 1
        If intLoopCounter <= lstTargetComputers.ListCount - 1 Then
            If lstTargetComputers.Selected(intLoopCounter) = True Then
                lstAvailableComputers.AddItem lstTargetComputers.List(intLoopCounter)
                lstTargetComputers.RemoveItem intLoopCounter
                intLoopCounter = intLoopCounter - 1
            End If
        End If
    Next
    
End Sub

Private Sub cmdSend_Click()
    If lstTargetComputers.ListCount > 0 And LenB(txtMessage.Text) > 0 Then
        sb_SendMessage
    Else
        Beep
        lblStatus.Caption = "Status: You have not selected any computers to send message to."
        lstAvailableComputers.SetFocus
        Beep
    End If
End Sub

Private Sub Form_Activate()
    txtRecvd.SelStart = Len(txtRecvd.Text)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Me.WindowState = vbMinimized
    End If
    bolMsgIcon = False
    Shell_NotifyIcon NIM_MODIFY, typNotIData
End Sub

Private Sub Form_Load()
    Dim lngSysMenuHwnd As Long  'system Menu Handle.

    'Allow only one instance to run at a time.
    If App.PrevInstance = True Then
        End
    End If
    
    Me.Icon = imgList.ListImages("Default").Picture
    
    'Add the About in the System Menu of this form.
    lngSysMenuHwnd = GetSystemMenu(Me.hwnd, 0) 'Gets system menu
    AppendMenuByString lngSysMenuHwnd, MF_SEPARATOR, NEW_MENUID, "" 'adds a seperator
    AppendMenuByString lngSysMenuHwnd, MF_STRING, NEW_MENUID + 1, "About NetSender v:" & App.Major & "." & App.Minor 'Adds menu item
    
    
    'Show the StartupScreen.
    frmAbout.Show
    frmAbout.ZOrder 0
    
    'Define the Width and Height of the Main Form.
    intWidth = Me.Width
    intHeight = Me.Height
    
    bolShowMainIcon = True
    bolMsgIcon = False
    
    Screen.MousePointer = vbHourglass
    
    'Get list of computers in the network.
    SERVERTYPE = SV_TYPE_WORKSTATION
    lstAvailableComputers.Clear
    fn_FillDomainTree SV_TYPE_DOMAIN_ENUM
    
    'Get the system name.
    txtFrom.Text = fn_GetLocalSystemName
    
    intDataLength = 1000
    
    'Start hooking for net send windows.
    sb_StartHook
    
    Screen.MousePointer = vbDefault
    
    'Unload Startup screen.
    Unload frmAbout
    
    'Place the Main Form in the system tray.
    sb_PlaceInSysTray
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'************************************************************************
'This procedure will determine when to show the menu when user clicks _
    on the icon on the system tray.
'************************************************************************
    Dim Action As Long
    
    'there are two display modes and we need to find out
    'which one the application is using
    
    If Me.ScaleMode = vbPixels Then
        Action = X
    Else
        Action = X / Screen.TwipsPerPixelX
    End If
    
Select Case Action

    Case WM_LBUTTONDBLCLK 'Left Button Double Click
        Me.WindowState = vbNormal 'put into taskbar
        Me.Show 'show form
        SetForegroundWindow Me.hwnd
    
    Case WM_RBUTTONUP 'Right Button Up
        If Me.Visible = False Then
            SetForegroundWindow Me.hwnd
            PopupMenu mnuPopUp
        End If
    End Select
    
End Sub


Private Sub sb_SendMessage()
'************************************************************************
'This is the function which will actually send the message.
'The function uses the API NetMessageBufferSend to send the message.
'Before the message is sent, it has to be converted to Unicode using the _
        StrConv Function.
'************************************************************************
    Dim lReturnCode As Long
    Dim sUnicodeToName As String
    Dim sUnicodeFromName As String
    Dim sUnicodeMessage As String
    Dim lMessageLength As Long
    Dim intPartCounter As Integer
    
    Screen.MousePointer = vbHourglass
    
    ' Get the local computer name and convert it to Unicode
    sUnicodeFromName = StrConv(txtFrom.Text, vbUnicode)
    
    'Break the message into parts.
    sb_BreakMessage
    
    'Loop thru the Target Computer list and send the message to all selected.
    For intLoopCounter = 0 To lstTargetComputers.ListCount
        If lstTargetComputers.List(intLoopCounter) <> "" Then
            
            Me.Caption = "Sending Message to " & lstTargetComputers.List(intLoopCounter)
            lblStatus.Caption = "Status: Sending Message to " & lstTargetComputers.List(intLoopCounter)
            DoEvents
            
            ' Convert the Target computer name to Unicode
            sUnicodeToName = StrConv(lstTargetComputers.List(intLoopCounter), vbUnicode)
            
            ' Send the message.
            'If message is in parts, send all the parts to the selected computer.
            For intPartCounter = 0 To UBound(strAryMsg)
                lblStatus.Caption = "Status: Part " & CStr(intPartCounter + 1) & " of Message being sent to " & lstTargetComputers.List(intLoopCounter)
                DoEvents
                
                lMessageLength = Len(strAryMsg(intPartCounter))
                lReturnCode = NetMessageBufferSend("", sUnicodeToName, sUnicodeFromName, strAryMsg(intPartCounter), lMessageLength)
                
                If lReturnCode <> 0 Then
                    'If error in sending any part, quit.
                    intPartCounter = UBound(strAryMsg)
                End If
                
            Next intPartCounter
            
            ' Provide some feedback about the send action
            If lReturnCode = 0 Then
                'Net Send Succeeded.
                Me.Caption = "Message successfully sent to " & lstTargetComputers.List(intLoopCounter)
                
                lblStatus.Caption = "Status: Message successfully sent to " & lstTargetComputers.List(intLoopCounter)
                DoEvents
                
                'Display the message sent in the Message Window.
                strRTFString = strRTFString & " \par Sent: \par " & Format$(Date, "dd/mmm/yyyy") & " " & Time & "            \b[" & lstTargetComputers.List(intLoopCounter) & "]\b0 \par \par " & Replace(txtMessage.Text, Chr(13), "\par ", , , vbBinaryCompare) & "\par \par " & MSG_END
                frmSend.txtRecvd.TextRTF = NS_RTF_HEADDER & strRTFString & NS_RTF_FOOTER
                frmSend.txtRecvd.SelStart = Len(frmSend.txtRecvd.Text)
                
            Else
                'Net Send Failed.
                Me.Caption = "Error - " & lstTargetComputers.List(intLoopCounter) & "  Return code: " & CStr(lReturnCode)
                lblStatus.Caption = "Status: There was an ERROR in sending message to " & lstTargetComputers.List(intLoopCounter)
                DoEvents
            End If
        End If
    Next intLoopCounter
    
    Me.Caption = "NetSender v:" & App.Major & "." & App.Minor & " - [Send Net Message]"
    
    ' Default pointer
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Paint()
    bolMsgIcon = False
    Shell_NotifyIcon NIM_MODIFY, typNotIData
    txtRecvd.SelStart = Len(txtRecvd.Text)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'************************************************************************
'In Query Unlload, remove the Tray Icon.
'************************************************************************
    Shell_NotifyIcon NIM_DELETE, typNotIData 'Remove from System Tray.
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Then
        Me.Width = intWidth
        Me.Height = intHeight
    Else
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'************************************************************************
'This procedure Stops teh hooking and unregisters the Hook.
'************************************************************************
    Call RegisterShellHook(hwnd, RSH_DEREGISTER)
    SetWindowLong hwnd, GWL_WNDPROC, OldProc
End Sub


Private Sub mnuPopUp_About_Click()
    frmAbout.bolShowButton = True
    frmAbout.Show vbModal
End Sub

Private Sub mnuPopUp_Exit_Click()
    Unload Me
End Sub

Private Sub mnuPopUp_Restore_Click()
    Me.WindowState = vbNormal
    Me.Show
End Sub


Private Sub sb_PlaceInSysTray()
'************************************************************************
'This procedure will place the form only in the System Tray with an Icon.
'************************************************************************
    Me.Show 'form must be fully visible
    Me.Refresh
    
    'Prepare Both icons for System Tray Use
    
    With typNotIData 'with system tray
        .cbSize = Len(typNotIData)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = imgList.ListImages("Default").Picture 'Use form's icon in tray
        .szTip = "NetSender - [Developed By Anup Jishnu]" & vbNullChar 'tooltip text
    End With
    
    With typFlashNotIData 'with system tray
        .cbSize = Len(typNotIData)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = imgList.ListImages("Msg").Picture     'New Message Flashing Icon
        .szTip = "New Incomming Message" & vbNullChar 'tooltip text
    End With
    
    
    Shell_NotifyIcon NIM_ADD, typNotIData 'add to tray
End Sub


Private Sub timeFlash_Timer()
'************************************************************************
'The following will Flash the System Tray Icon when new messages arrive. _
    It is controlled by the Boolean Variable bolShowMainIcon.
'************************************************************************
    If bolMsgIcon = True Then
        If bolShowMainIcon = True Then
            Shell_NotifyIcon NIM_MODIFY, typNotIData
            bolShowMainIcon = False
        Else
            Shell_NotifyIcon NIM_MODIFY, typFlashNotIData
            bolShowMainIcon = True
        End If
    End If
End Sub

Private Sub txtRecvd_DblClick()
'************************************************************************
'When the user Double Clicks on the Rich Text Box, we need to ensure that _
    1. If the user clicked on the Sender Name, search for the sender name _
        in the Available List and transfer it to Target list, removing all _
        other from Target list. _
    2. If the user did not click on teh sender name, dont do anything. _
    3. If the Computer name is not found, keep the Target List blank.
'************************************************************************
    Dim intStartSpace As Integer
    Dim intEndSpace As Integer
    'Dim strSelText As String
    
    If txtRecvd.SelStart = 0 Then
        txtRecvd.SelStart = 1
    End If
    
    intStartSpace = InStrRev(txtRecvd.Text, " ", txtRecvd.SelStart, vbTextCompare)
    intEndSpace = InStr(txtRecvd.SelStart, txtRecvd.Text, "]", vbTextCompare)
    
    If intEndSpace = 0 Or intStartSpace = 0 Then
        Exit Sub
    End If
    
    If Mid$(txtRecvd.Text, intStartSpace + 1, 1) = "[" And Mid$(txtRecvd.Text, intEndSpace, 1) = "]" Then
        
        txtRecvd.SelStart = intStartSpace + 1
        txtRecvd.SelLength = (intEndSpace - 1) - (intStartSpace + 1)
        
        If InStr(1, txtRecvd.SelText, "[", vbTextCompare) > 0 Or InStr(1, txtRecvd.SelText, "]", vbTextCompare) > 0 Then
            Exit Sub
        End If
        
        'Remove all computers from the Traget List to Available List.
        cmdRemoveAllComps_Click
        
        'Shift the Selected (Reply) computer from Available to Target list.
        sb_ShiftToTarget txtRecvd.SelText
        sb_ShiftToTarget Left$(Right$(txtRecvd.SelText, Len(txtRecvd.SelText) - 1), Len(txtRecvd.SelText) - 2)
        
        txtMessage.Text = vbNullString
        txtMessage.SetFocus
    End If
End Sub

Private Sub sb_ShiftToTarget(strCompName As String)
'************************************************************************
'This procedure will transfer the Selected Computer from Available to Target List.
'************************************************************************
    For intCounter = 0 To lstAvailableComputers.ListCount - 1
        If lstAvailableComputers.List(intCounter) = strCompName Then
            lstTargetComputers.AddItem strCompName
            lstAvailableComputers.RemoveItem intCounter
            Exit Sub
        End If
    Next
End Sub

Private Sub sb_StartHook()
'************************************************************************
'The following code will start the hooking.
'************************************************************************
    'Used for Hooking START
    uRegMsg = RegisterWindowMessage(ByVal "SHELLHOOK")
    Call RegisterShellHook(hwnd, RSH_REGISTER) ' Or RSH_REGISTER_TASKMAN Or RSH_REGISTER_PROGMAN)
    OldProc = GetWindowLong(hwnd, GWL_WNDPROC)
    SetWindowLong hwnd, GWL_WNDPROC, AddressOf WndProc
End Sub

Private Sub sb_BreakMessage()
'************************************************************************
'The function will check if the message has to be broken into parts _
    due to its length. If length of message is more than 1970 charecters, _
    it breaks it up into required parts and return the array.
'Net send internally sends only the first 895 charecters.
'************************************************************************
    
    Dim intMsgLength As Integer
    Dim intMsgStart As Integer
    Dim strTemp As String
    Dim intTempCounter As Integer
    Dim intNumberOfParts As Integer
    
    
    ReDim strAryMsg(0)
    strAryMsg(0) = vbNullString
    intMsgStart = 1
    txtMessage.Text = Trim(txtMessage.Text)
    
    'We have to replace the \ as in the RTF box, if it is _
    single does not get displayed along with the adjoining word.
    txtMessage.Text = Replace(Replace(txtMessage.Text, "\", "\\"), "\\{", "\\\{")
    
    intMsgLength = Len(txtMessage.Text)
    
    If intMsgLength <= MSG_LEN Then
        'If the length of the message is less than 895 charecters, _
        send it directly without breaking.
        strAryMsg(0) = StrConv(txtMessage.Text, vbUnicode)
        Exit Sub
    End If
    
    'Get the number of parts the message will be broken into.
    intNumberOfParts = CInt(intMsgLength / MSG_LEN)
    If intMsgLength Mod MSG_LEN > 0 Then
        intNumberOfParts = intNumberOfParts + 1
    End If
    
    'Format the first Part.
    intTempCounter = 1
    strAryMsg(0) = StrConv("Part " & Format(intTempCounter, "00") & " of " & Format(intNumberOfParts, "00") & vbCrLf & Mid(txtMessage.Text, intMsgStart, MSG_LEN), vbUnicode)
    intMsgStart = intMsgStart + MSG_LEN
    
    'Format the remaining parts.
    Do While intMsgStart < intMsgLength
        ReDim Preserve strAryMsg(intTempCounter)
        strAryMsg(intTempCounter) = StrConv("Part " & Format(intTempCounter + 1, "00") & " of " & Format(intNumberOfParts, "00") & vbCrLf & Mid(txtMessage.Text, intMsgStart, MSG_LEN), vbUnicode)
        intMsgStart = intMsgStart + MSG_LEN
        intTempCounter = intTempCounter + 1
    Loop
    
End Sub
