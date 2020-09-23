VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H0093BEE2&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2595
   ClientLeft      =   2310
   ClientTop       =   1620
   ClientWidth     =   5730
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1791.115
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H006DA8D8&
      Caption         =   "&OK"
      Height          =   345
      Left            =   4170
      MaskColor       =   &H00D6E7EF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2085
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designed for Windows 2000, XP and NT."
      Height          =   195
      Left            =   915
      TabIndex        =   6
      Top             =   2115
      Width           =   2940
   End
   Begin VB.Label lblMailID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "anupjishnu@hotmail.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2655
      MouseIcon       =   "frmAbout.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1455
      Width           =   1845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anup Jishnu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1995
      TabIndex        =   4
      Top             =   1147
      Width           =   1515
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H0093BEE2&
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 1.1"
      Height          =   225
      Left            =   3555
      TabIndex        =   3
      Top             =   600
      Width           =   930
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   285
      Top             =   90
      Width           =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1356.278
      Y2              =   1356.278
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackColor       =   &H0093BEE2&
      Caption         =   "Developed by:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   930
      TabIndex        =   1
      Top             =   1215
      Width           =   1035
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H006DA8D8&
      Caption         =   "NetSender"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5685
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bolShowButton As Boolean


Option Explicit
'================================================================================
' © Copyright 2002  AnupJishnu@hotmail.com  All rights reserved.
'================================================================================
' Module Name:      frmAbout
' Description:
' Created By:       Anup Jishnu
' Creation Date:    17/12/2002
'================================================================================
'

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblTitle.Width = Me.Width
    Image1.Picture = frmSend.imgList.ListImages("Default").Picture
    
    If bolShowButton = True Then
        cmdOK.Visible = True
    Else
        cmdOK.Visible = False
    End If
End Sub

Private Sub lblMailID_Click()
    Dim IEObj As New SHDocVw.InternetExplorer
    IEObj.Navigate "mailto:anupjishnu@hotmail.com"
    IEObj.Visible = True
End Sub
