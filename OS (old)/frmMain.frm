VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   Caption         =   "Menu Screen"
   ClientHeight    =   9285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   19770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   8625
      Width           =   19770
      _ExtentX        =   34872
      _ExtentY        =   1164
      ButtonWidth     =   2090
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Loggout"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Control Panel"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Delete"
            Object.Tag             =   ""
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraAdminPanel 
      Caption         =   "Admin Panel"
      Height          =   2535
      Left            =   10440
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CheckBox chkDDM 
         Caption         =   "Desktop Delete Mode?"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Image imgInternet 
      DataField       =   "Internet"
      Height          =   855
      Left            =   2640
      Tag             =   "www.Google.ca"
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   31
      Left            =   9360
      Top             =   5520
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   30
      Left            =   8400
      Top             =   5520
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   29
      Left            =   9360
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   28
      Left            =   8400
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   27
      Left            =   7440
      Top             =   5520
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   26
      Left            =   6480
      Top             =   5520
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   25
      Left            =   7440
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   24
      Left            =   6480
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   23
      Left            =   9360
      Top             =   3600
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   22
      Left            =   8400
      Top             =   3600
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   21
      Left            =   9360
      Top             =   2640
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   20
      Left            =   8400
      Top             =   2640
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   19
      Left            =   7440
      Top             =   3600
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   18
      Left            =   6480
      Top             =   3600
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   17
      Left            =   7440
      Top             =   2640
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   16
      Left            =   6480
      Top             =   2640
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   15
      Left            =   5520
      Top             =   5520
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   14
      Left            =   4560
      Top             =   5520
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   13
      Left            =   5520
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   12
      Left            =   4560
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   11
      Left            =   3600
      Top             =   5520
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   10
      Left            =   2640
      Top             =   5520
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   9
      Left            =   3600
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   8
      Left            =   2640
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   7
      Left            =   5520
      Top             =   3600
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   6
      Left            =   4560
      Top             =   3600
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   5
      Left            =   5520
      Top             =   2640
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   4
      Left            =   4560
      Top             =   2640
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   3
      Left            =   3600
      Top             =   3600
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   2
      Left            =   2640
      Top             =   3600
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Index           =   1
      Left            =   3600
      Top             =   2640
      Width           =   855
   End
   Begin VB.Image imgIcon 
      DataField       =   "Internet"
      Height          =   850
      Index           =   0
      Left            =   2640
      Tag             =   "www.Google.ca"
      Top             =   2640
      Width           =   850
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

'MAKE SURE THE DATA IS SET FIRST
frmControlPanel.Show
frmControlPanel.lblUserName = frmLogin.txtUser

Call LoadIcons

'HIDE THE CONTROL PANEL AS TO NOT FREAK OUT THE USER
frmControlPanel.Hide

'center the program
With Me

    .Top = (Screen.Height - .Height) / 2
    .Left = (Screen.Width - .Width) / 2

End With

End Sub

Private Sub imgIcon_DblClick(Index As Integer)

If chkDDM.Value = 1 Then
    Call DeleteIcon(Index)
    Exit Sub
End If

frmInternet.Show

Call ClickIcon(Index)

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)

If Button.Index = 1 Then
    Call Loggout
End If

If Button.Index = 2 Then
End If

If Button.Index = 3 Then
    If Button.Value = tbrPressed Then
        chkDDM.Value = 1
    Else
        chkDDM.Value = 0
    End If
End If

End Sub
