VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Operating System Login"
   ClientHeight    =   6510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "x"
      Height          =   375
      Left            =   8040
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   4815
      Begin VB.CommandButton cmdEnter 
         Caption         =   "Enter"
         Height          =   495
         Left            =   1440
         TabIndex        =   3
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtPassword 
         Height          =   405
         Left            =   1080
         TabIndex        =   2
         Text            =   "Password"
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox txtUser 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Text            =   "User"
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "Pass:"
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "User:"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Enter your User and Password"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnter_Click()
Dim User As String
Dim Pass As String
Dim i As String

'Make things easier to type out
User = txtUser.Text
Pass = txtPassword.Text

'Make sure the dir is there
If Dir(App.Path & "\System\Users\" & User & "\") <> "" Then

    'Get the user's password
    Open App.Path & "\System\Users\" & User & "\pass.txt" For Input As #1
    Input #1, i
    Close #1
    
    'see if its the same
    If Pass = i Then
    
        'Get the user's username
        Open App.Path & "\System\Users\" & User & "\user.txt" For Input As #1
        Input #1, i
        Close #1
        
        'Log on
        If User = i Then
            Call LogOn
        End If
        
    'password not valid (derp)
    Else
    MsgBox "Password not valid."
    End If
Else

'user not exist (derp)
MsgBox "User does not exist"

End If

End Sub

Private Sub cmdExit_Click()

'unloading the all frms
Unload frmInternet
Unload frmMain
Unload frmControlPanel
Unload Me

End Sub

Private Sub Form_Load()

'center the program
With Me

    .Top = (Screen.Height - .Height) / 2
    .Left = (Screen.Width - .Width) / 2

End With

End Sub

Private Sub txtPassword_GotFocus()

'making backspacing easier and setting the mask char
txtPassword.Text = ""
txtPassword.PasswordChar = "*"

End Sub

Private Sub txtUser_GotFocus()

'making backspacing easier
txtUser.Text = ""

End Sub
