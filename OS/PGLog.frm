VERSION 5.00
Begin VB.Form PGLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox textPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox textUsername 
      Height          =   345
      Left            =   2520
      TabIndex        =   2
      Top             =   1920
      Width           =   3015
   End
   Begin VB.CommandButton cmdCreateAccount 
      Caption         =   "Create"
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblError 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   6975
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "PGLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreateAccount_Click()

If textUsername.Text = "" Or textPassword.Text = "" Then Exit Sub

Call MODLogDat.SaveUser(textUsername.Text, textPassword.Text)

End Sub

Private Sub cmdLogin_Click()

Call MODLogDat.Login(textUsername.Text, textPassword.Text)

End Sub

Private Sub Form_Unload(Cancel As Integer)

End

End Sub
