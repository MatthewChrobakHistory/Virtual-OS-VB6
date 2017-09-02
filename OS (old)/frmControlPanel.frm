VERSION 5.00
Begin VB.Form frmControlPanel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtNewName 
         Height          =   285
         Left            =   3480
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdChangeName 
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblUserName 
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "User Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChangeName_Click()

If cmdChangeName.Caption = "Change" Then
    txtNewName.Visible = True
    cmdChangeName.Caption = "Save"
Else
    cmdChangeName.Caption = "Change"
    lblUserName.Caption = txtNewName.Text
    txtNewName.Visible = False
    'insert folder changing code here
    
End If

End Sub

Private Sub Form_Load()

cmdChangeName.Caption = "Change"

With Me

    .Top = (Screen.Height - .Height) / 2
    .Left = (Screen.Width - .Width) / 2

End With

End Sub
