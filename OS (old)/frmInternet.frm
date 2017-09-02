VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmInternet 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   19650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   19455
      Begin VB.PictureBox picBack 
         Height          =   450
         Left            =   120
         ScaleHeight     =   390
         ScaleWidth      =   390
         TabIndex        =   6
         Top             =   240
         Width           =   450
      End
      Begin VB.PictureBox picForward 
         Height          =   450
         Left            =   600
         ScaleHeight     =   390
         ScaleWidth      =   390
         TabIndex        =   5
         Top             =   240
         Width           =   450
      End
      Begin VB.PictureBox Picture1 
         Height          =   450
         Left            =   13200
         ScaleHeight     =   390
         ScaleWidth      =   390
         TabIndex        =   4
         Top             =   240
         Width           =   450
      End
      Begin VB.TextBox txtUrl 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Text            =   "Http:/"
         Top             =   360
         Width           =   10575
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "X"
         Height          =   375
         Left            =   18960
         TabIndex        =   2
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   495
         Left            =   12120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Url:"
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblState 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   720
         Width           =   10575
      End
      Begin VB.Label lblCurrentPage 
         Height          =   375
         Left            =   13920
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   3735
      End
   End
   Begin SHDocVwCtl.WebBrowser wbExplorer 
      Height          =   8175
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   19455
      ExtentX         =   34316
      ExtentY         =   14420
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmInternet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()

frmMain.Show

Unload Me

End Sub

Private Sub cmdSearch_Click()

wbExplorer.Navigate txtUrl.Text
lblCurrentPage.Caption = txtUrl.Text

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler:

With Me

    .Top = (Screen.Height - .Height) / 2
    .Left = (Screen.Width - .Width) / 2

End With

ErrorHandler:

Exit Sub

End Sub

Private Sub picBack_Click()

On Error GoTo ErrorHandler:

wbExplorer.GoBack
txtUrl.Text = wbExplorer.LocationURL

ErrorHandler:

Exit Sub

End Sub

Private Sub picForward_Click()

On Error GoTo ErrorHandler:

wbExplorer.GoForward
txtUrl.Text = wbExplorer.LocationURL

ErrorHandler:

Exit Sub

End Sub

Private Sub Picture1_Click()

wbExplorer.Refresh
txtUrl.Text = wbExplorer.LocationURL

End Sub

Private Sub wbExplorer_StatusTextChange(ByVal Text As String)

If wbExplorer.Busy = True Then Exit Sub

If lblCurrentPage.Caption <> wbExplorer.LocationURL Then
    txtUrl.Text = wbExplorer.LocationURL
    lblCurrentPage.Caption = wbExplorer.LocationURL
End If
lblState.Caption = wbExplorer.LocationName

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

    Case vbKeyUp
        MsgBox "LOL!"
        
End Select

End Sub

