Attribute VB_Name = "modLogic"
Public Function ClickIcon(ByVal Index As Long)

With frmMain
            
    If .imgIcon(Index).DataField = "Internet" Then
        frmInternet.Show
        .Hide
        frmInternet.wbExplorer.Navigate .imgIcon(Index).Tag
        MsgBox .imgIcon(Index).Tag
    End If
    
End With

End Function

Public Function GetIconDataField(ByVal Index As Long) As String
GetIconDataField = "None"

With frmMain

    If .imgIcon(Index).DataField = "Internet" Then GetIconDataField = "Internet"

End With

End Function

Public Sub LoadIcons()
Dim i As Long
Dim locL As Long
Dim locT As Long

For i = 0 To MAX_ICONS

    With frmMain
        '////////////
        '\/SetIcons\/
        '////////////
        If .imgIcon(i).DataField = "Internet" Then
            .imgIcon(i).Picture = LoadPicture(App.Path & "\System\Users\" & GetUserName & "\Programs\Internet\Icons\Explorer.ico")
        End If
        
        '\\\\\\\\\\\\
        '\/SetSpace\/
        '\\\\\\\\\\\\
        
        Open App.Path & "\System\Users\" & GetUserName & "\Desktop\" & i & "\locL.txt" For Input As #1
        Input #1, locL
        Close #1
        
        .imgIcon(i).Left = locL
        
        Open App.Path & "\System\Users\" & GetUserName & "\Desktop\" & i & "\locT.txt" For Input As #1
        Input #1, locT
        Close #1
        
        .imgIcon(i).Top = locT
        
    End With

Next

End Sub

Public Sub SaveIcons()
Dim i As Long

With frmMain
    For i = 0 To MAX_ICONS
        'set the tag
        Open App.Path & "\System\Users\" & GetUserName & "\Desktop\" & i & "\TAG.txt" For Output As #1
        Print #1, .imgIcon(i).Tag
        Close #1
        'set the field
        Open App.Path & "\System\Users\" & GetUserName & "\Desktop\" & i & "\FIELD.txt" For Output As #1
        Print #1, .imgIcon(i).DataField
        Close #1
        'set the left
        Open App.Path & "\System\Users\" & GetUserName & "\Desktop\" & i & "\locL.txt" For Output As #1
        Print #1, .imgIcon(i).Left
        Close #1
        Open App.Path & "\System\Users\" & GetUserName & "\Desktop\" & i & "\locT.txt" For Output As #1
        Print #1, .imgIcon(i).Top
        Close #1
    Next
End With
End Sub

Public Function GetUserName() As String

GetUserName = frmControlPanel.lblUserName.Caption

End Function

Public Sub DeleteIcon(ByVal Index As Long)

MsgBox "todo: make the system lol"

End Sub

Public Sub Loggout()

Call SaveIcons

Unload frmInternet
Unload frmMain
frmLogin.Show

End Sub

Public Sub LogOn()
    
    'show the desktop
    frmMain.Show
    Unload frmLogin
    

End Sub
