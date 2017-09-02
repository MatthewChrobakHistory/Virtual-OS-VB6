Attribute VB_Name = "MODLogDat"
Public Sub SaveUser(ByVal Username As String, ByVal Password As String)
Dim FExist As Byte
Dim i As String
Dim x As Byte

On Error GoTo errorhandler:

FExist = 0
Call MkDir(App.Path & "\System\Users\" & Username)
FExist = 1

Call MkDir(App.Path & "\System\Users\" & Username & "\Programs\")
For x = 1 To MAX_PROGRAMS
    Call MkDir(App.Path & "\System\Users\" & Username & "\Programs\" & x)
    Open App.Path & "\System\Users\" & Username & "\Programs\" & x & "\Name.txt" For Output As #1
    Print #1, "None"
    Close #1
Next
Call MkDir(App.Path & "\System\Users\" & Username & "\Libraries\")
Call MkDir(App.Path & "\System\Users\" & Username & "\Libraries\Music\")
Call MkDir(App.Path & "\System\Users\" & Username & "\Libraries\Documents\")
Call MkDir(App.Path & "\System\Users\" & Username & "\Libraries\Pictures\")
Call MkDir(App.Path & "\System\Users\" & Username & "\Libraries\Video\")

Open App.Path & "\System\Users\" & Username & "\Username.txt" For Output As #1
Print #1, Username
Close #1

Open App.Path & "\System\Users\" & Username & "\Password.txt" For Output As #1
Print #1, Password
Close #1

errorhandler:

If FExist = 0 Then PGLog.lblError.Caption = "The User already exists. Please select another name."

End Sub

Public Sub Login(ByVal Username As String, ByVal Password As String)
Dim UPassword As String
Dim FExist As Byte

On Error GoTo errorhandler:

FExist = 0

Open App.Path & "\System\Users\" & Username & "\Username.txt" For Input As #1
Input #1, UUsername
Close #1

FExist = 1

errorhandler:

If FExist = 0 Then
    PGLog.lblError.Caption = "User does not exist."
    Exit Sub
End If

Open App.Path & "\System\Users\" & Username & "\Password.txt" For Input As #1
Input #1, UPassword
Close #1

If UPassword <> Password Then
    PGLog.lblError.Caption = "Password did not match."
    Exit Sub
End If

PGLog.lblError.Caption = ""
    
User.Username = Username
User.Password = Password
PGDesk.Show
PGLog.Hide

End Sub
