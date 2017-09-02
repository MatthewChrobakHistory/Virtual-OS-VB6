Attribute VB_Name = "MODDeskDat"
Public Sub LoadDesktop()
Dim i As Byte
Dim Name As String

For i = 1 To MAX_ICONS

    PGDesk.picIcon(i).Visible = False

Next

For i = 1 To MAX_PROGRAMS
    Open App.Path & "\System\Users\" & User.Username & "\Programs\" & i & "\Name.txt" For Input As #1
    Input #1, Name
    Close #1
    User.Program(i).Name = Name
Next

End Sub

Public Sub InstallProgram()
Dim Program As String
Dim x As Byte
Dim Freeslot As Byte

x = 0

Program = InputBox("Type the proper name of the program you wish to install.", "Program Installation")

Freeslot = FindFreeProgramSlot

If Freeslot = 0 Then
    MsgBox "Not enough space!", vbCritical
    Exit Sub
End If

User.Program(Freeslot).Name = Program

MsgBox Freeslot
On Error GoTo errorhandler:
Open App.Path & "\System\Users\" & User.Username & "\Programs\" & Freeslot & "\Name.txt" For Output As #1
Print #1, Program
Close #1

x = 1

errorhandler:

If x = 0 Then MsgBox "File does not exist!", vbCritical
End Sub

Public Sub RunProgram()
Dim Program As String
Dim x As Byte

x = 0

Program = InputBox("Type the proper name of the program you wish to run.", "Program Running")

On Error GoTo errorhandler:

Shell App.Path & "\System\Addons\" & Program & "\client.exe", vbNormalFocus
x = 1

errorhandler:

If x = 0 Then MsgBox "File does not exist!", vbCritical

End Sub
