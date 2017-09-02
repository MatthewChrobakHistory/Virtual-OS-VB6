Attribute VB_Name = "MODDmp"
Public Sub DUMP()
Exit Sub
Open App.Path & "\System\Users\" & Username & "\Username.bin" For Binary As #1
Put #1, , Username
Close #1

Open App.Path & "\System\Users\" & Username & "\Username.bin" For Binary As #1
Get #1, , i
Close #1
MsgBox i

End Sub
