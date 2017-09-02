Attribute VB_Name = "MODLogic"
Public Function FindFreeProgramSlot() As Byte
Dim i As Byte
Dim Name As String

FindFreeProgramSlot = 0

For i = 1 To MAX_PROGRAMS
    Open App.Path & "\System\Users\" & User.Username & "\Programs\" & i & "\Name.txt" For Input As #1
    Input #1, Name
    Close #1
    If Name = "None" Then FindFreeProgramSlot = i
    Exit For
Next

End Function
