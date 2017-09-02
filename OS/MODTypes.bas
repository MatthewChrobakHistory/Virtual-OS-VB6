Attribute VB_Name = "MODTypes"
Option Explicit

Public User As UserRec

Private Type ProgramRec
    Name As String
End Type

Private Type UserRec
    Username As String
    Password As String
    Program(1 To MAX_PROGRAMS) As ProgramRec
End Type
