Option Compare Database
Option Explicit
' Erl check for the smoke test, in its own module so line 8 never moves.

' Fails on line 8 of this file; SourceWithLine (modErrors) should name that line.
Public Function ErrorLineSource() As String
    On Error GoTo Fail
    Err.Raise 5
    Exit Function
Fail:
    ErrorLineSource = SourceWithLine("test")
End Function
