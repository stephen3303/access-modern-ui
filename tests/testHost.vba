Option Compare Database
Option Explicit
Public callbackGrid As String
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    callbackGrid = grid
End Sub
