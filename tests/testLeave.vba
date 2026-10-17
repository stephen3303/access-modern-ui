Option Compare Database
Option Explicit
' Test fixture (modKitTests): a shell page that can be left only while
' TempVars("ShellLeave") is True.

Public Function CanLeave() As Boolean
    CanLeave = Nz(TempVars("ShellLeave"), True)
End Function
