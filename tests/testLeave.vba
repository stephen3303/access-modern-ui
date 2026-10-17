Option Compare Database
Option Explicit
' Test fixture (modKitTests): a shell page that can be left only while
' TempVars("ShellLeave") is True. While TempVars("ShellRefuse") is True it
' fails as a refused save does: KitRefuseSave, then Access's 2465.

Public Function CanLeave() As Boolean
    If Nz(TempVars("ShellRefuse"), False) Then KitRefuseSave: Err.Raise 2465
    CanLeave = Nz(TempVars("ShellLeave"), True)
End Function
