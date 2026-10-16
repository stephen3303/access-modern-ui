Option Compare Database
Option Explicit
' A bound form that stamps its records (modKitTests: StampRecord).
Private Sub Form_BeforeUpdate(Cancel As Integer)
    StampRecord Me
End Sub
