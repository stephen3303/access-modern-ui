Option Compare Database
Option Explicit
' A bound form wired to the kit as modKitLocks says (modKitTests: write conflicts).
Private Sub Form_BeforeUpdate(Cancel As Integer)
    KitNoteEdits Me
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    KitFormError Me, DataErr, Response
End Sub
