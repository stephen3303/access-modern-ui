Option Compare Database
Option Explicit
' Test fixture (modKitTests): a bound form that refuses every save, as a
' detail form does while a required field is empty.
Private Sub Form_BeforeUpdate(Cancel As Integer)
    Cancel = True
End Sub
