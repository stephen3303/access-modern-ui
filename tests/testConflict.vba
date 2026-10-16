Option Compare Database
Option Explicit
' A bound form that hands its errors to the kit (modKitTests: write conflicts).
Private Sub Form_Error(DataErr As Integer, Response As Integer)
    KitFormError Me, DataErr, Response
End Sub
