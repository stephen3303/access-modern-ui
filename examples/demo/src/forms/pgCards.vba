Option Compare Database
Option Explicit

' Cards and progress (see Build_pgCards).

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.sfItems.Form.RecordSource = "SELECT id, name, owner, category, status, Format(due_date,'dd mmm yyyy') AS due " & _
        "FROM demo_item ORDER BY due_date"
    Exit Sub
Err_Handler:
    HandleError "pgCards.Form_Load"
End Sub

' Row click in the table (subGrid): nothing to open on this page.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
End Sub

Private Sub btnRoundA_Click()
    On Error GoTo Err_Handler
    Notify "Primary button clicked.", vbInformation
    Exit Sub
Err_Handler:
    HandleError "pgCards.btnRoundA_Click"
End Sub

Private Sub btnRoundB_Click()
    On Error GoTo Err_Handler
    Notify "Outline button clicked.", vbInformation
    Exit Sub
Err_Handler:
    HandleError "pgCards.btnRoundB_Click"
End Sub
