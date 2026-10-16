Option Compare Database
Option Explicit

' Dashboard (see Build_pgDash): stat tiles over the items, and the flagged
' items in a card, with an empty state when none are flagged.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.sfFlagged.Form.RecordSource = "SELECT id, name, owner, category, status, Format(due_date,'dd mmm yyyy') AS due " & _
        "FROM demo_item WHERE flagged ORDER BY due_date"
    ShowFlagged
    Exit Sub
Err_Handler:
    HandleError "pgDash.Form_Load"
End Sub

' Row click in the table (subGrid): nothing to open on this page.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
End Sub

' The flagged table, or the empty state in its place.
Private Sub ShowFlagged()
    Dim n As Long, i As Long
    n = DCount("*", "demo_item", "flagged")
    SetVisible Me.sfFlagged, n > 0
    For i = 0 To 4
        SetVisible Me("hdrFlagged" & i), n > 0
    Next
    ShowEmpty Me, "emptyFlagged", n = 0
End Sub

' The empty state's button: flag the item due soonest, to see the table.
Private Sub emptyFlagged_b_Click()
    On Error GoTo Err_Handler
    CurrentDb.Execute "UPDATE demo_item SET flagged=True WHERE id=" & DMin("id", "demo_item"), dbFailOnError
    Me.sfFlagged.Form.Requery
    Me.Recalc
    ShowFlagged
    Exit Sub
Err_Handler:
    HandleError "pgDash.emptyFlagged_b_Click"
End Sub
