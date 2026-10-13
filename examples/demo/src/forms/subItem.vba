Option Compare Database
Option Explicit

' Detail panel of pgList: one item, editable. The page saves a changed item
' before it shows another; Save saves now and refreshes the list.

Private Sub txtPct_AfterUpdate()
    On Error GoTo Err_Handler
    Me.Recalc          ' redraw the progress capsule
    Exit Sub
Err_Handler:
    HandleError "subItem.txtPct_AfterUpdate"
End Sub

' The Flagged switch changed (Toggle's onChange): save the record at once,
' like a settings switch.
Public Sub FlagChanged()
    On Error GoTo Err_Handler
    If Me.Dirty Then Me.Dirty = False
    Exit Sub
Err_Handler:
    HandleError "subItem.FlagChanged"
End Sub

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    If Me.Dirty Then Me.Dirty = False
    Me.Parent.ItemSaved Me!id
    Exit Sub
Err_Handler:
    HandleError "subItem.btnSave_Click"
End Sub
