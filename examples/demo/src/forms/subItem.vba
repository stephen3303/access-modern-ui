Option Compare Database
Option Explicit

' Detail panel of pgList: one item, editable. The page saves a changed item
' before it shows another; Save checks the required fields (item and
' owner), saves now and refreshes the list.

' A saved item may have a new due date: redo the date picker's marks.
Private Sub Form_AfterUpdate()
    On Error GoTo Err_Handler
    MarkDueDates
    Exit Sub
Err_Handler:
    HandleError "subItem.Form_AfterUpdate"
End Sub

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
    If Not CheckRequired(Me, "txtName=Item;txtOwner=Owner") Then Exit Sub
    If Me.Dirty Then Me.Dirty = False
    Me.Parent.ItemSaved Me!id
    Exit Sub
Err_Handler:
    HandleError "subItem.btnSave_Click"
End Sub
