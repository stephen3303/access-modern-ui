Option Compare Database
Option Explicit

' Detail panel of pgList: one item, editable. The page saves a changed item
' before it shows another; Save checks the required fields (item and
' owner), saves now, refreshes the list and confirms with a toast.

Private Sub txtPct_AfterUpdate()
    On Error GoTo Err_Handler
    Me.Recalc          ' redraw the progress capsule
    Exit Sub
Err_Handler:
    HandleError "subItem.txtPct_AfterUpdate"
End Sub

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    If Not CheckRequired(Me, "txtName=Item;txtOwner=Owner") Then Exit Sub
    If Me.Dirty Then Me.Dirty = False
    Me.Parent.ItemSaved Me!id
    KitToast "Item saved.", "success"
    Exit Sub
Err_Handler:
    HandleError "subItem.btnSave_Click"
End Sub
