Option Compare Database
Option Explicit

' Detail panel of pgList: one item, editable. The page saves a changed item
' before it shows another; Save checks the required fields (item and
' owner), saves now and refreshes the list. Every other save (the page
' moving to another item, or the focus leaving the panel) checks them too,
' in Form_BeforeUpdate, and is refused while one is empty.

' The required inputs, for CheckRequired.
Private Const REQUIRED As String = "txtName=Item;txtOwner=Owner"

' Any save: refuse it while a required field is empty. CheckRequired shows
' its own message and handles its own errors.
Private Sub Form_BeforeUpdate(Cancel As Integer)
    Cancel = Not CheckRequired(Me, REQUIRED)
End Sub

Private Sub txtPct_AfterUpdate()
    On Error GoTo Err_Handler
    Me.Recalc          ' redraw the progress capsule
    Exit Sub
Err_Handler:
    HandleError "subItem.txtPct_AfterUpdate"
End Sub

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    ' Checked here as well as in Form_BeforeUpdate: from a click the focus
    ' can move to the first empty field, which it can't during an update.
    If Not CheckRequired(Me, REQUIRED) Then Exit Sub
    If Me.Dirty Then Me.Dirty = False
    Me.Parent.ItemSaved Me!id
    Exit Sub
Err_Handler:
    HandleError "subItem.btnSave_Click"
End Sub
