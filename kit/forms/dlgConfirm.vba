Option Compare Database
Option Explicit

' Type-to-confirm dialog for KitConfirmTyped (BuildConfirmDialog builds it),
' before a step that can't be undone. Shows the heading in the red band, the
' detail below it, and which word to type; the action button stays disabled
' until the box holds that word exactly (case counts), as GitHub asks for a
' repository's name before deleting it. Esc and Cancel close it unconfirmed;
' Enter does nothing, so a stray key never confirms.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.Caption = KitConfirmTitle()
    Me.lblHeading.Caption = KitConfirmHeading()
    Me.txtDetail = KitConfirmDetail()
    Me.lblAsk.Caption = "To confirm, type " & KitConfirmWord() & " in the box below."
    Me.btnConfirm.Caption = KitConfirmAction()
    SetEnabled Me.btnConfirm, False
    CenterOnApp Me                  ' over Access (modKitUtil)
    Me.txtValue.SetFocus
    Exit Sub
Err_Handler:
    ' The dialog still opens; with the button disabled nothing can be confirmed.
    LogError "dlgConfirm.Form_Load", KitConfirmHeading()
End Sub

' Each key typed: the button is enabled only while the box holds the word.
Private Sub txtValue_Change()
    On Error GoTo Err_Handler
    SetEnabled Me.btnConfirm, KitConfirmMatches(Nz(Me.txtValue.Text, ""))
    Exit Sub
Err_Handler:
    HandleError "dlgConfirm.txtValue_Change"
End Sub

Private Sub btnConfirm_Click()
    On Error GoTo Err_Handler
    ' The click moved the focus off the box, so its Value holds what was typed.
    If Not KitConfirmMatches(Nz(Me.txtValue.Value, "")) Then
        SetEnabled Me.btnConfirm, False
        Exit Sub
    End If
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgConfirm.btnConfirm_Click"
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgConfirm.btnCancel_Click"
End Sub
