Option Compare Database
Option Explicit

' Steps and sections (see Build_pgSteps). Back and Next move the step
' indicator; clicking a done step goes back to it. The accordion's first
' section holds the items table.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.acc_b1.Form.RecordSource = "SELECT id, name, owner, category, status, Format(due_date,'dd mmm yyyy') AS due " & _
        "FROM demo_item ORDER BY due_date"
    ShowStep
    Exit Sub
Err_Handler:
    HandleError "pgSteps.Form_Load"
End Sub

' Row click in the table (subGrid): nothing to open on this page.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
End Sub

' The step indicator's handler: a done step was clicked.
Public Sub StepBack(ByVal i As Long)
    On Error GoTo Err_Handler
    SetStep Me, "steps", i
    ShowStep
    Exit Sub
Err_Handler:
    HandleError "pgSteps.StepBack"
End Sub

Private Sub btnBack_Click()
    On Error GoTo Err_Handler
    If CurrentStep(Me, "steps") > 1 Then SetStep Me, "steps", CurrentStep(Me, "steps") - 1
    ShowStep
    Exit Sub
Err_Handler:
    HandleError "pgSteps.btnBack_Click"
End Sub

Private Sub btnNext_Click()
    On Error GoTo Err_Handler
    If CurrentStep(Me, "steps") <= 4 Then SetStep Me, "steps", CurrentStep(Me, "steps") + 1
    ShowStep
    Exit Sub
Err_Handler:
    HandleError "pgSteps.btnNext_Click"
End Sub

' What the current step asks for, and which buttons apply.
Private Sub ShowStep()
    Dim i As Long
    i = CurrentStep(Me, "steps")
    Me.lblStepBody.Caption = Choose(i, _
        "Step 1 of 4: describe the item. Next moves on; the circles before the current step turn into check marks.", _
        "Step 2 of 4: pick an owner. Click step 1's check mark to go back to it.", _
        "Step 3 of 4: review what you entered.", _
        "Step 4 of 4: submit. Next marks every step done.", _
        "All four steps are done. Back returns to step 4.")
    SetEnabled Me.btnBack, i > 1
    SetEnabled Me.btnNext, i <= 4
    Me.btnNext.Caption = IIf(i = 4, "Finish", "Next")
End Sub
