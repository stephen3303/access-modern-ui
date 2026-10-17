Option Compare Database
Option Explicit

' Window for connecting parts of a form to the AI (BuildAIDialogs builds it;
' KitAIConnectParts opens it, on the form named in its open argument). The
' user picks an open form, then the button that asks the AI, the box holding
' the question, optionally the box that shows the answer, and an instruction.
' Connect saves it (KitAIBindingSave) and points the button at KitAIRun now;
' Disconnect forgets it. Saved connections are applied by each form's own
' KitAIApplyBindings call when it opens. Choosing a button that already does
' something asks first, since a connection replaces its click until the form
' is reopened. Pages hosted in a main window are listed by their own names.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Dim f As Form, c As Control, names As String, start As String
    Me.lblError.Caption = ""
    For Each f In Forms
        If f.Name <> Me.Name Then AddName names, f.Name
        For Each c In f.Controls
            If c.ControlType = acSubform Then
                If HasSubform(c) Then AddName names, c.Form.Name
            End If
        Next
    Next
    Me.cboForm.RowSourceType = "Value List"
    Me.cboForm.RowSource = names
    start = Nz(Me.OpenArgs, "")
    If Len(start) > 0 Then
        Me.cboForm = start
        FormChosen
    End If
    CenterOnApp Me
    Me.cboForm.SetFocus
    Exit Sub
Err_Handler:
    LogError "dlgAIConnect.Form_Load"
End Sub

' F1 opens the help on this window (KitHelp).
Public Property Get HelpTopic() As String
    HelpTopic = "ai.connect"
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KitAIF1 KeyCode
End Sub

Private Function HasSubform(c As Control) As Boolean
    Dim n As String
    On Error Resume Next
    n = c.Form.Name
    HasSubform = (Err.Number = 0)
End Function

' Add a quoted name to a value list unless it is already there.
Private Sub AddName(names As String, ByVal nm As String)
    If InStr(names, """" & nm & """") > 0 Then Exit Sub
    If Len(names) > 0 Then names = names & ";"
    names = names & """" & nm & """"
End Sub

Private Sub SetError(ByVal text As String)
    Me.lblError.Caption = text
    Me.lblError.ForeColor = Th(CLR_ERROR)
End Sub

Private Sub SetNote(ByVal text As String)
    Me.lblError.Caption = text
    Me.lblError.ForeColor = Th(CLR_MUTED_FG)
End Sub

' The chosen form, or Nothing when it has closed.
Private Function ChosenForm() As Object
    Set ChosenForm = KitAIFindForm(Trim$(Nz(Me.cboForm.Value, "")))
End Function

' Fill the three lists from the chosen form's controls: buttons and labels
' can ask the AI, boxes can hold the question, and boxes and labels can show
' the answer.
Private Sub FormChosen()
    Dim frm As Object, c As Control, trig As String, inp As String, outp As String
    Me.lblError.Caption = ""
    Me.cboTrigger = Null: Me.cboInput = Null: Me.cboOutput = Null: Me.txtInstruction = Null
    Set frm = ChosenForm()
    If frm Is Nothing Then
        SetError "That form isn't open."
        Exit Sub
    End If
    For Each c In frm.Controls
        Select Case c.ControlType
            Case acCommandButton: AddName trig, c.Name
            Case acLabel: AddName trig, c.Name: AddName outp, c.Name
            Case acTextBox, acComboBox
                If c.Name <> "txtFocusSink" Then AddName inp, c.Name: AddName outp, c.Name
        End Select
    Next
    SetList Me.cboTrigger, trig
    SetList Me.cboInput, inp
    SetList Me.cboOutput, outp
    ShowSaved frm.Name
End Sub

Private Sub SetList(cbo As ComboBox, ByVal names As String)
    cbo.RowSourceType = "Value List"
    cbo.RowSource = names
End Sub

Private Sub cboForm_AfterUpdate()
    On Error GoTo Err_Handler
    FormChosen
    Exit Sub
Err_Handler:
    HandleError "dlgAIConnect.cboForm_AfterUpdate"
End Sub

' A button that is already connected shows its connection.
Private Sub cboTrigger_AfterUpdate()
    Dim parts As Variant, frm As Object
    On Error GoTo Err_Handler
    Set frm = ChosenForm()
    If frm Is Nothing Then Exit Sub
    parts = KitAIBindingParts(frm.Name, Nz(Me.cboTrigger.Value, ""))
    If IsEmpty(parts) Then Exit Sub
    Me.cboInput = IIf(Len(parts(0)) > 0, parts(0), Null)
    Me.cboOutput = IIf(Len(parts(1)) > 0, parts(1), Null)
    Me.txtInstruction = IIf(Len(parts(2)) > 0, parts(2), Null)
    Exit Sub
Err_Handler:
    HandleError "dlgAIConnect.cboTrigger_AfterUpdate"
End Sub

' The connections saved for the form, one per line.
Private Sub ShowSaved(ByVal formName As String)
    Dim t As Variant, parts As Variant, s As String
    For Each t In KitAIBindingTriggers(formName)
        parts = KitAIBindingParts(formName, CStr(t))
        If Not IsEmpty(parts) Then
            If Len(s) > 0 Then s = s & vbCrLf
            s = s & t & ": asks with " & parts(0) & ", answers in " & IIf(Len(parts(1)) > 0, parts(1), "a message")
        End If
    Next
    Me.txtSaved = IIf(Len(s) > 0, s, "Nothing is connected on this form yet.")
End Sub

Private Sub btnConnect_Click()
    Dim frm As Object, trig As String, inp As String, outp As String
    On Error GoTo Err_Handler
    Me.btnConnect.SetFocus                              ' commits the box being typed in
    Set frm = ChosenForm()
    trig = Nz(Me.cboTrigger.Value, ""): inp = Nz(Me.cboInput.Value, ""): outp = Nz(Me.cboOutput.Value, "")
    If frm Is Nothing Then SetError "Choose a form that is open.": Exit Sub
    If Len(trig) = 0 Then SetError "Choose the button that asks the AI.": Exit Sub
    If Len(inp) = 0 Then SetError "Choose the box with the question.": Exit Sub
    If trig = inp Or trig = outp Then SetError "The button can't also be the question or answer box.": Exit Sub
    If inp = outp Then SetError "The answer needs a different box from the question.": Exit Sub
    If ClickIsBusy(frm, trig) Then
        If KitMsgBox("""" & trig & """ already does something when it is clicked." & vbCrLf & vbCrLf & _
            "Connecting it replaces that until the form is opened again. Connect it anyway?", vbQuestion + vbYesNo, "Connect to AI") <> vbYes Then Exit Sub
    End If
    KitAIBindingSave frm.Name, trig, inp, outp, Trim$(Nz(Me.txtInstruction.Value, ""))
    KitAIApplyBindings frm
    ShowSaved frm.Name
    SetNote trig & " now asks the AI. Open forms apply this at once; others when they open."
    Exit Sub
Err_Handler:
    HandleError "dlgAIConnect.btnConnect_Click"
End Sub

' True when the control already has a click of its own (not one of ours).
Private Function ClickIsBusy(frm As Object, ByVal nm As String) As Boolean
    Dim s As String
    s = frm.Controls(nm).OnClick
    ClickIsBusy = (Len(s) > 0) And (Left$(s, 9) <> "=KitAIRun")
End Function

Private Sub btnDisconnect_Click()
    Dim frm As Object, trig As String
    On Error GoTo Err_Handler
    Set frm = ChosenForm()
    trig = Nz(Me.cboTrigger.Value, "")
    If frm Is Nothing Or Len(trig) = 0 Then SetError "Choose the form and the button to disconnect.": Exit Sub
    If IsEmpty(KitAIBindingParts(frm.Name, trig)) Then SetError trig & " isn't connected.": Exit Sub
    KitAIBindingDelete frm.Name, trig
    ShowSaved frm.Name
    SetNote trig & " is disconnected. It stops asking the AI when the form is opened again."
    Exit Sub
Err_Handler:
    HandleError "dlgAIConnect.btnDisconnect_Click"
End Sub

Private Sub btnClose_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgAIConnect.btnClose_Click"
End Sub
