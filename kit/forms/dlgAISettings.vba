Option Compare Database
Option Explicit

' AI settings for KitAISettings (BuildAIDialogs builds it). AI on or off, the
' service, model and address, where the key comes from, and which registered
' parts (AIRegisterPart) use AI. Save writes the choices (AIConfigure,
' AISetEnabled, AISetPartEnabled); a key typed here is kept in memory for this
' session only (AISetKey), never saved. Rows for parts that aren't registered
' stay hidden. Ends "ok" when saved.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Dim i As Long, n As Long, dy As Long
    Me.chkEnabled = AIEnabled()
    Me.cboProvider = AIProvider()
    Me.txtModel = GetSetting(APP_NAME, "AI", "Model", "")
    Me.txtEndpoint = GetSetting(APP_NAME, "AI", "Endpoint", "")
    Me.txtKeyVar = GetSetting(APP_NAME, "AI", "KeyVar", "")
    n = AIPartCount()
    If n > AIUiPartRows() Then n = AIUiPartRows()
    For i = 1 To AIUiPartRows()
        If i <= n Then
            Me("lblPart" & i).Caption = AIPartCaption(i)
            Me("chkPart" & i) = (GetSetting(APP_NAME, "AI", "Part_" & AIPartKey(i), "1") = "1")
        Else
            SetVisible Me("lblPart" & i), False
            SetVisible Me("chkPart" & i), False
        End If
    Next
    ShowKeyState
    CenterOnApp Me
    Exit Sub
Err_Handler:
    LogError "dlgAISettings.Form_Load"
End Sub

' Whether a key is available now, and where it comes from.
Private Sub ShowKeyState()
    Dim v As String
    v = Trim$(Nz(Me.txtKeyVar, ""))
    If Len(v) = 0 Then v = AIDefaultKeyVar(Nz(Me.cboProvider, AI_OPENAI))
    If Len(Environ$(v)) > 0 And Len(v) > 0 Then
        Me.lblKeyState.Caption = "A key is set in " & v & "."
    ElseIf Len(AIKey()) > 0 Then
        Me.lblKeyState.Caption = "A key is set for this session."
    ElseIf Nz(Me.cboProvider, AI_OPENAI) = AI_OLLAMA Then
        Me.lblKeyState.Caption = "This service needs no key."
    Else
        Me.lblKeyState.Caption = "No key found. Set " & IIf(Len(v) > 0, v, "an environment variable") & " and restart, or type one above."
    End If
End Sub

' Choosing a provider shows that provider's variable name as the example.
Private Sub cboProvider_AfterUpdate()
    On Error GoTo Err_Handler
    ShowKeyState
    Exit Sub
Err_Handler:
    HandleError "dlgAISettings.cboProvider_AfterUpdate"
End Sub

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    Dim i As Long, n As Long, key As String
    Me.btnSave.SetFocus                     ' commits the last field typed in
    Me.lblMsg.Caption = ""
    key = Trim$(Nz(Me.txtKey.Value, ""))
    On Error GoTo Refused
    AIConfigure Nz(Me.cboProvider, AI_OPENAI), Nz(Me.txtModel, ""), Nz(Me.txtEndpoint, ""), Nz(Me.txtKeyVar, "")
    On Error GoTo Err_Handler
    AISetEnabled Nz(Me.chkEnabled, False)
    n = AIPartCount()
    If n > AIUiPartRows() Then n = AIUiPartRows()
    For i = 1 To n
        AISetPartEnabled AIPartKey(i), Nz(Me("chkPart" & i), False)
    Next
    If Len(key) > 0 Then AISetKey key
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Refused:
    ' A choice the user can fix (a bad address, a key typed as a variable name).
    Me.lblMsg.Caption = Err.Description
    Exit Sub
Err_Handler:
    HandleError "dlgAISettings.btnSave_Click"
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgAISettings.btnCancel_Click"
End Sub
