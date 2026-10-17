Option Compare Database
Option Explicit

' AI settings window (BuildAIDialogs builds it; KitAISettings opens it). Edits
' the settings modKitAI reads: the kind of service, model and address, where
' the API key comes from, standing instructions, the reply limit and the
' timeout. The key itself is never saved: name an environment variable or a
' key file, or type a key for this session only (kept in memory, forgotten
' when Access closes). Choosing a kind of service fills in its usual address,
' model and key variable. Test saves and sends a short message. Ends "ok"
' when saved.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.cboProvider = KitAIValue("Provider")
    Me.txtModel = KitAIValue("Model")
    Me.txtEndpoint = KitAIValue("Endpoint")
    Me.txtKeyEnv = KitAIValue("KeyEnv")
    Me.txtKeyHeader = KitAIValue("KeyHeader")
    Me.txtKeyFile = KitAIValue("KeyFile")
    Me.txtSystem = KitAIValue("SystemPrompt")
    Me.txtMax = KitAIValue("MaxTokens")
    Me.txtTimeout = KitAIValue("TimeoutSec")
    ShowKeyStatus
    CenterOnApp Me
    Me.cboProvider.SetFocus
    Exit Sub
Err_Handler:
    LogError "dlgAISettings.Form_Load"
End Sub

' F1 opens the help on this window (KitHelp).
Public Property Get HelpTopic() As String
    HelpTopic = "ai.settings"
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KitAIF1 KeyCode
End Sub

Private Sub ShowKeyStatus()
    Dim src As String
    src = KitAIKeySource()
    If Len(src) = 0 Then
        Me.lblKeyStatus.Caption = "No API key found yet. A service on this computer needs none."
    Else
        Me.lblKeyStatus.Caption = "The key comes from " & src & "."
    End If
End Sub

Private Sub SetStatus(ByVal text As String, ByVal isError As Boolean)
    Me.lblStatus.Caption = text
    Me.lblStatus.ForeColor = IIf(isError, Th(CLR_ERROR), Th(CLR_MUTED_FG))
End Sub

' A new kind of service brings its own address, model and key variable.
Private Sub cboProvider_AfterUpdate()
    Dim p As String
    On Error GoTo Err_Handler
    p = Nz(Me.cboProvider.Value, "openai")
    Me.txtEndpoint = KitAIDefault(p, "Endpoint")
    Me.txtModel = KitAIDefault(p, "Model")
    Me.txtKeyEnv = KitAIDefault(p, "KeyEnv")
    Me.txtKeyHeader = Null
    Exit Sub
Err_Handler:
    HandleError "dlgAISettings.cboProvider_AfterUpdate"
End Sub

Private Sub btnBrowse_Click()
    Dim path As String
    On Error GoTo Err_Handler
    path = KitPickFile("Choose the API key file", "Text file|*.txt|All files|*.*", "aikey")
    If Len(path) > 0 Then Me.txtKeyFile = path
    Exit Sub
Err_Handler:
    HandleError "dlgAISettings.btnBrowse_Click"
End Sub

' The text of a box, trimmed.
Private Function BoxText(ByVal nm As String) As String
    BoxText = Trim$(Nz(Me.Controls(nm).Value, ""))
End Function

' Check the fields and save them. False, with the reason shown, when one is
' wrong. A value that matches the service's own default is saved as "" so the
' default keeps applying.
Private Function SaveFields() As Boolean
    Dim p As String, ep As String, maxT As String, tmo As String, k As String
    On Error GoTo Err_Handler
    Me.btnSave.SetFocus                                 ' commits the box being typed in
    p = Nz(Me.cboProvider.Value, "")
    ep = BoxText("txtEndpoint")
    maxT = BoxText("txtMax"): tmo = BoxText("txtTimeout")
    If Not KitAIKnownProvider(p) Then SetStatus "Choose the kind of service.", True: Exit Function
    If Len(BoxText("txtModel")) = 0 Then SetStatus "Type the model name.", True: Exit Function
    If Len(ep) = 0 Then SetStatus "Type the service address.", True: Exit Function
    If Left$(LCase$(ep), 8) <> "https://" And Left$(LCase$(ep), 16) <> "http://localhost" And Left$(LCase$(ep), 16) <> "http://127.0.0.1" Then
        SetStatus "The service address must start with https:// (only a service on this computer may use http://).", True
        Exit Function
    End If
    If Not IsNumeric(maxT) Then SetStatus "Longest reply must be a whole number, 0 or more.", True: Exit Function
    If Val(maxT) < 0 Or Val(maxT) <> Int(Val(maxT)) Then SetStatus "Longest reply must be a whole number, 0 or more.", True: Exit Function
    If Not IsNumeric(tmo) Then SetStatus "Wait time must be 5 to 600 seconds.", True: Exit Function
    If Val(tmo) < 5 Or Val(tmo) > 600 Then SetStatus "Wait time must be 5 to 600 seconds.", True: Exit Function

    KitAISaveSetting "Provider", p
    KitAISaveSetting "Endpoint", IIf(ep = KitAIDefault(p, "Endpoint"), "", ep)
    KitAISaveSetting "Model", IIf(BoxText("txtModel") = KitAIDefault(p, "Model"), "", BoxText("txtModel"))
    k = BoxText("txtKeyEnv")
    KitAISaveSetting "KeyEnv", IIf(k = KitAIDefault(p, "KeyEnv"), "", k)
    KitAISaveSetting "KeyHeader", BoxText("txtKeyHeader")
    KitAISaveSetting "KeyFile", BoxText("txtKeyFile")
    KitAISaveSetting "SystemPrompt", Trim$(Nz(Me.txtSystem.Value, ""))
    KitAISaveSetting "MaxTokens", CStr(CLng(Val(maxT)))
    KitAISaveSetting "TimeoutSec", CStr(CLng(Val(tmo)))
    If Len(Nz(Me.txtKey.Value, "")) > 0 Then
        KitAISetKey Me.txtKey.Value
        Me.txtKey = Null                                ' don't leave the key on the screen or in the box
    End If
    ShowKeyStatus
    SaveFields = True
    Exit Function
Err_Handler:
    HandleError "dlgAISettings.SaveFields"
End Function

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    If Not SaveFields() Then Exit Sub
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgAISettings.btnSave_Click"
End Sub

' Save, then send a short message to see that the settings work.
Private Sub btnTest_Click()
    Dim reply As String, problem As String
    On Error GoTo Err_Handler
    If Not SaveFields() Then Exit Sub
    SetStatus "Testing...", False
    DoEvents
    problem = KitAISend("Reply with the single word OK.", reply, False)
    If Len(problem) > 0 Then
        SetStatus problem, True
    Else
        SetStatus "It works. The service replied: " & Left$(reply, 80), False
    End If
    Exit Sub
Err_Handler:
    HandleError "dlgAISettings.btnTest_Click"
End Sub

Private Sub btnClose_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgAISettings.btnClose_Click"
End Sub
