Option Compare Database
Option Explicit

' Settings window for KitAISettings (BuildAISettingsDialog builds it): the AI
' service, its model and endpoint, the API key for this Access session, and
' which parts of the app (KitAIPart) may use AI. Choosing a service fills in
' its saved or default model and endpoint. Nothing is saved until Save. The
' key is never written anywhere; if the environment already supplies one the
' note under the box says so.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Dim i As Long, keys As String
    Me.cboAIService = KitAIProvider()
    LoadService KitAIProvider()
    Me.lstAIParts.RowSource = KitAIPartList()
    keys = "," & KitAIEnabledParts() & ","
    For i = 0 To Me.lstAIParts.ListCount - 1
        Me.lstAIParts.Selected(i) = (InStr(keys, "," & Me.lstAIParts.ItemData(i) & ",") > 0)
    Next
    If Me.lstAIParts.ListCount = 0 Then Me.lblAIParts.Caption = "This app has no parts that use AI."
    CenterOnApp Me
    Exit Sub
Err_Handler:
    HandleError "dlgAISettings.Form_Load"
End Sub

' Show the saved (or default) model and endpoint for a service.
Private Sub LoadService(ByVal provider As String)
    Me.txtAIModel = KitAIModel(provider)
    Me.txtAIEndpoint = KitAIEndpoint(provider)
    If Not AINeedsKey(provider) Then
        Me.lblAIKeyNote.Caption = "This service does not need a key."
    ElseIf Len(KitAIKey(provider)) > 0 Then
        Me.lblAIKeyNote.Caption = "A key is already available; leave this empty to keep using it."
    Else
        Me.lblAIKeyNote.Caption = "No key yet. Or set the environment variable KIT_AI_KEY_" & UCase$(provider) & "."
    End If
End Sub

Private Sub cboAIService_AfterUpdate()
    On Error GoTo Err_Handler
    If Len(Nz(Me.cboAIService.Value, "")) > 0 Then LoadService CStr(Me.cboAIService.Value)
    Exit Sub
Err_Handler:
    HandleError "dlgAISettings.cboAIService_AfterUpdate"
End Sub

Private Sub btnAISave_Click()
    Dim i As Long, keys As String, provider As String
    On Error GoTo Err_Handler
    Me.btnAISave.SetFocus                 ' commits the boxes
    provider = Nz(Me.cboAIService.Value, "")
    If Len(provider) = 0 Then
        KitMsgBox "Choose a service.", vbExclamation
        Exit Sub
    End If
    ' KitAIConfigure keeps a setting whose box is empty, so an endpoint the
    ' user cleared goes back to the default by saving that default.
    KitAIConfigure provider, Trim$(Nz(Me.txtAIModel.Value, "")), _
        IIf(Len(Trim$(Nz(Me.txtAIEndpoint.Value, ""))) > 0, Trim$(Me.txtAIEndpoint.Value), AIDefaultEndpoint(provider))
    If Len(Trim$(Nz(Me.txtAIKey.Value, ""))) > 0 Then KitAISetKey CStr(Me.txtAIKey.Value)
    For i = 0 To Me.lstAIParts.ListCount - 1
        If Me.lstAIParts.Selected(i) Then keys = keys & IIf(Len(keys) > 0, ",", "") & Me.lstAIParts.ItemData(i)
    Next
    KitAISetParts keys
    AISettingsSaved
    Me.txtAIKey = Null
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgAISettings.btnAISave_Click"
End Sub

Private Sub btnAICancel_Click()
    On Error GoTo Err_Handler
    Me.txtAIKey = Null
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgAISettings.btnAICancel_Click"
End Sub
