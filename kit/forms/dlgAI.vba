Option Compare Database
Option Explicit

' AI window for ShowAI (BuildAiDialog builds it). The user types a message
' and clicks Send; the reply shows below, and each Send continues the same
' conversation until New chat. Service, Model and Endpoint are saved as they
' are changed (AiConfigure). API key asks for the key with a masked box and
' keeps it in memory for this session only (AiSetKey); it is never saved.
' The list at the bottom holds the interface parts the app declared (AiBind):
' a ticked part may use the AI, and the buttons of ticked parts show on their
' forms. Esc or Close closes it. ShowAI's prefill arrives as OpenArgs.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Dim v As Variant, rows As String
    For Each v In AiProviders()
        rows = rows & IIf(Len(rows) > 0, ";", "") & Quoted(Split(v, "|")(0)) & ";" & Quoted(Split(v, "|")(1))
    Next
    With Me.cboProvider
        .RowSourceType = "Value List"
        .ColumnCount = 2
        .ColumnWidths = "0;" & .Width
        .BoundColumn = 1
        .RowSource = rows
    End With
    ShowSettings
    FillParts
    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.txtMessage = Me.OpenArgs
    ShowStatus ""
    CenterOnApp Me
    Me.txtMessage.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgAI.Form_Load"
End Sub

' Show the saved service, model and endpoint.
Private Sub ShowSettings()
    Me.cboProvider = AiProvider()
    Me.txtModel = AiModel()
    Me.txtEndpoint = AiEndpoint()
End Sub

' List the declared parts with the user's choice ticked.
Private Sub FillParts()
    Dim v As Variant, rows As String, i As Long
    For Each v In AiPartList()
        rows = rows & IIf(Len(rows) > 0, ";", "") & Quoted(Split(v, "|")(0)) & ";" & Quoted(Mid$(v, InStr(v, "|") + 1))
    Next
    Me.lstParts.RowSource = rows
    For i = 0 To Me.lstParts.ListCount - 1
        Me.lstParts.Selected(i) = AiPartEnabled(CStr(Me.lstParts.Column(0, i)))
    Next
    SetEnabled Me.lstParts, Me.lstParts.ListCount > 0
    Me.lblParts.Visible = (Me.lstParts.ListCount > 0)
    Me.lstParts.Visible = (Me.lstParts.ListCount > 0)
End Sub

' A value-list item: quoted so it may hold ; or ".
Private Function Quoted(ByVal s As String) As String
    Quoted = """" & Replace(s, """", """""") & """"
End Function

Private Sub ShowStatus(ByVal s As String)
    Me.lblStatus.Caption = s
End Sub

Private Sub cboProvider_AfterUpdate()
    On Error GoTo Err_Handler
    If Not IsNull(Me.cboProvider) Then AiConfigure provider:=CStr(Me.cboProvider)
    ShowSettings                    ' the new service's own model and endpoint
    Exit Sub
Err_Handler:
    HandleError "dlgAI.cboProvider_AfterUpdate"
End Sub

Private Sub txtModel_AfterUpdate()
    On Error GoTo Err_Handler
    AiConfigure model:=Trim$(Nz(Me.txtModel, ""))
    Me.txtModel = AiModel()         ' empty goes back to the service's default
    Exit Sub
Err_Handler:
    HandleError "dlgAI.txtModel_AfterUpdate"
End Sub

Private Sub txtEndpoint_AfterUpdate()
    On Error GoTo Err_Handler
    If Len(Trim$(Nz(Me.txtEndpoint, ""))) = 0 Then AiUseDefaultEndpoint Else AiConfigure endpoint:=Trim$(Me.txtEndpoint)
    Me.txtEndpoint = AiEndpoint()
    ShowStatus ""
    Exit Sub
Err_Handler:
    ShowStatus Err.Description
    Me.txtEndpoint = AiEndpoint()   ' back to the last good one
End Sub

Private Sub btnKey_Click()
    On Error GoTo Err_Handler
    Dim k As Variant
    k = KitInputBox("Paste the API key. It is kept in memory until Access closes and is never saved." & _
        vbCrLf & "Leave it empty to forget the key.", "API key", "", True)
    If Not IsNull(k) Then
        AiSetKey CStr(k)
        ShowStatus IIf(Len(CStr(k)) > 0, "Key set for this session.", "Key forgotten.")
    End If
    Exit Sub
Err_Handler:
    HandleError "dlgAI.btnKey_Click"
End Sub

Private Sub btnSend_Click()
    On Error GoTo Err_Handler
    Dim answer As String
    If Len(Trim$(Nz(Me.txtMessage, ""))) = 0 Then
        ShowStatus "Type a message first."
        Exit Sub
    End If
    If Not AiReady() Then
        ShowStatus "No API key: choose API key, or set " & AiKeyVar() & "."
        Exit Sub
    End If
    ShowStatus "Waiting for the reply" & ChrW(8230)
    Me.btnSend.Enabled = False
    DoCmd.Hourglass True
    DoEvents
    If AiTry(CStr(Me.txtMessage), answer, , True) Then
        Me.txtResponse = answer
        Me.txtMessage = Null
        ShowStatus ""
    Else
        ShowStatus AiLastError()
    End If
    DoCmd.Hourglass False
    Me.btnSend.Enabled = True
    Me.txtMessage.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgAI.btnSend_Click"
End Sub

Private Sub btnNew_Click()
    On Error GoTo Err_Handler
    AiNewChat
    Me.txtResponse = Null
    Me.txtMessage = Null
    ShowStatus "New conversation."
    Me.txtMessage.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgAI.btnNew_Click"
End Sub

' A click on the list ticks or unticks a part: save every part's choice.
Private Sub lstParts_AfterUpdate()
    On Error GoTo Err_Handler
    Dim i As Long
    For i = 0 To Me.lstParts.ListCount - 1
        AiSetPartEnabled CStr(Me.lstParts.Column(0, i)), Me.lstParts.Selected(i)
    Next
    Exit Sub
Err_Handler:
    HandleError "dlgAI.lstParts_AfterUpdate"
End Sub

Private Sub btnClose_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgAI.btnClose_Click"
End Sub
