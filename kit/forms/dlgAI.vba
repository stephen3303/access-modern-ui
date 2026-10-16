Option Compare Database
Option Explicit

' Chat window for ShowAI (BuildAIDialog builds it). The reply box shows the
' conversation so far; the user types in the box below it and clicks Ask
' (or presses Ctrl+Enter). Each message continues the same conversation
' (AIChat) until New chat. A failed request is shown in the reply box and
' the typed message stays in the box so it can be sent again. Copy reply puts
' the AI's last answer on the clipboard.

Private m_last As String            ' the AI's last answer

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.lblService.Caption = UCase$(AIProvider() & "  " & AIModel())
    Me.txtReply = IIf(AIReady(), "", "The AI is not ready. " & NotReadyReason())
    Me.btnCopy.Enabled = False
    Me.KeyPreview = True
    CenterOnApp Me
    Me.txtPrompt.SetFocus
    Exit Sub
Err_Handler:
    LogError "dlgAI.Form_Load"
End Sub

' Ctrl+Enter sends, as in most chat windows.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And (Shift And acCtrlMask) <> 0 Then
        KeyCode = 0
        btnAsk_Click
    End If
End Sub

Private Sub btnAsk_Click()
    Dim q As String, reply As String
    On Error GoTo Err_Handler
    q = Trim$(Nz(Me.txtPrompt.Value, ""))
    If Len(q) = 0 Then Exit Sub
    If AIReady() = False Then
        Me.txtReply = Me.txtReply & NotReadyReason() & vbCrLf & vbCrLf
        Exit Sub
    End If
    Me.txtPrompt.SetFocus
    Me.btnAsk.Enabled = False
    Me.txtReply = Nz(Me.txtReply, "") & "You: " & q & vbCrLf & vbCrLf & "Thinking" & ChrW(8230)
    Me.Repaint
    reply = AIChat(q, AIChatSystem())
    Me.txtReply = Left$(Me.txtReply, Len(Me.txtReply) - Len("Thinking" & ChrW(8230))) & "AI: " & reply & vbCrLf & vbCrLf
    m_last = reply
    Me.btnCopy.Enabled = True
    Me.txtPrompt = Null
    ScrollToEnd
    Me.btnAsk.Enabled = True
    Exit Sub
Err_Handler:
    Dim msg As String
    msg = Err.Description
    On Error Resume Next
    Me.txtReply = Left$(Me.txtReply, Len(Me.txtReply) - Len("Thinking" & ChrW(8230))) & "Error: " & msg & vbCrLf & vbCrLf
    Me.btnAsk.Enabled = True
End Sub

Private Sub btnCopy_Click()
    On Error GoTo Err_Handler
    If Len(m_last) = 0 Then Exit Sub
    If KitCopyText(m_last) Then KitToast "Copied", "success" Else KitToast "Another program is using the clipboard", "warning"
    Exit Sub
Err_Handler:
    HandleError "dlgAI.btnCopy_Click"
End Sub

Private Sub btnNew_Click()
    On Error GoTo Err_Handler
    AIChatReset
    m_last = ""
    Me.txtReply = ""
    Me.btnCopy.Enabled = False
    Me.txtPrompt.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgAI.btnNew_Click"
End Sub

Private Sub btnClose_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgAI.btnClose_Click"
End Sub

' Why a request can't be sent right now, in a sentence.
Private Function NotReadyReason() As String
    If Len(AIEndpoint()) = 0 Then
        NotReadyReason = "No AI address is set."
    ElseIf ClassRank(AppClassification()) > ClassRank(AIMaxLevel()) Then
        NotReadyReason = "AI is not permitted in a " & AppClassification() & " app."
    ElseIf Len(AIKeyVar()) > 0 Then
        NotReadyReason = "Set the Windows environment variable " & AIKeyVar() & " and restart " & APP_NAME & "."
    Else
        NotReadyReason = "No API key is available."
    End If
End Function

' Show the end of the conversation.
Private Sub ScrollToEnd()
    Me.txtReply.SetFocus
    Me.txtReply.SelStart = Len(Nz(Me.txtReply.Text, ""))
    Me.txtPrompt.SetFocus
End Sub
