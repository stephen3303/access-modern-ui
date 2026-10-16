Option Compare Database
Option Explicit

' Chat window for KitAIChat (BuildAIDialogs builds it). Type a message and click
' Send: the answer is added to the conversation above, which the AI sees on the
' next message until New chat. Copy answer puts the latest answer on the
' clipboard. A problem the user can fix (no key, a refused request) shows in the
' line under the conversation, not in a message box.

Private m_last As String            ' the latest answer

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.Caption = IIf(Len(AIChatTitle()) > 0, AIChatTitle(), "AI")
    Me.lblStatus.Caption = Me.AIStatusLine()
    CenterOnApp Me
    Me.txtMessage.SetFocus
    Exit Sub
Err_Handler:
    LogError "dlgAI.Form_Load"
End Sub

' Which service answers, or why it can't yet.
Public Function AIStatusLine() As String
    If AIKeyMissing() Then
        AIStatusLine = "No key yet: open the AI settings, or set " & AIKeyVar() & "."
    Else
        AIStatusLine = AIProvider() & " / " & AIModel()
    End If
End Function

Private Sub btnSend_Click()
    On Error GoTo Err_Handler
    Dim q As String, a As String
    Me.btnSend.SetFocus                     ' commits what was typed
    q = Trim$(Nz(Me.txtMessage.Value, ""))
    If Len(q) = 0 Then Exit Sub
    Me.lblStatus.ForeColor = Th(CLR_MUTED_FG)
    Me.lblStatus.Caption = "Waiting for the answer" & ChrW(8230)
    SetEnabled Me.btnSend, False
    DoEvents
    a = AIAsk(q, AIChatSystem(), True)
    m_last = a
    Me.txtAnswer.Value = Nz(Me.txtAnswer.Value, "") & IIf(Len(Nz(Me.txtAnswer.Value, "")) > 0, vbCrLf & vbCrLf, "") & _
        "You: " & q & vbCrLf & vbCrLf & "AI: " & a
    Me.txtMessage.Value = Null
    Me.lblStatus.Caption = Me.AIStatusLine()
    SetEnabled Me.btnSend, True
    Me.txtAnswer.SetFocus
    Me.txtAnswer.SelStart = Len(Nz(Me.txtAnswer.Value, ""))
    Me.txtMessage.SetFocus
    Exit Sub
Err_Handler:
    SetEnabled Me.btnSend, True
    If Err.Number = AI_ERR_NUMBER Then
        Me.lblStatus.ForeColor = Th(CLR_ERROR)
        Me.lblStatus.Caption = Err.Description
    Else
        HandleError "dlgAI.btnSend_Click"
    End If
End Sub

Private Sub btnCopy_Click()
    On Error GoTo Err_Handler
    If Len(m_last) = 0 Then Exit Sub
    KitCopyText m_last
    KitToast "Copied", "success"
    Exit Sub
Err_Handler:
    HandleError "dlgAI.btnCopy_Click"
End Sub

Private Sub btnNew_Click()
    On Error GoTo Err_Handler
    AIReset
    m_last = ""
    Me.txtAnswer.Value = Null
    Me.lblStatus.Caption = Me.AIStatusLine()
    Me.txtMessage.SetFocus
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
