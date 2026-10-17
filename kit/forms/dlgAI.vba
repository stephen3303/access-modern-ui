Option Compare Database
Option Explicit

' Chat window for KitAIChat and KitAIAssist (BuildAIDialog builds it). Type a
' message and click Send: the reply shows below, and follow-up messages
' continue the same conversation. "Use reply" (shown when the window was
' opened for a field) puts the reply in that field and closes; Copy puts it
' on the clipboard; Settings opens the AI settings. Errors from the service
' show in the status line, in words the user can act on.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.txtAIMsg = KitAISeed()
    SetVisible Me.btnAIUse, KitAIChatHasTarget()
    SetEnabled Me.btnAIUse, False
    SetEnabled Me.btnAICopy, False
    ShowReady
    CenterOnApp Me
    Me.txtAIMsg.SetFocus
    Me.txtAIMsg.SelStart = Len(Nz(Me.txtAIMsg.Text, ""))
    Exit Sub
Err_Handler:
    HandleError "dlgAI.Form_Load"
End Sub

' The seed text, which KitAIChat holds.
Private Function KitAISeed() As String
    KitAISeed = KitAIChatSeed()
End Function

Private Sub ShowReady()
    If KitAIReady() Then
        Me.lblAIStatus.Caption = KitAIProviderName(KitAIProvider()) & ", " & KitAIModel()
    Else
        Me.lblAIStatus.Caption = "Not set up yet: click Settings."
    End If
End Sub

Private Sub btnAISend_Click()
    Dim msg As String, reply As String
    On Error GoTo Fail
    Me.btnAISend.SetFocus                 ' commits the message box
    msg = Trim$(Nz(Me.txtAIMsg.Value, ""))
    If Len(msg) = 0 Then Exit Sub
    Me.lblAIStatus.Caption = "Waiting for the AI service" & ChrW(8230)
    SetEnabled Me.btnAISend, False
    Me.Repaint
    DoCmd.Hourglass True
    reply = KitAIAsk(msg, , True)
    DoCmd.Hourglass False
    Me.txtAIReply = reply
    Me.txtAIMsg = Null
    SetEnabled Me.btnAISend, True
    SetEnabled Me.btnAICopy, True
    SetEnabled Me.btnAIUse, True
    ShowReady
    Me.txtAIMsg.SetFocus
    Exit Sub
Fail:
    Dim num As Long, desc As String
    num = Err.Number: desc = Err.Description
    DoCmd.Hourglass False
    If num <> AI_ERR Then LogError "dlgAI.btnAISend_Click"
    Me.lblAIStatus.Caption = desc
    SetEnabled Me.btnAISend, True
End Sub

Private Sub btnAIUse_Click()
    On Error GoTo Err_Handler
    KitAIChatUse Nz(Me.txtAIReply.Value, "")
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgAI.btnAIUse_Click"
End Sub

Private Sub btnAICopy_Click()
    On Error GoTo Err_Handler
    If KitCopyText(Nz(Me.txtAIReply.Value, "")) Then KitToast "Copied", "success"
    Exit Sub
Err_Handler:
    HandleError "dlgAI.btnAICopy_Click"
End Sub

Private Sub btnAISettings_Click()
    On Error GoTo Err_Handler
    If KitAISettings() Then ShowReady
    Exit Sub
Err_Handler:
    HandleError "dlgAI.btnAISettings_Click"
End Sub

Private Sub btnAIClose_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgAI.btnAIClose_Click"
End Sub
