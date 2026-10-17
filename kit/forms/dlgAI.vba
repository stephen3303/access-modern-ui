Option Compare Database
Option Explicit

' AI assistant dialog (BuildAIDialog builds it, AIOpenAssistant opens it): a
' message box, a Send button and the reply above. The conversation continues
' between sends (modKitAI keeps it); New chat starts again. Ctrl+Enter in the
' message box sends. A failure is shown in the reply box, not as an error
' window. The text typed in OpenArgs is the first message.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    CenterOnApp Me
    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.txtMessage = Me.OpenArgs
    If Not AIConfigured() Then Me.txtReply = "The AI service isn't set up yet."
    Me.txtMessage.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgAI.Form_Load"
End Sub

Private Sub btnSend_Click()
    On Error GoTo Err_Handler
    Dim reply As String
    Me.btnSend.SetFocus                         ' commits the typed message
    If AITryAsk(Nz(Me.txtMessage.Value, ""), reply) Then
        Me.txtReply = reply
        Me.txtMessage = Null
    Else
        Me.txtReply = AILastError()
    End If
    Me.txtMessage.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgAI.btnSend_Click"
End Sub

Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And (Shift And acCtrlMask) <> 0 Then KeyCode = 0: btnSend_Click
End Sub

Private Sub btnNew_Click()
    On Error GoTo Err_Handler
    AIReset
    Me.txtReply = Null
    Me.txtMessage = Null
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
