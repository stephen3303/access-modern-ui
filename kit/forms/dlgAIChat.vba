Option Compare Database
Option Explicit

' Chat window (BuildAIDialogs builds it; KitAIChat opens it). Type a message
' and press Send or Enter (Ctrl+Enter starts a new line); the reply is added
' to the transcript above. Each opening starts a new conversation, and the
' reply follows on from what was said earlier in it. While a reply is awaited
' Send reads Cancel, and closing the window cancels the wait first. Copy
' reply puts the last reply on the clipboard; Settings opens the AI settings.

Private m_log As String         ' the transcript
Private m_last As String        ' the last reply
Private m_closing As Boolean    ' Close was asked for while a reply was awaited

Private Sub Form_Load()
    On Error GoTo Err_Handler
    KitAIReset
    Me.lblService.Caption = KitAIServiceSummary()
    Me.txtLog = ""
    Me.lblStatus.Caption = ""
    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.txtMessage = Me.OpenArgs
    CenterOnApp Me
    Me.txtMessage.SetFocus
    Exit Sub
Err_Handler:
    LogError "dlgAIChat.Form_Load"
End Sub

' F1 opens the help on this window (KitHelp).
Public Property Get HelpTopic() As String
    HelpTopic = "ai"
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KitAIF1 KeyCode
End Sub

' Closing while a reply is awaited: cancel the wait first, then close.
Private Sub Form_Unload(Cancel As Integer)
    If KitAIBusy() Then
        Cancel = True
        m_closing = True
        KitAICancel
    End If
End Sub

Private Sub SetStatus(ByVal text As String, ByVal isError As Boolean)
    Me.lblStatus.Caption = text
    Me.lblStatus.ForeColor = IIf(isError, Th(CLR_ERROR), Th(CLR_MUTED_FG))
End Sub

' Add one speaker's turn to the transcript and scroll to the end.
Private Sub AppendLog(ByVal who As String, ByVal text As String)
    If Len(m_log) > 0 Then m_log = m_log & vbCrLf & vbCrLf
    m_log = m_log & who & ":" & vbCrLf & text
    If Len(m_log) > 60000 Then m_log = Right$(m_log, 60000)      ' a text box holds about 64,000 characters
    Me.txtLog = m_log
    Me.txtLog.SetFocus
    Me.txtLog.SelStart = Len(m_log)
End Sub

Private Sub btnSend_Click()
    Dim msg As String, reply As String, problem As String
    On Error GoTo Err_Handler
    If KitAIBusy() Then KitAICancel: Exit Sub           ' Send reads Cancel while waiting
    Me.btnSend.SetFocus                                 ' commits the message
    msg = Trim$(Nz(Me.txtMessage.Value, ""))
    If Len(msg) = 0 Then
        SetStatus "Type a message first.", True
        Me.txtMessage.SetFocus
        Exit Sub
    End If
    Me.btnSend.Caption = "Cancel"
    Me.txtMessage.Enabled = False
    SetStatus "Waiting for the AI...", False
    problem = KitAISend(msg, reply, True)
    If m_closing Then DoCmd.Close acForm, Me.Name, acSaveNo: Exit Sub
    Me.btnSend.Caption = "Send"
    Me.txtMessage.Enabled = True
    If Len(problem) > 0 Then
        SetStatus problem, True
    Else
        AppendLog "You", msg
        AppendLog "AI", reply
        m_last = reply
        Me.txtMessage = Null
        SetStatus "", False
    End If
    Me.txtMessage.SetFocus
    Exit Sub
Err_Handler:
    On Error Resume Next
    Me.btnSend.Caption = "Send"
    Me.txtMessage.Enabled = True
    On Error GoTo 0
    HandleError "dlgAIChat.btnSend_Click"
End Sub

Private Sub btnNew_Click()
    On Error GoTo Err_Handler
    If KitAIBusy() Then Exit Sub
    KitAIReset
    m_log = "": m_last = ""
    Me.txtLog = ""
    SetStatus "", False
    Me.txtMessage.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgAIChat.btnNew_Click"
End Sub

Private Sub btnCopy_Click()
    On Error GoTo Err_Handler
    If Len(m_last) = 0 Then
        SetStatus "There is no reply to copy yet.", False
    ElseIf KitCopyText(m_last) Then
        KitToast "Copied", "success"
    Else
        KitToast "Couldn't copy. Try again.", "warning"
    End If
    Exit Sub
Err_Handler:
    HandleError "dlgAIChat.btnCopy_Click"
End Sub

Private Sub btnSettings_Click()
    On Error GoTo Err_Handler
    If KitAIBusy() Then Exit Sub
    If KitAISettings(False) Then Me.lblService.Caption = KitAIServiceSummary()
    Exit Sub
Err_Handler:
    HandleError "dlgAIChat.btnSettings_Click"
End Sub

Private Sub btnClose_Click()
    On Error GoTo Err_Handler
    If KitAIBusy() Then m_closing = True: KitAICancel: Exit Sub
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgAIChat.btnClose_Click"
End Sub
