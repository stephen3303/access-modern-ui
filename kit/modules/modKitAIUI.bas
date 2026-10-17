Option Compare Database
Option Explicit

' Access UI kit - the AI screens (the connection itself is in modKitAI).
'
'   KitAIChat                  a chat window (dlgAIChat): type a message, read the reply
'   KitAISettings              the settings window (dlgAISettings): service, model,
'                              where the API key comes from, standing instructions
'   KitAIConnectParts          the window (dlgAIConnect) where the user chooses which
'                              parts of a form are connected to the AI
'
' Connected parts. The user picks, for a form, a button (or link label) that
' asks the AI, the box that holds the question, and optionally the box that
' shows the answer (without one, the answer opens in a message) and an
' instruction for the AI. The choice is saved per Windows user under
' APP_NAME. A form applies its saved choices when it opens:
'
'   Private Sub Form_Load()
'       KitAIApplyBindings Me
'   End Sub
'
' which points the button's click at KitAIRun, so no code is written for it.
' Apply replaces the button's own click for that session, so choose a button
' that does nothing else. To offer the window: a "Connect to AI" button with
' =KitAIConnectParts([Form].[Name]) as its click, shown to the people who may
' change the form's layout.
'
' Build the dialogs with the app's others: BuildAIDialogs. Help topics for
' them are in kit\help\ai.txt, to copy into the app's help file. The app
' provides APP_NAME, HandleError and LogError (modErrors).

Public Const AI_CHAT_FORM As String = "dlgAIChat"
Public Const AI_SETTINGS_FORM As String = "dlgAISettings"
Public Const AI_CONNECT_FORM As String = "dlgAIConnect"
Private Const BIND_SECTION As String = "AIBindings"

' ---------------------------------------------------------------------------
' Opening the screens
' ---------------------------------------------------------------------------

' Open the chat window on a new conversation, with firstMessage typed in the
' message box (not sent).
Public Sub KitAIChat(Optional ByVal firstMessage As String = "")
    On Error GoTo Fail
    RunDialog AI_CHAT_FORM, firstMessage
    Exit Sub
Fail:
    HandleError "modKitAIUI.KitAIChat", "", "The AI window couldn't be opened."
End Sub

' Open the settings window. True when the user saved. dimBehind False for a
' window opened from another dialog.
Public Function KitAISettings(Optional ByVal dimBehind As Boolean = True) As Boolean
    On Error GoTo Fail
    KitAISettings = RunDialog(AI_SETTINGS_FORM, , dimBehind)
    Exit Function
Fail:
    HandleError "modKitAIUI.KitAISettings", "", "The AI settings couldn't be opened."
End Function

' Open the window for connecting a form's parts to the AI, on formName (a form
' that is open) when given. For a button's click: =KitAIConnectParts([Form].[Name])
Public Function KitAIConnectParts(Optional ByVal formName As String = "")
    On Error GoTo Fail
    RunDialog AI_CONNECT_FORM, formName
    Exit Function
Fail:
    HandleError "modKitAIUI.KitAIConnectParts", formName, "The window for connecting to the AI couldn't be opened."
End Function

' The service in one line for a window's header: its type, model and whether a
' key was found (never the key).
Public Function KitAIServiceSummary() As String
    Dim s As String
    s = KitAIProviderName(KitAIValue("Provider")) & ", model " & KitAIValue("Model")
    If Len(KitAIKeySource()) = 0 Then s = s & " - no API key found"
    KitAIServiceSummary = s
End Function

' A provider's name for people.
Public Function KitAIProviderName(ByVal provider As String) As String
    Dim p As Variant
    For Each p In KitAIProviders()
        If Split(p, "|")(0) = LCase$(provider) Then KitAIProviderName = Split(p, "|")(1): Exit Function
    Next
    KitAIProviderName = provider
End Function

' F1 in an AI dialog (its Form_KeyDown): open the kit help on the dialog's
' topic. KitHelp is called by name, so an app without modKitHelp still compiles.
Public Sub KitAIF1(KeyCode As Integer)
    If KeyCode <> vbKeyF1 Then Exit Sub
    KeyCode = 0
    On Error Resume Next
    Application.Run "KitHelp"
End Sub

' ---------------------------------------------------------------------------
' Connected parts of a form
' ---------------------------------------------------------------------------

' Save that clicking trigger (a button or label) on form formName sends the text
' of inputCtl to the AI and shows the answer in outputCtl ("" for a message),
' with instruction as the AI's standing instruction for it ("" for the
' SystemPrompt setting). Raises error 5 for a missing form, button or input.
Public Sub KitAIBindingSave(ByVal formName As String, ByVal trigger As String, ByVal inputCtl As String, _
        ByVal outputCtl As String, ByVal instruction As String)
    If Len(formName) = 0 Or Len(trigger) = 0 Or Len(inputCtl) = 0 Then
        Err.Raise 5, , "A connection needs a form, a button and an input box."
    End If
    instruction = Replace(Replace(Replace(instruction, vbCrLf, " "), vbCr, " "), vbLf, " ")
    SaveSetting APP_NAME, BIND_SECTION, formName & "|" & trigger, inputCtl & vbTab & outputCtl & vbTab & instruction
End Sub

' Forget the connection of trigger on form formName.
Public Sub KitAIBindingDelete(ByVal formName As String, ByVal trigger As String)
    On Error Resume Next
    DeleteSetting APP_NAME, BIND_SECTION, formName & "|" & trigger
End Sub

' The connection of trigger on form formName as Array(input, output,
' instruction), or Empty when there is none.
Public Function KitAIBindingParts(ByVal formName As String, ByVal trigger As String) As Variant
    Dim s As String, p() As String
    On Error Resume Next
    s = GetSetting(APP_NAME, BIND_SECTION, formName & "|" & trigger, "")
    On Error GoTo 0
    If Len(s) = 0 Then Exit Function
    p = Split(s, vbTab, 3)
    ReDim Preserve p(0 To 2)
    KitAIBindingParts = p
End Function

' The names of the controls of form formName that have a connection (an array,
' empty when none).
Public Function KitAIBindingTriggers(ByVal formName As String) As Variant
    Dim all As Variant, i As Long, k As String, names As String
    On Error Resume Next
    all = GetAllSettings(APP_NAME, BIND_SECTION)
    On Error GoTo 0
    If IsEmpty(all) Then KitAIBindingTriggers = Array(): Exit Function
    For i = LBound(all, 1) To UBound(all, 1)
        k = all(i, 0)
        If Left$(k, Len(formName) + 1) = formName & "|" Then
            If Len(names) > 0 Then names = names & vbTab
            names = names & Mid$(k, Len(formName) + 2)
        End If
    Next
    If Len(names) = 0 Then KitAIBindingTriggers = Array() Else KitAIBindingTriggers = Split(names, vbTab)
End Function

' Point the click of every connected control on frm at KitAIRun. Call it from
' the form's Load. A saved connection whose control is gone is skipped.
Public Sub KitAIApplyBindings(frm As Object)
    Dim t As Variant
    On Error GoTo Fail
    For Each t In KitAIBindingTriggers(frm.Name)
        If ControlExists(frm, CStr(t)) Then frm.Controls(CStr(t)).OnClick = "=KitAIRun([Form],""" & t & """)"
    Next
    Exit Sub
Fail:
    LogError "modKitAIUI.KitAIApplyBindings", "form " & frm.Name
End Sub

' The click of a connected control: send the text of its input box to the AI
' and show the answer where the connection says. Problems are shown in a
' message, never raised.
Public Function KitAIRun(frm As Form, ByVal trigger As String)
    Dim parts As Variant, msg As String, reply As String, problem As String
    On Error GoTo Err_Handler
    parts = KitAIBindingParts(frm.Name, trigger)
    If IsEmpty(parts) Then KitToast "This isn't connected to the AI.", "warning": Exit Function
    If Not ControlExists(frm, CStr(parts(0))) Then
        KitToast "The AI input box " & parts(0) & " isn't on this form.", "warning"
        Exit Function
    End If
    msg = Trim$(ControlText(frm, CStr(parts(0))))
    If Len(msg) = 0 Then KitToast "Type something for the AI first.", "info": Exit Function
    DoCmd.Hourglass True
    SysCmd acSysCmdSetStatus, "Waiting for the AI..."
    problem = KitAISend(msg, reply, False, CStr(parts(2)))
    DoCmd.Hourglass False
    SysCmd acSysCmdClearStatus
    If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, "AI": Exit Function
    If Len(parts(1)) > 0 And ControlExists(frm, CStr(parts(1))) Then
        SetControlText frm, CStr(parts(1)), reply
    Else
        KitMsgBox reply, vbInformation, "AI reply"
    End If
    Exit Function
Err_Handler:
    DoCmd.Hourglass False
    On Error Resume Next
    SysCmd acSysCmdClearStatus
    On Error GoTo 0
    HandleError "modKitAIUI.KitAIRun", frm.Name & "." & trigger
End Function

' The open form (or a subform in an open form, such as a page in the main
' window) named formName, or Nothing.
Public Function KitAIFindForm(ByVal formName As String) As Object
    Dim f As Form, c As Control
    For Each f In Forms
        If f.Name = formName Then Set KitAIFindForm = f: Exit Function
        For Each c In f.Controls
            If c.ControlType = acSubform Then
                If HasSubform(c) Then
                    If c.Form.Name = formName Then Set KitAIFindForm = c.Form: Exit Function
                End If
            End If
        Next
    Next
End Function

Private Function HasSubform(c As Control) As Boolean
    Dim n As String
    On Error Resume Next
    n = c.Form.Name
    HasSubform = (Err.Number = 0)
End Function

Private Function ControlExists(frm As Object, ByVal nm As String) As Boolean
    Dim c As Control
    On Error Resume Next
    Set c = frm.Controls(nm)
    ControlExists = (Err.Number = 0) And Not c Is Nothing
End Function

' What a connected control holds: a label's caption, or a box's value.
Private Function ControlText(frm As Object, ByVal nm As String) As String
    Dim c As Control
    Set c = frm.Controls(nm)
    If c.ControlType = acLabel Then ControlText = c.Caption Else ControlText = Nz(c.Value, "")
End Function

Private Sub SetControlText(frm As Object, ByVal nm As String, ByVal text As String)
    Dim c As Control
    Set c = frm.Controls(nm)
    If c.ControlType = acLabel Then c.Caption = text Else c.Value = text
End Sub

' ---------------------------------------------------------------------------
' Build
' ---------------------------------------------------------------------------

Public Sub BuildAIDialogs()
    Build_dlgAIChat
    Build_dlgAISettings
    Build_dlgAIConnect
End Sub

Private Sub FieldLabel(ByVal nm As String, ByVal cap As String, ByVal x As Long, ByVal y As Long, ByVal w As Long)
    Lbl nm, cap, x, y, w, px(16), 9, True, CLR_MUTED_FG
End Sub

' The chat window: the transcript (read-only, scrolls), the message box (Enter
' sends, Ctrl+Enter starts a new line) and Send (it reads Cancel while waiting),
' New chat, Copy reply, Settings and Close.
Private Sub Build_dlgAIChat()
    Const DLG_W As Long = 9000          ' 600 px
    Dim c As Control, pad As Long, w As Long, x As Long, y As Long, gap As Long, btnW As Long
    pad = Sp(SP_DIALOG): gap = Sp(SP_ROW_GAP)
    x = pad: w = DLG_W - 2 * pad: btnW = px(100)
    NewDialog "Ask AI", DLG_W, px(600)
    Ev F, "OnLoad", "OnKeyDown", "OnUnload": F.KeyPreview = True
    y = pad
    Lbl "lblTitle", "Ask AI", x, y, w, px(26), 13, True
    y = y + px(26)
    Lbl "lblService", "", x, y, w, px(18), 9, False, CLR_MUTED_FG
    y = y + px(18) + gap
    Set c = RO("txtLog", "", x, y, w, px(280), 10)
    c.BackStyle = 1: c.BackColor = CLR_SURFACE: c.BorderStyle = 1: c.BorderColor = CLR_BORDER
    c.LeftMargin = px(8): c.TopMargin = px(8): c.RightMargin = px(8)
    c.ScrollBars = 2
    c.TabStop = True                                ' keyboard users can scroll and select
    y = y + px(280) + gap
    FieldLabel "lblMessage", "Your message (Ctrl+Enter for a new line)", x, y, w
    y = y + px(18)
    Set c = Txt("txtMessage", "", x, y, w, px(72), 10)
    c.TopMargin = px(6): c.ScrollBars = 2
    y = y + px(72) + gap
    Lbl "lblStatus", "", x, y, w, px(36), 9, False, CLR_MUTED_FG
    y = y + px(36) + gap
    Set c = Btn("btnSend", "Send", x, y, btnW, px(32)): c.Default = True
    Btn "btnNew", "New chat", x + (btnW + gap), y, btnW, px(32), "outline"
    Btn "btnCopy", "Copy reply", x + 2 * (btnW + gap), y, btnW, px(32), "outline"
    Btn "btnSettings", "Settings", x + 3 * (btnW + gap), y, btnW, px(32), "outline"
    Set c = Btn("btnClose", "Close", x + w - btnW, y, btnW, px(32), "outline"): c.Cancel = True
    F.Section(acDetail).Height = y + px(32) + pad
    SaveForm AI_CHAT_FORM
End Sub

' The settings window: the service, model and address, where the API key comes
' from (an environment variable, a key file, or a key for this session), the
' standing instructions, the reply limit and timeout, with Test, Save and Close.
Private Sub Build_dlgAISettings()
    Const DLG_W As Long = 8400          ' 560 px
    Dim c As Control, pad As Long, w As Long, x As Long, x2 As Long, y As Long, half As Long, gap As Long, btnW As Long
    pad = Sp(SP_DIALOG): gap = Sp(SP_ROW_GAP)
    x = pad: w = DLG_W - 2 * pad: half = (w - Sp(16)) \ 2: x2 = x + half + Sp(16): btnW = px(96)
    NewDialog "AI settings", DLG_W, px(760)
    Ev F, "OnLoad", "OnKeyDown": F.KeyPreview = True
    y = pad
    Lbl "lblTitle", "AI settings", x, y, w, px(26), 13, True
    y = y + px(26) + Sp(16)
    FieldLabel "lblProvider", "Kind of service", x, y, half
    FieldLabel "lblModel", "Model", x2, y, half
    y = y + px(18)
    Set c = Cbo("cboProvider", "", "openai;OpenAI or compatible service;anthropic;Anthropic;gemini;Google Gemini", x, y, half, px(30), 2, "0;3000")
    Ev c, "AfterUpdate"
    Txt "txtModel", "", x2, y, half, px(30)
    y = y + px(30) + gap
    FieldLabel "lblEndpoint", "Service address (https://...)", x, y, w
    y = y + px(18)
    Txt "txtEndpoint", "", x, y, w, px(30)
    y = y + px(30) + px(4)
    Lbl "lblEndpointHint", "{model} in the address is replaced by the model name.", x, y, w, px(16), 8, False, CLR_MUTED_FG
    y = y + px(16) + Sp(16)
    FieldLabel "lblKeyEnv", "API key: environment variable", x, y, half
    FieldLabel "lblKeyHeader", "Key header (blank: the service's own)", x2, y, half
    y = y + px(18)
    Txt "txtKeyEnv", "", x, y, half, px(30)
    Txt "txtKeyHeader", "", x2, y, half, px(30)
    y = y + px(30) + gap
    FieldLabel "lblKeyFile", "API key: file whose first line is the key", x, y, w
    y = y + px(18)
    Txt "txtKeyFile", "", x, y, w - btnW - Sp(SP_ITEM_GAP), px(30)
    Btn "btnBrowse", "Browse...", x + w - btnW, y, btnW, px(30), "outline"
    y = y + px(30) + gap
    FieldLabel "lblKey", "API key for this session only", x, y, w
    y = y + px(18)
    Set c = Txt("txtKey", "", x, y, w, px(30)): c.InputMask = SECRET_MASK
    y = y + px(30) + px(4)
    Lbl "lblKeyStatus", "", x, y, w, px(32), 8, False, CLR_MUTED_FG
    y = y + px(32) + gap
    FieldLabel "lblSystem", "Standing instructions for the AI (optional)", x, y, w
    y = y + px(18)
    Set c = Txt("txtSystem", "", x, y, w, px(64), 10)
    c.TopMargin = px(6): c.ScrollBars = 2
    y = y + px(64) + gap
    FieldLabel "lblMax", "Longest reply (tokens; 0 leaves it out)", x, y, half
    FieldLabel "lblTimeout", "Wait for a reply (seconds)", x2, y, half
    y = y + px(18)
    Txt "txtMax", "", x, y, half, px(30)
    Txt "txtTimeout", "", x2, y, half, px(30)
    y = y + px(30) + gap
    Lbl "lblStatus", "", x, y, w, px(36), 9, False, CLR_MUTED_FG
    y = y + px(36) + gap
    Btn "btnTest", "Test", x, y, btnW, px(32), "outline"
    Set c = Btn("btnSave", "Save", x + w - 2 * btnW - Sp(SP_ITEM_GAP), y, btnW, px(32)): c.Default = True
    Set c = Btn("btnClose", "Close", x + w - btnW, y, btnW, px(32), "outline"): c.Cancel = True
    F.Section(acDetail).Height = y + px(32) + pad
    SaveForm AI_SETTINGS_FORM
End Sub

' The window for connecting parts of a form to the AI: the form, the button
' that asks, the box with the question, the box for the answer and an
' instruction, the connections already made, and Connect, Disconnect and Close.
Private Sub Build_dlgAIConnect()
    Const DLG_W As Long = 8400          ' 560 px
    Dim c As Control, pad As Long, w As Long, x As Long, x2 As Long, y As Long, half As Long, gap As Long, btnW As Long
    pad = Sp(SP_DIALOG): gap = Sp(SP_ROW_GAP)
    x = pad: w = DLG_W - 2 * pad: half = (w - Sp(16)) \ 2: x2 = x + half + Sp(16): btnW = px(104)
    NewDialog "Connect to AI", DLG_W, px(620)
    Ev F, "OnLoad", "OnKeyDown": F.KeyPreview = True
    y = pad
    Lbl "lblTitle", "Connect part of a form to the AI", x, y, w, px(26), 13, True
    y = y + px(26) + gap
    Lbl "lblIntro", "Choose a button that asks the AI, the box that holds the question, and where the answer goes. " & _
        "Clicking the button then sends the box's text to the AI.", x, y, w, px(36), 9, False, CLR_MUTED_FG
    y = y + px(36) + gap
    FieldLabel "lblForm", "Form (it must be open)", x, y, w
    y = y + px(18)
    Set c = Cbo("cboForm", "", "", x, y, w, px(30))
    Ev c, "AfterUpdate"
    y = y + px(30) + gap
    FieldLabel "lblTrigger", "Button that asks the AI", x, y, half
    FieldLabel "lblInput", "Box with the question", x2, y, half
    y = y + px(18)
    Set c = Cbo("cboTrigger", "", "", x, y, half, px(30))
    Ev c, "AfterUpdate"
    Cbo "cboInput", "", "", x2, y, half, px(30)
    y = y + px(30) + gap
    FieldLabel "lblOutput", "Box for the answer (blank: a message)", x, y, w
    y = y + px(18)
    Cbo "cboOutput", "", "", x, y, half, px(30)
    y = y + px(30) + gap
    FieldLabel "lblInstruction", "Instruction for the AI (optional)", x, y, w
    y = y + px(18)
    Set c = Txt("txtInstruction", "", x, y, w, px(56), 10)
    c.TopMargin = px(6): c.ScrollBars = 2
    y = y + px(56) + gap
    FieldLabel "lblSaved", "Connected on this form", x, y, w
    y = y + px(18)
    Set c = RO("txtSaved", "", x, y, w, px(64), 9, False, CLR_FG)
    c.BackStyle = 1: c.BackColor = CLR_SURFACE: c.BorderStyle = 1: c.BorderColor = CLR_BORDER
    c.LeftMargin = px(8): c.TopMargin = px(6): c.ScrollBars = 2
    c.TabStop = True
    y = y + px(64) + gap
    Lbl "lblError", "", x, y, w, px(36), 9, False, CLR_ERROR
    y = y + px(36) + gap
    Set c = Btn("btnConnect", "Connect", x, y, btnW, px(32)): c.Default = True
    Btn "btnDisconnect", "Disconnect", x + btnW + Sp(SP_ITEM_GAP), y, btnW, px(32), "outline"
    Set c = Btn("btnClose", "Close", x + w - btnW, y, btnW, px(32), "outline"): c.Cancel = True
    F.Section(acDetail).Height = y + px(32) + pad
    SaveForm AI_CONNECT_FORM
End Sub
