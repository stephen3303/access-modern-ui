Option Compare Database
Option Explicit

' Access UI kit - the AI chat and AI settings dialogs (the logic is modKitAI).
'
'   KitAIChat [title, system]   a window to type a message to the AI and read its
'                               answer, with the conversation kept while it is
'                               open (dlgAI). system tells the AI how to behave.
'   KitAISettings               the user's choices (dlgAISettings): AI on or off,
'                               the service (provider, model, address), the name
'                               of the environment variable that holds the key
'                               or a key for this session only, and which of the
'                               app's registered parts (AIRegisterPart) use AI.
'                               True when the user saved.
'   BuildAIDialogs              build step: builds both dialogs. Copy
'                               kit\forms\dlgAI.vba and dlgAISettings.vba to
'                               src\forms first (SaveForm attaches their code).
'
' The app provides APP_NAME, HandleError and LogError (modErrors).

Private m_title As String
Private m_system As String

' Open the chat window. Without the user having turned AI on, offers the
' settings first. Returns when the window closes.
Public Sub KitAIChat(Optional ByVal title As String = "", Optional ByVal system As String = "")
    If Not AIEnabled() Then
        If KitMsgBox("AI is turned off. Open the AI settings to turn it on?", vbYesNo + vbQuestion, "AI") <> vbYes Then Exit Sub
        If Not KitAISettings() Then Exit Sub
        If Not AIEnabled() Then Exit Sub
    End If
    m_title = title
    m_system = system
    AIReset
    RunDialog "dlgAI"
    AIReset                        ' the conversation isn't kept after the window closes
End Sub

' The title and system text KitAIChat was given (read by dlgAI).
Public Function AIChatTitle() As String
    AIChatTitle = m_title
End Function

Public Function AIChatSystem() As String
    AIChatSystem = m_system
End Function

' Open the settings dialog. True when the user saved.
Public Function KitAISettings() As Boolean
    KitAISettings = RunDialog("dlgAISettings")
End Function

' Build step: dlgAI and dlgAISettings (code in kit\forms).
Public Sub BuildAIDialogs()
    Build_dlgAI
    Build_dlgAISettings
End Sub

Private Sub Build_dlgAI()
    Const DLG_W As Long = 8100          ' 540 px
    Dim c As Control, pad As Long, w As Long, y As Long, gap As Long, btnW As Long
    pad = Sp(SP_DIALOG): gap = Sp(SP_ROW_GAP): w = DLG_W - 2 * pad: btnW = px(88)
    NewDialog "AI", DLG_W, px(500)
    Ev F, "OnLoad"
    y = pad
    ' The conversation: read-only, scrolls, selectable so it can be copied.
    Set c = Txt("txtAnswer", "", pad, y, w, px(250), 10)
    c.Locked = True: c.EnterKeyBehavior = True: c.ScrollBars = 2
    y = y + px(250) + gap
    Lbl "lblStatus", "", pad, y, w, px(20), 9, False, CLR_MUTED_FG
    y = y + px(20) + Sp(SP_ITEM_GAP)
    Set c = Txt("txtMessage", "", pad, y, w, px(80), 10)
    c.EnterKeyBehavior = True: c.ScrollBars = 2
    y = y + px(80) + gap
    Btn "btnSend", "Send", pad, y, btnW, px(30)
    Btn "btnCopy", "Copy answer", pad + btnW + gap, y, px(110), px(30), "outline"
    Btn "btnNew", "New chat", pad + btnW + px(110) + 2 * gap, y, btnW, px(30), "outline"
    Set c = Btn("btnClose", "Close", DLG_W - pad - btnW, y, btnW, px(30), "outline")
    c.Cancel = True
    y = y + px(30) + pad
    F.Section(acDetail).Height = y
    SaveForm "dlgAI"
End Sub

' Parts the settings list can show (AIRegisterPart allows the same number).
Public Function AIUiPartRows() As Long
    AIUiPartRows = 8
End Function

Private Sub Build_dlgAISettings()
    Const DLG_W As Long = 7200          ' 480 px
    Dim c As Control, pad As Long, w As Long, y As Long, gap As Long, i As Long, half As Long
    pad = Sp(SP_DIALOG): gap = Sp(SP_ROW_GAP): w = DLG_W - 2 * pad: half = (w - Sp(SP_ITEM_GAP)) \ 2
    NewDialog "AI settings", DLG_W, px(640)
    Ev F, "OnLoad"
    y = pad
    Chk "chkEnabled", "", pad, y + px(2)
    Lbl "lblEnabled", "Use AI in this app", pad + px(24), y, w - px(24), px(22), 10, True
    y = y + px(22) + Sp(SP_ITEM_GAP)
    Lbl "lblOff", "Text you send to AI leaves this computer. Turn this off to hide every AI feature.", _
        pad, y, w, px(34), 9, False, CLR_MUTED_FG
    y = y + px(34) + gap
    Lbl "lblProv", "Service", pad, y, half, px(18), 9, True
    Lbl "lblModel", "Model", pad + half + Sp(SP_ITEM_GAP), y, half, px(18), 9, True
    y = y + px(18)
    Cbo "cboProvider", "", "openai;anthropic;ollama", pad, y, half, px(30)
    Txt "txtModel", "", pad + half + Sp(SP_ITEM_GAP), y, half, px(30)
    y = y + px(30) + gap
    Lbl "lblEndpoint", "Address (leave blank for the service's own)", pad, y, w, px(18), 9, True
    y = y + px(18)
    Txt "txtEndpoint", "", pad, y, w, px(30)
    y = y + px(30) + px(4)
    Lbl "lblEndpointHint", "Must start with https://. Other services that use the OpenAI format work with the openai service and their own address.", _
        pad, y, w, px(34), 9, False, CLR_MUTED_FG
    y = y + px(34) + gap
    Lbl "lblKeyVar", "Environment variable that holds the key", pad, y, w, px(18), 9, True
    y = y + px(18)
    Txt "txtKeyVar", "", pad, y, w, px(30)
    y = y + px(30) + gap
    Lbl "lblKey", "Key for this session only (never saved)", pad, y, w, px(18), 9, True
    y = y + px(18)
    SecretBox "txtKey", "", pad, y, w, px(30)
    y = y + px(30) + px(4)
    Lbl "lblKeyState", "", pad, y, w, px(18), 9, False, CLR_MUTED_FG
    y = y + px(18) + gap
    Lbl "lblParts", "Where AI is used", pad, y, w, px(18), 9, True
    y = y + px(18)
    For i = 1 To AIUiPartRows()
        Chk "chkPart" & i, "", pad, y + px(2)
        Lbl "lblPart" & i, "", pad + px(24), y, w - px(24), px(22), 10
        y = y + px(24)
    Next
    y = y + gap
    Lbl "lblMsg", "", pad, y, w, px(34), 9, False, CLR_ERROR
    y = y + px(34) + gap
    Btn "btnSave", "Save", DLG_W - pad - 2 * px(88) - Sp(SP_ITEM_GAP), y, px(88), px(30)
    Set c = Btn("btnCancel", "Cancel", DLG_W - pad - px(88), y, px(88), px(30), "outline")
    c.Cancel = True
    y = y + px(30) + pad
    F.Section(acDetail).Height = y
    SaveForm "dlgAISettings"
End Sub
