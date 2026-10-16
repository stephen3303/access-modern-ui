Option Compare Database
Option Explicit

' Access UI kit - connect to an AI service over its web API.
'
' Send a message, get the reply as text. Two request formats cover most
' services: "anthropic" (the Claude Messages API) and "openai" (the Chat
' Completions format, which OpenAI, Azure OpenAI, Ollama, LM Studio, vLLM,
' OpenRouter and many others accept); "custom" is the openai format at an
' address and model you give. Change service with AIConfigure, no code change.
'
' Setup (once, by the person who runs the app):
'   AIConfigure "anthropic"                      ' or "openai", or:
'   AIConfigure "custom", "llama3", "http://localhost:11434/v1/chat/completions"
' Settings are saved per user in Program Settings\<APP_NAME>\AI.
'
' THE API KEY IS NEVER IN THE CODE OR SAVED. It is read from the Windows
' environment variable named by AIConfigure's keyVar (ANTHROPIC_API_KEY and
' OPENAI_API_KEY by default), or typed once per session with AIAskForKey
' (masked, kept in memory only until Access closes). A local "custom" service
' that needs no key works with none.
'
' Ask:     reply = AIAsk("Summarise this", "You are concise.")   one question
'          reply = AIChat("And in French?")                      keeps the conversation; AIChatReset
'          reply = AIAskTable("Which items are late?", "SELECT ...")  rows from a query as context
' Forms:   ShowAI                                  a chat window (dlgAI, built by BuildAIDialog)
'          AIConnect Me, "txtQuestion", "txtAnswer", "btnAsk"   the parts of YOUR form that talk to the AI
'          AIChooseParts Me                        ask the user which controls those are
' Class.:  data sent to an outside service leaves the database, so the AI is refused when the app's
'          classification (modClassification) is above the level set with AIConfigure's maxLevel
'          (UNCLASSIFIED until set).
'
' Errors: every request error is raised as error AI_ERR (vbObjectError + 4100) with a readable
' message (no key, no network, the service's own message, a timeout). Prompts, replies and keys
' are never written to the error log. A request waits for the service while the app stays
' responsive (DoEvents); only one request runs at a time.
'
' The app provides APP_NAME, LogError, KitMsgBox and KitInputBox (the kit's). Needs Windows
' (MSXML2.ServerXMLHTTP.6.0, ADODB.Stream); only https, or http to this computer.

Public Const AI_ERR As Long = vbObjectError + 4100

Private Const AI_SECTION As String = "AI"
Private Const AI_BINDINGS As String = "AIParts"
Private Const ANTHROPIC_URL As String = "https://api.anthropic.com/v1/messages"
Private Const OPENAI_URL As String = "https://api.openai.com/v1/chat/completions"
Private Const ANTHROPIC_VERSION As String = "2023-06-01"
Private Const DEFAULT_TIMEOUT As Long = 60          ' seconds
Private Const MAX_TURNS As Long = 40                ' messages kept in a conversation

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private m_key As String                 ' typed key, this session only
Private m_busy As Boolean               ' a request is running
Private m_n As Long                     ' conversation messages
Private m_role(0 To MAX_TURNS) As String
Private m_text(0 To MAX_TURNS) As String

' ---------------------------------------------------------------------------
' Settings
' ---------------------------------------------------------------------------

' The services with built-in request formats.
Public Function AIProviders() As Variant
    AIProviders = Array("anthropic", "openai", "custom")
End Function

' Choose the service and, optionally, its model, web address (needed for
' "custom"), the name of the environment variable holding the key, the longest
' reply in tokens, the seconds to wait, and maxLevel (the highest app
' classification allowed to use the AI, a modClassification level). Leave a
' value out to keep it (the service's default when never set). Raises
' error 5 for an unknown service, an address that is not https (or http to
' this computer), or a bad number.
Public Sub AIConfigure(ByVal provider As String, Optional ByVal model As String = "", _
        Optional ByVal endpoint As String = "", Optional ByVal keyVar As String = "", _
        Optional ByVal maxTokens As Long = 0, Optional ByVal timeoutSec As Long = 0, _
        Optional ByVal maxLevel As String = "")
    provider = LCase$(Trim$(provider))
    If InStr("|anthropic|openai|custom|", "|" & provider & "|") = 0 Then Err.Raise 5, , "Unknown AI service: " & provider
    If Len(endpoint) > 0 Then
        If Not SafeUrl(endpoint) Then Err.Raise 5, , "The AI address must start with https:// (or http:// to this computer)"
    End If
    If provider = "custom" And Len(endpoint) = 0 Then
        If provider <> AIProvider() Or Len(GetSetting(APP_NAME, AI_SECTION, "Endpoint", "")) = 0 Then Err.Raise 5, , "A custom AI service needs an address"
    End If
    If maxTokens < 0 Or timeoutSec < 0 Then Err.Raise 5, , "Tokens and seconds can't be negative"
    If Len(maxLevel) > 0 Then
        maxLevel = UCase$(Trim$(maxLevel))
        If ClassRank(maxLevel) < 0 Then Err.Raise 5, , "Unknown classification level: " & maxLevel
    End If
    If provider <> AIProvider() Then          ' a new service starts from its own defaults
        SaveSetting APP_NAME, AI_SECTION, "Model", ""
        SaveSetting APP_NAME, AI_SECTION, "Endpoint", ""
        SaveSetting APP_NAME, AI_SECTION, "KeyVar", ""
    End If
    SaveSetting APP_NAME, AI_SECTION, "Provider", provider
    If Len(model) > 0 Then SaveSetting APP_NAME, AI_SECTION, "Model", Trim$(model)
    If Len(endpoint) > 0 Then SaveSetting APP_NAME, AI_SECTION, "Endpoint", Trim$(endpoint)
    If Len(keyVar) > 0 Then SaveSetting APP_NAME, AI_SECTION, "KeyVar", Trim$(keyVar)
    If maxTokens > 0 Then SaveSetting APP_NAME, AI_SECTION, "MaxTokens", CStr(maxTokens)
    If timeoutSec > 0 Then SaveSetting APP_NAME, AI_SECTION, "Timeout", CStr(timeoutSec)
    If Len(maxLevel) > 0 Then SaveSetting APP_NAME, AI_SECTION, "MaxLevel", maxLevel
    AIChatReset
End Sub

' The service in use: "anthropic" until AIConfigure says otherwise.
Public Function AIProvider() As String
    Dim s As String
    s = LCase$(GetSetting(APP_NAME, AI_SECTION, "Provider", "anthropic"))
    If InStr("|anthropic|openai|custom|", "|" & s & "|") = 0 Then s = "anthropic"
    AIProvider = s
End Function

' The model name sent with each request.
Public Function AIModel() As String
    Dim s As String
    s = GetSetting(APP_NAME, AI_SECTION, "Model", "")
    If Len(s) = 0 Then
        Select Case AIProvider()
            Case "anthropic": s = "claude-sonnet-5-5"
            Case "openai": s = "gpt-4o-mini"
        End Select
    End If
    AIModel = s
End Function

' The web address requests go to.
Public Function AIEndpoint() As String
    Dim s As String
    s = GetSetting(APP_NAME, AI_SECTION, "Endpoint", "")
    If Len(s) = 0 Then
        Select Case AIProvider()
            Case "anthropic": s = ANTHROPIC_URL
            Case "openai": s = OPENAI_URL
        End Select
    End If
    AIEndpoint = s
End Function

' Name of the environment variable holding the key ("" for none).
Public Function AIKeyVar() As String
    Dim s As String
    s = GetSetting(APP_NAME, AI_SECTION, "KeyVar", "")
    If Len(s) = 0 Then
        Select Case AIProvider()
            Case "anthropic": s = "ANTHROPIC_API_KEY"
            Case "openai": s = "OPENAI_API_KEY"
        End Select
    End If
    AIKeyVar = s
End Function

' The highest app classification (modClassification level) allowed to use the AI.
Public Function AIMaxLevel() As String
    Dim s As String
    s = UCase$(GetSetting(APP_NAME, AI_SECTION, "MaxLevel", ""))
    If ClassRank(s) < 0 Then s = "UNCLASSIFIED"
    AIMaxLevel = s
End Function

' Hold a key for this session only (never saved). "" forgets it.
Public Sub AISetKey(ByVal key As String)
    m_key = Trim$(key)
End Sub

' Ask the user for the key (masked) and hold it for the session. True when
' one was entered. Use it when no environment variable is set.
Public Function AIAskForKey() As Boolean
    Dim v As Variant
    v = KitInputBox("Enter the API key for " & AIProvider() & ". It is kept only until you close " & APP_NAME & ".", _
        "AI service key", , True)
    If IsNull(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    AISetKey CStr(v)
    AIAskForKey = True
End Function

' True when a request could be sent: allowed by classification and a key available (or none needed).
Public Function AIReady() As Boolean
    AIReady = (Len(ClassProblem()) = 0) And (Len(KeyProblem()) = 0) And (Len(AIEndpoint()) > 0)
End Function

' ---------------------------------------------------------------------------
' Asking
' ---------------------------------------------------------------------------

' Send one question and return the reply. system (optional) sets the AI's
' role for this question. Nothing is remembered; see AIChat. Raises AI_ERR.
Public Function AIAsk(ByVal prompt As String, Optional ByVal system As String = "") As String
    Dim roles(0 To 0) As String, texts(0 To 0) As String
    roles(0) = "user": texts(0) = prompt
    AIAsk = Send(roles, texts, 1, system)
End Function

' Send a message as part of a conversation: the earlier messages and replies go
' with it, so "and in French?" makes sense. The newest MAX_TURNS messages are
' kept. A failed request leaves the conversation as it was. Raises AI_ERR.
Public Function AIChat(ByVal userText As String, Optional ByVal system As String = "") As String
    Dim reply As String, i As Long
    If m_n >= MAX_TURNS Then          ' drop the oldest exchange
        For i = 2 To m_n - 1
            m_role(i - 2) = m_role(i): m_text(i - 2) = m_text(i)
        Next
        m_n = m_n - 2
    End If
    m_role(m_n) = "user": m_text(m_n) = userText
    m_n = m_n + 1
    On Error GoTo Fail
    reply = Send(m_role, m_text, m_n, system)
    m_role(m_n) = "assistant": m_text(m_n) = reply
    m_n = m_n + 1
    AIChat = reply
    Exit Function
Fail:
    m_n = m_n - 1                     ' forget the unanswered message
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Start a new conversation.
Public Sub AIChatReset()
    m_n = 0
End Sub

' Ask about rows of the database: instruction followed by the rows sql returns
' (a header line, then tab-separated values, at most maxRows rows). The rows
' are sent to the AI service, so only choose data that may leave. Raises AI_ERR.
Public Function AIAskTable(ByVal instruction As String, ByVal sql As String, Optional ByVal maxRows As Long = 50, _
        Optional ByVal system As String = "") As String
    Dim rs As DAO.Recordset, s As String, i As Long, n As Long, v As Variant
    If maxRows < 1 Then maxRows = 1
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    For i = 0 To rs.Fields.Count - 1
        s = s & IIf(i > 0, vbTab, "") & rs.Fields(i).Name
    Next
    Do While Not rs.EOF And n < maxRows
        s = s & vbLf
        For i = 0 To rs.Fields.Count - 1
            v = rs.Fields(i).Value
            If IsNull(v) Then v = "" Else v = Replace(Replace(Replace(CStr(v), vbTab, " "), vbCr, " "), vbLf, " ")
            s = s & IIf(i > 0, vbTab, "") & v
        Next
        n = n + 1
        rs.MoveNext
    Loop
    If Not rs.EOF Then s = s & vbLf & "(more rows not shown)"
    rs.Close
    AIAskTable = AIAsk(instruction & vbLf & vbLf & s, system)
End Function

' ---------------------------------------------------------------------------
' Forms: the parts of an interface that talk to the AI
' ---------------------------------------------------------------------------

' Say which controls of form frm are the AI's: promptCtl is the text box the
' user types the question in, replyCtl the text box or label that gets the
' reply, and sendCtl (optional) the button that sends it: its Click is set to
' ask the AI. keepChat:=True makes the form a conversation (AIChat); otherwise
' each send is a single question. system gives the AI its role. The choice is
' saved (per user, per form), so a form needs it only once; AIRestoreParts
' brings it back when the form opens. Raises error 5 for a control the form
' doesn't have.
Public Sub AIConnect(frm As Object, ByVal promptCtl As String, ByVal replyCtl As String, _
        Optional ByVal sendCtl As String = "", Optional ByVal keepChat As Boolean = False, _
        Optional ByVal system As String = "")
    Dim nm As Variant
    For Each nm In Array(promptCtl, replyCtl, sendCtl)
        If Len(nm) > 0 Then
            If Not HasCtl(frm, CStr(nm)) Then Err.Raise 5, , "Form " & frm.Name & " has no control named " & nm
        End If
    Next
    SaveSetting APP_NAME, AI_BINDINGS, frm.Name, promptCtl & "|" & replyCtl & "|" & sendCtl & "|" & _
        IIf(keepChat, "1", "0") & "|" & Replace(system, "|", "/")
    WireParts frm, promptCtl, replyCtl, sendCtl
End Sub

' Reconnect a form to the controls saved for it (call it from Form_Load).
' True when the form had saved parts.
Public Function AIRestoreParts(frm As Object) As Boolean
    Dim p() As String
    On Error GoTo Fail
    p = Split(GetSetting(APP_NAME, AI_BINDINGS, frm.Name, ""), "|")
    If UBound(p) < 3 Then Exit Function
    If Not (HasCtl(frm, p(0)) And HasCtl(frm, p(1))) Then Exit Function
    If Len(p(2)) > 0 Then If Not HasCtl(frm, p(2)) Then p(2) = ""
    WireParts frm, p(0), p(1), p(2)
    AIRestoreParts = True
    Exit Function
Fail:
    LogError "modKitAI.AIRestoreParts", "form " & frm.Name
End Function

' Forget the parts saved for a form.
Public Sub AIForgetParts(ByVal formName As String)
    On Error Resume Next
    DeleteSetting APP_NAME, AI_BINDINGS, formName
End Sub

' Let the user say which controls of frm take the question, show the reply and
' send it (by name, as in the form's Design view). True when they were set.
' Meant for the person who builds or administers the app.
Public Function AIChooseParts(frm As Object) As Boolean
    Dim a As Variant, b As Variant, c As Variant, names As String, ctl As Control
    For Each ctl In frm.Controls
        If ctl.ControlType = acTextBox Or ctl.ControlType = acLabel Or ctl.ControlType = acCommandButton Then
            names = names & ctl.Name & ", "
        End If
    Next
    If Len(names) > 1500 Then names = Left$(names, 1500) & "..."
    a = KitInputBox("Name of the box where the question is typed." & vbCrLf & vbCrLf & "Controls: " & names, "AI question box")
    If IsNull(a) Then Exit Function
    b = KitInputBox("Name of the box or label that shows the reply.", "AI reply")
    If IsNull(b) Then Exit Function
    c = KitInputBox("Name of the button that sends the question (blank for none).", "AI send button")
    If IsNull(c) Then Exit Function
    On Error GoTo Bad
    AIConnect frm, Trim$(CStr(a)), Trim$(CStr(b)), Trim$(CStr(c))
    AIChooseParts = True
    Exit Function
Bad:
    KitMsgBox Err.Description, vbExclamation, "AI parts"
End Function

' Send the question in a form's connected question box and put the reply in its
' reply control (a button's Click calls this: =AIAskFromForm([Form])). Errors
' are shown in the reply control, not raised, so a click never stops the app.
Public Function AIAskFromForm(frm As Object) As Boolean
    Dim p() As String, q As String, reply As String
    On Error GoTo Fail
    p = Split(GetSetting(APP_NAME, AI_BINDINGS, frm.Name, ""), "|")
    If UBound(p) < 3 Then Err.Raise AI_ERR, "modKitAI", "This form has no AI parts set"
    q = Trim$(Nz(frm.Controls(p(0)).Value, ""))
    If Len(q) = 0 Then Exit Function
    SetReply frm.Controls(p(1)), "Thinking" & ChrW(8230)
    DoEvents
    If p(3) = "1" Then reply = AIChat(q, p(4)) Else reply = AIAsk(q, p(4))
    SetReply frm.Controls(p(1)), reply
    AIAskFromForm = True
    Exit Function
Fail:
    On Error Resume Next
    If UBound(p) >= 1 Then SetReply frm.Controls(p(1)), "AI error: " & Err.Description
End Function

' Open the chat window (dlgAI). system sets the AI's role for the session's chat.
Public Sub ShowAI(Optional ByVal system As String = "")
    TempVars("AISystem") = system
    RunDialog "dlgAI"
End Sub

' The role the chat window gives the AI (set by ShowAI).
Public Function AIChatSystem() As String
    AIChatSystem = Nz(TempVars("AISystem"), "")
End Function

' Build dlgAI, the chat window ShowAI opens. Call it once with the app's
' dialogs; copy kit\forms\dlgAI.vba to src\forms first.
Public Sub BuildAIDialog()
    Dim c As Control, w As Long, h As Long, y As Long
    w = px(560): h = px(460)
    NewDialog "Ask AI", w, h
    Ev F, "OnLoad"
    Lbl "lblService", "", px(20), px(14), w - px(40), px(18), 8, True, CLR_MUTED_FG
    Set c = RO("txtReply", "", px(20), px(36), w - px(40), px(274), 10, False, CLR_FG)
    c.TopMargin = 0: c.LeftMargin = px(8): c.RightMargin = px(8)
    c.BackStyle = 1: c.BackColor = CLR_INSET
    c.ScrollBars = 2
    c.TabStop = True
    Set c = Txt("txtPrompt", "", px(20), px(322), w - px(40), px(70), 10)
    c.EnterKeyBehavior = True
    c.ScrollBars = 2
    c.ControlTipText = "Type a message, then click Ask (Ctrl+Enter)"
    y = px(404)
    Set c = Btn("btnAsk", "Ask", w - px(20) - px(88), y, px(88), px(30))
    c.Default = False
    Btn "btnCopy", "Copy reply", w - px(20) - px(88) - px(8) - px(104), y, px(104), px(30), "outline"
    Btn "btnNew", "New chat", w - px(20) - px(88) - px(8) - px(104) - px(8) - px(96), y, px(96), px(30), "outline"
    Set c = Btn("btnClose", "Close", px(20), y, px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgAI"
End Sub

' ---------------------------------------------------------------------------
' Requests
' ---------------------------------------------------------------------------

' Build, send and read one request for the first n messages.
Private Function Send(roles() As String, texts() As String, ByVal n As Long, ByVal system As String) As String
    Dim body As String, url As String, status As Long, resp As String, prov As String
    Dim http As Object, t0 As Single, limit As Long, key As String, num As Long, d As String
    If m_busy Then Err.Raise AI_ERR, "modKitAI", "The AI is still working on the last request"
    If Len(Trim$(texts(0))) = 0 And n = 1 Then Err.Raise AI_ERR, "modKitAI", "There is no message to send"
    If Len(ClassProblem()) > 0 Then Err.Raise AI_ERR, "modKitAI", ClassProblem()
    If Len(KeyProblem()) > 0 Then Err.Raise AI_ERR, "modKitAI", KeyProblem()
    prov = AIProvider()
    url = AIEndpoint()
    If Len(url) = 0 Then Err.Raise AI_ERR, "modKitAI", "No AI address is set (AIConfigure)"
    If Not SafeUrl(url) Then Err.Raise AI_ERR, "modKitAI", "The AI address must start with https:// (or http:// to this computer)"
    If Len(AIModel()) = 0 Then Err.Raise AI_ERR, "modKitAI", "No AI model is set (AIConfigure)"
    body = AIRequestBody(prov, roles, texts, n, system)
    key = CurrentKey()
    limit = CLng(Val(GetSetting(APP_NAME, AI_SECTION, "Timeout", CStr(DEFAULT_TIMEOUT))))
    If limit < 5 Then limit = DEFAULT_TIMEOUT
    m_busy = True
    On Error GoTo Fail
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 10000, 10000, limit * 1000, limit * 1000
    http.Open "POST", url, True
    http.setRequestHeader "Content-Type", "application/json"
    If prov = "anthropic" Then
        http.setRequestHeader "x-api-key", key
        http.setRequestHeader "anthropic-version", ANTHROPIC_VERSION
    ElseIf Len(key) > 0 Then
        http.setRequestHeader "Authorization", "Bearer " & key
    End If
    http.send body
    t0 = Timer
    Do While http.readyState < 4
        DoEvents
        Sleep 30
        If Timer - t0 > limit Then http.abort: Err.Raise AI_ERR, "modKitAI", "The AI service didn't answer within " & limit & " seconds"
    Loop
    status = http.status
    resp = Utf8(http.responseBody)
    m_busy = False
    If status < 200 Or status >= 300 Then
        Err.Raise AI_ERR, "modKitAI", "The AI service refused the request (" & status & "): " & ServiceMessage(resp)
    End If
    Send = AIReplyText(prov, resp)
    If Len(Send) = 0 Then Err.Raise AI_ERR, "modKitAI", "The AI service sent back no text"
    Exit Function
Fail:
    m_busy = False
    num = Err.Number: d = Err.Description
    If num <> AI_ERR Then
        LogError "modKitAI.Send", prov & " " & Err.Number       ' never the prompt, reply or key
        d = "Couldn't reach the AI service (" & d & ")"
    End If
    Err.Raise AI_ERR, "modKitAI", d
End Function

' The JSON request body for service prov and the first n messages (also lets the checks see it).
Public Function AIRequestBody(ByVal prov As String, roles() As String, texts() As String, ByVal n As Long, _
        ByVal system As String) As String
    Dim s As String, i As Long, mt As Long   ' (for openai the system text goes first in messages)
    mt = CLng(Val(GetSetting(APP_NAME, AI_SECTION, "MaxTokens", "1024")))
    If mt < 1 Then mt = 1024
    s = "{""model"":" & JStr(AIModel())
    s = s & ",""max_tokens"":" & mt
    If prov = "anthropic" And Len(system) > 0 Then s = s & ",""system"":" & JStr(system)
    s = s & ",""messages"":["
    If prov <> "anthropic" And Len(system) > 0 Then s = s & "{""role"":""system"",""content"":" & JStr(system) & "},"
    For i = 0 To n - 1
        s = s & IIf(i > 0, ",", "") & "{""role"":" & JStr(roles(i)) & ",""content"":" & JStr(texts(i)) & "}"
    Next
    AIRequestBody = s & "]}"
End Function

' The reply text in a successful response from service prov ("" when it has none).
Public Function AIReplyText(ByVal prov As String, ByVal json As String) As String
    Dim p As Long, stopAt As Long, s As String, v As String
    If prov = "anthropic" Then
        p = InStr(json, """content""")
        If p = 0 Then Exit Function
        stopAt = InStr(p, json, """stop_reason""")
        If stopAt = 0 Then stopAt = Len(json) + 1
        Do
            v = JsonString(json, "text", p)
            If p = 0 Or p > stopAt Then Exit Do
            s = s & IIf(Len(s) > 0, vbCrLf, "") & v
        Loop
        AIReplyText = s
    Else
        p = InStr(json, """message""")
        If p = 0 Then Exit Function
        p = p + Len("""message""")
        AIReplyText = JsonString(json, "content", p)
    End If
End Function

' The service's own error message in an error response, else its first words.
Private Function ServiceMessage(ByVal json As String) As String
    Dim p As Long, s As String
    p = InStr(json, """error""")
    If p > 0 Then
        p = p + 7
        s = JsonString(json, "message", p)
    End If
    If Len(s) = 0 Then s = Left$(Replace(json, vbLf, " "), 200)
    ServiceMessage = s
End Function

' The string value of the next "key": "..." at or after pos, with its escapes
' resolved. pos moves to just after the value, or to 0 when there is no such
' key. A value that isn't a string (null) gives "".
Private Function JsonString(ByVal json As String, ByVal key As String, ByRef pos As Long) As String
    Dim p As Long, q As Long, ch As String, out As String, tok As String, n As Long
    tok = """" & key & """"
    p = pos
    Do
        p = InStr(p, json, tok)
        If p = 0 Then pos = 0: Exit Function
        q = p + Len(tok)
        Do While Mid$(json, q, 1) = " " Or Mid$(json, q, 1) = vbLf Or Mid$(json, q, 1) = vbCr: q = q + 1: Loop
        If Mid$(json, q, 1) = ":" Then Exit Do      ' a key, not a value that happens to match
        p = p + 1
    Loop
    q = q + 1
    Do While Mid$(json, q, 1) = " " Or Mid$(json, q, 1) = vbLf Or Mid$(json, q, 1) = vbCr: q = q + 1: Loop
    If Mid$(json, q, 1) <> """" Then pos = q: Exit Function
    q = q + 1
    n = Len(json)
    Do While q <= n
        ch = Mid$(json, q, 1)
        If ch = """" Then Exit Do
        If ch = "\" Then
            q = q + 1
            ch = Mid$(json, q, 1)
            Select Case ch
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case "t": out = out & vbTab
                Case "b": out = out & Chr$(8)
                Case "f": out = out & Chr$(12)
                Case "u"
                    out = out & ChrW(CLng("&H" & Mid$(json, q + 1, 4)))
                    q = q + 4
                Case Else: out = out & ch           ' \" \\ \/
            End Select
        Else
            out = out & ch
        End If
        q = q + 1
    Loop
    pos = q + 1
    JsonString = out
End Function

' A JSON string literal for s: quotes and backslashes escaped, control and
' non-ASCII characters as \uXXXX, so the request body is plain ASCII.
Private Function JStr(ByVal s As String) As String
    Dim i As Long, c As Long, out As String, ch As String
    out = """"
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        c = AscW(ch)
        If c < 0 Then c = c + 65536
        Select Case c
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 10: out = out & "\n"
            Case 13: out = out & "\r"
            Case 9: out = out & "\t"
            Case 32 To 126: out = out & ch
            Case Else: out = out & "\u" & Right$("0000" & Hex$(c), 4)
        End Select
    Next
    JStr = out & """"
End Function

' UTF-8 bytes to text.
Private Function Utf8(bytes As Variant) As String
    Dim st As Object
    If IsEmpty(bytes) Then Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open: st.Write bytes
    st.Position = 0: st.Type = 2: st.Charset = "utf-8"
    Utf8 = st.ReadText
    st.Close
End Function

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

' https, or http to this computer (a local model).
Private Function SafeUrl(ByVal url As String) As Boolean
    Dim u As String
    u = LCase$(Trim$(url))
    If Left$(u, 8) = "https://" Then SafeUrl = (Len(u) > 8): Exit Function
    SafeUrl = (Left$(u, 17) = "http://localhost:" Or Left$(u, 17) = "http://localhost/" Or _
        Left$(u, 17) = "http://127.0.0.1:" Or Left$(u, 17) = "http://127.0.0.1/")
End Function

' The key to send: the typed one, else the environment variable's.
Private Function CurrentKey() As String
    Dim v As String
    If Len(m_key) > 0 Then CurrentKey = m_key: Exit Function
    v = AIKeyVar()
    If Len(v) > 0 Then CurrentKey = Trim$(Environ$(v))
End Function

' Why no key is available, or "" when one is or none is needed (custom without a variable).
Private Function KeyProblem() As String
    If Len(CurrentKey()) > 0 Then Exit Function
    If AIProvider() = "custom" And Len(AIKeyVar()) = 0 Then Exit Function
    If Len(AIKeyVar()) > 0 Then
        KeyProblem = "No API key. Set the Windows environment variable " & AIKeyVar() & " and restart " & APP_NAME & ", or enter a key for this session."
    Else
        KeyProblem = "No API key. Enter a key for this session."
    End If
End Function

' Why the app's classification forbids the AI, or "".
Private Function ClassProblem() As String
    If ClassRank(AppClassification()) > ClassRank(AIMaxLevel()) Then
        ClassProblem = "AI is not permitted in a " & AppClassification() & " app (limit: " & AIMaxLevel() & ")"
    End If
End Function

Private Function HasCtl(frm As Object, ByVal nm As String) As Boolean
    Dim c As Control
    On Error Resume Next
    Set c = frm.Controls(nm)
    HasCtl = Not c Is Nothing
End Function

' Put text in a text box or label.
Private Sub SetReply(c As Object, ByVal s As String)
    If c.ControlType = acLabel Then c.Caption = Replace(s, "&", "&&") Else c.Value = s
End Sub

' Wire the send button's Click to ask the AI (the form's own code is untouched).
Private Sub WireParts(frm As Object, ByVal promptCtl As String, ByVal replyCtl As String, ByVal sendCtl As String)
    If Len(sendCtl) > 0 Then frm.Controls(sendCtl).OnClick = "=AIAskFromForm([Form])"
End Sub
