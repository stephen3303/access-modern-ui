Option Compare Database
Option Explicit

' Access UI kit - connect the app to an AI service over its web API: send a
' message, get the reply, use it in forms and data, and let the user choose
' which parts of the interface are AI-connected.
'
' Services: two request formats cover most of them.
'   "openai"     the chat-completions format: OpenAI, Azure OpenAI, and
'                the local servers that copy it (Ollama, LM Studio, vLLM).
'   "anthropic"  the Messages format.
' Another service with its own format is one more case in BuildRequest and
' ParseReply; nothing else changes.
'
' Set up once at startup (or let the user: AIConfigure's values are remembered
' in Program Settings\<APP_NAME>\AI, never the key):
'   AIConfigure "openai", "https://api.example.com/v1/chat/completions", "model-name", "MY_AI_KEY"
'   AIConfigure "anthropic", "https://api.example.com/v1/messages", "model-name", "MY_AI_KEY"
' Endpoint and model are the service's own, from its documentation. The
' endpoint must be https (http only for localhost). The fourth argument names
' the Windows environment variable that holds the API key. The key is never
' in code, files, the registry or the error log: it is read from that
' variable, or asked for once per session (a masked KitInputBox) and kept
' in memory only. AIClearKey forgets it.
'
' Messages and replies:
'   reply = AIAsk("Summarise this note: " & note)       one question; keeps the conversation
'   reply = AIAsk(text, , False)                         a one-off, no history
'   AIReset                                              start a new conversation
'   AILastReply(), AILastError(), AILastStatus()         the last call's outcome
' AIAsk raises an error (and logs it) when the call fails; AITryAsk returns
' False and AILastError() instead, for a form that shows the problem itself.
' Calls wait for the service (KitBusy shows); keep prompts short, and see
' AI_MAX_CONTEXT for how much data goes with them.
'
' Other systems and forms:
'   AIAskWithData prompt, sql        the query's rows (as CSV) go with the question
'   AIAskControl frm                 input control -> AI -> output control (below)
'   AISetHandler "modApp.OnAIReply"  every reply is also passed to this
'                                    Public Sub (prompt As String, reply As String)
'   AIOpenAssistant [prompt]         the chat dialog dlgAI (BuildAIDialog builds it)
'
' User-specified AI parts: the user (or the app's settings page) says which
' controls on a form are AI-connected. Per form there are three roles:
'   "input"   the control whose text is sent
'   "output"  the control that gets the reply
'   "send"    the button that sends
'   AISetPart "frmNote", "input", "txtBody"
'   AISetPart "frmNote", "output", "txtSummary"
'   AISetPart "frmNote", "send", "btnSummarise"
'   AIAttach Me                       in the form's Load: wires up the chosen send button
'   AIChoosePart Me, "output"         asks the user which control (by name)
'   AIPartsDescribe(Me)               the choices, for a settings page
' A form with no choices is left alone, so AIAttach is safe to add to every
' form. Choosing a part never needs a code change.
'
' Privacy: whatever is in the prompt, and the data AIAskWithData adds, goes
' to the service. Use only a service the organization has approved for the
' data, and mind the classification of what is sent (modClassification).

Public Const AI_MAX_CONTEXT As Long = 20000     ' characters of query data sent with a question
Public Const AI_TIMEOUT_MS As Long = 60000      ' wait for a reply
Public Const AI_DEFAULT_TOKENS As Long = 1024   ' reply limit (required by "anthropic")

Private Const SETTINGS_KEY As String = "AI"
Private Const ANTHROPIC_VERSION As String = "2023-06-01"
Private Const MAX_TURNS As Long = 40             ' messages kept in a conversation

Private m_provider As String
Private m_endpoint As String
Private m_model As String
Private m_keyVar As String
Private m_system As String
Private m_maxTokens As Long
Private m_key As String                          ' this session's key, memory only
Private m_hist As Collection                     ' "role" & vbTab & "text"
Private m_handler As String
Private m_reply As String
Private m_error As String
Private m_status As Long
Private m_loaded As Boolean

' JSON reader state
Private m_js As String
Private m_jp As Long

' ---------------------------------------------------------------------------
' Setup
' ---------------------------------------------------------------------------

' Choose the service. provider: "openai" or "anthropic" (the request format).
' keyVar: the environment variable holding the key. Remembered for next time.
' Raises error 5 for an unknown provider, an endpoint that isn't https (or
' localhost), or an empty model.
Public Sub AIConfigure(ByVal provider As String, ByVal endpoint As String, ByVal model As String, _
        Optional ByVal keyVar As String = "AI_API_KEY", Optional ByVal systemPrompt As String = "", _
        Optional ByVal maxTokens As Long = 0)
    provider = LCase$(Trim$(provider)): endpoint = Trim$(endpoint): model = Trim$(model)
    If provider <> "openai" And provider <> "anthropic" Then Err.Raise 5, , "Unknown AI provider: " & provider
    If Not EndpointOk(endpoint) Then Err.Raise 5, , "The AI endpoint must start with https:// (http:// only for localhost): " & endpoint
    If Len(model) = 0 Then Err.Raise 5, , "Name the AI model"
    If Len(Trim$(keyVar)) = 0 Then Err.Raise 5, , "Name the environment variable that holds the AI key"
    If maxTokens < 0 Then Err.Raise 5, , "maxTokens can't be negative"
    m_provider = provider: m_endpoint = endpoint: m_model = model
    m_keyVar = Trim$(keyVar): m_system = systemPrompt: m_maxTokens = maxTokens
    m_loaded = True
    SaveSetting APP_NAME, SETTINGS_KEY, "Provider", m_provider
    SaveSetting APP_NAME, SETTINGS_KEY, "Endpoint", m_endpoint
    SaveSetting APP_NAME, SETTINGS_KEY, "Model", m_model
    SaveSetting APP_NAME, SETTINGS_KEY, "KeyVar", m_keyVar
    SaveSetting APP_NAME, SETTINGS_KEY, "System", m_system
    SaveSetting APP_NAME, SETTINGS_KEY, "MaxTokens", CStr(m_maxTokens)
    AIReset
End Sub

' True when a service is configured (this session, or remembered).
Public Function AIConfigured() As Boolean
    LoadConfig
    AIConfigured = (Len(m_endpoint) > 0 And Len(m_model) > 0)
End Function

' The instructions sent with every question (the assistant's role).
Public Sub AISetSystem(ByVal text As String)
    LoadConfig
    m_system = text
    SaveSetting APP_NAME, SETTINGS_KEY, "System", m_system
End Sub

Private Sub LoadConfig()
    If m_loaded Then Exit Sub
    m_provider = GetSetting(APP_NAME, SETTINGS_KEY, "Provider", "")
    m_endpoint = GetSetting(APP_NAME, SETTINGS_KEY, "Endpoint", "")
    m_model = GetSetting(APP_NAME, SETTINGS_KEY, "Model", "")
    m_keyVar = GetSetting(APP_NAME, SETTINGS_KEY, "KeyVar", "AI_API_KEY")
    m_system = GetSetting(APP_NAME, SETTINGS_KEY, "System", "")
    m_maxTokens = Val(GetSetting(APP_NAME, SETTINGS_KEY, "MaxTokens", "0"))
    If m_provider <> "openai" And m_provider <> "anthropic" Then m_provider = "": m_endpoint = "": m_model = ""
    If Len(m_endpoint) > 0 Then
        If Not EndpointOk(m_endpoint) Then m_endpoint = "": m_model = ""
    End If
    m_loaded = True
End Sub

Private Function EndpointOk(ByVal url As String) As Boolean
    url = LCase$(url)
    If Left$(url, 8) = "https://" Then EndpointOk = (Len(url) > 8): Exit Function
    EndpointOk = (Left$(url, 17) = "http://localhost:" Or Left$(url, 17) = "http://localhost/" _
        Or Left$(url, 17) = "http://127.0.0.1:" Or Left$(url, 17) = "http://127.0.0.1/")
End Function

' ---------------------------------------------------------------------------
' The API key
' ---------------------------------------------------------------------------

' The key: this session's, else the environment variable's, else asked once
' (masked). "" when there is none (a local server may need none) and the user
' cancels.
Private Function ApiKey() As String
    Dim v As Variant
    If Len(m_key) = 0 Then m_key = Environ$(m_keyVar)
    If Len(m_key) = 0 Then
        v = KitInputBox("Enter the API key for the AI service. It is kept in memory until you close the app " & _
            "and is not saved. (To avoid this question, set the " & m_keyVar & " environment variable.)", _
            "AI service", "", True)
        If Not IsNull(v) Then m_key = CStr(v)
    End If
    ApiKey = m_key
End Function

' Forget the key held in memory.
Public Sub AIClearKey()
    m_key = ""
End Sub

' ---------------------------------------------------------------------------
' Asking
' ---------------------------------------------------------------------------

' Send a message and return the reply. keepHistory: the question and reply
' join the conversation (the service sees earlier turns); False for a
' one-off. systemPrompt replaces the configured instructions for this call.
' Raises the error (after logging it) when the call fails.
Public Function AIAsk(ByVal message As String, Optional ByVal systemPrompt As String = "", _
        Optional ByVal keepHistory As Boolean = True) As String
    On Error GoTo Fail
    If Not AICall(message, systemPrompt, keepHistory) Then Err.Raise vbObjectError + 5100, "modKitAI.AIAsk", m_error
    AIAsk = m_reply
    Exit Function
Fail:
    Dim num As Long, desc As String
    num = Err.Number: desc = Err.Description
    m_error = desc
    KitBusyEnd
    LogError "modKitAI.AIAsk", Left$(m_provider & " " & m_model, 60)
    Err.Raise num, "modKitAI.AIAsk", desc
End Function

' Like AIAsk, but a failure returns False with the reason in AILastError()
' (and nothing is raised), so a form can show it where it likes. The reply
' is in reply.
Public Function AITryAsk(ByVal message As String, reply As String, Optional ByVal systemPrompt As String = "", _
        Optional ByVal keepHistory As Boolean = True) As Boolean
    On Error GoTo Fail
    reply = ""
    AITryAsk = AICall(message, systemPrompt, keepHistory)
    If AITryAsk Then reply = m_reply
    Exit Function
Fail:
    m_error = Err.Description
    KitBusyEnd
    AITryAsk = False
End Function

' The last reply ("" after a failure), the last failure ("" after success)
' and the last HTTP status (0 when the service couldn't be reached).
Public Function AILastReply() As String
    AILastReply = m_reply
End Function
Public Function AILastError() As String
    AILastError = m_error
End Function
Public Function AILastStatus() As Long
    AILastStatus = m_status
End Function

' Start a new conversation.
Public Sub AIReset()
    Set m_hist = New Collection
End Sub

' Add rows from a query to the question: the SQL's rows go in as CSV, up to
' AI_MAX_CONTEXT characters (the rest is cut and the service is told), and
' the question follows. For summaries and questions about the app's data.
' Only what the SQL selects is sent; select just the columns that are needed.
Public Function AIAskWithData(ByVal question As String, ByVal sql As String, Optional ByVal keepHistory As Boolean = False) As String
    Dim rs As DAO.Recordset, f As DAO.Field, line As String, data As String, cut As Boolean
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    For Each f In rs.Fields
        line = line & IIf(Len(line) > 0, ",", "") & CsvField(f.Name)
    Next
    data = line & vbLf
    Do While Not rs.EOF
        line = ""
        For Each f In rs.Fields
            line = line & IIf(Len(line) > 0, ",", "") & CsvField(f.Value)
        Next
        If Len(data) + Len(line) + 1 > AI_MAX_CONTEXT Then cut = True: Exit Do
        data = data & line & vbLf
        rs.MoveNext
    Loop
    rs.Close
    AIAskWithData = AIAsk("Data (CSV" & IIf(cut, ", truncated" , "") & "):" & vbLf & data & vbLf & question, , keepHistory)
End Function

' Name a Public Sub (prompt As String, reply As String) in a standard module
' that is called after every successful reply: log it, store it, feed another
' system. "" to stop. An error in the handler is logged and doesn't affect
' the reply.
Public Sub AISetHandler(ByVal procName As String)
    m_handler = Trim$(procName)
End Sub

' ---------------------------------------------------------------------------
' The call itself
' ---------------------------------------------------------------------------

Private Function AICall(ByVal message As String, ByVal systemPrompt As String, ByVal keepHistory As Boolean) As Boolean
    Dim body As String, key As String, http As Object, hdr As String, text As String
    m_reply = "": m_error = "": m_status = 0
    LoadConfig
    If Len(Trim$(message)) = 0 Then m_error = "Type a message first.": Exit Function
    If Not AIConfigured() Then m_error = "The AI service isn't set up yet (AIConfigure).": Exit Function
    If Len(systemPrompt) = 0 Then systemPrompt = m_system
    If m_hist Is Nothing Then AIReset
    If Not keepHistory Then
        body = BuildRequest(systemPrompt, Nothing, message)
    Else
        body = BuildRequest(systemPrompt, m_hist, message)
    End If
    key = ApiKey()
    If Len(key) = 0 And Left$(LCase$(m_endpoint), 5) = "https" Then m_error = "No API key was entered.": Exit Function

    KitBusy "Waiting for the AI service..."
    On Error GoTo Fail
    Set http = CreateObject("MSXML2.ServerXMLHTTP.60")
    http.setTimeouts 10000, 10000, AI_TIMEOUT_MS, AI_TIMEOUT_MS
    http.Open "POST", m_endpoint, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    If m_provider = "anthropic" Then
        http.setRequestHeader "x-api-key", key
        http.setRequestHeader "anthropic-version", ANTHROPIC_VERSION
    ElseIf InStr(LCase$(m_endpoint), ".openai.azure.com") > 0 Then
        http.setRequestHeader "api-key", key
    ElseIf Len(key) > 0 Then
        http.setRequestHeader "Authorization", "Bearer " & key
    End If
    http.send body
    m_status = http.Status
    text = http.responseText
    KitBusyEnd
    On Error GoTo 0

    If m_status < 200 Or m_status > 299 Then
        m_error = "The AI service answered " & m_status & ": " & ServiceError(text)
        If m_status = 401 Or m_status = 403 Then m_key = ""        ' ask again next time
        Exit Function
    End If
    m_reply = ParseReply(text)
    If Len(m_reply) = 0 Then m_error = "The AI service sent no text.": Exit Function
    If keepHistory Then
        m_hist.Add "user" & vbTab & message
        m_hist.Add "assistant" & vbTab & m_reply
        Do While m_hist.Count > MAX_TURNS: m_hist.Remove 1: Loop
        Do While m_hist.Count > 0
            If Left$(m_hist(1), 4) = "user" Then Exit Do
            m_hist.Remove 1                       ' the conversation starts with the user
        Loop
    End If
    AICall = True
    CallHandler message, m_reply
    Exit Function
Fail:
    m_error = "Couldn't reach the AI service: " & Err.Description
    KitBusyEnd
    m_status = 0
End Function

Private Sub CallHandler(ByVal prompt As String, ByVal reply As String)
    If Len(m_handler) = 0 Then Exit Sub
    On Error GoTo Fail
    Application.Run m_handler, prompt, reply
    Exit Sub
Fail:
    LogError "modKitAI.CallHandler", Left$(m_handler, 60)
End Sub

' The JSON body for a question after the history (a Collection of
' "role" & vbTab & "text", or Nothing).
Private Function BuildRequest(ByVal systemPrompt As String, hist As Collection, ByVal message As String) As String
    Dim msgs As String, i As Long, p() As String, n As Long, s As String
    If m_provider = "openai" And Len(systemPrompt) > 0 Then
        msgs = JsonMsg("system", systemPrompt)
    End If
    If Not hist Is Nothing Then
        For i = 1 To hist.Count
            p = Split(hist(i), vbTab, 2)
            msgs = msgs & IIf(Len(msgs) > 0, ",", "") & JsonMsg(p(0), p(1))
        Next
    End If
    msgs = msgs & IIf(Len(msgs) > 0, ",", "") & JsonMsg("user", message)
    n = m_maxTokens
    If n = 0 And m_provider = "anthropic" Then n = AI_DEFAULT_TOKENS
    s = "{""model"":" & JsonStr(m_model)
    If m_provider = "anthropic" Then
        s = s & ",""max_tokens"":" & n
        If Len(systemPrompt) > 0 Then s = s & ",""system"":" & JsonStr(systemPrompt)
    ElseIf n > 0 Then
        s = s & ",""max_tokens"":" & n
    End If
    BuildRequest = s & ",""messages"":[" & msgs & "]}"
End Function

Private Function JsonMsg(ByVal role As String, ByVal text As String) As String
    JsonMsg = "{""role"":" & JsonStr(role) & ",""content"":" & JsonStr(text) & "}"
End Function

' The reply text from a response body: the first choice's message
' ("openai") or the text blocks of the content ("anthropic").
Private Function ParseReply(ByVal body As String) As String
    Dim root As Variant, v As Variant, blk As Variant, s As String
    On Error GoTo Bad
    JsonParse body, root
    If m_provider = "anthropic" Then
        Set v = JsonPath(root, "content")
        For Each blk In v
            If JsonText(JsonPath(blk, "type")) = "text" Then s = s & JsonText(JsonPath(blk, "text"))
        Next
    Else
        s = JsonText(JsonPath(root, "choices[0].message.content"))
    End If
    ParseReply = s
    Exit Function
Bad:
    ParseReply = ""
End Function

' The service's own error message from an error body, else the start of the body.
Private Function ServiceError(ByVal body As String) As String
    Dim root As Variant, s As String
    On Error Resume Next
    JsonParse body, root
    s = JsonText(JsonPath(root, "error.message"))
    If Len(s) = 0 Then s = JsonText(JsonPath(root, "error"))
    If Len(s) = 0 Then s = JsonText(JsonPath(root, "message"))
    On Error GoTo 0
    If Len(s) = 0 Then s = Left$(Replace(body, vbLf, " "), 200)
    ServiceError = Left$(s, 300)
End Function

' ---------------------------------------------------------------------------
' JSON (small and strict enough for these bodies): objects are
' Scripting.Dictionary, arrays are Collection, numbers Double, true/false
' Boolean, null Null.
' ---------------------------------------------------------------------------

' A JSON string literal for text.
Public Function JsonStr(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536       ' AscW is signed
        Select Case code
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 8: out = out & "\b"
            Case 9: out = out & "\t"
            Case 10: out = out & "\n"
            Case 12: out = out & "\f"
            Case 13: out = out & "\r"
            Case Is < 32: out = out & "\u" & Right$("0000" & Hex$(code), 4)
            Case Else: out = out & ch
        End Select
    Next
    JsonStr = """" & out & """"
End Function

Private Sub JsonParse(ByVal text As String, out As Variant)
    m_js = text: m_jp = 1
    JsonValue out
    m_js = ""
End Sub

Private Sub JsonSkip()
    Do While m_jp <= Len(m_js)
        Select Case Mid$(m_js, m_jp, 1)
            Case " ", vbTab, vbCr, vbLf: m_jp = m_jp + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

Private Sub JsonValue(out As Variant)
    Dim ch As String
    JsonSkip
    If m_jp > Len(m_js) Then Err.Raise 5, , "Unexpected end of JSON"
    ch = Mid$(m_js, m_jp, 1)
    Select Case ch
        Case "{": JsonObject out
        Case "[": JsonArray out
        Case """": out = JsonString()
        Case "t": JsonWord "true": out = True
        Case "f": JsonWord "false": out = False
        Case "n": JsonWord "null": out = Null
        Case Else: JsonNumber out
    End Select
End Sub

Private Sub JsonWord(ByVal w As String)
    If Mid$(m_js, m_jp, Len(w)) <> w Then Err.Raise 5, , "Bad JSON at " & m_jp
    m_jp = m_jp + Len(w)
End Sub

Private Sub JsonNumber(out As Variant)
    Dim st As Long, ch As String
    st = m_jp
    Do While m_jp <= Len(m_js)
        ch = Mid$(m_js, m_jp, 1)
        If InStr("0123456789+-.eE", ch) = 0 Then Exit Do
        m_jp = m_jp + 1
    Loop
    If m_jp = st Then Err.Raise 5, , "Bad JSON at " & m_jp
    out = Val(Mid$(m_js, st, m_jp - st))
End Sub

Private Function JsonString() As String
    Dim sb As String, ch As String, e As String
    m_jp = m_jp + 1                                 ' the opening quote
    Do While m_jp <= Len(m_js)
        ch = Mid$(m_js, m_jp, 1)
        If ch = """" Then m_jp = m_jp + 1: JsonString = sb: Exit Function
        If ch = "\" Then
            e = Mid$(m_js, m_jp + 1, 1)
            Select Case e
                Case "n": sb = sb & vbLf
                Case "r": sb = sb & vbCr
                Case "t": sb = sb & vbTab
                Case "b": sb = sb & Chr$(8)
                Case "f": sb = sb & Chr$(12)
                Case "u"
                    sb = sb & ChrW(CLng("&H" & Mid$(m_js, m_jp + 2, 4)))
                    m_jp = m_jp + 4
                Case Else: sb = sb & e              ' \" \\ \/
            End Select
            m_jp = m_jp + 2
        Else
            sb = sb & ch
            m_jp = m_jp + 1
        End If
    Loop
    Err.Raise 5, , "Unterminated JSON string"
End Function

Private Sub JsonObject(out As Variant)
    Dim d As Object, k As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    m_jp = m_jp + 1
    JsonSkip
    If Mid$(m_js, m_jp, 1) = "}" Then m_jp = m_jp + 1: Set out = d: Exit Sub
    Do
        JsonSkip
        If Mid$(m_js, m_jp, 1) <> """" Then Err.Raise 5, , "Bad JSON object at " & m_jp
        k = JsonString()
        JsonSkip
        If Mid$(m_js, m_jp, 1) <> ":" Then Err.Raise 5, , "Bad JSON object at " & m_jp
        m_jp = m_jp + 1
        JsonValue v
        If IsObject(v) Then Set d(k) = v Else d(k) = v
        JsonSkip
        Select Case Mid$(m_js, m_jp, 1)
            Case ",": m_jp = m_jp + 1
            Case "}": m_jp = m_jp + 1: Set out = d: Exit Sub
            Case Else: Err.Raise 5, , "Bad JSON object at " & m_jp
        End Select
    Loop
End Sub

Private Sub JsonArray(out As Variant)
    Dim c As Collection, v As Variant
    Set c = New Collection
    m_jp = m_jp + 1
    JsonSkip
    If Mid$(m_js, m_jp, 1) = "]" Then m_jp = m_jp + 1: Set out = c: Exit Sub
    Do
        JsonValue v
        c.Add v
        JsonSkip
        Select Case Mid$(m_js, m_jp, 1)
            Case ",": m_jp = m_jp + 1
            Case "]": m_jp = m_jp + 1: Set out = c: Exit Sub
            Case Else: Err.Raise 5, , "Bad JSON array at " & m_jp
        End Select
    Loop
End Sub

' Follow a path such as "choices[0].message.content" through parsed JSON;
' Null when any step is missing.
Private Function JsonPath(ByVal node As Variant, ByVal path As String) As Variant
    Dim part As Variant, name As String, idx As Long, br As Long
    For Each part In Split(path, ".")
        br = InStr(part, "[")
        If br > 0 Then
            name = Left$(part, br - 1)
            idx = CLng(Mid$(part, br + 1, InStr(part, "]") - br - 1))
        Else
            name = part: idx = -1
        End If
        If Len(name) > 0 Then
            If Not IsObject(node) Then JsonPath = Null: Exit Function
            If TypeName(node) <> "Dictionary" Then JsonPath = Null: Exit Function
            If Not node.Exists(name) Then JsonPath = Null: Exit Function
            If IsObject(node(name)) Then Set node = node(name) Else node = node(name)
        End If
        If idx >= 0 Then
            If Not IsObject(node) Then JsonPath = Null: Exit Function
            If TypeName(node) <> "Collection" Then JsonPath = Null: Exit Function
            If idx + 1 > node.Count Then JsonPath = Null: Exit Function
            If IsObject(node(idx + 1)) Then Set node = node(idx + 1) Else node = node(idx + 1)
        End If
    Next
    If IsObject(node) Then Set JsonPath = node Else JsonPath = node
End Function

' A parsed value as text ("" for Null, an object or an array).
Private Function JsonText(ByVal v As Variant) As String
    If IsObject(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    JsonText = CStr(v)
End Function

' ---------------------------------------------------------------------------
' AI-connected parts of a form (the user's choice)
' ---------------------------------------------------------------------------

Private Function PartKey(ByVal formName As String, ByVal role As String) As String
    PartKey = formName & "|" & LCase$(role)
End Function

Private Function PartRoleOk(ByVal role As String) As Boolean
    role = LCase$(role)
    PartRoleOk = (role = "input" Or role = "output" Or role = "send")
End Function

' Choose the control for a role on a form ("input", "output" or "send");
' remembered per user. controlName "" removes the choice. The control is
' checked when the form is open (AIChoosePart) or attached (AIAttach).
Public Sub AISetPart(ByVal formName As String, ByVal role As String, ByVal controlName As String)
    If Not PartRoleOk(role) Then Err.Raise 5, , "An AI part is input, output or send: " & role
    If Len(formName) = 0 Then Err.Raise 5, , "Name the form"
    If Len(controlName) = 0 Then
        On Error Resume Next
        DeleteSetting APP_NAME, SETTINGS_KEY & "Parts", PartKey(formName, role)
    Else
        SaveSetting APP_NAME, SETTINGS_KEY & "Parts", PartKey(formName, role), controlName
    End If
End Sub

' The control chosen for a role on a form, "" for none.
Public Function AIGetPart(ByVal formName As String, ByVal role As String) As String
    AIGetPart = GetSetting(APP_NAME, SETTINGS_KEY & "Parts", PartKey(formName, role), "")
End Function

' Ask the user which control is the form's input, output or send button: a
' prompt listing the form's suitable controls by name. Returns the name chosen,
' "" when cancelled or left empty (which removes the choice). Re-attaches the form.
Public Function AIChoosePart(frm As Form, ByVal role As String) As String
    Dim c As Control, names As String, v As Variant, nm As String
    If Not PartRoleOk(role) Then Err.Raise 5, , "An AI part is input, output or send: " & role
    For Each c In frm.Controls
        If PartControlOk(c, role) Then names = names & IIf(Len(names) > 0, ", ", "") & c.Name
    Next
    If Len(names) = 0 Then
        KitMsgBox "This form has no control that can be the " & role & ".", vbInformation, APP_NAME
        Exit Function
    End If
    v = KitInputBox("Which control is the " & role & "? Enter its name, or leave empty to remove the choice." & _
        vbCrLf & vbCrLf & names, "AI-connected " & role, AIGetPart(frm.Name, role))
    If IsNull(v) Then Exit Function
    nm = Trim$(CStr(v))
    If Len(nm) > 0 Then
        If Not ControlExists(frm, nm) Then
            KitMsgBox "'" & nm & "' isn't a control on this form.", vbExclamation, APP_NAME
            Exit Function
        End If
        If Not PartControlOk(frm.Controls(nm), role) Then
            KitMsgBox "'" & nm & "' can't be the " & role & " (" & IIf(role = "send", "a button", "a text box or combo box") & " is needed).", vbExclamation, APP_NAME
            Exit Function
        End If
    End If
    AISetPart frm.Name, role, nm
    AIAttach frm
    AIChoosePart = nm
End Function

' A readable list of the form's choices, for a settings page.
Public Function AIPartsDescribe(frm As Form) As String
    Dim r As Variant, s As String, nm As String
    For Each r In Array("input", "output", "send")
        nm = AIGetPart(frm.Name, CStr(r))
        s = s & r & ": " & IIf(Len(nm) > 0, nm, "(none)") & vbCrLf
    Next
    AIPartsDescribe = s
End Function

Private Function ControlExists(frm As Form, ByVal nm As String) As Boolean
    Dim c As Control
    For Each c In frm.Controls
        If StrComp(c.Name, nm, vbTextCompare) = 0 Then ControlExists = True: Exit Function
    Next
End Function

Private Function PartControlOk(c As Control, ByVal role As String) As Boolean
    Select Case LCase$(role)
        Case "send": PartControlOk = (c.ControlType = acCommandButton)
        Case Else: PartControlOk = (c.ControlType = acTextBox Or c.ControlType = acComboBox)
    End Select
End Function

' Wire up the form's chosen send button so a click runs AIAskControl. Add to
' the form's Load. Only the send button's OnClick is set, and only when the
' user chose one that is on the form; a form with no choices is unchanged.
Public Sub AIAttach(frm As Form)
    Dim nm As String
    nm = AIGetPart(frm.Name, "send")
    If Len(nm) = 0 Then Exit Sub
    If Not ControlExists(frm, nm) Then Exit Sub
    If frm.Controls(nm).ControlType <> acCommandButton Then Exit Sub
    frm.Controls(nm).OnClick = "=AIAskControl([Form])"
End Sub

' The send button's action: send the input control's text, put the reply in
' the output control. A failure is shown in the output control (and the
' status bar) rather than raised, so a click never opens an error window.
' Without an output choice the reply goes to a message box. Returns the reply.
Public Function AIAskControl(frm As Form) As String
    Dim inp As String, outp As String, reply As String, ok As Boolean
    inp = AIGetPart(frm.Name, "input"): outp = AIGetPart(frm.Name, "output")
    If Len(inp) = 0 Or Not ControlExists(frm, inp) Then
        KitMsgBox "Choose which control holds the message (AIChoosePart).", vbInformation, APP_NAME
        Exit Function
    End If
    ok = AITryAsk(Nz(frm.Controls(inp).Value, ""), reply)
    If Not ok Then reply = AILastError()
    If Len(outp) > 0 And ControlExists(frm, outp) Then
        On Error Resume Next
        frm.Controls(outp).Value = reply
        If Err.Number <> 0 Then LogError "modKitAI.AIAskControl", outp   ' e.g. a locked or calculated control
        On Error GoTo 0
    Else
        KitMsgBox reply, IIf(ok, vbInformation, vbExclamation), APP_NAME
    End If
    If ok Then AIAskControl = reply
End Function

' ---------------------------------------------------------------------------
' The assistant dialog
' ---------------------------------------------------------------------------

' Open the chat dialog dlgAI (BuildAIDialog builds it), with an optional
' first message already typed in.
Public Sub AIOpenAssistant(Optional ByVal prompt As String = "")
    DoCmd.OpenForm "dlgAI", , , , , acDialog, prompt
End Sub

' Build time: the chat dialog dlgAI, a message box, a reply box, Send, New
' chat and Close. Its code-behind is kit\forms\dlgAI.vba.
Public Sub BuildAIDialog()
    Dim c As Control, w As Long, h As Long
    w = px(560): h = px(480)
    NewDialog "AI assistant", w, h
    Ev F, "OnLoad"
    Set c = Lbl("lblReply", "Reply", px(20), px(16), px(200), px(20), 9, True, CLR_MUTED_FG)
    Set c = RO("txtReply", "", px(20), px(38), px(520), px(250), 10, False, CLR_FG)
    c.ScrollBars = 2
    Set c = Lbl("lblMessage", "Message", px(20), px(298), px(200), px(20), 9, True, CLR_MUTED_FG)
    Set c = Txt("txtMessage", "", px(20), px(320), px(520), px(80), 10)
    c.EnterKeyBehavior = True                       ' Enter adds a line; Send sends
    c.ScrollBars = 2
    c.OnKeyDown = "[Event Procedure]"                ' Ctrl+Enter sends
    Set c = Btn("btnSend", "Send", px(20), px(420), px(88), px(30))
    Set c = Btn("btnNew", "New chat", px(116), px(420), px(88), px(30), "outline")
    Set c = Btn("btnClose", "Close", px(452), px(420), px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgAI"
End Sub
