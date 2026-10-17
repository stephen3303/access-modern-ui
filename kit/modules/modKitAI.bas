Option Compare Database
Option Explicit

' Access UI kit - talk to an AI service through its web API.
'
' Send a message, get the reply as text:
'
'   reply = KitAIAsk("Summarize this note: " & Me.txtNote)         ' raises on failure
'   problem = KitAISend(Me.txtQuestion, reply)                      ' "" when reply is filled in
'   If Len(problem) > 0 Then KitMsgBox problem, vbExclamation
'
' The screens (a chat window, settings, and connecting parts of a form to the
' AI) are in modKitAIUI; this module has no screens and works from code, a
' macro or a scheduled task.
'
' Services. Three request formats cover the common services, chosen by the
' Provider setting:
'   openai     OpenAI's chat completions format, which Azure OpenAI, Ollama,
'              LM Studio, vLLM and most gateways also speak (set Endpoint)
'   anthropic  Anthropic's Messages API
'   gemini     Google's generateContent API
' Endpoint and Model default to the provider's own and can be changed in the
' settings, so a new service or model needs no code. {model} in an endpoint is
' replaced by the model name (Gemini's address has it in the path).
'
' Settings (saved per Windows user under APP_NAME, "AI"; KitAIValue reads them,
' KitAISaveSetting writes them, KitAIUseProvider sets the provider and its
' defaults):
'   Provider      openai, anthropic or gemini (default openai)
'   Endpoint      the full web address the request is posted to
'   Model         the model name
'   KeyEnv        name of the environment variable that holds the API key
'   KeyFile       path of a text file whose first line is the API key
'   KeyHeader     header the key goes in, when it isn't the provider's own
'                 (Azure OpenAI wants "api-key")
'   SystemPrompt  standing instructions sent with every message
'   MaxTokens     longest reply to ask for (default 1024; 0 leaves it out,
'                 which models that reject the limit need)
'   TimeoutSec    how long to wait for the reply (default 60)
'
' API keys. No key is in the code or in a setting. The key is read, in this
' order, from the key given for this session (KitAISetKey, kept in memory
' only), the environment variable named by KeyEnv (the provider's usual one,
' such as OPENAI_API_KEY, when KeyEnv is blank; a user or system variable
' works as well as one the process started with), or the file named by
' KeyFile. A service on this computer (localhost) needs no key. Keep a key
' file in a folder only its user can read. The key is never logged or shown.
'
' Conversation. By default each reply is added to the conversation (up to
' AI_MAX_HISTORY messages) so the next message follows on; KitAIReset starts
' again, and keepHistory:=False sends a message on its own.
'
' Other systems:
'   * Data in: KitAIContextFromSql turns a query into text for the prompt, and
'     KitAIAskAboutSql sends a question with the data.
'   * Replies out: KitAISetReplyHandler names a procedure called with each
'     message and reply (log it, file it, send it on).
'   * What goes out: KitAISetMessageFilter names a procedure that sees each
'     message first and returns what to send, or "" to stop it (to remove
'     names, or check ClassificationPermitted).
'   * KitAIRun and KitAIApplyBindings (modKitAIUI) tie a form's own boxes and
'     buttons to the AI.
' What is sent leaves the computer for the service: send only what the
' service is approved for.
'
' The request uses MSXML2.ServerXMLHTTP.6.0, so it follows the computer's
' WinHTTP proxy settings (netsh winhttp show proxy). Access stays responsive
' while it waits, and KitAICancel stops the wait.
'
' The app provides APP_NAME and LogError (modErrors). Needs modKitJson and,
' for key files, modCsv (ReadTextFile).

Public Const AI_MAX_HISTORY As Long = 40         ' messages kept in the conversation
Public Const AI_ERROR As Long = vbObjectError + 5100

Private Const SETTINGS_KEY As String = "AI"

Private m_key As String                  ' this session's key (never saved)
Private m_hist As Collection             ' the conversation: Array(role, text)
Private m_busy As Boolean                ' a request is waiting for its reply
Private m_cancel As Boolean
Private m_handler As String              ' procedure called with each message and reply
Private m_filter As String               ' procedure that sees each message first
Private m_mockOn As Boolean              ' tests and offline demos: a stored reply instead of the web
Private m_mockStatus As Long
Private m_mockBody As String

' ---------------------------------------------------------------------------
' Settings
' ---------------------------------------------------------------------------

' The providers, as "key|name" for a list.
Public Function KitAIProviders() As Variant
    KitAIProviders = Array("openai|OpenAI or compatible service", "anthropic|Anthropic", "gemini|Google Gemini")
End Function

' A provider's own Endpoint, Model or KeyEnv ("" for any other name).
Public Function KitAIDefault(ByVal provider As String, ByVal setting As String) As String
    Select Case LCase$(provider) & "." & setting
        Case "openai.Endpoint": KitAIDefault = "https://api.openai.com/v1/chat/completions"
        Case "openai.Model": KitAIDefault = "gpt-4o-mini"
        Case "openai.KeyEnv": KitAIDefault = "OPENAI_API_KEY"
        Case "anthropic.Endpoint": KitAIDefault = "https://api.anthropic.com/v1/messages"
        Case "anthropic.Model": KitAIDefault = "claude-sonnet-5-5"
        Case "anthropic.KeyEnv": KitAIDefault = "ANTHROPIC_API_KEY"
        Case "gemini.Endpoint": KitAIDefault = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
        Case "gemini.Model": KitAIDefault = "gemini-2.0-flash"
        Case "gemini.KeyEnv": KitAIDefault = "GEMINI_API_KEY"
    End Select
End Function

' True for a provider this module can talk to.
Public Function KitAIKnownProvider(ByVal provider As String) As Boolean
    Select Case LCase$(provider)
        Case "openai", "anthropic", "gemini": KitAIKnownProvider = True
    End Select
End Function

' A setting's value: the saved one, else the provider's default (Endpoint,
' Model, KeyEnv), else the setting's own default.
Public Function KitAIValue(ByVal setting As String) As String
    Dim s As String
    On Error Resume Next
    s = GetSetting(APP_NAME, SETTINGS_KEY, setting, "")
    On Error GoTo 0
    If Len(s) = 0 Then
        Select Case setting
            Case "Provider": s = "openai"
            Case "Endpoint", "Model", "KeyEnv": s = KitAIDefault(KitAIValue("Provider"), setting)
            Case "MaxTokens": s = "1024"
            Case "TimeoutSec": s = "60"
        End Select
    End If
    KitAIValue = s
End Function

' Save a setting; "" removes it, so the default applies again.
Public Sub KitAISaveSetting(ByVal setting As String, ByVal value As String)
    If Len(value) = 0 Then
        On Error Resume Next
        DeleteSetting APP_NAME, SETTINGS_KEY, setting
        On Error GoTo 0
    Else
        SaveSetting APP_NAME, SETTINGS_KEY, setting, value
    End If
End Sub

' Use a provider: saves it and clears Endpoint, Model and KeyEnv, so its own
' defaults apply, then any of them given. Raises error 5 for an unknown
' provider.
Public Sub KitAIUseProvider(ByVal provider As String, Optional ByVal endpoint As String = "", _
        Optional ByVal model As String = "", Optional ByVal keyEnv As String = "")
    If Not KitAIKnownProvider(provider) Then Err.Raise 5, , "Not an AI provider this kit knows: " & provider
    KitAISaveSetting "Provider", LCase$(provider)
    KitAISaveSetting "Endpoint", endpoint
    KitAISaveSetting "Model", model
    KitAISaveSetting "KeyEnv", keyEnv
End Sub

' ---------------------------------------------------------------------------
' The API key
' ---------------------------------------------------------------------------

' Use key for this session only. It is kept in memory, never saved; "" forgets it.
Public Sub KitAISetKey(ByVal key As String)
    m_key = Trim$(key)
End Sub

' The key to send: the session's, else the environment variable's, else the key
' file's; "" when there is none (or the key file can't be read, which is logged).
Public Function KitAIApiKey() As String
    Dim s As String, path As String
    If Len(m_key) > 0 Then KitAIApiKey = m_key: Exit Function
    s = EnvValue(KitAIValue("KeyEnv"))
    If Len(s) > 0 Then KitAIApiKey = s: Exit Function
    path = Trim$(KitAIValue("KeyFile"))
    If Len(path) > 0 Then KitAIApiKey = KeyFromFile(path)
End Function

' Where the key comes from, in words, or "" when there is none. Never the key.
Public Function KitAIKeySource() As String
    Dim env As String, path As String
    If Len(m_key) > 0 Then KitAIKeySource = "a key entered for this session": Exit Function
    env = KitAIValue("KeyEnv")
    If Len(EnvValue(env)) > 0 Then KitAIKeySource = "the environment variable " & env: Exit Function
    path = Trim$(KitAIValue("KeyFile"))
    If Len(path) > 0 Then
        If Len(KeyFromFile(path)) > 0 Then KitAIKeySource = "the key file " & path
    End If
End Function

' The named variable from this process, else from the user's, else the system's
' variables (one set after Access started is found too). "" when blank or unset.
Private Function EnvValue(ByVal name As String) As String
    Dim sh As Object, s As String, scope As Variant
    name = Trim$(name)
    If Len(name) = 0 Then Exit Function
    s = Environ$(name)
    If Len(s) = 0 Then
        On Error Resume Next
        Set sh = CreateObject("WScript.Shell")
        For Each scope In Array("Process", "User", "System")
            s = sh.Environment(CStr(scope)).Item(name)
            If Len(s) > 0 Then Exit For
        Next
        On Error GoTo 0
    End If
    EnvValue = Trim$(s)
End Function

' The first non-blank line of the key file, or "" when it is missing or empty
' (an unreadable file is logged, without its contents).
Private Function KeyFromFile(ByVal path As String) As String
    Dim line As Variant
    On Error GoTo Fail
    If Not FileExists(path) Then Exit Function
    For Each line In Split(Replace(ReadTextFile(path), vbCr, ""), vbLf)
        If Len(Trim$(line)) > 0 Then KeyFromFile = Trim$(line): Exit Function
    Next
    Exit Function
Fail:
    LogError "modKitAI.KeyFromFile", "key file " & path
End Function

' ---------------------------------------------------------------------------
' Sending a message
' ---------------------------------------------------------------------------

' Send message and return the reply; raises AI_ERROR (the problem in its
' description) when there is none. See KitAISend.
Public Function KitAIAsk(ByVal message As String, Optional ByVal keepHistory As Boolean = True, _
        Optional ByVal system As String = "") As String
    Dim problem As String, reply As String
    problem = KitAISend(message, reply, keepHistory, system)
    If Len(problem) > 0 Then Err.Raise AI_ERROR, "modKitAI.KitAIAsk", problem
    KitAIAsk = reply
End Function

' Send message to the AI service. Returns "" and fills in reply, or returns what
' went wrong in words for the user (reply is then ""). keepHistory adds the
' message and reply to the conversation; system replaces the SystemPrompt
' setting for this message. Never raises for a failed request.
Public Function KitAISend(ByVal message As String, ByRef reply As String, Optional ByVal keepHistory As Boolean = True, _
        Optional ByVal system As String = "") As String
    Dim provider As String, endpoint As String, model As String, key As String
    Dim body As String, url As String, headers As Variant, status As Long, raw As String
    Dim work As Collection, problem As String, sent As String, v As Variant
    reply = ""
    On Error GoTo Fail
    If m_busy Then KitAISend = "Another question is still waiting for its reply.": Exit Function
    sent = message
    If Len(m_filter) > 0 Then
        v = Application.Run(m_filter, message)
        If IsNull(v) Then sent = "" Else sent = CStr(v)
        If Len(Trim$(sent)) = 0 Then KitAISend = "The message was stopped before it was sent.": Exit Function
    End If
    If Len(Trim$(sent)) = 0 Then KitAISend = "Type a message first.": Exit Function
    provider = LCase$(KitAIValue("Provider"))
    If Not KitAIKnownProvider(provider) Then KitAISend = "The AI provider """ & provider & """ isn't one this app can use.": Exit Function
    endpoint = Trim$(KitAIValue("Endpoint"))
    model = Trim$(KitAIValue("Model"))
    If Len(endpoint) = 0 Then KitAISend = "No AI service address is set. Open the AI settings.": Exit Function
    If Len(model) = 0 Then KitAISend = "No AI model is set. Open the AI settings.": Exit Function
    If Left$(LCase$(endpoint), 8) <> "https://" And Not IsLocalEndpoint(endpoint) And Not m_mockOn Then
        KitAISend = "The AI service address must start with https:// (only a service on this computer may use http://)."
        Exit Function
    End If
    key = KitAIApiKey()
    If Len(key) = 0 And Not IsLocalEndpoint(endpoint) And Not m_mockOn Then KitAISend = NoKeyMessage(): Exit Function
    If Len(system) = 0 Then system = KitAIValue("SystemPrompt")

    Set work = HistoryCopy()
    work.Add Array("user", sent)
    body = KitAIRequestBody(provider, model, system, work, CLng(Val(KitAIValue("MaxTokens"))))
    url = Replace(endpoint, "{model}", model)
    headers = KitAIHeaders(provider, key, KitAIValue("KeyHeader"))

    m_busy = True: m_cancel = False
    problem = PostJson(url, headers, body, CLng(Val(KitAIValue("TimeoutSec"))), status, raw)
    m_busy = False
    If Len(problem) > 0 Then KitAISend = problem: Exit Function
    problem = KitAIReplyProblem(provider, status, raw, reply)
    If Len(problem) > 0 Then reply = "": KitAISend = problem: Exit Function

    If keepHistory Then
        Set m_hist = work
        m_hist.Add Array("assistant", reply)
        TrimHistory
    End If
    CallReplyHandler sent, reply
    Exit Function
Fail:
    m_busy = False
    reply = ""
    LogError "modKitAI.KitAISend", "provider " & provider & ", model " & model
    KitAISend = "The AI request failed: " & Err.Description
End Function

Private Function NoKeyMessage() As String
    Dim env As String
    env = KitAIValue("KeyEnv")
    NoKeyMessage = "There is no API key for the AI service. Set the " & IIf(Len(env) > 0, env, "key's") & _
        " environment variable, choose a key file, or enter a key for this session in the AI settings."
End Function

' True for an address on this computer, which needs no key and may use http://.
Private Function IsLocalEndpoint(ByVal url As String) As Boolean
    Dim s As String
    s = LCase$(url)
    IsLocalEndpoint = (Left$(s, 16) = "http://localhost") Or (Left$(s, 17) = "https://localhost") Or _
        (Left$(s, 16) = "http://127.0.0.1") Or (Left$(s, 17) = "https://127.0.0.1") Or _
        (Left$(s, 12) = "http://[::1]") Or (Left$(s, 13) = "https://[::1]")
End Function

' Stop waiting for a reply (from a Cancel button; the wait lets events run).
Public Sub KitAICancel()
    m_cancel = True
End Sub

' True while a request is waiting for its reply.
Public Function KitAIBusy() As Boolean
    KitAIBusy = m_busy
End Function

' ---------------------------------------------------------------------------
' The conversation
' ---------------------------------------------------------------------------

' Start the conversation again.
Public Sub KitAIReset()
    Set m_hist = Nothing
End Sub

' How many messages (the user's and the AI's) the conversation holds.
Public Function KitAIHistoryCount() As Long
    If Not m_hist Is Nothing Then KitAIHistoryCount = m_hist.Count
End Function

Private Function HistoryCopy() As Collection
    Dim c As Collection, m As Variant
    Set c = New Collection
    If Not m_hist Is Nothing Then
        For Each m In m_hist
            c.Add m
        Next
    End If
    Set HistoryCopy = c
End Function

' Keep the newest AI_MAX_HISTORY messages, starting on a user message.
Private Sub TrimHistory()
    Do While m_hist.Count > AI_MAX_HISTORY
        m_hist.Remove 1
    Loop
    Do While m_hist.Count > 0
        If m_hist(1)(0) = "user" Then Exit Do
        m_hist.Remove 1
    Loop
End Sub

' ---------------------------------------------------------------------------
' Requests and replies (the formats of each service)
' ---------------------------------------------------------------------------

' The JSON body for provider: history is the whole conversation so far, as
' Array(role, text) with the roles "user" and "assistant", ending with the
' message to answer. system "" sends none; maxTokens 0 leaves the limit out
' (Anthropic needs one, so 1024 is used).
Public Function KitAIRequestBody(ByVal provider As String, ByVal model As String, ByVal system As String, _
        history As Collection, ByVal maxTokens As Long) As String
    Dim m As Variant, s As String, first As Boolean
    first = True
    Select Case LCase$(provider)
        Case "anthropic"
            If maxTokens <= 0 Then maxTokens = 1024
            s = "{""model"":" & JsonQuote(model) & ",""max_tokens"":" & maxTokens
            If Len(system) > 0 Then s = s & ",""system"":" & JsonQuote(system)
            s = s & ",""messages"":["
            For Each m In history
                If Not first Then s = s & ","
                first = False
                s = s & "{""role"":" & JsonQuote(CStr(m(0))) & ",""content"":" & JsonQuote(CStr(m(1))) & "}"
            Next
            s = s & "]}"
        Case "gemini"
            s = "{""contents"":["
            For Each m In history
                If Not first Then s = s & ","
                first = False
                s = s & "{""role"":" & JsonQuote(IIf(m(0) = "assistant", "model", "user")) & _
                    ",""parts"":[{""text"":" & JsonQuote(CStr(m(1))) & "}]}"
            Next
            s = s & "]"
            If Len(system) > 0 Then s = s & ",""systemInstruction"":{""parts"":[{""text"":" & JsonQuote(system) & "}]}"
            If maxTokens > 0 Then s = s & ",""generationConfig"":{""maxOutputTokens"":" & maxTokens & "}"
            s = s & "}"
        Case Else                                ' openai, and services that speak it
            s = "{""model"":" & JsonQuote(model) & ",""messages"":["
            If Len(system) > 0 Then
                s = s & "{""role"":""system"",""content"":" & JsonQuote(system) & "}"
                first = False
            End If
            For Each m In history
                If Not first Then s = s & ","
                first = False
                s = s & "{""role"":" & JsonQuote(CStr(m(0))) & ",""content"":" & JsonQuote(CStr(m(1))) & "}"
            Next
            s = s & "]"
            If maxTokens > 0 Then s = s & ",""max_tokens"":" & maxTokens
            s = s & "}"
    End Select
    KitAIRequestBody = s
End Function

' The extra request headers, as "Name: value" strings (Content-Type is added by
' the sender). keyHeader replaces the provider's own header for the key and
' sends the key as it is. No key, no key header.
Public Function KitAIHeaders(ByVal provider As String, ByVal key As String, Optional ByVal keyHeader As String = "") As Variant
    Dim h As String, extra As String
    keyHeader = Trim$(keyHeader)
    If Len(key) > 0 Then
        If Len(keyHeader) > 0 Then
            h = keyHeader & ": " & key
        Else
            Select Case LCase$(provider)
                Case "anthropic": h = "x-api-key: " & key
                Case "gemini": h = "x-goog-api-key: " & key
                Case Else: h = "Authorization: Bearer " & key
            End Select
        End If
    End If
    If LCase$(provider) = "anthropic" Then extra = "anthropic-version: 2023-06-01"
    If Len(h) > 0 And Len(extra) > 0 Then
        KitAIHeaders = Array(h, extra)
    ElseIf Len(h) > 0 Then
        KitAIHeaders = Array(h)
    ElseIf Len(extra) > 0 Then
        KitAIHeaders = Array(extra)
    Else
        KitAIHeaders = Array()
    End If
End Function

' Read the service's answer. Returns "" and the reply text, or the problem in
' words (an error the service reported, a refused key, no text in the answer).
' status is the HTTP status and raw the body.
Public Function KitAIReplyProblem(ByVal provider As String, ByVal status As Long, ByVal raw As String, ByRef reply As String) As String
    Dim root As Variant, why As String, i As Long, n As Long, parts As String, t As String
    reply = ""
    On Error GoTo BadJson
    If Len(Trim$(raw)) = 0 Then
        If status >= 400 Or status = 0 Then KitAIReplyProblem = StatusProblem(status, ""): Exit Function
        KitAIReplyProblem = "The AI service sent an empty answer.": Exit Function
    End If
    Set root = JsonParse(raw)
    On Error GoTo 0
    If Not IsObject(root) Then GoTo NoText
    why = KitAIErrorText(root)
    If status >= 400 Or Len(why) > 0 Then KitAIReplyProblem = StatusProblem(status, why): Exit Function
    Select Case LCase$(provider)
        Case "anthropic"
            n = JsonCount(root, "content")
            For i = 0 To n - 1
                If JsonText(root, "content." & i & ".type") = "text" Then parts = parts & JsonText(root, "content." & i & ".text")
            Next
        Case "gemini"
            n = JsonCount(root, "candidates.0.content.parts")
            For i = 0 To n - 1
                parts = parts & JsonText(root, "candidates.0.content.parts." & i & ".text")
            Next
            If Len(parts) = 0 Then
                t = JsonText(root, "promptFeedback.blockReason")
                If Len(t) > 0 Then KitAIReplyProblem = "The AI service blocked the message (" & t & "). Try rewording it.": Exit Function
            End If
        Case Else
            parts = JsonText(root, "choices.0.message.content")
    End Select
    If Len(parts) = 0 Then GoTo NoText
    reply = parts
    Exit Function
NoText:
    KitAIReplyProblem = "The AI service answered with no text."
    Exit Function
BadJson:
    If status >= 400 Then KitAIReplyProblem = StatusProblem(status, "") Else KitAIReplyProblem = "The AI service sent an answer this app couldn't read."
End Function

' The message inside a service's error answer ({"error":{"message":...}},
' {"error":"..."} or {"message":...}), or "".
Public Function KitAIErrorText(root As Variant) As String
    Dim s As String
    s = JsonText(root, "error.message")
    If Len(s) = 0 Then s = JsonText(root, "error")
    If Len(s) = 0 Then s = JsonText(root, "message")
    KitAIErrorText = s
End Function

Private Function StatusProblem(ByVal status As Long, ByVal why As String) As String
    Dim s As String
    Select Case status
        Case 0: s = "The AI service couldn't be reached."
        Case 200 To 399: s = "The AI service reported a problem."
        Case 401, 403: s = "The AI service didn't accept the API key (HTTP " & status & ")."
        Case 404: s = "The AI service address or model wasn't found (HTTP 404)."
        Case 429: s = "The AI service is busy or the usage limit was reached (HTTP 429). Try again shortly."
        Case Else
            If status >= 500 Then s = "The AI service had a problem (HTTP " & status & ")." Else s = "The AI service refused the request (HTTP " & status & ")."
    End Select
    If Len(why) > 0 Then s = s & " " & Left$(why, 400)
    StatusProblem = s
End Function

' ---------------------------------------------------------------------------
' The web request
' ---------------------------------------------------------------------------

' POST body to url. Returns "" with the HTTP status and the answer's text, or
' why the request couldn't be made or finished (not a bad status: the caller
' reads that from the answer).
Private Function PostJson(ByVal url As String, headers As Variant, ByVal body As String, ByVal timeoutSec As Long, _
        ByRef status As Long, ByRef raw As String) As String
    Dim req As Object, h As Variant, p As Long, t0 As Single, waited As Single
    On Error GoTo Fail
    status = 0: raw = ""
    If m_mockOn Then status = m_mockStatus: raw = m_mockBody: Exit Function
    If timeoutSec < 5 Then timeoutSec = 60
    Set req = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    req.setTimeouts 10000, 15000, 30000, timeoutSec * 1000
    req.Open "POST", url, True
    req.setRequestHeader "Content-Type", "application/json"
    req.setRequestHeader "Accept", "application/json"
    For Each h In headers
        p = InStr(h, ":")
        If p > 1 Then req.setRequestHeader Left$(h, p - 1), Trim$(Mid$(h, p + 1))
    Next
    req.send body
    t0 = Timer
    Do While Not req.waitForResponse(1)
        DoEvents
        waited = Timer - t0
        If waited < 0 Then waited = waited + 86400
        If m_cancel Then
            req.abort
            PostJson = "Cancelled."
            Exit Function
        End If
        If waited > timeoutSec Then
            req.abort
            PostJson = "The AI service didn't answer within " & timeoutSec & " seconds."
            Exit Function
        End If
    Loop
    status = req.status
    raw = Utf8ToText(req.responseBody)
    Exit Function
Fail:
    ' The key is in the headers, never in the message: report only what went wrong.
    PostJson = "The AI service couldn't be reached: " & Err.Description
    LogError "modKitAI.PostJson", "host " & HostOf(url)         ' not the whole address: it could carry a key
End Function

' The host part of a web address ("api.openai.com").
Private Function HostOf(ByVal url As String) As String
    Dim p As Long, q As Long
    p = InStr(url, "//")
    p = IIf(p > 0, p + 2, 1)
    q = InStr(p, url, "/")
    If q = 0 Then q = Len(url) + 1
    HostOf = Mid$(url, p, q - p)
End Function

' An answer's bytes read as UTF-8 text.
Private Function Utf8ToText(bytes As Variant) As String
    Dim st As Object
    If Not IsArray(bytes) Then Exit Function
    If UBound(bytes) < LBound(bytes) Then Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write bytes
    st.Position = 0
    st.Type = 2
    st.Charset = "utf-8"
    Utf8ToText = st.ReadText
    st.Close
End Function

' Answer with status and body instead of using the web, until KitAIClearMock:
' for tests and for trying the screens with no service. Not for production.
Public Sub KitAIUseMock(ByVal status As Long, ByVal body As String)
    m_mockOn = True: m_mockStatus = status: m_mockBody = body
End Sub

Public Sub KitAIClearMock()
    m_mockOn = False: m_mockStatus = 0: m_mockBody = ""
End Sub

' ---------------------------------------------------------------------------
' Other systems
' ---------------------------------------------------------------------------

' Call procedure procName (a Public Sub or Function of the app) with the message
' and the reply after each successful KitAISend; "" stops. A failure in it is
' logged and doesn't affect the reply.
'   Public Sub OnAiReply(ByVal message As String, ByVal reply As String)
Public Sub KitAISetReplyHandler(ByVal procName As String)
    m_handler = Trim$(procName)
End Sub

' Call procedure procName with each message before it is sent. It returns the
' text to send (changed if need be), or "" to stop the message. A failure in it
' stops the message and is reported to the user.
'   Public Function AiFilter(ByVal message As String) As String
Public Sub KitAISetMessageFilter(ByVal procName As String)
    m_filter = Trim$(procName)
End Sub

Private Sub CallReplyHandler(ByVal message As String, ByVal reply As String)
    If Len(m_handler) = 0 Then Exit Sub
    On Error GoTo Fail
    Application.Run m_handler, message, reply
    Exit Sub
Fail:
    LogError "modKitAI.CallReplyHandler", "procedure " & m_handler
End Sub

' A query's rows as text for a prompt: a header line of the column names, then
' one line per row with the values split by tabs (line breaks and tabs inside a
' value become spaces). Stops at maxRows rows or maxChars characters and says
' so. Raises the query's error if it fails.
Public Function KitAIContextFromSql(ByVal sql As String, Optional ByVal maxRows As Long = 50, _
        Optional ByVal maxChars As Long = 20000) As String
    Dim rs As DAO.Recordset, i As Long, n As Long, line As String, buf As String, cut As Boolean
    Dim num As Long, desc As String
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    On Error GoTo Fail
    For i = 0 To rs.Fields.Count - 1
        line = line & IIf(i > 0, vbTab, "") & rs.Fields(i).Name
    Next
    buf = line
    Do While Not rs.EOF
        If n >= maxRows Then cut = True: Exit Do
        line = ""
        For i = 0 To rs.Fields.Count - 1
            line = line & IIf(i > 0, vbTab, "") & OneLine(rs.Fields(i).Value)
        Next
        If Len(buf) + Len(line) + 2 > maxChars Then cut = True: Exit Do
        buf = buf & vbLf & line
        n = n + 1
        rs.MoveNext
    Loop
    If cut Then buf = buf & vbLf & "(more rows not shown)"
    rs.Close
    KitAIContextFromSql = buf
    Exit Function
Fail:
    num = Err.Number: desc = Err.Description
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    Err.Raise num, "modKitAI.KitAIContextFromSql", desc
End Function

Private Function OneLine(ByVal v As Variant) As String
    Dim s As String
    If IsNull(v) Then Exit Function
    On Error Resume Next
    s = CStr(v)
    If Err.Number <> 0 Then s = "(not text)"
    On Error GoTo 0
    s = Replace(Replace(Replace(Replace(s, vbCrLf, " "), vbLf, " "), vbCr, " "), vbTab, " ")
    OneLine = s
End Function

' Ask question about the rows a query returns (KitAIContextFromSql sends them
' with it). Returns "" and fills in reply, or the problem, like KitAISend.
' The conversation isn't kept unless keepHistory is True.
Public Function KitAIAskAboutSql(ByVal question As String, ByVal sql As String, ByRef reply As String, _
        Optional ByVal keepHistory As Boolean = False, Optional ByVal maxRows As Long = 50) As String
    Dim data As String
    On Error GoTo Fail
    data = KitAIContextFromSql(sql, maxRows)
    KitAIAskAboutSql = KitAISend("Here is data from the database, as tab-separated rows:" & vbLf & data & vbLf & vbLf & question, _
        reply, keepHistory)
    Exit Function
Fail:
    reply = ""
    LogError "modKitAI.KitAIAskAboutSql", "query " & Left$(sql, 120)
    KitAIAskAboutSql = "The data for the question couldn't be read: " & Err.Description
End Function
