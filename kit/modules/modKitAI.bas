Option Compare Database
Option Explicit

' Access UI kit - connect the app to an AI service over its web API: send a
' message, get the reply back, put it in the app's own forms and fields, and
' let the user choose which parts of the interface use it.
'
'   KitAIConfigure "anthropic"                 ' once, at start-up (or the user does, in KitAISettings)
'   reply = KitAIAsk("Summarize: " & Me.txtNotes)       ' one question, one answer
'   reply = KitAIAsk("And in one line?", , True)         ' keepHistory: a conversation
'   KitAIFill Me.txtSummary, "Summarize: " & Me.txtNotes  ' straight into a field
'   KitAIAssist Me, "txtNotes", "Fix the spelling"       ' the chat window, seeded with the field
'   KitAIChat                                            ' the chat window on its own
'
' Services (the provider key). One request/response shape each, so adding a
' service means one more branch in AIFormat, AIDefaultEndpoint and
' AIDefaultModel:
'   openai     OpenAI chat completions
'   anthropic  Anthropic messages
'   gemini     Google Gemini generateContent
'   azure      Azure OpenAI (chat completions; the endpoint is your deployment's URL)
'   local      any OpenAI-compatible server on this PC or network (Ollama, LM Studio); no key
'   custom     any other OpenAI-compatible service (you give the endpoint; key optional)
' The endpoint, model and system prompt can be changed per service
' (KitAIConfigure, or the user in KitAISettings). They are saved per Windows
' user under APP_NAME ("AI" settings); the key is not.
'
' The API KEY is never in the code or the database. KitAIKey() reads, in this
' order: the key set for this session (KitAISetKey; KitAISettings asks for it
' and holds it only until Access closes), the environment variable
' KIT_AI_KEY_<PROVIDER> (KIT_AI_KEY_ANTHROPIC), then KIT_AI_API_KEY. Set one
' in the user's Windows environment (a User variable) so nobody types it.
'
' Interface parts. The app names the places that can use AI, once, at start-up:
'   KitAIPart "notes", "Notes assistant"
'   KitAIPart "summary", "Summary on the dashboard"
' and tags the controls that belong to each one on its form: Tag = "ai:notes".
' The user ticks the parts they want in KitAISettings (nothing is on until
' they do, so nothing is sent to a service they have not chosen). A form's
' Load then calls KitAIApplyParts Me, which hides the tagged controls whose
' part is off, and code checks KitAIPartOn("notes") before offering AI.
'
' Dialogs: BuildAIDialog builds dlgAI (the chat) and BuildAISettingsDialog
' builds dlgAISettings; call both with the app's dialogs and copy
' kit\forms\dlgAI.vba and dlgAISettings.vba to src\forms.
'
' Errors. Everything that talks to the service raises an error numbered
' AI_ERR, with the service's own message ("Incorrect API key provided") or
' what went wrong ("No API key"); a window or field helper shows it to the
' user, and code calling KitAIAsk handles it. Needs Windows' MSXML (always
' installed) and, for the reply text, ADODB.

Public Const AI_ERR As Long = vbObjectError + 5400
Public Const AI_TAG As String = "ai:"           ' a control's Tag: this, then its part key

Private Const SETTINGS_KEY As String = "AI"
Private Const MAX_TURNS As Long = 20            ' messages kept in a conversation
Private Const MAX_TOKENS As Long = 1024         ' longest reply asked for
Private Const DEFAULT_TIMEOUT_S As Long = 60
Private Const DEFAULT_SYSTEM As String = "You are a concise assistant inside a business database application. " & _
    "Answer plainly, without markdown, and say so when you are not sure."

Private m_key As String                         ' the key for this session
Private m_history As Collection                 ' conversation: Array(role, text), role "user" or "assistant"
Private m_parts As Collection                   ' registered parts: Array(key, caption)

' The chat window (dlgAI): what it opens with, where "Use reply" puts the
' reply, and the reply that was used.
Private m_chatOpen As Boolean
Private m_chatSeed As String
Private m_chatTarget As Control
Private m_chatUsed As String
Private m_saved As Boolean                      ' dlgAISettings saved

' ---------------------------------------------------------------------------
' Services and settings
' ---------------------------------------------------------------------------

' The service keys, in the order a picker shows them.
Public Function KitAIProviders() As Variant
    KitAIProviders = Array("openai", "anthropic", "gemini", "azure", "local", "custom")
End Function

' What a service is called to the user.
Public Function KitAIProviderName(ByVal provider As String) As String
    Select Case LCase$(provider)
        Case "openai": KitAIProviderName = "OpenAI"
        Case "anthropic": KitAIProviderName = "Anthropic (Claude)"
        Case "gemini": KitAIProviderName = "Google Gemini"
        Case "azure": KitAIProviderName = "Azure OpenAI"
        Case "local": KitAIProviderName = "Local server (Ollama, LM Studio)"
        Case "custom": KitAIProviderName = "Other OpenAI-compatible service"
        Case Else: KitAIProviderName = provider
    End Select
End Function

' The request/response shape a service uses: "openai", "anthropic" or "gemini".
' Raises error 5 for an unknown service.
Public Function AIFormat(ByVal provider As String) As String
    Select Case LCase$(provider)
        Case "openai", "azure", "local", "custom": AIFormat = "openai"
        Case "anthropic": AIFormat = "anthropic"
        Case "gemini": AIFormat = "gemini"
        Case Else: Err.Raise 5, , "Unknown AI service: " & provider
    End Select
End Function

Public Function AIDefaultEndpoint(ByVal provider As String) As String
    Select Case LCase$(provider)
        Case "openai": AIDefaultEndpoint = "https://api.openai.com/v1/chat/completions"
        Case "anthropic": AIDefaultEndpoint = "https://api.anthropic.com/v1/messages"
        Case "gemini": AIDefaultEndpoint = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
        Case "local": AIDefaultEndpoint = "http://localhost:11434/v1/chat/completions"
        Case "azure", "custom": AIDefaultEndpoint = ""          ' the user's own
        Case Else: Err.Raise 5, , "Unknown AI service: " & provider
    End Select
End Function

Public Function AIDefaultModel(ByVal provider As String) As String
    Select Case LCase$(provider)
        Case "openai": AIDefaultModel = "gpt-4o-mini"
        Case "anthropic": AIDefaultModel = "claude-sonnet-5-5"
        Case "gemini": AIDefaultModel = "gemini-2.0-flash"
        Case "local": AIDefaultModel = "llama3.1"
        Case "azure", "custom": AIDefaultModel = ""
        Case Else: Err.Raise 5, , "Unknown AI service: " & provider
    End Select
End Function

' True for a service that cannot be used without a key.
Public Function AINeedsKey(ByVal provider As String) As Boolean
    AINeedsKey = (LCase$(provider) <> "local" And LCase$(provider) <> "custom")
End Function

' The service in use (openai until one is chosen).
Public Function KitAIProvider() As String
    Dim p As String
    p = LCase$(GetSetting(APP_NAME, SETTINGS_KEY, "Provider", "openai"))
    If InProviders(p) Then KitAIProvider = p Else KitAIProvider = "openai"
End Function

Private Function InProviders(ByVal p As String) As Boolean
    Dim v As Variant
    For Each v In KitAIProviders()
        If v = p Then InProviders = True: Exit Function
    Next
End Function

' The endpoint, model and system prompt saved for a service, else its defaults.
Public Function KitAIEndpoint(Optional ByVal provider As String = "") As String
    If Len(provider) = 0 Then provider = KitAIProvider()
    KitAIEndpoint = GetSetting(APP_NAME, SETTINGS_KEY, provider & "_Endpoint", AIDefaultEndpoint(provider))
End Function

Public Function KitAIModel(Optional ByVal provider As String = "") As String
    If Len(provider) = 0 Then provider = KitAIProvider()
    KitAIModel = GetSetting(APP_NAME, SETTINGS_KEY, provider & "_Model", AIDefaultModel(provider))
End Function

Public Function KitAISystemPrompt() As String
    KitAISystemPrompt = GetSetting(APP_NAME, SETTINGS_KEY, "System", DEFAULT_SYSTEM)
End Function

' Choose the service and, optionally, its model, endpoint and the system
' prompt (what the AI is told it is). Leave a part "" to keep what is saved
' for it. Raises error 5 for an unknown service. Starts a new conversation.
Public Sub KitAIConfigure(ByVal provider As String, Optional ByVal model As String = "", _
        Optional ByVal endpoint As String = "", Optional ByVal systemPrompt As String = "")
    provider = LCase$(provider)
    If Not InProviders(provider) Then Err.Raise 5, , "Unknown AI service: " & provider
    SaveSetting APP_NAME, SETTINGS_KEY, "Provider", provider
    If Len(model) > 0 Then SaveSetting APP_NAME, SETTINGS_KEY, provider & "_Model", model
    If Len(endpoint) > 0 Then SaveSetting APP_NAME, SETTINGS_KEY, provider & "_Endpoint", endpoint
    If Len(systemPrompt) > 0 Then SaveSetting APP_NAME, SETTINGS_KEY, "System", systemPrompt
    KitAIReset
End Sub

' ---------------------------------------------------------------------------
' API key (never stored)
' ---------------------------------------------------------------------------

' Hold a key for this Access session only. Pass "" to forget it.
Public Sub KitAISetKey(ByVal key As String)
    m_key = Trim$(key)
End Sub

' The key to use for a service: the session key, else KIT_AI_KEY_<SERVICE>,
' else KIT_AI_API_KEY, else "".
Public Function KitAIKey(Optional ByVal provider As String = "") As String
    If Len(provider) = 0 Then provider = KitAIProvider()
    If Len(m_key) > 0 Then KitAIKey = m_key: Exit Function
    KitAIKey = Trim$(Environ$("KIT_AI_KEY_" & UCase$(provider)))
    If Len(KitAIKey) = 0 Then KitAIKey = Trim$(Environ$("KIT_AI_API_KEY"))
End Function

' True when the chosen service has what it needs to be called: a key if it
' needs one, and an endpoint and model.
Public Function KitAIReady() As Boolean
    Dim p As String
    p = KitAIProvider()
    If AINeedsKey(p) And Len(KitAIKey(p)) = 0 Then Exit Function
    KitAIReady = (Len(KitAIEndpoint(p)) > 0 And Len(KitAIModel(p)) > 0)
End Function

' ---------------------------------------------------------------------------
' Message in, reply out
' ---------------------------------------------------------------------------

' Send message to the chosen service and return its reply. systemPrompt ""
' uses the saved one. keepHistory True continues one conversation (the last
' MAX_TURNS messages are sent each time; KitAIReset starts again); False asks
' a single question that is not remembered. Raises AI_ERR (see the top).
Public Function KitAIAsk(ByVal message As String, Optional ByVal systemPrompt As String = "", _
        Optional ByVal keepHistory As Boolean = False) As String
    Dim p As String, turns As Collection, body As String, url As String, headers As String
    Dim reply As String, i As Long
    p = KitAIProvider()
    If Len(Trim$(message)) = 0 Then Err.Raise AI_ERR, , "There is no message to send."
    If AINeedsKey(p) And Len(KitAIKey(p)) = 0 Then
        Err.Raise AI_ERR, , "No API key for " & KitAIProviderName(p) & ". Set the environment variable KIT_AI_KEY_" & _
            UCase$(p) & " (or KIT_AI_API_KEY), or enter a key in the AI settings."
    End If
    If Len(KitAIEndpoint(p)) = 0 Then Err.Raise AI_ERR, , "No endpoint is set for " & KitAIProviderName(p) & "."
    If Len(KitAIModel(p)) = 0 Then Err.Raise AI_ERR, , "No model is set for " & KitAIProviderName(p) & "."
    If Len(systemPrompt) = 0 Then systemPrompt = KitAISystemPrompt()

    Set turns = New Collection
    If keepHistory Then
        If m_history Is Nothing Then Set m_history = New Collection
        For i = 1 To m_history.Count
            turns.Add m_history(i)
        Next
    End If
    turns.Add Array("user", message)
    Do While turns.Count > MAX_TURNS
        turns.Remove 1
    Loop

    body = AIRequestBody(p, KitAIModel(p), systemPrompt, turns)
    url = Replace(KitAIEndpoint(p), "{model}", KitAIModel(p))
    headers = AIHeaders(p, KitAIKey(p))
    reply = AIReplyText(p, AIPost(url, headers, body))
    If Len(reply) = 0 Then Err.Raise AI_ERR, , "The AI service sent no text back."

    If keepHistory Then
        turns.Add Array("assistant", reply)
        Do While turns.Count > MAX_TURNS
            turns.Remove 1
        Loop
        Set m_history = turns
    End If
    KitAIAsk = reply
End Function

' Forget the conversation KitAIAsk (keepHistory) has been keeping.
Public Sub KitAIReset()
    Set m_history = Nothing
End Sub

' Ask and write the reply into a text box or field (ctl). Returns True when it
' was written; False when the AI failed (the user is told) or the control is
' read only. Marks the form's record as changed through the control's own
' Value, so a bound form saves it like typing.
Public Function KitAIFill(ctl As Control, ByVal message As String, Optional ByVal systemPrompt As String = "") As Boolean
    Dim reply As String
    On Error GoTo Fail
    If ctl.Locked Or Not ctl.Enabled Then Exit Function
    DoCmd.Hourglass True
    reply = KitAIAsk(message, systemPrompt)
    DoCmd.Hourglass False
    ctl.Value = reply
    KitAIFill = True
    Exit Function
Fail:
    ShowAIError "modKitAI.KitAIFill", Err.Number, Err.Description
End Function

' Show an AI error to the user (a KitMsgBox), and log it unless it is one of
' ours (a missing key or a refused request is not a bug in the app).
Private Sub ShowAIError(ByVal source As String, ByVal num As Long, ByVal desc As String)
    On Error Resume Next
    DoCmd.Hourglass False
    If num <> AI_ERR Then LogError source
    KitMsgBox desc, vbExclamation, "AI service"
End Sub

' ---------------------------------------------------------------------------
' Requests and replies (pure text, no network)
' ---------------------------------------------------------------------------

' The JSON body for a service. turns is a Collection of Array(role, text),
' role "user" or "assistant", oldest first.
Public Function AIRequestBody(ByVal provider As String, ByVal model As String, ByVal systemPrompt As String, _
        turns As Collection) As String
    Dim s As String, i As Long, t As Variant, role As String
    Select Case AIFormat(provider)
        Case "openai"
            s = "{""model"":" & JsonQuote(model) & ",""messages"":["
            If Len(systemPrompt) > 0 Then s = s & "{""role"":""system"",""content"":" & JsonQuote(systemPrompt) & "},"
            For i = 1 To turns.Count
                t = turns(i)
                s = s & "{""role"":" & JsonQuote(CStr(t(0))) & ",""content"":" & JsonQuote(CStr(t(1))) & "}" & IIf(i < turns.Count, ",", "")
            Next
            s = s & "]}"
        Case "anthropic"
            s = "{""model"":" & JsonQuote(model) & ",""max_tokens"":" & MAX_TOKENS
            If Len(systemPrompt) > 0 Then s = s & ",""system"":" & JsonQuote(systemPrompt)
            s = s & ",""messages"":["
            For i = 1 To turns.Count
                t = turns(i)
                s = s & "{""role"":" & JsonQuote(CStr(t(0))) & ",""content"":" & JsonQuote(CStr(t(1))) & "}" & IIf(i < turns.Count, ",", "")
            Next
            s = s & "]}"
        Case "gemini"
            s = "{"
            If Len(systemPrompt) > 0 Then s = s & """systemInstruction"":{""parts"":[{""text"":" & JsonQuote(systemPrompt) & "}]},"
            s = s & """contents"":["
            For i = 1 To turns.Count
                t = turns(i)
                role = IIf(CStr(t(0)) = "assistant", "model", "user")
                s = s & "{""role"":""" & role & """,""parts"":[{""text"":" & JsonQuote(CStr(t(1))) & "}]}" & IIf(i < turns.Count, ",", "")
            Next
            s = s & "],""generationConfig"":{""maxOutputTokens"":" & MAX_TOKENS & "}}"
    End Select
    AIRequestBody = s
End Function

' The request headers for a service, one "Name: value" per line.
Public Function AIHeaders(ByVal provider As String, ByVal key As String) As String
    Dim s As String
    s = "Content-Type: application/json"
    Select Case LCase$(provider)
        Case "anthropic"
            s = s & vbLf & "x-api-key: " & key & vbLf & "anthropic-version: 2023-06-01"
        Case "gemini"
            s = s & vbLf & "x-goog-api-key: " & key
        Case "azure"
            s = s & vbLf & "api-key: " & key
        Case Else
            If Len(key) > 0 Then s = s & vbLf & "Authorization: Bearer " & key
    End Select
    AIHeaders = s
End Function

' The reply text in a service's response body. Raises AI_ERR with the
' service's own message when the body is an error, or has no text.
Public Function AIReplyText(ByVal provider As String, ByVal json As String) As String
    Dim root As Variant, errMsg As Variant, node As Variant, s As String, i As Long
    On Error GoTo BadJson
    JsonParseInto json, root
    On Error GoTo 0
    errMsg = JsonGet(root, "error.message")
    If Not IsEmpty(errMsg) Then
        If Len(CStr(errMsg)) > 0 Then Err.Raise AI_ERR, , CStr(errMsg)
    End If
    Select Case AIFormat(provider)
        Case "openai"
            s = JsonText(JsonGet(root, "choices.0.message.content"))
        Case "anthropic"
            JsonNodeInto JsonGet(root, "content"), node
            If IsObject(node) Then
                For i = 1 To node.Count
                    If JsonText(JsonGet(node(i), "type")) = "text" Then s = s & JsonText(JsonGet(node(i), "text"))
                Next
            End If
        Case "gemini"
            JsonNodeInto JsonGet(root, "candidates.0.content.parts"), node
            If IsObject(node) Then
                For i = 1 To node.Count
                    s = s & JsonText(JsonGet(node(i), "text"))
                Next
            End If
    End Select
    AIReplyText = Trim$(s)
    Exit Function
BadJson:
    Err.Raise AI_ERR, , "The AI service sent a reply this app cannot read."
End Function

' A JSON string literal for s: quoted, with control characters and anything
' outside plain ASCII escaped (\uXXXX), so the text reaches the service
' intact whatever encoding the request goes out in.
Public Function JsonQuote(ByVal s As String) As String
    Dim i As Long, c As Long, out As String
    out = """"
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        Select Case c
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 8: out = out & "\b"
            Case 9: out = out & "\t"
            Case 10: out = out & "\n"
            Case 12: out = out & "\f"
            Case 13: out = out & "\r"
            Case 32 To 126: out = out & Mid$(s, i, 1)
            Case Else: out = out & "\u" & Right$("0000" & Hex$(c), 4)
        End Select
    Next
    JsonQuote = out & """"
End Function

' ---------------------------------------------------------------------------
' Minimal JSON reader: objects become Scripting.Dictionary, arrays become
' Collection (1-based), strings, numbers, True/False and Null as VBA values.
' ---------------------------------------------------------------------------

' Parse json into out. Raises error 5 for text that is not JSON.
Public Sub JsonParseInto(ByVal json As String, ByRef out As Variant)
    Dim p As Long
    p = 1
    JsonValue json, p, out
    JsonSkip json, p
    If p <= Len(json) Then Err.Raise 5, , "Unexpected text after the JSON value at " & p
End Sub

' The value at path in a parsed node: names for objects, 0-based numbers for
' arrays ("choices.0.message.content"). Empty when any step is missing.
Public Function JsonGet(node As Variant, ByVal path As String) As Variant
    Dim seg As Variant, cur As Variant, idx As Long
    JsonNodeInto node, cur
    For Each seg In Split(path, ".")
        If Not IsObject(cur) Then Exit Function         ' ran out of nodes: Empty
        If TypeName(cur) = "Collection" Then
            If Not IsNumeric(seg) Then Exit Function
            idx = CLng(seg) + 1
            If idx < 1 Or idx > cur.Count Then Exit Function
            JsonNodeInto cur(idx), cur
        Else
            If Not cur.Exists(CStr(seg)) Then Exit Function
            JsonNodeInto cur(CStr(seg)), cur
        End If
    Next
    JsonNodeInto cur, JsonGet
End Function

' Copy a value, object or not, into out.
Private Sub JsonNodeInto(v As Variant, ByRef out As Variant)
    If IsObject(v) Then Set out = v Else out = v
End Sub

' A parsed value as text: "" for Empty, Null or a node.
Private Function JsonText(v As Variant) As String
    If IsObject(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    JsonText = CStr(v)
End Function

Private Sub JsonSkip(ByVal s As String, ByRef p As Long)
    Do While p <= Len(s)
        Select Case Mid$(s, p, 1)
            Case " ", vbTab, vbLf, vbCr: p = p + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

Private Sub JsonValue(ByVal s As String, ByRef p As Long, ByRef out As Variant)
    JsonSkip s, p
    If p > Len(s) Then Err.Raise 5, , "The JSON ends too soon"
    Select Case Mid$(s, p, 1)
        Case "{": JsonObject s, p, out
        Case "[": JsonArray s, p, out
        Case """": out = JsonString(s, p)
        Case Else: JsonLiteral s, p, out
    End Select
End Sub

Private Sub JsonObject(ByVal s As String, ByRef p As Long, ByRef out As Variant)
    Dim d As Object, k As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    p = p + 1
    JsonSkip s, p
    If Mid$(s, p, 1) = "}" Then
        p = p + 1
        Set out = d
        Exit Sub
    End If
    Do
        JsonSkip s, p
        If Mid$(s, p, 1) <> """" Then Err.Raise 5, , "A JSON name was expected at " & p
        k = JsonString(s, p)
        JsonSkip s, p
        If Mid$(s, p, 1) <> ":" Then Err.Raise 5, , "A JSON colon was expected at " & p
        p = p + 1
        v = Empty
        JsonValue s, p, v
        If IsObject(v) Then Set d(k) = v Else d(k) = v
        JsonSkip s, p
        Select Case Mid$(s, p, 1)
            Case ",": p = p + 1
            Case "}": p = p + 1: Exit Do
            Case Else: Err.Raise 5, , "A JSON comma or } was expected at " & p
        End Select
    Loop
    Set out = d
End Sub

Private Sub JsonArray(ByVal s As String, ByRef p As Long, ByRef out As Variant)
    Dim a As Collection, v As Variant
    Set a = New Collection
    p = p + 1
    JsonSkip s, p
    If Mid$(s, p, 1) = "]" Then
        p = p + 1
        Set out = a
        Exit Sub
    End If
    Do
        v = Empty
        JsonValue s, p, v
        a.Add v
        JsonSkip s, p
        Select Case Mid$(s, p, 1)
            Case ",": p = p + 1
            Case "]": p = p + 1: Exit Do
            Case Else: Err.Raise 5, , "A JSON comma or ] was expected at " & p
        End Select
    Loop
    Set out = a
End Sub

' The string starting at the quote at p; p ends after the closing quote.
Private Function JsonString(ByVal s As String, ByRef p As Long) As String
    Dim out As String, ch As String, n As Long, q As Long
    p = p + 1
    Do
        q = p
        Do While q <= Len(s)                            ' a run with nothing to unescape
            ch = Mid$(s, q, 1)
            If ch = """" Or ch = "\" Then Exit Do
            q = q + 1
        Loop
        If q > Len(s) Then Err.Raise 5, , "A JSON string is not closed"
        out = out & Mid$(s, p, q - p)
        p = q
        If ch = """" Then
            p = p + 1
            JsonString = out
            Exit Function
        End If
        ch = Mid$(s, p + 1, 1)                          ' the escape
        p = p + 2
        Select Case ch
            Case """", "\", "/": out = out & ch
            Case "b": out = out & vbBack
            Case "f": out = out & Chr$(12)
            Case "n": out = out & vbLf
            Case "r": out = out & vbCr
            Case "t": out = out & vbTab
            Case "u"
                If p + 3 > Len(s) Then Err.Raise 5, , "A JSON \u escape is cut short"
                n = CLng("&H" & Mid$(s, p, 4))
                If n > 32767 Then n = n - 65536         ' ChrW takes a signed 16-bit value
                out = out & ChrW(n)
                p = p + 4
            Case Else: Err.Raise 5, , "A JSON escape is not valid at " & (p - 1)
        End Select
    Loop
End Function

Private Sub JsonLiteral(ByVal s As String, ByRef p As Long, ByRef out As Variant)
    Dim q As Long, tok As String
    q = p
    Do While q <= Len(s)
        Select Case Mid$(s, q, 1)
            Case ",", "}", "]", " ", vbTab, vbLf, vbCr: Exit Do
        End Select
        q = q + 1
    Loop
    tok = Mid$(s, p, q - p)
    p = q
    Select Case tok
        Case "true": out = True
        Case "false": out = False
        Case "null": out = Null
        Case Else
            If Len(tok) = 0 Or Not IsNumeric(tok) Then Err.Raise 5, , "A JSON value was expected at " & q
            out = Val(tok)                              ' Val ignores the PC's decimal separator
    End Select
End Sub

' ---------------------------------------------------------------------------
' The network call
' ---------------------------------------------------------------------------

' POST body to url with headers ("Name: value" lines) and return the reply
' body, decoded as UTF-8. A status outside 2xx raises AI_ERR with the
' service's own message from the body when it has one.
Private Function AIPost(ByVal url As String, ByVal headers As String, ByVal body As String) As String
    Dim http As Object, line As Variant, at As Long, text As String, msg As String, root As Variant
    Dim secs As Long
    secs = Val(GetSetting(APP_NAME, SETTINGS_KEY, "TimeoutSeconds", CStr(DEFAULT_TIMEOUT_S)))
    If secs < 5 Then secs = DEFAULT_TIMEOUT_S
    On Error GoTo NoReach
    Set http = CreateObject("MSXML2.ServerXMLHTTP.60")
    http.setTimeouts 10000, 10000, 30000, secs * 1000   ' resolve, connect, send, receive
    http.Open "POST", url, False
    For Each line In Split(headers, vbLf)
        at = InStr(line, ":")
        If at > 1 Then http.setRequestHeader Trim$(Left$(line, at - 1)), Trim$(Mid$(line, at + 1))
    Next
    http.send body
    text = Utf8Text(http.responseBody)
    On Error GoTo 0
    If http.Status < 200 Or http.Status > 299 Then
        On Error Resume Next
        JsonParseInto text, root
        msg = JsonText(JsonGet(root, "error.message"))
        On Error GoTo 0
        If Len(msg) = 0 Then msg = "The AI service answered " & http.Status & " " & http.statusText & "."
        Err.Raise AI_ERR, , msg
    End If
    AIPost = text
    Exit Function
NoReach:
    Err.Raise AI_ERR, , "Could not reach the AI service (" & Err.Description & "). Check the network connection and the endpoint."
End Function

' Bytes (a Byte array from responseBody) as text, read as UTF-8.
Private Function Utf8Text(body As Variant) As String
    Dim st As Object
    If IsEmpty(body) Then Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1                                         ' adTypeBinary
    st.Open
    st.Write body
    st.Position = 0
    st.Type = 2                                         ' adTypeText
    st.Charset = "utf-8"
    Utf8Text = st.ReadText
    st.Close
End Function

' ---------------------------------------------------------------------------
' Interface parts the user chooses
' ---------------------------------------------------------------------------

' Register a place in the app that can use AI (see the top). Registering the
' same key again changes its caption. A key is letters and digits.
Public Sub KitAIPart(ByVal key As String, ByVal caption As String)
    Dim i As Long
    If Len(key) = 0 Or key Like "*[!A-Za-z0-9_]*" Then Err.Raise 5, , "An AI part key is letters, digits and _: " & key
    If m_parts Is Nothing Then Set m_parts = New Collection
    For i = 1 To m_parts.Count
        If m_parts(i)(0) = key Then m_parts.Remove i: Exit For
    Next
    m_parts.Add Array(key, caption)
End Sub

' The registered parts as "key;caption;" pairs for a value-list list box.
Public Function KitAIPartList() As String
    Dim i As Long, s As String
    If m_parts Is Nothing Then Exit Function
    For i = 1 To m_parts.Count
        s = s & m_parts(i)(0) & ";""" & Replace(m_parts(i)(1), """", "'") & """;"
    Next
    KitAIPartList = s
End Function

' The part keys the user turned on, comma separated.
Public Function KitAIEnabledParts() As String
    KitAIEnabledParts = GetSetting(APP_NAME, SETTINGS_KEY, "Parts", "")
End Function

' True when the user has turned on the part.
Public Function KitAIPartOn(ByVal key As String) As Boolean
    KitAIPartOn = (InStr("," & KitAIEnabledParts() & ",", "," & key & ",") > 0)
End Function

' Turn parts on: keys is the comma-separated list of the ones that are on
' (the rest go off). Keys that are not registered are dropped.
Public Sub KitAISetParts(ByVal keys As String)
    Dim item As Variant, k As String, kept As String, i As Long
    For Each item In Split(keys, ",")
        k = Trim$(CStr(item))
        If Len(k) > 0 And Not m_parts Is Nothing Then
            For i = 1 To m_parts.Count
                If m_parts(i)(0) = k Then kept = kept & IIf(Len(kept) > 0, ",", "") & k: Exit For
            Next
        End If
    Next
    SaveSetting APP_NAME, SETTINGS_KEY, "Parts", kept
End Sub

' Show each control on frm whose Tag is "ai:<key>" when the user has that
' part on and the service is usable, hide it otherwise. Call it from the
' form's Load (and again after KitAISettings closes).
Public Sub KitAIApplyParts(frm As Form)
    Dim c As Control, tg As String, on_ As Boolean
    For Each c In frm.Controls
        tg = Nz(c.Tag, "")
        If Left$(tg, Len(AI_TAG)) = AI_TAG Then
            on_ = KitAIPartOn(Mid$(tg, Len(AI_TAG) + 1)) And KitAIReady()
            SetVisible c, on_
        End If
    Next
End Sub

' ---------------------------------------------------------------------------
' Windows
' ---------------------------------------------------------------------------

' Open the AI settings window (service, model, endpoint, key for this session
' and the parts to use). Returns True when the user saved changes.
Public Function KitAISettings() As Boolean
    On Error GoTo Fail
    m_saved = False
    DimApp
    DoCmd.OpenForm "dlgAISettings", WindowMode:=acDialog
    UndimApp
    KitAISettings = m_saved
    Exit Function
Fail:
    LogError "modKitAI.KitAISettings", "dlgAISettings did not open"
    UndimApp
End Function

' dlgAISettings tells the module it saved.
Public Sub AISettingsSaved()
    m_saved = True
End Sub

' Open the chat window. seed is the text the message starts with; target is a
' text box that "Use reply" fills (without one the button is hidden). Returns
' the reply the user chose to use, "" when none.
Public Function KitAIChat(Optional ByVal seed As String = "", Optional target As Control = Nothing) As String
    If m_chatOpen Then Exit Function
    On Error GoTo Fail
    m_chatOpen = True
    m_chatSeed = seed
    Set m_chatTarget = target
    m_chatUsed = ""
    KitAIReset
    DimApp
    DoCmd.OpenForm "dlgAI", WindowMode:=acDialog
    UndimApp
    m_chatOpen = False
    KitAIChat = m_chatUsed
    m_chatUsed = ""
    Set m_chatTarget = Nothing
    Exit Function
Fail:
    LogError "modKitAI.KitAIChat", "dlgAI did not open"
    If m_chatOpen Then UndimApp
    m_chatOpen = False
    Set m_chatTarget = Nothing
End Function

' Open the chat seeded with an instruction and the text of a field on frm
' ("Fix the spelling", then the field's text); "Use reply" puts the reply in
' that field. For a button: =KitAIAssist([Form],"txtNotes","Fix the spelling")
Public Function KitAIAssist(frm As Form, ByVal controlName As String, Optional ByVal instruction As String = "") As String
    Dim c As Control, seed As String
    On Error GoTo Fail
    Set c = frm.Controls(controlName)
    seed = instruction
    If Len(Nz(c.Value, "")) > 0 Then seed = seed & IIf(Len(seed) > 0, vbCrLf & vbCrLf, "") & CStr(c.Value)
    If c.Locked Or Not c.Enabled Then
        KitAIAssist = KitAIChat(seed)
    Else
        KitAIAssist = KitAIChat(seed, c)
    End If
    Exit Function
Fail:
    ShowAIError "modKitAI.KitAIAssist", Err.Number, Err.Description
End Function

' What dlgAI needs to know.
Public Function KitAIChatSeed() As String
    KitAIChatSeed = m_chatSeed
End Function

Public Function KitAIChatHasTarget() As Boolean
    KitAIChatHasTarget = Not m_chatTarget Is Nothing
End Function

' dlgAI's "Use reply": put the reply in the target and remember it.
Public Sub KitAIChatUse(ByVal reply As String)
    If Not m_chatTarget Is Nothing Then
        On Error Resume Next
        m_chatTarget.Value = reply
        On Error GoTo 0
    End If
    m_chatUsed = reply
End Sub

' Build dlgAI, the chat window. Call it once with the app's dialogs.
Public Sub BuildAIDialog()
    Dim c As Control, w As Long, h As Long, pad As Long, y As Long, msgH As Long, replyH As Long
    w = px(640): h = px(600)
    pad = Sp(SP_DIALOG)
    msgH = px(110): replyH = px(250)
    NewDialog "AI assistant", w, h
    Ev F, "OnLoad"
    Set c = Lbl("lblAIIcon", ChrW(&HE99A), pad, pad, px(22), px(28), 13, False, CLR_PRIMARY)
    c.FontName = "Segoe MDL2 Assets": c.FontWeight = 400
    Lbl "lblAIHeader", "AI assistant", pad + px(30), pad, w - 2 * pad - px(30), px(28), 13, True
    y = pad + px(28) + Sp(SP_ROW_GAP)
    Lbl "lblAIMsg", "Your message", pad, y, w - 2 * pad, px(20), 9, True, CLR_MUTED_FG
    y = y + px(22)
    Set c = Txt("txtAIMsg", "", pad, y, w - 2 * pad, msgH, 10)
    c.EnterKeyBehavior = True: c.ScrollBars = 2
    y = y + msgH + Sp(SP_ROW_GAP)
    Set c = Btn("btnAISend", "Send", w - pad - px(88), y, px(88), px(30))
    Lbl "lblAIStatus", "", pad, y, w - 2 * pad - px(100), px(30), 9, False, CLR_MUTED_FG
    y = y + px(30) + Sp(SP_ROW_GAP)
    Lbl "lblAIReplyHead", "Reply", pad, y, w - 2 * pad, px(20), 9, True, CLR_MUTED_FG
    y = y + px(22)
    Set c = Txt("txtAIReply", "", pad, y, w - 2 * pad, replyH, 10)
    c.Locked = True: c.ScrollBars = 2
    c.OnGotFocus = "": c.OnLostFocus = ""            ' read only, so no focus ring
    y = h - pad - px(30)
    Btn "btnAISettings", "Settings", pad, y, px(88), px(30), "outline"
    Btn "btnAICopy", "Copy", pad + px(96), y, px(88), px(30), "outline"
    Btn "btnAIUse", "Use reply", w - pad - 2 * px(88) - px(8), y, px(88), px(30)
    Set c = Btn("btnAIClose", "Close", w - pad - px(88), y, px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgAI"
End Sub

' Build dlgAISettings. Call it once with the app's dialogs.
Public Sub BuildAISettingsDialog()
    Dim c As Control, w As Long, h As Long, pad As Long, y As Long, fieldW As Long, i As Long, p As Variant
    Dim list As String
    w = px(560): h = px(560)
    pad = Sp(SP_DIALOG)
    fieldW = w - 2 * pad
    NewDialog "AI settings", w, h
    Ev F, "OnLoad"
    Set c = Lbl("lblAISIcon", ChrW(&HE713), pad, pad, px(22), px(28), 13, False, CLR_PRIMARY)
    c.FontName = "Segoe MDL2 Assets": c.FontWeight = 400
    Lbl "lblAISHeader", "AI settings", pad + px(30), pad, w - 2 * pad - px(30), px(28), 13, True
    y = pad + px(28) + Sp(SP_ROW_GAP)
    For Each p In KitAIProviders()
        list = list & p & ";""" & KitAIProviderName(CStr(p)) & """;"
    Next
    Lbl "lblAISService", "Service", pad, y, fieldW, px(20), 9, True, CLR_MUTED_FG
    Set c = Cbo("cboAIService", "", list, pad, y + px(22), fieldW, px(30), 2, "0;" & (fieldW - px(30)))
    Ev c, "AfterUpdate"
    y = y + px(22) + px(30) + Sp(SP_ROW_GAP)
    Lbl "lblAIModel", "Model", pad, y, fieldW, px(20), 9, True, CLR_MUTED_FG
    Txt "txtAIModel", "", pad, y + px(22), fieldW, px(30), 10
    y = y + px(22) + px(30) + Sp(SP_ROW_GAP)
    Lbl "lblAIEndpoint", "Endpoint (web address)", pad, y, fieldW, px(20), 9, True, CLR_MUTED_FG
    Txt "txtAIEndpoint", "", pad, y + px(22), fieldW, px(30), 10
    y = y + px(22) + px(30) + Sp(SP_ROW_GAP)
    Lbl "lblAIKey", "API key (kept only until Access closes)", pad, y, fieldW, px(20), 9, True, CLR_MUTED_FG
    SecretBox "txtAIKey", "", pad, y + px(22), fieldW, px(30)
    y = y + px(22) + px(30) + Sp(SP_ROW_GAP)
    Lbl "lblAIKeyNote", "", pad, y, fieldW, px(18), 8, False, CLR_MUTED_FG
    y = y + px(18) + Sp(SP_ROW_GAP)
    Lbl "lblAIParts", "Parts of the app that use AI", pad, y, fieldW, px(20), 9, True, CLR_MUTED_FG
    Set c = Lst("lstAIParts", "", pad, y + px(22), fieldW, h - (y + px(22)) - pad - px(30) - Sp(16), 2, "0;" & (fieldW - px(30)))
    c.ColumnHeads = False: c.MultiSelect = 1: c.FontSize = 10
    y = h - pad - px(30)
    Set c = Btn("btnAISave", "Save", w - pad - 2 * px(88) - px(8), y, px(88), px(30))
    c.Default = True
    Set c = Btn("btnAICancel", "Cancel", w - pad - px(88), y, px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgAISettings"
End Sub
