Option Compare Database
Option Explicit

' Access UI kit - ask an AI service a question over HTTPS and use the answer.
'
' AIAsk sends a message (and the conversation so far, when asked) to the
' chat service the user has chosen and returns its answer as text:
'     answer = AIAsk("Summarize this note: " & note, "You write one-line summaries.")
' AIAskInto does the same between two controls on a form, so a button can
' read a text box and fill another. modKitAIUI adds the chat and settings
' dialogs.
'
' Services. The provider names the request and answer format, and the
' endpoint is the address it is sent to, so one module serves several services:
'   openai      OpenAI's chat completions format. Many other services speak it
'               too (Azure OpenAI, OpenRouter, LM Studio, vLLM, a company
'               gateway): choose openai and set their endpoint and model.
'   anthropic   Anthropic's Messages format.
'   ollama      A local Ollama server's chat format (no key needed).
' AIConfigure saves the choice for the Windows user (registry, as the palette
' is), so users can point an app at the service their organization allows.
'
' The API key is never in the code, the database or the registry. AIKey reads
' it, in this order, from the key the user typed this session (AISetKey, kept
' in memory only and cleared by AIClearKey) and from an environment variable
' (OPENAI_API_KEY, ANTHROPIC_API_KEY, or the name given to AIConfigure). The
' request goes only to https addresses, or to this computer, so a key is never
' sent in the clear.
'
' Interface parts. An app names the places AI can be used with
' AIRegisterPart, and tags the controls that belong to each with "ai:<key>"
' (a button's Tag "ai:notes"). The user decides, in the AI settings, whether AI
' is on at all and which parts use it (AIEnabled, AIPartEnabled). AIApplyParts
' hides the tagged controls of the parts that are off, and AIAskInto does
' nothing for a part that is off. AI is off until a user turns it on.
'
' Text sent to the service leaves this computer. Whether a note may be sent
' is the caller's decision: check it before calling AIAsk.
'
' Errors a user can fix (no key, a refused request, no answer) are raised
' with number AI_ERR_NUMBER and a plain message; entry points show it with
' HandleError or an inline message. The app provides APP_NAME, HandleError
' and RaiseAgain (modErrors).

Public Const AI_OPENAI As String = "openai"
Public Const AI_ANTHROPIC As String = "anthropic"
Public Const AI_OLLAMA As String = "ollama"
Public Const AI_ERR_NUMBER As Long = vbObjectError + 5100
Public Const AI_TIMEOUT_MS As Long = 90000        ' wait for an answer
Public Const AI_MAX_TOKENS As Long = 1024         ' answer length limit (required by anthropic)

Private Const SETTINGS_KEY As String = "AI"
Private Const ANTHROPIC_VERSION As String = "2023-06-01"
Private Const MAX_PARTS As Long = 12

Private m_key As String                   ' the key typed this session
Private m_role() As String                ' the conversation, oldest first
Private m_text() As String
Private m_n As Long
Private m_partKey(1 To MAX_PARTS) As String
Private m_partCap(1 To MAX_PARTS) As String
Private m_parts As Long

' Json reader state (JsonParse).
Private m_js As String
Private m_jp As Long

' ---------------------------------------------------------------- settings

' True for a provider this module can talk to.
Public Function AIProviderValid(ByVal provider As String) As Boolean
    Select Case LCase$(provider)
        Case AI_OPENAI, AI_ANTHROPIC, AI_OLLAMA: AIProviderValid = True
    End Select
End Function

' The chosen provider (openai until another is chosen).
Public Function AIProvider() As String
    Dim p As String
    p = LCase$(GetSetting(APP_NAME, SETTINGS_KEY, "Provider", AI_OPENAI))
    If AIProviderValid(p) Then AIProvider = p Else AIProvider = AI_OPENAI
End Function

Public Function AIDefaultEndpoint(ByVal provider As String) As String
    Select Case LCase$(provider)
        Case AI_ANTHROPIC: AIDefaultEndpoint = "https://api.anthropic.com/v1/messages"
        Case AI_OLLAMA: AIDefaultEndpoint = "http://localhost:11434/api/chat"
        Case Else: AIDefaultEndpoint = "https://api.openai.com/v1/chat/completions"
    End Select
End Function

' A starting model name for the provider. Services retire models, so the
' user can replace it (AIConfigure).
Public Function AIDefaultModel(ByVal provider As String) As String
    Select Case LCase$(provider)
        Case AI_ANTHROPIC: AIDefaultModel = "claude-sonnet-5-5"
        Case AI_OLLAMA: AIDefaultModel = "llama3.1"
        Case Else: AIDefaultModel = "gpt-4o-mini"
    End Select
End Function

' The environment variable that holds the provider's key.
Public Function AIDefaultKeyVar(ByVal provider As String) As String
    Select Case LCase$(provider)
        Case AI_ANTHROPIC: AIDefaultKeyVar = "ANTHROPIC_API_KEY"
        Case AI_OLLAMA: AIDefaultKeyVar = ""
        Case Else: AIDefaultKeyVar = "OPENAI_API_KEY"
    End Select
End Function

Public Function AIModel() As String
    Dim m As String
    m = Trim$(GetSetting(APP_NAME, SETTINGS_KEY, "Model", ""))
    If Len(m) = 0 Then m = AIDefaultModel(AIProvider())
    AIModel = m
End Function

Public Function AIEndpoint() As String
    Dim e As String
    e = Trim$(GetSetting(APP_NAME, SETTINGS_KEY, "Endpoint", ""))
    If Len(e) = 0 Then e = AIDefaultEndpoint(AIProvider())
    AIEndpoint = e
End Function

' The name of the environment variable that holds the key.
Public Function AIKeyVar() As String
    Dim v As String
    v = Trim$(GetSetting(APP_NAME, SETTINGS_KEY, "KeyVar", ""))
    If Len(v) = 0 Then v = AIDefaultKeyVar(AIProvider())
    AIKeyVar = v
End Function

' Save the service to use. Leave model, endpoint or keyVar blank for the
' provider's default. Changing the provider clears the saved model and
' endpoint unless new ones are given, so one provider's are never sent to
' another. keyVar is a variable's name, never a key.
Public Sub AIConfigure(ByVal provider As String, Optional ByVal model As String = "", _
        Optional ByVal endpoint As String = "", Optional ByVal keyVar As String = "")
    Dim problem As String
    provider = LCase$(Trim$(provider))
    If Not AIProviderValid(provider) Then
        Err.Raise AI_ERR_NUMBER, "modKitAI.AIConfigure", "Unknown AI provider """ & provider & """. Use openai, anthropic or ollama."
    End If
    problem = AIEndpointProblem(endpoint)
    If Len(problem) > 0 Then Err.Raise AI_ERR_NUMBER, "modKitAI.AIConfigure", problem
    If LooksLikeKey(keyVar) Then
        Err.Raise AI_ERR_NUMBER, "modKitAI.AIConfigure", "Give the name of an environment variable, not the key itself. The key is never saved."
    End If
    SaveSetting APP_NAME, SETTINGS_KEY, "Provider", provider
    SaveSetting APP_NAME, SETTINGS_KEY, "Model", Trim$(model)
    SaveSetting APP_NAME, SETTINGS_KEY, "Endpoint", Trim$(endpoint)
    SaveSetting APP_NAME, SETTINGS_KEY, "KeyVar", Trim$(keyVar)
End Sub

' "" when endpoint is blank (the default is used), https, or on this
' computer; otherwise what is wrong with it.
Public Function AIEndpointProblem(ByVal endpoint As String) As String
    endpoint = LCase$(Trim$(endpoint))
    If Len(endpoint) = 0 Then Exit Function
    If Left$(endpoint, 8) = "https://" Then Exit Function
    If IsLocalUrl(endpoint) Then Exit Function
    AIEndpointProblem = "The AI address must start with https:// (http:// is allowed only for this computer), so the key is never sent in the clear."
End Function

' True for an http address on this computer (localhost or 127.0.0.1).
Private Function IsLocalUrl(ByVal url As String) As Boolean
    Dim p As Variant
    For Each p In Array("http://localhost", "http://127.0.0.1")
        If Left$(url, Len(p)) = p Then
            If Len(url) = Len(p) Then
                IsLocalUrl = True
            ElseIf InStr(":/?", Mid$(url, Len(p) + 1, 1)) > 0 Then
                IsLocalUrl = True
            End If
        End If
    Next
End Function

' True when s is shaped like an API key rather than a variable name: it
' has a dash or is long. A variable name is letters, digits and underscores.
Private Function LooksLikeKey(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    LooksLikeKey = (InStr(s, "-") > 0) Or (Len(s) > 64) Or (InStr(s, " ") > 0)
End Function

' ------------------------------------------------------------------ the key

' Keep a key for this session only (memory; never written anywhere).
Public Sub AISetKey(ByVal key As String)
    m_key = Trim$(key)
End Sub

Public Sub AIClearKey()
    m_key = ""
End Sub

' The key to use: the one typed this session, else the environment variable.
Public Function AIKey() As String
    Dim v As String
    If Len(m_key) > 0 Then AIKey = m_key: Exit Function
    v = AIKeyVar()
    If Len(v) > 0 Then AIKey = Trim$(Environ$(v))
End Function

' True when a key is needed and none is available. A key is not needed for
' ollama, or for an http address on this computer.
Public Function AIKeyMissing() As Boolean
    If AIProvider() = AI_OLLAMA Then Exit Function
    If IsLocalUrl(LCase$(AIEndpoint())) Then Exit Function
    AIKeyMissing = (Len(AIKey()) = 0)
End Function

' ----------------------------------------------------------- the user's parts

' True when the user has turned AI on (off until they do).
Public Function AIEnabled() As Boolean
    AIEnabled = (GetSetting(APP_NAME, SETTINGS_KEY, "Enabled", "0") = "1")
End Function

Public Sub AISetEnabled(ByVal isOn As Boolean)
    SaveSetting APP_NAME, SETTINGS_KEY, "Enabled", IIf(isOn, "1", "0")
End Sub

' Name a place in the app where AI can be used, for the user's settings list.
' key is short and has no spaces ("notes"); caption is what the user reads
' ("Summarize notes"). Registering a key again changes its caption. Call at
' startup, before the settings dialog opens.
Public Sub AIRegisterPart(ByVal key As String, ByVal caption As String)
    Dim i As Long
    key = LCase$(Trim$(key))
    If Len(key) = 0 Or InStr(key, " ") > 0 Then
        Err.Raise AI_ERR_NUMBER, "modKitAI.AIRegisterPart", "An AI part key is a short word with no spaces."
    End If
    For i = 1 To m_parts
        If m_partKey(i) = key Then m_partCap(i) = caption: Exit Sub
    Next
    If m_parts >= MAX_PARTS Then
        Err.Raise AI_ERR_NUMBER, "modKitAI.AIRegisterPart", "At most " & MAX_PARTS & " AI parts can be registered."
    End If
    m_parts = m_parts + 1
    m_partKey(m_parts) = key
    m_partCap(m_parts) = caption
End Sub

Public Function AIPartCount() As Long
    AIPartCount = m_parts
End Function

Public Function AIPartKey(ByVal i As Long) As String
    If i >= 1 And i <= m_parts Then AIPartKey = m_partKey(i)
End Function

Public Function AIPartCaption(ByVal i As Long) As String
    If i >= 1 And i <= m_parts Then AIPartCaption = m_partCap(i)
End Function

' True when AI is on and the user has not turned this part off (parts are on
' until turned off).
Public Function AIPartEnabled(ByVal key As String) As Boolean
    If Not AIEnabled() Then Exit Function
    AIPartEnabled = (GetSetting(APP_NAME, SETTINGS_KEY, "Part_" & LCase$(key), "1") = "1")
End Function

Public Sub AISetPartEnabled(ByVal key As String, ByVal isOn As Boolean)
    SaveSetting APP_NAME, SETTINGS_KEY, "Part_" & LCase$(key), IIf(isOn, "1", "0")
End Sub

' Show the controls of frm tagged "ai:<key>" for the parts that are on and
' hide the rest. Call it from the form's Load, and after the settings change.
Public Sub AIApplyParts(frm As Form)
    Dim c As Control, t As String
    For Each c In frm.Controls
        t = ""
        On Error Resume Next
        t = c.Tag
        On Error GoTo 0
        If LCase$(Left$(t, 3)) = "ai:" Then SetVisible c, AIPartEnabled(Mid$(t, 4))
    Next
End Sub

' --------------------------------------------------------------- the chat

' Forget the conversation.
Public Sub AIReset()
    m_n = 0
    Erase m_role
    Erase m_text
End Sub

' Add a message to the conversation. role is "user" or "assistant".
Public Sub AIAddMessage(ByVal role As String, ByVal text As String)
    role = LCase$(role)
    If role <> "user" And role <> "assistant" Then
        Err.Raise AI_ERR_NUMBER, "modKitAI.AIAddMessage", "A message's role is user or assistant."
    End If
    m_n = m_n + 1
    ReDim Preserve m_role(1 To m_n)
    ReDim Preserve m_text(1 To m_n)
    m_role(m_n) = role
    m_text(m_n) = text
End Sub

Public Function AIMessageCount() As Long
    AIMessageCount = m_n
End Function

' Send prompt to the chosen service and return its answer. system sets how
' the AI should behave. keepHistory:=True continues the conversation (this
' prompt and the answer are added to it); otherwise each call starts fresh.
' Raises AI_ERR_NUMBER with a plain message when there is no key, the service
' refuses or can't be reached, or its answer can't be read.
Public Function AIAsk(ByVal prompt As String, Optional ByVal system As String = "", _
        Optional ByVal keepHistory As Boolean = False) As String
    Dim provider As String, url As String, body As String, raw As String
    Dim status As Long, reply As String, problem As String, n As Long
    If Len(Trim$(prompt)) = 0 Then
        Err.Raise AI_ERR_NUMBER, "modKitAI.AIAsk", "Type a message first."
    End If
    provider = AIProvider()
    url = AIEndpoint()
    problem = AIEndpointProblem(url)
    If Len(problem) > 0 Then Err.Raise AI_ERR_NUMBER, "modKitAI.AIAsk", problem
    If AIKeyMissing() Then
        Err.Raise AI_ERR_NUMBER, "modKitAI.AIAsk", "No AI key is available. Set the " & AIKeyVar() & _
            " environment variable and restart " & APP_NAME & ", or enter a key for this session in the AI settings."
    End If
    If Not keepHistory Then AIReset
    n = m_n
    AIAddMessage "user", prompt
    On Error GoTo Fail
    body = AIRequestBody(provider, AIModel(), system, AI_MAX_TOKENS)
    raw = AIPost(url, AIHeaders(provider, AIKey()), body, status)
    problem = ""
    If status >= 200 And status < 300 Then
        reply = AIParseResponse(provider, raw, problem)
    Else
        problem = AIFailureText(provider, raw, status)
    End If
    If Len(problem) > 0 Then Err.Raise AI_ERR_NUMBER, "modKitAI.AIAsk", problem
    AIAddMessage "assistant", reply
    AIAsk = reply
    Exit Function
Fail:
    ' Take this prompt back out, so a retry doesn't send it twice.
    m_n = n
    If Err.Number = AI_ERR_NUMBER Then Err.Raise Err.Number, Err.Source, Err.Description
    Err.Raise AI_ERR_NUMBER, "modKitAI.AIAsk", "The AI service could not be reached (" & Err.Description & ")."
End Function

' Read the text in the control inputName of frm, ask the AI, and put the
' answer in outputName (a text box or a label). Returns False without asking
' when the part named by partKey is off (see AIPartEnabled; leave partKey
' blank for a form that doesn't use parts) or the input is empty. Errors go
' to the caller, as AIAsk's do.
Public Function AIAskInto(frm As Form, ByVal inputName As String, ByVal outputName As String, _
        Optional ByVal partKey As String = "", Optional ByVal system As String = "", _
        Optional ByVal keepHistory As Boolean = False) As Boolean
    Dim q As String, a As String, c As Control
    If Len(partKey) > 0 Then
        If Not AIPartEnabled(partKey) Then Exit Function
    End If
    Set c = frm.Controls(inputName)
    If TypeOf c Is Access.Label Then q = c.Caption Else q = Nz(c.Value, "")
    If Len(Trim$(q)) = 0 Then Exit Function
    a = AIAsk(q, system, keepHistory)
    Set c = frm.Controls(outputName)
    If TypeOf c Is Access.Label Then c.Caption = a Else c.Value = a
    AIAskInto = True
End Function

' ------------------------------------------------- request and answer format

' The request body for the provider, with the conversation so far. system is
' left out when blank.
Public Function AIRequestBody(ByVal provider As String, ByVal model As String, ByVal system As String, _
        ByVal maxTokens As Long) As String
    Dim msgs As String, i As Long, s As String
    For i = 1 To m_n
        If Len(msgs) > 0 Then msgs = msgs & ","
        msgs = msgs & "{""role"":" & JsonString(m_role(i)) & ",""content"":" & JsonString(m_text(i)) & "}"
    Next
    Select Case LCase$(provider)
        Case AI_ANTHROPIC
            s = "{""model"":" & JsonString(model) & ",""max_tokens"":" & maxTokens
            If Len(system) > 0 Then s = s & ",""system"":" & JsonString(system)
            s = s & ",""messages"":[" & msgs & "]}"
        Case Else
            ' openai and ollama carry the system text as the first message.
            If Len(system) > 0 Then msgs = "{""role"":""system"",""content"":" & JsonString(system) & "}" & IIf(Len(msgs) > 0, ",", "") & msgs
            s = "{""model"":" & JsonString(model) & ",""messages"":[" & msgs & "]"
            If LCase$(provider) = AI_OLLAMA Then s = s & ",""stream"":false"
            s = s & "}"
    End Select
    AIRequestBody = s
End Function

' The request headers for the provider as "Name: value" lines. The key
' goes in a header only, never in the body or the address.
Public Function AIHeaders(ByVal provider As String, ByVal key As String) As Variant
    Dim h As Collection, out() As String, i As Long
    Set h = New Collection
    h.Add "Content-Type: application/json"
    Select Case LCase$(provider)
        Case AI_ANTHROPIC
            h.Add "anthropic-version: " & ANTHROPIC_VERSION
            If Len(key) > 0 Then h.Add "x-api-key: " & key
        Case Else
            If Len(key) > 0 Then h.Add "Authorization: Bearer " & key
    End Select
    ReDim out(0 To h.Count - 1)
    For i = 1 To h.Count
        out(i - 1) = h(i)
    Next
    AIHeaders = out
End Function

' The answer text in a successful reply. When the service sent an error, or
' the reply can't be read, returns "" and puts the reason in problem.
Public Function AIParseResponse(ByVal provider As String, ByVal json As String, ByRef problem As String) As String
    Dim root As Variant, blocks As Object, b As Variant, s As String, ans As String
    problem = ""
    On Error GoTo Unreadable
    JsonParse json, root
    If Not IsObject(root) Then GoTo Unreadable
    s = JsonErrorText(root)
    If Len(s) > 0 Then problem = "The AI service reported an error: " & s: Exit Function
    Select Case LCase$(provider)
        Case AI_ANTHROPIC
            ' content is a list of blocks; the answer is the text blocks joined.
            If Not root.Exists("content") Then GoTo Unreadable
            If TypeName(root("content")) <> "Collection" Then GoTo Unreadable
            Set blocks = root("content")
            For Each b In blocks
                If TypeName(b) = "Dictionary" Then
                    If b.Exists("text") Then s = s & CStr(b("text"))
                End If
            Next
            ans = s
        Case AI_OLLAMA
            ans = JsonText(root, "message/content")
        Case Else
            ans = JsonText(root, "choices/0/message/content")
    End Select
    AIParseResponse = ans
    If Len(ans) = 0 Then problem = "The AI service sent an empty answer."
    Exit Function
Unreadable:
    AIParseResponse = ""
    problem = "The AI service's answer could not be read."
End Function

' The message for a reply with an HTTP status outside 200-299: the service's
' own message when it sent one, else the status alone. Never includes the key.
Public Function AIFailureText(ByVal provider As String, ByVal raw As String, ByVal status As Long) As String
    Dim root As Variant, s As String, hint As String
    On Error Resume Next
    JsonParse raw, root
    If IsObject(root) Then s = JsonErrorText(root)
    On Error GoTo 0
    Select Case status
        Case 401, 403: hint = " Check the key."
        Case 404: hint = " Check the address and the model name."
        Case 429: hint = " Too many requests; wait a moment and try again."
    End Select
    AIFailureText = "The AI service answered HTTP " & status & IIf(Len(s) > 0, ": " & s, ".") & hint
End Function

' The message inside a service's error answer: {"error":{"message":"..."}},
' {"error":"..."}, or "" when the answer holds no error.
Private Function JsonErrorText(root As Variant) As String
    Dim e As String
    If Not root.Exists("error") Then Exit Function
    e = JsonText(root, "error/message")
    If Len(e) = 0 Then e = JsonText(root, "error")
    If Len(e) = 0 Then e = "unknown error"
    JsonErrorText = e
End Function

' -------------------------------------------------------------------- HTTP

' POST body to url with the headers ("Name: value" lines) and return the
' reply text; status gets the HTTP status. Waits up to AI_TIMEOUT_MS.
' The body is sent and the reply read as UTF-8.
Public Function AIPost(ByVal url As String, headers As Variant, ByVal body As String, ByRef status As Long) As String
    Dim http As Object, h As Variant, p As Long, payload As Variant
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 15000, 15000, 30000, AI_TIMEOUT_MS
    http.Open "POST", url, False
    For Each h In headers
        p = InStr(h, ": ")
        If p > 1 Then http.SetRequestHeader Left$(h, p - 1), Mid$(h, p + 2)
    Next
    payload = Utf8Bytes(body)
    http.Send payload
    status = http.status
    AIPost = Utf8Text(http.ResponseBody)
End Function

' s as UTF-8 bytes (no byte order mark).
Private Function Utf8Bytes(ByVal s As String) As Variant
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText s
    st.Position = 0
    st.Type = 1
    st.Position = 3                       ' past the byte order mark
    Utf8Bytes = st.Read
    st.Close
End Function

' UTF-8 bytes as text ("" when there are none).
Private Function Utf8Text(bytes As Variant) As String
    Dim st As Object
    If Not IsArray(bytes) Then Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write bytes
    st.Position = 0
    st.Type = 2
    st.Charset = "utf-8"
    Utf8Text = st.ReadText
    st.Close
End Function

' --------------------------------------------------------------------- JSON

' s as a JSON string literal, quotes included.
Public Function JsonString(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        Select Case code
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 8: out = out & "\b"
            Case 9: out = out & "\t"
            Case 10: out = out & "\n"
            Case 12: out = out & "\f"
            Case 13: out = out & "\r"
            Case 0 To 31: out = out & "\u" & Right$("0000" & Hex$(code), 4)
            Case Else: out = out & ch
        End Select
    Next
    JsonString = """" & out & """"
End Function

' Read JSON text into result: an object is a Scripting.Dictionary, an array
' a Collection, and other values are String, Double, Boolean or Null. Raises
' error 5 on text that isn't JSON.
Public Sub JsonParse(ByVal text As String, ByRef result As Variant)
    m_js = text
    m_jp = 1
    JsonValue result
    JsonSkip
    If m_jp <= Len(m_js) Then Err.Raise 5, "modKitAI.JsonParse", "Extra text after the JSON value"
    m_js = ""
End Sub

' The value at path in a parsed object, as text: names and array positions
' (from 0) separated by "/", for example "choices/0/message/content".
' "" when the path isn't there or ends at an object, an array or null.
Public Function JsonText(root As Variant, ByVal path As String) As String
    Dim cur As Variant, part As Variant, i As Long, parts() As String
    If Not IsObject(root) Then Exit Function
    Set cur = root
    parts = Split(path, "/")
    For i = 0 To UBound(parts)
        part = parts(i)
        Select Case TypeName(cur)
            Case "Dictionary"
                If Not cur.Exists(part) Then Exit Function
                If IsObject(cur(part)) Then
                    Set cur = cur(part)
                Else
                    If i < UBound(parts) Then Exit Function
                    If Not IsNull(cur(part)) Then JsonText = CStr(cur(part))
                    Exit Function
                End If
            Case "Collection"
                If Not IsNumeric(part) Then Exit Function
                If CLng(part) < 0 Or CLng(part) >= cur.Count Then Exit Function
                If IsObject(cur(CLng(part) + 1)) Then
                    Set cur = cur(CLng(part) + 1)
                Else
                    If i < UBound(parts) Then Exit Function
                    If Not IsNull(cur(CLng(part) + 1)) Then JsonText = CStr(cur(CLng(part) + 1))
                    Exit Function
                End If
            Case Else
                Exit Function
        End Select
    Next
End Function

Private Sub JsonSkip()
    Do While m_jp <= Len(m_js)
        Select Case Mid$(m_js, m_jp, 1)
            Case " ", vbTab, vbCr, vbLf: m_jp = m_jp + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

Private Sub JsonValue(ByRef out As Variant)
    Dim ch As String
    JsonSkip
    If m_jp > Len(m_js) Then Err.Raise 5, "modKitAI.JsonParse", "The JSON ends too soon"
    ch = Mid$(m_js, m_jp, 1)
    Select Case ch
        Case "{": Set out = JsonObject()
        Case "[": Set out = JsonArray()
        Case """": out = JsonStr()
        Case "t": JsonWord "true": out = True
        Case "f": JsonWord "false": out = False
        Case "n": JsonWord "null": out = Null
        Case Else: out = JsonNumber()
    End Select
End Sub

Private Sub JsonWord(ByVal w As String)
    If Mid$(m_js, m_jp, Len(w)) <> w Then Err.Raise 5, "modKitAI.JsonParse", "Unexpected text in the JSON"
    m_jp = m_jp + Len(w)
End Sub

Private Function JsonObject() As Object
    Dim d As Object, k As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1                       ' names match without regard to case
    m_jp = m_jp + 1
    JsonSkip
    If Mid$(m_js, m_jp, 1) = "}" Then m_jp = m_jp + 1: Set JsonObject = d: Exit Function
    Do
        JsonSkip
        If Mid$(m_js, m_jp, 1) <> """" Then Err.Raise 5, "modKitAI.JsonParse", "A name was expected in the JSON object"
        k = JsonStr()
        JsonSkip
        If Mid$(m_js, m_jp, 1) <> ":" Then Err.Raise 5, "modKitAI.JsonParse", "A colon was expected in the JSON object"
        m_jp = m_jp + 1
        JsonValue v
        If IsObject(v) Then Set d(k) = v Else d(k) = v
        JsonSkip
        Select Case Mid$(m_js, m_jp, 1)
            Case ",": m_jp = m_jp + 1
            Case "}": m_jp = m_jp + 1: Exit Do
            Case Else: Err.Raise 5, "modKitAI.JsonParse", "A comma or } was expected in the JSON object"
        End Select
    Loop
    Set JsonObject = d
End Function

Private Function JsonArray() As Collection
    Dim c As Collection, v As Variant
    Set c = New Collection
    m_jp = m_jp + 1
    JsonSkip
    If Mid$(m_js, m_jp, 1) = "]" Then m_jp = m_jp + 1: Set JsonArray = c: Exit Function
    Do
        JsonValue v
        c.Add v
        JsonSkip
        Select Case Mid$(m_js, m_jp, 1)
            Case ",": m_jp = m_jp + 1
            Case "]": m_jp = m_jp + 1: Exit Do
            Case Else: Err.Raise 5, "modKitAI.JsonParse", "A comma or ] was expected in the JSON array"
        End Select
    Loop
    Set JsonArray = c
End Function

' The string at the reader's position (on its opening quote), with escapes
' resolved.
Private Function JsonStr() As String
    Dim ch As String, out As String, hexa As String
    m_jp = m_jp + 1
    Do While m_jp <= Len(m_js)
        ch = Mid$(m_js, m_jp, 1)
        m_jp = m_jp + 1
        If ch = """" Then JsonStr = out: Exit Function
        If ch = "\" Then
            ch = Mid$(m_js, m_jp, 1)
            m_jp = m_jp + 1
            Select Case ch
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case "t": out = out & vbTab
                Case "b": out = out & Chr$(8)
                Case "f": out = out & Chr$(12)
                Case "u"
                    hexa = Mid$(m_js, m_jp, 4)
                    If Len(hexa) < 4 Or Not hexa Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then
                        Err.Raise 5, "modKitAI.JsonParse", "A bad \u escape in the JSON"
                    End If
                    out = out & ChrW(CLng("&H" & hexa))
                    m_jp = m_jp + 4
                Case Else: out = out & ch          ' \" \\ \/
            End Select
        Else
            out = out & ch
        End If
    Loop
    Err.Raise 5, "modKitAI.JsonParse", "A string in the JSON is not closed"
End Function

Private Function JsonNumber() As Double
    Dim start As Long, ch As String
    start = m_jp
    Do While m_jp <= Len(m_js)
        ch = Mid$(m_js, m_jp, 1)
        If InStr("0123456789+-.eE", ch) = 0 Then Exit Do
        m_jp = m_jp + 1
    Loop
    If m_jp = start Then Err.Raise 5, "modKitAI.JsonParse", "Unexpected text in the JSON"
    JsonNumber = Val(Mid$(m_js, start, m_jp - start))
End Function
