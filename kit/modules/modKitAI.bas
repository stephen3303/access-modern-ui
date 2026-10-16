Option Compare Database
Option Explicit

' Access UI kit - connect the app to an AI service over its web API.
'
'   answer = AiAsk("Summarise this note: " & Me.txtNote)
'
' Message input and response retrieval: AiAsk sends a message (and the
' conversation so far when keepHistory is True) and returns the reply text.
' AiTry does the same for a form: it returns False instead of raising, and
' AiLastError says why. AiLastStatus is the HTTP status of the last call.
'
' Compatible with different services: the same calls work for
'   "openai"      OpenAI chat completions
'   "anthropic"   Anthropic messages
'   "gemini"      Google Gemini generateContent
'   "compatible"  any OpenAI-style server: a local Ollama or LM Studio, or a
'                 company gateway (give its endpoint; a key is optional)
' Choose one with AiConfigure (or the dialog, ShowAI). The choice, model,
' endpoint and limits are saved per Windows user under APP_NAME (registry).
' Another service is a Select Case branch in AiDefault*, AiRequest and
' AiReply.
'
' API keys are never in code, the database or the registry. A key comes
' from the environment variable the service names (OPENAI_API_KEY,
' ANTHROPIC_API_KEY, GEMINI_API_KEY; AiConfigure keyVar:= names another) or
' from AiSetKey, which keeps it in memory for this session only (the dialog's
' "API key" button asks with a masked KitInputBox). A key is only sent over
' https, or to this computer (localhost).
'
' Integration with other systems:
'   AiAskWithData prompt, sql      sends the rows of a query with the message
'   AiOnResponse partKey, prompt, text   an optional Public Sub in the app's
'                                  modules, called after every AiRunPart so
'                                  the app can log, store or forward replies
' Everything sent leaves the computer: don't send data the service may not
' hold (see ClassificationPermits in modKitClassification).
'
' Interface parts the user connects: the app declares where AI could be used,
' and each user turns those parts on or off (ShowAI lists them).
'   AiBind "note", "Summarise note", "frmNote", "txtNote", "txtSummary", _
'       "Summarise this note in two sentences: {input}", "btnAiSummary"
'   Form_Load:          ApplyAiParts Me     (shows the button only when on)
'   btnAiSummary_Click: AiRunPart "note", Me   (reads txtNote, fills txtSummary)
' A part is off until the user ticks it in ShowAI (or AiSetPartEnabled).
'
' Needs from the app: APP_NAME, HandleError, RaiseAgain, LogError, and the
' dialog dlgAI: call BuildAiDialog once with the app's dialogs and copy
' kit\forms\dlgAI.vba to src\forms. Requires MSXML (Windows).

Private Const SETTINGS_KEY As String = "AI"

Private Const DEFAULT_TIMEOUT As Long = 60           ' seconds to wait for a reply
Private Const ANTHROPIC_VERSION As String = "2023-06-01"
Private Const MAX_DATA_CHARS As Long = 20000         ' cap on query text AiAskWithData sends

Private m_keys As Object                             ' provider -> key, this session only
Private m_hist As Collection                         ' Array(role, text) pairs of the chat
Private m_parts As Object                            ' key -> Array(caption, form, input, output, template, trigger)
Private m_lastError As String
Private m_lastStatus As Long
Private m_js As String                               ' JSON being parsed
Private m_jp As Long                                 ' and the position in it

' ---------------------------------------------------------------------------
' Services and settings
' ---------------------------------------------------------------------------

' The services as "key|Name", the choices for AiConfigure and the dialog.
Public Function AiProviders() As Variant
    AiProviders = Array("openai|OpenAI", "anthropic|Anthropic", "gemini|Google Gemini", "compatible|OpenAI-compatible (local or other)")
End Function

Private Function ValidProvider(ByVal p As String) As Boolean
    Dim v As Variant
    For Each v In AiProviders()
        If Split(v, "|")(0) = p Then ValidProvider = True
    Next
End Function

' The service's own endpoint. {model} is replaced by the model name.
Public Function AiDefaultEndpoint(ByVal provider As String) As String
    Select Case provider
        Case "openai": AiDefaultEndpoint = "https://api.openai.com/v1/chat/completions"
        Case "anthropic": AiDefaultEndpoint = "https://api.anthropic.com/v1/messages"
        Case "gemini": AiDefaultEndpoint = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
        Case "compatible": AiDefaultEndpoint = "http://localhost:11434/v1/chat/completions"   ' Ollama
    End Select
End Function

' A model to start with; models are retired often, so set the one you use.
Public Function AiDefaultModel(ByVal provider As String) As String
    Select Case provider
        Case "openai": AiDefaultModel = "gpt-4o-mini"
        Case "anthropic": AiDefaultModel = "claude-sonnet-5-5"
        Case "gemini": AiDefaultModel = "gemini-2.0-flash"
        Case "compatible": AiDefaultModel = "llama3.2"
    End Select
End Function

' The environment variable that holds the service's API key.
Public Function AiDefaultKeyVar(ByVal provider As String) As String
    Select Case provider
        Case "openai": AiDefaultKeyVar = "OPENAI_API_KEY"
        Case "anthropic": AiDefaultKeyVar = "ANTHROPIC_API_KEY"
        Case "gemini": AiDefaultKeyVar = "GEMINI_API_KEY"
        Case "compatible": AiDefaultKeyVar = "AI_API_KEY"
    End Select
End Function

' The service in use: the saved one, else OpenAI.
Public Function AiProvider() As String
    AiProvider = GetSetting(APP_NAME, SETTINGS_KEY, "Provider", "")
    If Not ValidProvider(AiProvider) Then AiProvider = "openai"
End Function

Public Function AiModel() As String
    AiModel = GetSetting(APP_NAME, SETTINGS_KEY, "Model", "")
    If Len(AiModel) = 0 Then AiModel = AiDefaultModel(AiProvider())
End Function

Public Function AiEndpoint() As String
    AiEndpoint = GetSetting(APP_NAME, SETTINGS_KEY, "Endpoint", "")
    If Len(AiEndpoint) = 0 Then AiEndpoint = AiDefaultEndpoint(AiProvider())
End Function

Public Function AiKeyVar() As String
    AiKeyVar = GetSetting(APP_NAME, SETTINGS_KEY, "KeyVar", "")
    If Len(AiKeyVar) = 0 Then AiKeyVar = AiDefaultKeyVar(AiProvider())
End Function

' Longest reply asked for, in tokens; 0 leaves it to the service (Anthropic
' needs one, so 1024 is sent for it then).
Public Function AiMaxTokens() As Long
    AiMaxTokens = Val(GetSetting(APP_NAME, SETTINGS_KEY, "MaxTokens", "0"))
End Function

Public Function AiTimeout() As Long
    AiTimeout = Val(GetSetting(APP_NAME, SETTINGS_KEY, "Timeout", ""))
    If AiTimeout < 5 Then AiTimeout = DEFAULT_TIMEOUT
End Function

' Standing instruction sent with every message ("" for none).
Public Function AiSystemPrompt() As String
    AiSystemPrompt = GetSetting(APP_NAME, SETTINGS_KEY, "System", "")
End Function

' Save the settings given (leave the others as they are). Changing the service
' forgets the model, endpoint and key variable chosen for the old one unless
' they are given too. Raises error 5 for an unknown service or a bad number.
'   AiConfigure "anthropic", model:="claude-sonnet-5-5"
'   AiConfigure "compatible", endpoint:="http://localhost:1234/v1/chat/completions", model:="qwen2.5"
Public Sub AiConfigure(Optional ByVal provider As String = "", Optional ByVal model As String = "", _
        Optional ByVal endpoint As String = "", Optional ByVal keyVar As String = "", _
        Optional ByVal maxTokens As Long = -1, Optional ByVal timeoutSec As Long = -1, _
        Optional ByVal system As Variant)
    If Len(provider) > 0 Then
        If Not ValidProvider(provider) Then Err.Raise 5, , "Unknown AI service: " & provider
        If provider <> AiProvider() Then
            DelSetting "Model": DelSetting "Endpoint": DelSetting "KeyVar"
        End If
        SaveSetting APP_NAME, SETTINGS_KEY, "Provider", provider
    End If
    If Len(endpoint) > 0 Then
        If Not EndpointAllowed(endpoint) Then Err.Raise 5, , "The endpoint must start with https:// (http:// only for this computer): " & endpoint
        SaveSetting APP_NAME, SETTINGS_KEY, "Endpoint", endpoint
    End If
    If Len(model) > 0 Then SaveSetting APP_NAME, SETTINGS_KEY, "Model", model
    If Len(keyVar) > 0 Then SaveSetting APP_NAME, SETTINGS_KEY, "KeyVar", keyVar
    If maxTokens > -1 Then
        If maxTokens > 200000 Then Err.Raise 5, , "Max tokens must be 0 to 200000: " & maxTokens
        SaveSetting APP_NAME, SETTINGS_KEY, "MaxTokens", CStr(maxTokens)
    End If
    If timeoutSec > -1 Then
        If timeoutSec < 5 Or timeoutSec > 600 Then Err.Raise 5, , "Timeout must be 5 to 600 seconds: " & timeoutSec
        SaveSetting APP_NAME, SETTINGS_KEY, "Timeout", CStr(timeoutSec)
    End If
    If Not IsMissing(system) Then SaveSetting APP_NAME, SETTINGS_KEY, "System", CStr(system)
End Sub

' Go back to the service's own endpoint.
Public Sub AiUseDefaultEndpoint()
    DelSetting "Endpoint"
End Sub

' Forget one saved setting so its default applies again.
Private Sub DelSetting(ByVal name As String)
    On Error Resume Next
    DeleteSetting APP_NAME, SETTINGS_KEY, name
End Sub

' True for https, or http to this computer (a local model server).
Private Function EndpointAllowed(ByVal url As String) As Boolean
    Dim u As String
    u = LCase$(Trim$(url))
    If Left$(u, 8) = "https://" Then
        EndpointAllowed = True
    ElseIf Left$(u, 7) = "http://" Then
        u = Mid$(u, 8)
        EndpointAllowed = (Left$(u, 10) = "localhost:" Or Left$(u, 10) = "localhost/" Or u = "localhost" _
            Or Left$(u, 10) = "127.0.0.1:" Or Left$(u, 10) = "127.0.0.1/" Or u = "127.0.0.1" _
            Or Left$(u, 6) = "[::1]:" Or Left$(u, 6) = "[::1]/")
    End If
End Function

' ---------------------------------------------------------------------------
' API key (never stored)
' ---------------------------------------------------------------------------

' Keep key for the service in use until Access closes. "" forgets it.
Public Sub AiSetKey(ByVal key As String)
    If m_keys Is Nothing Then Set m_keys = CreateObject("Scripting.Dictionary")
    If Len(key) = 0 Then
        If m_keys.Exists(AiProvider()) Then m_keys.Remove AiProvider()
    Else
        m_keys(AiProvider()) = Trim$(key)
    End If
End Sub

' The key for the service in use: the one AiSetKey kept, else the service's
' environment variable. "" when there is none.
Public Function AiKey() As String
    If Not m_keys Is Nothing Then
        If m_keys.Exists(AiProvider()) Then AiKey = m_keys(AiProvider()): Exit Function
    End If
    AiKey = Trim$(Environ$(AiKeyVar()))
End Function

' True when AiAsk has what it needs: a key, or a service that takes none.
Public Function AiReady() As Boolean
    AiReady = (Len(AiKey()) > 0 Or AiProvider() = "compatible")
End Function

' ---------------------------------------------------------------------------
' Asking
' ---------------------------------------------------------------------------

' Send message and return the reply text.
'   system    standing instruction for this call ("" uses AiSystemPrompt())
'   keepHistory  True remembers the exchange so the next call continues the
'             conversation (AiNewChat starts over); False asks on its own
' Raises an error that says what went wrong: no key, no connection, a refusal
' from the service (its own message), or a reply with no text.
Public Function AiAsk(ByVal message As String, Optional ByVal system As String = "", Optional ByVal keepHistory As Boolean = False) As String
    Dim provider As String, key As String, url As String, body As String, resp As String
    Dim msgs As Collection, answer As String
    On Error GoTo Fail
    If Len(Trim$(message)) = 0 Then Err.Raise 5, , "There is no message to send."
    provider = AiProvider()
    key = AiKey()
    If Len(key) = 0 And provider <> "compatible" Then
        Err.Raise vbObjectError + 5101, , "No API key. Set the " & AiKeyVar() & " environment variable and restart Access, or choose API key in the AI window."
    End If
    url = Replace(AiEndpoint(), "{model}", AiModel())
    If Not EndpointAllowed(url) Then Err.Raise 5, , "The AI endpoint must start with https:// : " & url
    If Len(system) = 0 Then system = AiSystemPrompt()
    Set msgs = New Collection
    If Not m_hist Is Nothing Then
        Dim h As Variant
        For Each h In m_hist: msgs.Add h: Next
    End If
    msgs.Add Array("user", message)
    body = AiRequest(provider, AiModel(), system, msgs, AiMaxTokens())
    resp = Post(provider, url, key, body)
    answer = AiReply(provider, resp)
    If keepHistory Then
        If m_hist Is Nothing Then Set m_hist = New Collection
        m_hist.Add Array("user", message)
        m_hist.Add Array("assistant", answer)
    End If
    m_lastError = ""
    AiAsk = answer
    Exit Function
Fail:
    m_lastError = Err.Description
    RaiseAgain "modKitAI.AiAsk", provider
End Function

' Like AiAsk for code that shows the failure itself: True with the reply in
' answer, or False with AiLastError set (answer is "").
Public Function AiTry(ByVal message As String, ByRef answer As String, Optional ByVal system As String = "", _
        Optional ByVal keepHistory As Boolean = False) As Boolean
    On Error GoTo Fail
    answer = AiAsk(message, system, keepHistory)
    AiTry = True
    Exit Function
Fail:
    answer = ""
    m_lastError = Err.Description
End Function

' The message and the rows of a query in one request, for questions about the
' app's data ("which of these is overdue?"). The rows go after the message as
' tab-separated text with a header row, at most maxRows rows and about 20,000
' characters; a longer result is cut and says so.
Public Function AiAskWithData(ByVal message As String, ByVal sql As String, Optional ByVal maxRows As Long = 50, _
        Optional ByVal system As String = "", Optional ByVal keepHistory As Boolean = False) As String
    AiAskWithData = AiAsk(message & vbLf & vbLf & "Data:" & vbLf & AiDataText(sql, maxRows), system, keepHistory)
End Function

' The rows of a query as tab-separated text with a header row (tabs and line
' breaks inside a value become spaces).
Public Function AiDataText(ByVal sql As String, Optional ByVal maxRows As Long = 50) As String
    Dim rs As DAO.Recordset, i As Long, n As Long, line As String, sb As String, cut As Boolean
    On Error GoTo Fail
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    For i = 0 To rs.Fields.Count - 1
        line = line & IIf(i > 0, vbTab, "") & rs.Fields(i).Name
    Next
    sb = line
    Do While Not rs.EOF
        If n >= maxRows Or Len(sb) > MAX_DATA_CHARS Then cut = True: Exit Do
        line = ""
        For i = 0 To rs.Fields.Count - 1
            line = line & IIf(i > 0, vbTab, "") & Flat(rs.Fields(i).Value)
        Next
        sb = sb & vbLf & line
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    If cut Then sb = sb & vbLf & "(more rows not shown)"
    AiDataText = sb
    Exit Function
Fail:
    RaiseAgain "modKitAI.AiDataText", sql
End Function

Private Function Flat(ByVal v As Variant) As String
    If IsNull(v) Then Exit Function
    If VarType(v) = vbArray + vbByte Then Flat = "(binary)": Exit Function
    Flat = Replace(Replace(Replace(Replace(CStr(v), vbTab, " "), vbCrLf, " "), vbCr, " "), vbLf, " ")
End Function

' Why the last AiAsk or AiTry failed; "" after a success.
Public Function AiLastError() As String
    AiLastError = m_lastError
End Function

' HTTP status of the last call to the service (0: it never answered).
Public Function AiLastStatus() As Long
    AiLastStatus = m_lastStatus
End Function

' Start the conversation over: the next keepHistory call remembers nothing before it.
Public Sub AiNewChat()
    Set m_hist = Nothing
End Sub

' Messages remembered in the conversation (a question and its reply are two).
Public Function AiHistoryCount() As Long
    If Not m_hist Is Nothing Then AiHistoryCount = m_hist.Count
End Function

' ---------------------------------------------------------------------------
' One request per service
' ---------------------------------------------------------------------------

' The JSON body for the service. msgs holds Array(role, text) pairs, role
' "user" or "assistant". Public so a test or another service's code can see it.
Public Function AiRequest(ByVal provider As String, ByVal model As String, ByVal system As String, _
        msgs As Collection, ByVal maxTokens As Long) As String
    Dim m As Variant, s As String, role As String
    Select Case provider
        Case "anthropic"
            If maxTokens <= 0 Then maxTokens = 1024
            For Each m In msgs
                s = s & IIf(Len(s) > 0, ",", "") & "{""role"":""" & m(0) & """,""content"":" & JsonStr(CStr(m(1))) & "}"
            Next
            AiRequest = "{""model"":" & JsonStr(model) & ",""max_tokens"":" & maxTokens & _
                IIf(Len(system) > 0, ",""system"":" & JsonStr(system), "") & ",""messages"":[" & s & "]}"
        Case "gemini"
            For Each m In msgs
                role = IIf(m(0) = "assistant", "model", "user")
                s = s & IIf(Len(s) > 0, ",", "") & "{""role"":""" & role & """,""parts"":[{""text"":" & JsonStr(CStr(m(1))) & "}]}"
            Next
            AiRequest = "{""contents"":[" & s & "]" & _
                IIf(Len(system) > 0, ",""systemInstruction"":{""parts"":[{""text"":" & JsonStr(system) & "}]}", "") & _
                IIf(maxTokens > 0, ",""generationConfig"":{""maxOutputTokens"":" & maxTokens & "}", "") & "}"
        Case "openai", "compatible"
            If Len(system) > 0 Then s = "{""role"":""system"",""content"":" & JsonStr(system) & "}"
            For Each m In msgs
                s = s & IIf(Len(s) > 0, ",", "") & "{""role"":""" & m(0) & """,""content"":" & JsonStr(CStr(m(1))) & "}"
            Next
            AiRequest = "{""model"":" & JsonStr(model) & ",""messages"":[" & s & "]" & _
                IIf(maxTokens > 0, ",""max_tokens"":" & maxTokens, "") & "}"
        Case Else
            Err.Raise 5, , "Unknown AI service: " & provider
    End Select
End Function

' The reply text in the service's JSON answer. Raises the service's own error
' message when the answer is an error, and an error when it holds no text.
Public Function AiReply(ByVal provider As String, ByVal json As String) As String
    Dim root As Variant, v As Variant, part As Variant, s As String
    ParseInto root, json
    If TypeName(root) <> "Dictionary" Then Err.Raise vbObjectError + 5102, , "The AI service's answer was not understood: " & Left$(json, 200)
    If root.Exists("error") Then Err.Raise vbObjectError + 5103, , "The AI service reported: " & ErrorText(root("error"))
    Select Case provider
        Case "anthropic"
            If root.Exists("content") Then
                For Each part In root("content")
                    If TypeName(part) = "Dictionary" Then
                        If part.Exists("text") Then s = s & part("text")
                    End If
                Next
            End If
        Case "gemini"
            If root.Exists("candidates") Then
                If root("candidates").Count > 0 Then
                    Set v = root("candidates")(1)
                    If v.Exists("content") Then
                        If v("content").Exists("parts") Then
                            For Each part In v("content")("parts")
                                If part.Exists("text") Then s = s & part("text")
                            Next
                        End If
                    End If
                End If
            End If
        Case Else
            If root.Exists("choices") Then
                If root("choices").Count > 0 Then
                    Set v = root("choices")(1)
                    If v.Exists("message") Then
                        If v("message").Exists("content") Then
                            If Not IsNull(v("message")("content")) Then s = v("message")("content")
                        End If
                    End If
                End If
            End If
    End Select
    If Len(s) = 0 Then Err.Raise vbObjectError + 5104, , "The AI service sent no text." & IIf(provider = "gemini", " (It may have blocked the message.)", "")
    AiReply = s
End Function

' The message inside an error object (or the error itself when it is text).
Private Function ErrorText(ByVal e As Variant) As String
    If TypeName(e) = "Dictionary" Then
        If e.Exists("message") Then ErrorText = CStr(e("message")): Exit Function
    End If
    If IsObject(e) Then ErrorText = "unknown error" Else ErrorText = CStr(e)
End Function

' ---------------------------------------------------------------------------
' HTTP
' ---------------------------------------------------------------------------

' POST body to url with the service's headers and return the answer text.
' A refusal (HTTP 400 and up) raises the service's message when it sent one.
Private Function Post(ByVal provider As String, ByVal url As String, ByVal key As String, ByVal body As String) As String
    Dim http As Object, t As Long, text As String, root As Variant
    On Error GoTo Fail
    m_lastStatus = 0
    Set http = CreateObject("MSXML2.ServerXMLHTTP.60")
    t = AiTimeout() * 1000
    http.setTimeouts 10000, 15000, 30000, t
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    Select Case provider
        Case "anthropic"
            http.setRequestHeader "x-api-key", key
            http.setRequestHeader "anthropic-version", ANTHROPIC_VERSION
        Case "gemini"
            http.setRequestHeader "x-goog-api-key", key
        Case Else
            If Len(key) > 0 Then http.setRequestHeader "Authorization", "Bearer " & key
    End Select
    http.send body
    m_lastStatus = http.Status
    text = Utf8Text(http.responseBody)
    If m_lastStatus >= 400 Then
        On Error Resume Next
        ParseInto root, text
        If TypeName(root) = "Dictionary" Then
            If root.Exists("error") Then text = ErrorText(root("error"))
        End If
        On Error GoTo Fail
        Err.Raise vbObjectError + 5105, , "The AI service refused the request (HTTP " & m_lastStatus & "): " & Left$(text, 300)
    End If
    Post = text
    Exit Function
Fail:
    If Err.Number <> vbObjectError + 5105 And m_lastStatus = 0 Then
        Err.Raise vbObjectError + 5106, , "Could not reach the AI service (" & Err.Description & "). Check the connection, the endpoint and the timeout."
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Bytes of a UTF-8 answer as text.
Private Function Utf8Text(ByVal bytes As Variant) As String
    Dim st As Object
    If IsEmpty(bytes) Then Exit Function
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

' ---------------------------------------------------------------------------
' JSON
' ---------------------------------------------------------------------------

' s as a JSON string literal. Everything outside plain ASCII is written as
' \uXXXX, so the request is ASCII whatever the text holds.
Public Function JsonStr(ByVal s As String) As String
    Dim i As Long, c As Long, out As String
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        Select Case c
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 8: out = out & "\b"
            Case 9: out = out & "\t"
            Case 10: out = out & "\n"
            Case 12: out = out & "\f"
            Case 13: out = out & "\r"
            Case Is < 32, Is > 126: out = out & "\u" & Right$("0000" & Hex$(c), 4)
            Case Else: out = out & Chr$(c)
        End Select
    Next
    JsonStr = """" & out & """"
End Function

' Parse JSON text: an object is a Scripting.Dictionary, an array a Collection
' (1-based), a string a String, a number a Double, true/false a Boolean and
' null Null. Raises error 5 for text that is not JSON.
' Use Set for an object or an array: Set d = JsonParse(text). ParseInto
' (below) fills a Variant whatever the JSON holds.
Public Function JsonParse(ByVal text As String) As Variant
    Dim v As Variant
    ParseInto v, text
    If IsObject(v) Then Set JsonParse = v Else JsonParse = v
End Function

' Parse text into out, an object or a plain value alike.
Private Sub ParseInto(ByRef out As Variant, ByVal text As String)
    m_js = text: m_jp = 1
    JsonValue out
    JsonSkip
    If m_jp <= Len(m_js) Then JsonFail "extra text after the value"
    m_js = ""
End Sub

Private Sub JsonFail(ByVal why As String)
    Dim at As Long
    at = m_jp
    m_js = ""
    Err.Raise 5, , "Not valid JSON (" & why & " at character " & at & ")."
End Sub

Private Sub JsonSkip()
    Do While m_jp <= Len(m_js)
        Select Case Mid$(m_js, m_jp, 1)
            Case " ", vbTab, vbLf, vbCr: m_jp = m_jp + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

Private Sub JsonValue(ByRef out As Variant)
    Dim ch As String, st As Long
    JsonSkip
    If m_jp > Len(m_js) Then JsonFail "the text ended early"
    ch = Mid$(m_js, m_jp, 1)
    Select Case ch
        Case "{": JsonObject out
        Case "[": JsonArray out
        Case """": out = JsonString()
        Case "t": JsonWord "true": out = True
        Case "f": JsonWord "false": out = False
        Case "n": JsonWord "null": out = Null
        Case "-", "0" To "9"
            st = m_jp
            Do While m_jp <= Len(m_js)
                If InStr("+-0123456789.eE", Mid$(m_js, m_jp, 1)) = 0 Then Exit Do
                m_jp = m_jp + 1
            Loop
            out = Val(Mid$(m_js, st, m_jp - st))
        Case Else: JsonFail "unexpected '" & ch & "'"
    End Select
End Sub

Private Sub JsonWord(ByVal w As String)
    If Mid$(m_js, m_jp, Len(w)) <> w Then JsonFail "expected " & w
    m_jp = m_jp + Len(w)
End Sub

Private Sub JsonObject(ByRef out As Variant)
    Dim d As Object, k As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    m_jp = m_jp + 1
    JsonSkip
    If Mid$(m_js, m_jp, 1) = "}" Then m_jp = m_jp + 1: Set out = d: Exit Sub
    Do
        JsonSkip
        If Mid$(m_js, m_jp, 1) <> """" Then JsonFail "expected a name"
        k = JsonString()
        JsonSkip
        If Mid$(m_js, m_jp, 1) <> ":" Then JsonFail "expected ':'"
        m_jp = m_jp + 1
        JsonValue v
        If IsObject(v) Then Set d(k) = v Else d(k) = v
        JsonSkip
        Select Case Mid$(m_js, m_jp, 1)
            Case ",": m_jp = m_jp + 1
            Case "}": m_jp = m_jp + 1: Exit Do
            Case Else: JsonFail "expected ',' or '}'"
        End Select
    Loop
    Set out = d
End Sub

Private Sub JsonArray(ByRef out As Variant)
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
            Case "]": m_jp = m_jp + 1: Exit Do
            Case Else: JsonFail "expected ',' or ']'"
        End Select
    Loop
    Set out = c
End Sub

' The string at m_jp (its opening quote), with escapes decoded.
Private Function JsonString() As String
    Dim ch As String, st As Long, out As String, n As Long
    m_jp = m_jp + 1
    st = m_jp
    Do
        If m_jp > Len(m_js) Then JsonFail "a string was not closed"
        ch = Mid$(m_js, m_jp, 1)
        If ch = """" Then
            out = out & Mid$(m_js, st, m_jp - st)
            m_jp = m_jp + 1
            Exit Do
        ElseIf ch = "\" Then
            out = out & Mid$(m_js, st, m_jp - st)
            m_jp = m_jp + 1
            ch = Mid$(m_js, m_jp, 1)
            Select Case ch
                Case "n": out = out & vbLf
                Case "t": out = out & vbTab
                Case "r": out = out & vbCr
                Case "b": out = out & Chr$(8)
                Case "f": out = out & Chr$(12)
                Case "u"
                    If m_jp + 4 > Len(m_js) Then JsonFail "a \u escape was cut short"
                    n = Val("&H" & Mid$(m_js, m_jp + 1, 4) & "&")
                    out = out & ChrW(n)
                    m_jp = m_jp + 4
                Case """", "\", "/": out = out & ch
                Case Else: JsonFail "bad escape \" & ch
            End Select
            m_jp = m_jp + 1
            st = m_jp
        Else
            m_jp = m_jp + 1
        End If
    Loop
    JsonString = out
End Function

' ---------------------------------------------------------------------------
' Interface parts the user connects to the AI
' ---------------------------------------------------------------------------

' Declare a place in the app where AI can be used. Call it at start-up (it is
' remembered for the session, not saved); the user's on/off choice is saved.
'   key         short name ("note")
'   caption     what the user sees in the list ("Summarise note")
'   formName    the form the part is on ("" for any form)
'   inputCtl    control whose text is the message ("" for none)
'   outputCtl   control that receives the reply ("" to only return it)
'   template    the message sent, with {input} for the input text ("" is just it)
'   triggerCtl  button that starts it; ApplyAiParts shows it only while on
'   defaultOn   True starts the part on; the default is off until the user ticks it
Public Sub AiBind(ByVal key As String, ByVal caption As String, Optional ByVal formName As String = "", _
        Optional ByVal inputCtl As String = "", Optional ByVal outputCtl As String = "", _
        Optional ByVal template As String = "", Optional ByVal triggerCtl As String = "", _
        Optional ByVal defaultOn As Boolean = False)
    If Len(key) = 0 Or InStr(key, "|") > 0 Or InStr(key, """") > 0 Or InStr(key, ";") > 0 Then Err.Raise 5, , "An AI part's key must be text without | ; or quotes: " & key
    If m_parts Is Nothing Then Set m_parts = CreateObject("Scripting.Dictionary")
    m_parts(key) = Array(caption, formName, inputCtl, outputCtl, template, triggerCtl, defaultOn)
End Sub

' Remove a declared part.
Public Sub AiUnbind(ByVal key As String)
    If m_parts Is Nothing Then Exit Sub
    If m_parts.Exists(key) Then m_parts.Remove key
End Sub

' The declared parts as "key|Caption", in the order they were declared.
Public Function AiPartList() As Variant
    Dim k As Variant, out() As String, n As Long
    If m_parts Is Nothing Then AiPartList = Array(): Exit Function
    If m_parts.Count = 0 Then AiPartList = Array(): Exit Function
    ReDim out(0 To m_parts.Count - 1)
    For Each k In m_parts.Keys
        out(n) = k & "|" & m_parts(k)(0): n = n + 1
    Next
    AiPartList = out
End Function

' True when the user has this part turned on (or, never chosen, it was declared on).
Public Function AiPartEnabled(ByVal key As String) As Boolean
    Dim d As String, s As String
    If m_parts Is Nothing Then Exit Function
    If Not m_parts.Exists(key) Then Exit Function
    d = IIf(m_parts(key)(6), "1", "0")
    s = GetSetting(APP_NAME, SETTINGS_KEY, "Part_" & key, d)
    AiPartEnabled = (s = "1")
End Function

' Turn a part on or off for this Windows user (saved).
Public Sub AiSetPartEnabled(ByVal key As String, ByVal on_ As Boolean)
    If m_parts Is Nothing Then Err.Raise 5, , "No AI parts are declared."
    If Not m_parts.Exists(key) Then Err.Raise 5, , "Unknown AI part: " & key
    SaveSetting APP_NAME, SETTINGS_KEY, "Part_" & key, IIf(on_, "1", "0")
End Sub

' Show or hide the trigger buttons on frm: a part's button is visible while
' the part is on. Call it from the form's Load and after the user changes
' the parts (ShowAI). Parts declared for other forms are left alone.
Public Sub ApplyAiParts(frm As Form)
    Dim k As Variant, p As Variant, c As Control
    If m_parts Is Nothing Then Exit Sub
    For Each k In m_parts.Keys
        p = m_parts(k)
        If Len(p(5)) > 0 And (Len(p(1)) = 0 Or p(1) = frm.Name) Then
            Set c = Nothing
            On Error Resume Next
            Set c = frm.Controls(p(5))
            On Error GoTo 0
            If Not c Is Nothing Then SetVisible c, AiPartEnabled(CStr(k))
        End If
    Next
End Sub

' Run a part on frm: send the template with the input control's text and put
' the reply in the output control. False, with AiLastError set, when the part
' is off, has nothing to send or the call failed; nothing is raised, so a
' button handler can call it directly. Afterwards the app's optional
' AiOnResponse(partKey, prompt, text) runs.
Public Function AiRunPart(ByVal key As String, frm As Form) As Boolean
    Dim p As Variant, txt As String, prompt As String, answer As String
    On Error GoTo Fail
    If m_parts Is Nothing Then Err.Raise 5, , "No AI parts are declared."
    If Not m_parts.Exists(key) Then Err.Raise 5, , "Unknown AI part: " & key
    p = m_parts(key)
    If Not AiPartEnabled(key) Then
        m_lastError = p(0) & " is turned off. Turn it on in the AI window."
        Exit Function
    End If
    If Len(p(2)) > 0 Then txt = Trim$(Nz(frm.Controls(p(2)).Value, ""))
    If Len(p(2)) > 0 And Len(txt) = 0 Then
        m_lastError = "There is nothing to send."
        Exit Function
    End If
    prompt = IIf(Len(p(4)) > 0, Replace(p(4), "{input}", txt), txt)
    DoCmd.Hourglass True
    If Not AiTry(prompt, answer) Then DoCmd.Hourglass False: Exit Function
    DoCmd.Hourglass False
    If Len(p(3)) > 0 Then frm.Controls(p(3)).Value = answer
    On Error Resume Next
    Application.Run "AiOnResponse", key, prompt, answer      ' optional; absent is fine
    On Error GoTo 0
    AiRunPart = True
    Exit Function
Fail:
    DoCmd.Hourglass False
    m_lastError = Err.Description
End Function

' ---------------------------------------------------------------------------
' The AI window
' ---------------------------------------------------------------------------

' Open the AI window: type a message, read the reply, choose the service and
' model, set the API key for this session, and tick the interface parts that
' may use the AI. prefill starts the message box with that text. Returns
' True when the window was opened.
Public Function ShowAI(Optional ByVal prefill As String = "") As Boolean
    On Error GoTo Fail
    RunDialog "dlgAI", prefill
    ShowAI = True
    Exit Function
Fail:
    RaiseAgain "modKitAI.ShowAI"
End Function

' Build dlgAI, the AI window ShowAI opens. Call it once with the app's
' dialogs, and copy kit\forms\dlgAI.vba to src\forms.
Public Sub BuildAiDialog()
    Dim c As Control, w As Long, pad As Long, fw As Long
    w = px(640): pad = px(20): fw = px(600)
    NewDialog "AI assistant", w, px(668)
    Ev F, "OnLoad"
    Lbl "lblProv", "Service", pad, px(18), px(180), px(18), 9, True, CLR_MUTED_FG
    Cbo "cboProvider", "", "", pad, px(38), px(190), px(30)
    Lbl "lblModel", "Model", pad + px(202), px(18), px(180), px(18), 9, True, CLR_MUTED_FG
    Txt "txtModel", "", pad + px(202), px(38), px(228), px(30)
    Btn "btnKey", "API key", pad + px(442), px(38), px(158), px(30), "outline"
    Lbl "lblEndpoint", "Endpoint (empty for the service's own)", pad, px(78), fw, px(18), 9, True, CLR_MUTED_FG
    Txt "txtEndpoint", "", pad, px(98), fw, px(30)
    Lbl "lblMsg", "Message", pad, px(140), fw, px(18), 9, True, CLR_MUTED_FG
    Set c = Txt("txtMessage", "", pad, px(160), fw, px(96))
    c.EnterKeyBehavior = True: c.ScrollBars = 2: c.VerticalAnchor = 0
    c.TopMargin = 60: c.LeftMargin = 90: c.RightMargin = 90
    Btn "btnSend", "Send", pad, px(266), px(100), px(32)
    Btn "btnNew", "New chat", pad + px(110), px(266), px(100), px(32), "outline"
    Set c = Lbl("lblStatus", "", pad + px(224), px(266), px(376), px(32), 9, False, CLR_MUTED_FG)
    c.TextAlign = 3
    Lbl "lblReply", "Response", pad, px(310), fw, px(18), 9, True, CLR_MUTED_FG
    Set c = Txt("txtResponse", "", pad, px(330), fw, px(150))
    c.Locked = True: c.ScrollBars = 2: c.TopMargin = 60: c.LeftMargin = 90: c.RightMargin = 90
    Lbl "lblParts", "Interface parts connected to the AI (tick the ones that may use it)", pad, px(494), fw, px(18), 9, True, CLR_MUTED_FG
    Set c = Lst("lstParts", "", pad, px(514), fw, px(96), 2, "0;" & (fw - px(40)))
    c.ColumnHeads = False
    c.MultiSelect = 1                                   ' simple: a click ticks or unticks
    Set c = Btn("btnClose", "Close", w - pad - px(88), px(622), px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgAI"
End Sub
