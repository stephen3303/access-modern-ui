Option Compare Database
Option Explicit

' Access UI kit - talk to an AI service over its web API.
'
' One conversation, one service at a time. The app picks a service with
' AIConfigure (a preset for "anthropic" or "openai", or any other service that
' speaks one of those two request formats: Azure OpenAI, OpenRouter, Ollama,
' LM Studio and so on, by giving its endpoint), gives the key to AISetKey, and
' calls AIAsk "question" to get the reply as text. Earlier turns are kept and
' sent with each question, so the model sees the conversation (AIReset starts
' a new one; AIMaxMessages caps how much is sent).
'
' Adapters. A service is described by four settings, not by code: the request
' format ("anthropic" Messages or "openai" Chat Completions), the endpoint, how
' the key is sent (an "x-api-key" header or "Authorization: Bearer"), and the
' model. AIBuildBody and AIParseReply are the only places that know the two
' formats, so a third format is one more branch in each.
'
' The key is never in the code or saved by the kit. AISetKey keeps it in memory
' for the session; with none set, the environment variable for the service
' (ANTHROPIC_API_KEY or OPENAI_API_KEY, or the one named in AIConfigure) is
' used. AIAskForKey asks the user for it in a masked box. AISaveConfig saves
' the provider, endpoint and model (not the key) per user so they survive
' restarts; AILoadConfig reads them back.
'
' Binding the interface. AIBind frm, "txtAsk", "txtReply", "btnSend" [, "lblAIStatus"]
' names the form's input box, response box, Send button and optional status
' label. Which controls those are is the app's choice (or, with
' AIBindFromSetting, the user's: AISaveBinding writes the names to the user's
' settings). The Send button's OnClick becomes =AISendClick("frmName"), which
' sends the input, shows "Thinking..." in the status label while it waits, puts
' the reply in the response box and shows an error in the status label (or the
' response box when there is none) instead of a dialog. Then it tells the rest
' of the app, so the reply can go into other systems: the form's optional
' Public Sub AIReplied(ByVal reply As String), and the procedure named by
' AISetReplyHandler "modName.Proc" (called with the reply and the question).
'
' The call is synchronous: Access stops responding for the seconds the service
' takes (AISetTimeout bounds it, default 60 s). The bound click redraws the
' status label first so the user sees why.
'
' The app provides APP_NAME (the saved settings' key) and RaiseAgain (modErrors);
' LogError is used by the bound click. KitInputBox and KitMsgBox are the kit's.

Public Const AI_ANTHROPIC As String = "anthropic"
Public Const AI_OPENAI As String = "openai"

' Error numbers AIAsk raises (vbObjectError + n), so a caller can tell them apart.
Public Const AI_ERR_NOT_CONFIGURED As Long = vbObjectError + 5101   ' no provider, endpoint or model
Public Const AI_ERR_NO_KEY As Long = vbObjectError + 5102           ' no key set
Public Const AI_ERR_SERVICE As Long = vbObjectError + 5103          ' the service answered with an error
Public Const AI_ERR_BAD_REPLY As Long = vbObjectError + 5104        ' the answer had no text in it

Private Const SETTINGS_KEY As String = "AI"
Private Const ANTHROPIC_URL As String = "https://api.anthropic.com/v1/messages"
Private Const ANTHROPIC_VERSION As String = "2023-06-01"
Private Const OPENAI_URL As String = "https://api.openai.com/v1/chat/completions"

Private m_format As String        ' request format: AI_ANTHROPIC or AI_OPENAI
Private m_provider As String      ' preset name or the caller's own label
Private m_endpoint As String
Private m_model As String
Private m_keyEnv As String        ' environment variable that holds the key
Private m_key As String           ' set by AISetKey, kept in memory only
Private m_system As String        ' system prompt
Private m_maxTokens As Long       ' 0 leaves it out (OpenAI reasoning models refuse max_tokens)
Private m_timeoutSec As Long
Private m_maxMessages As Long
Private m_history As Collection   ' Array(role, text), oldest first
Private m_binds As Collection     ' Array(input, output, send, status), keyed by form name
Private m_handler As String       ' "modName.Proc" called with (reply, question)
Private m_busy As Boolean         ' a bound send is running

' ---- Configuration ----

' Choose the service. provider is "anthropic" or "openai" for a preset, or any
' label of your own with format and endpoint given ("ollama", format "openai",
' endpoint "http://localhost:11434/v1/chat/completions"). model is the service's
' model id; it is not defaulted because ids change. keyEnv names the environment
' variable that holds the key, default ANTHROPIC_API_KEY / OPENAI_API_KEY /
' <PROVIDER>_API_KEY. Starts a new conversation.
Public Sub AIConfigure(ByVal provider As String, ByVal model As String, _
        Optional ByVal endpoint As String = "", Optional ByVal fmt As String = "", _
        Optional ByVal keyEnv As String = "")
    Dim p As String
    p = LCase$(Trim$(provider))
    If Len(p) = 0 Then Err.Raise AI_ERR_NOT_CONFIGURED, "modKitAI.AIConfigure", "Name the AI service."
    fmt = LCase$(Trim$(fmt))
    If Len(fmt) = 0 Then
        If p = AI_ANTHROPIC Then fmt = AI_ANTHROPIC Else fmt = AI_OPENAI
    End If
    If fmt <> AI_ANTHROPIC And fmt <> AI_OPENAI Then
        Err.Raise AI_ERR_NOT_CONFIGURED, "modKitAI.AIConfigure", "Unknown request format '" & fmt & "' (use anthropic or openai)."
    End If
    endpoint = Trim$(endpoint)
    If Len(endpoint) = 0 Then
        If p = AI_ANTHROPIC Then
            endpoint = ANTHROPIC_URL
        ElseIf p = AI_OPENAI Then
            endpoint = OPENAI_URL
        Else
            Err.Raise AI_ERR_NOT_CONFIGURED, "modKitAI.AIConfigure", "Give the endpoint for '" & provider & "'."
        End If
    End If
    If Len(Trim$(keyEnv)) = 0 Then
        If p = AI_ANTHROPIC Then
            keyEnv = "ANTHROPIC_API_KEY"
        ElseIf p = AI_OPENAI Then
            keyEnv = "OPENAI_API_KEY"
        Else
            keyEnv = UCase$(p) & "_API_KEY"
        End If
    End If
    If m_provider <> p Then m_key = ""                  ' a key belongs to one service
    m_provider = p
    m_format = fmt
    m_endpoint = endpoint
    m_model = Trim$(model)
    m_keyEnv = Trim$(keyEnv)
    If m_maxTokens = 0 Then m_maxTokens = 1024
    If m_timeoutSec = 0 Then m_timeoutSec = 60
    If m_maxMessages = 0 Then m_maxMessages = 20
    AIReset
End Sub

' True when a service and model are chosen.
Public Function AIConfigured() As Boolean
    AIConfigured = (Len(m_endpoint) > 0) And (Len(m_model) > 0)
End Function

' The key for this session. Held in memory only; never saved by the kit.
Public Sub AISetKey(ByVal key As String)
    m_key = Trim$(key)
End Sub

' True when a key is available (set, or in the service's environment variable).
Public Function AIHasKey() As Boolean
    AIHasKey = Len(CurrentKey()) > 0
End Function

' Ask the user for the key in a masked box and keep it for the session. False
' when the user cancelled or typed nothing.
Public Function AIAskForKey() As Boolean
    Dim a As Variant
    a = KitInputBox("Enter the API key for " & m_provider & ". It is kept for this session only.", _
        "AI service key", "", masked:=True)
    If IsNull(a) Then Exit Function
    If Len(Trim$(CStr(a))) = 0 Then Exit Function
    AISetKey CStr(a)
    AIAskForKey = True
End Function

' Instructions that frame every answer ("You answer questions about our
' maintenance records. Be brief."). "" for none.
Public Sub AISetSystemPrompt(ByVal text As String)
    m_system = text
End Sub

' Longest reply to ask for, in tokens. 0 leaves it out of the request.
Public Sub AISetMaxTokens(ByVal n As Long)
    If n < 0 Then n = 0
    m_maxTokens = n
End Sub

' Seconds to wait for the service to answer (5 to 600).
Public Sub AISetTimeout(ByVal seconds As Long)
    If seconds < 5 Then seconds = 5
    If seconds > 600 Then seconds = 600
    m_timeoutSec = seconds
End Sub

' Most earlier messages (questions and replies) sent with a new question.
Public Sub AIMaxMessages(ByVal n As Long)
    If n < 1 Then n = 1
    m_maxMessages = n
End Sub

' Start a new conversation (forget the earlier questions and replies).
Public Sub AIReset()
    Set m_history = New Collection
End Sub

' Procedure to call after every bound reply, to pass it to another system:
' "modOutbox.SaveReply", declared (ByVal reply As String, ByVal question As String).
' "" removes it.
Public Sub AISetReplyHandler(ByVal procName As String)
    m_handler = Trim$(procName)
End Sub

' Save the provider, endpoint, model and key variable (not the key) for this
' Windows user, for AILoadConfig at the next start.
Public Sub AISaveConfig()
    SaveSetting APP_NAME, SETTINGS_KEY, "Provider", m_provider
    SaveSetting APP_NAME, SETTINGS_KEY, "Format", m_format
    SaveSetting APP_NAME, SETTINGS_KEY, "Endpoint", m_endpoint
    SaveSetting APP_NAME, SETTINGS_KEY, "Model", m_model
    SaveSetting APP_NAME, SETTINGS_KEY, "KeyEnv", m_keyEnv
End Sub

' Restore what AISaveConfig saved. False (nothing changed) when none was saved
' or it is incomplete.
Public Function AILoadConfig() As Boolean
    Dim p As String, m As String
    p = GetSetting(APP_NAME, SETTINGS_KEY, "Provider", "")
    m = GetSetting(APP_NAME, SETTINGS_KEY, "Model", "")
    If Len(p) = 0 Or Len(m) = 0 Then Exit Function
    AIConfigure p, m, GetSetting(APP_NAME, SETTINGS_KEY, "Endpoint", ""), _
        GetSetting(APP_NAME, SETTINGS_KEY, "Format", ""), GetSetting(APP_NAME, SETTINGS_KEY, "KeyEnv", "")
    AILoadConfig = True
End Function

' ---- Asking ----

' Send question (with context text placed before it, when given) and return the
' reply text. The turn is added to the conversation only when the service
' answered. Errors are raised with the numbers above; see AIAsked for a call
' that reports them instead.
Public Function AIAsk(ByVal question As String, Optional ByVal context As String = "") As String
    Dim body As String, answer As String, reply As String, msg As String
    On Error GoTo Fail
    If Not AIConfigured() Then Err.Raise AI_ERR_NOT_CONFIGURED, , "No AI service is set up. Call AIConfigure first."
    If Len(Trim$(question)) = 0 Then Err.Raise 5, , "Type a question first."
    If Len(context) > 0 Then msg = "Context:" & vbLf & context & vbLf & vbLf & "Question:" & vbLf & question Else msg = question
    AddMessage "user", msg
    On Error GoTo FailTurn
    body = AIBuildBody(m_format, m_model, m_system, m_maxTokens, m_history)
    answer = HttpPost(body)
    reply = AIParseReply(m_format, answer)
    AddMessage "assistant", reply
    AIAsk = reply
    Exit Function
FailTurn:
    If m_history.Count > 0 Then m_history.Remove m_history.Count    ' drop the question that got no answer
Fail:
    RaiseAgain "modKitAI.AIAsk", "service " & m_provider & ", model " & m_model
End Function

' AIAsk that never raises: True with the reply in reply, or False with the
' reason in errMsg (written for the user).
Public Function AIAsked(ByVal question As String, ByRef reply As String, ByRef errMsg As String, _
        Optional ByVal context As String = "") As Boolean
    On Error GoTo Fail
    reply = AIAsk(question, context)
    errMsg = ""
    AIAsked = True
    Exit Function
Fail:
    reply = ""
    errMsg = Err.Description
    LogError "modKitAI.AIAsked", "service " & m_provider
End Function

' ---- Bound controls ----

' Connect a form's controls to the AI: inputName (where the user types),
' outputName (where the reply goes), sendName (the Send button) and optionally
' statusName (a label for "Thinking..." and errors). The form can be open in
' Form view. Raises if a named control is missing. Binding again replaces the
' earlier binding for that form.
Public Sub AIBind(frm As Form, ByVal inputName As String, ByVal outputName As String, _
        ByVal sendName As String, Optional ByVal statusName As String = "")
    Dim nm As Variant
    On Error GoTo Fail
    For Each nm In Array(inputName, outputName, sendName)
        RequireControl frm, CStr(nm)
    Next
    If Len(statusName) > 0 Then RequireControl frm, statusName
    If m_binds Is Nothing Then Set m_binds = New Collection
    On Error Resume Next
    m_binds.Remove frm.Name
    On Error GoTo Fail
    m_binds.Add Array(inputName, outputName, sendName, statusName), frm.Name
    frm.Controls(sendName).OnClick = "=AISendClick(""" & frm.Name & """)"
    Exit Sub
Fail:
    RaiseAgain "modKitAI.AIBind", "form " & frm.Name
End Sub

' Remove the binding on a form (the Send button's OnClick is cleared).
Public Sub AIUnbind(frm As Form)
    Dim b As Variant
    On Error Resume Next
    If m_binds Is Nothing Then Exit Sub
    b = m_binds(frm.Name)
    If IsEmpty(b) Then Exit Sub
    frm.Controls(b(2)).OnClick = ""
    m_binds.Remove frm.Name
End Sub

' Save which controls AIBind should use on a form, for this Windows user, so
' they can be changed without editing code.
Public Sub AISaveBinding(ByVal formName As String, ByVal inputName As String, ByVal outputName As String, _
        ByVal sendName As String, Optional ByVal statusName As String = "")
    SaveSetting APP_NAME, SETTINGS_KEY, "Bind_" & formName, _
        inputName & "|" & outputName & "|" & sendName & "|" & statusName
End Sub

' AIBind with the names AISaveBinding saved for the form, or defaultSpec
' ("input|output|send|status") when none were saved. False when neither exists.
Public Function AIBindFromSetting(frm As Form, Optional ByVal defaultSpec As String = "") As Boolean
    Dim spec As String, p() As String
    spec = GetSetting(APP_NAME, SETTINGS_KEY, "Bind_" & frm.Name, defaultSpec)
    If Len(spec) = 0 Then Exit Function
    p = Split(spec & "|||", "|")
    If Len(p(0)) = 0 Or Len(p(1)) = 0 Or Len(p(2)) = 0 Then Exit Function
    AIBind frm, p(0), p(1), p(2), p(3)
    AIBindFromSetting = True
End Function

' The Send button's click (set by AIBind; call it yourself from a button's code
' too). Returns True when a reply was shown.
Public Function AISendClick(ByVal formName As String) As Boolean
    Dim frm As Form, b As Variant, q As String, reply As String, errMsg As String
    If m_busy Then Exit Function
    On Error GoTo Fail
    b = BindingOf(formName)
    Set frm = Forms(formName)
    q = Trim$(Nz(frm.Controls(b(0)).Value, ""))
    If Len(q) = 0 Then
        ShowStatus frm, b, "Type a question first.", True
        Exit Function
    End If
    m_busy = True
    SetEnabled frm.Controls(b(2)), False
    ShowStatus frm, b, "Thinking...", False
    DoEvents                                    ' paint the status before Access blocks on the network
    If AIAsked(q, reply, errMsg) Then
        frm.Controls(b(1)).Value = reply
        frm.Controls(b(0)).Value = Null
        ShowStatus frm, b, "", False
        AISendClick = True
    Else
        ShowStatus frm, b, errMsg, True
    End If
    SetEnabled frm.Controls(b(2)), True
    m_busy = False
    If AISendClick Then Announce frm, reply, q
    Exit Function
Fail:
    m_busy = False
    LogError "modKitAI.AISendClick", "form " & formName
End Function

' ---- Request and reply adapters (no network; testable) ----

' The JSON request for a format. history holds Array(role, text) items, oldest
' first, ending with the user's question; system may be "". maxTokens 0 omits
' the limit (the Anthropic format requires one, so it gets 1024).
Public Function AIBuildBody(ByVal fmt As String, ByVal model As String, ByVal system As String, _
        ByVal maxTokens As Long, history As Collection) As String
    Dim s As String, i As Long, sep As String, m As Variant
    s = "{""model"":" & JsonQuote(model)
    If fmt = AI_ANTHROPIC Then
        If maxTokens <= 0 Then maxTokens = 1024
        s = s & ",""max_tokens"":" & maxTokens
        If Len(system) > 0 Then s = s & ",""system"":" & JsonQuote(system)
    ElseIf maxTokens > 0 Then
        s = s & ",""max_tokens"":" & maxTokens
    End If
    s = s & ",""messages"":["
    If fmt = AI_OPENAI And Len(system) > 0 Then
        s = s & "{""role"":""system"",""content"":" & JsonQuote(system) & "}"
        sep = ","
    End If
    For i = Window(history) + 1 To history.Count
        m = history(i)
        s = s & sep & "{""role"":" & JsonQuote(CStr(m(0))) & ",""content"":" & JsonQuote(CStr(m(1))) & "}"
        sep = ","
    Next
    AIBuildBody = s & "]}"
End Function

' The reply text in a successful response. Raises AI_ERR_SERVICE with the
' service's own message when the response is an error, and AI_ERR_BAD_REPLY
' when there is no text.
Public Function AIParseReply(ByVal fmt As String, ByVal json As String) As String
    Dim p As Long, t As String, errText As String
    errText = AIParseError(json)
    If Len(errText) > 0 Then Err.Raise AI_ERR_SERVICE, , errText
    If fmt = AI_ANTHROPIC Then
        ' content is a list of blocks; the text blocks, joined, are the reply.
        p = JsonFindKey(json, "content", 1)
        Do While p > 0
            p = JsonFindKey(json, "text", p)
            If p > 0 Then t = t & JsonReadString(json, p, p)
        Loop
    Else
        p = JsonFindKey(json, "message", 1)
        If p > 0 Then p = JsonFindKey(json, "content", p)
        If p > 0 Then t = JsonReadString(json, p, p)
    End If
    If Len(t) = 0 Then Err.Raise AI_ERR_BAD_REPLY, , "The AI service sent no text."
    AIParseReply = t
End Function

' The message inside an error response ({"error":{"message":"..."}} or
' {"error":"..."}), else "".
Public Function AIParseError(ByVal json As String) As String
    Dim p As Long, q As Long, c As String
    p = JsonFindKey(json, "error", 1)
    If p = 0 Then Exit Function
    c = NextChar(json, p)
    If c = """" Then
        AIParseError = JsonReadString(json, p, q)
    ElseIf c = "{" Then
        q = JsonFindKey(json, "message", p)
        If q > 0 Then AIParseError = JsonReadString(json, q, q) Else AIParseError = "The AI service reported an error."
    End If
End Function

' A string as a JSON string literal. Non-ASCII characters are written as \uXXXX
' so the request body is plain ASCII whatever the text.
Public Function JsonQuote(ByVal s As String) As String
    Dim i As Long, code As Long, out As String, ch As String
    out = """"
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        Select Case code
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 8: out = out & "\b"
            Case 9: out = out & "\t"
            Case 10: out = out & "\n"
            Case 12: out = out & "\f"
            Case 13: out = out & "\r"
            Case Is < 32, Is > 126: out = out & "\u" & Right$("0000" & Hex$(code), 4)
            Case Else: out = out & ch
        End Select
    Next
    JsonQuote = out & """"
End Function

' ---- Private: JSON ----

' Position just after the colon of the first object key named key at or after
' start, skipping over string values so text inside them never matches. 0 when
' there is none.
Private Function JsonFindKey(ByVal json As String, ByVal key As String, ByVal start As Long) As Long
    Dim i As Long, n As Long, s As String, j As Long
    n = Len(json)
    i = start
    Do While i <= n
        If Mid$(json, i, 1) = """" Then
            s = JsonReadString(json, i, j)      ' j is after the closing quote
            i = j
            If s = key Then
                Do While i <= n
                    If InStr(" " & vbTab & vbCr & vbLf, Mid$(json, i, 1)) = 0 Then Exit Do
                    i = i + 1
                Loop
                If Mid$(json, i, 1) = ":" Then JsonFindKey = i + 1: Exit Function
            End If
        Else
            i = i + 1
        End If
    Loop
End Function

' The string value starting at the first quote at or after pos, decoded; after
' is set to the position following its closing quote. "" (after = pos) when the
' next value is not a string.
Private Function JsonReadString(ByVal json As String, ByVal pos As Long, ByRef after As Long) As String
    Dim i As Long, n As Long, ch As String, out As String, code As Long
    n = Len(json)
    i = pos
    Do While i <= n
        If InStr(" " & vbTab & vbCr & vbLf, Mid$(json, i, 1)) = 0 Then Exit Do
        i = i + 1
    Loop
    If Mid$(json, i, 1) <> """" Then after = pos: Exit Function
    i = i + 1
    Do While i <= n
        ch = Mid$(json, i, 1)
        If ch = """" Then Exit Do
        If ch = "\" Then
            i = i + 1
            ch = Mid$(json, i, 1)
            Select Case ch
                Case "n": out = out & vbLf
                Case "t": out = out & vbTab
                Case "r": out = out & vbCr
                Case "b": out = out & Chr$(8)
                Case "f": out = out & Chr$(12)
                Case "u"
                    code = CLng("&H" & Mid$(json, i + 1, 4))
                    If code > 32767 Then code = code - 65536
                    out = out & ChrW$(code)
                    i = i + 4
                Case Else: out = out & ch           ' \" \\ \/
            End Select
        Else
            out = out & ch
        End If
        i = i + 1
    Loop
    after = i + 1
    JsonReadString = out
End Function

' The first character at or after pos that is not white space.
Private Function NextChar(ByVal json As String, ByVal pos As Long) As String
    Do While pos <= Len(json)
        If InStr(" " & vbTab & vbCr & vbLf, Mid$(json, pos, 1)) = 0 Then NextChar = Mid$(json, pos, 1): Exit Function
        pos = pos + 1
    Loop
End Function

' ---- Private: conversation and network ----

' Add a turn, creating the history when the app never called AIReset.
Private Sub AddMessage(ByVal role As String, ByVal text As String)
    If m_history Is Nothing Then Set m_history = New Collection
    m_history.Add Array(role, text)
End Sub

' How many of the oldest messages to leave out so at most m_maxMessages are
' sent, keeping the first sent message a user turn (both formats require it).
Private Function Window(history As Collection) As Long
    Dim cap As Long, skip As Long
    cap = m_maxMessages
    If cap < 1 Then cap = 20
    skip = history.Count - cap
    If skip < 0 Then skip = 0
    If skip > 0 Then
        If history(skip + 1)(0) <> "user" And skip < history.Count Then skip = skip + 1
    End If
    Window = skip
End Function

' The key to send: AISetKey's, else the service's environment variable.
Private Function CurrentKey() As String
    CurrentKey = m_key
    If Len(CurrentKey) = 0 And Len(m_keyEnv) > 0 Then CurrentKey = Trim$(Environ$(m_keyEnv))
End Function

' POST body to the endpoint and return the response text (UTF-8). A service
' error status becomes AI_ERR_SERVICE with the service's message when it sent one.
Private Function HttpPost(ByVal body As String) As String
    Dim req As Object, key As String, text As String, st As Object, msg As String
    key = CurrentKey()
    If Len(key) = 0 Then
        Err.Raise AI_ERR_NO_KEY, , "No API key. Set the " & m_keyEnv & " environment variable or enter the key."
    End If
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
    req.Open "POST", m_endpoint, False
    req.SetTimeouts 10000, 10000, 30000, m_timeoutSec * 1000
    req.SetRequestHeader "Content-Type", "application/json"
    If m_format = AI_ANTHROPIC Then
        req.SetRequestHeader "x-api-key", key
        req.SetRequestHeader "anthropic-version", ANTHROPIC_VERSION
    Else
        req.SetRequestHeader "Authorization", "Bearer " & key
    End If
    On Error GoTo NoReach
    req.Send body
    On Error GoTo 0
    Set st = CreateObject("ADODB.Stream")      ' ResponseText guesses the charset; the services send UTF-8
    st.Type = 1
    st.Open
    st.Write req.ResponseBody
    st.Position = 0
    st.Type = 2
    st.Charset = "utf-8"
    text = st.ReadText
    st.Close
    If req.Status < 200 Or req.Status >= 300 Then
        msg = AIParseError(text)
        If Len(msg) = 0 Then msg = "The AI service answered with status " & req.Status & " " & req.StatusText & "."
        Err.Raise AI_ERR_SERVICE, , msg
    End If
    HttpPost = text
    Exit Function
NoReach:
    Err.Raise AI_ERR_SERVICE, , "Could not reach the AI service (" & Err.Description & ")."
End Function

' ---- Private: bound controls ----

' The binding for a form, or an error naming the form.
Private Function BindingOf(ByVal formName As String) As Variant
    On Error Resume Next
    BindingOf = m_binds(formName)
    If IsEmpty(BindingOf) Then
        On Error GoTo 0
        Err.Raise 5, , "Form " & formName & " is not bound to the AI. Call AIBind first."
    End If
End Function

' Raise a clear error when frm has no control called nm.
Private Sub RequireControl(frm As Form, ByVal nm As String)
    Dim c As Control
    On Error Resume Next
    Set c = frm.Controls(nm)
    On Error GoTo 0
    If c Is Nothing Then Err.Raise 2465, , "Form " & frm.Name & " has no control named " & nm & "."
End Sub

' Show text in the status label, else (an error with no label) in the response box.
Private Sub ShowStatus(frm As Form, b As Variant, ByVal text As String, ByVal isError As Boolean)
    If Len(b(3)) > 0 Then
        frm.Controls(b(3)).Caption = text
        frm.Controls(b(3)).ForeColor = IIf(isError, Th(CLR_DESTRUCTIVE), Th(CLR_MUTED_FG))
    ElseIf isError Then
        frm.Controls(b(1)).Value = text
    End If
    DoEvents
End Sub

' Tell the rest of the app about a reply: the form's AIReplied, then the
' registered handler. A hook that fails is logged and never undoes the reply.
Private Sub Announce(frm As Form, ByVal reply As String, ByVal question As String)
    On Error Resume Next
    CallByName frm, "AIReplied", VbMethod, reply
    If Err.Number <> 0 And Err.Number <> 438 And Err.Number <> 450 Then LogError "modKitAI.AIReplied", "form " & frm.Name
    Err.Clear
    If Len(m_handler) > 0 Then
        Application.Run m_handler, reply, question
        If Err.Number <> 0 Then LogError "modKitAI.Announce", "handler " & m_handler
    End If
End Sub
