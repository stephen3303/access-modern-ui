Option Compare Database
Option Explicit

' Access UI kit - connect parts of the app to an AI service over its web API.
'
' One call sends a message and returns the reply:
'   reply = KitAISend("Summarize this record: " & Me.txtNotes)
' and one call wires a form's controls to it:
'   KitAIAsk Me                       ' the form's AI input -> service -> its AI output
'
' Any service that takes a JSON request over HTTPS works, because the request
' is described by settings, not by code. A provider profile has:
'   Endpoint      the URL to POST to
'   Headers       "Name: value|Name: value"; {key} is replaced by the API key
'   Payload       the JSON body with placeholders: {prompt}, {system} and {model}
'                 (each becomes a quoted, escaped JSON string) and {maxtokens}
'                 (a number)
'   ResponsePath  where the reply text sits in the JSON answer, such as
'                 choices[0].message.content
'   Model, System, MaxTokens, TimeoutSec, KeyEnv, MaxClass
' KitAIUseProvider loads a ready-made profile ("anthropic", "openai", "ollama"
' for a local server, "custom" for a blank one); KitAISetting changes any field
' afterwards, so an OpenAI-compatible service (Azure, a gateway, LM Studio) is
' the "openai" profile with another Endpoint. Settings are saved per Windows
' user under APP_NAME, in the "AI" key (the registry).
'
' The API key is never in the code or in the saved settings. It is read from
' the Windows environment variable named by the KeyEnv setting (ANTHROPIC_API_KEY
' by default for "anthropic", OPENAI_API_KEY for "openai"), or held in memory
' for the session with KitAISetKey (for a key the user types in). A service that
' needs no key (a local one) leaves KeyEnv blank.
'
' Which parts of the interface are AI-connected is the user's choice, saved
' the same way. A bound control has a role:
'   input     its text is the message (the first non-empty input is sent)
'   context   its text is added to the message as "Name: value" lines
'   output    the reply is written here (a text box, or a label's caption)
' KitAIBind form, control, role and KitAIUnbind change the list in code;
' KitAIConnectControls form, anchor lets the user do it with two menus. An
' app marks the connected controls with KitAIMark (a thin accent outline).
'
' What is sent leaves the machine, so the app's classification limits it:
' KitAISend refuses when AppClassification (modKitClassification) is above
' the profile's MaxClass (UNCLASSIFIED unless the profile is set higher, for a
' service the organization has approved for that level).
'
' Calls are synchronous: Access waits (KitBusy shows a message while
' KitAIAsk does). The request times out after TimeoutSec (60).
'
' The app provides APP_NAME, LogError and HandleError (modErrors).

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)

Public Const AI_SETTINGS_KEY As String = "AI"
Public Const AI_ROLE_INPUT As String = "input"
Public Const AI_ROLE_CONTEXT As String = "context"
Public Const AI_ROLE_OUTPUT As String = "output"
Public Const AI_MARK_TAG As String = "aiconnected"

Private m_key As String           ' a key held for the session (KitAISetKey)
Private m_text As String          ' the JSON being parsed
Private m_pos As Long             ' and the position in it

' ---------------------------------------------------------------------------
' Provider profiles and settings
' ---------------------------------------------------------------------------

' The fields of a profile, in the order the settings are listed.
Public Function AIFields() As Variant
    AIFields = Array("Provider", "Endpoint", "Headers", "Payload", "ResponsePath", "Model", _
        "System", "MaxTokens", "TimeoutSec", "KeyEnv", "MaxClass")
End Function

' The profile a provider name stands for, as "Field=value" lines (Split on
' vbLf). "" for an unknown name.
Public Function AIPreset(ByVal provider As String) As String
    Dim q As String
    q = """"
    Select Case LCase$(Trim$(provider))
    Case "anthropic"
        AIPreset = "Endpoint=https://api.anthropic.com/v1/messages" & vbLf & _
            "Headers=x-api-key: {key}|anthropic-version: 2023-06-01|content-type: application/json" & vbLf & _
            "Payload={""model"":{model},""max_tokens"":{maxtokens},""system"":{system},""messages"":[{""role"":""user"",""content"":{prompt}}]}" & vbLf & _
            "ResponsePath=content[0].text" & vbLf & "Model=claude-sonnet-4-5" & vbLf & "KeyEnv=ANTHROPIC_API_KEY"
    Case "openai"
        AIPreset = "Endpoint=https://api.openai.com/v1/chat/completions" & vbLf & _
            "Headers=Authorization: Bearer {key}|content-type: application/json" & vbLf & _
            "Payload={""model"":{model},""messages"":[{""role"":""system"",""content"":{system}},{""role"":""user"",""content"":{prompt}}]}" & vbLf & _
            "ResponsePath=choices[0].message.content" & vbLf & "Model=gpt-4o-mini" & vbLf & "KeyEnv=OPENAI_API_KEY"
    Case "ollama"
        AIPreset = "Endpoint=http://localhost:11434/v1/chat/completions" & vbLf & _
            "Headers=content-type: application/json" & vbLf & _
            "Payload={""model"":{model},""messages"":[{""role"":""system"",""content"":{system}},{""role"":""user"",""content"":{prompt}}]}" & vbLf & _
            "ResponsePath=choices[0].message.content" & vbLf & "Model=llama3.1" & vbLf & "KeyEnv="
    Case "custom"
        AIPreset = "Endpoint=" & vbLf & "Headers=content-type: application/json" & vbLf & _
            "Payload={""model"":{model},""prompt"":{prompt}}" & vbLf & "ResponsePath=" & vbLf & "Model=" & vbLf & "KeyEnv="
    End Select
End Function

' Make provider the selected profile: its preset fills Endpoint, Headers,
' Payload, ResponsePath, Model and KeyEnv (replacing what was there), and
' System, MaxTokens, TimeoutSec and MaxClass keep their values. Raises error
' 5 for a name that isn't anthropic, openai, ollama or custom.
Public Sub KitAIUseProvider(ByVal provider As String)
    Dim p As String, line As Variant, k As Long
    p = AIPreset(provider)
    If Len(p) = 0 Then Err.Raise 5, "modKitAI", "Not a known AI provider (anthropic, openai, ollama, custom): " & provider
    KitAISetting "Provider", LCase$(Trim$(provider))
    For Each line In Split(p, vbLf)
        k = InStr(line, "=")
        KitAISetting Left$(line, k - 1), Mid$(line, k + 1)
    Next
End Sub

' The saved value of a field, or its default when none is saved:
' MaxTokens 1024, TimeoutSec 60, MaxClass UNCLASSIFIED, System a short
' assistant instruction, everything else "".
Public Function KitAISetting(ByVal field As String, Optional ByVal setTo As Variant) As String
    Dim def As String
    Select Case field
        Case "MaxTokens": def = "1024"
        Case "TimeoutSec": def = "60"
        Case "MaxClass": def = "UNCLASSIFIED"
        Case "System": def = "You are a concise assistant inside a database application."
    End Select
    If Not IsMissing(setTo) Then
        SaveSetting APP_NAME, AI_SETTINGS_KEY, field, CStr(setTo)
        KitAISetting = CStr(setTo)
    Else
        KitAISetting = GetSetting(APP_NAME, AI_SETTINGS_KEY, field, def)
    End If
End Function

' Hold an API key in memory for this session only (never saved). Pass "" to
' forget it.
Public Sub KitAISetKey(ByVal key As String)
    m_key = key
End Sub

' The API key in use: the session key, else the variable KeyEnv names, else
' "". The key is never logged or shown.
Private Function AIKey() As String
    Dim v As String
    If Len(m_key) > 0 Then AIKey = m_key: Exit Function
    v = Trim$(KitAISetting("KeyEnv"))
    If Len(v) > 0 Then AIKey = Environ$(v)
End Function

' True when a request can be made: an endpoint is set, and a key is available
' if the headers need one.
Public Function KitAIReady() As Boolean
    KitAIReady = Len(AIProblem()) = 0
End Function

' Why a request can't be made now, in words for the user, or "".
Public Function AIProblem() As String
    Dim ep As String
    ep = Trim$(KitAISetting("Endpoint"))
    If Len(ep) = 0 Then AIProblem = "No AI service is set up. Choose a provider first.": Exit Function
    If Not EndpointAllowed(ep) Then
        AIProblem = "The AI service address must start with https:// (http:// is allowed only for this computer)."
        Exit Function
    End If
    If Len(Trim$(KitAISetting("Payload"))) = 0 Or Len(Trim$(KitAISetting("ResponsePath"))) = 0 Then
        AIProblem = "The AI service has no request body or reply path set."
        Exit Function
    End If
    If InStr(KitAISetting("Headers"), "{key}") > 0 And Len(AIKey()) = 0 Then
        AIProblem = "No API key. Set the " & IIf(Len(KitAISetting("KeyEnv")) > 0, KitAISetting("KeyEnv"), "key") & _
            " environment variable, or enter the key for this session."
    End If
End Function

' https always; http only for localhost and 127.0.0.1.
Private Function EndpointAllowed(ByVal ep As String) As Boolean
    ep = LCase$(Trim$(ep))
    If Left$(ep, 8) = "https://" Then EndpointAllowed = True: Exit Function
    EndpointAllowed = (Left$(ep, 17) = "http://localhost:" Or Left$(ep, 17) = "http://localhost/" _
        Or Left$(ep, 17) = "http://127.0.0.1:" Or Left$(ep, 17) = "http://127.0.0.1/")
End Function

' ---------------------------------------------------------------------------
' Sending a message
' ---------------------------------------------------------------------------

' Send prompt to the selected service and return the reply text. system
' replaces the profile's System setting for this call. Raises an error (5 for a
' refused or misconfigured request, 3000 for a failed call or a reply with no
' text) whose description is fit to show the user. Nothing is logged with the
' prompt or the key.
Public Function KitAISend(ByVal prompt As String, Optional ByVal system As String = "") As String
    Dim why As String, body As String, status As Long, reply As String
    If Len(Trim$(prompt)) = 0 Then Err.Raise 5, "modKitAI", "There is no message to send."
    why = AIProblem()
    If Len(why) > 0 Then Err.Raise 5, "modKitAI", why
    why = AIClassificationProblem()
    If Len(why) > 0 Then Err.Raise 5, "modKitAI", why
    If Len(system) = 0 Then system = KitAISetting("System")
    body = AIRequestBody(prompt, system)
    reply = AIPost(Trim$(KitAISetting("Endpoint")), AIHeaderList(), body, status)
    KitAISend = AIReplyText(reply, status)
End Function

' The app's classification is above what this service is approved for.
Private Function AIClassificationProblem() As String
    Dim limit As String
    limit = KitAISetting("MaxClass")
    If ClassificationLevel(limit) < 0 Then limit = KIT_CLASS_DEFAULT
    If ClassificationLevel(AppClassification()) > ClassificationLevel(limit) Then
        AIClassificationProblem = "This app is marked " & AppClassification() & ", which is above the " & limit & _
            " this AI service is approved for. Nothing was sent."
    End If
End Function

' The request body: the Payload setting with its placeholders filled in.
Public Function AIRequestBody(ByVal prompt As String, ByVal system As String) As String
    Dim s As String, n As Long
    s = KitAISetting("Payload")
    n = Val(KitAISetting("MaxTokens"))
    If n < 1 Then n = 1024
    ' One pass over the placeholders, so text inside a value is never read as one.
    AIRequestBody = FillPlaceholders(s, prompt, system, KitAISetting("Model"), n)
End Function

Private Function FillPlaceholders(ByVal tmpl As String, ByVal prompt As String, ByVal system As String, _
        ByVal model As String, ByVal maxTokens As Long) As String
    Dim out As String, i As Long, j As Long, name As String
    i = 1
    Do While i <= Len(tmpl)
        If Mid$(tmpl, i, 1) = "{" Then
            j = InStr(i, tmpl, "}")
            name = IIf(j > i, LCase$(Mid$(tmpl, i + 1, j - i - 1)), "")
            Select Case name
                Case "prompt": out = out & JsonQuote(prompt): i = j + 1
                Case "system": out = out & JsonQuote(system): i = j + 1
                Case "model": out = out & JsonQuote(model): i = j + 1
                Case "maxtokens": out = out & CStr(maxTokens): i = j + 1
                Case Else: out = out & "{": i = i + 1
            End Select
        Else
            out = out & Mid$(tmpl, i, 1): i = i + 1
        End If
    Loop
    FillPlaceholders = out
End Function

' The Headers setting as an array of "Name: value" with {key} filled in.
' Raises error 5 for a header with a line break or no colon.
Public Function AIHeaderList() As Variant
    Dim parts() As String, out() As String, i As Long, n As Long, h As String
    parts = Split(KitAISetting("Headers"), "|")
    ReDim out(0 To UBound(parts))
    For i = 0 To UBound(parts)
        h = Trim$(parts(i))
        If Len(h) > 0 Then
            If InStr(h, vbCr) > 0 Or InStr(h, vbLf) > 0 Or InStr(h, ":") < 2 Then _
                Err.Raise 5, "modKitAI", "An AI service header is not ""Name: value"": " & Left$(h, 30)
            If InStr(h, "{key}") > 0 Then
                If InStr(AIKey(), vbCr) > 0 Or InStr(AIKey(), vbLf) > 0 Then Err.Raise 5, "modKitAI", "The API key has a line break in it."
                h = Replace(h, "{key}", AIKey())
            End If
            out(n) = h: n = n + 1
        End If
    Next
    If n = 0 Then
        AIHeaderList = Array()
    Else
        ReDim Preserve out(0 To n - 1)
        AIHeaderList = out
    End If
End Function

' POST body and return the response text; status receives the HTTP status.
' Raises error 3000 for a network failure or timeout.
Private Function AIPost(ByVal url As String, headers As Variant, ByVal body As String, ByRef status As Long) As String
    Dim http As Object, h As Variant, k As Long, t As Long
    t = Val(KitAISetting("TimeoutSec"))
    If t < 5 Then t = 60
    On Error GoTo Fail
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 10000, 10000, 30000, t * 1000
    http.Open "POST", url, False
    For Each h In headers
        k = InStr(h, ":")
        http.setRequestHeader Trim$(Left$(h, k - 1)), Trim$(Mid$(h, k + 1))
    Next
    http.send body
    status = http.status
    AIPost = Utf8Text(http.responseBody)
    Exit Function
Fail:
    Err.Raise 3000, "modKitAI", "The AI service could not be reached (" & Err.Description & ")."
End Function

' Response bytes as text (the services answer in UTF-8).
Private Function Utf8Text(bytes As Variant) As String
    Dim st As Object
    If IsEmpty(bytes) Or IsNull(bytes) Then Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open: st.Write bytes
    st.Position = 0: st.Type = 2: st.Charset = "utf-8"
    Utf8Text = st.ReadText
    st.Close
End Function

' The reply text in a response: the ResponsePath value. Raises error 3000,
' with the service's own error message when it sent one, for a non-2xx status
' or a response without the text.
Public Function AIReplyText(ByVal response As String, ByVal status As Long) As String
    Dim v As Variant, msg As Variant
    v = JsonPath(response, KitAISetting("ResponsePath"))
    If status >= 200 And status < 300 Then
        If VarType(v) = vbString Then
            If Len(v) > 0 Then AIReplyText = v: Exit Function
        End If
        Err.Raise 3000, "modKitAI", "The AI service answered, but the reply had no text at " & KitAISetting("ResponsePath") & "."
    End If
    msg = JsonPath(response, "error.message")
    If VarType(msg) <> vbString Then msg = JsonPath(response, "error")
    If VarType(msg) <> vbString Then msg = JsonPath(response, "message")
    If VarType(msg) = vbString Then
        Err.Raise 3000, "modKitAI", "The AI service refused the request (" & status & "): " & Left$(msg, 300)
    End If
    Err.Raise 3000, "modKitAI", "The AI service refused the request (" & status & ")."
End Function

' ---------------------------------------------------------------------------
' Connecting controls
' ---------------------------------------------------------------------------

' The bound controls as "form|control|role" entries (an array of Strings).
Public Function KitAIBindings() As Variant
    Dim s As String
    s = GetSetting(APP_NAME, AI_SETTINGS_KEY, "Bindings", "")
    If Len(s) = 0 Then KitAIBindings = Array() Else KitAIBindings = Split(s, ";")
End Function

' Connect a control to the AI: role is input, context or output. A control has
' one role (binding it again changes it); a form can have several inputs and
' contexts and one output, and a second output replaces the first. Raises
' error 5 for a bad role or a name with "|" or ";".
Public Sub KitAIBind(ByVal formName As String, ByVal controlName As String, ByVal role As String)
    Dim list As Variant, e As Variant, p() As String, s As String
    role = LCase$(Trim$(role))
    If role <> AI_ROLE_INPUT And role <> AI_ROLE_CONTEXT And role <> AI_ROLE_OUTPUT Then _
        Err.Raise 5, "modKitAI", "An AI role is input, context or output: " & role
    If Len(formName) = 0 Or Len(controlName) = 0 Or InStr(formName & controlName, "|") > 0 Or InStr(formName & controlName, ";") > 0 Then _
        Err.Raise 5, "modKitAI", "Not a form and control name: " & formName & "." & controlName
    list = KitAIBindings()
    For Each e In list
        p = Split(e, "|")
        If Not (p(0) = formName And p(1) = controlName) Then
            If Not (role = AI_ROLE_OUTPUT And p(0) = formName And p(2) = AI_ROLE_OUTPUT) Then _
                s = s & IIf(Len(s) > 0, ";", "") & e
        End If
    Next
    s = s & IIf(Len(s) > 0, ";", "") & formName & "|" & controlName & "|" & role
    SaveSetting APP_NAME, AI_SETTINGS_KEY, "Bindings", s
End Sub

' Disconnect a control. True when it was connected.
Public Function KitAIUnbind(ByVal formName As String, ByVal controlName As String) As Boolean
    Dim e As Variant, p() As String, s As String
    For Each e In KitAIBindings()
        p = Split(e, "|")
        If p(0) = formName And p(1) = controlName Then
            KitAIUnbind = True
        Else
            s = s & IIf(Len(s) > 0, ";", "") & e
        End If
    Next
    If KitAIUnbind Then SaveSetting APP_NAME, AI_SETTINGS_KEY, "Bindings", s
End Function

' The role of a control, or "" when it isn't AI-connected.
Public Function KitAIRole(ByVal formName As String, ByVal controlName As String) As String
    Dim e As Variant, p() As String
    For Each e In KitAIBindings()
        p = Split(e, "|")
        If p(0) = formName And p(1) = controlName Then KitAIRole = p(2): Exit Function
    Next
End Function

' The names of a form's controls with a role (an array of Strings, possibly empty).
Public Function KitAIControls(ByVal formName As String, ByVal role As String) As Variant
    Dim e As Variant, p() As String, out() As String, n As Long
    ReDim out(0 To UBound(KitAIBindings()) + 1)
    For Each e In KitAIBindings()
        p = Split(e, "|")
        If p(0) = formName And p(2) = role Then out(n) = p(1): n = n + 1
    Next
    If n = 0 Then KitAIControls = Array() Else ReDim Preserve out(0 To n - 1): KitAIControls = out
End Function

' Outline the form's AI-connected controls with the accent color (CLR_PRIMARY),
' so users can see what is connected; others are left as they were. Call it
' from the form's Load, after ThemeForm. Controls not on the form are skipped.
Public Sub KitAIMark(frm As Form)
    Dim c As Control, role As String
    On Error GoTo Fail
    For Each c In frm.Controls
        role = KitAIRole(frm.Name, c.Name)
        If Len(role) > 0 Then
            On Error Resume Next
            c.BorderStyle = 1: c.BorderWidth = 2: c.BorderColor = Th(CLR_PRIMARY)
            c.ControlTipText = "Connected to the AI (" & role & ")"
            On Error GoTo Fail
        End If
    Next
    Exit Sub
Fail:
    LogError "modKitAI.KitAIMark", "form " & frm.Name
End Sub

' Let the user choose which controls on frm are AI-connected: a menu under
' anchor lists the form's text boxes, combo boxes and labels (a mark shows the
' connected ones); picking one opens a second menu for its role, or to remove
' it. Repeats until the user closes the first menu. True when anything changed.
Public Function KitAIConnectControls(frm As Form, anchor As Control) As Boolean
    Dim c As Control, items() As Variant, n As Long, pick As String, role As String, cur As String
    On Error GoTo Fail
    Do
        n = 0
        ReDim items(0 To 0)
        For Each c In frm.Controls
            If AICapable(c) Then
                If n >= MENU_ROWS Then Exit For
                ReDim Preserve items(0 To n)
                cur = KitAIRole(frm.Name, c.Name)
                items(n) = c.Name & "|" & c.Name & IIf(Len(cur) > 0, "  (" & cur & ")", "") & IIf(Len(cur) > 0, "|E73E", "")
                n = n + 1
            End If
        Next
        If n = 0 Then Exit Do
        pick = KitMenuList(anchor, items)
        If Len(pick) = 0 Then Exit Do
        role = KitMenu(anchor, "input|Message (sent to the AI)", "context|Context (added to the message)", _
            "output|Reply goes here", "-", "none|!Not connected")
        If Len(role) = 0 Then
        ElseIf role = "none" Then
            If KitAIUnbind(frm.Name, pick) Then KitAIConnectControls = True
        Else
            KitAIBind frm.Name, pick, role
            KitAIConnectControls = True
        End If
    Loop
    If KitAIConnectControls Then KitAIMark frm
    Exit Function
Fail:
    HandleError "modKitAI.KitAIConnectControls", "form " & frm.Name, "The AI connections couldn't be changed."
End Function

' KitMenu with an array of items.
Private Function KitMenuList(anchor As Control, items() As Variant) As String
    Dim list As Variant
    list = items
    If Not MenuOpen(anchor, list) Then Exit Function
    Do While IsFormOpen("dlgMenu")
        DoEvents
        Sleep 15
    Loop
    KitMenuList = MenuResult()
End Function

' Controls that can carry AI text.
Private Function AICapable(c As Control) As Boolean
    Select Case c.ControlType
        Case acTextBox, acComboBox, acLabel
            AICapable = (c.Name Like "[A-Za-z]*") And Not (c.Name Like "nb_*" Or c.Name Like "ni_*" Or c.Name Like "nt_*")
    End Select
End Function

' Send frm's AI input to the service and write the reply into its AI output.
' The message is the first non-empty input, with each context control added
' above it as "Caption/Name: value". A busy message shows while it waits. Does
' nothing, with a toast saying why, when the form has no input or output
' connected. Returns the reply, or "" when nothing came back (a message
' explained why).
Public Function KitAIAsk(frm As Form) As String
    Dim inputs As Variant, ctx As Variant, outs As Variant, n As Variant
    Dim prompt As String, msg As String, reply As String, busy As Boolean
    On Error GoTo Fail
    inputs = KitAIControls(frm.Name, AI_ROLE_INPUT)
    outs = KitAIControls(frm.Name, AI_ROLE_OUTPUT)
    If UBound(inputs) < 0 Or UBound(outs) < 0 Then
        KitToast "Choose an AI message box and a place for the reply first.", "warning"
        Exit Function
    End If
    For Each n In inputs
        msg = Trim$(ControlText(frm.Controls(n)))
        If Len(msg) > 0 Then Exit For
    Next
    If Len(msg) = 0 Then KitToast "Type a message for the AI first.", "warning": Exit Function
    ctx = KitAIControls(frm.Name, AI_ROLE_CONTEXT)
    For Each n In ctx
        If Len(ControlText(frm.Controls(n))) > 0 Then prompt = prompt & n & ": " & ControlText(frm.Controls(n)) & vbLf
    Next
    prompt = prompt & msg
    KitBusy "Asking the AI..."
    busy = True
    reply = KitAISend(prompt)
    KitBusyEnd
    busy = False
    SetControlText frm.Controls(outs(0)), reply
    KitAIAsk = reply
    Exit Function
Fail:
    Dim d As String
    d = Err.Description
    If busy Then KitBusyEnd
    KitMsgBox d, vbExclamation, "AI"
End Function

Private Function ControlText(c As Control) As String
    If c.ControlType = acLabel Then ControlText = c.Caption Else ControlText = Nz(c.Value, "")
End Function

Private Sub SetControlText(c As Control, ByVal s As String)
    If c.ControlType = acLabel Then c.Caption = s Else c.Value = s
End Sub

' ---------------------------------------------------------------------------
' JSON helpers
' ---------------------------------------------------------------------------

' s as a quoted JSON string. Everything outside printable ASCII is written as
' \uXXXX, so the request body is plain ASCII whatever the text.
Public Function JsonQuote(ByVal s As String) As String
    Dim i As Long, ch As Long, out As String
    out = """"
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1)) And &HFFFF&
        Select Case ch
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 10: out = out & "\n"
            Case 13: out = out & "\r"
            Case 9: out = out & "\t"
            Case 32 To 126: out = out & Chr$(ch)
            Case Else: out = out & "\u" & Right$("0000" & Hex$(ch), 4)
        End Select
    Next
    JsonQuote = out & """"
End Function

' The value at path in a JSON text: "a.b", "list[0].name". A string, a Double,
' a Boolean or Null for a value; a Dictionary or Collection for an object or
' array; Empty when the path isn't there or the text isn't JSON.
Public Function JsonPath(ByVal json As String, ByVal path As String) As Variant
    Dim cur As Variant, parts As Collection, p As Variant, idx As Long
    On Error GoTo Bad
    m_text = json: m_pos = 1
    JsonValue cur
    Set parts = PathParts(path)
    For Each p In parts
        If VarType(p) = vbLong Then
            If TypeName(cur) <> "Collection" Then Exit Function
            idx = p + 1
            If idx < 1 Or idx > cur.Count Then Exit Function
            AssignTo JsonPath, cur(idx)
        Else
            If TypeName(cur) <> "Dictionary" Then Exit Function
            If Not cur.Exists(p) Then Exit Function
            AssignTo JsonPath, cur(p)
        End If
        AssignTo cur, JsonPath
    Next
    If parts.Count = 0 Then AssignTo JsonPath, cur
    Exit Function
Bad:
    JsonPath = Empty
End Function

Private Sub AssignTo(ByRef dest As Variant, src As Variant)
    If IsObject(src) Then Set dest = src Else dest = src
End Sub

' "a.b[2].c" as keys (Strings) and indexes (Longs).
Private Function PathParts(ByVal path As String) As Collection
    Dim out As New Collection, i As Long, tok As String, ch As String, j As Long
    path = Trim$(path)
    i = 1
    Do While i <= Len(path)
        ch = Mid$(path, i, 1)
        Select Case ch
            Case ".", " "
                If Len(tok) > 0 Then out.Add tok: tok = ""
                i = i + 1
            Case "["
                If Len(tok) > 0 Then out.Add tok: tok = ""
                j = InStr(i, path, "]")
                If j = 0 Then Err.Raise 5, "modKitAI", "Bad reply path: " & path
                out.Add CLng(Val(Mid$(path, i + 1, j - i - 1)))
                i = j + 1
            Case Else
                tok = tok & ch: i = i + 1
        End Select
    Loop
    If Len(tok) > 0 Then out.Add tok
    Set PathParts = out
End Function

' --- a small JSON reader (m_text from m_pos) ---

Private Sub JsonSkip()
    Do While m_pos <= Len(m_text)
        Select Case Mid$(m_text, m_pos, 1)
            Case " ", vbTab, vbCr, vbLf: m_pos = m_pos + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

Private Sub JsonValue(ByRef out As Variant)
    Dim ch As String
    JsonSkip
    If m_pos > Len(m_text) Then Err.Raise 5, "modKitAI", "Not JSON"
    ch = Mid$(m_text, m_pos, 1)
    Select Case ch
        Case "{": JsonObject out
        Case "[": JsonArray out
        Case """": out = JsonString()
        Case "t": JsonWord "true": out = True
        Case "f": JsonWord "false": out = False
        Case "n": JsonWord "null": out = Null
        Case Else: out = JsonNumber()
    End Select
End Sub

Private Sub JsonWord(ByVal w As String)
    If Mid$(m_text, m_pos, Len(w)) <> w Then Err.Raise 5, "modKitAI", "Not JSON"
    m_pos = m_pos + Len(w)
End Sub

Private Function JsonNumber() As Double
    Dim st As Long
    st = m_pos
    Do While m_pos <= Len(m_text)
        If InStr("0123456789+-.eE", Mid$(m_text, m_pos, 1)) = 0 Then Exit Do
        m_pos = m_pos + 1
    Loop
    If m_pos = st Then Err.Raise 5, "modKitAI", "Not JSON"
    JsonNumber = Val(Mid$(m_text, st, m_pos - st))
End Function

Private Function JsonString() As String
    Dim out As String, ch As String, esc As String
    m_pos = m_pos + 1                       ' the opening quote
    Do While m_pos <= Len(m_text)
        ch = Mid$(m_text, m_pos, 1)
        If ch = """" Then m_pos = m_pos + 1: JsonString = out: Exit Function
        If ch = "\" Then
            esc = Mid$(m_text, m_pos + 1, 1)
            m_pos = m_pos + 2
            Select Case esc
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case "t": out = out & vbTab
                Case "b": out = out & vbBack
                Case "f": out = out & Chr$(12)
                Case "u"
                    out = out & ChrW(CLng("&H" & Mid$(m_text, m_pos, 4)))
                    m_pos = m_pos + 4
                Case Else: out = out & esc    ' " \ /
            End Select
        Else
            out = out & ch: m_pos = m_pos + 1
        End If
    Loop
    Err.Raise 5, "modKitAI", "Not JSON"
End Function

Private Sub JsonObject(ByRef out As Variant)
    Dim d As Object, key As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    m_pos = m_pos + 1
    JsonSkip
    If Mid$(m_text, m_pos, 1) = "}" Then m_pos = m_pos + 1: Set out = d: Exit Sub
    Do
        JsonSkip
        If Mid$(m_text, m_pos, 1) <> """" Then Err.Raise 5, "modKitAI", "Not JSON"
        key = JsonString()
        JsonSkip
        If Mid$(m_text, m_pos, 1) <> ":" Then Err.Raise 5, "modKitAI", "Not JSON"
        m_pos = m_pos + 1
        JsonValue v
        If IsObject(v) Then Set d(key) = v Else d(key) = v
        JsonSkip
        Select Case Mid$(m_text, m_pos, 1)
            Case ",": m_pos = m_pos + 1
            Case "}": m_pos = m_pos + 1: Set out = d: Exit Sub
            Case Else: Err.Raise 5, "modKitAI", "Not JSON"
        End Select
    Loop
End Sub

Private Sub JsonArray(ByRef out As Variant)
    Dim a As New Collection, v As Variant
    m_pos = m_pos + 1
    JsonSkip
    If Mid$(m_text, m_pos, 1) = "]" Then m_pos = m_pos + 1: Set out = a: Exit Sub
    Do
        JsonValue v
        a.Add v
        JsonSkip
        Select Case Mid$(m_text, m_pos, 1)
            Case ",": m_pos = m_pos + 1
            Case "]": m_pos = m_pos + 1: Set out = a: Exit Sub
            Case Else: Err.Raise 5, "modKitAI", "Not JSON"
        End Select
    Loop
End Sub
