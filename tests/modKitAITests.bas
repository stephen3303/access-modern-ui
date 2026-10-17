Option Compare Database
Option Explicit

' Tests for modKitAI that need no network or key: request and reply adapters for
' both formats, JSON quoting and decoding, error replies, conversation windowing
' and configuration. TestKit calls TestKitAI and adds the count to its own.

Private n As Long

Private Sub Ok(ByVal valid As Boolean, ByVal message As String)
    If Not valid Then Err.Raise 5, , message
    n = n + 1
End Sub

' Message number errNum raised by calling AIParseReply, else 0.
Private Function ParseErr(ByVal fmt As String, ByVal json As String) As Long
    On Error Resume Next
    AIParseReply fmt, json
    ParseErr = Err.Number
End Function

Public Function TestKitAI() As Long
    Dim h As Collection, raw As String, i As Long, bound As Long
    n = 0
    ' JSON quoting: escapes, control characters and non-ASCII as \u, round trip.
    Ok JsonQuote("a""b\c") = """a\""b\\c""", "JsonQuote did not escape quote and backslash"
    Ok JsonQuote("x" & vbCrLf & vbTab) = """x\r\n\t""", "JsonQuote did not escape control characters"
    Ok JsonQuote("caf" & ChrW$(233)) = """café""", "JsonQuote left a non-ASCII character"
    Ok JsonQuote(ChrW$(&H4E2D)) = """中""", "JsonQuote broke a character above &H7FFF"
    ' Requests.
    Set h = New Collection
    h.Add Array("user", "Hi ""there""")
    raw = AIBuildBody(AI_ANTHROPIC, "m1", "Be brief.", 0, h)
    Ok raw = "{""model"":""m1"",""max_tokens"":1024,""system"":""Be brief."",""messages"":[{""role"":""user"",""content"":""Hi \""there\""""}]}", _
        "Anthropic request wrong: " & raw
    raw = AIBuildBody(AI_OPENAI, "m2", "Be brief.", 0, h)
    Ok raw = "{""model"":""m2"",""messages"":[{""role"":""system"",""content"":""Be brief.""},{""role"":""user"",""content"":""Hi \""there\""""}]}", _
        "OpenAI request wrong: " & raw
    raw = AIBuildBody(AI_OPENAI, "m2", "", 200, h)
    Ok InStr(raw, """max_tokens"":200") > 0 And InStr(raw, """system""") = 0, "OpenAI request with no system prompt wrong: " & raw
    h.Add Array("assistant", "Hello")
    h.Add Array("user", "Again")
    raw = AIBuildBody(AI_ANTHROPIC, "m1", "", 50, h)
    Ok InStr(raw, "{""role"":""assistant"",""content"":""Hello""},{""role"":""user""") > 0, "Conversation turns not kept in order: " & raw
    ' Replies.
    Ok AIParseReply(AI_ANTHROPIC, "{""id"":""x"",""type"":""message"",""content"":[{""type"":""text"",""text"":""One\nTwo ""},{""type"":""text"",""text"":""é!""}],""usage"":{""output_tokens"":3}}") _
        = "One" & vbLf & "Two " & ChrW$(233) & "!", "Anthropic reply misread"
    Ok AIParseReply(AI_OPENAI, "{""choices"":[{""index"":0,""message"":{""role"":""assistant"",""content"":""Say \""hi\"" \\ ok""},""finish_reason"":""stop""}]}") _
        = "Say ""hi"" \ ok", "OpenAI reply misread"
    Ok AIParseReply(AI_OPENAI, "{ ""choices"" : [ { ""message"" : { ""content"" : ""spaced"" } } ] }") = "spaced", "Spaced JSON misread"
    ' A reply that talks about "error" and "content" is still just text.
    Ok AIParseReply(AI_OPENAI, "{""choices"":[{""message"":{""content"":""\""error\"": \""no\""""}}]}") = """error"": ""no""", "Text inside a reply read as a key"
    ' Errors.
    Ok AIParseError("{""type"":""error"",""error"":{""type"":""authentication_error"",""message"":""invalid x-api-key""}}") = "invalid x-api-key", "Anthropic error message misread"
    Ok AIParseError("{""error"":""plain text error""}") = "plain text error", "String error misread"
    Ok AIParseError("{""choices"":[]}") = "", "Error found in a good reply"
    Ok ParseErr(AI_OPENAI, "{""error"":{""message"":""Rate limit""}}") = AI_ERR_SERVICE, "Error reply not raised as AI_ERR_SERVICE"
    Ok ParseErr(AI_ANTHROPIC, "{""content"":[]}") = AI_ERR_BAD_REPLY, "Empty reply not raised as AI_ERR_BAD_REPLY"
    ' Configuration.
    On Error Resume Next
    AIConfigure "ollama", "llama"
    Ok Err.Number = AI_ERR_NOT_CONFIGURED, "A custom service with no endpoint was accepted"
    Err.Clear
    AIConfigure "anthropic", "m", fmt:="soap"
    Ok Err.Number = AI_ERR_NOT_CONFIGURED, "An unknown request format was accepted"
    Err.Clear
    On Error GoTo 0
    AIConfigure "ollama", "llama", "http://localhost:11434/v1/chat/completions", keyEnv:="KIT_AI_TEST_KEY_NOT_SET"
    Ok AIConfigured(), "Configured service not reported"
    Ok Not AIHasKey(), "A key was found that nobody set"
    AISetKey " abc "
    Ok AIHasKey(), "AISetKey not kept"
    AIConfigure "other", "m", "http://x/y"
    Ok Not AIHasKey(), "A key carried over to another service"
    ' A long conversation is cut to AIMaxMessages, starting on a user turn.
    AIMaxMessages 3
    Set h = New Collection
    For i = 1 To 6
        h.Add Array(IIf(i Mod 2 = 1, "user", "assistant"), "m" & i)
    Next
    raw = AIBuildBody(AI_OPENAI, "m", "", 0, h)
    Ok InStr(raw, "m4") = 0 And InStr(raw, "m5") > 0 And InStr(raw, """m6""") > 0, "Windowed conversation wrong: " & raw
    AIMaxMessages 20
    TestKitAI = n
End Function
