Option Compare Database
Option Explicit

' Access UI kit - a small JSON reader and writer, with no references to add
' (Scripting.Dictionary is created late, as the other kit modules do).
'
'   body = "{""model"":" & JsonQuote(model) & "}"        ' write: quote text
'   root = JsonParse(reply)                              ' read
'   text = JsonText(root, "choices.0.message.content")   ' by path, "" if absent
'
' JsonParse returns an object as a Scripting.Dictionary, an array as a
' Collection (1-based, as Collections are), and text, numbers (Double) and
' booleans as themselves; a JSON null is Null. A syntax error raises error 5
' naming the position. JsonGet and JsonText take a path of names and array
' positions split by dots; array positions start at 0 ("choices.0.message").
' JsonQuote writes ASCII only (anything else as \uXXXX), so a request body
' is the same in any code page.
' The parser keeps its place in module variables, so it isn't reentrant
' (nothing in it calls back out).

Private m_t As String      ' the text being read
Private m_p As Long        ' the next character
Private m_bs As Long       ' position of the next backslash at or after m_p: 0 none, -1 not looked for

' ---------------------------------------------------------------------------
' Writing
' ---------------------------------------------------------------------------

' s as a JSON string literal, quotes included.
Public Function JsonQuote(ByVal s As String) As String
    Dim i As Long, c As Long, run As Long, buf As String, esc As String
    run = 1
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        esc = ""
        Select Case c
            Case 34: esc = "\"""
            Case 92: esc = "\\"
            Case 8: esc = "\b"
            Case 9: esc = "\t"
            Case 10: esc = "\n"
            Case 12: esc = "\f"
            Case 13: esc = "\r"
            Case Is < 32, Is > 126: esc = "\u" & Right$("0000" & Hex$(c), 4)
        End Select
        If Len(esc) > 0 Then
            buf = buf & Mid$(s, run, i - run) & esc
            run = i + 1
        End If
    Next
    JsonQuote = """" & buf & Mid$(s, run) & """"
End Function

' ---------------------------------------------------------------------------
' Reading
' ---------------------------------------------------------------------------

' The value JSON text holds (see the top of the module).
Public Function JsonParse(ByVal text As String) As Variant
    Dim v As Variant
    m_t = text: m_p = 1: m_bs = -1
    On Error GoTo Fail
    ParseValue v
    SkipWs
    If m_p <= Len(m_t) Then JsonFail "Unexpected text after the value"
    m_t = ""
    If IsObject(v) Then Set JsonParse = v Else JsonParse = v
    Exit Function
Fail:
    Dim num As Long, desc As String
    num = Err.Number: desc = Err.Description
    m_t = ""
    Err.Raise num, "modKitJson.JsonParse", desc
End Function

' The value at path inside root (a parsed object or array), or Empty when any
' step is missing.
Public Function JsonGet(root As Variant, ByVal path As String) As Variant
    Dim node As Variant, part As Variant, idx As Long
    If IsObject(root) Then Set node = root Else node = root
    For Each part In Split(path, ".")
        If Not IsObject(node) Then JsonGet = Empty: Exit Function
        If TypeName(node) = "Collection" Then
            If Not IsNumeric(part) Then JsonGet = Empty: Exit Function
            idx = CLng(part) + 1
            If idx < 1 Or idx > node.Count Then JsonGet = Empty: Exit Function
            If IsObject(node(idx)) Then Set node = node(idx) Else node = node(idx)
        Else
            If Not node.Exists(CStr(part)) Then JsonGet = Empty: Exit Function
            If IsObject(node(CStr(part))) Then Set node = node(CStr(part)) Else node = node(CStr(part))
        End If
    Next
    If IsObject(node) Then Set JsonGet = node Else JsonGet = node
End Function

' The text at path, or dflt when it is missing, null or not a simple value.
Public Function JsonText(root As Variant, ByVal path As String, Optional ByVal dflt As String = "") As String
    Dim v As Variant
    v = JsonGet(root, path)
    If IsObject(v) Or IsEmpty(v) Or IsNull(v) Then JsonText = dflt Else JsonText = CStr(v)
End Function

' How many items the array at path holds (path "" for root itself); 0 when it
' is missing or isn't an array.
Public Function JsonCount(root As Variant, ByVal path As String) As Long
    Dim v As Variant
    If Len(path) = 0 Then v = root Else v = JsonGet(root, path)
    If IsObject(v) Then
        If TypeName(v) = "Collection" Then JsonCount = v.Count
    End If
End Function

' ---------------------------------------------------------------------------
' The parser
' ---------------------------------------------------------------------------

Private Sub JsonFail(ByVal why As String)
    Err.Raise 5, "modKitJson", why & " (JSON position " & m_p & ")"
End Sub

Private Sub SkipWs()
    Dim c As String
    Do While m_p <= Len(m_t)
        c = Mid$(m_t, m_p, 1)
        If c = " " Or c = vbTab Or c = vbLf Or c = vbCr Then m_p = m_p + 1 Else Exit Do
    Loop
End Sub

Private Sub ParseValue(ByRef v As Variant)
    SkipWs
    Select Case Mid$(m_t, m_p, 1)
        Case "{": Set v = ParseObject()
        Case "[": Set v = ParseArray()
        Case """": v = ParseString()
        Case "t": ExpectWord "true": v = True
        Case "f": ExpectWord "false": v = False
        Case "n": ExpectWord "null": v = Null
        Case "-", "0" To "9": v = ParseNumber()
        Case "": JsonFail "The text ended early"
        Case Else: JsonFail "Unexpected character " & Mid$(m_t, m_p, 1)
    End Select
End Sub

Private Sub ExpectWord(ByVal w As String)
    If Mid$(m_t, m_p, Len(w)) <> w Then JsonFail "Expected " & w
    m_p = m_p + Len(w)
End Sub

Private Function ParseObject() As Object
    Dim d As Object, k As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    m_p = m_p + 1
    SkipWs
    If Mid$(m_t, m_p, 1) = "}" Then m_p = m_p + 1: Set ParseObject = d: Exit Function
    Do
        SkipWs
        If Mid$(m_t, m_p, 1) <> """" Then JsonFail "Expected a name in quotes"
        k = ParseString()
        SkipWs
        If Mid$(m_t, m_p, 1) <> ":" Then JsonFail "Expected :"
        m_p = m_p + 1
        v = Empty
        ParseValue v
        If IsObject(v) Then Set d(k) = v Else d(k) = v
        SkipWs
        Select Case Mid$(m_t, m_p, 1)
            Case ",": m_p = m_p + 1
            Case "}": m_p = m_p + 1: Set ParseObject = d: Exit Function
            Case Else: JsonFail "Expected , or }"
        End Select
    Loop
End Function

Private Function ParseArray() As Collection
    Dim a As Collection, v As Variant
    Set a = New Collection
    m_p = m_p + 1
    SkipWs
    If Mid$(m_t, m_p, 1) = "]" Then m_p = m_p + 1: Set ParseArray = a: Exit Function
    Do
        v = Empty
        ParseValue v
        a.Add v
        SkipWs
        Select Case Mid$(m_t, m_p, 1)
            Case ",": m_p = m_p + 1
            Case "]": m_p = m_p + 1: Set ParseArray = a: Exit Function
            Case Else: JsonFail "Expected , or ]"
        End Select
    Loop
End Function

' The string at m_p (on its opening quote), with escapes resolved. Plain runs
' are copied whole, so a long reply isn't built a character at a time.
Private Function ParseString() As String
    Dim q As Long, b As Long, buf As String, c As String, hx As String
    m_p = m_p + 1
    Do
        q = InStr(m_p, m_t, """")
        If q = 0 Then JsonFail "A string is not closed"
        If m_bs <> 0 And m_bs < m_p Then m_bs = InStr(m_p, m_t, "\")
        b = m_bs
        If b = 0 Or b > q Then
            ParseString = buf & Mid$(m_t, m_p, q - m_p)
            m_p = q + 1
            Exit Function
        End If
        buf = buf & Mid$(m_t, m_p, b - m_p)
        c = Mid$(m_t, b + 1, 1)
        m_p = b + 2
        Select Case c
            Case """", "\", "/": buf = buf & c
            Case "b": buf = buf & vbBack
            Case "f": buf = buf & Chr$(12)
            Case "n": buf = buf & vbLf
            Case "r": buf = buf & vbCr
            Case "t": buf = buf & vbTab
            Case "u"
                hx = Mid$(m_t, m_p, 4)
                If Len(hx) < 4 Then JsonFail "A \u escape is cut short"
                If Not (hx Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]") Then JsonFail "A \u escape is not four hex digits"
                buf = buf & ChrW(CLng("&H" & hx))
                m_p = m_p + 4
            Case Else: JsonFail "Unknown escape \" & c
        End Select
    Loop
End Function

Private Function ParseNumber() As Double
    Dim s As Long, c As String
    s = m_p
    Do While m_p <= Len(m_t)
        c = Mid$(m_t, m_p, 1)
        If c Like "[0-9.eE+-]" Then m_p = m_p + 1 Else Exit Do
    Loop
    ParseNumber = Val(Mid$(m_t, s, m_p - s))
End Function
