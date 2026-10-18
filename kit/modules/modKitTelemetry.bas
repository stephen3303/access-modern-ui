Option Compare Database
Option Explicit

' Access UI kit - optional error reporting to Sentry, so a site's admin sees
' errors from every copy of the app in one place, grouped, with which build
' they started in.
'
' Off unless a Personnel admin turns it on for the site with
' KitTelemetryTurnOn, which also asks for the site's classification marking
' and refuses any marking but UNCLASSIFIED. Every send checks the saved
' marking again. Never use it on SIPRNet or any classified network: leave this
' module out of that build altogether (LogError reaches it through
' Application.Run, so an app without it builds and runs the same).
'
' What one report holds: the app's name, the error number, the procedure that
' caught it (module.procedure and line), the page key, the front end's build,
' the Access version, the time, and a random id for the install (made on first
' use and kept in the registry). Never the error's description (an Access
' message can quote field values), the context, the call trail, user names,
' the Windows account or the computer name. A procedure name that carries more
' than "module.procedure line N" is cut back to that.
'
' A report never stops or slows the user: it is sent in the background with a
' short timeout, nothing is shown when it fails, and nothing is retried. Each
' session sends at most TELEMETRY_SESSION_MAX reports and repeats none (the same
' error number in the same procedure is sent once a session).
'
' Settings live in kit_setting, a table in the app's shared back end, so one
' admin turns reporting on for every user. At startup, after KitEnsureBackend,
' the app calls KitEnsureSettingTable with one of its linked tables; it adds
' the table to that data file when it lacks it and links it into the front end:
'   telemetry_dsn       the Sentry project's DSN (a send-only client key)
'   telemetry_marking   the marking the admin confirmed (UNCLASSIFIED)
'
' The app provides APP_NAME and FrontEndBuild (modErrors); modErrors.LogError
' calls KitTelemetryReport after it writes each entry.

Public Const KIT_SETTING_TABLE As String = "kit_setting"
Public Const TELEMETRY_DSN_KEY As String = "telemetry_dsn"
Public Const TELEMETRY_MARKING_KEY As String = "telemetry_marking"
Public Const TELEMETRY_MARKING As String = "UNCLASSIFIED"
Public Const TELEMETRY_SESSION_MAX As Long = 20

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef st As SYSTEMTIME)

Private Const TIMEOUT_MS As Long = 4000
Private Const MAX_PENDING As Long = 5

Private m_sent As Object            ' fingerprints reported this session
Private m_pending As Collection     ' requests still in flight (kept alive until they finish)

' Build or upgrade step: add kit_setting to db (default the current database;
' pass the back end's for a split app) when it lacks it. True when it was added.
Public Function KitAddSettingTable(Optional db As DAO.Database) As Boolean
    Dim td As DAO.TableDef
    On Error GoTo Fail
    If db Is Nothing Then Set db = CurrentDb
    db.TableDefs.Refresh
    For Each td In db.TableDefs
        If StrComp(td.Name, KIT_SETTING_TABLE, vbTextCompare) = 0 Then Exit Function
    Next
    db.Execute "CREATE TABLE " & KIT_SETTING_TABLE & " (setting_key TEXT(64) CONSTRAINT pk_kit_setting PRIMARY KEY, " & _
        "setting_value TEXT(255))", dbFailOnError
    KitAddSettingTable = True
    Exit Function
Fail:
    RaiseAgain "modKitTelemetry.KitAddSettingTable"
End Function

' Startup step for a split app: make sure kit_setting is in the data file
' probeTable is linked to, and linked into this front end the same way
' (password included). A front end whose probeTable is a local table keeps
' kit_setting locally. True when the table was added or linked.
Public Function KitEnsureSettingTable(ByVal probeTable As String) As Boolean
    Dim db As DAO.Database, cn As String, td As DAO.TableDef, be As DAO.Database
    On Error GoTo Fail
    Set db = CurrentDb
    For Each td In db.TableDefs
        If StrComp(td.Name, KIT_SETTING_TABLE, vbTextCompare) = 0 Then Exit Function
    Next
    cn = db.TableDefs(probeTable).Connect
    If Len(cn) = 0 Then
        KitEnsureSettingTable = KitAddSettingTable(db)
        Exit Function
    End If
    Set be = DBEngine.OpenDatabase(KitBackendPath(probeTable), False, False, ConnectPassword(cn))
    KitAddSettingTable be
    be.Close
    Set td = db.CreateTableDef(KIT_SETTING_TABLE)
    td.Connect = cn
    td.SourceTableName = KIT_SETTING_TABLE
    db.TableDefs.Append td
    db.TableDefs.Refresh
    KitEnsureSettingTable = True
    Exit Function
Fail:
    RaiseAgain "modKitTelemetry.KitEnsureSettingTable", probeTable
End Function

' A kit_setting value, or "" when it is unset or the table can't be read.
Public Function KitSetting(ByVal key As String) As String
    On Error Resume Next
    KitSetting = Nz(DLookup("setting_value", KIT_SETTING_TABLE, "setting_key=" & SqlQuote(key)), "")
End Function

' Save a kit_setting value; "" removes it.
Public Sub SetKitSetting(ByVal key As String, ByVal value As String)
    On Error GoTo Fail
    CurrentDb.Execute "DELETE FROM " & KIT_SETTING_TABLE & " WHERE setting_key=" & SqlQuote(key), dbFailOnError
    If Len(value) > 0 Then
        CurrentDb.Execute "INSERT INTO " & KIT_SETTING_TABLE & " (setting_key, setting_value) VALUES (" & _
            SqlQuote(key) & ", " & SqlQuote(Left$(value, 255)) & ")", dbFailOnError
    End If
    Exit Sub
Fail:
    RaiseAgain "modKitTelemetry.SetKitSetting", key
End Sub

' Turn reporting on for the site. Only an active Personnel admin can (error 5
' otherwise). marking is the site's classification as the admin confirmed it;
' anything but UNCLASSIFIED is refused. Returns "" when reporting is on, or
' why it was refused (a malformed DSN, the marking).
Public Function KitTelemetryTurnOn(ByVal dsn As String, ByVal marking As String) As String
    On Error GoTo Fail
    PersonnelRequireAdmin
    dsn = Trim$(dsn)
    If StrComp(Trim$(marking), TELEMETRY_MARKING, vbTextCompare) <> 0 Then
        KitTelemetryTurnOn = "Error reporting is only for unclassified sites. It stays off."
        Exit Function
    End If
    If Len(TelemetryUrl(dsn)) = 0 Then
        KitTelemetryTurnOn = "That isn't a Sentry DSN. Copy it from the Sentry project's Client Keys page; it starts with https://."
        Exit Function
    End If
    SetKitSetting TELEMETRY_MARKING_KEY, TELEMETRY_MARKING
    SetKitSetting TELEMETRY_DSN_KEY, dsn
    Exit Function
Fail:
    RaiseAgain "modKitTelemetry.KitTelemetryTurnOn"
End Function

' Turn reporting off for the site. Only an active Personnel admin can.
Public Sub KitTelemetryTurnOff()
    On Error GoTo Fail
    PersonnelRequireAdmin
    SetKitSetting TELEMETRY_DSN_KEY, ""
    SetKitSetting TELEMETRY_MARKING_KEY, ""
    Exit Sub
Fail:
    RaiseAgain "modKitTelemetry.KitTelemetryTurnOff"
End Sub

' True when reporting is on: a valid DSN and the UNCLASSIFIED marking saved.
Public Function KitTelemetryIsOn() As Boolean
    KitTelemetryIsOn = (KitSetting(TELEMETRY_MARKING_KEY) = TELEMETRY_MARKING) And _
        Len(TelemetryUrl(KitSetting(TELEMETRY_DSN_KEY))) > 0
End Function

' What a report holds, for the screen where an admin turns it on.
Public Function KitTelemetryFields() As String
    KitTelemetryFields = "Each error report holds only: the app's name, the error number, the procedure it happened in, " & _
        "the page, the app's build, the Access version, the time and a random id for this copy of the app. " & _
        "Never the error message, the record it was working on, names, Windows accounts or computer names."
End Function

' Report one error. Called by LogError after it writes the entry; never raises
' and never shows anything. True when a report was started.
Public Function KitTelemetryReport(ByVal errNumber As Long, ByVal source As String) As Boolean
    Dim url As String, dsn As String, proc As String, fp As String, req As Object
    On Error GoTo Fail
    If KitSetting(TELEMETRY_MARKING_KEY) <> TELEMETRY_MARKING Then Exit Function
    dsn = KitSetting(TELEMETRY_DSN_KEY)
    url = TelemetryUrl(dsn)
    If Len(url) = 0 Then Exit Function
    If m_sent Is Nothing Then Set m_sent = CreateObject("Scripting.Dictionary")
    If m_pending Is Nothing Then Set m_pending = New Collection
    proc = TelemetryName(source, True)
    fp = errNumber & "|" & StripLine(proc)
    If m_sent.Exists(fp) Or m_sent.Count >= TELEMETRY_SESSION_MAX Then Exit Function
    PruneFinished
    If m_pending.Count >= MAX_PENDING Then Exit Function    ' Sentry is slow: drop this one, a later one may go
    Set req = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    req.setTimeouts TIMEOUT_MS, TIMEOUT_MS, TIMEOUT_MS, TIMEOUT_MS
    req.Open "POST", url, True
    req.setRequestHeader "Content-Type", "application/x-sentry-envelope"
    req.setRequestHeader "X-Sentry-Auth", "Sentry sentry_version=7, sentry_key=" & DsnKey(dsn) & ", sentry_client=access-kit/1.0"
    req.send KitTelemetryEnvelope(errNumber, source, dsn)
    m_pending.Add req
    m_sent(fp) = True               ' only once it was started, so a dropped one can go later
    KitTelemetryReport = True
    Exit Function
Fail:
    ' Reporting is best effort: an error here must never reach the user or the log.
End Function

' The envelope for one report (Sentry's envelope format: a header line, an item
' header line, the event). Public so the tests can check what would be sent.
Public Function KitTelemetryEnvelope(ByVal errNumber As Long, ByVal source As String, ByVal dsn As String) As String
    Dim id As String, ts As String, app As String, proc As String, procNoLine As String
    Dim page As String, build As String, av As String, title As String, ev As String
    id = LCase$(RandomHex(16))
    ts = UtcStamp()
    app = TelemetryName(APP_NAME)
    proc = TelemetryName(source, True)
    procNoLine = StripLine(proc)
    page = TelemetryName(TelemetryPage())
    build = TelemetryName(TelemetryBuild())
    av = TelemetryName(Application.Version & "." & Application.Build)
    title = "Logged error " & errNumber & IIf(Len(proc) > 0, " in " & proc, "")
    ev = "{""event_id"":""" & id & """,""timestamp"":""" & ts & """,""platform"":""other"",""level"":""error""," & _
        """logger"":""access-kit"",""release"":" & JsonText(app & "@" & build) & "," & _
        """message"":{""formatted"":" & JsonText(title) & "},""tags"":{" & _
        """app"":" & JsonText(app) & ",""error_number"":""" & errNumber & """," & _
        """procedure"":" & JsonText(procNoLine) & ",""page"":" & JsonText(page) & "," & _
        """fe_build"":" & JsonText(build) & ",""access_version"":" & JsonText(av) & "," & _
        """install"":" & JsonText(InstallId()) & "}," & _
        """fingerprint"":[""access-kit""," & JsonText(app) & ",""" & errNumber & """," & JsonText(procNoLine) & "]}"
    KitTelemetryEnvelope = "{""event_id"":""" & id & """,""dsn"":" & JsonText(dsn) & ",""sent_at"":""" & ts & """}" & vbLf & _
        "{""type"":""event""}" & vbLf & ev & vbLf
End Function

' The envelope URL for a DSN (https://KEY[:SECRET]@HOST[:PORT][/PATH]/PROJECT),
' or "" when dsn isn't one. Public for the tests.
Public Function TelemetryUrl(ByVal dsn As String) As String
    Dim rest As String, at As Long, slash As Long, project As String, hostPath As String
    dsn = Trim$(dsn)
    If LCase$(Left$(dsn, 8)) <> "https://" Then Exit Function
    rest = Mid$(dsn, 9)
    at = InStr(rest, "@")
    If at < 2 Then Exit Function
    If Len(DsnKey(dsn)) = 0 Then Exit Function
    rest = Mid$(rest, at + 1)
    If Right$(rest, 1) = "/" Then rest = Left$(rest, Len(rest) - 1)
    slash = InStrRev(rest, "/")
    If slash < 2 Then Exit Function
    project = Mid$(rest, slash + 1)
    hostPath = Left$(rest, slash - 1)
    If Len(project) = 0 Or project Like "*[!0-9]*" Then Exit Function
    If hostPath Like "*[!A-Za-z0-9.:/_-]*" Or Left$(hostPath, 1) = "/" Then Exit Function
    TelemetryUrl = "https://" & hostPath & "/api/" & project & "/envelope/"
End Function

' A value cut back to the name at its start, so free text can't ride along: a
' procedure keeps "module.procedure" and an optional " line N"; anything else
' (an app name, page key, build or version) keeps letters, digits, spaces and
' . _ : - only. Public for the tests.
Public Function TelemetryName(ByVal s As String, Optional ByVal procedure As Boolean = False) As String
    Dim i As Long, ch As String, out As String, rest As String, digits As String
    s = Trim$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not (ch Like "[A-Za-z0-9_.]" Or (Not procedure And ch Like "[ :-]")) Then Exit For
        out = out & ch
    Next
    If procedure And Len(out) > 0 Then
        rest = Mid$(s, Len(out) + 1)
        If Left$(rest, 6) = " line " Then
            For i = 7 To Len(rest)
                If Not Mid$(rest, i, 1) Like "#" Then Exit For
                digits = digits & Mid$(rest, i, 1)
            Next
            If Len(digits) > 0 Then out = out & " line " & digits
        End If
    End If
    TelemetryName = Trim$(Left$(out, 100))
End Function

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

' The key part of a DSN's user info (before any ":secret"), or "" when it
' isn't hex.
Private Function DsnKey(ByVal dsn As String) As String
    Dim userInfo As String, at As Long
    dsn = Trim$(dsn)
    at = InStr(dsn, "@")
    If at <= 9 Then Exit Function
    userInfo = Mid$(dsn, 9, at - 9)
    If InStr(userInfo, ":") > 0 Then userInfo = Left$(userInfo, InStr(userInfo, ":") - 1)
    If Len(userInfo) = 0 Or userInfo Like "*[!0-9A-Fa-f]*" Then Exit Function
    DsnKey = userInfo
End Function

' The ";PWD=..." part of a linked table's Connect, as OpenDatabase takes it
' ("" when the file has no password).
Private Function ConnectPassword(ByVal cn As String) As String
    Dim p As Long, e As Long
    p = InStr(1, cn, ";PWD=", vbTextCompare)
    If p = 0 Then Exit Function
    e = InStr(p + 1, cn, ";")
    If e > 0 Then ConnectPassword = Mid$(cn, p, e - p) Else ConnectPassword = Mid$(cn, p)
End Function

Private Function StripLine(ByVal proc As String) As String
    Dim p As Long
    p = InStr(proc, " line ")
    If p > 0 Then StripLine = Left$(proc, p - 1) Else StripLine = proc   ' not IIf: it would run Left$(proc, -1)
End Function

' A random id for this install, made on first use and kept in the registry.
Private Function InstallId() As String
    On Error Resume Next
    InstallId = GetSetting(APP_NAME, "Telemetry", "Install", "")
    If Len(InstallId) <> 32 Then
        InstallId = LCase$(RandomHex(16))
        SaveSetting APP_NAME, "Telemetry", "Install", InstallId
    End If
End Function

Private Function TelemetryPage() As String
    On Error Resume Next
    If IsFormOpen("frmMain") Then TelemetryPage = Forms("frmMain").PageKey
End Function

Private Function TelemetryBuild() As String
    On Error Resume Next
    TelemetryBuild = Application.Run("FrontEndBuild")
End Function

' Drop finished requests (readyState 4) from the in-flight list.
Private Sub PruneFinished()
    Dim i As Long
    On Error Resume Next
    For i = m_pending.Count To 1 Step -1
        If m_pending(i).readyState = 4 Then m_pending.Remove i
    Next
End Sub

' Now in UTC, as Sentry wants it: 2026-10-18T02:40:16Z.
Private Function UtcStamp() As String
    Dim st As SYSTEMTIME
    GetSystemTime st
    UtcStamp = Format$(st.wYear, "0000") & "-" & Format$(st.wMonth, "00") & "-" & Format$(st.wDay, "00") & "T" & _
        Format$(st.wHour, "00") & ":" & Format$(st.wMinute, "00") & ":" & Format$(st.wSecond, "00") & "Z"
End Function

' s as a JSON string. Only names reach here, but quotes and backslashes are
' escaped and control characters dropped all the same.
Private Function JsonText(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case AscW(ch)
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 0 To 31
            Case Else: out = out & ch
        End Select
    Next
    JsonText = """" & out & """"
End Function

Private Function SqlQuote(ByVal s As String) As String
    SqlQuote = "'" & Replace(s, "'", "''") & "'"
End Function
