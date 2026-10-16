Option Compare Database
Option Explicit

' Access UI kit - the shared Personnel directory: who a person is, which apps
' they may open, and signing them in.
'
' The directory is one encrypted data file, Personnel_be.accdb, in its own
' restricted folder on the share, that every app links to. It holds two
' tables, and nothing else holds people's names:
'   kit_person      one row per person: Windows account and SID, sign-in
'                   name, rank and name, command, email, app password hash,
'                   directory admin, active (no SSN or DoD ID)
'   kit_person_app  which apps (app_code "SABER", "SPM", "MG") a person may open
' Each app keeps its own roles and links them to person_id.
'
' Startup, straight after the app's own KitEnsureBackend:
'   Select Case PersonnelStatus("item")          ' item: one of the app's linked tables
'       Case PERSONNEL_READY: signed = PersonnelSignInWindows(APP_CODE)
'       Case Else: ... the setup screen (find the file, create it, or ask the admin)
'   End Select
' PersonnelStatus links the file by itself when it is in the default place:
' a Personnel folder beside the app's data folder (S:\Apps\Data\app.accdb
' gives S:\Apps\Personnel\Personnel_be.accdb).
'
' Signing in sets TempVars PersonID, Username (what KitUser() stamps),
' PersonName ("LT J. Smith") and PersonIsAdmin. PersonnelSignInWindows
' matches the Windows account (WinSid, then WinAccount) with no questions;
' after Sign out, PersonnelSignIn checks a sign-in name and app password.
'
' The file's password: PersonnelKey(), the kit's default unless the app sets
' TempVars("PersonnelKey") before the first call. Every app that shares the
' file must use the same key. The folder's permissions are the real control:
' the key travels in each front end's links.

Public Const PERSONNEL_FILE As String = "Personnel_be.accdb"
Public Const PERSONNEL_FOLDER As String = "Personnel"
Public Const PERSON_TABLE As String = "kit_person"
Public Const PERSON_APP_TABLE As String = "kit_person_app"

' PersonnelStatus results.
Public Const PERSONNEL_READY As Long = 0          ' linked and reachable
Public Const PERSONNEL_NOT_SET_UP As Long = 1     ' no file linked or in the default place: offer find or create
Public Const PERSONNEL_LOST As Long = 2           ' linked, but the file can't be reached: offer find, never create

Private Const DEFAULT_KEY As String = "kit-personnel-1"
Private Const MIN_PASSWORD_LEN As Long = 8

' The Personnel file's password (see the top of the module).
Public Function PersonnelKey() As String
    On Error Resume Next                        ' TempVars may be unset
    PersonnelKey = Nz(TempVars("PersonnelKey"), "")
    On Error GoTo 0
    If Len(PersonnelKey) = 0 Then PersonnelKey = DEFAULT_KEY
End Function

' The linked Personnel file, or "" when the tables aren't linked.
Public Function PersonnelPath() As String
    PersonnelPath = KitBackendPath(PERSON_TABLE)
End Function

' Where the Personnel file is looked for when none is linked: a Personnel
' folder beside the folder of appTable's data file, or beside the front end
' when appTable isn't linked.
Public Function PersonnelDefaultPath(ByVal appTable As String) As String
    Dim folder As String, parent As String
    folder = KitBackendPath(appTable)
    If Len(folder) > 0 Then folder = Left$(folder, InStrRev(folder, "\") - 1) Else folder = CurrentProject.Path
    If InStrRev(folder, "\") > 0 Then         ' "S:" (a file at a drive root) has no parent
        parent = Left$(folder, InStrRev(folder, "\") - 1)
        If Not (parent Like "\\*" And InStr(3, parent, "\") = 0) Then folder = parent   ' never up to just "\\server"
    End If
    PersonnelDefaultPath = folder & "\" & PERSONNEL_FOLDER & "\" & PERSONNEL_FILE
End Function

' Whether the Personnel directory can be used (PERSONNEL_READY), and if not
' why. Links the file found in the default place, so a second app finds it
' with no questions.
Public Function PersonnelStatus(ByVal appTable As String) As Long
    Dim def As String
    On Error GoTo Fail
    If Len(PersonnelPath()) > 0 Then
        If KitBackendReachable(PERSON_TABLE) Then PersonnelStatus = PERSONNEL_READY: Exit Function
    End If
    def = PersonnelDefaultPath(appTable)
    If Len(PersonnelFileProblem(def)) = 0 Then
        LinkPersonnelFile def
        PersonnelStatus = PERSONNEL_READY
    ElseIf Len(PersonnelPath()) > 0 Then
        PersonnelStatus = PERSONNEL_LOST
    Else
        PersonnelStatus = PERSONNEL_NOT_SET_UP
    End If
    Exit Function
Fail:
    RaiseAgain "modPersonnel.PersonnelStatus", appTable
End Function

' Why path can't be used as the Personnel file ("" when it can): missing,
' not opening with PersonnelKey(), or without the two tables.
Public Function PersonnelFileProblem(ByVal path As String) As String
    Dim db As DAO.Database
    If Not FileExists(path) Then PersonnelFileProblem = "There is no file at " & path & ".": Exit Function
    On Error Resume Next
    Set db = DBEngine.OpenDatabase(path, False, True, ";PWD=" & PersonnelKey())
    On Error GoTo 0
    If db Is Nothing Then
        PersonnelFileProblem = "That file can't be opened as the Personnel file. It may be another app's file, or use a different key."
        Exit Function
    End If
    If Not (HasTable(db, PERSON_TABLE) And HasTable(db, PERSON_APP_TABLE)) Then
        PersonnelFileProblem = "That file isn't a Personnel file: it has no " & PERSON_TABLE & " table."
    End If
    db.Close
End Function

' Link the Personnel tables to path (replacing any links to another file),
' after checking it with PersonnelFileProblem (error 5 with its words).
Public Sub LinkPersonnelFile(ByVal path As String)
    Dim problem As String
    problem = PersonnelFileProblem(path)
    If Len(problem) > 0 Then Err.Raise 5, "modPersonnel.LinkPersonnelFile", problem
    KitLinkTables path, Array(PERSON_TABLE, PERSON_APP_TABLE), PersonnelKey()
End Sub

' Create a new, encrypted Personnel file at path, link it, and add the
' current Windows user as its first person and directory admin, with
' access to appCode. Refuses (error 5) while a Personnel file is linked, so
' nobody makes a second directory by mistake, or when path exists.
' Returns the admin's person_id.
Public Function CreatePersonnelFile(ByVal path As String, ByVal appCode As String, ByVal firstName As String, _
        ByVal lastName As String, Optional ByVal rank As String = "") As Long
    Dim db As DAO.Database, n As Long, d As String, src As String
    On Error GoTo Fail
    If Len(PersonnelPath()) > 0 Then Err.Raise 5, "modPersonnel", "A Personnel file is already linked: " & PersonnelPath()
    If FileExists(path) Then Err.Raise 5, "modPersonnel", "A file already exists at " & path & "."
    EnsureFolder Left$(path, InStrRev(path, "\") - 1)
    Set db = DBEngine.CreateDatabase(path, dbLangGeneral & ";pwd=" & PersonnelKey())
    CreatePersonnelTables db
    db.Close
    Set db = Nothing
    LinkPersonnelFile path
    CreatePersonnelFile = AddPerson(firstName, lastName, rank, WinAccount(), "", True)
    KitExecute "UPDATE " & PERSON_TABLE & " SET windows_sid=" & SqlOrNull(WinSid()) & " WHERE person_id=" & CreatePersonnelFile
    SetPersonApp CreatePersonnelFile, appCode, True
    Exit Function
Fail:
    If Not db Is Nothing Then
        n = Err.Number: d = Err.Description: src = Err.Source        ' Close resets Err
        On Error Resume Next
        db.Close
        On Error GoTo 0
        Err.Number = n: Err.Description = d: Err.Source = src
    End If
    RaiseAgain "modPersonnel.CreatePersonnelFile", path
End Function

' Add a person; returns their person_id. windowsAccount (DOMAIN\user) may be
' ""; username defaults to its user part. Raises error 5 when the sign-in
' name or Windows account is already someone's.
Public Function AddPerson(ByVal firstName As String, ByVal lastName As String, Optional ByVal rank As String = "", _
        Optional ByVal windowsAccount As String = "", Optional ByVal username As String = "", _
        Optional ByVal isAdmin As Boolean = False, Optional ByVal middleInitial As String = "", _
        Optional ByVal commandName As String = "", Optional ByVal email As String = "") As Long
    On Error GoTo Fail
    firstName = Trim$(firstName): lastName = Trim$(lastName)
    If Len(firstName) = 0 Or Len(lastName) = 0 Then Err.Raise 5, "modPersonnel", "A person needs a first and last name."
    If Len(username) = 0 Then username = SignInNameFor(windowsAccount)
    If Len(username) = 0 Then username = LCase$(Left$(firstName, 1) & Replace(lastName, " ", ""))
    If Not IsNull(DLookup("person_id", PERSON_TABLE, "username=" & SqlText(username))) Then
        Err.Raise 5, "modPersonnel", "The sign-in name " & username & " is already used."
    End If
    If Len(windowsAccount) > 0 Then
        If Not IsNull(DLookup("person_id", PERSON_TABLE, "windows_account=" & SqlText(windowsAccount))) Then
            Err.Raise 5, "modPersonnel", "The Windows account " & windowsAccount & " is already someone's."
        End If
    End If
    KitExecute "INSERT INTO " & PERSON_TABLE & " (windows_account, username, rank, first_name, middle_initial, last_name, " & _
        "display_name, command, email, is_admin, is_active, " & StampColumns() & ") VALUES (" & _
        SqlOrNull(windowsAccount) & ", " & SqlText(username) & ", " & SqlOrNull(rank) & ", " & SqlText(firstName) & ", " & _
        SqlOrNull(Left$(middleInitial, 1)) & ", " & SqlText(lastName) & ", " & _
        SqlText(PersonDisplayName(rank, firstName, lastName)) & ", " & SqlOrNull(commandName) & ", " & SqlOrNull(email) & ", " & _
        IIf(isAdmin, "True", "False") & ", True, " & StampValues() & ")"
    AddPerson = DLookup("person_id", PERSON_TABLE, "username=" & SqlText(username))
    Exit Function
Fail:
    RaiseAgain "modPersonnel.AddPerson", firstName & " " & lastName
End Function

' Give (or take away) a person's access to appCode.
Public Sub SetPersonApp(ByVal personId As Long, ByVal appCode As String, ByVal allowed As Boolean)
    Dim crit As String
    On Error GoTo Fail
    crit = "person_id=" & personId & " AND app_code=" & SqlText(appCode)
    If IsNull(DLookup("person_id", PERSON_APP_TABLE, crit)) Then
        KitExecute "INSERT INTO " & PERSON_APP_TABLE & " (person_id, app_code, is_active, granted_at, granted_by) VALUES (" & _
            personId & ", " & SqlText(appCode) & ", " & IIf(allowed, "True", "False") & ", Now(), " & SqlText(KitUser()) & ")"
    Else
        KitExecute "UPDATE " & PERSON_APP_TABLE & " SET is_active=" & IIf(allowed, "True", "False") & _
            ", granted_at=Now(), granted_by=" & SqlText(KitUser()) & " WHERE " & crit
    End If
    Exit Sub
Fail:
    RaiseAgain "modPersonnel.SetPersonApp", personId & " " & appCode
End Sub

' True when the person is active and may open appCode.
Public Function PersonCanUse(ByVal personId As Long, ByVal appCode As String) As Boolean
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT p.person_id FROM " & PERSON_TABLE & " AS p INNER JOIN " & PERSON_APP_TABLE & _
        " AS a ON p.person_id = a.person_id WHERE p.person_id=" & personId & " AND p.is_active AND a.is_active AND a.app_code=" & _
        SqlText(appCode), dbOpenSnapshot)
    PersonCanUse = Not rs.EOF
    rs.Close
End Function

' "LT J. Smith": rank, first initial, last name.
Public Function PersonDisplayName(ByVal rank As String, ByVal firstName As String, ByVal lastName As String) As String
    PersonDisplayName = Trim$(Trim$(rank) & " " & IIf(Len(Trim$(firstName)) > 0, UCase$(Left$(Trim$(firstName), 1)) & ". ", "") & Trim$(lastName))
End Function

' Sign in the person whose Windows account is running Access: by SID, else
' by DOMAIN\user (saving the SID, so a renamed account still matches).
' True when they are active and may open appCode; then the session is set
' and a toast says who. False, changing nothing, otherwise.
Public Function PersonnelSignInWindows(ByVal appCode As String, Optional ByVal quiet As Boolean = False) As Boolean
    Dim id As Variant, sid As String
    On Error GoTo Fail
    id = Null                                   ' Empty would skip the DOMAIN\user lookup
    sid = WinSid()
    If Len(sid) > 0 Then id = DLookup("person_id", PERSON_TABLE, "windows_sid=" & SqlText(sid))
    If IsNull(id) And Len(WinAccount()) > 0 Then
        id = DLookup("person_id", PERSON_TABLE, "windows_account=" & SqlText(WinAccount()) & " AND windows_sid Is Null")
        If Not IsNull(id) And Len(sid) > 0 Then
            KitExecute "UPDATE " & PERSON_TABLE & " SET windows_sid=" & SqlText(sid) & " WHERE person_id=" & id
        End If
    End If
    If IsNull(id) Then Exit Function
    If Not PersonCanUse(id, appCode) Then Exit Function
    StartSession id, quiet
    PersonnelSignInWindows = True
    Exit Function
Fail:
    RaiseAgain "modPersonnel.PersonnelSignInWindows", appCode
End Function

' The login screen's check: "" and the session set when username and
' password match an active person who may open appCode, else what to tell
' the user (never which of the two was wrong). A person with no password
' yet can't sign in here: PersonnelNeedsPassword says so, and
' PersonnelSetPassword sets it.
Public Function PersonnelSignIn(ByVal username As String, ByVal password As String, ByVal appCode As String) As String
    Dim rs As DAO.Recordset, ok As Boolean
    On Error GoTo Fail
    Set rs = CurrentDb.OpenRecordset("SELECT person_id, password_hash FROM " & PERSON_TABLE & _
        " WHERE is_active AND username=" & SqlText(Trim$(username)), dbOpenSnapshot)
    If Not rs.EOF Then ok = VerifyPassword(password, Nz(rs!password_hash, ""))
    If ok Then
        If Not PersonCanUse(rs!person_id, appCode) Then
            PersonnelSignIn = "You don't have access to this app yet. Ask a Personnel admin to add it."
        Else
            StartSession rs!person_id, False
        End If
    Else
        PersonnelSignIn = "That sign-in name and password don't match."
    End If
    rs.Close
    Exit Function
Fail:
    RaiseAgain "modPersonnel.PersonnelSignIn", username
End Function

' True when username belongs to an active person who has no app password
' yet (the login screen then asks them to set one).
Public Function PersonnelNeedsPassword(ByVal username As String) As Boolean
    Dim h As Variant
    h = DLookup("Nz(password_hash,'')", PERSON_TABLE, "is_active AND username=" & SqlText(Trim$(username)))
    If Not IsNull(h) Then PersonnelNeedsPassword = (Len(h) = 0)
End Function

' Set a person's app password (salted hash). Returns "" or why the password
' is refused.
Public Function PersonnelSetPassword(ByVal personId As Long, ByVal password As String) As String
    If Len(password) < MIN_PASSWORD_LEN Then
        PersonnelSetPassword = "Use at least " & MIN_PASSWORD_LEN & " characters."
        Exit Function
    End If
    KitExecute "UPDATE " & PERSON_TABLE & " SET password_hash=" & SqlText(HashPassword(password)) & ", " & StampSql() & _
        " WHERE person_id=" & personId
End Function

' Save the current Windows account to the signed-in person's record ("Use
' my Windows sign-in next time"). Refused (error 5) when the account is
' already someone else's.
Public Sub PersonnelSaveWindowsAccount()
    Dim other As Variant
    If SignedInPersonID() = 0 Then Err.Raise 5, "modPersonnel", "Nobody is signed in."
    other = DLookup("person_id", PERSON_TABLE, "windows_account=" & SqlText(WinAccount()) & " AND person_id<>" & SignedInPersonID())
    If Not IsNull(other) Then Err.Raise 5, "modPersonnel", "This Windows account is already saved for someone else."
    KitExecute "UPDATE " & PERSON_TABLE & " SET windows_account=" & SqlText(WinAccount()) & ", windows_sid=" & _
        SqlOrNull(WinSid()) & ", " & StampSql() & " WHERE person_id=" & SignedInPersonID()
End Sub

' Sign out: clear the session. The app then shows its login screen (not a
' Windows sign-in), so another person can sign in.
Public Sub PersonnelSignOut()
    Dim v As Variant
    On Error Resume Next                        ' one that was never set
    For Each v In Array("PersonID", "Username", "PersonName", "PersonIsAdmin")
        TempVars.Remove v
    Next
End Sub

' The signed-in person's person_id, or 0.
Public Function SignedInPersonID() As Long
    On Error Resume Next                        ' TempVars may be unset
    SignedInPersonID = Nz(TempVars("PersonID"), 0)
End Function

' True when the signed-in person is a directory admin.
Public Function SignedInIsAdmin() As Boolean
    On Error Resume Next
    SignedInIsAdmin = Nz(TempVars("PersonIsAdmin"), False)
End Function

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

Private Sub StartSession(ByVal personId As Long, ByVal quiet As Boolean)
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT username, display_name, is_admin FROM " & PERSON_TABLE & _
        " WHERE person_id=" & personId, dbOpenSnapshot)
    TempVars("PersonID") = personId
    TempVars("Username") = rs!username.Value
    TempVars("PersonName") = Nz(rs!display_name, rs!username)
    TempVars("PersonIsAdmin") = CBool(rs!is_admin)
    rs.Close
    If Not quiet Then KitToast "Signed in as " & TempVars("PersonName"), "success"
End Sub

' The sign-in name for a Windows account: "NAVY\jsmith" gives "jsmith".
Private Function SignInNameFor(ByVal windowsAccount As String) As String
    SignInNameFor = LCase$(Mid$(windowsAccount, InStrRev(windowsAccount, "\") + 1))
End Function

Private Sub CreatePersonnelTables(db As DAO.Database)
    db.Execute "CREATE TABLE " & PERSON_TABLE & " (person_id COUNTER CONSTRAINT pk_" & PERSON_TABLE & " PRIMARY KEY, " & _
        "windows_account TEXT(150), windows_sid TEXT(184), username TEXT(64) NOT NULL, rank TEXT(20), " & _
        "first_name TEXT(60) NOT NULL, middle_initial TEXT(1), last_name TEXT(60) NOT NULL, display_name TEXT(100), " & _
        "command TEXT(100), email TEXT(150), password_hash TEXT(255), is_admin YESNO, is_active YESNO, " & _
        "created_by TEXT(150), created_at DATETIME, updated_by TEXT(150), updated_at DATETIME)", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX ux_person_username ON " & PERSON_TABLE & " (username)", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX ux_person_account ON " & PERSON_TABLE & " (windows_account) WITH IGNORE NULL", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX ux_person_sid ON " & PERSON_TABLE & " (windows_sid) WITH IGNORE NULL", dbFailOnError
    db.Execute "CREATE TABLE " & PERSON_APP_TABLE & " (person_id LONG NOT NULL, app_code TEXT(20) NOT NULL, is_active YESNO, " & _
        "granted_at DATETIME, granted_by TEXT(150), CONSTRAINT pk_" & PERSON_APP_TABLE & " PRIMARY KEY (person_id, app_code), " & _
        "CONSTRAINT fk_person_app FOREIGN KEY (person_id) REFERENCES " & PERSON_TABLE & " (person_id))", dbFailOnError
End Sub

Private Function HasTable(db As DAO.Database, ByVal nm As String) As Boolean
    Dim td As DAO.TableDef
    For Each td In db.TableDefs
        If StrComp(td.Name, nm, vbTextCompare) = 0 Then HasTable = True: Exit Function
    Next
End Function

' Text as a quoted SQL literal.
Private Function SqlText(ByVal s As String) As String
    SqlText = "'" & Replace(s, "'", "''") & "'"
End Function

' Text as a quoted SQL literal, or Null when empty.
Private Function SqlOrNull(ByVal s As String) As String
    If Len(Trim$(s)) = 0 Then SqlOrNull = "Null" Else SqlOrNull = SqlText(Trim$(s))
End Function
