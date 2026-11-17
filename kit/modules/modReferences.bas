Option Compare Database
Option Explicit

' Access UI kit - the shared Reference library: one list of reference
' documents (instructions, manuals, notices) that every app cites, kept the
' way the Personnel directory is kept: one file on the shared drive, used by
' every app, never copied into an app's own data file. The web kit
' (navy-web-apps kit/server/KitReferences.ps1) reads and writes the same
' file and tables, so a reference added in an Access app shows in the web
' apps and the other way round. Keep the two in step.
'
' The file: References_be.accdb, beside the Personnel file (the same folder
' as Personnel_be.accdb). It has no database password: nothing in it is
' personal, and it holds only what may be read by anyone who can open the
' apps. The web makes it from its Data files page; a Personnel admin can make
' it here too (CreateReferencesFile).
'
' Tables (made with CREATE TABLE when missing, so an older file catches up on
' its own):
'   kit_reference         one row per document: ref_number (OPNAVINST
'                         3120.32D), ref_title, ref_revision (CH-1), pub_date,
'                         ref_status (active, superseded, cancelled),
'                         superseded_by (the ref_id that replaced it),
'                         doc_link (a shared-drive path or an http(s)
'                         address), ref_marking, ref_notes, source_app (the
'                         app it was first added from), row_version (+1 on
'                         every change, for the stale-save check), created_*
'                         and updated_*.
'   kit_reference_editor  who may change and retire references (Library
'                         editors). Personnel admins always may, and they
'                         choose the editors.
'   kit_reference_audit   every add, change, status change and editor change,
'                         with who and from which PC.
'   kit_reference_setting name/value, used by the web's first-run import.
' The first three are linked into the front end; the settings table is not.
'
' Who may do what (Stephen, 2026-11-09): anyone signed in to an app adds a
' reference, so drafting is never held up; Library editors and Personnel
' admins change, retire and supersede. Nobody deletes: a reference other
' files cite is retired (cancelled or superseded), never removed.
'
' The same document twice: two references are the same when their numbers
' match (case, spaces, dashes, underscores and dots ignored) and their
' revisions match; a reference without a number matches on its title
' instead. RefAdd hands back the one already there rather than adding it
' again.
'
' Startup, after KitPersonnelStart (modReferencesUI wraps this):
'   ReferencesStatus()     links the file found beside the Personnel file
' For app code (a citation keeps a ref_id and a page number in the app's own
' data file, as the web's comment_citation does):
'   RefAdd number, title, revision, pubDate, link, marking, notes, appCode, added
'   RefUpdate id, ..., status, supersededBy, version     editors only
'   RefLabel(id)            "NUMBER Title (REV)", as a citation reads
'   RefPickSql()            the active references, for a PickBox
'   RefOpen(id, page)       opens the document (a PDF in Edge at the page)
'   RefIsEditor()           may the signed-in person change references
'   RefSetMarkingCheck "AppMarkingProblem"   the app's own marking rules
'
' 32-bit and 64-bit Office: no Declare and no ActiveX here. Windows is
' reached late bound (WScript.Shell, Shell.Application) and through Shell,
' which work the same in both.

Public Const REF_FILE As String = "References_be.accdb"
Public Const REF_TABLE As String = "kit_reference"
Public Const REF_EDITOR_TABLE As String = "kit_reference_editor"
Public Const REF_AUDIT_TABLE As String = "kit_reference_audit"
Public Const REF_SETTING_TABLE As String = "kit_reference_setting"

' ReferencesStatus results.
Public Const REFS_READY As Long = 0           ' linked and reachable
Public Const REFS_NOT_SET_UP As Long = 1      ' no file linked or beside the Personnel file: a Personnel admin can make one
Public Const REFS_LOST As Long = 2            ' linked, but the file can't be reached: offer find, never create
Public Const REFS_NO_PERSONNEL As Long = 3    ' no Personnel file linked, so there is nowhere to look

' The statuses a reference can have (ref_status).
Public Const REF_ACTIVE As String = "active"
Public Const REF_SUPERSEDED As String = "superseded"
Public Const REF_CANCELLED As String = "cancelled"

' The widest marking kept (the web's KIT_MARKING_MAX): room for REL TO lists.
Public Const REF_MARKING_MAX As Long = 100
' The largest page a citation may name (the web's /open route).
Public Const REF_PAGE_MAX As Long = 99999

' The only document types RefOpen opens, whatever the link says: anything
' else may run code when opened, so it can only be copied.
Public Const REF_OPEN_TYPES As String = ".pdf.doc.docx.xls.xlsx.ppt.pptx.txt.rtf.xps.oxps.png.jpg.jpeg.tif.tiff."

' The marking levels a reference may carry, lowest first (the web's
' KIT_MARKING_LEVELS). A REL TO list may follow: "SECRET//REL TO USA, FVEY".
Public Const REF_MARKING_LEVELS As String = "UNCLASSIFIED;CUI;CONFIDENTIAL;SECRET"

Private m_markingCheck As String     ' the app's marking check (RefSetMarkingCheck), or ""

' ---------------------------------------------------------------------------
' The file
' ---------------------------------------------------------------------------

' The linked References file, or "" when the tables aren't linked.
Public Function ReferencesPath() As String
    ReferencesPath = KitBackendPath(REF_TABLE)
End Function

' Where the References file belongs: beside the linked Personnel file (as
' the web's Resolve-KitReferencesPath). "" when no Personnel file is linked.
Public Function ReferencesDefaultPath() As String
    Dim pe As String
    pe = PersonnelPath()
    If Len(pe) = 0 Or InStrRev(pe, "\") = 0 Then Exit Function
    ReferencesDefaultPath = Left$(pe, InStrRev(pe, "\")) & REF_FILE
End Function

' Whether the Reference library can be used (REFS_READY), and if not why.
' Links the file beside the Personnel file when it is there, so every app
' finds it with no questions.
Public Function ReferencesStatus() As Long
    Dim def As String
    On Error GoTo Fail
    If Len(ReferencesPath()) > 0 Then
        If KitBackendReachable(REF_TABLE) Then ReferencesStatus = REFS_READY: Exit Function
    End If
    def = ReferencesDefaultPath()
    If Len(def) > 0 Then
        If Len(ReferencesFileProblem(def)) = 0 Then
            LinkReferencesFile def
            ReferencesStatus = REFS_READY
            Exit Function
        End If
    End If
    If Len(ReferencesPath()) > 0 Then
        ReferencesStatus = REFS_LOST
    ElseIf Len(def) = 0 Then
        ReferencesStatus = REFS_NO_PERSONNEL
    Else
        ReferencesStatus = REFS_NOT_SET_UP
    End If
    Exit Function
Fail:
    RaiseAgain "modReferences.ReferencesStatus"
End Function

' True when the library is linked and reachable now.
Public Function ReferencesReady() As Boolean
    If Len(ReferencesPath()) = 0 Then Exit Function
    ReferencesReady = KitBackendReachable(REF_TABLE)
End Function

' Why path can't be used as the References file ("" when it can): missing,
' not opening (it has a password, or isn't an Access file), or without the
' kit_reference table.
Public Function ReferencesFileProblem(ByVal path As String) As String
    Dim db As DAO.Database
    If Not FileExists(path) Then ReferencesFileProblem = "There is no file at " & path & ".": Exit Function
    On Error Resume Next
    Set db = DBEngine.OpenDatabase(path, False, True)
    On Error GoTo 0
    If db Is Nothing Then
        ReferencesFileProblem = "That file can't be opened as the Reference library. It may be another app's file."
        Exit Function
    End If
    If Not RefHasTable(db, REF_TABLE) Then ReferencesFileProblem = "That file isn't a Reference library: it has no " & REF_TABLE & " table."
    db.Close
End Function

' Link the library's tables to path (replacing any links to another file),
' after checking it (ReferencesFileProblem, error 5 with its words) and
' making any table it lacks.
Public Sub LinkReferencesFile(ByVal path As String)
    Dim problem As String, n As Long, d As String, src As String
    KitBusy "Connecting to the Reference library..."   ' a file on a share can take a while
    On Error GoTo Fail
    problem = ReferencesFileProblem(path)
    If Len(problem) > 0 Then Err.Raise 5, "modReferences.LinkReferencesFile", problem
    EnsureReferencesTables path
    KitLinkTables path, Array(REF_TABLE, REF_EDITOR_TABLE, REF_AUDIT_TABLE)
    KitBusyEnd
    Exit Sub
Fail:
    ' KitBusyEnd resets Err, so keep the error and put it back.
    n = Err.Number: d = Err.Description: src = Err.Source
    KitBusyEnd
    Err.Number = n: Err.Description = d: Err.Source = src
    RaiseAgain "modReferences.LinkReferencesFile", path
End Sub

' Make a new, empty References file at path (normally ReferencesDefaultPath)
' and link it. Personnel admins only. Refused (error 5) while a library is
' linked and reachable, or when a file is already at path.
Public Sub CreateReferencesFile(ByVal path As String)
    Dim db As DAO.Database, made As Boolean, n As Long, d As String, src As String
    On Error GoTo Fail
    If Not PersonnelIsAdminNow() Then Err.Raise 5, "modReferences", "Only a Personnel admin can set up the Reference library."
    If ReferencesReady() Then Err.Raise 5, "modReferences", "A Reference library is already linked: " & ReferencesPath()
    If FileExists(path) Then Err.Raise 5, "modReferences", "A file already exists at " & path & "."
    EnsureFolder Left$(path, InStrRev(path, "\") - 1)
    Set db = DBEngine.CreateDatabase(path, dbLangGeneral)
    made = True
    CreateReferencesTables db
    db.Close
    Set db = Nothing
    LinkReferencesFile path
    Exit Sub
Fail:
    n = Err.Number: d = Err.Description: src = Err.Source        ' Close resets Err
    On Error Resume Next
    If Not db Is Nothing Then db.Close
    If made Then Kill path
    On Error GoTo 0
    Err.Number = n: Err.Description = d: Err.Source = src
    RaiseAgain "modReferences.CreateReferencesFile", path
End Sub

' The tables a References file lacks (an older or half-made file), and room
' for REL TO markings in a ref_marking made 30 wide. Widening needs nobody
' else in the file; when it can't, the narrower width stands (RefMarkingWidth).
Private Sub EnsureReferencesTables(ByVal path As String)
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(path, False, False)
    On Error GoTo Fail
    CreateReferencesTables db
    db.TableDefs.Refresh
    If db.TableDefs(REF_TABLE).Fields("ref_marking").Size < REF_MARKING_MAX Then
        On Error Resume Next                    ' someone else has it open: keep the width
        db.Execute "ALTER TABLE " & REF_TABLE & " ALTER COLUMN ref_marking TEXT(" & REF_MARKING_MAX & ")", dbFailOnError
        On Error GoTo Fail
    End If
    db.Close
    Exit Sub
Fail:
    Dim n As Long, d As String, src As String
    n = Err.Number: d = Err.Description: src = Err.Source
    On Error Resume Next
    db.Close
    On Error GoTo 0
    Err.Number = n: Err.Description = d: Err.Source = src
    RaiseAgain "modReferences.EnsureReferencesTables", path
End Sub

' The same tables, columns and indexes as the web's Initialize-KitReferencesTables.
Private Sub CreateReferencesTables(db As DAO.Database)
    If Not RefHasTable(db, REF_TABLE) Then
        db.Execute "CREATE TABLE " & REF_TABLE & " (ref_id COUNTER CONSTRAINT pk_kit_reference PRIMARY KEY, ref_number TEXT(100), " & _
            "ref_title TEXT(255) NOT NULL, ref_revision TEXT(30), pub_date DATETIME, ref_status TEXT(12) NOT NULL, superseded_by LONG, " & _
            "doc_link TEXT(255), ref_marking TEXT(" & REF_MARKING_MAX & "), ref_notes LONGTEXT, source_app TEXT(30), row_version LONG NOT NULL, " & _
            "created_by_id LONG, created_by TEXT(100), created_at DATETIME, updated_by TEXT(100), updated_at DATETIME)", dbFailOnError
        db.Execute "CREATE INDEX ix_kit_reference_number ON " & REF_TABLE & " (ref_number)", dbFailOnError
    End If
    If Not RefHasTable(db, REF_EDITOR_TABLE) Then
        db.Execute "CREATE TABLE " & REF_EDITOR_TABLE & " (person_id LONG CONSTRAINT pk_kit_reference_editor PRIMARY KEY, " & _
            "person_name TEXT(100), granted_by TEXT(100), granted_at DATETIME)", dbFailOnError
    End If
    If Not RefHasTable(db, REF_AUDIT_TABLE) Then
        db.Execute "CREATE TABLE " & REF_AUDIT_TABLE & " (audit_id COUNTER CONSTRAINT pk_kit_reference_audit PRIMARY KEY, logged_at DATETIME, " & _
            "actor_id LONG, actor_name TEXT(100), ref_id LONG, [action] TEXT(30), detail TEXT(255), app_key TEXT(30), " & _
            "windows_account TEXT(150), device TEXT(100))", dbFailOnError
        db.Execute "CREATE INDEX ix_kit_reference_audit_ref ON " & REF_AUDIT_TABLE & " (ref_id)", dbFailOnError
    End If
    If Not RefHasTable(db, REF_SETTING_TABLE) Then
        db.Execute "CREATE TABLE " & REF_SETTING_TABLE & " (setting_name TEXT(50) CONSTRAINT pk_kit_reference_setting PRIMARY KEY, " & _
            "setting_value TEXT(255), updated_at DATETIME)", dbFailOnError
    End If
End Sub

' The width ref_marking has in the linked file (REF_MARKING_MAX once widened).
Public Function RefMarkingWidth() As Long
    On Error Resume Next
    RefMarkingWidth = CurrentDb.TableDefs(REF_TABLE).Fields("ref_marking").Size
    If RefMarkingWidth = 0 Then RefMarkingWidth = REF_MARKING_MAX
End Function

' ---------------------------------------------------------------------------
' Matching and labels
' ---------------------------------------------------------------------------

' A number or revision as it is compared: capitals, with spaces, tabs,
' dashes, underscores and dots left out ("OPNAVINST 3120.32D" and
' "opnavinst3120.32 d" match), as the web's Get-KitRefKey.
Public Function RefKey(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = UCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case " ", vbTab, vbCr, vbLf, "-", "_", "."
            Case Else: out = out & ch
        End Select
    Next
    RefKey = out
End Function

' A title as it is compared when there is no number: lower case, with runs
' of white space as one space, trimmed (the web's Get-KitRefTitleKey).
Public Function RefTitleKey(ByVal s As String) As String
    RefTitleKey = LCase$(RefOneSpace(s))
End Function

' s with tabs and line breaks as spaces, runs of spaces as one, trimmed.
Private Function RefOneSpace(ByVal s As String) As String
    s = Replace(Replace(Replace(s, vbTab, " "), vbCr, " "), vbLf, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    RefOneSpace = Trim$(s)
End Function

' The reference that is the same document as number, title and revision
' (see the top of the module), leaving out exceptId; 0 when there is none.
Public Function RefFindSame(ByVal number As String, ByVal title As String, ByVal revision As String, _
        Optional ByVal exceptId As Long = 0) As Long
    Dim rs As DAO.Recordset, nk As String, rk As String, tk As String
    nk = RefKey(number): rk = RefKey(revision): tk = RefTitleKey(title)
    Set rs = CurrentDb.OpenRecordset("SELECT ref_id, ref_number, ref_title, ref_revision FROM " & REF_TABLE, dbOpenSnapshot)
    Do While Not rs.EOF
        If rs!ref_id <> exceptId And StrComp(RefKey(Nz(rs!ref_revision, "")), rk, vbBinaryCompare) = 0 Then
            If Len(nk) > 0 Then
                If StrComp(RefKey(Nz(rs!ref_number, "")), nk, vbBinaryCompare) = 0 Then RefFindSame = rs!ref_id: Exit Do
            ElseIf Len(Nz(rs!ref_number, "")) = 0 Then
                If StrComp(RefTitleKey(Nz(rs!ref_title, "")), tk, vbBinaryCompare) = 0 Then RefFindSame = rs!ref_id: Exit Do
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
End Function

' "NUMBER Title (REV)", as a citation reads (the web's Format-KitReference).
Public Function RefLabelOf(ByVal number As String, ByVal title As String, ByVal revision As String) As String
    Dim s As String
    s = title
    If Len(number) > 0 Then s = number & " " & title
    If Len(revision) > 0 Then s = s & " (" & revision & ")"
    RefLabelOf = s
End Function

' The label of reference id, or "" when there is none.
Public Function RefLabel(ByVal id As Long) As String
    Dim rs As DAO.Recordset
    If id <= 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT ref_number, ref_title, ref_revision FROM " & REF_TABLE & " WHERE ref_id=" & id, dbOpenSnapshot)
    If Not rs.EOF Then RefLabel = RefLabelOf(Nz(rs!ref_number, ""), Nz(rs!ref_title, ""), Nz(rs!ref_revision, ""))
    rs.Close
End Function

' The label as an Access SQL expression over kit_reference's columns.
Public Function RefLabelSql() As String
    RefLabelSql = "IIf(Nz(ref_number,'')='','',ref_number & ' ') & ref_title & IIf(Nz(ref_revision,'')='','',' (' & ref_revision & ')')"
End Function

' The active references for a PickBox (modMenu): ref_id, the label, and
' "Document" in muted text when a document is linked.
Public Function RefPickSql() As String
    RefPickSql = "SELECT ref_id, " & RefLabelSql() & " AS ref_label, IIf(Nz(doc_link,'')='','','Document') AS ref_doc FROM " & _
        REF_TABLE & " WHERE ref_status='" & REF_ACTIVE & "'"
End Function

' The rows of the library page: ref_id, Number, Title, Revision, Date,
' Status, Document; matching term (number, title, revision or notes) and, unless
' withRetired, active only. term is matched as typed: [ * ? # are not wildcards.
Public Function RefLibrarySql(ByVal term As String, ByVal withRetired As Boolean) As String
    Dim w As String, t As String
    If Not withRetired Then w = "ref_status='" & REF_ACTIVE & "'"
    term = Trim$(term)
    If Len(term) > 0 Then
        t = "'*" & RefLikeText(term) & "*'"
        w = IIf(Len(w) > 0, w & " AND ", "") & "(ref_number LIKE " & t & " OR ref_title LIKE " & t & _
            " OR ref_revision LIKE " & t & " OR ref_notes LIKE " & t & ")"
    End If
    RefLibrarySql = "SELECT ref_id, ref_number AS [Number], ref_title AS [Title], ref_revision AS [Revision], " & _
        "Format(pub_date,'yyyy-mm-dd') AS [Date], StrConv(ref_status,3) AS [Status], " & _
        "IIf(Nz(doc_link,'')='','','Linked') AS [Document] FROM " & REF_TABLE & _
        IIf(Len(w) > 0, " WHERE " & w, "") & " ORDER BY ref_number, ref_title, ref_revision"
End Function

' s for a LIKE pattern, quoted for SQL: [ * ? # are matched as themselves.
Private Function RefLikeText(ByVal s As String) As String
    s = Replace(s, "[", "[[]")
    s = Replace(s, "*", "[*]")
    s = Replace(s, "?", "[?]")
    s = Replace(s, "#", "[#]")
    RefLikeText = Replace(s, "'", "''")
End Function

' ---------------------------------------------------------------------------
' Input
' ---------------------------------------------------------------------------

' The app's own marking rules (a SABER-style CAPCO check, or the ceiling of
' the system's classification): the name of a public function in the app,
' taking the marking and returning "" when it may be kept or a sentence
' saying why not. "" goes back to the kit's own check only.
Public Sub RefSetMarkingCheck(ByVal functionName As String)
    m_markingCheck = functionName
End Sub

' marking as it is kept: trimmed, capitals, UNCLASSIFIED when blank.
Public Function RefMarkingText(ByVal marking As String) As String
    marking = UCase$(RefOneSpace(marking))
    If Len(marking) = 0 Then marking = "UNCLASSIFIED"
    RefMarkingText = marking
End Function

' Why marking can't be kept ("" when it can): a level (REF_MARKING_LEVELS),
' optionally followed by "//" and more (REL TO countries), no longer than
' the column; then the app's own check (RefSetMarkingCheck).
Public Function RefMarkingProblem(ByVal marking As String) As String
    Dim lvl As String, v As Variant, ok As Boolean
    marking = RefMarkingText(marking)
    lvl = marking
    If InStr(lvl, "//") > 0 Then lvl = Trim$(Left$(lvl, InStr(lvl, "//") - 1))
    For Each v In Split(REF_MARKING_LEVELS, ";")
        If lvl = v Then ok = True
    Next
    If Not ok Then
        RefMarkingProblem = "The marking must start with UNCLASSIFIED, CUI, CONFIDENTIAL or SECRET."
    ElseIf Right$(marking, 2) = "//" Then
        RefMarkingProblem = "Finish the marking after //, or take the // off."
    ElseIf Len(marking) > RefMarkingWidth() Then
        RefMarkingProblem = "The References file holds markings of at most " & RefMarkingWidth() & _
            " characters until it is updated (open the app while nobody else has the file open), so name fewer countries."
    ElseIf Len(m_markingCheck) > 0 Then
        RefMarkingProblem = Nz(Application.Run(m_markingCheck, marking), "")
    End If
End Function

' Why link can't be a reference's document link ("" when it can): a web
' address (https://...) or a file path (\\server\share\... or S:\...), at
' most 255 characters, as the web's Get-KitReferenceInput.
Public Function RefLinkProblem(ByVal link As String) As String
    link = RefLinkText(link)
    If Len(link) = 0 Then Exit Function
    If Len(link) > 255 Then RefLinkProblem = "A link can be at most 255 characters.": Exit Function
    If (LCase$(link) Like "http://?*" Or LCase$(link) Like "https://?*") Then
        If InStr(link, " ") = 0 And InStr(link, vbTab) = 0 Then Exit Function
    ElseIf link Like "\\?*\?*" Then
        If InStr(Mid$(link, 3, InStr(3, link, "\") - 3), "/") = 0 Then Exit Function
    ElseIf link Like "[A-Za-z]:\?*" Then
        Exit Function
    End If
    RefLinkProblem = "The link must be a web address (https://...) or a file path (\\server\share\... or S:\...)."
End Function

' link as it is kept: trimmed, without the quotes Windows' "Copy as path" adds.
Public Function RefLinkText(ByVal link As String) As String
    link = Trim$(link)
    If Len(link) >= 2 And Left$(link, 1) = """" And Right$(link, 1) = """" Then link = Trim$(Mid$(link, 2, Len(link) - 2))
    RefLinkText = link
End Function

' Why these can't be a reference ("" when they can): the web's rules and
' words. pubDate is Null or a date.
Public Function RefInputProblem(ByVal number As String, ByVal title As String, ByVal revision As String, _
        ByVal pubDate As Variant, ByVal link As String, ByVal marking As String, ByVal notes As String) As String
    Dim p As String
    title = RefOneSpace(title)
    If Len(title) = 0 Then p = "Enter the reference's title."
    If Len(p) = 0 And Len(title) > 255 Then p = "A title can be at most 255 characters."
    If Len(p) = 0 And Len(Trim$(number)) > 100 Then p = "A number can be at most 100 characters."
    If Len(p) = 0 And Len(Trim$(revision)) > 30 Then p = "A revision can be at most 30 characters."
    If Len(p) = 0 And Not IsNull(pubDate) Then
        If Not IsDate(pubDate) Then
            p = "The date must be a real date."
        ElseIf Year(CDate(pubDate)) < 1900 Or Year(CDate(pubDate)) > 2100 Then
            p = "The date must be a real date."
        End If
    End If
    If Len(p) = 0 Then p = RefLinkProblem(link)
    If Len(p) = 0 Then p = RefMarkingProblem(marking)
    If Len(p) = 0 And Len(Trim$(notes)) > 2000 Then p = "Notes can be at most 2000 characters."
    RefInputProblem = p
End Function

' ---------------------------------------------------------------------------
' Who may change references
' ---------------------------------------------------------------------------

' True when the signed-in person may change, retire and supersede
' references now: a Personnel admin (PersonnelIsAdminNow) or a Library editor.
Public Function RefIsEditor() As Boolean
    If SignedInPersonID() = 0 Then Exit Function
    If PersonnelIsAdminNow() Then RefIsEditor = True: Exit Function
    RefIsEditor = RefIsEditorId(SignedInPersonID())
End Function

' True when personId is on the Library editors list (Personnel admins aren't
' listed: they always may).
Public Function RefIsEditorId(ByVal personId As Long) As Boolean
    RefIsEditorId = (DCount("*", REF_EDITOR_TABLE, "person_id=" & personId) > 0)
End Function

' Make personId a Library editor, or not. Personnel admins only (error 5).
Public Sub RefSetEditor(ByVal personId As Long, ByVal editor As Boolean)
    Dim ws As DAO.Workspace, db As DAO.Database, rs As DAO.Recordset, nm As String, inTrans As Boolean, has As Boolean
    On Error GoTo Fail
    If Not PersonnelIsAdminNow() Then Err.Raise 5, "modReferences", "Only Personnel admins choose the Library editors."
    nm = Nz(DLookup("display_name", PERSON_TABLE, "person_id=" & personId), "")
    If Len(nm) = 0 Then Err.Raise 5, "modReferences", "Choose a person from the Personnel directory."
    nm = Left$(nm, 100)
    has = RefIsEditorId(personId)
    If has = editor Then Exit Sub
    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb
    ws.BeginTrans: inTrans = True
    If editor Then
        Set rs = db.OpenRecordset(REF_EDITOR_TABLE, dbOpenDynaset, dbAppendOnly)
        rs.AddNew
        rs!person_id = personId: rs!person_name = nm: rs!granted_by = RefActorName(): rs!granted_at = Now()
        rs.Update
        rs.Close
        RefAudit db, 0, "EditorAdded", nm & " is now a Library editor.", ""
    Else
        db.Execute "DELETE FROM " & REF_EDITOR_TABLE & " WHERE person_id=" & personId, dbFailOnError
        RefAudit db, 0, "EditorRemoved", nm & " is no longer a Library editor.", ""
    End If
    ws.CommitTrans: inTrans = False
    Exit Sub
Fail:
    If inTrans Then RefRollback ws
    RaiseAgain "modReferences.RefSetEditor", CStr(personId)
End Sub

' The Library editors, for a list: person_id, Name, Granted by, Granted.
Public Function RefEditorsSql() As String
    RefEditorsSql = "SELECT person_id, person_name AS [Name], granted_by AS [Granted by], " & _
        "Format(granted_at,'yyyy-mm-dd') AS [Granted] FROM " & REF_EDITOR_TABLE & " ORDER BY person_name"
End Function

' ---------------------------------------------------------------------------
' Adding and changing
' ---------------------------------------------------------------------------

' Add a reference, or find the one already there (see the top of the
' module): returns its ref_id, with added True when it is new. Anyone signed
' in may add. appCode is the app it was added from ("SABER"). Refused (error
' 5, with the reason) when nobody is signed in or a field breaks the rules
' (RefInputProblem).
Public Function RefAdd(ByVal number As String, ByVal title As String, ByVal revision As String, ByVal pubDate As Variant, _
        ByVal link As String, ByVal marking As String, ByVal notes As String, ByVal appCode As String, _
        Optional added As Boolean) As Long
    Dim ws As DAO.Workspace, db As DAO.Database, rs As DAO.Recordset, problem As String, id As Long, inTrans As Boolean, who As String
    On Error GoTo Fail
    added = False
    If SignedInPersonID() = 0 Then Err.Raise 5, "modReferences", "Sign in first to add a reference."
    problem = RefInputProblem(number, title, revision, pubDate, link, marking, notes)
    If Len(problem) > 0 Then Err.Raise 5, "modReferences", problem
    number = Trim$(number): title = RefOneSpace(title): revision = Trim$(revision)
    link = RefLinkText(link): marking = RefMarkingText(marking): notes = Trim$(notes)
    who = RefActorName()
    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb
    ws.BeginTrans: inTrans = True
    id = RefFindSame(number, title, revision)
    If id = 0 Then
        Set rs = db.OpenRecordset(REF_TABLE, dbOpenDynaset, dbAppendOnly)
        rs.AddNew
        rs!ref_number = RefNull(number): rs!ref_title = title: rs!ref_revision = RefNull(revision)
        If IsNull(pubDate) Then rs!pub_date = Null Else rs!pub_date = DateValue(CDate(pubDate))
        rs!ref_status = REF_ACTIVE: rs!doc_link = RefNull(link): rs!ref_marking = marking
        rs!ref_notes = RefNull(notes): rs!source_app = RefNull(Left$(appCode, 30)): rs!row_version = 1
        rs!created_by_id = SignedInPersonID(): rs!created_by = who: rs!created_at = Now()
        rs!updated_by = who: rs!updated_at = Now()
        rs.Update
        rs.Bookmark = rs.LastModified
        id = rs!ref_id
        rs.Close
        RefAudit db, id, "Added", RefLabelOf(number, title, revision) & IIf(Len(appCode) > 0, ", from " & appCode, "") & ".", appCode
        added = True
    End If
    ws.CommitTrans: inTrans = False
    RefAdd = id
    Exit Function
Fail:
    If inTrans Then RefRollback ws
    RaiseAgain "modReferences.RefAdd", title
End Function

' The row_version of reference id (0 when there is none): pass it back to
' RefUpdate, so a change someone else saved in the meantime isn't overwritten.
Public Function RefVersion(ByVal id As Long) As Long
    RefVersion = Nz(DLookup("row_version", REF_TABLE, "ref_id=" & id), 0)
End Function

' Change reference id: its fields, status (active, superseded or cancelled)
' and, when superseded, the reference that replaced it. version is the
' row_version read when the change began (RefVersion). Library editors and
' Personnel admins only. Refused (error 5, with the reason) for anyone else,
' a field that breaks the rules, a change someone else saved first, a
' number and revision another reference already has, or a supersede that
' points at itself, nothing, or a reference superseded by this one.
' appCode goes in the audit row.
Public Sub RefUpdate(ByVal id As Long, ByVal number As String, ByVal title As String, ByVal revision As String, _
        ByVal pubDate As Variant, ByVal link As String, ByVal marking As String, ByVal notes As String, _
        ByVal status As String, ByVal supersededBy As Long, ByVal version As Long, Optional ByVal appCode As String = "")
    Dim ws As DAO.Workspace, db As DAO.Database, rs As DAO.Recordset, problem As String, inTrans As Boolean
    Dim before As String, after As String, oldStatus As String, what As String, action As String, who As String, toSup As Variant
    On Error GoTo Fail
    If Not RefIsEditor() Then Err.Raise 5, "modReferences", _
        "Only Library editors and Personnel admins can change a reference. Ask one to make the change."
    problem = RefInputProblem(number, title, revision, pubDate, link, marking, notes)
    If Len(problem) > 0 Then Err.Raise 5, "modReferences", problem
    status = LCase$(Trim$(status))
    If Len(status) = 0 Then status = REF_ACTIVE
    If status <> REF_ACTIVE And status <> REF_SUPERSEDED And status <> REF_CANCELLED Then _
        Err.Raise 5, "modReferences", "The status must be active, superseded or cancelled."
    If status <> REF_SUPERSEDED Then supersededBy = 0
    number = Trim$(number): title = RefOneSpace(title): revision = Trim$(revision)
    link = RefLinkText(link): marking = RefMarkingText(marking): notes = Trim$(notes)
    who = RefActorName()
    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb
    ws.BeginTrans: inTrans = True
    Set rs = db.OpenRecordset("SELECT * FROM " & REF_TABLE & " WHERE ref_id=" & id, dbOpenDynaset)
    If rs.EOF Then Err.Raise 5, "modReferences", "No such reference."
    If rs!row_version <> version Then Err.Raise 5, "modReferences", Nz(rs!updated_by, "Someone") & _
        " changed this reference after you opened it. Close it and open it again to see their change, then make yours again."
    If RefFindSame(number, title, revision, id) <> 0 Then Err.Raise 5, "modReferences", _
        RefLabelOf(number, title, revision) & " is already in the library."
    If status = REF_SUPERSEDED Then
        If supersededBy <= 0 Then Err.Raise 5, "modReferences", "Choose the reference that replaced it."
        If supersededBy = id Then Err.Raise 5, "modReferences", "A reference can't supersede itself."
        toSup = DLookup("superseded_by", REF_TABLE, "ref_id=" & supersededBy)
        If DCount("*", REF_TABLE, "ref_id=" & supersededBy) = 0 Then Err.Raise 5, "modReferences", _
            "Choose the reference that replaced it from the library."
        If Nz(toSup, 0) = id Then Err.Raise 5, "modReferences", RefLabel(supersededBy) & _
            " is marked as superseded by this one; two references can't supersede each other."
    End If
    before = RefLabelOf(Nz(rs!ref_number, ""), Nz(rs!ref_title, ""), Nz(rs!ref_revision, ""))
    oldStatus = Nz(rs!ref_status, REF_ACTIVE)
    rs.Edit
    rs!ref_number = RefNull(number): rs!ref_title = title: rs!ref_revision = RefNull(revision)
    If IsNull(pubDate) Then rs!pub_date = Null Else rs!pub_date = DateValue(CDate(pubDate))
    rs!ref_status = status
    If supersededBy > 0 Then rs!superseded_by = supersededBy Else rs!superseded_by = Null
    rs!doc_link = RefNull(link): rs!ref_marking = marking: rs!ref_notes = RefNull(notes)
    rs!row_version = rs!row_version + 1: rs!updated_by = who: rs!updated_at = Now()
    rs.Update
    rs.Close
    after = RefLabelOf(number, title, revision)
    If status <> oldStatus Then what = "status " & oldStatus & " -> " & status
    If StrComp(before, after, vbBinaryCompare) <> 0 Then what = what & IIf(Len(what) > 0, "; ", "") & "was " & before
    If status = oldStatus Then
        action = "Changed"
    ElseIf status = REF_ACTIVE Then
        action = "Restored"
    Else
        action = StrConv(status, vbProperCase)
    End If
    RefAudit db, id, action, after & IIf(Len(what) > 0, ": " & what, "") & ".", appCode
    ws.CommitTrans: inTrans = False
    Exit Sub
Fail:
    If inTrans Then RefRollback ws
    RaiseAgain "modReferences.RefUpdate", CStr(id)
End Sub

' One kit_reference_audit row, in the caller's transaction: who (the
' signed-in person), what, and from which Windows account and PC.
Private Sub RefAudit(db As DAO.Database, ByVal refId As Long, ByVal action As String, ByVal detail As String, ByVal appCode As String)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(REF_AUDIT_TABLE, dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!logged_at = Now()
    If SignedInPersonID() > 0 Then rs!actor_id = SignedInPersonID()
    rs!actor_name = RefNull(RefActorName())
    If refId > 0 Then rs!ref_id = refId
    rs!action = Left$(action, 30)
    rs!detail = Left$(detail, 255)
    rs!app_key = RefNull(Left$(appCode, 30))
    rs!windows_account = RefNull(Left$(WinAccount(), 150))
    rs!device = RefNull(Left$(WinComputer(), 100))
    rs.Update
    rs.Close
End Sub

' The history of reference id, newest first, for a list: When, Who, What, App, PC.
Public Function RefHistorySql(ByVal id As Long) As String
    RefHistorySql = "SELECT audit_id, Format(logged_at,'yyyy-mm-dd hh:nn') AS [When], actor_name AS [Who], [action] & ': ' & detail AS [What], " & _
        "app_key AS [App], device AS [PC] FROM " & REF_AUDIT_TABLE & " WHERE ref_id=" & id & " ORDER BY audit_id DESC"
End Function

' The signed-in person's name as the library stores it ("LT J. Smith").
Private Function RefActorName() As String
    On Error Resume Next                        ' TempVars may be unset
    RefActorName = Left$(Nz(TempVars("PersonName"), Nz(TempVars("Username"), "")), 100)
End Function

Private Sub RefRollback(ws As DAO.Workspace)
    Dim n As Long, d As String, src As String
    n = Err.Number: d = Err.Description: src = Err.Source
    On Error Resume Next
    ws.Rollback
    On Error GoTo 0
    Err.Number = n: Err.Description = d: Err.Source = src
End Sub

Private Function RefNull(ByVal s As String) As Variant
    If Len(s) = 0 Then RefNull = Null Else RefNull = s
End Function

Private Function RefHasTable(db As DAO.Database, ByVal nm As String) As Boolean
    Dim td As DAO.TableDef
    For Each td In db.TableDefs
        If StrComp(td.Name, nm, vbTextCompare) = 0 Then RefHasTable = True: Exit Function
    Next
End Function

' ---------------------------------------------------------------------------
' Opening a document at a page
' ---------------------------------------------------------------------------
'
' A citation keeps a page number beside its paragraph (a PDF can be addressed
' by page, not by paragraph). A PDF on the shared drive opens in Microsoft
' Edge as file:///...#page=N (Edge is on every Windows 10/11 PC, and Adobe
' Reader's own command line takes the page differently); a web link opens as
' link#page=N in Edge, which honors it too. Anything else opens in its usual
' program. Only the library's own link is opened, and only document types
' (REF_OPEN_TYPES), never programs or scripts.

' What RefOpen would open for reference id at page (0 for none): target is
' set to the file or address, and the answer is "" when it can be opened, or
' the reason it can't. page outside 1 to REF_PAGE_MAX counts as none. Opens
' nothing, so the tests can check it.
Public Function RefOpenCheck(ByVal id As Long, ByVal page As Long, target As String) As String
    Dim link As String, rest As String, ext As String, rs As DAO.Recordset, lbl As String
    target = ""
    If page < 1 Or page > REF_PAGE_MAX Then page = 0
    Set rs = CurrentDb.OpenRecordset("SELECT ref_number, ref_title, ref_revision, doc_link FROM " & REF_TABLE & " WHERE ref_id=" & id, dbOpenSnapshot)
    If rs.EOF Then rs.Close: RefOpenCheck = "No such reference.": Exit Function
    lbl = RefLabelOf(Nz(rs!ref_number, ""), Nz(rs!ref_title, ""), Nz(rs!ref_revision, ""))
    link = Nz(rs!doc_link, "")
    rs.Close
    If Len(link) = 0 Then RefOpenCheck = lbl & " has no document linked in the library.": Exit Function
    If LCase$(link) Like "http://?*" Or LCase$(link) Like "https://?*" Then
        target = link
        If page > 0 Then
            If InStr(target, "#") > 0 Then target = Left$(target, InStr(target, "#") - 1)
            target = target & "#page=" & page
        End If
        Exit Function
    End If
    ' The drive's colon only: anything else is a stream (x.pdf:evil) or a device path.
    rest = link
    If link Like "[A-Za-z]:\*" Then rest = Mid$(link, 3)
    If InStr(rest, ":") > 0 Or Left$(link, 4) = "\\?\" Or Left$(link, 4) = "\\.\" Or InStr(link, "/") > 0 Then
        RefOpenCheck = "That link is not a path this PC can open.": Exit Function
    End If
    ext = RefExtension(link)
    If InStr(REF_OPEN_TYPES, "." & ext & ".") = 0 Or Len(ext) = 0 Then
        RefOpenCheck = IIf(Len(ext) > 0, "." & ext & " files", "Files without a type") & _
            " are not opened from the library. Copy the path instead.": Exit Function
    End If
    If Not FileExists(link) Then
        RefOpenCheck = "The document is not at " & link & " (or the drive is not connected on this PC).": Exit Function
    End If
    target = link
    If ext = "pdf" And page > 0 Then target = RefFileUri(link) & "#page=" & page
End Function

' Open reference id's document at page (0 for none). Returns "" when it
' opened at that page, or a sentence for the person: why it can't open, or
' that it opened at page 1 (no Edge on this PC).
Public Function RefOpen(ByVal id As Long, ByVal page As Long) As String
    Dim target As String, problem As String, edge As String, note As String
    On Error GoTo Fail
    problem = RefOpenCheck(id, page, target)
    If Len(problem) > 0 Then RefOpen = problem: Exit Function
    If page < 1 Or page > REF_PAGE_MAX Then page = 0
    If page > 0 And (Left$(LCase$(target), 5) = "file:" Or (Left$(LCase$(target), 4) = "http" And InStr(target, "#page=") > 0)) Then
        edge = RefEdgePath()
        If Len(edge) > 0 Then
            Shell """" & edge & """ """ & target & """", vbNormalFocus
            Exit Function
        End If
        If Left$(LCase$(target), 5) = "file:" Then
            note = "Microsoft Edge is not on this PC, so the PDF opened in its usual program at page 1. Go to page " & page & "."
            target = DLookup("doc_link", REF_TABLE, "ref_id=" & id)
        End If
    End If
    ' Shell.Application opens a document in its usual program and a web
    ' address in the default browser, without Access's hyperlink warning.
    CreateObject("Shell.Application").ShellExecute target, "", "", "open", 1
    RefOpen = note
    Exit Function
Fail:
    RefOpen = "Windows could not open it: " & Err.Description
End Function

' The file type of path without its dot, in lower case ("pdf"), or "".
Private Function RefExtension(ByVal path As String) As String
    Dim nm As String
    nm = Mid$(path, InStrRev(path, "\") + 1)
    If InStrRev(nm, ".") > 0 Then RefExtension = LCase$(Mid$(nm, InStrRev(nm, ".") + 1))
End Function

' Microsoft Edge on this PC, or "". Looked up in App Paths (32-bit Access on
' 64-bit Windows reads the WOW6432Node copy, which Edge also writes), then in
' both Program Files folders.
Public Function RefEdgePath() As String
    Dim sh As Object, v As Variant, p As String
    On Error Resume Next
    Set sh = CreateObject("WScript.Shell")
    For Each v In Array("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe\", _
                        "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe\")
        p = ""
        p = sh.RegRead(v)
        If Len(p) > 0 Then
            p = Replace(p, """", "")
            If FileExists(p) Then RefEdgePath = p: Exit Function
        End If
    Next
    For Each v In Array(Environ$("ProgramFiles(x86)"), Environ$("ProgramW6432"), Environ$("ProgramFiles"))
        If Len(v) > 0 Then
            p = v & "\Microsoft\Edge\Application\msedge.exe"
            If FileExists(p) Then RefEdgePath = p: Exit Function
        End If
    Next
End Function

' The file: address of a path, as .NET's Uri.AbsoluteUri writes it:
' S:\Pubs\x y.pdf gives file:///S:/Pubs/x%20y.pdf, \\srv\share\x.pdf gives
' file://srv/share/x.pdf. Characters outside the safe set are written as
' %XX of their UTF-8 bytes.
Public Function RefFileUri(ByVal path As String) As String
    Dim body As String, i As Long, ch As String, out As String, code As Long, lo As Long
    If Left$(path, 2) = "\\" Then
        body = Mid$(path, 3): out = "file://"
    Else
        body = path: out = "file:///"
    End If
    body = Replace(body, "\", "/")
    i = 1
    Do While i <= Len(body)
        ch = Mid$(body, i, 1)
        code = AscW(ch) And &HFFFF&
        If (ch Like "[A-Za-z0-9]") Or InStr("-._~/:!$&'()*+,;=@", ch) > 0 Then
            out = out & ch
        ElseIf code >= &HD800& And code <= &HDBFF& And i < Len(body) Then
            lo = AscW(Mid$(body, i + 1, 1)) And &HFFFF&
            out = out & RefUtf8Hex(&H10000 + (code - &HD800&) * &H400& + (lo - &HDC00&))
            i = i + 1
        Else
            out = out & RefUtf8Hex(code)
        End If
        i = i + 1
    Loop
    RefFileUri = out
End Function

' A code point as %XX of its UTF-8 bytes.
Private Function RefUtf8Hex(ByVal cp As Long) As String
    If cp < &H80& Then
        RefUtf8Hex = RefHexByte(cp)
    ElseIf cp < &H800& Then
        RefUtf8Hex = RefHexByte(&HC0& Or (cp \ &H40&)) & RefHexByte(&H80& Or (cp And &H3F&))
    ElseIf cp < &H10000 Then
        RefUtf8Hex = RefHexByte(&HE0& Or (cp \ &H1000&)) & RefHexByte(&H80& Or ((cp \ &H40&) And &H3F&)) & _
            RefHexByte(&H80& Or (cp And &H3F&))
    Else
        RefUtf8Hex = RefHexByte(&HF0& Or (cp \ &H40000)) & RefHexByte(&H80& Or ((cp \ &H1000&) And &H3F&)) & _
            RefHexByte(&H80& Or ((cp \ &H40&) And &H3F&)) & RefHexByte(&H80& Or (cp And &H3F&))
    End If
End Function

Private Function RefHexByte(ByVal b As Long) As String
    RefHexByte = "%" & Right$("0" & Hex$(b), 2)
End Function

' The caption of an open button beside a citation (the web's RefOpenButton):
' "Open p. 12", or "Open document" with no page.
Public Function RefOpenCaption(ByVal page As Long) As String
    If page >= 1 And page <= REF_PAGE_MAX Then RefOpenCaption = "Open p. " & page Else RefOpenCaption = "Open document"
End Function

' Why page can't be a citation's page ("" when it can; blank is no page).
Public Function RefPageProblem(ByVal page As Variant) As String
    If IsNull(page) Then Exit Function
    If Len(Trim$(CStr(page))) = 0 Then Exit Function
    If Not IsNumeric(page) Then RefPageProblem = "The page must be a whole number from 1 to " & REF_PAGE_MAX & ".": Exit Function
    If CDbl(page) <> Int(CDbl(page)) Or CDbl(page) < 1 Or CDbl(page) > REF_PAGE_MAX Then _
        RefPageProblem = "The page must be a whole number from 1 to " & REF_PAGE_MAX & "."
End Function
