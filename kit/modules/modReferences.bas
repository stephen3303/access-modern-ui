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
'   kit_reference         one row per publication, kept current (Stephen,
'                         2026-11-17): ref_number, the short title (OPNAVINST
'                         3120.32), ref_title, ref_revision (the current
'                         change, CH-1), pub_date (the date of the current
'                         revision), ref_status (active, superseded, cancelled),
'                         superseded_by (the ref_id that replaced it),
'                         doc_link (a shared-drive path or an http(s)
'                         address), ref_marking, ref_notes, source_app (the
'                         app it was first added from), row_version (+1 on
'                         every change, for the stale-save check), created_*
'                         and updated_*.
'   kit_reference_detail  ref_id, ref_publisher (OPNAV), ref_edition (the
'                         current edition, D): its own table so an older file
'                         gains them with CREATE TABLE while others have it
'                         open (ALTER TABLE needs the file to itself).
'   kit_reference_editor  who may change and retire references (Library
'                         editors). Personnel admins always may, and they
'                         choose the editors.
'   kit_reference_audit   every add, change, status change and editor change,
'                         with who and from which PC.
'   kit_reference_setting name/value, used by the web's first-run import.
' All but the settings table are linked into the front end.
'
' Stephen's fields (2026-11-17), with the columns that keep them: Short Title
' (Publication Number) ref_number, Title ref_title, Publisher ref_publisher,
' Classification ref_marking, Current Edition ref_edition, Current Change
' ref_revision, Date of current Revision pub_date. Import and export use
' those seven columns in that order, then Document, Notes, Status and
' Superseded by (RefExport; RefImportRead, RefImportCheck, RefImportApply).
'
' Who may do what (Stephen, 2026-11-09): anyone signed in to an app adds a
' reference, so drafting is never held up; Library editors and Personnel
' admins change, retire and supersede. Nobody deletes: a reference other
' files cite is retired (cancelled or superseded), never removed.
'
' The same publication twice: two references are the same when their short
' titles match (case, spaces, dashes, underscores and dots ignored), whatever
' their edition and change, since one record is kept current through new
' editions and changes; a reference without a short title matches on its
' title instead. RefAdd hands back the one already there rather than adding
' it again.
'
' A citation prints the short title with its edition (RefCiteText): a one-
' or two-letter edition joins it (OPNAVINST 3120.32 + D = OPNAVINST
' 3120.32D), any other follows after a space; the title when there is no
' short title.
'
' Startup, after KitPersonnelStart (modReferencesUI wraps this):
'   ReferencesStatus()     links the file found beside the Personnel file
' For app code (a citation keeps a ref_id and a page number in the app's own
' data file, as the web's comment_citation does):
'   RefAdd number, title, revision, pubDate, link, marking, notes, appCode, added, publisher, edition
'   RefUpdate id, ..., status, supersededBy, version, appCode, publisher, edition     editors only
'   RefLabel(id)            "NUMBERED Title (CHANGE)", for lists
'   RefCiteText(id)         "OPNAVINST 3120.32D", as a citation prints
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
Public Const REF_DETAIL_TABLE As String = "kit_reference_detail"

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
        If KitBackendReachable(REF_TABLE) Then
            ' A front end linked before publisher and edition came: link (and
            ' make, when the file lacks it) kit_reference_detail too.
            If Len(KitBackendPath(REF_DETAIL_TABLE)) = 0 Then LinkReferencesFile ReferencesPath()
            ReferencesStatus = REFS_READY
            Exit Function
        End If
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
    KitLinkTables path, Array(REF_TABLE, REF_EDITOR_TABLE, REF_AUDIT_TABLE, REF_DETAIL_TABLE)
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
    If Not RefHasTable(db, REF_DETAIL_TABLE) Then
        db.Execute "CREATE TABLE " & REF_DETAIL_TABLE & " (ref_id LONG CONSTRAINT pk_kit_reference_detail PRIMARY KEY, " & _
            "ref_publisher TEXT(100), ref_edition TEXT(30))", dbFailOnError
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

' The reference that is the same publication (see the top of the module),
' leaving out exceptId; 0 when there is none: the same short title, an
' active one first, whatever the edition and change; with no short title,
' the same title. revision is no longer compared (kept so callers need not
' change).
Public Function RefFindSame(ByVal number As String, ByVal title As String, Optional ByVal revision As String = "", _
        Optional ByVal exceptId As Long = 0) As Long
    Dim rs As DAO.Recordset, nk As String, tk As String, same As Boolean
    nk = RefKey(number): tk = RefTitleKey(title)
    Set rs = CurrentDb.OpenRecordset("SELECT ref_id, ref_number, ref_title, ref_status FROM " & REF_TABLE & " ORDER BY ref_id", dbOpenSnapshot)
    Do While Not rs.EOF
        same = False
        If rs!ref_id <> exceptId Then
            If Len(nk) > 0 Then
                same = (StrComp(RefKey(Nz(rs!ref_number, "")), nk, vbBinaryCompare) = 0)
            ElseIf Len(Nz(rs!ref_number, "")) = 0 Then
                same = (StrComp(RefTitleKey(Nz(rs!ref_title, "")), tk, vbBinaryCompare) = 0)
            End If
        End If
        If same Then
            If Nz(rs!ref_status, REF_ACTIVE) = REF_ACTIVE Then RefFindSame = rs!ref_id: Exit Do
            If RefFindSame = 0 Then RefFindSame = rs!ref_id
        End If
        rs.MoveNext
    Loop
    rs.Close
End Function

' The short title with its edition, as a citation prints it (the web's
' Join-KitRefNumber): a one- or two-letter edition joins the number
' (OPNAVINST 3120.32 + D = OPNAVINST 3120.32D), any other follows after a
' space (NAVSEA S9086 Rev 3). "" with no short title.
Public Function RefNumberOf(ByVal number As String, ByVal edition As String) As String
    number = Trim$(number): edition = Trim$(edition)
    If Len(number) = 0 Then Exit Function
    If Len(edition) = 0 Then
        RefNumberOf = number
    ElseIf edition Like "[A-Za-z]" Or edition Like "[A-Za-z][A-Za-z]" Then
        RefNumberOf = number & edition
    Else
        RefNumberOf = number & " " & edition
    End If
End Function

' What a citation prints: the short title with its edition, or the title
' when there is no short title (Stephen, 2026-11-17).
Public Function RefCiteOf(ByVal number As String, ByVal edition As String, ByVal title As String) As String
    RefCiteOf = RefNumberOf(number, edition)
    If Len(RefCiteOf) = 0 Then RefCiteOf = Trim$(title)
End Function

' "NUMBERED Title (CHANGE)", for lists (the web's Format-KitReference): the
' short title with its edition, the title, the change in brackets.
Public Function RefLabelOf(ByVal number As String, ByVal title As String, ByVal revision As String, _
        Optional ByVal edition As String = "") As String
    Dim s As String, n As String
    n = RefNumberOf(number, edition)
    s = title
    If Len(n) > 0 Then s = n & " " & title
    If Len(revision) > 0 Then s = s & " (" & revision & ")"
    RefLabelOf = s
End Function

' kit_reference (r) with its publisher and edition (d), for a FROM clause.
Public Function RefFromSql() As String
    RefFromSql = REF_TABLE & " AS r LEFT JOIN " & REF_DETAIL_TABLE & " AS d ON r.ref_id = d.ref_id"
End Function

' The label of reference id, or "" when there is none.
Public Function RefLabel(ByVal id As Long) As String
    Dim rs As DAO.Recordset
    If id <= 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT r.ref_number, r.ref_title, r.ref_revision, d.ref_edition FROM " & RefFromSql() & _
        " WHERE r.ref_id=" & id, dbOpenSnapshot)
    If Not rs.EOF Then RefLabel = RefLabelOf(Nz(rs!ref_number, ""), Nz(rs!ref_title, ""), Nz(rs!ref_revision, ""), Nz(rs!ref_edition, ""))
    rs.Close
End Function

' What a citation of reference id prints (RefCiteOf), or "" when there is none.
Public Function RefCiteText(ByVal id As Long) As String
    Dim rs As DAO.Recordset
    If id <= 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT r.ref_number, r.ref_title, d.ref_edition FROM " & RefFromSql() & _
        " WHERE r.ref_id=" & id, dbOpenSnapshot)
    If Not rs.EOF Then RefCiteText = RefCiteOf(Nz(rs!ref_number, ""), Nz(rs!ref_edition, ""), Nz(rs!ref_title, ""))
    rs.Close
End Function

' The publisher and the edition of reference id ("" for none).
Public Function RefPublisher(ByVal id As Long) As String
    RefPublisher = Nz(DLookup("ref_publisher", REF_DETAIL_TABLE, "ref_id=" & id), "")
End Function

Public Function RefEdition(ByVal id As Long) As String
    RefEdition = Nz(DLookup("ref_edition", REF_DETAIL_TABLE, "ref_id=" & id), "")
End Function

' The short title with its edition as an Access SQL expression over r and d
' (RefFromSql), as RefNumberOf. Like ignores case in Access SQL.
Private Function RefNumberSql() As String
    RefNumberSql = "IIf(Nz(r.ref_number,'')='','',r.ref_number & IIf(Nz(d.ref_edition,'')='',''," & _
        "IIf(d.ref_edition Like '[A-Z]' Or d.ref_edition Like '[A-Z][A-Z]',d.ref_edition,' ' & d.ref_edition)))"
End Function

' The label as an Access SQL expression over r and d (RefFromSql).
Public Function RefLabelSql() As String
    RefLabelSql = "IIf(" & RefNumberSql() & "='',''," & RefNumberSql() & " & ' ') & r.ref_title & " & _
        "IIf(Nz(r.ref_revision,'')='','',' (' & r.ref_revision & ')')"
End Function

' The active references for a PickBox (modMenu): ref_id, the label, and
' "Document" in muted text when a document is linked.
Public Function RefPickSql() As String
    RefPickSql = "SELECT r.ref_id, " & RefLabelSql() & " AS ref_label, IIf(Nz(r.doc_link,'')='','','Document') AS ref_doc FROM " & _
        RefFromSql() & " WHERE r.ref_status='" & REF_ACTIVE & "'"
End Function

' The rows of the library page, in Stephen's order: ref_id, Short title,
' Title, Publisher, Classification, Edition, Change, Revised, Status,
' Document; matching term
' (short title, title, publisher, edition, change or notes) and, unless
' withRetired, active only. term is matched as typed: [ * ? # are not wildcards.
Public Function RefLibrarySql(ByVal term As String, ByVal withRetired As Boolean) As String
    Dim w As String, t As String
    If Not withRetired Then w = "r.ref_status='" & REF_ACTIVE & "'"
    term = Trim$(term)
    If Len(term) > 0 Then
        t = "'*" & RefLikeText(term) & "*'"
        w = IIf(Len(w) > 0, w & " AND ", "") & "(r.ref_number LIKE " & t & " OR r.ref_title LIKE " & t & _
            " OR r.ref_revision LIKE " & t & " OR r.ref_notes LIKE " & t & " OR d.ref_publisher LIKE " & t & _
            " OR " & RefNumberSql() & " LIKE " & t & ")"
    End If
    RefLibrarySql = "SELECT r.ref_id, r.ref_number AS [Short title], r.ref_title AS [Title], d.ref_publisher AS [Publisher], " & _
        "r.ref_marking AS [Classification], d.ref_edition AS [Edition], r.ref_revision AS [Change], " & _
        "Format(r.pub_date,'yyyy-mm-dd') AS [Revised], StrConv(r.ref_status,3) AS [Status], " & _
        "IIf(Nz(r.doc_link,'')='','','Linked') AS [Document] FROM " & RefFromSql() & _
        IIf(Len(w) > 0, " WHERE " & w, "") & " ORDER BY r.ref_number, r.ref_title"
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
        ByVal pubDate As Variant, ByVal link As String, ByVal marking As String, ByVal notes As String, _
        Optional ByVal publisher As String = "", Optional ByVal edition As String = "") As String
    Dim p As String
    title = RefOneSpace(title)
    If Len(title) = 0 Then p = "Enter the reference's title."
    If Len(p) = 0 And Len(title) > 255 Then p = "A title can be at most 255 characters."
    If Len(p) = 0 And Len(Trim$(number)) > 100 Then p = "A short title can be at most 100 characters."
    If Len(p) = 0 And Len(RefOneSpace(publisher)) > 100 Then p = "A publisher can be at most 100 characters."
    If Len(p) = 0 And Len(Trim$(edition)) > 30 Then p = "An edition can be at most 30 characters."
    If Len(p) = 0 And Len(Trim$(revision)) > 30 Then p = "A change can be at most 30 characters."
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
        Optional added As Boolean, Optional ByVal publisher As String = "", Optional ByVal edition As String = "") As Long
    Dim ws As DAO.Workspace, db As DAO.Database, rs As DAO.Recordset, problem As String, id As Long, inTrans As Boolean, who As String
    On Error GoTo Fail
    added = False
    If SignedInPersonID() = 0 Then Err.Raise 5, "modReferences", "Sign in first to add a reference."
    problem = RefInputProblem(number, title, revision, pubDate, link, marking, notes, publisher, edition)
    If Len(problem) > 0 Then Err.Raise 5, "modReferences", problem
    number = Trim$(number): title = RefOneSpace(title): revision = Trim$(revision)
    link = RefLinkText(link): marking = RefMarkingText(marking): notes = Trim$(notes)
    publisher = RefOneSpace(publisher): edition = Trim$(edition)
    who = RefActorName()
    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb
    ws.BeginTrans: inTrans = True
    id = RefFindSame(number, title)
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
        RefSetDetail db, id, publisher, edition
        RefAudit db, id, "Added", RefLabelOf(number, title, revision, edition) & IIf(Len(appCode) > 0, ", from " & appCode, "") & ".", appCode
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
' short title another reference already has, or a supersede that points at
' itself, nothing, or a reference superseded by this one. appCode goes in
' the audit row. publisher and edition left out keep what the reference has
' (a caller written before they came).
Public Sub RefUpdate(ByVal id As Long, ByVal number As String, ByVal title As String, ByVal revision As String, _
        ByVal pubDate As Variant, ByVal link As String, ByVal marking As String, ByVal notes As String, _
        ByVal status As String, ByVal supersededBy As Long, ByVal version As Long, Optional ByVal appCode As String = "", _
        Optional ByVal publisher As Variant, Optional ByVal edition As Variant)
    Dim ws As DAO.Workspace, db As DAO.Database, rs As DAO.Recordset, problem As String, inTrans As Boolean
    Dim before As String, after As String, oldStatus As String, what As String, action As String, who As String, toSup As Variant
    Dim pub As String, ed As String, oldEd As String
    On Error GoTo Fail
    If Not RefIsEditor() Then Err.Raise 5, "modReferences", _
        "Only Library editors and Personnel admins can change a reference. Ask one to make the change."
    oldEd = RefEdition(id)
    If IsMissing(publisher) Then pub = RefPublisher(id) Else pub = RefOneSpace(Nz(publisher, ""))
    If IsMissing(edition) Then ed = oldEd Else ed = Trim$(Nz(edition, ""))
    problem = RefInputProblem(number, title, revision, pubDate, link, marking, notes, pub, ed)
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
    If RefFindSame(number, title, "", id) <> 0 Then Err.Raise 5, "modReferences", _
        IIf(Len(number) > 0, number, title) & " is already in the library: keep one record for it and change that one."
    If status = REF_SUPERSEDED Then
        If supersededBy <= 0 Then Err.Raise 5, "modReferences", "Choose the reference that replaced it."
        If supersededBy = id Then Err.Raise 5, "modReferences", "A reference can't supersede itself."
        toSup = DLookup("superseded_by", REF_TABLE, "ref_id=" & supersededBy)
        If DCount("*", REF_TABLE, "ref_id=" & supersededBy) = 0 Then Err.Raise 5, "modReferences", _
            "Choose the reference that replaced it from the library."
        If Nz(toSup, 0) = id Then Err.Raise 5, "modReferences", RefLabel(supersededBy) & _
            " is marked as superseded by this one; two references can't supersede each other."
    End If
    before = RefLabelOf(Nz(rs!ref_number, ""), Nz(rs!ref_title, ""), Nz(rs!ref_revision, ""), oldEd)
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
    RefSetDetail db, id, pub, ed
    after = RefLabelOf(number, title, revision, ed)
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

' Publisher and edition of reference id (kit_reference_detail), in the
' caller's transaction: the row is replaced, or dropped when both are blank.
Private Sub RefSetDetail(db As DAO.Database, ByVal id As Long, ByVal publisher As String, ByVal edition As String)
    Dim rs As DAO.Recordset
    db.Execute "DELETE FROM " & REF_DETAIL_TABLE & " WHERE ref_id=" & id, dbFailOnError
    If Len(publisher) = 0 And Len(edition) = 0 Then Exit Sub
    Set rs = db.OpenRecordset(REF_DETAIL_TABLE, dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!ref_id = id: rs!ref_publisher = RefNull(Left$(publisher, 100)): rs!ref_edition = RefNull(Left$(edition, 30))
    rs.Update
    rs.Close
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
    Set rs = CurrentDb.OpenRecordset("SELECT doc_link FROM " & REF_TABLE & " WHERE ref_id=" & id, dbOpenSnapshot)
    If rs.EOF Then rs.Close: RefOpenCheck = "No such reference.": Exit Function
    link = Nz(rs!doc_link, "")
    rs.Close
    lbl = RefLabel(id)
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

' ---------------------------------------------------------------------------
' Import and export (Stephen, 2026-11-17)
' ---------------------------------------------------------------------------
'
' The columns, in this order: Short Title (Publication Number), Title,
' Publisher, Classification, Current Edition, Current Change, Date of
' current Revision; then Document, Notes, Status and Superseded by, which a
' file may leave out. The same layout as the web's Reference library page,
' so a sheet exported from one imports into the other. Import reads the
' older headings too (Number, Revision, Date, Marking, SABER's
' instruction_number ...).
'
' Import is checked first (RefImportCheck, nothing written) and written
' only when no row has a problem (RefImportApply, one transaction). One
' record is kept per short title: a row whose short title is in the library
' updates that record (a new edition or change, a new date), and a blank
' cell keeps what the library has.

Private Const REF_EXPORT_QUERY As String = "zz_qryRefExport"
Private Const REF_IMPORT_TABLE As String = "zz_reference_import"

' The export's headings, in Stephen's order, for a sheet's first row.
Public Function RefExportHeadings() As Variant
    RefExportHeadings = Array("Short Title (Publication Number)", "Title", "Publisher", "Classification", "Current Edition", _
        "Current Change", "Date of current Revision", "Document", "Notes", "Status", "Superseded by")
End Function

' The whole library (retired references too) in the export's columns.
Public Function RefExportSql() As String
    Dim h As Variant
    h = RefExportHeadings()
    RefExportSql = "SELECT r.ref_number AS [" & h(0) & "], r.ref_title AS [" & h(1) & "], d.ref_publisher AS [" & h(2) & "], " & _
        "r.ref_marking AS [" & h(3) & "], d.ref_edition AS [" & h(4) & "], r.ref_revision AS [" & h(5) & "], " & _
        "r.pub_date AS [" & h(6) & "], r.doc_link AS [" & h(7) & "], r.ref_notes AS [" & h(8) & "], " & _
        "r.ref_status AS [" & h(9) & "], s.ref_number AS [" & h(10) & "] " & _
        "FROM (" & RefFromSql() & ") LEFT JOIN " & REF_TABLE & " AS s ON r.superseded_by = s.ref_id " & _
        "ORDER BY r.ref_number, r.ref_title"
End Function

' Write the library to path: an Excel workbook when path ends in .xlsx,
' else CSV. Replaces a file already there. Returns the references written.
Public Function RefExport(ByVal path As String) As Long
    Dim n As Long, d As String, src As String
    KitBusy "Exporting the Reference library..."
    On Error GoTo Fail
    If FileExists(path) Then Kill path
    If LCase$(Right$(path, 5)) = ".xlsx" Then
        RefDropQuery
        CurrentDb.CreateQueryDef REF_EXPORT_QUERY, RefExportSql()
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, REF_EXPORT_QUERY, path, True
        RefExport = DCount("*", REF_EXPORT_QUERY)
        RefDropQuery
    Else
        RefExport = ExportSqlToCsv(RefExportSql(), path)
    End If
    KitBusyEnd
    Exit Function
Fail:
    n = Err.Number: d = Err.Description: src = Err.Source        ' clean-up resets Err
    RefDropQuery
    KitBusyEnd
    Err.Number = n: Err.Description = d: Err.Source = src
    RaiseAgain "modReferences.RefExport", path
End Function

Private Sub RefDropQuery()
    On Error Resume Next
    CurrentDb.QueryDefs.Delete REF_EXPORT_QUERY
End Sub

' A heading as it is compared: lower case letters and digits only
' ("Short Title (Publication Number)" is "shorttitlepublicationnumber"), as
' the web's headerKey.
Public Function RefHeaderKey(ByVal s As String) As String
    Dim i As Long, ch As String
    s = LCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[a-z0-9]" Then RefHeaderKey = RefHeaderKey & ch
    Next
End Function

' The headings each field may have in a sheet (after RefHeaderKey), as the
' web's ALIASES.
Private Function RefImportAliases(ByVal field As String) As Variant
    Select Case field
        Case "number": RefImportAliases = Array("shorttitlepublicationnumber", "shorttitle", "number", "refnumber", "referencenumber", _
            "instructionnumber", "publicationnumber", "pubnumber", "docnumber", "designator")
        Case "title": RefImportAliases = Array("title", "reftitle", "instructionname", "publicationtitle", "publication", "name", "subject")
        Case "publisher": RefImportAliases = Array("publisher", "issuedby", "issuer", "originator")
        Case "edition": RefImportAliases = Array("currentedition", "edition")
        Case "revision": RefImportAliases = Array("currentchange", "change", "revision", "rev")
        Case "date": RefImportAliases = Array("dateofcurrentrevision", "date", "publicationdate", "pubdate", "dated", "effectivedate", "revisiondate")
        Case "link": RefImportAliases = Array("link", "documenturl", "url", "documentlink", "path", "document", "file")
        Case "marking": RefImportAliases = Array("marking", "classification")
        Case "notes": RefImportAliases = Array("notes", "note", "remarks", "comments")
        Case "status": RefImportAliases = Array("status")
        Case "sup": RefImportAliases = Array("supersededby", "replacedby")
        Case "active": RefImportAliases = Array("isactive", "active")
    End Select
End Function

' The first of field's headings the row has a value for, or "".
Private Function RefImportPick(ByVal d As Object, ByVal field As String) As String
    Dim k As Variant
    For Each k In RefImportAliases(field)
        If d.Exists(k) Then
            If Len(d(k)) > 0 Then RefImportPick = d(k): Exit Function
        End If
    Next
End Function

' A cell as text: a date as yyyy-mm-dd, trimmed.
Private Function RefCellText(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then RefCellText = Format$(v, "yyyy-mm-dd") Else RefCellText = Trim$(CStr(v))
End Function

' The rows of a .xlsx, .xls or .csv file whose first row names the columns
' (at least Title), as dictionaries with line (its row in the sheet),
' number, title, publisher, marking, edition, revision, date (as written),
' link, notes, status and sup (the short title in Superseded by). Rows with
' nothing in them are left out. Error 5 for a file it can't use.
Public Function RefImportRead(ByVal path As String) As Collection
    Dim raw As New Collection, out As New Collection, d As Object, src As Variant, k As Variant, ext As String
    Dim rs As DAO.Recordset, fld As DAO.Field, row As Object, f As Variant, hasTitle As Boolean, hasAny As Boolean, i As Long, act As String
    Dim n As Long, ds As String, so As String, hdr As Collection
    On Error GoTo Fail
    ext = LCase$(Mid$(path, InStrRev(path, ".") + 1))
    Select Case ext
    Case "csv"
        For Each src In ReadCsvDicts(path)
            Set d = CreateObject("Scripting.Dictionary")
            For Each k In src.Keys
                d(RefHeaderKey(k)) = RefCellText(src(k))
            Next
            raw.Add d
        Next
        ' The headings, even with no rows under them.
        Set d = CreateObject("Scripting.Dictionary")
        Set hdr = ParseCsv(ReadTextFile(path))
        If hdr.Count > 0 Then
            For Each k In hdr(1)
                d(RefHeaderKey(k)) = ""
            Next
        End If
        hasTitle = RefHasTitleHeading(d)
    Case "xlsx", "xls"
        On Error Resume Next                    ' left from an earlier import that failed
        CurrentDb.TableDefs.Delete REF_IMPORT_TABLE
        On Error GoTo Fail
        DoCmd.TransferSpreadsheet acImport, IIf(ext = "xls", acSpreadsheetTypeExcel8, acSpreadsheetTypeExcel12Xml), REF_IMPORT_TABLE, path, True
        Set rs = CurrentDb.OpenRecordset(REF_IMPORT_TABLE, dbOpenSnapshot)
        Set d = CreateObject("Scripting.Dictionary")
        For Each fld In rs.Fields
            d(RefHeaderKey(fld.Name)) = ""
        Next
        hasTitle = RefHasTitleHeading(d)
        Do Until rs.EOF
            Set d = CreateObject("Scripting.Dictionary")
            For Each fld In rs.Fields
                d(RefHeaderKey(fld.Name)) = RefCellText(fld.Value)
            Next
            raw.Add d
            rs.MoveNext
        Loop
        rs.Close
        CurrentDb.TableDefs.Delete REF_IMPORT_TABLE
    Case Else
        Err.Raise 5, "modReferences", "Import reads .xlsx, .xls or .csv files."
    End Select
    If Not hasTitle Then Err.Raise 5, "modReferences", _
        "The first row must name the columns, with at least Title. Export the library to see the layout."
    For Each d In raw
        i = i + 1
        Set row = CreateObject("Scripting.Dictionary")
        row("line") = i + 1                         ' row 1 is the headings
        hasAny = False
        For Each f In Array("number", "title", "publisher", "marking", "edition", "revision", "date", "link", "notes", "status", "sup")
            row(f) = RefImportPick(d, f)
            If Len(row(f)) > 0 Then hasAny = True
        Next
        row("status") = LCase$(row("status"))
        act = LCase$(RefImportPick(d, "active"))
        If Len(row("status")) = 0 And (act = "false" Or act = "no" Or act = "0" Or act = "n") Then row("status") = REF_CANCELLED
        row("marking") = UCase$(row("marking"))
        If hasAny Then out.Add row
    Next
    Set RefImportRead = out
    Exit Function
Fail:
    n = Err.Number: ds = Err.Description: so = Err.Source
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    CurrentDb.TableDefs.Delete REF_IMPORT_TABLE
    On Error GoTo 0
    Err.Number = n: Err.Description = ds: Err.Source = so
    RaiseAgain "modReferences.RefImportRead", path
End Function

Private Function RefHasTitleHeading(ByVal headings As Object) As Boolean
    Dim k As Variant
    For Each k In RefImportAliases("title")
        If headings.Exists(k) Then RefHasTitleHeading = True: Exit Function
    Next
End Function

' A row's date: Null when blank, the date, or "bad".
Private Function RefImportDate(ByVal s As String) As Variant
    If Len(s) = 0 Then RefImportDate = Null: Exit Function
    If s Like "####-##-##*" Then s = Left$(s, 10)
    If IsDate(s) Then RefImportDate = DateValue(CDate(s)) Else RefImportDate = "bad"
End Function

' The key that makes two rows the same publication: the short title's
' (RefKey), or "t:" and the title's when there is none.
Private Function RefRowKey(ByVal number As String, ByVal title As String) As String
    If Len(RefKey(number)) > 0 Then RefRowKey = RefKey(number) Else RefRowKey = "t:" & RefTitleKey(title)
End Function

' Check rows read by RefImportRead against the library, writing nothing.
' Each row gets result ("add", "update", "unchanged" or "error"), note (why
' an error), changes (what an update changes, comma separated) and id (the
' library reference it updates). Library editors and Personnel admins only.
Public Function RefImportCheck(ByVal rows As Collection) As Collection
    Dim row As Object, seen As Object, rs As DAO.Recordset, k As String, problem As String, ch As String
    Dim dt As Variant, cur As Long, other As Object, libC As Long, fileC As Long, fileId As Long, fileFree As Boolean, target As Long
    If Not RefIsEditor() Then Err.Raise 5, "modReferences", "Only Library editors and Personnel admins can import references."
    Set seen = CreateObject("Scripting.Dictionary")
    For Each row In rows
        row("result") = "add": row("note") = "": row("changes") = "": row("id") = 0
        row("title") = RefOneSpace(row("title")): row("publisher") = RefOneSpace(row("publisher"))
        row("link") = RefLinkText(row("link"))
        dt = RefImportDate(row("date"))
        problem = ""
        If Not IsNull(dt) Then
            If VarType(dt) = vbString Then problem = "The date must be a real date (yyyy-mm-dd)."
        End If
        If Len(problem) = 0 Then problem = RefInputProblem(row("number"), row("title"), row("revision"), IIf(IsNull(dt), Null, dt), _
            row("link"), row("marking"), row("notes"), row("publisher"), row("edition"))
        If Len(problem) = 0 And Len(row("status")) > 0 Then
            If row("status") <> REF_ACTIVE And row("status") <> REF_SUPERSEDED And row("status") <> REF_CANCELLED Then _
                problem = "Status must be active, superseded or cancelled (or blank)."
        End If
        If Len(problem) = 0 And Len(row("sup")) > 0 Then
            If Len(row("status")) > 0 And row("status") <> REF_SUPERSEDED Then
                problem = """Superseded by"" is only for a superseded reference."
            Else
                row("status") = REF_SUPERSEDED
            End If
        End If
        If Len(problem) = 0 And row("status") = REF_SUPERSEDED And Len(row("sup")) = 0 Then _
            problem = "A superseded reference needs ""Superseded by"" (the short title of the one that replaced it)."
        k = RefRowKey(row("number"), row("title"))
        If Len(problem) = 0 And seen.Exists(k) Then problem = "The same publication is on row " & seen(k) & " of this file."
        If Len(problem) > 0 Then
            row("result") = "error": row("note") = problem
        Else
            seen(k) = row("line")
            row("dateValue") = dt
            cur = RefFindSame(row("number"), row("title"))
            If cur > 0 Then
                row("id") = cur
                Set rs = CurrentDb.OpenRecordset("SELECT r.*, d.ref_publisher, d.ref_edition FROM " & RefFromSql() & " WHERE r.ref_id=" & cur, dbOpenSnapshot)
                row("curStatus") = Nz(rs!ref_status, REF_ACTIVE): row("curSup") = Nz(rs!superseded_by, 0)
                ch = ""
                If StrComp(row("title"), Nz(rs!ref_title, ""), vbBinaryCompare) <> 0 Then ch = ch & ",title"
                ' One record is kept current: a new edition or change updates it; blank cells keep what it has.
                If Len(row("publisher")) > 0 And StrComp(row("publisher"), Nz(rs!ref_publisher, ""), vbBinaryCompare) <> 0 Then ch = ch & ",publisher"
                If Len(row("edition")) > 0 And StrComp(row("edition"), Nz(rs!ref_edition, ""), vbBinaryCompare) <> 0 Then ch = ch & ",edition"
                If Len(row("revision")) > 0 And StrComp(row("revision"), Nz(rs!ref_revision, ""), vbBinaryCompare) <> 0 Then ch = ch & ",change"
                If Not IsNull(dt) Then
                    If IsNull(rs!pub_date) Then
                        ch = ch & ",date"
                    ElseIf DateValue(rs!pub_date) <> dt Then
                        ch = ch & ",date"
                    End If
                End If
                If Len(row("link")) > 0 And row("link") <> Nz(rs!doc_link, "") Then ch = ch & ",link"
                If Len(row("marking")) > 0 And RefMarkingText(row("marking")) <> Nz(rs!ref_marking, "UNCLASSIFIED") Then ch = ch & ",marking"
                If Len(row("notes")) > 0 And row("notes") <> Nz(rs!ref_notes, "") Then ch = ch & ",notes"
                If Len(row("status")) > 0 And row("status") <> row("curStatus") And Len(row("sup")) = 0 Then ch = ch & ",status"
                rs.Close
                row("changes") = Mid$(ch, 2)
                row("result") = IIf(Len(ch) > 0, "update", "unchanged")
            End If
        End If
    Next
    ' "Superseded by" names a publication by short title: in the library, or added by this file.
    For Each row In rows
        If row("result") <> "error" And Len(row("sup")) > 0 Then
            k = RefKey(row("sup"))
            libC = 0: fileC = 0: fileId = 0: fileFree = False
            Set rs = CurrentDb.OpenRecordset("SELECT ref_id, ref_number, ref_status, superseded_by FROM " & REF_TABLE & _
                " ORDER BY IIf(ref_status='" & REF_ACTIVE & "',0,1), ref_id DESC", dbOpenSnapshot)
            Do While Not rs.EOF
                If StrComp(RefKey(Nz(rs!ref_number, "")), k, vbBinaryCompare) = 0 And rs!ref_id <> row("id") Then
                    If libC = 0 Then libC = rs!ref_id
                End If
                rs.MoveNext
            Loop
            rs.Close
            ' A row of this file with that short title is the one it means (RefImportApply does the same).
            For Each other In rows
                If Not other Is row And other("result") <> "error" Then
                    If StrComp(RefKey(other("number")), k, vbBinaryCompare) = 0 Then
                        fileC = fileC + 1
                        If fileC = 1 Then
                            fileId = other("id")
                            ' Left as it is by this file (not superseded or restored by it).
                            fileFree = (Len(other("sup")) = 0 And (Len(other("status")) = 0 Or other("status") = REF_SUPERSEDED))
                        End If
                    End If
                End If
            Next
            If fileC > 0 Then target = fileId Else target = libC      ' 0: a reference this file adds
            If libC = 0 And fileC = 0 Then
                row("result") = "error"
                If StrComp(RefKey(row("number")), k, vbBinaryCompare) = 0 Then
                    row("note") = "A reference can't supersede itself."
                Else
                    row("note") = "Superseded by " & row("sup") & ": no reference has that short title."
                End If
            Else
                ' Two references can't supersede each other (as the change page refuses).
                For Each other In rows
                    If Not other Is row And other("result") <> "error" And Len(other("sup")) > 0 And Len(RefKey(row("number"))) > 0 Then
                        If StrComp(RefKey(other("sup")), RefKey(row("number")), vbBinaryCompare) = 0 And _
                           StrComp(RefKey(other("number")), k, vbBinaryCompare) = 0 Then
                            row("result") = "error"
                            row("note") = "Superseded by " & row("sup") & ", which this file marks as superseded by this one; two references can't supersede each other."
                        End If
                    End If
                Next
                If row("result") <> "error" And row("id") > 0 And target > 0 And (fileC = 0 Or fileFree) Then
                    If Nz(DLookup("superseded_by", REF_TABLE, "ref_id=" & target), 0) = row("id") Then
                        row("result") = "error"
                        row("note") = row("sup") & " is marked as superseded by this one; two references can't supersede each other."
                    End If
                End If
                If row("result") <> "error" And row("id") > 0 Then
                    ' Already superseded by that one: nothing to change.
                    If Not (target > 0 And row("curStatus") = REF_SUPERSEDED And row("curSup") = target) Then
                        row("changes") = row("changes") & IIf(Len(row("changes")) > 0, ",", "") & "superseded by"
                        row("result") = "update"
                    End If
                End If
            End If
        End If
    Next
    Set RefImportCheck = rows
End Function

' How many checked rows have result ("add", "update", "unchanged" or "error").
Public Function RefImportCount(ByVal checked As Collection, ByVal result As String) As Long
    Dim row As Object
    For Each row In checked
        If row("result") = result Then RefImportCount = RefImportCount + 1
    Next
End Function

' What a check found, for a message: the counts, then the problem rows (up
' to 12) or the changes.
Public Function RefImportReport(ByVal checked As Collection) As String
    Dim row As Object, s As String, n As Long
    s = RefImportCount(checked, "add") & " new, " & RefImportCount(checked, "update") & " to change, " & _
        RefImportCount(checked, "unchanged") & " already there"
    If RefImportCount(checked, "error") > 0 Then s = s & ", " & RefImportCount(checked, "error") & " with a problem"
    s = s & "."
    For Each row In checked
        If n >= 12 Then s = s & vbCrLf & "...": Exit For
        If row("result") = "error" Then
            s = s & vbCrLf & "Row " & row("line") & ": " & row("note"): n = n + 1
        ElseIf row("result") = "update" And RefImportCount(checked, "error") = 0 Then
            s = s & vbCrLf & "Row " & row("line") & ": " & IIf(Len(row("number")) > 0, row("number"), row("title")) & " (" & row("changes") & ")": n = n + 1
        End If
    Next
    RefImportReport = s
End Function

' Write checked rows (RefImportCheck) in one transaction: the new ones, the
' changes, then what supersedes what. Refused (error 5) when a row has a
' problem, so a file goes in whole or not at all. Returns "Added N, changed N."
' Library editors and Personnel admins only.
Public Function RefImportApply(ByVal checked As Collection) As String
    Dim ws As DAO.Workspace, db As DAO.Database, rs As DAO.Recordset, row As Object, other As Object, inTrans As Boolean
    Dim who As String, added As Long, changed As Long, ch As String, k As String, toId As Long, st As String
    On Error GoTo Fail
    If Not RefIsEditor() Then Err.Raise 5, "modReferences", "Only Library editors and Personnel admins can import references."
    If RefImportCount(checked, "error") > 0 Then Err.Raise 5, "modReferences", _
        "Some rows have a problem, so nothing was imported. Fix them and import the file again."
    who = RefActorName()
    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb
    ws.BeginTrans: inTrans = True
    For Each row In checked
        If row("result") = "add" Then
            st = row("status")
            If st <> REF_CANCELLED Then st = REF_ACTIVE         ' superseded is set below
            Set rs = db.OpenRecordset(REF_TABLE, dbOpenDynaset, dbAppendOnly)
            rs.AddNew
            rs!ref_number = RefNull(Trim$(row("number"))): rs!ref_title = row("title"): rs!ref_revision = RefNull(Trim$(row("revision")))
            rs!pub_date = row("dateValue")
            rs!ref_status = st: rs!doc_link = RefNull(row("link")): rs!ref_marking = RefMarkingText(row("marking"))
            rs!ref_notes = RefNull(Trim$(row("notes"))): rs!source_app = "import": rs!row_version = 1
            rs!created_by_id = SignedInPersonID(): rs!created_by = who: rs!created_at = Now()
            rs!updated_by = who: rs!updated_at = Now()
            rs.Update
            rs.Bookmark = rs.LastModified
            row("id") = rs!ref_id
            rs.Close
            RefSetDetail db, row("id"), row("publisher"), Trim$(row("edition"))
            RefAudit db, row("id"), "Imported", "Import added " & RefLabelOf(row("number"), row("title"), row("revision"), row("edition")) & ".", "import"
            added = added + 1
        ElseIf row("result") = "update" Then
            ch = "," & row("changes") & ","
            Set rs = db.OpenRecordset("SELECT * FROM " & REF_TABLE & " WHERE ref_id=" & row("id"), dbOpenDynaset)
            rs.Edit
            If InStr(ch, ",title,") > 0 Then rs!ref_title = row("title")
            If InStr(ch, ",change,") > 0 Then rs!ref_revision = RefNull(Trim$(row("revision")))
            If InStr(ch, ",date,") > 0 Then rs!pub_date = row("dateValue")
            If InStr(ch, ",link,") > 0 Then rs!doc_link = RefNull(row("link"))
            If InStr(ch, ",marking,") > 0 Then rs!ref_marking = RefMarkingText(row("marking"))
            If InStr(ch, ",notes,") > 0 Then rs!ref_notes = RefNull(Trim$(row("notes")))
            If InStr(ch, ",status,") > 0 Then rs!ref_status = row("status"): rs!superseded_by = Null
            rs!row_version = rs!row_version + 1: rs!updated_by = who: rs!updated_at = Now()
            rs.Update
            rs.Close
            If InStr(ch, ",publisher,") > 0 Or InStr(ch, ",edition,") > 0 Then
                ' Blank cells keep what the library has.
                RefSetDetail db, row("id"), IIf(Len(row("publisher")) > 0, row("publisher"), RefPublisher(row("id"))), _
                    IIf(Len(row("edition")) > 0, Trim$(row("edition")), RefEdition(row("id")))
            End If
            RefAudit db, row("id"), "Imported", "Import changed " & Replace(row("changes"), ",", ", ") & " of " & _
                RefLabelOf(row("number"), row("title"), row("revision"), row("edition")) & ".", "import"
            changed = changed + 1
        End If
    Next
    ' What supersedes what: a row of this file first, else the library's (an active one first).
    For Each row In checked
        If Len(row("sup")) > 0 And (row("result") = "add" Or InStr("," & row("changes") & ",", ",superseded by,") > 0) Then
            k = RefKey(row("sup")): toId = 0
            For Each other In checked
                If Not other Is row And other("id") > 0 Then
                    If StrComp(RefKey(other("number")), k, vbBinaryCompare) = 0 Then toId = other("id"): Exit For
                End If
            Next
            If toId = 0 Then
                Set rs = db.OpenRecordset("SELECT ref_id, ref_number FROM " & REF_TABLE & " WHERE ref_id<>" & row("id") & _
                    " ORDER BY IIf(ref_status='" & REF_ACTIVE & "',0,1), ref_id DESC", dbOpenSnapshot)
                Do While Not rs.EOF
                    If StrComp(RefKey(Nz(rs!ref_number, "")), k, vbBinaryCompare) = 0 Then toId = rs!ref_id: Exit Do
                    rs.MoveNext
                Loop
                rs.Close
            End If
            db.Execute "UPDATE " & REF_TABLE & " SET ref_status='" & REF_SUPERSEDED & "', superseded_by=" & toId & _
                IIf(row("result") = "add", "", ", row_version=row_version+1") & " WHERE ref_id=" & row("id"), dbFailOnError
        End If
    Next
    RefAudit db, 0, "Import", "Import: " & added & " added, " & changed & " changed, " & RefImportCount(checked, "unchanged") & " already there.", "import"
    ws.CommitTrans: inTrans = False
    RefImportApply = "Added " & added & ", changed " & changed & "."
    Exit Function
Fail:
    If inTrans Then RefRollback ws
    RaiseAgain "modReferences.RefImportApply"
End Function
