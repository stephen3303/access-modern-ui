Option Compare Database
Option Explicit

' Add or change one reference in the shared library (BuildReferenceDialogs
' builds it; dlgReferences and KitRefAdd open it). OpenArgs: "new",
' "new|<title typed in the citation>" or a ref_id. Anyone signed in adds;
' only Library editors and Personnel admins change (status and Superseded
' by too). Adding a document that is already there picks that one instead.
' Ends "ok" with TempVars("RefId") set to the reference saved or found.

Private m_id As Long            ' 0 while adding
Private m_ver As Long           ' row_version when it was opened (the stale-save check)

' Changing is for Library editors and Personnel admins only, checked against
' the files now (RefIsEditor); adding needs someone signed in. RunDialog
' treats the cancelled open as not ok. The builder wires OnOpen to this, in
' place of the usual =ThemeForm([Form]) hook, so it themes the form too.
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler
    ThemeForm Me
    If Left$(Nz(Me.OpenArgs, "new"), 3) = "new" Then
        If SignedInPersonID() = 0 Then KitMsgBox "Sign in first to add a reference.", vbExclamation: Cancel = True
    ElseIf Not RefIsEditor() Then
        KitMsgBox "Only Library editors and Personnel admins can change a reference. Ask one to make the change.", vbExclamation
        Cancel = True
    End If
    Exit Sub
Err_Handler:
    Cancel = True
    HandleError "dlgReference.Form_Open"
End Sub

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Dim rs As DAO.Recordset, a As String
    a = Nz(Me.OpenArgs, "new")
    Me.lblError.Caption = ""
    If Left$(a, 3) = "new" Then
        m_id = 0
        Me.lblTitle.Caption = "Add reference"
        Me.cboMarking = "UNCLASSIFIED"
        If InStr(a, "|") > 0 Then Me.txtTitle = Mid$(a, InStr(a, "|") + 1)
        ShowChangeRows False
    Else
        m_id = CLng(a)
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & REF_TABLE & " WHERE ref_id=" & m_id, dbOpenSnapshot)
        If rs.EOF Then rs.Close: KitMsgBox "That reference is no longer in the library.", vbExclamation: DoCmd.Close acForm, Me.Name, acSaveNo: Exit Sub
        Me.lblTitle.Caption = "Change " & RefLabelOf(Nz(rs!ref_number, ""), rs!ref_title, Nz(rs!ref_revision, ""))
        Me.txtNumber = rs!ref_number: Me.txtRevision = rs!ref_revision: Me.txtTitle = rs!ref_title
        Me.txtDate = rs!pub_date: Me.cboMarking = Nz(rs!ref_marking, "UNCLASSIFIED")
        Me.txtLink = rs!doc_link: Me.txtNotes = rs!ref_notes
        Me.cboStatus = Nz(rs!ref_status, REF_ACTIVE)
        Me.txtSupKey = rs!superseded_by
        m_ver = rs!row_version
        rs.Close
        KitPickShow Me.txtSup
        ShowChangeRows True
    End If
    CenterOnApp Me
    Exit Sub
Err_Handler:
    HandleError "dlgReference.Form_Load"
End Sub

' The status row, its hint and History show only when changing; adding
' closes the gap they leave so Save sits under the notes.
Private Sub ShowChangeRows(ByVal show As Boolean)
    Dim c As Control, top As Long, delta As Long
    If show Then
        SetSupEnabled
        Exit Sub
    End If
    top = Me.lblStatus.Top
    delta = Me.lblError.Top - top
    For Each c In Me.Controls
        If IsChangeRow(c) Then c.Visible = False
    Next
    For Each c In Me.Controls
        If Not IsChangeRow(c) And c.Top > top Then c.Top = c.Top - delta
    Next
    Me.Section(acDetail).Height = Me.Section(acDetail).Height - delta
    Me.InsideHeight = Me.InsideHeight - delta
End Sub

' True for a control of the status row, History included. The builder tags
' them "change"; the Superseded by picker can't carry that Tag (PickBox keeps
' its row source there), so it is named.
Private Function IsChangeRow(c As Control) As Boolean
    IsChangeRow = (c.Tag = "change" Or c.Name = "txtSup" Or c.Name = "txtSupKey")
End Function

' Superseded by applies only to a superseded reference.
Private Sub SetSupEnabled()
    SetEnabled Me.txtSup, Nz(Me.cboStatus, "") = REF_SUPERSEDED
End Sub

Private Sub cboStatus_AfterUpdate()
    On Error GoTo Err_Handler
    SetSupEnabled
    Exit Sub
Err_Handler:
    HandleError "dlgReference.cboStatus_AfterUpdate"
End Sub

Private Sub txtSup_KeyDown(KeyCode As Integer, Shift As Integer)
    KitPickKey Me.txtSup, KeyCode
End Sub

' F1 opens the help on adding and changing references (KitHelp).
Public Property Get HelpTopic() As String
    HelpTopic = "reflib.reference"
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    RefF1 KeyCode
End Sub

' The Windows file picker for the document link, starting in its folder.
Private Sub btnBrowse_Click()
    On Error GoTo Err_Handler
    Dim path As String
    path = KitPickFile("Choose the reference document (on the shared drive, so everyone can open it)", _
        "Documents|*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.txt;*.rtf;*.xps;*.oxps;*.png;*.jpg;*.jpeg;*.tif;*.tiff|All files|*.*", _
        "references")
    If Len(path) = 0 Then Exit Sub
    Me.txtLink = path
    Me.lblError.Caption = RefPickNote(path)
    Exit Sub
Err_Handler:
    HandleError "dlgReference.btnBrowse_Click"
End Sub

' What to know about a picked file (the web's Get-KitReferencePickNote):
' a type the library won't open, or a file only this PC can reach.
Private Function RefPickNote(ByVal path As String) As String
    Dim ext As String
    If InStrRev(path, ".") > InStrRev(path, "\") Then ext = LCase$(Mid$(path, InStrRev(path, ".") + 1))
    If Len(ext) = 0 Or InStr(REF_OPEN_TYPES, "." & ext & ".") = 0 Then
        RefPickNote = IIf(Len(ext) > 0, "." & ext & " files", "Files without a type") & _
            " are not opened from the library, so people will only be able to copy the path."
    ElseIf path Like "[A-Za-z]:\*" Then
        On Error Resume Next
        If CreateObject("Scripting.FileSystemObject").GetDrive(Left$(path, 2)).DriveType <> 3 Then  ' 3: network
            RefPickNote = "That file is on this PC's own " & Left$(path, 2) & " drive, so nobody else can open it. Choose the copy on the shared drive."
        End If
    End If
End Function

' Open the saved document (a change); while adding there is nothing saved to open.
Private Sub btnOpenDoc_Click()
    On Error GoTo Err_Handler
    If m_id = 0 Then KitMsgBox "Save the reference first, then open its document from the library.", vbInformation: Exit Sub
    KitRefOpen m_id
    Exit Sub
Err_Handler:
    HandleError "dlgReference.btnOpenDoc_Click"
End Sub

' What happened to this reference, newest first.
Private Sub btnHistory_Click()
    On Error GoTo Err_Handler
    Dim rs As DAO.Recordset, s As String, n As Long
    Set rs = CurrentDb.OpenRecordset(RefHistorySql(m_id), dbOpenSnapshot)
    Do While Not rs.EOF And n < 12
        s = s & rs![When] & "  " & Nz(rs![Who], "") & "  " & Nz(rs![What], "") & vbCrLf
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    If Len(s) = 0 Then s = "No history yet."
    KitMsgBox s, vbInformation, "History"
    Exit Sub
Err_Handler:
    HandleError "dlgReference.btnHistory_Click"
End Sub

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    Dim id As Long, added As Boolean, sup As Long, lbl As String
    Me.btnSave.SetFocus                         ' commits the box being typed in
    Me.lblError.Caption = ""
    If m_id = 0 Then
        id = RefAdd(Nz(Me.txtNumber, ""), Nz(Me.txtTitle, ""), Nz(Me.txtRevision, ""), Me.txtDate.Value, Nz(Me.txtLink, ""), _
            Nz(Me.cboMarking, ""), Nz(Me.txtNotes, ""), RefUiAppCode(), added)
        TempVars("RefId") = id
        TempVars("DialogResult") = "ok"
        lbl = RefLabel(id)
        If added Then
            KitToast "Added " & lbl, "success"
        Else
            KitMsgBox lbl & " is already in the library, so that one is used.", vbInformation, "Add reference"
        End If
    Else
        If Not IsNull(Me.txtSupKey) Then sup = CLng(Me.txtSupKey)
        RefUpdate m_id, Nz(Me.txtNumber, ""), Nz(Me.txtTitle, ""), Nz(Me.txtRevision, ""), Me.txtDate.Value, Nz(Me.txtLink, ""), _
            Nz(Me.cboMarking, ""), Nz(Me.txtNotes, ""), Nz(Me.cboStatus, REF_ACTIVE), sup, m_ver, RefUiAppCode()
        TempVars("RefId") = m_id
        TempVars("DialogResult") = "ok"
        KitToast "Saved " & RefLabel(m_id), "success"
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    If Err.Number = 5 Then Me.lblError.Caption = Err.Description: Exit Sub   ' a refusal: say why, keep what was typed
    HandleError "dlgReference.btnSave_Click", CStr(m_id), "The reference couldn't be saved."
End Sub

Private Sub btnClose_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgReference.btnClose_Click"
End Sub
