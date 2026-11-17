Option Compare Database
Option Explicit

' The Reference library page (BuildReferenceDialogs builds it;
' KitReferencesManage opens it). Lists the shared library's references; the
' search box narrows the list as you type and Show retired adds the
' superseded and cancelled ones. Add reference (anyone signed in) and Change
' (Library editors and Personnel admins; a double-click too) open
' dlgReference; Open shows the document; Library editors (Personnel admins)
' opens dlgRefEditors.

Private m_editor As Boolean

Private Sub Form_Load()
    On Error GoTo Err_Handler
    m_editor = RefIsEditor()
    SetVisible Me.btnEdit, m_editor
    SetVisible Me.btnImport, m_editor
    SetVisible Me.btnEditors, PersonnelIsAdminNow()
    If m_editor Then
        Me.lblTip.Caption = "Double-click a reference to change, retire or supersede it. Open shows its document."
    Else
        Me.lblTip.Caption = "Only Library editors change or retire a reference. Open shows its document."
    End If
    LoadRefs
    CenterOnApp Me
    Exit Sub
Err_Handler:
    HandleError "dlgReferences.Form_Load"
End Sub

' F1 opens the help on this page (KitHelp).
Public Property Get HelpTopic() As String
    HelpTopic = "reflib"
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    RefF1 KeyCode
End Sub

' Fill the list with the references matching the search, keeping the
' selected one when it is still in it.
Private Sub LoadRefs(Optional ByVal term As Variant)
    Dim keep As Variant, i As Long
    If IsMissing(term) Then term = Nz(Me.txtSearch, "")
    keep = Me.lstRefs.Value
    Me.lstRefs.RowSource = RefLibrarySql(CStr(term), Nz(Me.chkRetired, False))
    ' With only the headings row left, the list is hidden and the note says so.
    SetVisible Me.lstRefs, Me.lstRefs.ListCount > 1
    SetVisible Me.lblNone, Me.lstRefs.ListCount <= 1
    Me.lstRefs.Value = Null
    If IsNull(keep) Then Exit Sub
    For i = 1 To Me.lstRefs.ListCount - 1           ' row 0 is the headings
        If CStr(Me.lstRefs.ItemData(i)) = CStr(keep) Then Me.lstRefs.Value = keep: Exit For
    Next i
End Sub

Private Sub txtSearch_Change()
    On Error GoTo Err_Handler
    LoadRefs Nz(Me.txtSearch.Text, "")
    Exit Sub
Err_Handler:
    HandleError "dlgReferences.txtSearch_Change"
End Sub

Private Sub chkRetired_AfterUpdate()
    On Error GoTo Err_Handler
    LoadRefs
    Exit Sub
Err_Handler:
    HandleError "dlgReferences.chkRetired_AfterUpdate"
End Sub

Private Sub btnAdd_Click()
    On Error GoTo Err_Handler
    If RunDialog(REF_FORM, "new") Then
        LoadRefs
        Me.lstRefs.Value = Nz(TempVars("RefId"), Null)
    End If
    Exit Sub
Err_Handler:
    HandleError "dlgReferences.btnAdd_Click"
End Sub

Private Sub lstRefs_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler
    If m_editor Then EditSelected Else OpenSelected
    Exit Sub
Err_Handler:
    HandleError "dlgReferences.lstRefs_DblClick"
End Sub

Private Sub btnEdit_Click()
    On Error GoTo Err_Handler
    EditSelected
    Exit Sub
Err_Handler:
    HandleError "dlgReferences.btnEdit_Click"
End Sub

Private Sub btnOpen_Click()
    On Error GoTo Err_Handler
    OpenSelected
    Exit Sub
Err_Handler:
    HandleError "dlgReferences.btnOpen_Click"
End Sub

Private Sub EditSelected()
    If IsNull(Me.lstRefs.Value) Then KitMsgBox "Pick a reference in the list first.", vbInformation: Exit Sub
    If RunDialog(REF_FORM, CStr(Me.lstRefs.Value)) Then LoadRefs
End Sub

Private Sub OpenSelected()
    If IsNull(Me.lstRefs.Value) Then KitMsgBox "Pick a reference in the list first.", vbInformation: Exit Sub
    KitRefOpen Me.lstRefs.Value
End Sub

' The whole library (retired references too) in Stephen's seven columns,
' then Document, Notes, Status and Superseded by: an Excel workbook or CSV.
Private Sub btnExport_Click()
    On Error GoTo Err_Handler
    Dim path As String, n As Long
    path = KitSaveAs("Export the Reference library", "Reference library " & Format$(Date, "yyyy-mm-dd") & ".xlsx", _
        "Excel workbook|*.xlsx|CSV file|*.csv", "references-export")
    If Len(path) = 0 Then Exit Sub
    n = RefExport(path)
    KitToast "Exported " & n & IIf(n = 1, " reference", " references") & " to " & Mid$(path, InStrRev(path, "\") + 1), "success"
    Exit Sub
Err_Handler:
    HandleError "dlgReferences.btnExport_Click"
End Sub

' A sheet in the export's layout (Library editors): checked first, then
' written whole after a yes, or not at all when a row has a problem.
Private Sub btnImport_Click()
    On Error GoTo Err_Handler
    Dim path As String, rows As Collection, report As String, done As String
    path = KitPickFile("Import references (the first row names the columns, as the export writes them)", _
        "Spreadsheets|*.xlsx;*.xls;*.csv", "references-import")
    If Len(path) = 0 Then Exit Sub
    KitBusy "Checking the file..."
    Set rows = RefImportCheck(RefImportRead(path))
    KitBusyEnd
    If rows.Count = 0 Then KitMsgBox "The sheet has a header row but no references under it.", vbInformation, "Import references": Exit Sub
    report = RefImportReport(rows)
    If RefImportCount(rows, "error") > 0 Then
        KitMsgBox report & vbCrLf & vbCrLf & "Nothing was imported. Fix these rows and import the file again.", vbExclamation, "Import references"
        Exit Sub
    End If
    If RefImportCount(rows, "add") + RefImportCount(rows, "update") = 0 Then
        KitMsgBox "Every row is already in the library as it is. Nothing to import.", vbInformation, "Import references"
        Exit Sub
    End If
    If KitMsgBox(report & vbCrLf & vbCrLf & "Import these now?", vbQuestion + vbYesNo, "Import references") <> vbYes Then Exit Sub
    done = RefImportApply(rows)
    KitToast "Imported: " & done, "success"
    LoadRefs
    Exit Sub
Err_Handler:
    KitBusyEnd
    If Err.Number = 5 Then KitMsgBox Err.Description, vbExclamation, "Import references": Exit Sub
    HandleError "dlgReferences.btnImport_Click"
End Sub

Private Sub btnEditors_Click()
    On Error GoTo Err_Handler
    RunDialog REF_EDITORS_FORM
    Exit Sub
Err_Handler:
    HandleError "dlgReferences.btnEditors_Click"
End Sub

Private Sub btnClose_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgReferences.btnClose_Click"
End Sub
