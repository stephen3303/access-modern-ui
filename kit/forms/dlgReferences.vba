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
