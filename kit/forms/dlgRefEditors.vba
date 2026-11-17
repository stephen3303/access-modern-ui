Option Compare Database
Option Explicit

' Library editors (BuildReferenceDialogs builds it; the library page's
' Library editors button opens it): who may change, retire and supersede
' references in every app. Personnel admins only; they always may.

' Checked against the Personnel file now (PersonnelIsAdminNow). The builder
' wires OnOpen to this in place of =ThemeForm([Form]), so it themes the form too.
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler
    ThemeForm Me
    If Not PersonnelIsAdminNow() Then
        KitMsgBox "Only Personnel admins choose the Library editors.", vbExclamation
        Cancel = True
    End If
    Exit Sub
Err_Handler:
    Cancel = True
    HandleError "dlgRefEditors.Form_Open"
End Sub

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.lblError.Caption = ""
    LoadEditors
    CenterOnApp Me
    Exit Sub
Err_Handler:
    HandleError "dlgRefEditors.Form_Load"
End Sub

Public Property Get HelpTopic() As String
    HelpTopic = "reflib.editors"
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    RefF1 KeyCode
End Sub

Private Sub txtPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    KitPickKey Me.txtPerson, KeyCode
End Sub

Private Sub LoadEditors()
    Me.lstEditors.RowSource = RefEditorsSql()
    SetVisible Me.lstEditors, Me.lstEditors.ListCount > 1
    SetVisible Me.lblNone, Me.lstEditors.ListCount <= 1
End Sub

Private Sub btnAdd_Click()
    On Error GoTo Err_Handler
    Me.lblError.Caption = ""
    If IsNull(Me.txtPersonKey) Then Me.lblError.Caption = "Choose a person first.": Exit Sub
    RefSetEditor CLng(Me.txtPersonKey), True
    Me.txtPersonKey = Null: Me.txtPerson = Null
    LoadEditors
    Exit Sub
Err_Handler:
    If Err.Number = 5 Then Me.lblError.Caption = Err.Description: Exit Sub
    HandleError "dlgRefEditors.btnAdd_Click"
End Sub

Private Sub btnRemove_Click()
    On Error GoTo Err_Handler
    Me.lblError.Caption = ""
    If IsNull(Me.lstEditors.Value) Then Me.lblError.Caption = "Pick an editor in the list first.": Exit Sub
    RefSetEditor CLng(Me.lstEditors.Value), False
    LoadEditors
    Exit Sub
Err_Handler:
    If Err.Number = 5 Then Me.lblError.Caption = Err.Description: Exit Sub
    HandleError "dlgRefEditors.btnRemove_Click"
End Sub

Private Sub btnClose_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgRefEditors.btnClose_Click"
End Sub
