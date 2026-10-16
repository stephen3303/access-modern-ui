Option Compare Database
Option Explicit

' First-run Personnel setup (BuildPersonnelDialogs builds it; KitPersonnelStart
' opens it with RunDialog). Shows one panel at a time (controls tagged
' choice, create or ask) and fits the dialog to it. Ends "ok" when the
' Personnel file is linked: found, or created with the current Windows user
' signed in as its first admin. Closing any other way leaves the app to quit.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Dim lost As Boolean
    lost = (PersonnelUiSetupMode() = PERSONNEL_LOST)
    If lost Then
        Me.lblTitle.Caption = "The Personnel file can't be reached"
        Me.lblWhy.Caption = APP_NAME & " signs people in from the shared Personnel file, which isn't at " & _
            PersonnelPath() & " any more. If it has moved, find it. Otherwise ask your admin."
    Else
        Me.lblWhy.Caption = APP_NAME & " signs people in from a shared Personnel file, and it can't find one. " & _
            "If Personnel is already set up for another app, find its file."
    End If
    ShowPanel "choice"
    If lost Then
        ' Never offer a second directory while one is linked.
        SetVisible Me.btnCreate, False
        SetVisible Me.lblCreateHint, False
        Me.btnNotAdmin.Top = Me.btnCreate.Top
        Me.lblNotAdminHint.Top = Me.lblCreateHint.Top
        FitToPanel "choice"
    End If
    Me.txtFolder = Left$(PersonnelDefaultPath(PersonnelUiAppTable()), InStrRev(PersonnelDefaultPath(PersonnelUiAppTable()), "\") - 1)
    Me.txtAsk = PersonnelAdminNote()
    CenterOnApp Me
    Exit Sub
Err_Handler:
    LogError "dlgPersonnelSetup.Form_Load"
End Sub

' Show the controls tagged panel (and the untagged title), hide the rest,
' and fit the dialog to them.
Private Sub ShowPanel(ByVal panel As String)
    Dim c As Control
    For Each c In Me.Controls
        If Len(c.Tag) > 0 Then
            If c.Tag = panel Then SetVisible c, True
        End If
    Next
    For Each c In Me.Controls
        If Len(c.Tag) > 0 And c.Tag <> panel Then SetVisible c, False
    Next
    FitToPanel panel
End Sub

' Height to the lowest visible control of panel, plus the dialog padding.
Private Sub FitToPanel(ByVal panel As String)
    Dim c As Control, lowest As Long
    For Each c In Me.Controls
        If c.Tag = panel And c.Visible Then
            If c.Top + c.Height > lowest Then lowest = c.Top + c.Height
        End If
    Next
    Me.Section(acDetail).Height = lowest + Me.lblTitle.Top
    Me.InsideHeight = Me.Section(acDetail).Height
    CenterOnApp Me
End Sub

Private Sub btnFind_Click()
    On Error GoTo Err_Handler
    Dim path As String, problem As String
    Do
        path = KitPickFile("Find the Personnel file", "Personnel file|" & PERSONNEL_FILE & "|Access database|*.accdb", "personnel")
        If Len(path) = 0 Then Exit Sub
        problem = PersonnelFileProblem(path)
        If Len(problem) = 0 Then Exit Do
        KitMsgBox problem & vbCrLf & vbCrLf & "Pick the Personnel file again, or Cancel.", vbExclamation
    Loop
    LinkPersonnelFile path
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnelSetup.btnFind_Click", path
End Sub

Private Sub btnCreate_Click()
    On Error GoTo Err_Handler
    ShowPanel "create"
    Me.lblTitle.Caption = "Create the Personnel file"
    Me.txtFirst.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnelSetup.btnCreate_Click"
End Sub

Private Sub btnNotAdmin_Click()
    On Error GoTo Err_Handler
    ShowPanel "ask"
    Me.lblTitle.Caption = "Ask your Personnel admin"
    Me.btnCopy.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnelSetup.btnNotAdmin_Click"
End Sub

Private Sub btnBack_Click()
    On Error GoTo Err_Handler
    ShowPanel "choice"
    Me.lblTitle.Caption = "Set up the Personnel directory"
    Me.btnCreate.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnelSetup.btnBack_Click"
End Sub

Private Sub btnFolder_Click()
    On Error GoTo Err_Handler
    Dim folder As String
    folder = KitPickFolder("Folder for the Personnel file", "personnel")
    If Len(folder) > 0 Then Me.txtFolder = folder
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnelSetup.btnFolder_Click"
End Sub

' Create the file, make the current Windows user its first admin with the
' password typed, and sign them in.
Private Sub btnCreateGo_Click()
    On Error GoTo Err_Handler
    Dim firstName As String, lastName As String, pwd As String, path As String, id As Long, problem As String
    Me.btnCreateGo.SetFocus                     ' commits the box being typed in
    firstName = Trim$(Nz(Me.txtFirst, "")): lastName = Trim$(Nz(Me.txtLast, "")): pwd = Nz(Me.txtPwd, "")
    If Len(firstName) = 0 Or Len(lastName) = 0 Then KitMsgBox "Type your first and last name.", vbExclamation: Exit Sub
    If StrComp(pwd, Nz(Me.txtPwd2, ""), vbBinaryCompare) <> 0 Then KitMsgBox "The two passwords don't match.", vbExclamation: Exit Sub
    If Len(pwd) < 8 Then KitMsgBox "Use at least 8 characters for your password.", vbExclamation: Exit Sub
    path = Nz(Me.txtFolder, "")
    If Right$(path, 1) <> "\" Then path = path & "\"
    path = path & PERSONNEL_FILE
    If FileExists(path) Then
        KitMsgBox "There is already a Personnel file in that folder. Go Back and choose Find the Personnel file.", vbExclamation
        Exit Sub
    End If
    DoCmd.Hourglass True
    id = CreatePersonnelFile(path, PersonnelUiAppCode(), firstName, lastName, Trim$(Nz(Me.txtRank, "")))
    problem = PersonnelSetPassword(id, pwd)
    DoCmd.Hourglass False
    If Len(problem) > 0 Then KitMsgBox problem, vbExclamation
    PersonnelSignInWindows PersonnelUiAppCode()
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    DoCmd.Hourglass False
    HandleError "dlgPersonnelSetup.btnCreateGo_Click", path, "The Personnel file couldn't be created."
End Sub

Private Sub btnCopy_Click()
    On Error GoTo Err_Handler
    If KitCopyText(Nz(Me.txtAsk, "")) Then KitToast "Copied", "success"
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnelSetup.btnCopy_Click"
End Sub

Private Sub btnCloseApp_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnelSetup.btnCloseApp_Click"
End Sub
