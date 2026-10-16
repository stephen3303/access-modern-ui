Option Compare Database
Option Explicit

' First-run Personnel setup (BuildPersonnelDialogs builds it; KitPersonnelStart
' opens it with RunDialog). Shows one panel at a time (controls tagged
' choice, create or ask) and fits the dialog to it. Ends "ok" when the
' Personnel file is linked: found, or created with the current Windows user
' signed in as its first admin. Closing any other way leaves the app to quit.

Private m_tops As Object          ' each tagged control's designed Top, by name
Private m_fullHeight As Long      ' the designed Detail height (the tallest panel)

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Dim lost As Boolean, c As Control
    Set m_tops = CreateObject("Scripting.Dictionary")
    For Each c In Me.Controls
        If Len(c.Tag) > 0 Then m_tops(c.Name) = c.Top
    Next
    m_fullHeight = Me.Section(acDetail).Height
    lost = (PersonnelUiSetupMode() = PERSONNEL_LOST)
    If lost Then
        Me.lblTitle.Caption = "The Personnel file can't be reached"
        Me.lblWhy.Caption = APP_NAME & " can't reach its Personnel file at " & PersonnelPath() & _
            ". If it has moved, find it. Otherwise ask your admin."
        ' Never offer a second directory while one is linked: Create leaves
        ' every panel, and I'm not the admin takes its place.
        Me.btnCreate.Tag = "never": Me.lblCreateHint.Tag = "never"
        m_tops("btnNotAdmin") = m_tops("btnCreate")
        m_tops("lblNotAdminHint") = m_tops("lblCreateHint")
    Else
        Me.lblWhy.Caption = APP_NAME & " signs people in from a shared Personnel file, and it can't find one. " & _
            "If Personnel is already set up for another app, find its file."
    End If
    Me.txtFolder = Left$(PersonnelDefaultPath(PersonnelUiAppTable()), InStrRev(PersonnelDefaultPath(PersonnelUiAppTable()), "\") - 1)
    Me.txtAsk = PersonnelAdminNote()
    FitWhy
    ShowPanel "choice"
    Exit Sub
Err_Handler:
    LogError "dlgPersonnelSetup.Form_Load"
End Sub

' F1 opens the help on this screen (KitHelp).
Public Property Get HelpTopic() As String
    HelpTopic = "directory.setup"
End Property

' Show the controls tagged panel (and the untagged title) at their designed
' places, hide the rest, and fit the dialog to the panel. Access won't make
' the Detail shorter than its lowest control, hidden or not, so hidden
' controls move to the top first; the Detail grows back to its designed
' height before the shown ones return to their places.
Private Sub ShowPanel(ByVal panel As String)
    Dim c As Control
    Me.Section(acDetail).Height = m_fullHeight
    For Each c In Me.Controls
        If Len(c.Tag) > 0 Then
            If c.Tag <> panel Then
                SetVisible c, False
                c.Top = 0
            End If
        End If
    Next
    For Each c In Me.Controls
        If c.Tag = panel Then
            c.Top = m_tops(c.Name)
            SetVisible c, True
        End If
    Next
    FitToPanel panel
    Me.Repaint                                  ' no trace of the last panel's text
End Sub

' Make lblWhy as tall as its text (two lines usually, four or more with a
' long path) and move the choices under it to match. The lines are counted
' by wrapping the words as the label does, since a long path can't break at
' a space and leaves the line before it short.
Private Sub FitWhy()
    Dim lines As Long, dy As Long, k As Variant
    lines = WrapLines(Me.lblWhy.Caption, Me.lblWhy.Width - 60)    ' less the label's inner margins
    dy = lines * CLng(10 * 20 * 1.36) + 30 - Me.lblWhy.Height
    If dy = 0 Then Exit Sub
    m_fullHeight = m_fullHeight + IIf(dy > 0, dy, 0)
    Me.Section(acDetail).Height = m_fullHeight
    Me.lblWhy.Height = Me.lblWhy.Height + dy
    For Each k In m_tops.Keys
        If Me.Controls(k).Tag = "choice" And m_tops(k) > m_tops("lblWhy") Then m_tops(k) = m_tops(k) + dy
    Next
End Sub

' Lines text takes at 10pt in a label avail twips wide, wrapping word by
' word. A word wider than the line (a long path) breaks across lines.
Private Function WrapLines(ByVal text As String, ByVal avail As Long) As Long
    Dim para As Variant, wd As Variant, n As Long, used As Long, ww As Long, gap As Long
    gap = MeasureText("a b", 10, 400) - MeasureText("ab", 10, 400)
    For Each para In Split(Replace(text, vbCrLf, vbLf), vbLf)
        n = n + 1: used = 0
        For Each wd In Split(CStr(para), " ")
            If Len(wd) > 0 Then
                ww = MeasureText(CStr(wd), 10, 400)
                If used > 0 And used + gap + ww > avail Then n = n + 1: used = 0
                If used > 0 Then used = used + gap
                used = used + ww
                Do While used > avail
                    n = n + 1: used = used - avail
                Loop
            End If
        Next
    Next
    If n < 1 Then n = 1
    WrapLines = n
End Function

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
    PersonnelUiCreated                          ' KitPersonnelStart opens Add person next
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    DoCmd.Hourglass False
    HandleError "dlgPersonnelSetup.btnCreateGo_Click", path, "The Personnel file couldn't be created."
End Sub

Private Sub btnCopy_Click()
    On Error GoTo Err_Handler
    If KitCopyText(Nz(Me.txtAsk, "")) Then KitToast "Copied", "success" Else KitToast "Couldn't copy. Try again.", "warning"
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
