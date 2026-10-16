Option Compare Database
Option Explicit

' Access UI kit - several users on one back end: record locks and write
' conflicts handled the same way in every app, in the kit's look.
'
' Two kinds of trouble come up once a back end is shared:
'   - Locks. Another user (or another of this user's windows) is writing the
'     same record or table this instant: errors 3008, 3009, 3186-3188, 3197,
'     3202, 3211, 3212, 3218, 3260-3262. They clear in a moment, so the kit
'     retries a few times, waiting a little longer each time, before it
'     gives up with a plain message.
'   - Write conflicts. Another user saved the record after this user started
'     editing it (Access's "Write Conflict" window, form error 7787, or 7878
'     "The data has been changed"). Retrying can't help; the user decides.
'
' For code:
'   n = KitExecute("UPDATE ...")       ' an action query, retried on locks
'   KitSaveRecord Me                   ' save a bound form's record, retried
' and, around DAO edits of your own:
'       attempt = 0
'   TryAgain:
'       On Error GoTo Fail
'       rs.Edit: rs!status = "Closed": rs.Update
'       Exit Sub
'   Fail:
'       If LockWait(Err.Number, attempt) Then Resume TryAgain
'       HandleError "pgItems.CloseItem"
'
' For bound forms, add this to the form's code-behind:
'   Private Sub Form_Error(DataErr As Integer, Response As Integer)
'       KitFormError Me, DataErr, Response
'   End Sub
' It replaces Access's Write Conflict window with a KitMsgBox: Yes keeps the
' user's changes on top of the other user's saved record (they save again),
' No drops them and shows the other user's. A lock gets "wait a moment, then
' save again", with the user's changes kept. Other errors are left to Access.

Public Const LOCK_TRIES As Long = 5            ' retries after the first try
Public Const LOCK_WAIT_MS As Long = 120        ' first wait; each retry waits longer
Public Const ERR_WRITE_CONFLICT As Long = 7787 ' Access's Write Conflict window
Public Const ERR_DATA_CHANGED As Long = 7878   ' "The data has been changed"

Private Const LOCK_MSG As String = "Someone else is saving this record right now. Wait a moment, then save again."
Private Const LOCK_RAISE As String = "Another user is saving the same records right now. Wait a moment, then try again."

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private m_testAnswer As Long    ' tests only: the conflict answer, 0 to ask
Private m_saving As Boolean     ' KitSaveRecord is saving: KitFormError leaves locks to it
Private m_formLock As Long      ' the lock error KitFormError passed on to KitSaveRecord

' True for the errors a moment's wait can clear (see the top of the module).
Public Function IsLockError(ByVal errNum As Long) As Boolean
    Select Case errNum
        Case 3008, 3009, 3186, 3187, 3188, 3197, 3202, 3211, 3212, 3218, 3260, 3261, 3262
            IsLockError = True
    End Select
End Function

' In an error handler: when errNum is a lock error and fewer than LOCK_TRIES
' retries have been made, wait (longer each time, with a random part so two
' users don't retry in step), count the retry in attempt and return True, so
' the handler resumes and tries again. Otherwise return False.
Public Function LockWait(ByVal errNum As Long, ByRef attempt As Long) As Boolean
    If Not IsLockError(errNum) Then Exit Function
    If attempt >= LOCK_TRIES Then Exit Function
    attempt = attempt + 1
    Randomize
    Pause LOCK_WAIT_MS * attempt + CLng(Rnd() * LOCK_WAIT_MS)
    LockWait = True
End Function

' Run an action query on the current database with dbFailOnError, retrying
' while its records are locked. Returns the records it changed. When the
' locks don't clear, raises the lock error with a plain description for the
' caller's HandleError to show; any other error is raised unchanged.
Public Function KitExecute(ByVal sql As String) As Long
    Dim db As DAO.Database, attempt As Long, num As Long
    Set db = CurrentDb
TryAgain:
    On Error GoTo Fail
    db.Execute sql, dbFailOnError
    KitExecute = db.RecordsAffected
    Exit Function
Fail:
    If LockWait(Err.Number, attempt) Then Resume TryAgain
    If IsLockError(Err.Number) Then
        num = Err.Number
        LogError "modKitLocks.KitExecute", "locked after " & LOCK_TRIES & " retries: " & Left$(sql, 200)
        On Error GoTo 0
        Err.Raise num, "modKitLocks.KitExecute", LOCK_RAISE
    End If
    RaiseAgain "modKitLocks.KitExecute", Left$(sql, 200)
End Function

' Save frm's changed record, retrying while it is locked. Returns True when
' the record is saved or had no changes. Returns False when the form refused
' (its BeforeUpdate cancelled), the user chose in KitFormError, or the lock
' didn't clear (the user is told). Other errors are raised.
Public Function KitSaveRecord(frm As Form) As Boolean
    Dim attempt As Long, num As Long
TryAgain:
    On Error GoTo Fail
    m_saving = True
    m_formLock = 0
    If frm.Dirty Then frm.Dirty = False
    m_saving = False
    KitSaveRecord = True
    Exit Function
Fail:
    m_saving = False
    num = Err.Number
    ' A form wired to KitFormError gets the lock in Form_Error first; it
    ' notes it and stays quiet, and the save fails here as refused (2101 etc.).
    If m_formLock <> 0 Then num = m_formLock
    m_formLock = 0
    If LockWait(num, attempt) Then Resume TryAgain
    If IsLockError(num) Then
        LogError "modKitLocks.KitSaveRecord", frm.Name & " locked after " & LOCK_TRIES & " retries"
        KitMsgBox LOCK_MSG, vbExclamation
        Exit Function
    End If
    Select Case num
        Case 2101, 2115, 2169, 3020     ' refused by BeforeUpdate or Form_Error: it has said why
            Exit Function
    End Select
    RaiseAgain "modKitLocks.KitSaveRecord", frm.Name
End Function

' A bound form's Form_Error (see the top of the module): handles write
' conflicts and locks in the kit's look and sets Response so Access shows
' nothing; leaves every other error to Access.
Public Sub KitFormError(frm As Form, ByVal DataErr As Integer, Response As Integer)
    On Error GoTo Fail
    Select Case DataErr
        Case ERR_WRITE_CONFLICT, ERR_DATA_CHANGED
            Response = acDataErrContinue
            ResolveConflict frm
        Case Else
            If IsLockError(DataErr) Then
                Response = acDataErrContinue
                If m_saving Then
                    m_formLock = DataErr        ' KitSaveRecord retries, and tells the user if it can't
                Else
                    KitMsgBox LOCK_MSG, vbExclamation
                End If
            End If
    End Select
    Exit Sub
Fail:
    LogError "modKitLocks.KitFormError", frm.Name & " error " & DataErr
End Sub

' Tests only: answer the next conflicts with vbYes or vbNo instead of
' asking (0 asks again).
Public Sub SetConflictAnswer(ByVal answer As Long)
    m_testAnswer = answer
End Sub

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

' Another user saved frm's record while this user was editing it. Note the
' user's changes, drop them, show the saved record, then ask: Yes puts the
' user's changes back on top (the record is changed again, for them to
' save), No leaves the other user's version.
Private Sub ResolveConflict(frm As Form)
    Dim names As Collection, fields As Collection, vals As Collection, i As Long, list As String, answer As Long
    Set names = New Collection
    Set fields = New Collection
    Set vals = New Collection
    ChangedFields frm, names, fields, vals
    For i = 1 To fields.Count
        If i <= 5 Then list = list & IIf(i > 1, ", ", "") & fields(i)
    Next
    If fields.Count > 5 Then list = list & " and " & (fields.Count - 5) & " more"
    frm.Undo
    frm.Refresh                     ' the other user's saved values
    If names.Count = 0 Then Exit Sub
    answer = m_testAnswer
    If answer = 0 Then
        answer = KitMsgBox("Someone else saved this record while you were editing it." & vbCrLf & vbCrLf & _
            "You changed: " & list & "." & vbCrLf & vbCrLf & _
            "Keep your changes? Yes puts them back over the other person's, then you save again. " & _
            "No keeps the other person's version.", vbYesNo + vbExclamation, "Record changed by someone else")
    End If
    If answer <> vbYes Then Exit Sub
    For i = 1 To names.Count
        frm.Controls(names(i)).Value = vals(i)
    Next
End Sub

' The bound controls on frm whose value differs from the saved one: each
' one's name into names, the field it shows (for the message) into fields,
' and its value into vals.
Private Sub ChangedFields(frm As Form, names As Collection, fields As Collection, vals As Collection)
    Dim c As Control, src As String
    For Each c In frm.Controls
        Select Case c.ControlType
            Case acTextBox, acComboBox, acCheckBox, acToggleButton, acListBox, acOptionGroup
                src = c.ControlSource
                If Len(src) > 0 And Left$(src, 1) <> "=" Then
                    If Not SameValue(c.Value, c.OldValue) Then
                        names.Add c.Name
                        fields.Add Replace(Replace(src, "[", ""), "]", "")
                        vals.Add c.Value
                    End If
                End If
        End Select
    Next
End Sub

Private Function SameValue(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsNull(a) Or IsNull(b) Then
        SameValue = IsNull(a) And IsNull(b)
    Else
        SameValue = (StrComp(CStr(a), CStr(b), vbBinaryCompare) = 0)
    End If
End Function

' Wait ms milliseconds, letting Access paint.
Private Sub Pause(ByVal ms As Long)
    Dim t As Single
    t = Timer
    Do
        DoEvents
        Sleep 15
    Loop While (Timer - t) * 1000 < ms And Timer >= t
End Sub
