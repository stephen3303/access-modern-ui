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
' The question comes after the failed save is over: from KitSaveRecord when
' it was saving, or, for a save Access started itself (moving to another
' record, closing), from the form's timer, borrowed for one tick and then
' put back as it was.

Public Const LOCK_TRIES As Long = 5            ' retries after the first try
Public Const LOCK_WAIT_MS As Long = 120        ' first wait; each retry waits longer
Public Const ERR_WRITE_CONFLICT As Long = 7787 ' Access's Write Conflict window
Public Const ERR_DATA_CHANGED As Long = 7878   ' "The data has been changed"

Private Const LOCK_MSG As String = "Someone else is saving this record right now. Wait a moment, then save again."
Private Const LOCK_RAISE As String = "Another user is saving the same records right now. Wait a moment, then try again."

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private m_testAnswer As Long    ' tests only: the conflict answer, 0 to ask
Private m_inSave As Boolean     ' KitSaveRecord is running (its retries wait with DoEvents)
Private m_saving As Boolean     ' KitSaveRecord is saving: KitFormError leaves locks and conflicts to it
Private m_formLock As Long      ' the lock error KitFormError passed on to KitSaveRecord
' A write conflict noted by KitFormError, resolved once the failed save is
' over (Undo, Refresh and setting values fail inside Form_Error).
Private m_conflictForm As String
Private m_names As Collection   ' changed controls
Private m_fields As Collection  ' what the user calls them, for the message
Private m_vals As Collection    ' the user's values
Private m_timerForm As String   ' the form whose timer KitLocksTick borrowed
Private m_oldOnTimer As String  ' that form's own timer settings
Private m_oldInterval As Long

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
    If m_inSave Then Exit Function              ' a click during a retry's wait: the first save carries on
    m_inSave = True
    On Error GoTo Fail
    KitSaveRecord = SaveTry(frm)
    m_inSave = False
    Exit Function
Fail:
    m_inSave = False
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
            NoteConflict frm
            ' A save Access started itself (moving to another record, closing)
            ' has no KitSaveRecord to finish it: borrow the form's timer.
            If Not m_saving Then StartConflictTimer frm
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

' The form's timer, borrowed by KitFormError: the failed save is over, so
' put the form's timer back and let the user decide. Public only so the
' form's OnTimer can call it.
Public Function KitLocksTick() As Boolean
    Dim frm As Form, nm As String
    On Error GoTo Fail
    nm = m_timerForm
    m_timerForm = ""
    If Len(nm) = 0 Then Exit Function
    Set frm = Forms(nm)
    frm.TimerInterval = m_oldInterval
    frm.OnTimer = m_oldOnTimer
    If m_conflictForm = nm Then ResolveConflict frm  ' unless a KitSaveRecord got there first
    Exit Function
Fail:
    m_conflictForm = ""
    LogError "modKitLocks.KitLocksTick", nm
End Function

' Tests only: answer the next conflicts with vbYes or vbNo instead of
' asking (0 asks again).
Public Sub SetConflictAnswer(ByVal answer As Long)
    m_testAnswer = answer
End Sub

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

' KitSaveRecord's work, retried while locked (KitSaveRecord keeps a second
' click out while it waits).
Private Function SaveTry(frm As Form) As Boolean
    Dim attempt As Long, num As Long
TryAgain:
    On Error GoTo Fail
    m_formLock = 0
    m_conflictForm = ""
    m_saving = True
    If frm.Dirty Then frm.Dirty = False
    m_saving = False
    If Len(m_conflictForm) > 0 Then GoTo Conflict   ' Access gave up the save without an error
    SaveTry = Not frm.Dirty
    Exit Function
Conflict:
    ResolveConflict frm                 ' after the failed save, never inside it
    Exit Function
Fail:
    m_saving = False
    If Len(m_conflictForm) > 0 Then Resume Conflict ' a write conflict: the user decides
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
    RaiseAgain "modKitLocks.SaveTry", frm.Name
End Function

' Form_Error, during the failed save: note which bound controls the user
' changed and their values, for ResolveConflict to use afterwards.
Private Sub NoteConflict(frm As Form)
    Dim c As Control, src As String
    Set m_names = New Collection
    Set m_fields = New Collection
    Set m_vals = New Collection
    m_conflictForm = frm.Name
    For Each c In frm.Controls
        Select Case c.ControlType
            Case acTextBox, acComboBox, acCheckBox, acToggleButton, acListBox, acOptionGroup
                src = c.ControlSource
                If Len(src) > 0 And Left$(src, 1) <> "=" Then
                    If Not SameValue(c.Value, c.OldValue) Then
                        m_names.Add c.Name
                        m_fields.Add FieldCaption(c, src)
                        m_vals.Add c.Value
                    End If
                End If
        End Select
    Next
End Sub

' After the failed save: drop the user's changes, show the other user's
' saved record, then ask. Yes puts the noted values back on top (the record
' is changed again, for them to save); No leaves the other user's version.
Private Sub ResolveConflict(frm As Form)
    Dim i As Long, list As String, answer As Long, names As Collection, vals As Collection, fields As Collection
    Set names = m_names: Set vals = m_vals: Set fields = m_fields
    m_conflictForm = ""
    frm.Undo
    frm.Refresh                     ' the other user's saved values
    If names Is Nothing Then Exit Sub
    If names.Count = 0 Then Exit Sub
    For i = 1 To fields.Count
        If i <= 5 Then list = list & IIf(i > 1, ", ", "") & fields(i)
    Next
    If fields.Count > 5 Then list = list & " and " & (fields.Count - 5) & " more"
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

' Run ResolveConflict from the form's timer, once Access has finished the
' save it started; the form's own timer settings come back first.
Private Sub StartConflictTimer(frm As Form)
    If m_timerForm = frm.Name Then Exit Sub         ' already borrowed
    m_timerForm = frm.Name
    m_oldOnTimer = frm.OnTimer
    m_oldInterval = frm.TimerInterval
    frm.OnTimer = "=KitLocksTick()"
    frm.TimerInterval = 1
End Sub

' What the user calls the field c shows: its attached label's caption
' without a trailing colon or "*", else the field name.
Private Function FieldCaption(c As Control, ByVal src As String) As String
    Dim cap As String
    On Error Resume Next                ' no attached label
    cap = c.Controls(0).Caption
    On Error GoTo 0
    cap = Trim$(Replace(cap, "&", ""))
    Do While Len(cap) > 0 And (Right$(cap, 1) = ":" Or Right$(cap, 1) = "*")
        cap = Trim$(Left$(cap, Len(cap) - 1))
    Loop
    If Len(cap) = 0 Then cap = Replace(Replace(src, "[", ""), "]", "")
    FieldCaption = cap
End Function

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
