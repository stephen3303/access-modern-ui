Option Compare Database
Option Explicit

' Access UI kit - who is using the app, and who created or last changed each
' record.
'
' WinUser() is the Windows account name, read from Windows itself
' (GetUserNameW): the USERNAME environment variable can be changed by
' anyone before Access starts, so it is only the fallback. KitUser() is the
' name to record: the app's signed-in user (TempVars("Username"), set by an
' app with its own sign-in) or, without one, WinUser().
'
' Stamps: a table that several users edit gets four columns, added by
' AddStampFields at build or upgrade time:
'   created_by TEXT(150), created_at DATETIME, updated_by TEXT(150), updated_at DATETIME
' A bound form fills them from its BeforeUpdate:
'   Private Sub Form_BeforeUpdate(Cancel As Integer)
'       StampRecord Me
'   End Sub
' and code that writes with SQL adds StampSql():
'   KitExecute "UPDATE item SET status='Closed', " & StampSql() & " WHERE id=" & id
'   KitExecute "INSERT INTO item (name, " & StampColumns() & ") VALUES ('Pump', " & StampValues() & ")"

Public Const STAMP_CREATED_BY As String = "created_by"
Public Const STAMP_CREATED_AT As String = "created_at"
Public Const STAMP_UPDATED_BY As String = "updated_by"
Public Const STAMP_UPDATED_AT As String = "updated_at"
Public Const STAMP_USER_LEN As Long = 150      ' the by columns' width, as error_log.username

Private Declare PtrSafe Function GetUserNameW Lib "advapi32" (ByVal lpBuffer As LongPtr, nSize As Long) As Long
Private Declare PtrSafe Function GetComputerNameW Lib "kernel32" (ByVal lpBuffer As LongPtr, nSize As Long) As Long

Private m_winUser As String
Private m_computer As String

' The Windows account running Access ("jsmith"), read once from Windows; the
' USERNAME environment variable only if Windows doesn't answer.
Public Function WinUser() As String
    Dim buf As String, n As Long
    If Len(m_winUser) = 0 Then
        n = 257                                 ' UNLEN + 1
        buf = String$(n, vbNullChar)
        If GetUserNameW(StrPtr(buf), n) <> 0 And n > 1 Then m_winUser = Left$(buf, n - 1)
        If Len(m_winUser) = 0 Then m_winUser = Environ$("USERNAME")
    End If
    WinUser = m_winUser
End Function

' The computer's name ("NAVY-PC-042"), read once from Windows; the
' COMPUTERNAME environment variable only if Windows doesn't answer.
Public Function WinComputer() As String
    Dim buf As String, n As Long
    If Len(m_computer) = 0 Then
        n = 256
        buf = String$(n, vbNullChar)
        If GetComputerNameW(StrPtr(buf), n) <> 0 Then m_computer = Left$(buf, n)
        If Len(m_computer) = 0 Then m_computer = Environ$("COMPUTERNAME")
    End If
    WinComputer = m_computer
End Function

' The name to record as the person doing something: the app's signed-in
' user (TempVars("Username")) when it has one, else WinUser().
Public Function KitUser() As String
    Dim u As String
    On Error Resume Next                        ' TempVars may be unset
    u = Nz(TempVars("Username"), "")
    On Error GoTo 0
    If Len(u) = 0 Then u = WinUser()
    KitUser = Left$(u, STAMP_USER_LEN)
End Function

' A bound form's BeforeUpdate: fill the stamp fields its record source has.
' A new record gets created_by and created_at; every save sets updated_by
' and updated_at. Fields the record source lacks are skipped, so a form can
' call it before its table has all four.
Public Sub StampRecord(frm As Form)
    Dim t As Date, u As String
    t = Now()
    u = KitUser()
    If frm.NewRecord Then
        SetField frm, STAMP_CREATED_BY, u
        SetField frm, STAMP_CREATED_AT, t
    End If
    SetField frm, STAMP_UPDATED_BY, u
    SetField frm, STAMP_UPDATED_AT, t
End Sub

' For an UPDATE's SET list: "updated_by='jsmith', updated_at=Now()".
Public Function StampSql() As String
    StampSql = STAMP_UPDATED_BY & "=" & SqlText(KitUser()) & ", " & STAMP_UPDATED_AT & "=Now()"
End Function

' For an INSERT's column list, with StampValues: all four stamp columns.
Public Function StampColumns() As String
    StampColumns = STAMP_CREATED_BY & ", " & STAMP_CREATED_AT & ", " & STAMP_UPDATED_BY & ", " & STAMP_UPDATED_AT
End Function

' For an INSERT's VALUES list, matching StampColumns.
Public Function StampValues() As String
    Dim u As String
    u = SqlText(KitUser())
    StampValues = u & ", Now(), " & u & ", Now()"
End Function

' Build or upgrade step: add the stamp columns table lacks, in db (default
' the current database; pass the back end's to change a linked table). Only
' adds columns, never changes or removes one. Returns how many it added.
Public Function AddStampFields(ByVal table As String, Optional db As DAO.Database) As Long
    Dim cols As Variant, i As Long, td As DAO.TableDef
    On Error GoTo Fail
    If db Is Nothing Then Set db = CurrentDb
    cols = Array(STAMP_CREATED_BY & " TEXT(" & STAMP_USER_LEN & ")", STAMP_CREATED_AT & " DATETIME", _
        STAMP_UPDATED_BY & " TEXT(" & STAMP_USER_LEN & ")", STAMP_UPDATED_AT & " DATETIME")
    Set td = db.TableDefs(table)
    For i = 0 To UBound(cols)
        If Not HasField(td, Split(cols(i), " ")(0)) Then
            db.Execute "ALTER TABLE [" & table & "] ADD COLUMN " & cols(i), dbFailOnError
            AddStampFields = AddStampFields + 1
        End If
    Next
    Exit Function
Fail:
    RaiseAgain "modKitUser.AddStampFields", "table " & table
End Function

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

' Set field nm of frm's current record, when its record source has it.
Private Sub SetField(frm As Form, ByVal nm As String, ByVal v As Variant)
    If Not HasField(frm.Recordset, nm) Then Exit Sub
    frm(nm).Value = v                           ' the record being saved, not the recordset's
End Sub

' True when obj (a TableDef or Recordset) has field nm.
Private Function HasField(obj As Object, ByVal nm As String) As Boolean
    Dim fld As DAO.Field
    For Each fld In obj.Fields
        If StrComp(fld.Name, nm, vbTextCompare) = 0 Then HasField = True: Exit Function
    Next
End Function

' Text as a quoted SQL literal.
Private Function SqlText(ByVal s As String) As String
    SqlText = "'" & Replace(s, "'", "''") & "'"
End Function
