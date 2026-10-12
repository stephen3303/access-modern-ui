Option Compare Database
Option Explicit

' Access UI kit - small file and form helpers the other kit modules use.

' True when a file exists.
Public Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Len(path) > 0) And (Len(Dir$(path, vbNormal)) > 0)
End Function

' True when a folder exists.
Public Function FolderExists(ByVal path As String) As Boolean
    On Error Resume Next
    FolderExists = (Len(path) > 0) And ((GetAttr(path) And vbDirectory) = vbDirectory)
End Function

' Create a folder and any missing parents.
Public Sub EnsureFolder(ByVal path As String)
    Dim parts() As String, i As Long, cur As String
    If FolderExists(path) Then Exit Sub
    parts = Split(path, "\")
    cur = parts(0)
    Dim first As Long
    first = 1
    If Left$(path, 2) = Chr$(92) & Chr$(92) Then
        If UBound(parts) < 3 Then Err.Raise 5, , "A UNC folder needs a server and share"
        cur = Chr$(92) & Chr$(92) & parts(2) & Chr$(92) & parts(3)
        If Not FolderExists(cur) Then Err.Raise 76, , "Network share unavailable: " & cur
        first = 4
    End If
    For i = first To UBound(parts)
        If Len(parts(i)) > 0 Then
            cur = cur & "\" & parts(i)
            If Not FolderExists(cur) Then MkDir cur
        End If
    Next
End Sub

' True when the form is open in any view.
Public Function IsFormOpen(ByVal name As String) As Boolean
    IsFormOpen = (SysCmd(acSysCmdGetObjectState, acForm, name) <> 0)
End Function

' Open a modal dialog and report whether it was confirmed.
Public Function RunDialog(ByVal formName As String, Optional ByVal openArgs As String = "") As Boolean
    TempVars("DialogResult") = ""
    DoCmd.OpenForm formName, , , , , acDialog, openArgs
    RunDialog = (Nz(TempVars("DialogResult"), "") = "ok")
End Function
