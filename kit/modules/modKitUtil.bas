Option Compare Database
Option Explicit

' Access UI kit - small file and form helpers the other kit modules use.

Private Const ERR_DISABLE_FOCUSED As Long = 2164   ' can't disable the focused control
Private Const ERR_HIDE_FOCUSED As Long = 2165      ' can't hide the focused control

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
    For i = 1 To UBound(parts)
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

' Show or hide a control. Access raises error 2165 when hiding the control that
' has the focus, or a subform, tab control or option group that holds it (a
' button that hides itself, a tab switch that moved the focus onto the page).
' Try first; only on 2165 move the focus away and try again.
Public Sub SetVisible(c As Control, ByVal v As Boolean)
    If c.Visible = v Then Exit Sub
    On Error GoTo Focused
    c.Visible = v
    Exit Sub
Focused:
    If Err.Number <> ERR_HIDE_FOCUSED Then Err.Raise Err.Number, Err.Source, Err.Description
    Resume Park
Park:
    On Error GoTo 0
    ParkFocus c
    c.Visible = v
End Sub

' Enable or disable a control. Access raises error 2164 when disabling the
' control that has the focus (a button that disables itself when clicked), so
' this moves the focus away first, like SetVisible.
Public Sub SetEnabled(c As Control, ByVal v As Boolean)
    If c.Enabled = v Then Exit Sub
    On Error GoTo Focused
    c.Enabled = v
    Exit Sub
Focused:
    If Err.Number <> ERR_DISABLE_FOCUSED Then Err.Raise Err.Number, Err.Source, Err.Description
    Resume Park
Park:
    On Error GoTo 0
    ParkFocus c
    c.Enabled = v
End Sub

' Move the focus out of c: to the form's txtFocusSink, or, when the form has
' none or it can't take the focus, to the first tab stop that c doesn't hold.
Private Sub ParkFocus(c As Control)
    Dim frm As Form, o As Control
    Set frm = OwnerForm(c)
    If HasFocusSink(frm) Then
        If Not Holds(c, frm.Controls("txtFocusSink")) Then
            If TryFocus(frm.Controls("txtFocusSink")) Then Exit Sub
        End If
    End If
    For Each o In frm.Controls
        If o.Name <> c.Name Then
            If Not Holds(c, o) Then
                If TryFocus(o) Then Exit Sub
            End If
        End If
    Next
End Sub

' The form a control sits on (its Parent can be a tab page, an option group or,
' for an attached label, another control).
Private Function OwnerForm(c As Control) As Form
    Dim p As Object
    Set p = c.Parent
    Do Until TypeOf p Is Form
        Set p = p.Parent
    Loop
    Set OwnerForm = p
End Function

' True when o sits inside c (on one of its tab pages or in its option group).
' Names are unique on a form, tab pages included, so they identify controls.
Private Function Holds(c As Control, o As Control) As Boolean
    Dim p As Object
    Set p = o.Parent
    Do Until TypeOf p Is Form
        If p.Name = c.Name Then Holds = True: Exit Function
        Set p = p.Parent
    Loop
End Function

Private Function HasFocusSink(frm As Form) As Boolean
    Dim o As Control
    On Error Resume Next
    Set o = frm.Controls("txtFocusSink")
    HasFocusSink = (Err.Number = 0)
End Function

' Give o the focus if it can take it. Labels, lines and images have no TabStop
' (reading it raises an error), and a control on a hidden page refuses SetFocus.
Private Function TryFocus(o As Control) As Boolean
    On Error GoTo Nope
    If o.TabStop And o.Visible And o.Enabled Then
        o.SetFocus
        TryFocus = True
    End If
Nope:
End Function

' Open a modal dialog and report whether it was confirmed.
Public Function RunDialog(ByVal formName As String, Optional ByVal openArgs As String = "") As Boolean
    TempVars("DialogResult") = ""
    DoCmd.OpenForm formName, , , , , acDialog, openArgs
    RunDialog = (Nz(TempVars("DialogResult"), "") = "ok")
End Function
