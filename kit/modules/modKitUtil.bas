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

' Show or hide a control. Access raises error 2165 when hiding the control that
' has the focus, or a subform, tab control or option group that holds it (a
' button that hides itself, a tab switch that moved the focus onto the page).
' Try first; only on 2165 move the focus away and try again. When nothing else
' on the form can take the focus, disabling c drops it (Access allows that),
' then c is hidden and its Enabled setting put back.
' A rounded pill's image (<name>_bg) goes with it, see PillImageVisible.
Public Sub SetVisible(c As Control, ByVal v As Boolean)
    Dim wasEnabled As Boolean
    PillImageVisible c, v
    If c.Visible = v Then Exit Sub
    On Error GoTo Focused
    c.Visible = v
    Exit Sub
Focused:
    If Err.Number <> ERR_HIDE_FOCUSED Then Err.Raise Err.Number, Err.Source, Err.Description
    Resume Park
Park:
    On Error GoTo 0
    If ParkFocus(c) Then
        c.Visible = v
    Else
        wasEnabled = c.Enabled
        c.Enabled = False
        c.Visible = v
        c.Enabled = wasEnabled
    End If
End Sub

' The pill image <name>_bg behind c (RoundCF, PillLabel), if c has one, follows
' c: hiding c hides it, so a card that is closed or cleared leaves no empty
' pill behind. Showing c shows a bound image again (its ControlSource picks
' the color, or none); a painted image stays as the app's paint left it,
' since only a fresh paint knows its color.
Private Sub PillImageVisible(c As Control, ByVal v As Boolean)
    Dim im As Control, src As String
    On Error Resume Next
    Set im = OwnerForm(c).Controls(c.Name & "_bg")
    ' Read through Properties: im.ControlSource raises 438 on an image control.
    src = Nz(im.Properties("ControlSource").Value, "")
    On Error GoTo 0
    If im Is Nothing Then Exit Sub
    If Not v Then
        im.Visible = False
    ElseIf Len(src) > 0 Then
        im.Visible = True
    End If
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
' False when nothing could take it.
Private Function ParkFocus(c As Control) As Boolean
    Dim frm As Form, o As Control
    Set frm = OwnerForm(c)
    If c.Name <> "txtFocusSink" And HasFocusSink(frm) Then
        If Not Holds(c, frm.Controls("txtFocusSink")) Then
            If TryFocus(frm.Controls("txtFocusSink")) Then ParkFocus = True: Exit Function
        End If
    End If
    For Each o In frm.Controls
        If o.Name <> c.Name Then
            If Not Holds(c, o) Then
                If TryFocus(o) Then ParkFocus = True: Exit Function
            End If
        End If
    Next
End Function

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

' Open a modal dialog and report whether it was confirmed. The app behind it
' is dimmed while it is open (DimApp, modBackdrop) unless dimBehind is False.
' The dim is removed when the dialog closes, and also when it fails to open,
' before the error goes back to the caller.
Public Function RunDialog(ByVal formName As String, Optional ByVal openArgs As String = "", _
                          Optional ByVal dimBehind As Boolean = True) As Boolean
    Dim dimmed As Boolean, num As Long, desc As String, src As String
    On Error GoTo Fail
    TempVars("DialogResult") = ""
    If dimBehind Then DimApp: dimmed = True
    DoCmd.OpenForm formName, , , , , acDialog, openArgs
    If dimmed Then UndimApp: dimmed = False
    RunDialog = (Nz(TempVars("DialogResult"), "") = "ok")
    Exit Function
Fail:
    ' UndimApp resets Err, so keep the error to pass it on.
    num = Err.Number: desc = Err.Description: src = Err.Source
    If dimmed Then UndimApp
    Err.Number = num: Err.Description = desc: Err.Source = src
    RaiseAgain "modKitUtil.RunDialog", formName
End Function
