Option Compare Database
Option Explicit

' Access UI kit - run-time helpers for pages built with modUIKit: tables that
' scale their columns (GridFit), list/detail pages (SelectFirstRow,
' ShowDetail, ClearDetail), tabs and filter chips with counts (SetCount,
' StyleChip), the splash screen's status line (SplashStatus) and messages in
' the kit's look (KitMsgBox). Colors set here go through Th() so they follow
' the palette.

' Width the table's vertical scroll bar takes (twips).
Private Const SCROLL_W As Long = 300
' Widest a form can be: 22 inches (twips).
Private Const MAX_FORM_W As Long = 31680
' The DateBox the date picker (dlgDatePick) is open for (PickDate).
Private m_dateTarget As Control
' The message the message dialog (dlgMessage) is showing (KitMsgBox), and the
' button chosen in it.
Private m_msgOpen As Boolean
Private m_msgText As String
Private m_msgStyle As Long
Private m_msgTitle As String
Private m_msgAnswer As Long
' The splash screen BuildSplash builds (SplashStatus, CloseSplash).
Private Const SPLASH_FORM As String = "frmSplash"

' ---------------------------------------------------------------------------
' Tables
' ---------------------------------------------------------------------------

' Proportional fit for a table built with BuildGrid(..., fit:=True): scale
' every column of the table shown in subform control sf to the control's
' current width, and move the captions GridHeaders drew on the parent with
' prefix. Columns never get narrower than designed. Call it from the host's
' Form_Resize.
Public Sub GridFit(sf As Control, Optional ByVal prefix As String = "")
    Dim host As Form, g As Form, c As Control, i As Long, t() As String
    Dim design As Long, avail As Long, ratio As Double, newW As Long
    Set host = sf.Parent
    Set g = sf.Form
    If Left$(Nz(g.Tag, ""), 5) <> "grid:" Then Exit Sub
    design = Val(Mid$(g.Tag, 6))
    ' An anchored control's Width keeps its design value, so add how much
    ' wider than designed the host is.
    avail = sf.Width
    If sf.HorizontalAnchor = 2 Then avail = avail + host.InsideWidth - host.Width
    ratio = (avail - SCROLL_W) / (design - SCROLL_W)
    If ratio < 1 Then ratio = 1
    If design * ratio > MAX_FORM_W Then ratio = MAX_FORM_W / design
    newW = CLng(design * ratio)
    ' Access ignores g.Width = n on a form shown in a subform: the form only
    ' grows when a control's right edge pushes it, and moving a column past
    ' its edge raises error 5. So widen the row background first (it pushes
    ' the form out), place the columns, then shrink the background.
    ' Resizing the form can leave an earlier reference to it dead (error 5 on
    ' any use), so it is read from sf again before each use.
    If newW > g.Width Then
        If HasCtl(g, "txtBg") Then g.Controls("txtBg").Width = newW
        Set g = sf.Form
        If HasCtl(g, "boxRowLine") Then g.Controls("boxRowLine").Width = newW
    End If
    Do
        Set g = sf.Form
        Set c = Nothing
        On Error Resume Next
        Set c = g.Controls("txtCol" & i)
        On Error GoTo 0
        If c Is Nothing Then Exit Do
        t = Split(c.Tag, "|")
        If UBound(t) >= 2 Then
            FitCtl c, CLng(Val(t(1)) * ratio), CLng(Val(t(2)) * ratio) - px(8)
            Set g = sf.Form
            Set c = g.Controls("txtCol" & i)
            If HasCtl(g, c.Name & "_bg") Then FitCtl g.Controls(c.Name & "_bg"), c.Left, c.Width
        End If
        If Len(prefix) > 0 Then
            If HasCtl(host, prefix & i) Then
                With host.Controls(prefix & i)
                    t = Split(.Tag, "|")
                    If UBound(t) >= 2 Then FitCtl host.Controls(prefix & i), _
                        CLng(Val(t(0)) + Val(t(1)) * ratio), CLng((Val(t(2)) + px(8)) * ratio) - px(8)
                End With
            End If
        End If
        i = i + 1
    Loop
    Set g = sf.Form
    If HasCtl(g, "txtBg") Then FitCtl g.Controls("txtBg"), 0, newW
    Set g = sf.Form
    If HasCtl(g, "boxRowLine") Then FitCtl g.Controls("boxRowLine"), 0, newW
    Set g = sf.Form
    ' Narrowing the form is cosmetic (no empty strip to scroll to), and
    ' Access may ignore it.
    If newW < g.Width Then
        On Error Resume Next
        g.Width = newW
    End If
End Sub

' Move and size a control horizontally.
Private Sub FitCtl(c As Control, ByVal x As Long, ByVal w As Long)
    If w < 15 Then w = 15
    If c.Left <> x Then c.Left = x
    If c.Width <> w Then c.Width = w
End Sub

' True when the form has a control named nm.
Private Function HasCtl(f As Form, ByVal nm As String) As Boolean
    Dim c As Control
    If Len(nm) = 0 Then Exit Function
    On Error Resume Next
    Set c = f.Controls(nm)
    HasCtl = Not c Is Nothing
End Function

' Show a table's vertical scroll bar only when its rows don't all fit:
' Access draws the bar on a continuous form whatever its content. A BuildGrid
' table calls this from its OnCurrent, which fires whenever its record source
' is set or requeried (and once before the host sets it: no record source,
' no rows). Rows are counted only until one more than fits, so a large linked
' table isn't read to the end on every row change. A failure is logged, not
' shown: it only leaves the bar.
Public Function GridScroll(g As Form)
    Dim n As Long, sb As Integer, rowH As Long, fits As Long
    On Error GoTo Fail
    rowH = g.Section(acDetail).Height
    If Len(g.RecordSource) > 0 And rowH > 0 Then
        fits = g.InsideHeight \ rowH
        With g.RecordsetClone
            If Not (.BOF And .EOF) Then .MoveFirst   ' a new clone's position isn't defined
            Do While Not .EOF And n <= fits
                n = n + 1
                .MoveNext
            Loop
        End With
    End If
    ' Only the vertical bar (2) changes; a horizontal one (1, set by ResizeGrid) stays.
    sb = (g.ScrollBars And 1) Or IIf(n * rowH > g.InsideHeight, 2, 0)
    If g.ScrollBars <> sb Then g.ScrollBars = sb
    Exit Function
Fail:
    If Err.Number = 7951 Then Exit Function     ' no recordset yet: nothing to scroll
    LogError "modKitUI.GridScroll", g.Name
End Function

' GridScroll for every BuildGrid table on frm and on the forms inside it.
' Call it from the window's Form_Resize: a table anchored to the window
' changes height without its OnCurrent firing. A subform still loading (the
' window resizes as it opens) is skipped; GridScroll logs its own failures.
Public Sub GridScrollAll(frm As Form)
    Dim c As Control
    On Error Resume Next
    For Each c In frm.Controls
        If c.ControlType = acSubform Then
            If Len(c.SourceObject) > 0 Then
                If c.Form.Tag Like "grid:*" Then GridScroll c.Form Else GridScrollAll c.Form
            End If
        End If
    Next
End Sub

' Select the first row of a BuildGrid table and tell the host, as a click
' would (its GridClick grid, False). Returns the row's id, or Null when the
' table is empty (then nothing is called). Call it after the list loads or
' its filter changes.
Public Function SelectFirstRow(sf As Control) As Variant
    Dim g As Form
    SelectFirstRow = Null
    Set g = sf.Form
    If g.RecordsetClone.RecordCount = 0 Then
        g.txtSel = Null
        Exit Function
    End If
    g.Recordset.MoveFirst
    g.txtSel = g!id
    g.Recalc
    SelectFirstRow = g.txtSel
    sf.Parent.GridClick sf.Name, False
End Function

' ---------------------------------------------------------------------------
' List and detail pages
' ---------------------------------------------------------------------------

' Show the records matching where in the detail panel: subform control
' detailCtl on host, showing form formName. A changed record in the panel is
' saved first, and the empty-state control emptyCtl is hidden. The panel's
' form is only swapped when formName changes. Swapping a subform while one of
' its own events is running is unsafe, so call this from the list (GridClick)
' or the page, never from inside the detail form.
Public Sub ShowDetail(host As Form, ByVal detailCtl As String, ByVal formName As String, ByVal where As String, _
                      Optional ByVal emptyCtl As String = "lblDetailEmpty")
    Dim sf As Control
    On Error GoTo Fail
    Set sf = host.Controls(detailCtl)
    SaveDetail sf
    If sf.SourceObject <> formName And sf.SourceObject <> "Form." & formName Then sf.SourceObject = formName
    sf.Form.Filter = where
    sf.Form.FilterOn = True
    SetVisible sf, True
    If HasCtl(host, emptyCtl) Then SetVisible host.Controls(emptyCtl), False
    Exit Sub
Fail:
    RaiseAgain "modKitUI.ShowDetail", formName & " where " & where
End Sub

' Empty the detail panel (saving a changed record first) and show the
' empty-state control instead.
Public Sub ClearDetail(host As Form, ByVal detailCtl As String, Optional ByVal emptyCtl As String = "lblDetailEmpty")
    Dim sf As Control
    On Error GoTo Fail
    Set sf = host.Controls(detailCtl)
    SaveDetail sf
    SetVisible sf, False
    sf.SourceObject = ""
    If HasCtl(host, emptyCtl) Then SetVisible host.Controls(emptyCtl), True
    Exit Sub
Fail:
    RaiseAgain "modKitUI.ClearDetail", detailCtl
End Sub

' Save the record shown in a detail panel if it changed.
Private Sub SaveDetail(sf As Control)
    If Len(sf.SourceObject) = 0 Then Exit Sub
    If sf.Form.Dirty Then sf.Form.Dirty = False
End Sub

' ---------------------------------------------------------------------------
' Date fields
' ---------------------------------------------------------------------------

' Click on a DateBox's calendar button: open the date picker for the box nm
' and wait for it to close. Does nothing while the box is disabled or locked,
' or its form is read-only (AllowEdits off, unless on a record being added).
Public Function PickDate(frm As Form, ByVal nm As String)
    Dim c As Control
    On Error GoTo Fail
    Set c = frm.Controls(nm)
    If Not c.Enabled Or c.Locked Then Exit Function
    If Not frm.AllowEdits And Not (frm.NewRecord And frm.AllowAdditions) Then Exit Function
    c.SetFocus
    Set m_dateTarget = c
    DoCmd.OpenForm "dlgDatePick", WindowMode:=acDialog
    Set m_dateTarget = Nothing
    Exit Function
Fail:
    Set m_dateTarget = Nothing
    HandleError "modKitUI.PickDate", nm
End Function

' The open date picker's box: its value (Null when none is open), and
' setting it. Setting a value does not fire the box's own update events; the
' record's BeforeUpdate still checks it when it is saved.
Public Function DateTargetValue() As Variant
    DateTargetValue = Null
    If Not m_dateTarget Is Nothing Then DateTargetValue = m_dateTarget.Value
End Function

Public Sub SetDateTarget(ByVal v As Variant)
    If m_dateTarget Is Nothing Then Exit Sub
    m_dateTarget.Value = v
End Sub

' ---------------------------------------------------------------------------
' Splash screen
' ---------------------------------------------------------------------------

' Show msg on the splash screen's status line (frmSplash, BuildSplash) while
' the app starts, for example "Opening the data file...". Does nothing once
' the splash has closed, so start-up code can call it either way.
Public Sub SplashStatus(ByVal msg As String)
    If Not IsFormOpen(SPLASH_FORM) Then Exit Sub
    Forms(SPLASH_FORM).Controls("lblSplashStatus").Caption = msg
    Forms(SPLASH_FORM).Repaint
End Sub

' Close the splash screen if it is open. KitMsgBox calls it: the splash is a
' pop-up, so a message opened during start-up could sit behind it. Access can
' refuse (2585) while the splash's own Timer is still running; the splash then
' closes itself when that event ends, so the refusal is ignored.
Public Sub CloseSplash()
    If Not IsFormOpen(SPLASH_FORM) Then Exit Sub
    On Error Resume Next
    DoCmd.Close acForm, SPLASH_FORM, acSaveNo
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' Messages
' ---------------------------------------------------------------------------

' MsgBox in the kit's look: opens the message dialog (dlgMessage, built by
' BuildMessageDialog) and waits for a button. Windows centres a MsgBox on the
' monitor, so on a large or second screen it opens outside Access and Access
' waits behind it; dlgMessage is centred over Access like the app's other
' dialogs. Takes the buttons vbOKOnly, vbOKCancel, vbYesNo and vbYesNoCancel,
' an icon (vbInformation, vbExclamation, vbQuestion, vbCritical) and
' vbDefaultButton2/3, and returns vbOK, vbCancel, vbYes or vbNo like MsgBox.
' Any other button set, a message while one is already showing, or a dialog
' that can't open (not built yet) uses MsgBox itself.
Public Function KitMsgBox(ByVal msg As String, Optional ByVal style As VbMsgBoxStyle = vbOKOnly, _
        Optional ByVal title As String = "") As VbMsgBoxResult
    Dim btns As Long
    btns = style And 7
    CloseSplash                    ' a start-up message must not open behind the splash
    If m_msgOpen Then GoTo Plain
    If Not (btns = vbOKOnly Or btns = vbOKCancel Or btns = vbYesNo Or btns = vbYesNoCancel) Then GoTo Plain
    On Error GoTo Failed
    m_msgText = msg
    m_msgStyle = style
    m_msgTitle = title
    m_msgAnswer = 0
    m_msgOpen = True
    DoCmd.OpenForm "dlgMessage", WindowMode:=acDialog
    m_msgOpen = False
    If m_msgAnswer = 0 Then        ' closed some other way: as Esc
        m_msgAnswer = Choose(btns + 1, vbOK, vbCancel, vbCancel, vbCancel, vbNo, vbCancel)
    End If
    KitMsgBox = m_msgAnswer
    Exit Function
Failed:
    ' Only this message's own dialog failed to open; a message already showing
    ' (Plain from the first line) keeps m_msgOpen set until it closes.
    m_msgOpen = False
Plain:
    KitMsgBox = MsgBox(msg, style, title)
End Function

' What the open message dialog shows, and its answer (dlgMessage calls these).
Public Function KitMsgText() As String
    KitMsgText = m_msgText
End Function

Public Function KitMsgStyle() As Long
    KitMsgStyle = m_msgStyle
End Function

Public Function KitMsgTitle() As String
    KitMsgTitle = m_msgTitle
End Function

Public Sub SetKitMsgAnswer(ByVal answer As Long)
    m_msgAnswer = answer
End Sub

' ---------------------------------------------------------------------------
' Tabs and filter chips
' ---------------------------------------------------------------------------

' Show a count after a tab or chip caption, "Name (n)"; Null shows the name
' alone. The caption without the count is kept in the control's Tag, so the
' count can change any number of times.
Public Sub SetCount(c As Control, ByVal n As Variant)
    Dim cap As String
    If Len(c.Tag) = 0 Then
        cap = c.Caption
        If Right$(cap, 1) = ")" And InStrRev(cap, " (") > 0 Then cap = Left$(cap, InStrRev(cap, " (") - 1)
        c.Tag = cap
    End If
    c.Caption = c.Tag & IIf(IsNull(n), "", " (" & n & ")")
End Sub

' Filter chip (a label): filled CLR_PRIMARY with white text when on,
' transparent with muted text when off. n, when given, is its count (SetCount).
Public Sub StyleChip(lbl As Control, ByVal isOn As Boolean, Optional ByVal n As Variant)
    lbl.BackStyle = IIf(isOn, 1, 0)
    lbl.BackColor = Th(CLR_PRIMARY)
    lbl.ForeColor = IIf(isOn, Th(CLR_WHITE), Th(CLR_MUTED_FG))
    If Not IsMissing(n) Then SetCount lbl, n
End Sub
