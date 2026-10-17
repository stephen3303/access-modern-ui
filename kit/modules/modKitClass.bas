Option Compare Database
Option Explicit

' Access UI kit - the classification banner, and the classification labels
' the app permits.
'
' The app has one SELECTED CLASSIFICATION, the highest it is accredited to
' hold. It shows as a colored banner across the top and bottom of the main
' window (and a report's page header and footer), and it limits which
' classification labels (markings) a record, document or export may carry: a
' label is permitted when its level is no higher than the selected one.
'
' Levels, lowest first, with the banner color of each (fixed colors, the same
' in every palette):
'   UNCLASSIFIED (green)   CUI (purple)   CONFIDENTIAL (blue)   SECRET (red)
'   TOP SECRET (orange)    TOP SECRET//SCI (yellow, dark text)
' A label's level is the marking's classification (CUI//NOFORN is CUI) and a
' label with an SCI control (TOP SECRET//SI/TK) is TOP SECRET//SCI, so an
' app selected at TOP SECRET permits TOP SECRET//NOFORN but not
' TOP SECRET//SI.
'
' Build (once, with the app's forms):
'   ShellBuild APP_NAME, "...", pages, links, PAGE_W, PAGE_H, CLASS_BANNER_H, CLASS_BANNER_H
'   BuildClassBanner "lblClassTop", 0, SHELL_W + PAGE_W
'   BuildClassBanner "lblClassBottom", CLASS_BANNER_H + PAGE_H, SHELL_W + PAGE_W, 1
'   SaveForm SHELL_FORM
' (ShellLoad paints the banners; any other form calls KitClassApplyBanners Me
' from its Load, and a report from its Open.)
'
' Run time:
'   KitClassSelect "SECRET"                    ' an administrator's choice (saved in the database)
'   KitClassChoose Me.btnClass                 ' a menu of the levels, then KitClassSelect
'   Me.cboMarking.RowSource = KitClassRowSource()   ' only permitted labels
'   If Not KitClassGuard(Me.cboMarking) Then Cancel = True   ' in BeforeUpdate
' Who may change the selected classification is the app's rule (an admin
' check before KitClassSelect or KitClassChoose); the kit only stores it.
' Until one is selected the app is UNCLASSIFIED, the most restrictive.

Public Const CLASS_BANNER_H As Long = 300       ' a banner's height: 20px (twips)
Public Const CLASS_TAG As String = "class:banner"   ' a banner label's Tag
Private Const CLASS_PROP As String = "Classification"   ' the database property

Private Const LEVEL_COUNT As Long = 6

' ---------------------------------------------------------------------------
' Levels
' ---------------------------------------------------------------------------

' The level names, lowest first (index 0 to 5).
Public Function KitClassLevels() As Variant
    KitClassLevels = Array("UNCLASSIFIED", "CUI", "CONFIDENTIAL", "SECRET", "TOP SECRET", "TOP SECRET//SCI")
End Function

' The level number of a level name or a label (0 to 5), -1 when it is not a
' classification. Case and spacing around // do not matter. A label's level
' is its marking's, raised to TOP SECRET//SCI by an SCI control.
Public Function KitClassLevel(ByVal text As String) As Long
    Dim t As String, base As String, rest As String, at As Long, i As Long
    t = UCase$(Replace(Replace(Trim$(text), " //", "//"), "// ", "//"))
    KitClassLevel = -1
    If Len(t) = 0 Then Exit Function
    at = InStr(t, "//")
    If at > 0 Then
        base = Left$(t, at - 1)
        rest = Mid$(t, at + 2)
    Else
        base = t
    End If
    For i = 0 To 4
        If base = KitClassLevels()(i) Then KitClassLevel = i: Exit For
    Next
    If KitClassLevel = 4 And HasSci(rest) Then KitClassLevel = 5
End Function

' True when the controls after TOP SECRET// name an SCI compartment system.
Private Function HasSci(ByVal rest As String) As Boolean
    Dim part As Variant
    rest = Replace(rest, "//", "/")
    For Each part In Split(rest, "/")
        Select Case Trim$(CStr(part))
            Case "SCI", "SI", "SI-G", "TK", "HCS", "HCS-P", "KLM", "RSV": HasSci = True
        End Select
    Next
End Function

' ---------------------------------------------------------------------------
' Labels
' ---------------------------------------------------------------------------

' Every label the kit knows, the commonly used markings, lowest first. An app
' that needs more adds them to KitClassLabels.
Public Function KitClassLabels() As Variant
    KitClassLabels = Array( _
        "UNCLASSIFIED", "UNCLASSIFIED//FOUO", _
        "CUI", "CUI//NOFORN", "CUI//FEDCON", "CUI//SP-PRVCY", "CUI//SP-PROPIN", "CUI//SP-CTI", "CUI//SP-EXPT", _
        "CONFIDENTIAL", "CONFIDENTIAL//NOFORN", "CONFIDENTIAL//REL TO USA, FVEY", _
        "SECRET", "SECRET//NOFORN", "SECRET//REL TO USA, FVEY", "SECRET//ORCON", "SECRET//NOFORN/ORCON", _
        "TOP SECRET", "TOP SECRET//NOFORN", "TOP SECRET//REL TO USA, FVEY", _
        "TOP SECRET//SI", "TOP SECRET//SI/TK", "TOP SECRET//SI/TK//NOFORN", "TOP SECRET//HCS", "TOP SECRET//SCI")
End Function

' The labels the selected classification permits, in the same order.
Public Function KitClassPermittedLabels() As Variant
    Dim labels As Variant, out() As String, i As Long, n As Long, sel As Long
    labels = KitClassLabels()
    sel = KitClassLevel(KitClassSelected())
    ReDim out(0 To UBound(labels))
    For i = 0 To UBound(labels)
        If KitClassLevel(CStr(labels(i))) <= sel Then out(n) = labels(i): n = n + 1
    Next
    ReDim Preserve out(0 To n - 1)
    KitClassPermittedLabels = out
End Function

' True when the selected classification permits label (a classification the
' kit knows, at or below the selected level).
Public Function KitClassPermitted(ByVal label As String) As Boolean
    Dim lvl As Long
    lvl = KitClassLevel(label)
    KitClassPermitted = (lvl >= 0 And lvl <= KitClassLevel(KitClassSelected()))
End Function

' "" when label is permitted, else a sentence saying why not (for a message).
Public Function KitClassProblem(ByVal label As String) As String
    If Len(Trim$(label)) = 0 Then Exit Function
    If KitClassLevel(label) < 0 Then
        KitClassProblem = """" & label & """ is not a classification label."
    ElseIf Not KitClassPermitted(label) Then
        KitClassProblem = """" & label & """ is above this application's classification (" & KitClassSelected() & ")."
    End If
End Function

' For a bound form's BeforeUpdate or a control's BeforeUpdate: True when the
' control's label is permitted (or empty); otherwise tells the user why and
' returns False, so the caller sets Cancel.
'   If Not KitClassGuard(Me.cboMarking) Then Cancel = True
Public Function KitClassGuard(ctl As Control) As Boolean
    Dim msg As String
    msg = KitClassProblem(CStr(Nz(ctl.Value, "")))
    If Len(msg) = 0 Then KitClassGuard = True: Exit Function
    KitMsgBox msg, vbExclamation, "Classification"
End Function

' A combo box's Value List of the permitted labels (RowSourceType "Value
' List"). Quote marks in a label are doubled; none of the kit's have any.
Public Function KitClassRowSource() As String
    Dim v As Variant, s As String
    For Each v In KitClassPermittedLabels()
        s = s & IIf(Len(s) > 0, ";", "") & """" & Replace(CStr(v), """", """""") & """"
    Next
    KitClassRowSource = s
End Function

' Point a combo box at the permitted labels, and clear a value it holds that
' is no longer permitted (the selected classification was lowered). Returns
' True when it cleared one.
Public Function KitClassFillCombo(cbo As Control) As Boolean
    cbo.RowSourceType = "Value List"
    cbo.RowSource = KitClassRowSource()
    cbo.LimitToList = True
    If Len(KitClassProblem(CStr(Nz(cbo.Value, "")))) > 0 Then
        cbo.Value = Null
        KitClassFillCombo = True
    End If
End Function

' A SQL condition for a text field holding labels: only permitted ones, e.g.
' "[marking] IN ('UNCLASSIFIED','CUI',...)". Use it to hide records marked
' above the application's classification.
Public Function KitClassInSql(ByVal field As String) As String
    Dim v As Variant, s As String
    For Each v In KitClassPermittedLabels()
        s = s & IIf(Len(s) > 0, ",", "") & "'" & Replace(CStr(v), "'", "''") & "'"
    Next
    KitClassInSql = "[" & field & "] IN (" & s & ")"
End Function

' ---------------------------------------------------------------------------
' The selected classification
' ---------------------------------------------------------------------------

' The selected classification: a name from KitClassLevels, UNCLASSIFIED until
' one is selected.
Public Function KitClassSelected() As String
    Dim v As String, i As Long
    On Error Resume Next
    v = CStr(CurrentDb.Properties(CLASS_PROP))
    On Error GoTo 0
    i = IndexOfLevel(Trim$(v))
    If i < 0 Then i = 0
    KitClassSelected = KitClassLevels()(i)
End Function

' The index of v in KitClassLevels, -1 when v is not a level name.
Private Function IndexOfLevel(ByVal v As String) As Long
    Dim i As Long
    IndexOfLevel = -1
    For i = 0 To LEVEL_COUNT - 1
        If UCase$(v) = KitClassLevels()(i) Then IndexOfLevel = i: Exit For
    Next
End Function

' Select the classification (a name from KitClassLevels), save it in the
' database and repaint every open banner. Raises error 5 for any other text.
Public Sub KitClassSelect(ByVal level As String)
    Dim db As DAO.Database, prop As DAO.Property, i As Long, n As Long, d As String
    i = IndexOfLevel(Trim$(level))
    If i < 0 Then Err.Raise 5, , "Not a classification: " & level
    level = KitClassLevels()(i)
    Set db = CurrentDb
    On Error Resume Next
    db.Properties(CLASS_PROP) = level
    If Err.Number = 3270 Then         ' property not found: create it
        Err.Clear
        Set prop = db.CreateProperty(CLASS_PROP, dbText, level)
        db.Properties.Append prop
    End If
    If Err.Number <> 0 Then
        n = Err.Number: d = Err.Description
        On Error GoTo 0
        Err.Raise n, , "Could not save the classification: " & d
    End If
    On Error GoTo 0
    KitClassRepaintAll
End Sub

' Open a menu of the levels under anchor and select the one picked (the
' current one is marked). Returns True when the selection changed. The app
' decides who may call it.
Public Function KitClassChoose(anchor As Control) As Boolean
    Dim items(0 To LEVEL_COUNT - 1) As Variant, i As Long, pick As String, cur As String
    cur = KitClassSelected()
    For i = 0 To LEVEL_COUNT - 1
        items(i) = "L" & i & "|" & KitClassLevels()(i) & IIf(KitClassLevels()(i) = cur, "|E73E", "")
    Next
    pick = KitMenu(anchor, items(0), items(1), items(2), items(3), items(4), items(5))
    If Len(pick) = 0 Then Exit Function
    i = CLng(Mid$(pick, 2))
    If KitClassLevels()(i) = cur Then Exit Function
    KitClassSelect CStr(KitClassLevels()(i))
    KitClassChoose = True
End Function

' ---------------------------------------------------------------------------
' Banner
' ---------------------------------------------------------------------------

' The banner color of a level name or label.
Public Function KitClassColor(ByVal text As String) As Long
    Select Case KitClassLevel(text)
        Case 1: KitClassColor = CLR_CLASS_CUI
        Case 2: KitClassColor = CLR_CLASS_CONFIDENTIAL
        Case 3: KitClassColor = CLR_CLASS_SECRET
        Case 4: KitClassColor = CLR_CLASS_TOPSECRET
        Case 5: KitClassColor = CLR_CLASS_SCI
        Case Else: KitClassColor = CLR_CLASS_UNCLASSIFIED
    End Select
End Function

' The text color on that banner: dark on yellow, white on the rest.
Public Function KitClassForeColor(ByVal text As String) As Long
    If KitClassLevel(text) = 5 Then KitClassForeColor = CLR_CLASS_SCI_FG Else KitClassForeColor = Th(CLR_WHITE)
End Function

' Build a banner label nm across the form at y, w wide, showing the selected
' classification (Tag CLASS_TAG, painted by KitClassApplyBanners). vAnchor 0
' keeps it to the top as the window resizes, 1 to the bottom. In a report it
' goes in section sec (page header or footer).
Public Function BuildClassBanner(ByVal nm As String, ByVal y As Long, ByVal w As Long, Optional ByVal vAnchor As Integer = 0, _
        Optional ByVal sec As Integer = acDetail) As Control
    Dim c As Control
    Set c = Lbl(nm, "UNCLASSIFIED", 0, y, w, CLASS_BANNER_H, 8, True, CLR_WHITE, sec)
    c.BackStyle = 1
    c.BackColor = CLR_CLASS_UNCLASSIFIED
    c.TextAlign = 2
    c.FontWeight = 700
    c.Tag = CLASS_TAG
    c.ControlTipText = "Classification of this application"
    On Error Resume Next                        ' a report's controls take no anchors
    Anchor c, 2, vAnchor
    On Error GoTo 0
    Set BuildClassBanner = c
End Function

' Paint one banner label with the selected classification.
Public Sub KitClassPaint(lbl As Control)
    Dim sel As String
    sel = KitClassSelected()
    lbl.Caption = sel
    lbl.BackColor = KitClassColor(sel)
    lbl.ForeColor = KitClassForeColor(sel)
End Sub

' Paint every banner (Tag CLASS_TAG) on a form or report.
Public Sub KitClassApplyBanners(frm As Object)
    Dim c As Control
    For Each c In frm.Controls
        If c.ControlType = acLabel Then
            If Nz(c.Tag, "") = CLASS_TAG Then KitClassPaint c
        End If
    Next
End Sub

' Repaint the banners of every open form and report (after a new selection).
Public Sub KitClassRepaintAll()
    Dim f As Form, r As Report
    On Error Resume Next
    For Each f In Forms
        KitClassApplyBanners f
    Next
    For Each r In Reports
        KitClassApplyBanners r
    Next
End Sub
