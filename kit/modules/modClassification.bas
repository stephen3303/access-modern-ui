Option Compare Database
Option Explicit

' Access UI kit - the classification banner, and the labels an app may use.
'
' The app has one selected classification (its ceiling): ClassSet "SECRET".
' It limits the labels the app permits: only markings at or below that level
' (SECRET allows UNCLASSIFIED, CUI, CONFIDENTIAL and SECRET, not TOP SECRET).
' Label lists in forms come from ClassLabels() / ClassFillList; a value typed
' or imported is checked with ClassLabelPermitted(). The banner shows the
' selected classification, in its marking color, across the top and/or bottom
' of a form.
'
' Build (frmMain, or any form; after ShellBuild leave room with its top and
' bottom arguments, CLASS_BANNER_H):
'   ClassBannerBuild "clsTop", 0, 0, formWidth
'   ClassBannerBuild "clsBottom", 0, formHeight - CLASS_BANNER_H, formWidth
' Run time (the form's Load, and after ClassChoose):
'   ClassBannerShow Me
' A label list: ClassFillList Me.cboMarking
' A "Classification" sidebar link or button:
'   =ClassChoose(Me.btnClass)      (a menu of the levels; the banner updates)
'
' The levels and labels below are generic and unclassified. They are the
' commonly used marking names only; an app whose policy differs replaces them
' with ClassSetLevels / ClassSetLabels at startup. Markings are not
' validated against any specific regulation, and the banner is a visual
' reminder, not an access control.
'
' The choice is kept in TempVars("ClassLevel") for the session and remembered
' in the app's settings (Program Settings\<APP_NAME>\Classification).

Public Const CLASS_BANNER_H As Long = 300       ' the banner's height, 20px (twips)

Private Const SETTINGS_KEY As String = "Classification"
Private Const TEMP_LEVEL As String = "ClassLevel"
Private Const BANNER_PREFIX As String = "cls"   ' banner names start with this

' Levels, lowest first: "NAME|RGB(back)|RGB(text)" as hex RRGGBB pairs.
Private m_levels As Variant
' Labels: "level index|label", any order; shown in this order.
Private m_labels As Variant

Private m_ready As Boolean

' ---------------------------------------------------------------------------
' Levels and labels
' ---------------------------------------------------------------------------

Private Sub Init()
    If m_ready Then Exit Sub
    ' The customary banner colors: green, purple, blue, red, orange.
    m_levels = Array( _
        "UNCLASSIFIED|007A33|FFFFFF", _
        "CUI|502B85|FFFFFF", _
        "CONFIDENTIAL|0033A0|FFFFFF", _
        "SECRET|C8102E|FFFFFF", _
        "TOP SECRET|FF671F|000000")
    m_labels = Array( _
        "0|UNCLASSIFIED", _
        "0|UNCLASSIFIED//FOR OFFICIAL USE ONLY", _
        "1|CUI", _
        "1|CUI//SP-PRVCY", _
        "1|CUI//SP-PROPIN", _
        "1|CUI//NOFORN", _
        "1|CUI//REL TO USA, FVEY", _
        "2|CONFIDENTIAL", _
        "2|CONFIDENTIAL//NOFORN", _
        "2|CONFIDENTIAL//REL TO USA, FVEY", _
        "3|SECRET", _
        "3|SECRET//NOFORN", _
        "3|SECRET//ORCON", _
        "3|SECRET//REL TO USA, FVEY", _
        "3|SECRET//RELIDO", _
        "4|TOP SECRET", _
        "4|TOP SECRET//NOFORN", _
        "4|TOP SECRET//SCI", _
        "4|TOP SECRET//REL TO USA, FVEY")
    m_ready = True
End Sub

' Replace the levels, lowest first: each "NAME|RRGGBB back|RRGGBB text".
' Clears the labels (set them with ClassSetLabels). Raises error 5 for a bad
' entry.
Public Sub ClassSetLevels(levels As Variant)
    Dim i As Long, p() As String
    Init
    If Not IsArray(levels) Then Err.Raise 5, , "Classification levels must be an array"
    For i = LBound(levels) To UBound(levels)
        p = Split(levels(i), "|")
        If UBound(p) <> 2 Then Err.Raise 5, , "Bad classification level: " & levels(i)
        If Len(Trim$(p(0))) = 0 Or Len(p(1)) <> 6 Or Len(p(2)) <> 6 Then Err.Raise 5, , "Bad classification level: " & levels(i)
    Next
    m_levels = levels
    m_labels = Array()
End Sub

' Replace the labels: each "level index|label", the index counting from 0 in
' the levels' order. Raises error 5 for a bad entry.
Public Sub ClassSetLabels(labels As Variant)
    Dim i As Long, p() As String
    Init
    If Not IsArray(labels) Then Err.Raise 5, , "Classification labels must be an array"
    For i = LBound(labels) To UBound(labels)
        p = Split(labels(i), "|", 2)
        If UBound(p) <> 1 Then Err.Raise 5, , "Bad classification label: " & labels(i)
        If Not IsNumeric(p(0)) Then Err.Raise 5, , "Bad classification label: " & labels(i)
        If CLng(p(0)) < 0 Or CLng(p(0)) > UBound(m_levels) Then Err.Raise 5, , "Bad classification label: " & labels(i)
    Next
    m_labels = labels
End Sub

' Number of levels.
Public Function ClassLevelCount() As Long
    Init
    ClassLevelCount = UBound(m_levels) + 1
End Function

' The name of level i (0 = lowest).
Public Function ClassLevelName(ByVal i As Long) As String
    Init
    If i < 0 Or i > UBound(m_levels) Then Err.Raise 5, , "No classification level " & i
    ClassLevelName = Split(m_levels(i), "|")(0)
End Function

' The index of a level name ("secret" or "SECRET"), -1 when unknown.
Public Function ClassLevelIndex(ByVal level As String) As Long
    Dim i As Long
    Init
    ClassLevelIndex = -1
    level = UCase$(Trim$(level))
    For i = 0 To UBound(m_levels)
        If Split(m_levels(i), "|")(0) = level Then ClassLevelIndex = i: Exit Function
    Next
End Function

' ---------------------------------------------------------------------------
' The selected classification
' ---------------------------------------------------------------------------

' The selected level's name: this session's choice, else the remembered one,
' else the lowest level. A remembered name that is no longer a level is
' ignored.
Public Function ClassCurrent() As String
    Dim s As String
    Init
    On Error Resume Next                        ' TempVars may be unset
    s = Nz(TempVars(TEMP_LEVEL), "")
    On Error GoTo 0
    If ClassLevelIndex(s) < 0 Then
        s = GetSetting(APP_NAME, SETTINGS_KEY, "Level", "")
        If ClassLevelIndex(s) < 0 Then s = ClassLevelName(0)
        TempVars(TEMP_LEVEL) = s
    End If
    ClassCurrent = UCase$(s)
End Function

' Select the app's classification and remember it. Raises error 5 for an
' unknown level. Banners already open show it after ClassBannerShow.
Public Sub ClassSet(ByVal level As String)
    Dim i As Long
    i = ClassLevelIndex(level)
    If i < 0 Then Err.Raise 5, , "Unknown classification level: " & level
    TempVars(TEMP_LEVEL) = ClassLevelName(i)
    SaveSetting APP_NAME, SETTINGS_KEY, "Level", ClassLevelName(i)
End Sub

' ---------------------------------------------------------------------------
' Permitted labels
' ---------------------------------------------------------------------------

' The labels the selected classification permits: those at or below it, in
' list order. Empty array when none.
Public Function ClassLabels() As Variant
    Dim out() As String, n As Long, i As Long, p() As String, cur As Long
    Init
    cur = ClassLevelIndex(ClassCurrent())
    ReDim out(0 To UBound(m_labels) + 1)
    For i = 0 To UBound(m_labels)
        p = Split(m_labels(i), "|", 2)
        If CLng(p(0)) <= cur Then out(n) = p(1): n = n + 1
    Next
    If n = 0 Then
        ClassLabels = Array()
    Else
        ReDim Preserve out(0 To n - 1)
        ClassLabels = out
    End If
End Function

' True when the label is in the list and its level is at or below the
' selected classification. Case and spacing around // are ignored.
Public Function ClassLabelPermitted(ByVal label As String) As Boolean
    Dim v As Variant
    label = Normalize(label)
    If Len(label) = 0 Then Exit Function
    For Each v In ClassLabels()
        If Normalize(CStr(v)) = label Then ClassLabelPermitted = True: Exit Function
    Next
End Function

' The label's level index, -1 when the label isn't in the list.
Public Function ClassLabelLevel(ByVal label As String) As Long
    Dim i As Long, p() As String
    Init
    ClassLabelLevel = -1
    label = Normalize(label)
    For i = 0 To UBound(m_labels)
        p = Split(m_labels(i), "|", 2)
        If Normalize(p(1)) = label Then ClassLabelLevel = CLng(p(0)): Exit Function
    Next
End Function

' A save check: Null or "" when the label is permitted, else the reason, for
' a form's BeforeUpdate or an import:
'   msg = ClassLabelProblem(Me.cboMarking): If Len(msg) > 0 Then KitMsgBox msg: Cancel = True
Public Function ClassLabelProblem(ByVal label As String) As String
    Dim lvl As Long
    If Len(Trim$(label)) = 0 Then ClassLabelProblem = "Choose a classification marking.": Exit Function
    lvl = ClassLabelLevel(label)
    If lvl < 0 Then
        ClassLabelProblem = "'" & label & "' is not a marking this app recognises."
    ElseIf Not ClassLabelPermitted(label) Then
        ClassLabelProblem = "'" & label & "' is above this app's classification (" & ClassCurrent() & ")."
    End If
End Function

' Fill a combo or list box with the permitted labels (Value List). A value the
' control already holds that is no longer permitted is cleared. Returns the
' number of labels.
Public Function ClassFillList(c As Control) As Long
    Dim v As Variant, s As String, n As Long
    For Each v In ClassLabels()
        s = s & IIf(Len(s) > 0, ";", "") & """" & Replace(CStr(v), """", """""") & """"
        n = n + 1
    Next
    c.RowSourceType = "Value List"
    c.RowSource = s
    If c.ControlType = acComboBox Then c.LimitToList = True
    If Not IsNull(c.Value) Then
        If Not ClassLabelPermitted(CStr(c.Value)) Then c.Value = Null
    End If
    ClassFillList = n
End Function

' Upper case, // with no spaces around it, runs of spaces as one.
Private Function Normalize(ByVal s As String) As String
    s = UCase$(Trim$(s))
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    s = Replace(s, " //", "//")
    s = Replace(s, "// ", "//")
    Normalize = s
End Function

' ---------------------------------------------------------------------------
' The banner
' ---------------------------------------------------------------------------

' Build time: a banner band at (x, y), w wide, on the form being built
' (NewForm), named nm (starts with "cls"). Bold 9pt centered text; the color
' and text are set by ClassBannerShow. Stretches with the form's width.
Public Sub ClassBannerBuild(ByVal nm As String, ByVal x As Long, ByVal y As Long, ByVal w As Long)
    Dim c As Control
    If LCase$(Left$(nm, Len(BANNER_PREFIX))) <> BANNER_PREFIX Then Err.Raise 5, , "A banner name starts with '" & BANNER_PREFIX & "': " & nm
    Init
    Set c = Box(nm, x, y, w, CLASS_BANNER_H, LevelColor(0, 1))
    c.Tag = "class:banner"
    Anchor c, 2, 0
    Set c = Lbl(nm & "_text", ClassLevelName(0), x, y, w, CLASS_BANNER_H, 9, True, vbWhite)
    c.TextAlign = 2
    c.Tag = "class:banner"
    Anchor c, 2, 0
End Sub

' Show the selected classification (or the label, when given) on every
' banner of the form, in its marking color. Call from the form's Load and
' after the classification changes. Does nothing for a form with no banner.
' A label above the selected classification shows the selected one.
Public Sub ClassBannerShow(frm As Form, Optional ByVal label As String = "")
    Dim c As Control, lvl As Long, txt As String
    Init
    txt = ClassCurrent()
    lvl = ClassLevelIndex(txt)
    If Len(label) > 0 Then
        If ClassLabelPermitted(label) Then txt = UCase$(Trim$(label)): lvl = ClassLabelLevel(label)
    End If
    For Each c In frm.Controls
        If c.Tag = "class:banner" Then
            If c.ControlType = acRectangle Then
                c.BackColor = LevelColor(lvl, 1)
            ElseIf c.ControlType = acLabel Then
                c.Caption = txt
                c.ForeColor = LevelColor(lvl, 2)
            End If
        End If
    Next
End Sub

' A menu of the levels under anchor; the picked one becomes the selected
' classification, and the banners on the anchor's form are refreshed. Returns
' the picked level, "" when closed with nothing picked. Levels over a label
' already in use are the app's to check (ClassLabelProblem).
Public Function ClassChoose(anchor As Control) As String
    Dim items() As Variant, i As Long, key As String, cur As String
    Init
    cur = ClassCurrent()
    ReDim items(0 To UBound(m_levels))
    For i = 0 To UBound(m_levels)
        items(i) = CStr(i) & "|" & ClassLevelName(i) & IIf(ClassLevelName(i) = cur, "|E73E", "")
    Next
    key = KitMenuFromArray(anchor, items)
    If Len(key) = 0 Then Exit Function
    ClassSet ClassLevelName(CLng(key))
    ClassBannerShow anchor.Parent
    ClassChoose = ClassLevelName(CLng(key))
End Function

' KitMenu takes its items as a ParamArray; pass an array of up to 8 on.
Private Function KitMenuFromArray(anchor As Control, items() As Variant) As String
    Select Case UBound(items)
        Case 0: KitMenuFromArray = KitMenu(anchor, items(0))
        Case 1: KitMenuFromArray = KitMenu(anchor, items(0), items(1))
        Case 2: KitMenuFromArray = KitMenu(anchor, items(0), items(1), items(2))
        Case 3: KitMenuFromArray = KitMenu(anchor, items(0), items(1), items(2), items(3))
        Case 4: KitMenuFromArray = KitMenu(anchor, items(0), items(1), items(2), items(3), items(4))
        Case 5: KitMenuFromArray = KitMenu(anchor, items(0), items(1), items(2), items(3), items(4), items(5))
        Case 6: KitMenuFromArray = KitMenu(anchor, items(0), items(1), items(2), items(3), items(4), items(5), items(6))
        Case 7: KitMenuFromArray = KitMenu(anchor, items(0), items(1), items(2), items(3), items(4), items(5), items(6), items(7))
        Case Else: Err.Raise 5, , "ClassChoose lists at most 8 levels"
    End Select
End Function

' Color of level i: part 1 = back, 2 = text. Marking colors are fixed (they
' don't follow the palette), so they are not Th() colors.
Private Function LevelColor(ByVal i As Long, ByVal part As Long) As Long
    Dim hex6 As String
    hex6 = Split(m_levels(i), "|")(part)
    LevelColor = RGB(CLng("&H" & Left$(hex6, 2)), CLng("&H" & Mid$(hex6, 3, 2)), CLng("&H" & Mid$(hex6, 5, 2)))
End Function
