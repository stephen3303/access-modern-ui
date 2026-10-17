Option Compare Database
Option Explicit

' Access UI kit - the classification banner: a bar across the top and bottom of
' frmMain that names the app's selected classification in the standard color,
' and a ranked list of markings, so the selected level limits which labels the
' rest of the app may use.
'
' Levels, lowest to highest (a level is a color and a rank):
'   0 UNCLASSIFIED (green)   1 CUI (purple)   2 CONFIDENTIAL (blue)
'   3 SECRET (red)           4 TOP SECRET (orange)   5 TOP SECRET//SCI (yellow)
' The list of markings (ClassMarkings) holds the common ones for each level:
' UNCLASSIFIED//FOUO, CUI//SP-PRVCY, SECRET//NOFORN, TOP SECRET//SCI and so on.
' A marking is permitted when its level is no higher than the selected one: an
' app set to SECRET permits UNCLASSIFIED through SECRET//REL TO USA, FVEY and
' refuses TOP SECRET. Matching ignores case and spaces around "//".
'
' In the build step, leave room for the bars and add them after ShellBuild:
'   ShellBuild title, subtitle, pages, links, PAGE_W, PAGE_H, BannerHeight(), BannerHeight()
'   BannerBuild
'   SaveForm SHELL_FORM
' At startup (before frmMain opens), select the app's classification:
'   SetClassification "SECRET//NOFORN"
' ShellLoad then paints the bars, and SetClassification repaints them while
' frmMain is open. Wherever the app takes a label from the user or from data
' (a record's marking, an export's header), check it:
'   If Not ClassPermits(lbl) Then KitMsgBox ClassProblem(lbl), vbExclamation, "Classification"
'   cboMarking.RowSource = ClassMarkingList()    ' a combo box limited to the permitted labels
'
' The selected marking is held for the session. SetClassification remember:=True
' also saves it in the registry (as the sidebar's collapsed state is), and it is
' read back when none was set; with neither, the app is UNCLASSIFIED. Colors are
' the CLR_CLASS_* constants that tools\make_themes.py writes into modTheme; they
' are the same in every palette. The app provides APP_NAME.

Public Const CLASS_UNCLASSIFIED As Long = 0
Public Const CLASS_CUI As Long = 1
Public Const CLASS_CONFIDENTIAL As Long = 2
Public Const CLASS_SECRET As Long = 3
Public Const CLASS_TOP_SECRET As Long = 4
Public Const CLASS_TS_SCI As Long = 5

Public Const BANNER_TAG As String = "banner"      ' the bars' Tag
Private Const BANNER_PX As Long = 22              ' a bar's height (px)

Private m_selected As String    ' the selected marking; "" until set or read

' ---------------------------------------------------------------------------
' The list
' ---------------------------------------------------------------------------

' The markings in rank order, each "Label|level". Every level has its plain
' name first.
Private Function MarkingTable() As Variant
    MarkingTable = Array( _
        "UNCLASSIFIED|0", "UNCLASSIFIED//FOUO|0", _
        "CUI|1", "CUI//SP-PRVCY|1", "CUI//SP-PROPIN|1", "CUI//SP-EXPT|1", "CUI//SP-CTI|1", "FOR OFFICIAL USE ONLY|1", _
        "CONFIDENTIAL|2", "CONFIDENTIAL//NOFORN|2", "CONFIDENTIAL//REL TO USA, FVEY|2", _
        "SECRET|3", "SECRET//NOFORN|3", "SECRET//REL TO USA, FVEY|3", _
        "TOP SECRET|4", "TOP SECRET//NOFORN|4", "TOP SECRET//REL TO USA, FVEY|4", _
        "TOP SECRET//SCI|5", "TOP SECRET//SCI//NOFORN|5")
End Function

' The name of a level ("SECRET"), or "" for a level out of range.
Public Function ClassLevelName(ByVal level As Long) As String
    Select Case level
        Case CLASS_UNCLASSIFIED: ClassLevelName = "UNCLASSIFIED"
        Case CLASS_CUI: ClassLevelName = "CUI"
        Case CLASS_CONFIDENTIAL: ClassLevelName = "CONFIDENTIAL"
        Case CLASS_SECRET: ClassLevelName = "SECRET"
        Case CLASS_TOP_SECRET: ClassLevelName = "TOP SECRET"
        Case CLASS_TS_SCI: ClassLevelName = "TOP SECRET//SCI"
    End Select
End Function

' A level's banner color, or -1 for a level out of range.
Public Function ClassLevelBack(ByVal level As Long) As Long
    Select Case level
        Case CLASS_UNCLASSIFIED: ClassLevelBack = CLR_CLASS_UNCLASS
        Case CLASS_CUI: ClassLevelBack = CLR_CLASS_CUI
        Case CLASS_CONFIDENTIAL: ClassLevelBack = CLR_CLASS_CONF
        Case CLASS_SECRET: ClassLevelBack = CLR_CLASS_SECRET
        Case CLASS_TOP_SECRET: ClassLevelBack = CLR_CLASS_TS
        Case CLASS_TS_SCI: ClassLevelBack = CLR_CLASS_SCI
        Case Else: ClassLevelBack = -1
    End Select
End Function

' The text color on a level's banner: black on the orange and yellow bars,
' white on the rest.
Public Function ClassLevelFore(ByVal level As Long) As Long
    ClassLevelFore = IIf(level >= CLASS_TOP_SECRET, CLR_CLASS_DARK, CLR_WHITE)
End Function

' The level of a marking: 0 to 5, or -1 when it is not in the list. Case and
' spaces around "//" and at the ends are ignored.
Public Function ClassLevelOf(ByVal label As String) As Long
    Dim m As Variant, p() As String
    label = NormalizeMarking(label)
    ClassLevelOf = -1
    If Len(label) = 0 Then Exit Function
    For Each m In MarkingTable()
        p = Split(m, "|")
        If NormalizeMarking(p(0)) = label Then ClassLevelOf = CLng(p(1)): Exit Function
    Next
End Function

' Upper case, with no spaces around "//" or at the ends.
Private Function NormalizeMarking(ByVal s As String) As String
    s = UCase$(Trim$(s))
    Do While InStr(s, " //") > 0 Or InStr(s, "// ") > 0
        s = Replace(Replace(s, " //", "//"), "// ", "//")
    Loop
    NormalizeMarking = s
End Function

' The markings up to maxLevel, in rank order, as an array. Without maxLevel
' (-1), up to the selected level: what the app permits.
Public Function ClassMarkings(Optional ByVal maxLevel As Long = -1) As Variant
    Dim m As Variant, p() As String, out() As String, n As Long
    If maxLevel < 0 Then maxLevel = ClassSelectedLevel()
    ReDim out(0 To UBound(MarkingTable()))
    For Each m In MarkingTable()
        p = Split(m, "|")
        If CLng(p(1)) <= maxLevel Then out(n) = p(0): n = n + 1
    Next
    If n = 0 Then ClassMarkings = Array(): Exit Function
    ReDim Preserve out(0 To n - 1)
    ClassMarkings = out
End Function

' The permitted markings as a value list for a combo box's RowSource
' (RowSourceType "Value List"): "UNCLASSIFIED;CUI;..." . Labels have no semicolons.
Public Function ClassMarkingList(Optional ByVal maxLevel As Long = -1) As String
    ClassMarkingList = Join(ClassMarkings(maxLevel), ";")
End Function

' ---------------------------------------------------------------------------
' The selection
' ---------------------------------------------------------------------------

' Select the app's classification: any marking in the list (error 5 for one
' that isn't). The open frmMain's bars are repainted. remember saves the
' choice in the registry for the next session.
Public Sub SetClassification(ByVal label As String, Optional ByVal remember As Boolean = False)
    If ClassLevelOf(label) < 0 Then Err.Raise 5, , "SetClassification: not a known marking: " & label
    m_selected = CanonicalMarking(label)
    If remember Then SaveSetting APP_NAME, "UI", "Classification", m_selected
    If IsFormOpen(SHELL_FORM) Then BannerRefresh Forms(SHELL_FORM)
End Sub

' The selected marking as the list spells it: set this session, else saved by
' SetClassification remember:=True, else UNCLASSIFIED.
Public Function ClassSelected() As String
    If Len(m_selected) = 0 Then
        m_selected = GetSetting(APP_NAME, "UI", "Classification", "")
        If ClassLevelOf(m_selected) < 0 Then m_selected = "UNCLASSIFIED"
    End If
    ClassSelected = m_selected
End Function

' The selected marking's level, 0 to 5.
Public Function ClassSelectedLevel() As Long
    ClassSelectedLevel = ClassLevelOf(ClassSelected())
End Function

' A marking spelled as the list spells it ("secret // noforn" -> "SECRET//NOFORN").
Private Function CanonicalMarking(ByVal label As String) As String
    Dim m As Variant, p() As String
    label = NormalizeMarking(label)
    For Each m In MarkingTable()
        p = Split(m, "|")
        If NormalizeMarking(p(0)) = label Then CanonicalMarking = p(0): Exit Function
    Next
End Function

' True when label is in the list and no higher than the selected level. A label
' that isn't in the list is never permitted.
Public Function ClassPermits(ByVal label As String) As Boolean
    Dim lv As Long
    lv = ClassLevelOf(label)
    ClassPermits = (lv >= 0 And lv <= ClassSelectedLevel())
End Function

' "" when label is permitted, else a sentence for the user saying why not.
Public Function ClassProblem(ByVal label As String) As String
    Dim lv As Long
    lv = ClassLevelOf(label)
    If lv < 0 Then
        ClassProblem = """" & Trim$(label) & """ is not a marking this app knows."
    ElseIf lv > ClassSelectedLevel() Then
        ClassProblem = Trim$(label) & " is higher than this app's classification, " & ClassSelected() & "."
    End If
End Function

' ---------------------------------------------------------------------------
' The bars
' ---------------------------------------------------------------------------

' A bar's height in twips: the top and bottom arguments of ShellBuild.
Public Function BannerHeight() As Long
    BannerHeight = px(BANNER_PX)
End Function

' Add the two bars to the form under construction (F): lblBannerTop across the
' top and lblBannerBottom across the bottom, each stretching with the form.
' The form leaves BannerHeight() free at each end (ShellBuild's top and
' bottom). Both are painted by BannerRefresh when the form opens.
Public Sub BannerBuild()
    Dim c As Control, h As Long
    h = BannerHeight()
    Set c = BannerBar("lblBannerTop", 0, h): Anchor c, 2, 0
    Set c = BannerBar("lblBannerBottom", F.Section(acDetail).Height - h, h): Anchor c, 2, 1
End Sub

' One bar at y, in the build palette's UNCLASSIFIED until BannerRefresh.
Private Function BannerBar(ByVal nm As String, ByVal y As Long, ByVal h As Long) As Control
    Dim c As Control
    Set c = Lbl(nm, "UNCLASSIFIED", 0, y, F.Width, h, 9, True, CLR_WHITE)
    c.BackStyle = 1
    c.BackColor = CLR_CLASS_UNCLASS
    c.TextAlign = 2
    c.Tag = BANNER_TAG
    c.ControlTipText = "The classification of this app"
    Set BannerBar = c
End Function

' Show the selected classification on every bar (control tagged "banner") of
' frm: its marking, in its level's colors. ShellLoad calls it, and
' SetClassification does while frmMain is open.
Public Sub BannerRefresh(frm As Form)
    Dim c As Control, lv As Long, mark As String
    mark = ClassSelected()
    lv = ClassLevelOf(mark)
    For Each c In frm.Controls
        If c.Tag = BANNER_TAG Then
            c.Caption = mark
            c.BackColor = ClassLevelBack(lv)
            c.ForeColor = ClassLevelFore(lv)
        End If
    Next
End Sub
