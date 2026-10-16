Option Compare Database
Option Explicit

' Access UI kit - classification banner and markings.
'
' The app has one selected classification (a level: UNCLASSIFIED, CUI,
' CONFIDENTIAL, SECRET, TOP SECRET or TOP SECRET//SCI), saved per user in
' Program Settings\<APP_NAME>\Classification, value Level (UNCLASSIFIED until
' one is chosen). That level limits the markings the app permits: a marking is
' permitted when its level is no higher than the app's, so a SECRET app takes
' UNCLASSIFIED, CUI, CONFIDENTIAL and SECRET markings and refuses TOP SECRET.
' The banner shows the app's marking (or a permitted, more specific one such as
' SECRET//NOFORN) in the standard color for its level. Banner colors are the
' same in every palette. This only labels and limits what the app accepts; it
' does not protect data, and the app's owner decides who may change the level.
'
' Build:   ClassBanner "bnrTop", 0, 0, PAGE_W          (modUIKit-style builder)
' Run:     ShowClassBanner Me                          (frmMain's Form_Load)
' Check:   If Not ClassPermits(txtMarking) Then ...    or  RequireClass marking

Private Const CLASS_SECTION As String = "Classification"
Private Const BANNER_TAG As String = "classbanner"
Public Const CLASS_BANNER_H As Long = 360              ' 24 px (twips)

' The levels a user can select, lowest first.
Public Function ClassLevels() As Variant
    ClassLevels = Array("UNCLASSIFIED", "CUI", "CONFIDENTIAL", "SECRET", "TOP SECRET", "TOP SECRET//SCI")
End Function

' Every known marking as "label|level rank" (rank = position in ClassLevels).
' The list of markings is commonly used US government banner text; add more here.
Private Function ClassTable() As Variant
    ClassTable = Array( _
        "UNCLASSIFIED|0", "UNCLASSIFIED//FOUO|1", "UNCLASSIFIED//FOR OFFICIAL USE ONLY|1", _
        "CUI|1", "CUI//NOFORN|1", "CUI//FEDCON|1", "CUI//REL TO USA, FVEY|1", _
        "CONFIDENTIAL|2", "CONFIDENTIAL//NOFORN|2", "CONFIDENTIAL//REL TO USA, FVEY|2", _
        "SECRET|3", "SECRET//NOFORN|3", "SECRET//REL TO USA, FVEY|3", "SECRET//ORCON|3", _
        "TOP SECRET|4", "TOP SECRET//NOFORN|4", "TOP SECRET//REL TO USA, FVEY|4", _
        "TOP SECRET//SCI|5", "TOP SECRET//SCI//NOFORN|5")
End Function

' Level rank of a marking (0 to 5), or -1 when the marking isn't in the list.
' Case and surrounding spaces don't matter.
Public Function ClassRank(ByVal label As String) As Long
    Dim e As Variant, p As Long
    label = UCase$(Trim$(label))
    ClassRank = -1
    For Each e In ClassTable()
        p = InStr(e, "|")
        If Left$(e, p - 1) = label Then ClassRank = CLng(Mid$(e, p + 1)): Exit Function
    Next
End Function

' The app's selected classification (a name from ClassLevels).
Public Function AppClassification() As String
    Dim s As String
    On Error Resume Next
    s = UCase$(GetSetting(APP_NAME, CLASS_SECTION, "Level", ""))
    On Error GoTo 0
    If LevelRank(s) < 0 Then s = "UNCLASSIFIED"
    AppClassification = s
End Function

' Select the app's classification and show it on the banners of the open
' forms. Raises error 5 for a name that isn't in ClassLevels. Lowering the
' level drops a banner marking the new level doesn't permit.
Public Sub SetAppClassification(ByVal level As String)
    level = UCase$(Trim$(level))
    If LevelRank(level) < 0 Then Err.Raise 5, , "Unknown classification: " & level
    SaveSetting APP_NAME, CLASS_SECTION, "Level", level
    If Not ClassPermits(BannerMarking()) Then SaveSetting APP_NAME, CLASS_SECTION, "Marking", level
    RefreshClassBanners
End Sub

' Markings the app permits, lowest first. maxLevel defaults to the app's.
Public Function ClassMarkings(Optional ByVal maxLevel As String = "") As Variant
    Dim e As Variant, p As Long, top As Long, n As Long, out() As String
    If Len(maxLevel) = 0 Then maxLevel = AppClassification()
    top = LevelRank(UCase$(Trim$(maxLevel)))
    If top < 0 Then Err.Raise 5, , "Unknown classification: " & maxLevel
    ReDim out(0 To UBound(ClassTable()))
    For Each e In ClassTable()
        p = InStr(e, "|")
        If CLng(Mid$(e, p + 1)) <= top Then out(n) = Left$(e, p - 1): n = n + 1
    Next
    ReDim Preserve out(0 To n - 1)
    ClassMarkings = out
End Function

' True when the app permits the marking: it is in the list and no higher than
' the app's classification. Unknown and empty markings are not permitted.
Public Function ClassPermits(ByVal label As String) As Boolean
    Dim r As Long
    r = ClassRank(label)
    ClassPermits = (r >= 0) And (r <= LevelRank(AppClassification()))
End Function

' Raise error 5 when ClassPermits is False, naming the app's classification.
Public Sub RequireClass(ByVal label As String)
    If Not ClassPermits(label) Then
        Err.Raise 5, , "The marking '" & label & "' is not permitted in a " & AppClassification() & " app"
    End If
End Sub

' Choose the banner's marking from the permitted ones (by default it is the
' app's classification). Raises error 5 for a marking ClassPermits refuses.
Public Sub SetBannerMarking(ByVal label As String)
    RequireClass label
    SaveSetting APP_NAME, CLASS_SECTION, "Marking", UCase$(Trim$(label))
    RefreshClassBanners
End Sub

' The marking the banner shows now: the one chosen if the app still permits it,
' else the app's classification.
Public Function BannerMarking() As String
    Dim s As String
    On Error Resume Next
    s = UCase$(GetSetting(APP_NAME, CLASS_SECTION, "Marking", ""))
    On Error GoTo 0
    If Not ClassPermits(s) Then s = AppClassification()
    BannerMarking = s
End Function

' Ask the user to pick the app's classification from a menu under anchor (a
' button). Returns the level picked, or "" when nothing was picked. The caller
' decides who may use it.
Public Function ChooseClassification(anchor As Control) As String
    Dim lv As Variant, items() As Variant, i As Long, k As String
    lv = ClassLevels()
    ReDim items(0 To UBound(lv))
    For i = 0 To UBound(lv)
        items(i) = lv(i) & "|" & lv(i)
    Next
    k = KitMenu(anchor, items(0), items(1), items(2), items(3), items(4), items(5))
    If Len(k) > 0 Then SetAppClassification k
    ChooseClassification = k
End Function

' Banner colors for a level rank: the standard marking colors (green, purple,
' blue, red, orange, yellow). Returns Array(back, fore).
Public Function ClassColors(ByVal rank As Long) As Variant
    Select Case rank
        Case 0: ClassColors = Array(RGB(0, 122, 51), RGB(255, 255, 255))
        Case 1: ClassColors = Array(RGB(80, 45, 130), RGB(255, 255, 255))
        Case 2: ClassColors = Array(RGB(0, 51, 160), RGB(255, 255, 255))
        Case 3: ClassColors = Array(RGB(200, 16, 46), RGB(255, 255, 255))
        Case 4: ClassColors = Array(RGB(255, 103, 31), RGB(0, 0, 0))
        Case Else: ClassColors = Array(RGB(255, 205, 0), RGB(0, 0, 0))
    End Select
End Function

' Build a banner strip on the form under construction: a colored panel nm
' (Tag "classbanner") with the marking centered in bold in nm_text. Put one at
' the top and one at the bottom, as paper documents do. h defaults to
' CLASS_BANNER_H. It starts as UNCLASSIFIED; ShowClassBanner sets the real
' marking when the form loads. Anchored across the form's width; pass
' atBottom:=True for a banner that stays at the bottom edge. Returns the panel.
Public Function ClassBanner(ByVal nm As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, _
        Optional ByVal h As Long = CLASS_BANNER_H, Optional ByVal atBottom As Boolean = False) As Control
    Dim pnl As Control, t As Control, k As Variant
    k = ClassColors(0)
    Set pnl = Box(nm, x, y, w, h, CLng(k(0)))
    pnl.Tag = BANNER_TAG
    Anchor pnl, 2, IIf(atBottom, 1, 0)
    Set t = Lbl(nm & "_text", "UNCLASSIFIED", x, y, w, h, 10, True, CLng(k(1)))
    t.TextAlign = 2
    Anchor t, 2, IIf(atBottom, 1, 0)
    Set ClassBanner = pnl
End Function

' Show the current marking on every banner of frm (the controls ClassBanner
' built). A form without banners is left alone.
Public Sub ShowClassBanner(frm As Object)
    Dim c As Control, k As Variant, m As String
    m = BannerMarking()
    k = ClassColors(ClassRank(m))
    On Error GoTo Fail
    For Each c In frm.Controls
        If c.Tag = BANNER_TAG Then
            c.BackColor = CLng(k(0))
            With frm.Controls(c.Name & "_text")
                .Caption = Replace(m, "&", "&&")    ' a label takes & as an access key
                .ForeColor = CLng(k(1))
            End With
            c.ControlTipText = "Classification: " & m
        End If
    Next
    Exit Sub
Fail:
    LogError "modClassification.ShowClassBanner", "form " & frm.Name
End Sub

' Repaint the banners of every open form after the level or marking changed.
Private Sub RefreshClassBanners()
    Dim f As Object
    On Error Resume Next
    For Each f In Forms
        ShowClassBanner f
    Next
End Sub

' Rank of a level name, or -1.
Private Function LevelRank(ByVal level As String) As Long
    Dim i As Long, lv As Variant
    lv = ClassLevels()
    LevelRank = -1
    For i = 0 To UBound(lv)
        If lv(i) = level Then LevelRank = i: Exit Function
    Next
End Function
