Option Compare Database
Option Explicit

' Access UI kit - the classification banner and the classification limit.
'
' The app has one selected classification (AppClassification), the highest
' marking it may hold. The banner builder (Banner, modUIKit) shows it in the
' standard marking color at the top and bottom of a form, and every place the
' app lets a user pick a marking offers only the labels the selection permits:
'
'   cboMarking.RowSourceType = "Value List"
'   cboMarking.RowSource = PermittedClassificationList()
'   problem = ClassificationProblem(Me.cboMarking)     ' "" when it is allowed
'
' A label is permitted when it is a known marking at or below the selected
' classification. Markings are ranked by level, lowest first:
'   0 UNCLASSIFIED
'   1 UNCLASSIFIED//FOUO, CUI
'   2 CONFIDENTIAL (and //NOFORN)
'   3 SECRET (and //NOFORN, //REL TO USA, FVEY)
'   4 TOP SECRET (and //NOFORN)
'   5 TOP SECRET//SCI (and //NOFORN)
' so with SECRET selected, UNCLASSIFIED, CUI and every SECRET label are
' permitted and TOP SECRET is not. A label that isn't in the list is never
' permitted. To use other markings, edit Markings() and the color choices in
' ClassificationColor.
'
' The selection is saved per Windows user under APP_NAME (the registry), and
' is UNCLASSIFIED until someone chooses. This guards the app's own screens and
' rules; it does not protect the data file, which needs the folder's own
' permissions. An app that sets the selection centrally calls
' SetAppClassification at startup. KitChooseClassification opens a menu under
' a button for a user to choose; show that button to administrators only
' (SignedInIsAdmin, modPersonnel). Lowering the selection does not change
' records already marked higher: the app checks them with
' ClassificationPermitted.
'
' Banner labels carry the Tag "classbanner". ThemeForm refreshes them as each
' form opens, and SetAppClassification refreshes every open form, so a change
' shows at once. The colors are fixed (CLR_CLASS_*, make_themes.py): palettes
' never repaint them.
'
' The app provides APP_NAME and LogError (modErrors).

Public Const KIT_CLASS_DEFAULT As String = "UNCLASSIFIED"
Public Const CLASS_BANNER_TAG As String = "classbanner"
Private Const SETTINGS_KEY As String = "Classification"

' Every marking, "label|level", lowest level first.
Private Function Markings() As Variant
    Markings = Array( _
        "UNCLASSIFIED|0", _
        "UNCLASSIFIED//FOUO|1", _
        "CUI|1", _
        "CONFIDENTIAL|2", _
        "CONFIDENTIAL//NOFORN|2", _
        "SECRET|3", _
        "SECRET//NOFORN|3", _
        "SECRET//REL TO USA, FVEY|3", _
        "TOP SECRET|4", _
        "TOP SECRET//NOFORN|4", _
        "TOP SECRET//SCI|5", _
        "TOP SECRET//SCI//NOFORN|5")
End Function

' ---------------------------------------------------------------------------
' The marking list
' ---------------------------------------------------------------------------

' Every known marking, lowest first (an array of Strings).
Public Function ClassificationLabels() As Variant
    Dim m As Variant, out() As String, i As Long
    m = Markings()
    ReDim out(0 To UBound(m))
    For i = 0 To UBound(m)
        out(i) = Split(m(i), "|")(0)
    Next
    ClassificationLabels = out
End Function

' The known marking that text names, or "" when it names none. Ignores case
' and extra spaces, and accepts the short forms U, FOUO, C, S, TS and
' TS//SCI and "UNCLASSIFIED//FOR OFFICIAL USE ONLY".
Public Function NormalizeClassification(ByVal text As String) As String
    Dim s As String, m As Variant
    s = UCase$(Trim$(text))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Replace(Replace(s, " //", "//"), "// ", "//")
    Select Case s
        Case "U": s = "UNCLASSIFIED"
        Case "FOUO", "FOR OFFICIAL USE ONLY", "UNCLASSIFIED//FOR OFFICIAL USE ONLY": s = "UNCLASSIFIED//FOUO"
        Case "C": s = "CONFIDENTIAL"
        Case "S": s = "SECRET"
        Case "TS": s = "TOP SECRET"
        Case "TS//SCI": s = "TOP SECRET//SCI"
    End Select
    For Each m In Markings()
        If Split(m, "|")(0) = s Then NormalizeClassification = s: Exit Function
    Next
End Function

' The level (0 to 5) of a marking, or -1 when text names none.
Public Function ClassificationLevel(ByVal text As String) As Long
    Dim s As String, m As Variant, p() As String
    ClassificationLevel = -1
    s = NormalizeClassification(text)
    If Len(s) = 0 Then Exit Function
    For Each m In Markings()
        p = Split(m, "|")
        If p(0) = s Then ClassificationLevel = CLng(p(1)): Exit Function
    Next
End Function

' ---------------------------------------------------------------------------
' The app's selected classification and what it permits
' ---------------------------------------------------------------------------

' The selected classification: the saved choice, or UNCLASSIFIED when none
' is saved or the saved text is no longer a known marking.
Public Function AppClassification() As String
    Dim s As String
    On Error Resume Next
    s = NormalizeClassification(GetSetting(APP_NAME, SETTINGS_KEY, "Selected", KIT_CLASS_DEFAULT))
    On Error GoTo 0
    If Len(s) = 0 Then s = KIT_CLASS_DEFAULT
    AppClassification = s
End Function

' Select the app's classification and refresh every open banner. Raises error
' 5 when label isn't a known marking.
Public Sub SetAppClassification(ByVal label As String)
    Dim s As String
    s = NormalizeClassification(label)
    If Len(s) = 0 Then Err.Raise 5, , "Not a known classification marking: " & label
    SaveSetting APP_NAME, SETTINGS_KEY, "Selected", s
    UpdateClassificationBanners
End Sub

' True when label is a known marking at or below the selected classification.
Public Function ClassificationPermitted(ByVal label As String) As Boolean
    Dim lvl As Long
    lvl = ClassificationLevel(label)
    If lvl < 0 Then Exit Function
    ClassificationPermitted = (lvl <= ClassificationLevel(AppClassification()))
End Function

' The labels the selection permits, lowest first (an array of Strings).
Public Function PermittedClassifications() As Variant
    Dim all As Variant, out() As String, i As Long, n As Long
    all = ClassificationLabels()
    ReDim out(0 To UBound(all))
    For i = 0 To UBound(all)
        If ClassificationPermitted(CStr(all(i))) Then out(n) = all(i): n = n + 1
    Next
    ReDim Preserve out(0 To n - 1)             ' never empty: the lowest marking is always permitted
    PermittedClassifications = out
End Function

' The permitted labels as a combo box's value list: "A";"B";...
Public Function PermittedClassificationList() As String
    Dim labels As Variant, i As Long, s As String
    labels = PermittedClassifications()
    For i = 0 To UBound(labels)
        If i > 0 Then s = s & ";"
        s = s & """" & labels(i) & """"
    Next
    PermittedClassificationList = s
End Function

' Why label can't be used, in words for the user, or "" when it can.
Public Function ClassificationProblem(ByVal label As String) As String
    If Len(Trim$(label)) = 0 Then
        ClassificationProblem = "Choose a classification."
    ElseIf ClassificationLevel(label) < 0 Then
        ClassificationProblem = Trim$(label) & " isn't a classification marking this app knows."
    ElseIf Not ClassificationPermitted(label) Then
        ClassificationProblem = NormalizeClassification(label) & " is above this app's classification (" & AppClassification() & ")."
    End If
End Function

' ---------------------------------------------------------------------------
' Banner colors and refresh
' ---------------------------------------------------------------------------

' The banner color for a marking (gray for an unknown one).
Public Function ClassificationColor(ByVal label As String) As Long
    Select Case ClassificationLevel(label)
        Case 0: ClassificationColor = CLR_CLASS_UNCLASSIFIED
        Case 1: ClassificationColor = CLR_CLASS_CUI
        Case 2: ClassificationColor = CLR_CLASS_CONFIDENTIAL
        Case 3: ClassificationColor = CLR_CLASS_SECRET
        Case 4: ClassificationColor = CLR_CLASS_TOP_SECRET
        Case 5: ClassificationColor = CLR_CLASS_SCI
        Case Else: ClassificationColor = CLR_BAND_NE
    End Select
End Function

' The text color that reads on ClassificationColor(label): black on the orange
' and yellow banners, white on the rest.
Public Function ClassificationTextColor(ByVal label As String) As Long
    Select Case ClassificationLevel(label)
        Case 4, 5: ClassificationTextColor = CLR_CLASS_TEXT_DARK
        Case Else: ClassificationTextColor = CLR_WHITE
    End Select
End Function

' Show the selected classification on every banner of frm (controls tagged
' "classbanner"), or of every open form when frm is left out.
Public Sub UpdateClassificationBanners(Optional frm As Object)
    Dim f As Object
    If frm Is Nothing Then
        For Each f In Forms
            ApplyBanners f
        Next
    Else
        ApplyBanners frm
    End If
End Sub

Private Sub ApplyBanners(frm As Object)
    Dim c As Control, label As String
    On Error GoTo Fail
    label = AppClassification()
    For Each c In frm.Controls
        If c.Tag = CLASS_BANNER_TAG Then
            c.Caption = label
            c.BackStyle = 1
            c.BackColor = ClassificationColor(label)
            c.ForeColor = ClassificationTextColor(label)
        End If
    Next
    Exit Sub
Fail:
    LogError "modKitClassification.ApplyBanners", "form " & frm.Name
End Sub

' ---------------------------------------------------------------------------
' Choosing
' ---------------------------------------------------------------------------

' Open a menu under anchor to choose the app's classification (a level at a
' time: UNCLASSIFIED, CUI, CONFIDENTIAL, SECRET, TOP SECRET, TOP SECRET//SCI;
' the //NOFORN and other forms are labels under those). True when the user
' picked a different one, which is then saved and shown on every banner.
' Anyone can call it: the app decides who sees the button.
Public Function KitChooseClassification(anchor As Control) As Boolean
    Dim pick As String
    On Error GoTo Fail
    pick = KitMenu(anchor, "UNCLASSIFIED|UNCLASSIFIED", "CUI|CUI", "CONFIDENTIAL|CONFIDENTIAL", _
        "SECRET|SECRET", "TOP SECRET|TOP SECRET", "TOP SECRET//SCI|TOP SECRET//SCI")
    If Len(pick) = 0 Then Exit Function
    If pick = AppClassification() Then Exit Function
    SetAppClassification pick
    KitChooseClassification = True
    Exit Function
Fail:
    HandleError "modKitClassification.KitChooseClassification", pick, "The classification couldn't be changed."
End Function
