Option Compare Database
Option Explicit

' Access UI kit - classification banner and permitted classification labels.
'
' The app is run at one classification, the highest marking its data may
' carry. That choice (the "ceiling") limits which labels the app accepts: a
' SECRET app permits UNCLASSIFIED, CUI, CONFIDENTIAL and SECRET labels but
' refuses TOP SECRET ones. The banner shows the ceiling's marking in its
' conventional color (green UNCLASSIFIED, purple CUI, blue CONFIDENTIAL, red
' SECRET, orange TOP SECRET, yellow TS//SCI), fixed in every palette.
'
' Build: put a banner at the top and bottom of the shell (modUIKit.Banner):
'   Banner "lblBannerTop", 0, PAGE_W
'   Banner "lblBannerBottom", PAGE_H - px(20), PAGE_W, vAnchor:=1
' Run time: frmMain's Load calls
'   ApplyBanner Me
' and a record's label combo box is limited to the permitted markings:
'   FillClassificationCombo Me.cboMarking
' A save that carries a label checks it first:
'   If Not RequireClassification(Me.cboMarking) Then Cancel = True
' Choosing the app's classification (an admin setting) is
'   SetClassification "SECRET": ApplyBanner Forms("frmMain")
' The choice is saved in the database (property KitClassification). Without
' one the app is UNCLASSIFIED, the strictest ceiling, so an app that forgot to
' choose never permits more than it should.
'
' The marking list is ClassificationMarkings(): one place to edit, in rank
' order. A label is permitted when its rank is at or below the ceiling's.

Public Const CLASS_PROP As String = "KitClassification"
Public Const CLASS_DEFAULT As String = "UNCLASSIFIED"

Private m_class As String                       ' "" until read or set

' The levels an app can be run at, lowest rank first. These are the
' choices for SetClassification.
Public Function ClassificationLevels() As Variant
    ClassificationLevels = Array("UNCLASSIFIED", "CUI", "CONFIDENTIAL", "SECRET", "TOP SECRET", "TOP SECRET//SCI")
End Function

' Commonly used markings as "label|level": the label a record may carry and
' the level it belongs to. Edit this list to add or remove markings; each
' level must be one of ClassificationLevels().
Public Function ClassificationMarkings() As Variant
    ClassificationMarkings = Array( _
        "UNCLASSIFIED|UNCLASSIFIED", _
        "UNCLASSIFIED//FOR OFFICIAL USE ONLY|UNCLASSIFIED", _
        "CUI|CUI", _
        "CUI//SP-CTI|CUI", _
        "CUI//SP-PRVCY|CUI", _
        "CUI//SP-PROPIN|CUI", _
        "CUI//NOFORN|CUI", _
        "CONFIDENTIAL|CONFIDENTIAL", _
        "CONFIDENTIAL//NOFORN|CONFIDENTIAL", _
        "SECRET|SECRET", _
        "SECRET//NOFORN|SECRET", _
        "SECRET//REL TO USA, FVEY|SECRET", _
        "TOP SECRET|TOP SECRET", _
        "TOP SECRET//NOFORN|TOP SECRET", _
        "TOP SECRET//SCI|TOP SECRET//SCI")
End Function

' 1-based rank of a level (UNCLASSIFIED 1 ... TOP SECRET//SCI 6), 0 when
' level is not one of ClassificationLevels(). Case and spacing around the
' text don't matter.
Public Function ClassificationRank(ByVal level As String) As Long
    Dim lv As Variant, i As Long
    level = UCase$(Trim$(level))
    lv = ClassificationLevels()
    For i = 0 To UBound(lv)
        If lv(i) = level Then ClassificationRank = i + 1: Exit Function
    Next
End Function

' The level a label belongs to ("SECRET//NOFORN" -> "SECRET"). A label that
' is not in the list but starts with a known level and "//" or is a level
' itself takes that level; anything else is "" (not a known marking).
Public Function ClassificationOf(ByVal label As String) As String
    Dim m As Variant, p() As String, lv As Variant, i As Long, best As String
    label = UCase$(Trim$(label))
    If Len(label) = 0 Then Exit Function
    For Each m In ClassificationMarkings()
        p = Split(m, "|")
        If p(0) = label Then ClassificationOf = p(1): Exit Function
    Next
    lv = ClassificationLevels()                  ' the longest level the label starts with
    For i = 0 To UBound(lv)
        If label = lv(i) Or Left$(label, Len(lv(i)) + 2) = lv(i) & "//" Then
            If Len(lv(i)) > Len(best) Then best = lv(i)
        End If
    Next
    ClassificationOf = best
End Function

' The app's classification: the one SetClassification saved, else
' UNCLASSIFIED.
Public Function KitClassification() As String
    Dim db As DAO.Database, s As String
    If Len(m_class) = 0 Then
        m_class = CLASS_DEFAULT
        Set db = CurrentDb
        On Error Resume Next
        s = CStr(db.Properties(CLASS_PROP))
        On Error GoTo 0
        If ClassificationRank(s) > 0 Then m_class = UCase$(Trim$(s))
    End If
    KitClassification = m_class
End Function

' Choose the app's classification and save it in the database. Raises error 5
' for a level that is not in ClassificationLevels(). Call ApplyBanner on the
' open shell afterwards, and refill any label combo boxes.
Public Sub SetClassification(ByVal level As String)
    Dim db As DAO.Database, prop As DAO.Property, n As Long, d As String
    level = UCase$(Trim$(level))
    If ClassificationRank(level) = 0 Then Err.Raise 5, , "Unknown classification: " & level
    Set db = CurrentDb
    On Error Resume Next
    db.Properties(CLASS_PROP) = level
    If Err.Number = 3270 Then                    ' property not found: create it
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
    m_class = level
End Sub

' True when the app, at its classification, permits label: the label is a
' known marking (or a known level with a "//" suffix) at or below the
' ceiling. An empty or unknown label is not permitted.
Public Function ClassificationPermits(ByVal label As String) As Boolean
    Dim r As Long
    r = ClassificationRank(ClassificationOf(label))
    ClassificationPermits = (r > 0 And r <= ClassificationRank(KitClassification()))
End Function

' The labels the app permits, in list order (ClassificationMarkings() filtered
' by the ceiling), as an array of strings.
Public Function PermittedClassifications() As Variant
    Dim m As Variant, p() As String, out() As String, n As Long, ceil As Long
    ceil = ClassificationRank(KitClassification())
    ReDim out(0 To UBound(ClassificationMarkings()))
    For Each m In ClassificationMarkings()
        p = Split(m, "|")
        If ClassificationRank(p(1)) <= ceil Then out(n) = p(0): n = n + 1
    Next
    ReDim Preserve out(0 To n - 1)
    PermittedClassifications = out
End Function

' Limit a combo box to the permitted labels: a value list, no entry outside
' it. A value already in the box that is no longer permitted (the ceiling
' was lowered) is cleared so a refused label can't be saved by accident.
Public Sub FillClassificationCombo(cbo As Control)
    Dim labels As Variant
    labels = PermittedClassifications()
    cbo.RowSourceType = "Value List"
    cbo.RowSource = """" & Join(labels, """;""") & """"
    cbo.LimitToList = True
    If Not IsNull(cbo.Value) Then
        If Not ClassificationPermits(CStr(cbo.Value)) Then cbo.Value = Null
    End If
End Sub

' True when label is permitted. When it is not, tells the user why and returns
' False (call it from a save: If Not RequireClassification(x) Then Cancel = True).
' An empty label is refused too: a record must carry a marking.
Public Function RequireClassification(ByVal label As String) As Boolean
    If ClassificationPermits(label) Then
        RequireClassification = True
    ElseIf Len(Trim$(label)) = 0 Then
        KitMsgBox "Choose a classification label for this record.", vbExclamation, "Classification"
    Else
        KitMsgBox "This app is classified " & KitClassification() & ". It can't hold a " & label & " label.", vbExclamation, "Classification"
    End If
End Function

' The banner's fill for a level (a fixed color: the same in every palette).
Public Function ClassificationBack(ByVal level As String) As Long
    Select Case ClassificationRank(level)
        Case 2: ClassificationBack = CLR_CLASS_CUI
        Case 3: ClassificationBack = CLR_CLASS_CONFIDENTIAL
        Case 4: ClassificationBack = CLR_CLASS_SECRET
        Case 5: ClassificationBack = CLR_CLASS_TOP_SECRET
        Case 6: ClassificationBack = CLR_CLASS_SCI
        Case Else: ClassificationBack = CLR_CLASS_UNCLASSIFIED
    End Select
End Function

' The banner's text color for a level: dark on the orange and yellow
' banners, white on the others.
Public Function ClassificationFore(ByVal level As String) As Long
    If ClassificationRank(level) >= 5 Then
        ClassificationFore = CLR_CLASS_DARK_TEXT
    Else
        ClassificationFore = CLR_WHITE
    End If
End Function

' Show the app's classification on every banner of frm (the labels whose name
' starts with "lblBanner", built by modUIKit.Banner): its marking as the text
' and the marking's colors. Call it from the shell's Load and after
' SetClassification. A form without banners is left alone.
Public Sub ApplyBanner(frm As Form)
    Dim c As Control, level As String
    level = KitClassification()
    For Each c In frm.Controls
        If c.ControlType = acLabel Then
            If Left$(c.Name, 9) = "lblBanner" Then
                c.Caption = level
                c.BackColor = ClassificationBack(level)
                c.ForeColor = ClassificationFore(level)
            End If
        End If
    Next
End Sub
