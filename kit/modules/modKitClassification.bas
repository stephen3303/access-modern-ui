Option Compare Database
Option Explicit

' Access UI kit - the classification banner and the markings the app permits.
'
' The app has one selected classification, the highest it handles
' (AppClassification, kept in the front end's database property
' KitClassification; SetAppClassification changes it). A banner bar across the
' top and bottom of the main form shows it in the standard banner color, and
' it limits which classification labels the app accepts: a label is permitted
' when its level is no higher than the app's. An app classified SECRET takes
' UNCLASSIFIED, CUI, CONFIDENTIAL, SECRET and SECRET//NOFORN, and refuses
' TOP SECRET. Until a classification is chosen the banner reads
' "CLASSIFICATION NOT SET" in gray, and only UNCLASSIFIED labels are permitted.
'
' The list, from lowest to highest level. Each marking has the color of its
' banner (CLR_CLASS_* are FIXED colors: never repainted by a palette):
'   UNCLASSIFIED                 green, white text
'   CUI                          purple, white text
'   CONFIDENTIAL                 blue, white text
'   SECRET, SECRET//NOFORN       red, white text
'   TOP SECRET, TOP SECRET//NOFORN       orange, dark text
'   TOP SECRET//SCI, TOP SECRET//SCI//NOFORN   yellow, dark text
' A dissemination control such as //NOFORN does not change the level.
'
' Using it:
'   Build, in the shell form:   ClassBanner "banTop", 0, w, 0
'                               ClassBanner "banBottom", h - px(20), w, 1
'   Form_Load of that form:     ClassApplyBanners Me
'   A button that changes it:   ChooseClassification Me.btnClassification
'   A field that holds a label: Me.cboMarking.RowSource = PermittedClassLabels()
'                               If Not CheckClassLabel(label, problem) Then ...
' The banner shows the app's classification only. Marking each record,
' report page or export with its own label, and checking existing records
' when the classification is lowered, are the app's job (ClassLabelPermitted).
'
' The app provides APP_NAME, HandleError, LogError and RaiseAgain (modErrors).

Public Const CLASS_COUNT As Long = 9
Public Const CLASS_BANNER_TAG As String = "classbanner"
Public Const CLASS_NOT_SET_TEXT As String = "CLASSIFICATION NOT SET"
Private Const CLASS_PROP As String = "KitClassification"

' The marking at position i (1 to CLASS_COUNT), lowest level first.
Public Function ClassLabel(ByVal i As Long) As String
    Select Case i
        Case 1: ClassLabel = "UNCLASSIFIED"
        Case 2: ClassLabel = "CUI"
        Case 3: ClassLabel = "CONFIDENTIAL"
        Case 4: ClassLabel = "SECRET"
        Case 5: ClassLabel = "SECRET//NOFORN"
        Case 6: ClassLabel = "TOP SECRET"
        Case 7: ClassLabel = "TOP SECRET//NOFORN"
        Case 8: ClassLabel = "TOP SECRET//SCI"
        Case 9: ClassLabel = "TOP SECRET//SCI//NOFORN"
    End Select
End Function

' A marking in its standard form (upper case, no spaces around //), or ""
' when it isn't one of the list.
Public Function ClassNormalize(ByVal label As String) As String
    Dim i As Long
    label = UCase$(Trim$(label))
    Do While InStr(label, "  ") > 0
        label = Replace(label, "  ", " ")
    Loop
    label = Replace(Replace(label, " //", "//"), "// ", "//")
    For i = 1 To CLASS_COUNT
        If ClassLabel(i) = label Then ClassNormalize = label: Exit Function
    Next
End Function

' The level of a marking: 0 UNCLASSIFIED, 1 CUI, 2 CONFIDENTIAL, 3 SECRET,
' 4 TOP SECRET, 5 TOP SECRET//SCI. -1 for anything not in the list.
Public Function ClassRank(ByVal label As String) As Long
    ClassRank = -1
    Select Case ClassNormalize(label)
        Case "UNCLASSIFIED": ClassRank = 0
        Case "CUI": ClassRank = 1
        Case "CONFIDENTIAL": ClassRank = 2
        Case "SECRET", "SECRET//NOFORN": ClassRank = 3
        Case "TOP SECRET", "TOP SECRET//NOFORN": ClassRank = 4
        Case "TOP SECRET//SCI", "TOP SECRET//SCI//NOFORN": ClassRank = 5
    End Select
End Function

' The banner color of a marking; gray for anything else (not set).
Public Function ClassBack(ByVal label As String) As Long
    Select Case ClassRank(label)
        Case 0: ClassBack = CLR_CLASS_UNCLASSIFIED
        Case 1: ClassBack = CLR_CLASS_CUI
        Case 2: ClassBack = CLR_CLASS_CONFIDENTIAL
        Case 3: ClassBack = CLR_CLASS_SECRET
        Case 4: ClassBack = CLR_CLASS_TOP_SECRET
        Case 5: ClassBack = CLR_CLASS_SCI
        Case Else: ClassBack = CLR_BAND_NE
    End Select
End Function

' The text color that reads on the marking's banner: dark on orange and
' yellow, white on the rest.
Public Function ClassFore(ByVal label As String) As Long
    Select Case ClassRank(label)
        Case 4, 5: ClassFore = CLR_CLASS_INK
        Case Else: ClassFore = CLR_WHITE
    End Select
End Function

' ------------------------------------------------------ the app's classification

' The app's classification as one of the list's markings, or "" when none has
' been chosen (or the saved value is not on the list).
Public Function AppClassification() As String
    Dim v As String
    On Error Resume Next
    v = CStr(CurrentDb.Properties(CLASS_PROP))
    On Error GoTo 0
    AppClassification = ClassNormalize(v)
End Function

' Choose the app's classification, saved in the database. Raises error 5 for
' a marking that isn't on the list. Open forms are not repainted: call
' ClassApplyAll.
Public Sub SetAppClassification(ByVal label As String)
    Dim db As DAO.Database, prop As DAO.Property, n As Long, d As String, std As String
    std = ClassNormalize(label)
    If Len(std) = 0 Then Err.Raise 5, "modKitClassification.SetAppClassification", "Not a classification marking: " & label
    Set db = CurrentDb
    On Error Resume Next
    db.Properties(CLASS_PROP) = std
    If Err.Number = 3270 Then         ' property not found: create it
        Err.Clear
        Set prop = db.CreateProperty(CLASS_PROP, dbText, std)
        db.Properties.Append prop
    End If
    If Err.Number <> 0 Then
        n = Err.Number: d = Err.Description
        On Error GoTo 0
        Err.Raise n, "modKitClassification.SetAppClassification", "Could not save the classification: " & d
    End If
    On Error GoTo 0
End Sub

' True when label is on the list and no higher than the app's classification.
' With no classification chosen only UNCLASSIFIED is permitted.
Public Function ClassLabelPermitted(ByVal label As String) As Boolean
    Dim r As Long, limit As Long
    r = ClassRank(label)
    If r < 0 Then Exit Function
    If Len(AppClassification()) = 0 Then limit = 0 Else limit = ClassRank(AppClassification())
    ClassLabelPermitted = (r <= limit)
End Function

' The permitted markings as a value list for a combo box, lowest first:
' "UNCLASSIFIED;CUI;CONFIDENTIAL;SECRET;SECRET//NOFORN" for a SECRET app.
Public Function PermittedClassLabels() As String
    Dim i As Long
    For i = 1 To CLASS_COUNT
        If ClassLabelPermitted(ClassLabel(i)) Then
            PermittedClassLabels = PermittedClassLabels & IIf(Len(PermittedClassLabels) > 0, ";", "") & ClassLabel(i)
        End If
    Next
End Function

' Check a label a user typed or picked. True when it is permitted; otherwise
' False with the reason in problem, ready for an inline message. Matches as
' ClassNormalize does ("secret // noforn" is SECRET//NOFORN).
Public Function CheckClassLabel(ByVal label As String, ByRef problem As String) As Boolean
    Dim std As String
    problem = ""
    std = ClassNormalize(label)
    If Len(std) = 0 Then problem = "Choose a classification marking from the list.": Exit Function
    If ClassLabelPermitted(std) Then CheckClassLabel = True: Exit Function
    If Len(AppClassification()) = 0 Then
        problem = APP_NAME & "'s classification has not been set, so only UNCLASSIFIED can be used."
    Else
        problem = std & " is above this app's classification (" & AppClassification() & ")."
    End If
End Function

' ------------------------------------------------------------------ the banner

' A banner bar the width of the form, tagged for ClassApplyBanners. Built in
' the app's classification now; ClassApplyBanners keeps it current. Use one
' at the top (y = 0, vAnchor 0) and one at the bottom (the form's height
' less 20px, vAnchor 1) of the main form. 20px high. Forms only.
Public Function ClassBanner(ByVal nm As String, ByVal y As Long, ByVal w As Long, _
        Optional ByVal vAnchor As Integer = 0, Optional ByVal sec As Integer = acDetail) As Control
    Dim c As Control, std As String
    std = AppClassification()
    Set c = Lbl(nm, IIf(Len(std) > 0, std, CLASS_NOT_SET_TEXT), 0, y, w, px(20), 8, True, ClassFore(std), sec)
    c.BackStyle = 1
    c.BackColor = ClassBack(std)
    c.TextAlign = 2
    c.FontWeight = 700
    c.Tag = CLASS_BANNER_TAG
    Anchor c, 2, vAnchor
    Set ClassBanner = c
End Function

' Show the app's classification on every banner of frm. Call it from the
' form's Load.
Public Sub ClassApplyBanners(frm As Form)
    Dim c As Control, std As String, t As String
    std = AppClassification()
    For Each c In frm.Controls
        t = ""
        On Error Resume Next
        t = c.Tag
        On Error GoTo 0
        If t = CLASS_BANNER_TAG Then
            c.Caption = IIf(Len(std) > 0, std, CLASS_NOT_SET_TEXT)
            c.BackColor = Th(ClassBack(std))
            c.ForeColor = Th(ClassFore(std))
        End If
    Next
End Sub

' Repaint the banners on every open form (after the classification changes).
Public Sub ClassApplyAll()
    Dim frm As Form
    For Each frm In Forms
        ClassApplyBanners frm
    Next
End Sub

' A menu under anchor to choose the app's classification, applied to every
' open form and saved. Choosing a lower level than the current one asks first,
' since records already marked above it stay as they are. True when the
' classification was changed. Call it from a button the app shows only to
' the people who may change it.
Public Function ChooseClassification(anchor As Control) As Boolean
    Dim items(0 To CLASS_COUNT - 1) As Variant, i As Long, pick As String, cur As String, label As String
    cur = AppClassification()
    For i = 1 To CLASS_COUNT
        items(i - 1) = CStr(i) & "|" & ClassLabel(i) & IIf(ClassLabel(i) = cur, "|E73E", "")
    Next
    pick = KitMenu(anchor, items(0), items(1), items(2), items(3), items(4), items(5), items(6), items(7), items(8))
    If Len(pick) = 0 Then Exit Function
    label = ClassLabel(CLng(pick))
    If label = cur Then Exit Function
    If Len(cur) > 0 Then
        If ClassRank(label) < ClassRank(cur) Then
            If KitMsgBox("Lower this app's classification to " & label & "?" & vbCrLf & vbCrLf & _
                    "Records already marked above " & label & " are not changed. Check them before continuing.", _
                    vbYesNo + vbExclamation + vbDefaultButton2, "Change classification") <> vbYes Then Exit Function
        End If
    End If
    SetAppClassification label
    ClassApplyAll
    KitToast "Classification set to " & label, "info"
    ChooseClassification = True
End Function
