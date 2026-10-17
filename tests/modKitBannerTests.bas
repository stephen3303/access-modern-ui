Option Compare Database
Option Explicit

' Tests for modKitBanner: the ranked marking list, the selected level limiting
' what is permitted, and the bars built on a form and painted from the
' selection. TestKit calls TestKitBanner and adds the count to its own.

Private n As Long

Private Sub Ok(ByVal valid As Boolean, ByVal message As String)
    If Not valid Then Err.Raise 5, , message
    n = n + 1
End Sub

' Error number raised by SetClassification label, else 0.
Private Function SetErr(ByVal label As String) As Long
    On Error Resume Next
    SetClassification label
    SetErr = Err.Number
End Function

' True when list (an array) holds label.
Private Function Holds(list As Variant, ByVal label As String) As Boolean
    Dim v As Variant
    For Each v In list
        If v = label Then Holds = True
    Next
End Function

Public Function TestKitBanner() As Long
    Dim frm As Form, lv As Long
    n = 0
    ' Levels and their colors.
    Ok ClassLevelName(CLASS_SECRET) = "SECRET" And ClassLevelName(CLASS_TS_SCI) = "TOP SECRET//SCI" And ClassLevelName(9) = "", _
        "Level names are wrong"
    Ok ClassLevelBack(CLASS_UNCLASSIFIED) = CLR_CLASS_UNCLASS And ClassLevelBack(CLASS_CUI) = CLR_CLASS_CUI _
        And ClassLevelBack(CLASS_CONFIDENTIAL) = CLR_CLASS_CONF And ClassLevelBack(CLASS_SECRET) = CLR_CLASS_SECRET _
        And ClassLevelBack(CLASS_TOP_SECRET) = CLR_CLASS_TS And ClassLevelBack(CLASS_TS_SCI) = CLR_CLASS_SCI _
        And ClassLevelBack(-1) = -1, "Level colors are wrong"
    Ok ClassLevelFore(CLASS_SECRET) = CLR_WHITE And ClassLevelFore(CLASS_TOP_SECRET) = CLR_CLASS_DARK _
        And ClassLevelFore(CLASS_TS_SCI) = CLR_CLASS_DARK, "Banner text colors are wrong"
    ' Markings and their ranks; case and spacing ignored; unknown is -1.
    Ok ClassLevelOf("UNCLASSIFIED") = 0 And ClassLevelOf("unclassified//fouo") = 0 And ClassLevelOf("CUI//SP-PRVCY") = 1 _
        And ClassLevelOf("CONFIDENTIAL//NOFORN") = 2 And ClassLevelOf(" Secret // Noforn ") = 3 _
        And ClassLevelOf("TOP SECRET") = 4 And ClassLevelOf("TOP SECRET//SCI//NOFORN") = 5, "Marking levels are wrong"
    Ok ClassLevelOf("") = -1 And ClassLevelOf("SECRET//BANANA") = -1 And ClassLevelOf("SECRE") = -1, "An unknown marking got a level"
    ' The list climbs in rank order and has every level.
    lv = 0
    Dim v As Variant, seen As String, prev As Long
    For Each v In ClassMarkings(CLASS_TS_SCI)
        Ok ClassLevelOf(CStr(v)) >= prev, "Markings are out of rank order at " & v
        prev = ClassLevelOf(CStr(v))
        If InStr(seen, "," & prev & ",") = 0 Then seen = seen & "," & prev & ",": lv = lv + 1
    Next
    Ok lv = 6, "The list is missing a level"
    ' The selection limits the list and the check.
    SetClassification "secret//noforn"
    Ok ClassSelected() = "SECRET//NOFORN" And ClassSelectedLevel() = CLASS_SECRET, "SetClassification didn't select the marking"
    Ok Holds(ClassMarkings(), "CUI//SP-PRVCY") And Holds(ClassMarkings(), "SECRET//REL TO USA, FVEY") _
        And Not Holds(ClassMarkings(), "TOP SECRET") And Not Holds(ClassMarkings(), "TOP SECRET//SCI"), "The permitted list isn't limited to SECRET"
    Ok Holds(ClassMarkings(CLASS_CUI), "CUI") And Not Holds(ClassMarkings(CLASS_CUI), "CONFIDENTIAL"), "ClassMarkings maxLevel isn't used"
    Ok Left$(ClassMarkingList(), 13) = "UNCLASSIFIED;" And InStr(ClassMarkingList(), "TOP SECRET") = 0, "ClassMarkingList is wrong"
    Ok ClassPermits("UNCLASSIFIED") And ClassPermits("cui") And ClassPermits("SECRET") And ClassPermits("Secret//REL TO USA, FVEY"), _
        "A lower or equal marking wasn't permitted"
    Ok Not ClassPermits("TOP SECRET") And Not ClassPermits("TOP SECRET//SCI") And Not ClassPermits("") And Not ClassPermits("SECRET//BANANA"), _
        "A higher or unknown marking was permitted"
    Ok ClassProblem("CUI") = "" And InStr(ClassProblem("TOP SECRET"), "higher than") > 0 And InStr(ClassProblem("SECRET//BANANA"), "not a marking") > 0, _
        "ClassProblem didn't say why"
    SetClassification "UNCLASSIFIED"
    Ok ClassPermits("UNCLASSIFIED//FOUO") And Not ClassPermits("CUI") And UBound(ClassMarkings()) = 1, "UNCLASSIFIED permits more than its own markings"
    SetClassification "TOP SECRET//SCI"
    Ok ClassPermits("TOP SECRET//SCI//NOFORN") And ClassPermits("TOP SECRET"), "The highest level doesn't permit every marking"
    Ok SetErr("SECRET//BANANA") = 5 And ClassSelected() = "TOP SECRET//SCI", "SetClassification took an unknown marking"
    ' The bars: built by BannerBuild, painted by ShellLoad from the selection.
    SetClassification "CUI"
    ShellBuild "Shell", "", Array("a|Alpha|subRing|E80F"), Empty, px(500), px(300), BannerHeight(), BannerHeight()
    BannerBuild
    Ok F.Controls("lblBannerTop").Top = 0 And F.Controls("lblBannerBottom").Top = F.Section(acDetail).Height - BannerHeight() _
        And F.Controls("boxSide").Top = BannerHeight() And F.Controls("sfPage").Top = BannerHeight(), "The bars don't frame the sidebar and page"
    Ok F.Controls("lblBannerTop").HorizontalAnchor = 2 And F.Controls("lblBannerBottom").VerticalAnchor = 1, "The bars don't stretch with the form"
    SaveForm "frmMain", "testShellBare"
    DoCmd.OpenForm "frmMain"
    Set frm = Forms("frmMain")
    Ok frm.lblBannerTop.Caption = "CUI" And frm.lblBannerTop.BackColor = CLR_CLASS_CUI And frm.lblBannerBottom.ForeColor = CLR_WHITE, _
        "ShellLoad didn't paint the bars"
    SetClassification "TOP SECRET"
    Ok frm.lblBannerTop.Caption = "TOP SECRET" And frm.lblBannerBottom.BackColor = CLR_CLASS_TS And frm.lblBannerBottom.ForeColor = CLR_CLASS_DARK, _
        "SetClassification didn't repaint the open bars"
    DoCmd.Close acForm, "frmMain", acSaveNo
    SetClassification "UNCLASSIFIED"
    TestKitBanner = n
End Function
