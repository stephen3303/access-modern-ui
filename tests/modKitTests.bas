Option Compare Database
Option Explicit
Private checks As Long
Private Sub Check(ByVal valid As Boolean, ByVal message As String)
    If Not valid Then Err.Raise 5, , message
    checks = checks + 1
End Sub
Public Function TestKit(ByVal root As String) As String
    Dim grid As Form, host As Form, c As Control, cols As Variant, i As Long
    Dim rows As DAO.Recordset, ids As Object, height As Long
    Dim text As String, fitted As String, measured As Long
    Dim db As DAO.Database, badName As Long, badScale As Long, badView As Long
    ' A kit error message (HandleError) would sit unseen behind the hidden Access
    ' and stall the run while it reports success; g_Silent raises it instead, so
    ' the run fails and names it.
    g_Silent = True
    g_SrcRoot = root & "\kit": UseBuildPalette
    CurrentDb.Execute "CREATE TABLE test_rows (id LONG PRIMARY KEY, category TEXT(60), description TEXT(255))"
    For i = 1 To 23
        CurrentDb.Execute "INSERT INTO test_rows VALUES (" & i & ",'Long category','Description " & i & "')"
    Next
    cols = Array("category|Category|160|c", "description|Description|300|ts")
    BuildGrid "subTest", cols, px(600)
    BuildGrid "subTestName", Array("name|Name|200|s"), px(300)
    NewForm "", 0, px(200), px(60)
    Txt "ringC", "", 0, 0, px(100), px(30)
    Btn "ringSave", "Save", px(110), 0, px(80), px(30)
    SaveForm "subRing"
    BuildToast
    BuildMenuDialog
    BuildPopoverDialog
    NewForm "test_rows", 0, px(300), px(40)
    Ev F, "BeforeUpdate"
    Txt "txtDesc", "description", 0, 0, px(280), px(30)
    SaveForm "subRefuse", "testRefuse"
    NewForm "", 0, px(700), px(700)
    GridHeaders "head", cols, px(24), px(60)
    GridHeaders "sortHead", cols, px(24), px(40), sortSf:="sfRows"
    GridHeaders "fillHead", cols, px(24), px(770), w:=px(600)
    GridHeaders "fitHead", cols, px(24), px(790), w:=px(600), fit:=True
    SubF "sfRows", "Form.subTest", px(24), px(82), px(600), px(180)
    SubF "sfReuse", "Form.subTest", px(24), px(320), px(600), px(180)
    SubF "sfName", "Form.subTestName", px(24), px(700), px(300), px(50)
    SubF "sfRefuse", "Form.subRefuse", px(340), px(700), px(300), px(40)
    GridFooter "sfRows", "pager", "id"
    Lbl "filter0", "All (23)", 0, 0, px(100), px(30)
    Lbl "filter1", "Retirement (2)", 0, 0, px(100), px(30)
    For i = 0 To 2
        Img "stepMarker" & i, 0, px(530), 420, 420
        Lbl "stepNumber" & i, CStr(i + 1), 0, px(530), 420, 420
        Lbl "stepLabel" & i, "Step", 0, px(560), 420, 300
        Lbl "stepDate" & i, "", 0, px(580), 420, 300
        If i < 2 Then Box "stepConnector" & i, 0, px(540), 1000, 15, CLR_BORDER
    Next
    Set c = Lbl("pillPaint", "Open", 0, px(620), px(80), px(20)): PillLabel c
    Set c = Txt("pillBound", "=""Open""", px(100), px(620), px(80), px(20))
    CF c, "True", CLR_PRIMARY, CLR_WHITE: RoundCF c
    Txt "ringA", "", px(200), px(620), px(80), px(30)
    Cbo "ringB", "", "One;Two", px(290), px(620), px(80), px(30)
    SubF "sfRing", "Form.subRing", px(400), px(620), px(200), px(60)
    Set c = CreateControl(F.Name, acTextBox, acDetail, "", "", px(610), px(620), px(60), px(30))
    c.Name = "plainTxt"                    ' takes the focus with no ring hooks of its own
    Txt "reqA", "", 0, px(660), px(80), px(30)
    Txt "reqB", "", px(100), px(660), px(80), px(30)
    Toggle "tglTest", "", px(220), px(663), onChange:="ToggleChanged"
    Segmented "segTest", "", px(700), px(660), px(240), "day|Day;week|Week;month|Month", onChange:="SegmentChanged"
    Alert "alTest", px(700), px(720), px(320), px(64), "warning", "Heads up", "Check the due date.", True
    Breadcrumb "bc", px(700), px(800), px(320), 5, "CrumbPicked", "home|Home", "items|Items", "i42|Pump overhaul"
    Stepper "st", px(700), px(840), px(360), 2, "StepPicked", "Details", "Review", "Done"
    Accordion "acc", px(1100), px(40), px(300), False, "General|Form.subRing|60|open", "Advanced|Form.subRing|60", "Notes||40"
    Accordion "acs", px(1450), px(40), px(300), True, "One||40|open", "Two||40"
    Avatar "avTest", "=""Ada Lovelace""", px(1100), px(500)
    StatTile "kpi", "Open items", "=12", px(1100), px(560), px(240), , "=-3.5", "vs last week", "E8F4", "#,##0", False
    EmptyState "emp", px(1100), px(700), px(300), "E721", "No items match", "Try another search.", "Clear search", "", True
    SearchBox "srch", px(1450), px(560), px(240), "SearchTyped"
    Rating "rt", "", px(1450), px(610), onChange:="RatingChanged"
    Slider "sl", "", px(1450), px(660), px(300), 0, 100, 10, "SliderChanged", True
    SecretBox "secret", "", px(400), px(700), px(160), px(30)     ' clear of the ring, required and toggle fields
    SaveForm "frmTest", "testHost"
    DoCmd.OpenForm "frmTest"
    Set host = Forms("frmTest"): Set grid = host.sfRows.Form
    ' Captions are as wide as their cells; a filling last column's reaches the table's edge.
    Check host.fillHead0.Width = grid.txtCol0.Width And host.fillHead0.Left = host.sfRows.Left + grid.txtCol0.Left, "Caption is not over its cells"
    ' Designed width (the Tag; an anchored control's Width follows the window): 600 - 8 - 160 - 8 - 8.
    Check Val(Split(host.fillHead1.Tag, "|")(2)) = px(416) And host.fillHead1.HorizontalAnchor = 2, "Last caption does not reach the table's edge"
    Check host.fitHead1.Width = px(292) And host.fitHead1.HorizontalAnchor = 0, "Fit table's last caption is not its column's width"
    Check FitSingleLineText(Null, px(100)) = "", "Null presentation text differs"
    Check FitSingleLineText("A" & vbCrLf & "B" & vbLf & "C", px(100)) = "A B C", "Presentation text retains line breaks"
    measured = MeasureText("Exact fit", 9, 400)
    Check FitSingleLineText("Exact fit", measured) = "Exact fit", "Exact-fit text was truncated"
    Check FitSingleLineText("Text", 0) = "", "Zero-width text overflows"
    Check IsNull(GridMark(Null, "a")), "Null search cell differs"
    Check GridMark("a<b & c", "") = "a&lt;b &amp; c", "Unsearched cell not escaped"
    Check GridMark("a   b", "") = "a &nbsp; b", "Cell spaces collapsed"
    Check GridMark(" a", "") = "&nbsp;a", "Leading space collapsed"
    Check GridMark("Pump overhaul", "pump") Like "*</font> overhaul", "Space after a match changed"
    text = GridMark("Pump pump", "PUMP")
    Check UBound(Split(text, "</font>")) = 2 And InStr(text, ">Pump</font> <") > 0, "Search matches not marked"
    measured = MeasureText(ChrW(&H2026), 9, 400)
    Check FitSingleLineText("Long text", measured - 1) = "", "Tiny cell emits overflowing ellipsis"
    text = "Wide WWW and narrow iii characters in a long description"
    fitted = FitSingleLineText(text, px(90))
    Check Right$(fitted, 1) = ChrW(&H2026), "Clipped text lacks ellipsis"
    Check MeasureText(fitted, 9, 400) <= px(90), "Measured presentation text overflows"
    Check Len(fitted) < Len(text), "Clipped text was not shortened"
    Check text = "Wide WWW and narrow iii characters in a long description", "Presentation fitting changed original text"
    Set c = host.head1
    c.Width = px(85)
    FitHeaderText c, "A long description heading", 9, 8
    Check c.ControlTipText = "A long description heading", "Header does not expose full caption"
    Check c.FontSize = 8, "Narrow header did not reduce font to minimum"
    Check Right$(c.Caption, 1) = ChrW(&H2026), "Narrow header lacks ellipsis"
    Check MeasureText(c.Caption, c.FontSize, IIf(c.FontBold, 700, 400)) <= c.Width - c.LeftMargin - c.RightMargin, "Header text exceeds its inner width"
    c.Width = px(400)
    FitHeaderText c, "A long description heading", 9, 8
    Check c.Caption = "A long description heading" And c.FontSize = 9, "Widened header did not restore caption and preferred font"
    FitHeaderText c, String$(300, "W"), 9, 8
    Check Len(c.ControlTipText) = 255, "Native header tip exceeds Access limit"
    FitHeaderText c, "Description", 9, 8
    LayoutFilterRow host, Array("filter0", "filter1"), px(24), px(10), px(600)
    Check host.filter0.Width = host.filter1.Width, "Filter widths differ"
    Check host.filter0.Top = host.filter1.Top, "Filters do not share one row"
    ' Spacing: unset is the guide's values; a choice scales gaps, is saved and can be undone.
    Check SpacingScale() = 1 And Sp(SP_CARD_GAP) = px(12), "Default spacing changed the guide's gaps"
    Check host.filter1.Left = host.filter0.Left + host.filter0.Width + px(4), "Default filter gap is not 4px"
    SetSpacing "compact"
    Check Sp(SP_CARD_GAP) = px(9) And Sp(SP_PAGE) = px(24), "Compact spacing did not scale gaps"
    SetSpacing "Comfortable"
    Check Sp(SP_ITEM_GAP) = px(5) And Sp(SP_ROW_GAP) = px(10), "Comfortable spacing did not scale gaps"
    LayoutFilterRow host, Array("filter0", "filter1"), px(24), px(10), px(600)
    Check host.filter1.Left = host.filter0.Left + host.filter0.Width + px(5), "Filter row ignores the chosen spacing"
    Set db = CurrentDb
    Check CDbl(db.Properties("KitSpacing")) = 1.25, "Spacing choice was not saved"
    On Error Resume Next
    SetSpacing "roomy": badName = Err.Number: Err.Clear
    SetSpacing 3: badScale = Err.Number: Err.Clear
    On Error GoTo 0
    Check badName = 5, "Unknown spacing name accepted"
    Check badScale = 5, "Out-of-range spacing accepted"
    Check SpacingScale() = 1.25, "Rejected spacing changed the scale"
    SetSpacing "default"
    LayoutFilterRow host, Array("filter0", "filter1"), px(24), px(10), px(600)
    LayoutNumberedTimeline host, "step", 3, px(24), px(500)
    LayoutNumberedTimeline host, "step", 3, px(24), px(650)
    Check host.stepMarker0.Width = 420 And host.stepMarker0.Height = 420, "Timeline circle expanded"
    Check host.stepNumber0.Width = 420, "Timeline number expanded"
    Check host.stepConnector0.Width = px(650) \ 3 - 480, "Timeline connector did not adjust"
    host.pillPaint_bg.Visible = True
    SetVisible host.pillPaint, False
    Check Not host.pillPaint_bg.Visible, "Hidden painted pill left its image"
    SetVisible host.pillPaint, True
    Check host.pillPaint.Visible And Not host.pillPaint_bg.Visible, "Painted pill image shown without a paint"
    SetVisible host.pillBound, False
    Check Not host.pillBound.Visible And Not host.pillBound_bg.Visible, "Hidden bound pill left its image"
    SetVisible host.pillBound, True
    Check host.pillBound.Visible And host.pillBound_bg.Visible, "Shown bound pill lost its image"
    ' SecretBox: hidden on open, the eye shows then hides it, HideSecrets hides it again.
    Check host.secret.Width = px(126) And host.eye_secret.Left = px(530), "Secret field pair is not 160px"
    Check host.secret.InputMask = SECRET_MASK, "Secret field opens shown"
    ToggleSecret host, "secret"
    Check host.secret.InputMask = "" And host.eye_secret.ControlTipText = "Hide", "Eye did not show the secret"
    Check Screen.ActiveControl.Name = "secret", "Eye left the focus off its field"
    ToggleSecret host, "secret"
    Check host.secret.InputMask = SECRET_MASK And host.eye_secret.Caption = ChrW(EYE_SHOW), "Eye did not hide the secret"
    ToggleSecret host, "secret"
    HideSecrets host
    Check host.secret.InputMask = SECRET_MASK And host.eye_secret.ControlTipText = "Show", "HideSecrets left the secret shown"
    grid.RecordSource = "SELECT * FROM test_rows ORDER BY id"
    ' Click-to-sort captions: ascending with an up arrow, then reversed; the arrow moves with the column.
    Check host.head1.OnClick = "" And Split(host.head1.Tag, "|")(3) = "", "Unsortable caption has a sort click"
    Check host.sortHead1.OnClick = "=GridSort([Form],""sfRows"",""sortHead"",1)", "Sortable caption has no sort click"
    GridSort host, "sfRows", "sortHead", 1
    Check grid.OrderByOn And grid.OrderBy Like "*description], [[]id]", "First caption click did not sort ascending"
    Check host.sortHead1.Caption = "DESCRIPTION " & ChrW(&H25B2), "Ascending caption lacks its arrow"
    GridSort host, "sfRows", "sortHead", 1
    Check grid.OrderBy Like "*description] DESC, [[]id]", "Second caption click did not reverse the sort"
    Check host.sortHead1.Caption = "DESCRIPTION " & ChrW(&H25BC), "Descending caption lacks its arrow"
    With grid.RecordsetClone
        .MoveFirst
        Check !description = "Description 9", "Rows are not in descending order"
    End With
    GridSort host, "sfRows", "sortHead", 0
    Check host.sortHead0.Caption = "CATEGORY " & ChrW(&H25B2) And host.sortHead1.Caption = "DESCRIPTION", "Arrow stayed on the old column"
    grid.OrderBy = "": grid.OrderByOn = False
    height = host.sfRows.Height
    LayoutTableFooter host, "sfRows", "pager", "id"
    Check Not host.pagerCount.Visible, "Unmarked table has footer"
    Check host.sfRows.Height = height, "Unmarked table lost height"
    Check grid.Recordset.RecordCount = 23, "Unmarked table was filtered"
    host.sfRows.Tag = "AppTag"
    SetTablePagination host.sfRows, True
    Check host.sfRows.Tag = "AppTag;Pagination:on", "Opt-in lost app tag"
    ResizeGrid grid, px(650), cols
    AlignGridHeaders host, "sfRows", "head", Array("txtCol0", "txtCol1")
    For i = 0 To 1
        Set c = grid.Controls("txtCol" & i)
        Check host.Controls("head" & i).Left = host.sfRows.Left + c.Left, "Header left differs"
        Check host.Controls("head" & i).Width = c.Width, "Header width differs"
        Check host.Controls("head" & i).TextAlign = c.TextAlign, "Header alignment differs"
        Check Abs(c.Top + c.Height \ 2 - grid.Section(acDetail).Height \ 2) <= 1, "Row vertical center differs"
    Next
    Check grid.txtCol1.Width > grid.txtCol0.Width, "Description did not expand"
    grid.Filter = "id > 3": grid.FilterOn = True
    LayoutTableFooter host, "sfRows", "pager", "id"
    Check grid.ScrollBars = 0, "Paginated grid retains vertical scrolling"
    Check host.pagerCount.Visible, "Marked footer hidden"
    Check host.sfRows.Height < height, "Footer did not reserve space"
    Check TableFullFilter(grid) = "id > 3", "Full export scope differs"
    Set ids = CreateObject("Scripting.Dictionary")
    Do
        Set rows = grid.Recordset.Clone
        Do While Not rows.EOF
            Check rows!id > 3, "Pager leaked scope"
            Check Not ids.Exists(CStr(rows!id)), "Duplicate page record"
            ids.Add CStr(rows!id), True: rows.MoveNext
        Loop
        rows.Close
        If Not MoveTablePage(host, "sfRows", "pager", "id", 1) Then Exit Do
    Loop
    Check ids.Count = 20, "Pager omitted records"
    Check Not host.pagerNext.Enabled, "Last page permits advance"
    ' Sorting a paginated table starts again at page 1 of the whole filtered set.
    GridSort host, "sfRows", "sortHead", 1
    With grid.RecordsetClone
        .MoveFirst
        Check !description = "Description 10", "Sorted pager kept the old order's page"
    End With
    Check host.pagerCount.Caption Like "Showing 1-* of 20", "Sorted pager did not return to page 1"
    grid.OrderBy = "": grid.OrderByOn = False
    LayoutTableFooter host, "sfRows", "pager", "id"
    SetTablePagination host.sfRows, False
    LayoutTableFooter host, "sfRows", "pager", "id"
    Check host.sfRows.Tag = "AppTag", "Opt-out lost app tag"
    Check Not host.pagerCount.Visible, "Opt-out footer visible"
    Check grid.ScrollBars = 2, "Opt-out did not restore vertical scrolling"
    Check host.sfRows.Height = height, "Opt-out did not restore height"
    Check grid.Filter = "id > 3", "Opt-out did not restore filter"
    Check grid.Recordset.RecordCount = 20, "Opt-out did not restore all rows"
    Check Not MoveTablePage(host, "sfRows", "pager", "id", 1), "Unmarked table permits page move"
    Check CountedCaption("UPL", 2) = "UPL (2)", "Count format differs"
    Check TagWidth(Array("Short", "Longest category")) = PillTextWidth("Longest category"), "Widest label differs"
    Check FileExists(CStr(ColumnPillPng(CLR_PRIMARY, 240, 26, PillTextWidth("Longest category")))), "Pill image missing"
    Check FileExists(RoundedCardPng(600, 200, CLR_SURFACE, CLR_BORDER)), "Rounded card missing"
    Check Not Nz(host.tglTest_val, False), "Unbound switch did not start off"
    ToggleSwitch host, "tglTest"
    Check host.tglTest_val, "Switch click did not turn it on"
    Check InStr(Nz(host.tglTest_bg.Properties("ControlSource").Value, ""), "TogglePng(") > 0, "Switch track is not bound"
    ToggleSwitch host, "tglTest"
    Check Not host.tglTest_val, "Second switch click did not turn it off"
    Check host.toggleChanges = 2, "Switch onChange did not run once per change"
    host.tglTest_val.Locked = True
    ToggleSwitch host, "tglTest"
    Check Not host.tglTest_val, "Locked switch changed"
    Check host.toggleChanges = 2, "Locked switch ran onChange"
    host.tglTest_val.Locked = False
    Check FileExists(CStr(TogglePng(True, CLR_PRIMARY, CLR_MUTED_FG, CLR_SURFACE, 44, 24))), "Switch image missing"
    Check FileExists(CStr(RingPng(CLR_PRIMARY, 52, 32))), "Focus ring image missing"
    Check Not host.tglTest_ring.Visible, "Focus ring shows without focus"
    host.tglTest.SetFocus
    Check host.tglTest_ring.Visible, "Focused switch shows no ring"
    host.txtFocusSink.SetFocus
    Check Not host.tglTest_ring.Visible, "Focus ring stayed after focus left"
    host.ringA.SetFocus: host.tglTest.SetFocus
    Check host.ringA.BorderColor = CLR_INPUT_BORDER, "Input kept its focus ring while the switch had the focus"
    ' Segmented control: starts on the first item, a click picks and calls onChange
    ' once, the picked item and a locked control ignore clicks, SetSegment is quiet.
    Check SegmentValue(host, "segTest") = "day", "Unbound segmented control did not start on its first item"
    Check host.segTest_th1.Left = host.segTest_th0.Left + host.segTest_th0.Width, "Segmented options are not side by side"
    Check host.segTest_b2.Left + host.segTest_b2.Width <= host.segTest_bg.Left + host.segTest_bg.Width, "Segmented option overruns its track"
    Check InStr(Nz(host.segTest_th1.Properties("ControlSource").Value, ""), "PillPng(") > 0, "Segmented thumb is not drawn"
    Check host.segTest_t0.FormatConditions.Count = 1, "Segmented caption has no picked format"
    SegmentClick host, "segTest", "week"
    Check SegmentValue(host, "segTest") = "week", "Segmented click did not pick its item"
    Check host.segmentChanges = 1, "Segmented onChange did not run"
    SegmentClick host, "segTest", "week"
    Check host.segmentChanges = 1, "Click on the picked item ran onChange"
    SetSegment host, "segTest", "month"
    Check SegmentValue(host, "segTest") = "month" And host.segmentChanges = 1, "SetSegment did not pick quietly"
    On Error Resume Next
    SetSegment host, "segTest", "year": badName = Err.Number: Err.Clear
    On Error GoTo 0
    Check badName = 5 And SegmentValue(host, "segTest") = "month", "SetSegment accepted an unknown item"
    host.segTest_val.Locked = True
    SegmentClick host, "segTest", "day"
    Check SegmentValue(host, "segTest") = "month" And host.segmentChanges = 1, "Locked segmented control changed"
    host.segTest_val.Locked = False
    Check Not host.segTest_ring.Visible, "Segmented focus ring shows without focus"
    host.segTest_b1.SetFocus
    Check host.segTest_ring.Visible, "Focused segmented option shows no ring"
    host.txtFocusSink.SetFocus
    Check Not host.segTest_ring.Visible, "Segmented focus ring stayed after focus left"
    ' Alert: built in its variant with a title; ShowAlert recolors it and moves an
    ' untitled message up; the close mark hides every part; unknown variants fail.
    Check host.alTest.BackColor = CLR_WARNING_BG And host.alTest_bar.BackColor = CLR_WARN_TEXT, "Alert is not in its variant's colors"
    Check host.alTest_title.Visible And host.alTest_msg.Top > host.alTest_title.Top, "Alert message is not below its title"
    Check AlertShown(host, "alTest"), "Alert does not show"
    ShowAlert host, "alTest", "Saved.", "success"
    Check host.alTest.BackColor = CLR_SUCCESS_BG And host.alTest_icon.ForeColor = CLR_GREEN_600, "ShowAlert did not recolor"
    Check Not host.alTest_title.Visible And host.alTest_msg.Top = host.alTest.Top + AlertMsgTop(host.alTest.Height, False), "Untitled alert message did not move up"
    Check host.alTest_msg.Caption = "Saved.", "ShowAlert did not set the message"
    HideAlert host, "alTest"
    Check Not AlertShown(host, "alTest") And Not host.alTest_x.Visible And Not host.alTest_bar.Visible, "HideAlert left parts showing"
    ShowAlert host, "alTest", "Could not save.", "destructive", "Error"
    Check AlertShown(host, "alTest") And host.alTest_title.Visible And host.alTest_x.Visible, "ShowAlert did not show it again"
    Check host.alTest_msg.Top = host.alTest.Top + AlertMsgTop(host.alTest.Height, True) And AlertMsgTop(px(36), False) = px(10), "Titled alert message did not move down"
    On Error Resume Next
    ShowAlert host, "alTest", "x", "loud": badName = Err.Number: Err.Clear
    On Error GoTo 0
    Check badName = 5, "ShowAlert accepted an unknown variant"
    ' Breadcrumb: built path in order, the last crumb current and not a link, a
    ' path wider than the trail folds its middle into an ellipsis, too many fail.
    Check host.bc0.Visible And host.bc2.Visible And Not host.bc3.Visible And Not host.bc_s3.Visible, "Breadcrumb shows the wrong crumbs"
    Check host.bc1.Left = host.bc_s1.Left + px(16) And host.bc_s1.Left = host.bc0.Left + host.bc0.Width, "Breadcrumb crumbs are not in a row"
    Check host.bc2.ForeColor = CLR_FG And host.bc0.ForeColor = CLR_MUTED_FG, "Breadcrumb current page is not set apart"
    CrumbClick host, "bc", 0
    Check host.lastCrumb = "home", "Breadcrumb click did not call its handler with the key"
    CrumbClick host, "bc", 2
    Check host.lastCrumb = "home", "Click on the current page's crumb called the handler"
    SetBreadcrumb host, "bc", "a|Alpha root", "b|Beta section", "c|Gamma subsection", "d|Delta page", "e|Epsilon detail"
    Check Not host.bc1.Visible And Not host.bc2.Visible And host.bc3.Visible And host.bc4.Visible, "Wide breadcrumb did not fold its middle"
    Check host.bc3.Caption = ChrW(&H2026) And host.bc3.Tag = "d" And host.bc3.ControlTipText = "Beta section > Gamma subsection > Delta page", "Folded crumbs are not named by the ellipsis"
    Check host.bc4.Left + host.bc4.Width <= host.bc.Left + host.bc.Width, "Folded breadcrumb overruns its width"
    CrumbClick host, "bc", 3
    Check host.lastCrumb = "d", "Ellipsis click did not go to the deepest folded crumb"
    SetBreadcrumb host, "bc", "x|One", "y|Two"
    Check host.bc1.Visible And Not host.bc2.Visible And host.bc1.Caption = "Two" And host.bc1.ForeColor = CLR_FG, "Shorter breadcrumb path left old crumbs"
    On Error Resume Next
    SetBreadcrumb host, "bc", "1|1", "2|2", "3|3", "4|4", "5|5", "6|6": badName = Err.Number: Err.Clear
    On Error GoTo 0
    Check badName = 5, "Breadcrumb took more items than slots"
    ' Stepper: done, current and later steps drawn apart; a done step's click goes
    ' back to it, later ones do nothing; SetStep moves it and rejects bad steps.
    Check host.st_f1.Visible And Not host.st_r1.Visible And host.st_n1.Caption = ChrW(&HE73E), "Done step is not checked"
    Check host.st_f2.Visible And host.st_n2.Caption = "2" And host.st_l2.FontWeight = 600 And host.st_l2.ForeColor = CLR_FG, "Current step is not set apart"
    Check host.st_r3.Visible And Not host.st_f3.Visible And host.st_l3.ForeColor = CLR_MUTED_FG, "Later step is not a ring"
    Check host.st_c1.BackColor = CLR_PRIMARY And host.st_c2.BackColor = CLR_INPUT_BORDER, "Step connectors are not colored by progress"
    Check host.st_c1.Left > host.st_f1.Left + host.st_f1.Width And host.st_c1.Left + host.st_c1.Width < host.st_f2.Left, "Step connector touches its circles"
    StepClick host, "st", 1
    Check host.lastStep = 1, "Done step click did not call the handler"
    StepClick host, "st", 3
    Check host.lastStep = 1, "Later step click called the handler"
    SetStep host, "st", 4
    Check CurrentStep(host, "st") = 4 And host.st_f3.Visible And host.st_n3.Caption = ChrW(&HE73E) And host.st_c2.BackColor = CLR_PRIMARY, "Stepper did not show every step done"
    SetStep host, "st", 1
    Check host.st_n1.Caption = "1" And host.st_r2.Visible And host.st_c1.BackColor = CLR_INPUT_BORDER, "Stepper did not go back to the first step"
    On Error Resume Next
    SetStep host, "st", 5: badName = Err.Number: Err.Clear
    On Error GoTo 0
    Check badName = 5 And CurrentStep(host, "st") = 1, "SetStep accepted a step out of range"
    ' Accordion: open and closed sections stacked, a header click opens or closes
    ' and moves the sections below, a one-open accordion closes the others.
    Check SectionOpen(host, "acc", 1) And Not SectionOpen(host, "acc", 2), "Accordion did not start with its open section"
    Check host.acc_v1.Caption = ChrW(&HE70D) And host.acc_v2.Caption = ChrW(&HE76C), "Accordion chevrons do not show open and closed"
    Check host.acc_b1.Top = host.acc.Top + px(48) And host.acc_h2.Top = host.acc_b1.Top + px(68), "Accordion sections are not stacked"
    AccordionClick host, "acc", 2
    Check SectionOpen(host, "acc", 2) And host.acc_h3.Top = host.acc_b2.Top + px(68) And host.acc_v2.Caption = ChrW(&HE70D), "Opened section did not push the next one down"
    AccordionClick host, "acc", 1
    Check Not SectionOpen(host, "acc", 1) And host.acc_h2.Top = host.acc.Top + px(41) And host.acc_u1.Top = host.acc.Top + px(40), "Closed section did not let the next one up"
    Check host.acc.Height = host.acc_b2.Top + px(68) + px(41) - host.acc.Top, "Accordion frame is not its sections' height"
    OpenSection host, "acs", 2, True
    Check SectionOpen(host, "acs", 2) And Not SectionOpen(host, "acs", 1), "One-open accordion left another section open"
    On Error Resume Next
    OpenSection host, "acc", 4, True: badName = Err.Number: Err.Clear
    On Error GoTo 0
    Check badName = 5, "OpenSection accepted a missing section"
    ' Avatar: initials from the first and last words, a stable tone per name, the
    ' empty name's own tone, and the built circle and initials.
    Check AvatarInitials("ETN1 Rivera") = "ER" And AvatarInitials("  ada   byron  lovelace ") = "AL", "Avatar initials differ"
    Check AvatarInitials("Plato") = "P" And AvatarInitials(Null) = "" And AvatarInitials("   ") = "", "One-word or empty avatar initials differ"
    Check AvatarTone("Ada Lovelace") = AvatarTone(" ada lovelace ") And AvatarTone("Ada Lovelace") <= 4, "Avatar tone is not stable"
    Check AvatarTone(Null) = 5 And AvatarTone("") = 5, "Empty name has no tone of its own"
    host.Recalc
    Check host.avTest = "AL" And host.avTest.FormatConditions.Count = 5, "Avatar does not show its initials"
    Check InStr(Nz(host.avTest_bg.Properties("ControlSource").Value, ""), "PillPng(Choose(AvatarTone(") > 0, "Avatar circle is not drawn by tone"
    Check host.avTest_bg.Width = px(32) And host.avTest.Width = px(32) And host.avTest.FontSize = 10, "Avatar is not 32px with 10pt initials"
    ' Stat tile: change badge text and tone, the built value and badge.
    Check StatDelta(8) = ChrW(&H25B2) & " 8%" And StatDelta(-3.5) = ChrW(&H25BC) & " 3.5%", "Stat change text differs"
    Check StatDelta(0) = "0%" And StatDelta(Null) = "" And StatDelta("x") = "", "Flat or empty stat change differs"
    Check StatTone(5, True) = 0 And StatTone(5, False) = 1 And StatTone(-2, True) = 1 And StatTone(0, True) = 2 And StatTone(Null, True) = 2, "Stat tone differs"
    ' (Comparing the badge's Value to a ChrW string inside a longer And chain
    ' crashed VBA 7 on 64-bit Access, so each part is read on its own.)
    text = Nz(host.kpi_d.Value, "")
    Check host.kpi.Value = 12, "Stat tile does not show its value"
    Check text = ChrW(&H25BC) & " 3.5%", "Stat tile does not show its change"
    Check host.kpi_d.FormatConditions.Count = 2, "Stat tile badge has no tone formats"
    Check host.kpi_ico.Caption = ChrW(&HE8F4) And host.kpi_cap.Left > host.kpi_dbg.Left + host.kpi_dbg.Width, "Stat tile icon or caption misplaced"
    ' Empty state: shown as built, hidden and shown again with a new title.
    Check host.emp_t.Visible And host.emp_b.Visible And host.emp_bg.Visible, "Empty state built shown is hidden"
    ShowEmpty host, "emp", False
    Check Not host.emp_t.Visible And Not host.emp_b.Visible And Not host.emp_ico.Visible And Not host.emp.Visible, "ShowEmpty left parts showing"
    ShowEmpty host, "emp", True, "Nothing here"
    Check host.emp_t.Visible And host.emp_t.Caption = "Nothing here" And host.emp_m.Caption = "Try another search.", "ShowEmpty did not show the new title"
    ' Search box: text in code calls the handler and shows the clear mark; the
    ' mark empties it, focuses it and calls the handler with ""; ring on the frame.
    Check Not host.srch_x.Visible, "Empty search box shows its clear mark"
    SetSearch host, "srch", "pump"
    Check host.lastSearch = "pump" And host.srch_x.Visible And SearchText(host, "srch") = "pump", "SetSearch did not search"
    SearchClear host, "srch"
    Check host.lastSearch = "" And Not host.srch_x.Visible And IsNull(host.srch.Value) And host.ActiveControl.Name = "srch", "Clear mark did not empty the search"
    Check host.srch_f.BorderColor = CLR_PRIMARY, "Focused search box lacks its ring"
    host.txtFocusSink.SetFocus
    Check host.srch_f.BorderColor = CLR_INPUT_BORDER, "Search box kept its ring"
    ' Rating: empty stars, a click fills up to it and calls onChange, a click on
    ' the value clears it, SetRating is quiet and checks its range.
    Check RatingValue(host, "rt") = 0 And host.rt_s1 = ChrW(59188), "Rating did not start empty"
    RatingClick host, "rt", 4
    text = Nz(host.rt_s4.Value, "")
    Check RatingValue(host, "rt") = 4 And text = ChrW(59189) And host.ratingChanges = 1, "Star click did not set the rating"
    text = Nz(host.rt_s5.Value, "")
    Check text = ChrW(59188), "Star past the rating is filled"
    RatingClick host, "rt", 4
    Check RatingValue(host, "rt") = 0 And host.ratingChanges = 2, "Click on the rating did not clear it"
    SetRating host, "rt", 3
    Check RatingValue(host, "rt") = 3 And host.ratingChanges = 2, "SetRating was not quiet"
    On Error Resume Next
    SetRating host, "rt", 6: badName = Err.Number: Err.Clear
    On Error GoTo 0
    Check badName = 5, "SetRating accepted a value past its stars"
    ' Slider: steps cover the track, a click moves it and calls onChange once,
    ' SetSlider rounds to a step and stays in range, SliderStep moves by steps.
    Check SliderValue(host, "sl") = 0 And host.sl_b0.Left = host.sl_bg.Left, "Slider did not start at its minimum"
    Check host.sl_b10.Left + host.sl_b10.Width = host.sl_bg.Left + host.sl_bg.Width, "Slider steps do not cover the track"
    SliderClick host, "sl", 7
    Check SliderValue(host, "sl") = 70 And host.sliderChanges = 1 And host.sl_b7.TabStop And Not host.sl_b0.TabStop, "Slider click did not move it"
    SliderClick host, "sl", 7
    Check host.sliderChanges = 1, "Click on the slider's step ran onChange"
    SetSlider host, "sl", 44
    Check SliderValue(host, "sl") = 40 And host.sliderChanges = 1, "SetSlider did not round quietly"
    SetSlider host, "sl", 500
    Check SliderValue(host, "sl") = 100, "SetSlider left the range"
    SliderStep host, "sl", -2
    Check SliderValue(host, "sl") = 80 And host.sliderChanges = 2, "SliderStep did not move two steps"
    Check host.sl_lbl = 80, "Slider value label differs"
    Check FileExists(CStr(SliderPng(37, CLR_PRIMARY, CLR_SECONDARY_BG, CLR_SURFACE, 200, 20))), "Slider image missing"
    Check IsNull(SliderPng(50, CLR_PRIMARY, CLR_SECONDARY_BG, CLR_SURFACE, 20, 20)), "Slider image drawn narrower than its knob"
    ' Popover: opens with its title and text, sized to the text, Esc closes it.
    Check KitPopover(host.tglTest, "Due date", "Work is due then."), "Popover did not open"
    Check IsFormOpen("dlgPopover"), "Popover form is not open"
    With Forms("dlgPopover")
        Check .Controls("lblTitle").Caption = "Due date" And .Controls("lblTitle").Visible, "Popover title differs"
        Check Abs(.InsideHeight - (px(12) + px(24) + POPOVER_LINE_H + px(12))) <= 15, "Popover is not sized to one line (" & .InsideHeight & ")"
    End With
    Dim pk As Integer
    pk = vbKeyEscape: PopoverKey Forms("dlgPopover"), pk
    Check Not IsFormOpen("dlgPopover"), "Esc did not close the popover"
    Check KitPopover(host.tglTest, "", "No title"), "Untitled popover did not open"
    Check Not Forms("dlgPopover").Controls("lblTitle").Visible And Forms("dlgPopover").Controls("lblText").Top = px(12), "Untitled popover kept its title"
    MenuClose Forms("dlgPopover")
    DoCmd.SelectObject acForm, "frmTest"
    ' Dropdown menu: placed under its button (right-aligned or above when it would
    ' run off), sized to its items, separators and disabled items skipped by the
    ' keys, a pick returns the key and closes it, Esc returns nothing.
    Dim mx As Long, my As Long, key As Integer
    MenuSpot 100, 100, 130, 130, 200, 150, 0, 0, 1000, 800, 4, mx, my
    Check mx = 100 And my = 134, "Menu is not under its button"
    MenuSpot 900, 100, 930, 130, 200, 150, 0, 0, 1000, 800, 4, mx, my
    Check mx = 730 And my = 134, "Menu at the right edge is not right-aligned"
    MenuSpot 100, 700, 130, 730, 200, 150, 0, 0, 1000, 800, 4, mx, my
    Check mx = 100 And my = 546, "Menu at the bottom is not above its button"
    Check MenuOpen(host.tglTest, Array("edit|Edit R&D|E70F", "-", "~arch|Archive", "del|!Delete|E74D")), "Menu did not open"
    Check IsFormOpen("dlgMenu"), "Menu form is not open"
    With Forms("dlgMenu")
        Check .Controls("lblCap0").Caption = "Edit R&&D" And .Controls("lblIcon0").Caption = ChrW(&HE70F), "Menu item caption or icon differs"
        Check .Controls("ruleSep1").Visible And Not .Controls("lblCap1").Visible, "Menu separator is not a line"
        Check .Controls("lblCap2").ForeColor = CLR_MUTED_FG And .Controls("lblCap3").ForeColor = CLR_DESTRUCTIVE, "Disabled or destructive item colors differ"
        Check Not .Controls("lblCap4").Visible, "Menu shows unused rows"
        Check .InsideHeight = px(4 + 32 + 9 + 32 + 32 + 4), "Menu is not sized to its items"
        key = vbKeyDown: MenuKey Forms("dlgMenu"), key
        Check .Controls("boxRow0").Visible, "Down did not light the first item"
        key = vbKeyDown: MenuKey Forms("dlgMenu"), key
        Check .Controls("boxRow3").Visible And Not .Controls("boxRow0").Visible, "Down did not skip the separator and the disabled item"
        MenuPick Forms("dlgMenu"), 2
        Check IsFormOpen("dlgMenu"), "Disabled item closed the menu"
    End With
    key = vbKeyReturn: MenuKey Forms("dlgMenu"), key
    Check Not IsFormOpen("dlgMenu") And MenuResult() = "del", "Enter did not pick the lit item"
    Check MenuOpen(host.tglTest, Array("a|Alpha", "b|Beta")), "Menu did not open again"
    key = vbKeyEscape: MenuKey Forms("dlgMenu"), key
    Check Not IsFormOpen("dlgMenu") And MenuResult() = "", "Esc did not close the menu with nothing picked"
    On Error Resume Next
    MenuOpen host.tglTest, Array("a|A", "b|B", "c|C", "d|D", "e|E", "f|F", "g|G", "h|H", "i|I", "j|J", "k|K"): badName = Err.Number: Err.Clear
    MenuOpen host.tglTest, Array("nokey"): badScale = Err.Number: Err.Clear
    On Error GoTo 0
    Check badName = 5 And badScale = 5 And Not IsFormOpen("dlgMenu"), "Menu took too many items or a bad one"
    DoCmd.SelectObject acForm, "frmTest"        ' the menu was the active window
    host.ringA.SetFocus
    Check host.ringA.BorderColor = CLR_PRIMARY, "Focused input lacks its focus ring"
    host.ringB.SetFocus
    Check host.ringA.BorderColor = CLR_INPUT_BORDER, "Input kept its focus ring after losing the focus"
    Check host.ringB.BorderColor = CLR_PRIMARY, "Focused combo box lacks its focus ring"
    host.ringA.BorderColor = CLR_DESTRUCTIVE: host.ringA.SetFocus: host.ringB.SetFocus
    Check host.ringA.BorderColor = CLR_DESTRUCTIVE, "Focus ring lost the input's own border color"
    host.sfRing.SetFocus: host.sfRing.Form.ringC.SetFocus
    Check host.sfRing.Form.ringC.BorderColor = CLR_PRIMARY, "Subform input lacks its focus ring"
    host.plainTxt.SetFocus
    Check host.sfRing.Form.ringC.BorderColor = CLR_INPUT_BORDER, "Subform input kept its ring when the focus left the subform"
    host.sfRing.SetFocus
    Check host.sfRing.Form.ringC.BorderColor = CLR_PRIMARY, "Subform input lost its ring when the focus came back"
    host.plainTxt.SetFocus: host.sfRing.SetFocus: host.sfRing.Form.ringSave.SetFocus
    Check host.sfRing.Form.ringC.BorderColor = CLR_INPUT_BORDER, "Subform input kept its ring while its Save button had the focus"
    host.ringB.Locked = True: host.ringA.SetFocus: host.ringB.SetFocus
    Check host.ringB.BorderColor = CLR_INPUT_BORDER, "Locked input shows a focus ring"
    ' The ring checks leave the focus on a host control; move it into the grid
    ' first, or Screen.ActiveControl stays on the host.
    host.sfRows.SetFocus: grid.txtCol0.SetFocus: grid.RowClick
    Check Screen.ActiveControl.Name = "txtFocusSink", "Row click left text focused (" & Screen.ActiveControl.Name & ")"
    Check host.callbackGrid = "sfRows", "Grid callback used wrong parent"
    Check CLng(grid.txtSel) = CLng(grid!id), "Row click lost record selection"
    host.sfReuse.Form.RecordSource = "SELECT * FROM test_rows ORDER BY id"
    host.sfReuse.Form.txtCol0.SetFocus: host.sfReuse.Form.RowClick
    Check host.callbackGrid = "sfReuse", "Reused grid callback selected first instance"
    ResizeGrid grid, px(200), cols
    Check grid.ScrollBars = 3, "Narrow grid clips minimum columns instead of scrolling"
    Check grid.txtCol1.Width >= px(292), "Narrow grid discarded Description width"
    GridHighlight host.sfRows, "desc"
    Check InStr(Nz(grid.txtCol1, ""), ">Desc</font>") > 0, "Table search match not marked"
    GridHighlight host.sfRows, ""
    Check InStr(Nz(grid.txtCol1, ""), "<font") = 0, "Cleared table search left a mark"
    host.sfName.Form.RecordSource = "SELECT id, description AS name FROM test_rows ORDER BY id"
    GridHighlight host.sfName, "desc"
    Check InStr(Nz(host.sfName.Form.txtCol0, ""), ">Desc</font>ription 1") > 0, "Search column named name shows the wrong value"
    host.reqA = "   "
    host.reqB = Null
    text = RequiredMissing(host, "reqA=First;reqB=Second", fitted)
    Check text = ChrW(&H2022) & " First" & vbCrLf & ChrW(&H2022) & " Second", "Required list differs"
    Check fitted = "reqA", "First missing required field differs"
    Check host.reqA.BorderColor = Th(CLR_DESTRUCTIVE), "Missing required field not outlined"
    host.reqA = "x"
    Check RequiredMissing(host, "reqA=First", fitted) = "" And fitted = "", "Filled required field reported"
    Check host.reqA.BorderColor = Th(CLR_INPUT_BORDER), "Filled required field kept its outline"
    DoCmd.OpenForm "dlgToast", WindowMode:=acHidden     ' as KitToast does, but errors raise here
    Forms("dlgToast").ShowToast "Saved R&D", "success", 3000
    Check Forms("dlgToast").lblMsg.Caption = "Saved R&&D", "Toast took & as an access key"
    KitToast "Saved", "success"
    Check Forms("dlgToast").lblMsg.Caption = "Saved", "Toast lost its message"
    Check Screen.ActiveForm.Name = "frmTest", "Toast took the focus"
    DoCmd.Close acForm, "dlgToast", acSaveNo
    Check SaveDetail(host, "sfRows"), "Unchanged detail panel reported a refused save"
    host.sfRefuse.Form.txtDesc = "Changed"
    Check Not SaveDetail(host, "sfRefuse"), "Refused detail save reported as saved"
    Check host.sfRefuse.Form.Dirty, "Refused detail save dropped the change"
    Check Not ShowDetail(host, "sfRefuse", "subRefuse", "id=2"), "ShowDetail moved on after a refused save"
    host.sfRefuse.Form.Undo
    DoCmd.Close acForm, "frmTest", acSaveNo
    ' Trust check: matching Trusted Locations, a codeless dlgTrust, and the AutoExec
    ' macro that opens it with the code off.
    Check PathCovers("C:\Apps\", False, "c:\apps") And Not PathCovers("C:\Apps", False, "C:\Apps\Sub"), "Trusted Location matching differs"
    Check PathCovers("C:\", True, "C:\Apps\Sub") And Not PathCovers("C:\Apps", True, "C:\AppsOld"), "Trusted Location subfolder matching differs"
    Check RecordPath("%USERPROFILE%/My%20Apps/App.accdb") = LCase$(Environ$("USERPROFILE") & "\my apps\app.accdb"), "Trusted document name not decoded"
    BuildTrustCheck "frmTest"
    DoCmd.OpenForm "dlgTrust", acDesign
    Check Not Forms("dlgTrust").HasModule And Len(Forms("dlgTrust").OnOpen) = 0 And Len(Forms("dlgTrust").OnClose) = 0, _
        "dlgTrust needs code to open or close"
    Check Forms("dlgTrust").btnClose.OnClick = "mcrTrustClose", "dlgTrust Close doesn't run its macro"
    DoCmd.Close acForm, "dlgTrust", acSaveNo
    text = CurrentProject.Path & "\autoexec_test.txt"
    Application.SaveAsText acMacro, "AutoExec", text
    fitted = Replace(ReadTextFile(text), vbNullChar, "")   ' SaveAsText may write UTF-16
    Kill text
    Check InStr(fitted, "IsTrusted") > 0 And InStr(fitted, "dlgTrust") > 0 And InStr(fitted, "TrustStartup") > 0, "AutoExec lost its trust rows"
    Application.SaveAsText acMacro, "mcrTrustClose", text
    fitted = Replace(Replace(Replace(Replace(ReadTextFile(text), vbNullChar, ""), " ", ""), vbCr, ""), vbLf, "")
    Kill text
    ' With the code off Access only runs Close with Save at Prompt, so No (2) or Yes (1) breaks it
    Check InStr(fitted, "dlgTrust") > 0 And InStr(fitted, "Argument=""dlgTrust""Argument=""1""") = 0 And _
        InStr(fitted, "Argument=""dlgTrust""Argument=""2""") = 0, "mcrTrustClose doesn't close with Save at Prompt"
    BuildMessageDialog
    DoCmd.OpenForm "dlgMessage"
    Set host = Forms("dlgMessage")
    Check host.btn1.Caption = "OK" And host.btn1.Visible And host.btn1.Default And host.btn1.Cancel, "Message dialog lacks its OK button"
    Check Not host.btn2.Visible And Not host.btn3.Visible, "Message dialog shows unused buttons"
    Check host.btn1.Left + host.btn1.Width = host.btn3.Left + host.btn3.Width, "Message OK button is not right-aligned"
    DoCmd.Close acForm, "dlgMessage", acSaveNo
    BuildInputDialog
    DoCmd.OpenForm "dlgInput"
    Set host = Forms("dlgInput")
    Check host.btnOK.Default And host.btnCancel.Cancel, "Input dialog lacks its Enter and Esc buttons"
    Check host.btnCancel.Left + host.btnCancel.Width = px(400), "Input Cancel button is not right-aligned"
    Check host.btnOK.Left + host.btnOK.Width + px(8) = host.btnCancel.Left, "Input OK button is not 8px left of Cancel"
    Check Not host.txtPrompt.Visible And host.txtValue.Top = px(22), "Input dialog without a prompt keeps the prompt's space"
    Check Len(host.txtValue.InputMask) = 0, "Unmasked input dialog masks its text"
    DoCmd.Close acForm, "dlgInput", acSaveNo
    ' SecretBox refuses a continuous form: one eye would show every row.
    NewForm "", 1, px(300), px(100)
    On Error Resume Next
    SecretBox "secret", "", 0, 0, px(160), px(30): badView = Err.Number: Err.Clear
    On Error GoTo 0
    DoCmd.Close acForm, F.Name, acSaveNo
    Check badView = 5, "SecretBox accepted a continuous form"
    NewForm "", 1, px(300), px(100)
    On Error Resume Next
    Segmented "seg", "", 0, 0, px(200), "a|A;b|B": badView = Err.Number: Err.Clear
    On Error GoTo 0
    DoCmd.Close acForm, F.Name, acSaveNo
    Check badView = 5, "Segmented accepted a continuous form"
    NewForm "", 0, px(300), px(100)
    On Error Resume Next
    Segmented "seg1", "", 0, 0, px(200), "a|A": badName = Err.Number: Err.Clear
    Segmented "seg2", "", 0, 0, px(200), "a|A;a|Again": badScale = Err.Number: Err.Clear
    Segmented "seg3", "", 0, 0, px(100), "a|A;b|B;c|C": badView = Err.Number: Err.Clear
    On Error GoTo 0
    DoCmd.Close acForm, F.Name, acSaveNo
    Check badName = 5, "Segmented accepted one item"
    Check badScale = 5, "Segmented accepted a repeated key"
    Check badView = 5, "Segmented accepted options narrower than 40px"
    NewForm "", 0, px(300), px(100)
    On Error Resume Next
    Alert "al", 0, 0, px(200), px(60), "loud": badName = Err.Number: Err.Clear
    On Error GoTo 0
    DoCmd.Close acForm, F.Name, acSaveNo
    Check badName = 5, "Alert accepted an unknown variant"
    NewForm "", 0, px(300), px(100)
    On Error Resume Next
    Breadcrumb "bc", 0, 0, px(200), 1, "x": badName = Err.Number: Err.Clear
    On Error GoTo 0
    DoCmd.Close acForm, F.Name, acSaveNo
    Check badName = 5, "Breadcrumb accepted one slot"
    NewForm "", 0, px(300), px(100)
    On Error Resume Next
    Stepper "st1", 0, 0, px(300), 1, "", "Only": badName = Err.Number: Err.Clear
    Stepper "st2", 0, 0, px(150), 1, "", "A", "B", "C": badScale = Err.Number: Err.Clear
    On Error GoTo 0
    DoCmd.Close acForm, F.Name, acSaveNo
    Check badName = 5, "Stepper accepted one step"
    Check badScale = 5, "Stepper accepted steps under 64px"
    NewForm "", 0, px(300), px(100)
    On Error Resume Next
    Accordion "ac1", 0, 0, px(300), False, "Title only": badName = Err.Number: Err.Clear
    Accordion "ac2", 0, 0, px(300), False, "Title||tall": badScale = Err.Number: Err.Clear
    On Error GoTo 0
    DoCmd.Close acForm, F.Name, acSaveNo
    Check badName = 5 And badScale = 5, "Accordion accepted a bad section"
    NewForm "", 1, px(300), px(100)
    Rating "rtRo", "", 0, 0, readOnly:=True
    On Error Resume Next
    Rating "rt1", "", 0, px(40): badName = Err.Number: Err.Clear
    Slider "sl1", "", 0, px(70), px(200): badView = Err.Number: Err.Clear
    On Error GoTo 0
    DoCmd.Close acForm, F.Name, acSaveNo
    Check badName = 5 And badView = 5, "Clickable rating or slider accepted a continuous form"
    NewForm "", 0, px(300), px(100)
    On Error Resume Next
    Rating "rt2", "", 0, 0, 2: badName = Err.Number: Err.Clear
    Slider "sl2", "", 0, px(40), px(200), 0, 100, 30: badScale = Err.Number: Err.Clear
    Slider "sl3", "", 0, px(70), px(200), 0, 100, 1: badView = Err.Number: Err.Clear
    On Error GoTo 0
    DoCmd.Close acForm, F.Name, acSaveNo
    Check badName = 5 And badScale = 5 And badView = 5, "Rating or slider accepted a bad count or step"
    ' Splash: CloseSplash and SplashStatus do nothing without one, the footer clears the
    ' subtitle, and the bitmap that hides Access's splash is 58 bytes.
    CloseSplash
    SplashStatus "Starting"
    BuildSplash "frmMain", "Subtitle"
    DoCmd.OpenForm "frmSplash", acDesign
    Set host = Forms("frmSplash")
    Check host.lblSplashSub.Top + host.lblSplashSub.Height <= host.lnSplash.Top, "Splash subtitle overlaps its footer"
    CloseSplash
    Check IsFormOpen("frmSplash"), "CloseSplash closed a splash in Design view"
    DoCmd.Close acForm, "frmSplash", acSaveNo
    HideAccessSplash
    text = CurrentProject.Path & "\" & Left$(CurrentProject.Name, InStrRev(CurrentProject.Name, ".") - 1) & ".bmp"
    Check FileLen(text) = 58, "HideAccessSplash did not write a 58-byte bitmap"
    Kill text
    Check CompareVersions("1.10.0", "1.9.2") = 1, "Version parts compared as text"
    Check CompareVersions("v1.2", "1.2.0") = 0, "Missing version part not read as 0"
    Check CompareVersions("1.0.0", "1.0.1-beta") = -1, "Version suffix not ignored"
    Check ReleasedVersion("") = "", "Blank release folder read a version"
    text = CurrentProject.Path & "\version-test"          ' the run's own folder (native-smoke.ps1)
    EnsureFolder text
    WriteTextFile text & "\" & KIT_VERSION_FILE, vbCrLf & " 1.4.0 " & vbCrLf & "notes"
    Check ReleasedVersion(text & "\") = "1.4.0", "Release folder version misread"
    Kill text & "\" & KIT_VERSION_FILE: RmDir text
    Check ReleasedVersion(text) = "", "Missing version file read a version"
    Check FederalHoliday(#10/12/2026#) = "Columbus Day", "Columbus Day 2026 missing"
    Check FederalHoliday(#5/25/2026#) = "Memorial Day", "Memorial Day 2026 missing"
    Check FederalHoliday(#11/26/2026#) = "Thanksgiving Day", "Thanksgiving 2026 missing"
    Check FederalHoliday(#7/3/2026#) = "Independence Day (observed)", "Saturday holiday not observed on Friday"
    Check FederalHoliday(#7/4/2026#) = "", "Saturday holiday also marked on its own day"
    Check FederalHoliday(#12/31/2021#) = "New Year's Day (observed)", "New Year's Day 2022 not observed on 31 December"
    Check FederalHoliday(#6/19/2020#) = "", "Juneteenth marked before 2021"
    Check FederalHoliday(#10/13/2026 3:00:00 PM#) = "", "Ordinary day marked as a holiday"
    Check FederalHoliday(#7/5/2027#) = "Independence Day (observed)", "Sunday holiday not observed on Monday"
    Check FederalHoliday(#1/1/2026#) = "New Year's Day", "New Year's Day 2026 missing"
    Check FederalHoliday(#1/19/2026#) = "Birthday of Martin Luther King, Jr.", "MLK Day 2026 missing"
    Check FederalHoliday(#2/16/2026#) = "Washington's Birthday", "Washington's Birthday 2026 missing"
    Check FederalHoliday(#6/19/2026#) = "Juneteenth National Independence Day", "Juneteenth 2026 missing"
    Check FederalHoliday(#9/7/2026#) = "Labor Day", "Labor Day 2026 missing"
    Check FederalHoliday(#11/11/2026#) = "Veterans Day", "Veterans Day 2026 missing"
    Check FederalHoliday(#12/25/2026#) = "Christmas Day", "Christmas Day 2026 missing"
    Check FederalHoliday(#11/11/1977#) = "", "Holiday rules applied before 1978"
    ClearMarkedDates
    MarkDate #10/20/2026 9:00:00 AM#, "Report due"
    MarkDate #10/20/2026#, "Drill"
    Check MarkedDate(#10/20/2026#) = "Report due; Drill", "Marked date labels differ"
    ClearMarkedDates
    Check MarkedDate(#10/20/2026#) = "", "ClearMarkedDates left a mark"
    ' Dialog placement (CenterOnApp): centred on the Access window, pulled back
    ' inside the monitor's work area, and never above or left of it.
    Check CenteredStart(0, 1000, 400, 0, 1000) = 300, "Dialog not centred on Access"
    Check CenteredStart(1920, 3840, 400, 1920, 3840) = 2680, "Dialog not centred on the second monitor"
    Check CenteredStart(1700, 2300, 400, 0, 1920) = 1520, "Dialog left past the work area's right edge"
    Check CenteredStart(-300, 300, 400, 0, 1920) = 0, "Dialog left past the work area's left edge"
    Check CenteredStart(0, 1000, 1200, 0, 1000) = 0, "Oversized dialog hides its title bar"
    Check ErrorLineSource() = "test line 8", "Error log did not record the failing line (Erl)"
    ' Backdrop: nested dims share one layer, and a dialog that fails to open
    ' leaves the app undimmed.
    ' One call per line: "DimApp:" at the start of a line is a line label, not a call.
    DimApp
    DimApp
    UndimApp
    Check AppDimmed(), "Inner dialog removed the outer dialog's dim"
    UndimApp
    Check Not AppDimmed(), "Dim left on after its last dialog"
    On Error Resume Next
    RunDialog "dlgNoSuchForm"
    badName = Err.Number
    On Error GoTo 0
    Check badName <> 0, "RunDialog hid a missing dialog"
    Check Not AppDimmed(), "Dim left on after a dialog failed to open"
    BackdropOnClose
    Check Not AppDimmed(), "A dialog closing with no dim waiting changed the count"
    ' Column chooser on a fit table: order, hiding, captions, bad layouts.
    Dim order() As Long, shown() As Boolean, sf As Control, capTag() As String
    Check ParseLayout("2,0,1h", 3, order, shown), "Valid column layout rejected"
    Check order(0) = 2 And Not shown(2), "Column layout parsed wrongly"
    Check Not ParseLayout("0,0,1", 3, order, shown), "Repeated column accepted"
    Check Not ParseLayout("0,1", 3, order, shown), "Short column layout accepted"
    Check Not ParseLayout("0h,1h,2h", 3, order, shown), "Layout with no shown column accepted"
    cols = Array("category|Category|160|", "description|Description|200|", "id|Id|80|r")
    BuildGrid "subTestFit", cols, px(600), fit:=True
    NewForm "", 0, px(700), px(300)
    GridHeaders "fit", cols, px(24), px(20)
    SubF "sfFit", "Form.subTestFit", px(24), px(42), px(600), px(180)
    SaveForm "frmTestFit", "testHost"
    DoCmd.OpenForm "frmTestFit"
    Set host = Forms("frmTestFit"): Set sf = host.sfFit
    sf.Form.RecordSource = "SELECT * FROM test_rows ORDER BY id"
    SetGridColumns sf, "fit", DefaultLayout(3)
    Set grid = sf.Form
    height = grid.txtCol2.Left + grid.txtCol2.Width          ' right edge of all the columns
    capTag = Split(host.fit2.Tag, "|")
    capTag(3) = "id"                                          ' as a sortable caption's
    host.fit2.Tag = Join(capTag, "|")
    SetGridColumns sf, "fit", "2,0,1h"
    Set grid = sf.Form
    Check Abs(grid.txtCol0.Left + grid.txtCol0.Width - height) <= px(2), "Shown columns did not fill the table as far as all of them"
    Check grid.txtCol2.Left < grid.txtCol0.Left, "Moved column did not move"
    Check Not grid.txtCol1.Visible And Not host.fit1.Visible, "Hidden column or caption still shows"
    Check Abs(host.fit2.Left - (sf.Left + grid.txtCol2.Left)) <= 15, "Caption did not follow its column"
    Check grid.txtCol0.Width > px(160), "Shown columns did not widen into the hidden one's space"
    Check grid.txtCol1.Width <= 15, "Hidden column kept its width (can push the form past its limit)"
    Check Split(host.fit2.Tag, "|")(3) = "id", "Caption lost a later Tag part (sort field)"
    Check host.fit2.Tag Like "*|fw=" & px(700), "Caption lost its form's design width"
    On Error Resume Next
    SetGridColumns sf, "fit", "0,1"
    Check Err.Number = 5, "Bad column layout was applied"
    On Error GoTo 0
    SetGridColumns sf, "fit", DefaultLayout(3)
    Check grid.txtCol0.Left < grid.txtCol1.Left And grid.txtCol1.Left < grid.txtCol2.Left, "Default layout lost design order"
    Check grid.txtCol1.Visible And host.fit1.Visible, "Default layout left a column hidden"
    DoCmd.Close acForm, "frmTestFit", acSaveNo
    ' Header filters (modGridFilter): funnels, value lists, and filters kept
    ' alongside the page's filter, sorting and the column chooser.
    Dim vals As Variant, more As Boolean
    CurrentDb.Execute "CREATE TABLE test_filter (id LONG PRIMARY KEY, category TEXT(60), qty LONG)"
    Set db = CurrentDb                     ' held: a field reached through a bare CurrentDb raises 3420
    db.TableDefs("test_filter").Fields("category").AllowZeroLength = True
    CurrentDb.Execute "INSERT INTO test_filter VALUES (1,'Alpha',5)"
    CurrentDb.Execute "INSERT INTO test_filter VALUES (2,'Beta',5)"
    CurrentDb.Execute "INSERT INTO test_filter VALUES (3,'Alpha',7)"
    CurrentDb.Execute "INSERT INTO test_filter VALUES (4,NULL,7)"
    CurrentDb.Execute "INSERT INTO test_filter VALUES (5,'',9)"
    CurrentDb.Execute "INSERT INTO test_filter VALUES (6,'O''Brien',9)"
    ' Qty has a sort field (a fifth part); Category has none.
    cols = Array("category|Category|160|", "qty|Qty|80|r|id", "=[qty]*2|Double|80|r")
    BuildGrid "subTestFilter", cols, px(500), fit:=True
    NewForm "", 0, px(640), px(300)
    GridHeaders "flt", cols, px(24), px(20), sortSf:="sfFlt", filterSf:="sfFlt"
    SubF "sfFlt", "Form.subTestFilter", px(24), px(42), px(500), px(180)
    SaveForm "frmTestFilter", "testHost"
    DoCmd.OpenForm "frmTestFilter"
    Set host = Forms("frmTestFilter"): Set sf = host.sfFlt
    sf.Form.RecordSource = "SELECT * FROM test_filter ORDER BY id"
    Check HasControl(host, "fltFilter0") And HasControl(host, "fltFilter1"), "Field column has no funnel"
    Check Not HasControl(host, "fltFilter2"), "Computed column got a funnel"
    Check host.flt0.Left + host.flt0.Width <= host.fltFilter0.Left, "Caption runs under its funnel"
    Check host.fltFilter0.Left + host.fltFilter0.Width <= host.flt1.Left, "Funnel overlaps the next caption"
    Check host.fltFilter0.ForeColor = Th(CLR_MUTED_FG), "Unfiltered funnel not muted"
    vals = GridColumnValues(sf, "flt", 0, more)
    Check UBound(vals) = 3 And Not more, "Column values not distinct, or blanks not merged"
    Check vals(0) = "Alpha" And vals(1) = "Beta" And vals(2) = "O'Brien" And IsNull(vals(3)), "Column values not sorted with blanks last"
    SetGridFilter sf, "qty < 9"
    Check RowCount(sf.Form) = 4, "Page filter not applied"
    SetColumnFilter sf, "flt", 0, Array("Alpha", Null)
    Check RowCount(sf.Form) = 3, "Header filter not combined with the page filter"
    Check GridColumnFiltered(sf, 0) And Not GridColumnFiltered(sf, 1), "Filtered columns reported wrongly"
    Check host.fltFilter0.BackStyle = 1 And host.fltFilter0.BackColor = Th(CLR_PRIMARY) And host.fltFilter0.ForeColor = Th(CLR_WHITE), "Filtered column's funnel not shown on a chip"
    Check host.fltFilter1.BackStyle = 0 And host.fltFilter1.ForeColor = Th(CLR_MUTED_FG), "Unfiltered column's funnel not muted"
    vals = GridColumnValues(sf, "flt", 1)
    Check UBound(vals) = 1 And vals(0) = 5 And vals(1) = 7, "Column values ignore the other filters, or not in sort-field order"
    vals = GridColumnValues(sf, "flt", 0)
    Check UBound(vals) = 2, "Column values narrowed by the column's own filter"
    SetGridFilter sf, ""
    Check RowCount(sf.Form) = 4, "Changing the page filter dropped the header filter, or blanks missed an empty value"
    SetColumnFilter sf, "flt", 1, Array(9)
    Check RowCount(sf.Form) = 1 And InStr(GridFilterWhere(sf), "[qty] IN (9)") > 0, "Second header filter not combined"
    GridSort host, "sfFlt", "flt", 1
    Check RowCount(sf.Form) = 1, "Sorting dropped the header filters"
    ' A Filter the page sets itself replaces the header filters: the next
    ' header call notices, and drops them and their funnels' color.
    sf.Form.Filter = "id = 2"
    sf.Form.FilterOn = True
    Check Not GridColumnFiltered(sf, 0), "Header filter kept after the page set Filter itself"
    Check host.fltFilter0.BackStyle = 0 And host.fltFilter0.ForeColor = Th(CLR_MUTED_FG), "Funnel still highlighted after its filter was replaced"
    SetColumnFilter sf, "flt", 0, Array("Alpha")
    Check RowCount(sf.Form) = 0, "Page's own Filter not kept as the page filter"
    ' The table is empty, but the column's kept value is still listed to clear.
    vals = GridColumnValues(sf, "flt", 0)
    Check UBound(vals) = 1 And vals(0) = "Alpha" And vals(1) = "Beta", "Column's own kept value not listed when other filters hide its rows"
    ClearGridFilters sf
    SetGridFilter sf, ""
    Check RowCount(sf.Form) = 6 And Len(GridFilterWhere(sf)) = 0, "Clearing every filter left rows hidden"
    SetColumnFilter sf, "flt", 0, Array("O'Brien")
    Check RowCount(sf.Form) = 1, "Value with a quote not matched"
    SetColumnFilter sf, "flt", 0, Empty
    Check RowCount(sf.Form) = 6 And Not GridColumnFiltered(sf, 0), "Removing a column's filter left it on"
    On Error Resume Next
    SetColumnFilter sf, "flt", 2, Array(1)
    Check Err.Number = 5, "Filtered a column with no funnel"
    On Error GoTo 0
    ' The funnels follow their captions when the table widens or columns hide.
    sf.Width = px(600)
    GridFit sf, "flt"
    Set grid = sf.Form
    Check host.fltFilter0.Left > host.flt0.Left And host.flt0.Left + host.flt0.Width <= host.fltFilter0.Left, "Funnel not kept right of its caption"
    Check host.fltFilter0.Left + host.fltFilter0.Width <= sf.Left + grid.txtCol1.Left + px(2), "Funnel not over its column after GridFit"
    SetGridColumns sf, "flt", "1,0h,2"
    Check Not host.fltFilter0.Visible And host.fltFilter1.Visible, "Hidden column's funnel still shows"
    SetGridColumns sf, "flt", DefaultLayout(3)
    Check host.fltFilter0.Visible, "Shown column's funnel stays hidden"
    DoCmd.Close acForm, "frmTestFilter", acSaveNo
    ' Help (modKitHelp): the topic file, the list, search and the help window.
    Dim helpPath As String, helpHtml As String, helpFrm As Form, helpErr As Long
    helpPath = root & "\help_test.txt"
    WriteTextFile helpPath, "Notes for the author." & vbCrLf & _
        "=== start | Getting Started" & vbCrLf & "Intro **bold** & <b>" & vbCrLf & "more text." & vbCrLf & vbCrLf & _
        "# Steps" & vbCrLf & "- one" & vbCrLf & "- two" & vbCrLf & _
        "=== detail | Detail" & vbCrLf & "Costs 5*2 [x] 100%." & vbCrLf & _
        "=== detail.roles | Roles" & vbCrLf & "Who can sign."
    LoadHelpTopics helpPath
    Check DCount("*", HELP_TABLE) = 3, "Help topics not loaded one per topic"
    Check DLookup("sort_order", HELP_TABLE, "topic_key='detail.roles'") = 3, "Help topics not in file order"
    Check InStr(DLookup("body", HELP_TABLE, "topic_key='start'"), "Notes") = 0, "Lines above the first topic loaded"
    Check HelpTitleFor("detail.roles") = "Detail  >  Roles", "Subtopic heading lacks its parent"
    helpHtml = HelpBodyHtml("start")
    Check InStr(helpHtml, "<div>Intro <strong>bold</strong> &amp; &lt;b&gt; more text.</div>") > 0, "Help paragraph not joined, escaped or bolded"
    Check InStr(helpHtml, "<font size=4><strong>Steps</strong></font>") > 0, "Help heading not drawn"
    Check InStr(helpHtml, "<ul><li>one</li><li>two</li></ul>") > 0 And Right$(helpHtml, 5) = "</ul>", "Help bullets not listed, or trailing spacer kept"
    Check InStr(HelpBodyHtml("start", "BOLD"), "BACKGROUND-COLOR:") > 0, "Search term not marked in the topic"
    Check HelpRows("") = 3 And HelpRows("sign") = 1 And HelpRows("Sign ") = 1, "Help search does not match text ignoring case"
    ' Literal matches: "*" in **bold** (start) and 5*2 (detail), "#" only in start's "# Steps".
    Check HelpRows("*") = 2 And HelpRows("5*2") = 1 And HelpRows("[x") = 1 And HelpRows("0%") = 1 And HelpRows("#") = 1, "Help search treats wildcards as wildcards"
    Check HelpRows("o'b") = 0 And HelpRows("zzz") = 0, "Help search matched missing text"
    Set rows = CurrentDb.OpenRecordset(HelpTopicsSql())
    rows.MoveLast
    Check Left$(rows.Fields(1), 6) = "      " And Left$(rows.Fields(0), 6) = "detail", "Subtopic not indented in the list"
    rows.Close
    WriteTextFile helpPath, "=== a | A" & vbCrLf & "=== A | Again"
    helpErr = 0: On Error Resume Next: LoadHelpTopics helpPath: helpErr = Err.Number: On Error GoTo 0
    Check helpErr = 5, "Repeated help key not refused"
    WriteTextFile helpPath, "=== a A"
    helpErr = 0: On Error Resume Next: LoadHelpTopics helpPath: helpErr = Err.Number: On Error GoTo 0
    Check helpErr = 5, "Help topic line without | not refused"
    WriteTextFile helpPath, "=== start | Getting Started" & vbCrLf & "Intro." & vbCrLf & "=== detail | Detail" & vbCrLf & _
        "=== detail.roles | Roles" & vbCrLf & "Who can sign."
    LoadHelpTopics helpPath
    Kill helpPath
    BuildHelpDialog
    DoCmd.OpenForm "dlgHelp", OpenArgs:="detail"
    Set helpFrm = Forms("dlgHelp")
    Check helpFrm.lstTopics = "detail" And helpFrm.lblHeader.Caption = "Detail", "Help window not open on its topic"
    Check helpFrm.Modal = False And helpFrm.lstTopics.ListCount = 3, "Help window modal, or topics not listed"
    Check HelpTopicOf(helpFrm) = "detail" And HelpTopicOf(Nothing) = "", "Form's HelpTopic not read"
    helpFrm.SetFocus
    Check HelpContextTopic() = "detail", "F1 in the help window changed its topic"
    KitHelp "nope"
    Check helpFrm.lstTopics = "start" And HelpCurrentTopic() = "start", "Unknown topic did not show the first topic"
    KitHelp "detail.roles"
    Check helpFrm.lblHeader.Caption = "Detail  >  Roles", "Open help did not switch topic"
    HelpSearch helpFrm, "sign"
    Check helpFrm.lstTopics.ListCount = 1 And Not helpFrm.lblNoMatch.Visible, "Help search did not narrow the list"
    Check InStr(helpFrm.txtBody, "BACKGROUND-COLOR:") > 0, "Help search not marked in the topic shown"
    HelpSearch helpFrm, "zzz"
    Check helpFrm.lstTopics.ListCount = 0 And helpFrm.lblNoMatch.Visible And Not helpFrm.lstTopics.Visible, "No-match note not shown, or the empty list covers it"
    HelpSearch helpFrm, "intro"
    Check HelpCurrentTopic() = "start" And helpFrm.lstTopics.Visible, "Search did not show the first match when the shown topic dropped out"
    HelpShowFirstMatch helpFrm
    Check HelpCurrentTopic() = "start", "Enter did not show the first match"
    HelpSearch helpFrm, ""
    Check helpFrm.lstTopics.ListCount = 3 And Not helpFrm.lblNoMatch.Visible And HelpCurrentTopic() = "start", "Clearing the search did not list every topic"
    DoCmd.Close acForm, "dlgHelp", acSaveNo
    DoCmd.OpenForm "frmTestFilter"
    Check HelpTopicOf(Forms("frmTestFilter")) = "", "Form without HelpTopic named a topic"
    DoCmd.Close acForm, "frmTestFilter", acSaveNo
    BuildHelpKey
    Check MacroExists("AutoKeys"), "F1 macro not written"
    ' Locks and write conflicts (modKitLocks).
    Dim lockTry As Long, conflictFrm As Form, saved As Boolean
    Check IsLockError(3218) And IsLockError(3260) And IsLockError(3188) And Not IsLockError(3022), "Lock errors not told apart"
    lockTry = 0
    Check LockWait(3218, lockTry) And lockTry = 1, "Lock error not retried"
    Check Not LockWait(3022, lockTry) And lockTry = 1, "Other error retried"
    lockTry = LOCK_TRIES
    Check Not LockWait(3260, lockTry), "Retried past LOCK_TRIES"
    CurrentDb.Execute "CREATE TABLE test_conflict (id LONG PRIMARY KEY, name TEXT(60), qty LONG)"
    Check KitExecute("INSERT INTO test_conflict VALUES (1, 'start', 1)") = 1, "KitExecute did not report the rows it changed"
    Check KitExecute("UPDATE test_conflict SET qty = 2 WHERE id = 99") = 0, "KitExecute counted rows it did not change"
    lockTry = 0: On Error Resume Next: KitExecute "UPDATE no_such_table SET x = 1": lockTry = Err.Number: On Error GoTo 0
    Check lockTry <> 0, "KitExecute hid a failing query"
    NewForm "test_conflict", 0, px(300), px(80)
    Ev F, "OnCurrent", "BeforeUpdate", "AfterUpdate", "OnError"
    Txt "txtName", "name", 0, 0, px(200), px(30)
    Txt "txtQty", "qty", 0, px(40), px(80), px(30)
    SaveForm "frmConflict", "testConflict"
    ' Theirs wins (No): another save lands while the form is changed; the
    ' form's save is refused and it shows the saved record.
    DoCmd.OpenForm "frmConflict"
    Set conflictFrm = Forms("frmConflict")
    conflictFrm.txtName = "mine"
    CurrentDb.Execute "UPDATE test_conflict SET name = 'theirs' WHERE id = 1", dbFailOnError
    SetConflictAnswer vbNo
    saved = KitSaveRecord(conflictFrm)
    Check Not saved And Not conflictFrm.Dirty And conflictFrm.txtName = "theirs", "Write conflict: No did not show the other user's record"
    Check DLookup("name", "test_conflict", "id=1") = "theirs", "Write conflict: No overwrote the other user's save"
    ' Mine wins (Yes): the change goes back on top of their record and saves.
    conflictFrm.txtName = "mine again"
    CurrentDb.Execute "UPDATE test_conflict SET qty = 7 WHERE id = 1", dbFailOnError
    SetConflictAnswer vbYes
    saved = KitSaveRecord(conflictFrm)
    Check Not saved And conflictFrm.Dirty And conflictFrm.txtName = "mine again" And conflictFrm.txtQty = 7, "Write conflict: Yes did not keep the change over the saved record"
    Check KitSaveRecord(conflictFrm), "Record not saved after keeping the change"
    Check DLookup("name", "test_conflict", "id=1") = "mine again" And DLookup("qty", "test_conflict", "id=1") = 7, "Kept change lost the other user's field"
    Check KitSaveRecord(conflictFrm), "Unchanged record not reported saved"
    ' After a save of their own, the user's next edit is compared with what
    ' they saved: the other user's later qty survives Yes.
    conflictFrm.txtQty = 5
    Check KitSaveRecord(conflictFrm), "Own change not saved"
    conflictFrm.txtName = "third"
    CurrentDb.Execute "UPDATE test_conflict SET qty = 9 WHERE id = 1", dbFailOnError
    SetConflictAnswer vbYes
    saved = KitSaveRecord(conflictFrm)
    Check Not saved And conflictFrm.txtQty = 9 And KitSaveRecord(conflictFrm), "Write conflict: Yes after an own save put a stale field back"
    Check DLookup("name", "test_conflict", "id=1") = "third" And DLookup("qty", "test_conflict", "id=1") = 9, "Write conflict after an own save lost a field"
    SetConflictAnswer 0
    DoCmd.Close acForm, "frmConflict", acSaveNo
    ' Current user and record stamps (modKitUser).
    Dim sf2 As Form, added As Long
    Check Len(WinUser()) > 0 And StrComp(WinUser(), Environ$("USERNAME"), vbTextCompare) = 0, "WinUser is not the Windows account"
    Check Len(WinComputer()) > 0 And StrComp(WinComputer(), Environ$("COMPUTERNAME"), vbTextCompare) = 0, "WinComputer is not this computer"
    Check StrComp(WinAccount(), Environ$("USERDOMAIN") & "\" & Environ$("USERNAME"), vbTextCompare) = 0, "WinAccount is not DOMAIN\user"
    Check WinSid() Like "S-1-5-*-*" And WinSid() = WinSid(), "WinSid is not this account's SID"
    TempVars("Username") = Null
    Check KitUser() = WinUser(), "KitUser without a sign-in is not the Windows account"
    TempVars("Username") = "O'Neil"
    Check KitUser() = "O'Neil", "KitUser ignores the signed-in user"
    Check StampSql() = "updated_by='O''Neil', updated_at=Now()", "StampSql not quoted"
    CurrentDb.Execute "CREATE TABLE test_stamp (id COUNTER PRIMARY KEY, name TEXT(60), created_by TEXT(150))"
    added = AddStampFields("test_stamp")
    Check added = 3 And AddStampFields("test_stamp") = 0, "AddStampFields did not add only the missing columns"
    CurrentDb.Execute "INSERT INTO test_stamp (name, " & StampColumns() & ") VALUES ('sql', " & StampValues() & ")", dbFailOnError
    Check DLookup("created_by", "test_stamp", "name='sql'") = "O'Neil" And Not IsNull(DLookup("updated_at", "test_stamp", "name='sql'")), "StampValues not written"
    NewForm "test_stamp", 0, px(300), px(40)
    F.AllowAdditions = True
    Ev F, "BeforeUpdate"
    Txt "txtName", "name", 0, 0, px(200), px(30)
    SaveForm "frmStamp", "testStamp"
    DoCmd.OpenForm "frmStamp"
    Set sf2 = Forms("frmStamp")
    DoCmd.GoToRecord acForm, "frmStamp", acNewRec
    sf2.txtName = "form"
    sf2.Dirty = False
    Check DLookup("created_by", "test_stamp", "name='form'") = "O'Neil" And DLookup("updated_by", "test_stamp", "name='form'") = "O'Neil", "New record not stamped"
    CurrentDb.Execute "UPDATE test_stamp SET created_by='first', updated_by='first' WHERE name='form'", dbFailOnError
    sf2.Requery
    sf2.Recordset.FindFirst "name='form'"
    sf2.txtName = "form2"
    TempVars("Username") = "second"
    sf2.Dirty = False
    Check DLookup("created_by", "test_stamp", "name='form2'") = "first" And DLookup("updated_by", "test_stamp", "name='form2'") = "second", "Changed record stamp wrong"
    DoCmd.Close acForm, "frmStamp", acSaveNo
    TempVars("Username") = Null
    ' Keyboard shortcuts (modKitKeys).
    Dim kf As Form
    NewForm "SELECT * FROM test_rows ORDER BY id", 0, px(400), px(80)
    Txt "txtDesc", "description", 0, 0, px(200), px(30)
    Txt "txtSearch", "", px(210), 0, px(150), px(30)
    SaveForm "frmKeys", "testKeys"
    DoCmd.OpenForm "frmKeys"
    Set kf = Forms("frmKeys")
    kf.SetFocus
    kf.txtDesc.SetFocus
    Check KitShortcut(SHORTCUT_NEW) And kf.lastAction = "new", "Form's own Shortcut not asked"
    kf.txtDesc = "Saved by Ctrl+S"
    Check kf.Dirty, "Test record not changed"
    Check KitShortcut(SHORTCUT_SAVE) And Not kf.Dirty And kf.lastAction = "save", "Ctrl+S did not save the changed record"
    Check DCount("*", "test_rows", "description='Saved by Ctrl+S'") = 1, "Ctrl+S save not in the table"
    Check KitShortcut(SHORTCUT_FIND) And Screen.ActiveControl.Name = "txtSearch", "Ctrl+F did not go to the search box"
    Check Not KitShortcut("nothing"), "Unknown shortcut reported done"
    DoCmd.Close acForm, "frmKeys", acSaveNo
    BuildShortcutKeys
    Check MacroExists("AutoKeys"), "Shortcut macro not written"
    g_Silent = False
    TestKit = checks & " kit checks passed"
End Function

' Rows a form shows (its recordset clone read to the end).
Private Function RowCount(g As Form) As Long
    With g.RecordsetClone
        If .BOF And .EOF Then Exit Function
        .MoveLast
        RowCount = .RecordCount
    End With
End Function

' True when frm has a control named nm.
Private Function HasControl(frm As Form, ByVal nm As String) As Boolean
    Dim c As Control
    On Error Resume Next
    Set c = frm.Controls(nm)
    HasControl = Not c Is Nothing
End Function

' Topics the help search finds for term.
Private Function HelpRows(ByVal term As String) As Long
    With CurrentDb.OpenRecordset(HelpTopicsSql(term))
        If Not (.BOF And .EOF) Then .MoveLast: HelpRows = .RecordCount
        .Close
    End With
End Function

' True when the database has macro nm.
Private Function MacroExists(ByVal nm As String) As Boolean
    Dim m As AccessObject
    For Each m In CurrentProject.AllMacros
        If m.Name = nm Then MacroExists = True
    Next
End Function
