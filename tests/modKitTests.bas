Option Compare Database
Option Explicit
Private Declare PtrSafe Function IsWindowEnabled Lib "user32" (ByVal hwnd As LongPtr) As Long
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
    BuildBusy
    BuildMenuDialog
    BuildPopoverDialog
    BuildPickDialog
    BuildPaletteDialog
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
    PickBox "pkTest", "", "SELECT id, description, category FROM test_rows", px(1450), px(720), px(240)
    SaveForm "frmTest", "testHost"
    DoCmd.OpenForm "frmTest"
    Set host = Forms("frmTest"): Set grid = host.sfRows.Form
    ' Captions are as wide as their cells; a filling last column's reaches the table's edge.
    Check host.fillHead0.Width = grid.txtCol0.Width And host.fillHead0.Left = host.sfRows.Left + grid.txtCol0.Left, "Caption is not over its cells"
    ' Designed width (the Tag; an anchored control's Width follows the window): 600 - 8 - 160 - 8 - 8.
    Check Val(Split(host.fillHead1.Tag, "|")(2)) = px(416) And host.fillHead1.HorizontalAnchor = 2, "Last caption does not reach the table's edge"
    Check host.fitHead1.Width = px(292) And host.fitHead1.HorizontalAnchor = 0, "Fit table's last caption is not its column's width"
    ' A right-click on a row goes to the host's GridRightClick, not Access's own menu.
    Check grid.txtCol0.OnMouseUp = "=RowMouseUp()" And grid.txtCol0.OnMouseDown = "=RowMouseDown()" And Not grid.ShortcutMenu, _
        "Table rows don't open the row menu on a right-click"
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
    Check host.segTest_t0.FormatConditions.Count = 1 And InStr(host.segTest_t0.FormatConditions(0).Expression1, """day""") > 0 _
        And InStr(host.segTest_t2.FormatConditions(0).Expression1, """month""") > 0, "Segmented caption is not bold for its own option"
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
    ' Keyboard reach: the close mark has a KeyButton named for screen readers that
    ' marks the mark while focused and hides with the alert.
    Check host.alTest_xk.Visible And host.alTest_xk.TabStop And host.alTest_xk.Caption = "Dismiss", "Alert close mark has no keyboard button"
    host.alTest_xk.SetFocus
    Check host.alTest_x.ForeColor = CLR_PRIMARY, "Focused close mark is not marked"
    host.txtFocusSink.SetFocus
    Check host.alTest_x.ForeColor = CLR_MUTED_FG, "Close mark kept its focus mark"
    HideAlert host, "alTest"
    Check Not host.alTest_xk.Visible, "Hidden alert left its keyboard button"
    ShowAlert host, "alTest", "Back.", "info"
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
    ' Keyboard reach: every crumb but the current page has a KeyButton over it,
    ' named with its caption; focus underlines the crumb.
    SetBreadcrumb host, "bc", "home|Home", "items|Items", "i42|Pump overhaul"
    Check host.bc0_k.Visible And host.bc1_k.Visible And Not host.bc2_k.Visible And Not host.bc3_k.Visible, "Breadcrumb keyboard buttons differ"
    Check host.bc1_k.Left = host.bc1.Left And host.bc1_k.Width = host.bc1.Width And host.bc1_k.Caption = "Items", "Crumb keyboard button is not over its crumb"
    host.bc0_k.SetFocus
    Check host.bc0.ForeColor = CLR_PRIMARY And host.bc0.FontUnderline, "Focused crumb is not marked"
    host.txtFocusSink.SetFocus
    Check host.bc0.ForeColor = CLR_MUTED_FG And Not host.bc0.FontUnderline, "Crumb kept its focus mark"
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
    ' Keyboard reach: only done steps take a Tab stop; names say the state.
    SetStep host, "st", 3
    Check host.st_k1.TabStop And host.st_k2.TabStop And Not host.st_k3.TabStop, "Step tab stops are not the done steps"
    Check host.st_k3.Caption = "Step 3 of 3, Done, current" And host.st_k1.Caption = "Step 1 of 3, Details, done", "Step button names differ"
    Check host.st_r3.Visible = False And host.st_f3.Visible, "Current step is not filled"
    SetStep host, "st", 1
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
    ' Accordion headers: names carry the state; focus underlines the title.
    Check host.acc_h2.Caption = "Advanced, expanded" And host.acc_h1.Caption = "General, collapsed", "Accordion header names differ"
    host.acc_h1.SetFocus
    Check host.acc_t1.ForeColor = CLR_PRIMARY And host.acc_t1.FontUnderline, "Focused accordion header is not marked"
    host.txtFocusSink.SetFocus
    Check host.acc_t1.ForeColor = CLR_FG And Not host.acc_t1.FontUnderline, "Accordion header kept its focus mark"
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
    SetSearch host, "srch", "x"
    Check host.srch_xk.Visible And host.srch_xk.Caption = "Clear search", "Search clear mark has no keyboard button"
    SetSearch host, "srch", ""
    Check Not host.srch_xk.Visible, "Empty search kept its clear button"
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
    Check host.rt_s5.ForeColor = CLR_MUTED_FG And host.rt_b3.Caption = "3 stars", "Rating empty star color or names differ"
    host.rt_b2.SetFocus
    Check host.rt_ring.Visible, "Focused rating shows no ring"
    host.txtFocusSink.SetFocus
    Check Not host.rt_ring.Visible, "Rating kept its ring"
    Check host.segTest_b1.Caption = "Week" And host.sl_b3.Caption = "30", "Segmented or slider buttons are unnamed"
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
    host.tglTest.ControlTipText = "Due date tip"
    Check KitPopover(host.tglTest, "Due date", "Work is due then."), "Popover did not open"
    Check host.tglTest.ControlTipText = "", "Anchor tooltip not cleared while the popover shows"
    Check IsFormOpen("dlgPopover"), "Popover form is not open"
    With Forms("dlgPopover")
        Check .Controls("lblTitle").Caption = "Due date" And .Controls("lblTitle").Visible, "Popover title differs"
        Check Abs(.InsideHeight - (px(12) + px(24) + POPOVER_LINE_H + px(12))) <= 15, "Popover is not sized to one line (" & .InsideHeight & ")"
    End With
    Dim pk As Integer
    pk = vbKeyEscape: PopoverKey Forms("dlgPopover"), pk
    Check Not IsFormOpen("dlgPopover"), "Esc did not close the popover"
    Check host.tglTest.ControlTipText = "Due date tip", "Anchor tooltip not restored when the popover closed"
    host.tglTest.ControlTipText = ""
    Check KitPopover(host.tglTest, "", "No title"), "Untitled popover did not open"
    Check Not Forms("dlgPopover").Controls("lblTitle").Visible And Forms("dlgPopover").Controls("lblText").Top = px(12), "Untitled popover kept its title"
    MenuClose Forms("dlgPopover")
    DoCmd.SelectObject acForm, "frmTest"
    TestKitMenu host
    Dim key As Integer
    ' Searchable picker: the list filters as the user types (wildcards taken
    ' literally) without taking the focus; Down and Enter pick; Esc puts the
    ' box back; leaving it empty clears the pick.
    Check PickLikeText("50%_[a]*?#") = "50%_[[]a][*][?][#]", "Picker took a wildcard as a wildcard"
    Check InStr(PickSql(host.pkTest, True), "q.[description] LIKE [pickText] OR q.[category] LIKE [pickText]") > 0, _
        "Picker doesn't search every column after the key"
    Check InStr(PickSql(host.pkTest, False), " WHERE ") = 0, "Picker filters with nothing typed"
    host.pkTest.SetFocus
    host.pkTest.Text = "ription 1"
    KitPickChanged host, "pkTest"
    Check IsFormOpen("dlgPick"), "Typing in the picker didn't open its list"
    With Forms("dlgPick")
        Check .Controls("lblMain0").Caption = "Description 1" And Left$(.Controls("lblSub0").Caption, 8) = "Long cat", _
            "Picker's first row differs"
        Check .Controls("lblMain7").Visible And .Controls("lblMain7").Caption = "Description 16", "Picker doesn't list 8 matches in order"
        Check .Controls("boxRow0").Visible, "Picker didn't light its first match"
    End With
    Check Screen.ActiveForm.Name = "frmTest" And Screen.ActiveControl.Name = "pkTest", "Picker list took the focus"
    key = vbKeyDown: KitPickKey host.pkTest, key
    Check key = 0 And Forms("dlgPick").Controls("boxRow1").Visible And Not Forms("dlgPick").Controls("boxRow0").Visible, _
        "Down didn't move the picker's light"
    key = vbKeyReturn: KitPickKey host.pkTest, key
    Check Not IsFormOpen("dlgPick") And host.pkTestKey = 10 And host.pkTest.Text = "Description 10", "Enter didn't pick the lit row"
    host.pkTest.Text = "zzz"
    KitPickChanged host, "pkTest"
    Check Forms("dlgPick").Controls("lblMain0").Caption = "No matches" And Not Forms("dlgPick").Controls("lblMain1").Visible, _
        "Picker with no matches doesn't say so"
    key = vbKeyReturn: KitPickKey host.pkTest, key
    Check IsFormOpen("dlgPick") And host.pkTestKey = 10, "Enter picked from an empty list"
    key = vbKeyEscape: KitPickKey host.pkTest, key
    Check Not IsFormOpen("dlgPick") And host.pkTest.Text = "Description 10" And host.pkTestKey = 10, "Esc didn't put the picker back"
    host.pkTest.Text = "[x"
    KitPickChanged host, "pkTest"                       ' a bracket is text, not a pattern error
    Check Forms("dlgPick").Controls("lblMain0").Caption = "No matches", "Picker took [ as a pattern"
    host.pkTest.Text = ""
    KitPickChanged host, "pkTest"
    Check Forms("dlgPick").Controls("lblMain0").Caption = "Description 1", "Empty picker doesn't list from the top"
    host.ringA.SetFocus
    PickTick Forms("dlgPick")
    Check Not IsFormOpen("dlgPick") And IsNull(host.pkTestKey), "Leaving the picker empty didn't clear the pick"
    host.pkTestKey = 3
    KitPickShow host.pkTest
    Check host.pkTest = "Description 3", "Picker doesn't show the picked row's text"
    On Error Resume Next
    PickBox "pkBad", "", "SELECT id, description FROM test_rows ORDER BY id", 0, 0, px(100): badName = Err.Number: Err.Clear
    On Error GoTo 0
    Check badName = 5, "PickBox took a rowSql with ORDER BY"
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
    KitTakeRefused
    KitRefuseSave
    Check KitTakeRefused() And Not KitTakeRefused(), "A refused save isn't noted once"
    Check host.reqA.BorderColor = Th(CLR_INPUT_BORDER), "Filled required field kept its outline"
    DoCmd.OpenForm "dlgToast", WindowMode:=acHidden     ' as KitToast does, but errors raise here
    Forms("dlgToast").ShowToast "Saved R&D", "success", 3000
    Check Forms("dlgToast").lblMsg.Caption = "Saved R&&D", "Toast took & as an access key"
    KitToast "Saved", "success"
    Check Forms("dlgToast").lblMsg.Caption = "Saved", "Toast lost its message"
    Check Screen.ActiveForm.Name = "frmTest", "Toast took the focus"
    DoCmd.Close acForm, "dlgToast", acSaveNo
    ' Busy pop-up: nested messages, the dim, the focus, and closing.
    KitBusy "Linking R&D"
    Check IsFormOpen("dlgBusy"), "KitBusy didn't open dlgBusy"
    Check Forms("dlgBusy").lblMsg.Caption = "Linking R&&D", "Busy pop-up took & as an access key"
    Check AppDimmed(), "KitBusy didn't dim the app"
    Check Screen.ActiveForm.Name = "frmTest", "Busy pop-up took the focus"
    Check IsWindowEnabled(Application.hWndAccessApp) = 0 And IsWindowEnabled(Forms("dlgBusy").hwnd) <> 0, _
        "Busy pop-up didn't block clicks and keys behind it"
    KitBusy "Exporting"
    Check KitBusyDepth() = 2 And Forms("dlgBusy").lblMsg.Caption = "Exporting", "Nested KitBusy didn't show its message"
    KitBusyEnd
    Check KitBusyDepth() = 1 And Forms("dlgBusy").lblMsg.Caption = "Linking R&&D", "Outer busy message didn't come back"
    KitBusy String(120, "x")
    Check Forms("dlgBusy").boxTrack.Top > Forms("dlgBusy").lblMsg.Top + 272, "Long busy message didn't move the track down"
    KitBusyEnd
    Check Abs(Forms("dlgBusy").boxTrack.Top + Forms("dlgBusy").boxTrack.Height + Forms("dlgBusy").lblMsg.Top _
        - Forms("dlgBusy").Section(acDetail).Height) <= 15, _
        "Busy pop-up padding below the track differs from above the message"
    KitBusyEnd
    Check KitBusyDepth() = 0 And Not IsFormOpen("dlgBusy"), "Last KitBusyEnd didn't close dlgBusy"
    Check Not AppDimmed(), "Last KitBusyEnd left the app dimmed"
    Check IsWindowEnabled(Application.hWndAccessApp) <> 0 And IsWindowEnabled(Forms("frmTest").hwnd) <> 0, _
        "Last KitBusyEnd left the app blocked"
    KitBusyEnd                                          ' one too many: nothing happens
    Check KitBusyDepth() = 0, "Extra KitBusyEnd changed the busy count"
    Check SaveDetail(host, "sfRows"), "Unchanged detail panel reported a refused save"
    host.sfRefuse.Form.txtDesc = "Changed"
    Check Not SaveDetail(host, "sfRefuse"), "Refused detail save reported as saved"
    Check host.sfRefuse.Form.Dirty, "Refused detail save dropped the change"
    Check Not ShowDetail(host, "sfRefuse", "subRefuse", "id=2"), "ShowDetail moved on after a refused save"
    host.sfRefuse.Form.Undo
    DoCmd.Close acForm, "frmTest", acSaveNo
    TestTrustCheck
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
    BuildSplash "frmMain", "Subtitle", , True
    DoCmd.OpenForm "frmSplash", acDesign
    Check Forms("frmSplash").Tag = "frmMain" & SPLASH_HIDDEN, "BuildSplash openHidden not saved in the splash's Tag"
    DoCmd.Close acForm, "frmSplash", acSaveNo
    BuildSplash "frmMain", "Subtitle"
    ' A table on a tab page: its subform's Parent is the Page, and the kit's
    ' table helpers (GridFit, funnels, column layouts) use OwnerForm instead.
    NewForm "", 0, px(300), px(200)
    TabCtl "tabOwner", 0, 0, px(300), px(200), "pgOwner|Owner"
    Set c = SubF("sfOwner", "", px(8), px(40), px(200), px(100), , "pgOwner")
    Check TypeOf c.Parent Is Page, "Subform on a tab page not parented by the page"
    Check OwnerForm(c).Name = F.Name, "OwnerForm did not find the form behind a tab page"
    DoCmd.Close acForm, F.Name, acSaveNo
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
    ' Literal matches: "*" only in 5*2 (detail); the ** around bold and the
    ' "# " and "- " markers in start aren't text the reader sees.
    Check HelpRows("*") = 1 And HelpRows("5*2") = 1 And HelpRows("[x") = 1 And HelpRows("0%") = 1, "Help search treats wildcards as wildcards"
    Check HelpRows("#") = 0 And HelpRows("- one") = 0 And HelpRows("Steps") = 1 And HelpRows("bold &") = 1, "Help search matches the markup, not the text"
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
    ' A form with no Shortcut of its own (no module at all), through the macro
    ' itself as a key press runs it.
    NewForm "", 0, px(400), px(80)
    Txt "txtOther", "", 0, 0, px(150), px(30)
    Txt "txtSearch", "", px(210), 0, px(150), px(30)
    SaveForm "frmKeysPlain"
    DoCmd.OpenForm "frmKeysPlain"
    Forms("frmKeysPlain").SetFocus
    Forms("frmKeysPlain").txtOther.SetFocus
    DoCmd.RunMacro "AutoKeys.^F"
    Check Screen.ActiveControl.Name = "txtSearch", "Ctrl+F (the macro) did not reach the search box of a form without Shortcut"
    DoCmd.RunMacro "AutoKeys.^S"
    DoCmd.RunMacro "AutoKeys.^N"
    DoCmd.Close acForm, "frmKeysPlain", acSaveNo
    ' A form with no search box: Ctrl+F does nothing, without an error.
    NewForm "", 0, px(400), px(80)
    Txt "txtOther", "", 0, 0, px(150), px(30)
    SaveForm "frmKeysNoSearch"
    DoCmd.OpenForm "frmKeysNoSearch"
    Forms("frmKeysNoSearch").SetFocus
    Forms("frmKeysNoSearch").txtOther.SetFocus
    DoCmd.RunMacro "AutoKeys.^F"
    DoCmd.RunMacro "AutoKeys.^S"
    Check Screen.ActiveControl.Name = "txtOther", "Ctrl+F moved the focus on a form without a search box"
    DoCmd.Close acForm, "frmKeysNoSearch", acSaveNo
    ' App shell (modKitShell): the first page, swapping pages from the timer,
    ' a page that refuses to be left, the hooks, hiding an item, collapsing.
    Dim shellFrm As Form, shellErr As Long, shellTop As Long, shellW As Long
    On Error Resume Next
    ShellBuild "Shell", "", Array("a|Alpha|subRing"), Empty, px(500), px(400): shellErr = Err.Number: Err.Clear
    ShellBuild "Shell", "", Array("a|Alpha|subRing|"), Empty, px(500), px(400): badName = Err.Number: Err.Clear
    ShellBuild "Shell", "", Array("a|Alpha|subRing|E80F"), Array("help|Help|E897|KitHelp()"), px(500), px(400): badScale = Err.Number: Err.Clear
    On Error GoTo 0
    Check shellErr = 5 And badName = 5 And badScale = 5, "ShellBuild took a bad page or link"
    NewForm "", 0, px(200), px(60)
    SaveForm "subLeave", "testLeave"
    ShellBuild "Shell", "Test", Array("a|Alpha|subRing|E80F", "b|Beta|subLeave|E8A5", "c|Gamma|subRing|E8FD"), _
        Array("help|Help|E897|=KitHelp()"), px(500), px(400)
    SaveForm "frmMain", "testShell"
    TempVars("ShellLeave") = True
    DoCmd.OpenForm "frmMain"
    Set shellFrm = Forms("frmMain")
    Check ShellPageKey() = "a" And shellFrm.sfPage.SourceObject = "subRing", "Shell didn't open on its first page"
    Check shellFrm.nb_a.BackStyle = 1 And shellFrm.nb_b.BackStyle = 0 And shellFrm.nt_a.FontWeight = 600 And shellFrm.nt_b.FontWeight = 400, _
        "Shell didn't light its page"
    Check shellFrm.nb_help.VerticalAnchor = 1 And shellFrm.nt_help.OnClick = "=KitHelp()", "Shell link isn't at the foot of the sidebar"
    ShellGo "b"
    Check shellFrm.TimerInterval = 10 And ShellPageKey() = "a", "ShellGo didn't wait for the calling event"
    ShellTimer shellFrm
    Check ShellPageKey() = "b" And shellFrm.sfPage.SourceObject = "subLeave" And shellFrm.nb_b.BackStyle = 1 And shellFrm.nb_a.BackStyle = 0, _
        "Shell didn't swap pages"
    Check TempVars("ShellShown") = "b", "Shell didn't call PageShown"
    TempVars("ShellLeave") = False
    ShellGo "c"
    Check shellFrm.TimerInterval = 0 And ShellPageKey() = "b" And Not shellFrm.CanLeavePage(), "Shell left a page that refused"
    TempVars("ShellLeave") = True
    TempVars("ShellRefuse") = True
    ShellGo "c"
    Check shellFrm.TimerInterval = 0 And ShellPageKey() = "b" And Not KitTakeRefused(), "Shell left a page whose save was refused"
    TempVars("ShellRefuse") = False
    ShellShow shellFrm, "secret"
    Check ShellPageKey() = "c", "Shell didn't follow RoutePage"
    ShellShow shellFrm, "detail"
    Check ShellPageKey() = "detail" And shellFrm.sfPage.SourceObject = "subLeave" And shellFrm.nb_c.BackStyle = 0, _
        "Shell didn't show a page from PageForm"
    ShellShow shellFrm, "nope"
    Check ShellPageKey() = "a" And shellFrm.sfPage.SourceObject = "subRing", "Shell didn't fall back to its first page"
    shellTop = shellFrm.nb_b.Top
    ShellShowItem shellFrm, "b", False
    Check Not shellFrm.nt_b.Visible And Not shellFrm.nb_b.Visible And shellFrm.nb_c.Top = shellTop And shellFrm.ni_c.Top = shellTop _
        And shellFrm.nt_c.Top = shellTop, "Hiding a sidebar item left a gap"
    ShellShowItem shellFrm, "b", True
    Check shellFrm.nt_b.Visible And shellFrm.nb_c.Top = shellTop + px(38), "Showing a sidebar item didn't make room"
    shellW = shellFrm.sfPage.Width
    ShellCollapse shellFrm, True, False
    Check shellFrm.boxSide.Width = SHELL_RAIL_W And shellFrm.sfPage.Left = SHELL_RAIL_W And shellFrm.sfPage.Width = shellW + SHELL_W - SHELL_RAIL_W, _
        "Collapsed sidebar didn't give its width to the page"
    Check Not shellFrm.nt_a.Visible And Not shellFrm.nt_help.Visible And Not shellFrm.lblBrand.Visible And shellFrm.ni_a.Visible _
        And shellFrm.nb_a.Width = px(44) And shellFrm.ni_help.Left = px(8) And shellFrm.lnkCollapse.Left = px(14), "Collapsed sidebar isn't a rail of glyphs"
    ShellCollapse shellFrm, False, False
    Check shellFrm.boxSide.Width = SHELL_W And shellFrm.sfPage.Left = SHELL_W And shellFrm.sfPage.Width = shellW And shellFrm.nt_a.Visible _
        And shellFrm.lblBrand.Visible And shellFrm.nb_a.Width = px(176) And shellFrm.ni_a.Left = px(20), "Opened sidebar didn't come back"
    ' Command palette (Ctrl+K): the shell's pages, the recent records, then its
    ' commands; typing narrows the list; the keys move; Enter goes.
    On Error Resume Next: DeleteSetting APP_NAME, PaletteRecentKey(): On Error GoTo 0
    KitRecent "Pump overhaul", "ShellGo(""c"")", "E8A5"
    KitRecent "Valve check", "ShellGo(""b"")"
    KitRecent "Pump overhaul", "ShellGo(""c"")", "E8A5"
    KitPalette
    Check IsFormOpen("dlgPalette"), "Ctrl+K didn't open the palette"
    Check Screen.ActiveForm.Name = "dlgPalette" And Screen.ActiveControl.Name = "txtSearch", "Palette didn't take the focus"
    Check PaletteRows() = 6 And PaletteRowText(0) = "Alpha" And PaletteRowText(2) = "Gamma" And PaletteRowText(3) = "Pump overhaul" _
        And PaletteRowText(4) = "Valve check" And PaletteRowText(5) = "Help", "Palette doesn't list pages, recent records (newest first, once), then commands"
    With Forms("dlgPalette")
        Check .lblSub0.Caption = "Page" And .lblSub3.Caption = "Recent" And .lblSub5.Caption = "Command" And .lblIcon3.Caption = ChrW(&HE8A5), _
            "Palette rows don't say what they are"
        Check PaletteHot() = 0 And .boxRow0.Visible And .barHot.Visible And .barHot.Top = .boxRow0.Top, "Palette didn't light its first row"
        Check .InsideWidth = px(PALETTE_W), "Palette isn't its design width"
        PaletteShow Forms("dlgPalette"), "A"
        Check PaletteRows() = 5 And PaletteRowText(0) = "Alpha" And PaletteRowText(1) = "Beta", "Palette doesn't put names starting with the text first"
        PaletteShow Forms("dlgPalette"), "check  VALVE"
        Check PaletteRows() = 1 And PaletteRowText(0) = "Valve check", "Palette doesn't match every word"
        PaletteShow Forms("dlgPalette"), "zzz"
        Check PaletteRows() = 0 And PaletteHot() = -1 And .lblEmpty.Visible And Not .lblCap0.Visible, "Palette shows no empty state"
        Check .InsideHeight = px(64 + 36 + 8 + 32), "Palette didn't shrink to its rows"
        PaletteShow Forms("dlgPalette"), ""
        key = vbKeyDown: PaletteKeyDown Forms("dlgPalette"), key
        key = vbKeyDown: PaletteKeyDown Forms("dlgPalette"), key
        Check PaletteHot() = 2 And .boxRow2.Visible And Not .boxRow0.Visible And .barHot.Top = .boxRow2.Top, "Down didn't move the light"
        For i = 1 To 3
            key = vbKeyUp: PaletteKeyDown Forms("dlgPalette"), key
        Next
        Check PaletteHot() = 5, "Up didn't go round to the last row"
        PaletteShow Forms("dlgPalette"), "ga"
    End With
    key = vbKeyReturn: PaletteKeyDown Forms("dlgPalette"), key
    Check Not IsFormOpen("dlgPalette") And shellFrm.TimerInterval = 10, "Enter in the palette didn't go to the page"
    ShellTimer shellFrm
    Check ShellPageKey() = "c", "Palette went to the wrong page"
    KitPalette
    key = vbKeyEscape: PaletteKeyDown Forms("dlgPalette"), key
    Check Not IsFormOpen("dlgPalette"), "Esc didn't close the palette"
    KitPalette
    shellFrm.SetFocus
    If IsFormOpen("dlgPalette") Then PaletteTick Forms("dlgPalette")
    Check Not IsFormOpen("dlgPalette"), "A click on the app window didn't close the palette"
    KitPalette
    KitPalette
    Check Not IsFormOpen("dlgPalette"), "Ctrl+K didn't close the open palette"
    DeleteSetting APP_NAME, PaletteRecentKey()
    DoCmd.Close acForm, "frmMain", acSaveNo
    Check Not IsFormOpen("frmMain"), "Shell didn't close"
    ' A frmMain with none of the optional hooks still opens and swaps pages.
    ShellBuild "Shell", "Test", Array("a|Alpha|subRing|E80F", "b|Beta|subLeave|E8A5"), Empty, px(500), px(400)
    SaveForm "frmMain", "testShellBare"
    DoCmd.OpenForm "frmMain"
    Set shellFrm = Forms("frmMain")
    ShellGo "b"
    ShellTimer shellFrm
    Check ShellPageKey() = "b" And shellFrm.sfPage.SourceObject = "subLeave", "Shell without hooks didn't swap pages"
    DoCmd.Close acForm, "frmMain", acSaveNo
    ' File pickers (modKitFiles): the parts that need no window.
    Dim fileErr As Long
    Check FileFilterText("CSV file|*.csv|Text|*.txt;*.log") = "CSV file (*.csv)" & vbNullChar & "*.csv" & vbNullChar & _
        "Text (*.txt;*.log)" & vbNullChar & "*.txt;*.log" & vbNullChar & vbNullChar, "File filter not in Windows' form"
    fileErr = 0: On Error Resume Next: FileFilterText "CSV file|*.csv|Text": fileErr = Err.Number: On Error GoTo 0
    Check fileErr = 5, "Unpaired file filter not refused"
    Check KitSafeFileName("Q3: a/b *final*?.") = "Q3_ a_b _final__" And KitSafeFileName(" # ") = "_" And KitSafeFileName("...") = "file", "Unsafe file name characters kept"
    Check KitSafeFileName("Weekly" & vbCrLf & "report" & vbTab & "v2" & Chr$(7)) = "Weekly report v2_" And KitSafeFileName("con.txt") = "_con.txt" And _
        KitSafeFileName("LPT1") = "_LPT1" And KitSafeFileName("Console") = "Console", "Control characters or device names kept in a file name"
    RememberFolder "test", root & "\some file.csv"
    Check LastFolder("test") = root, "Chosen file's folder not remembered"
    RememberFolder "test", root
    Check LastFolder("test") = root And LastFolder("other") = "", "Folder not remembered per purpose"
    RememberFolder "test", Environ$("SystemDrive") & "\report.csv"
    Check LastFolder("test") = Environ$("SystemDrive") & "\", "Drive root remembered without its backslash"
    RememberFolder "test", root
    Check WithoutFilterExtension("Q3 report.CSV", "CSV file|*.csv|Excel workbook|*.xlsx") = "Q3 report" And _
        WithoutFilterExtension("Q3.report", "CSV file|*.csv") = "Q3.report" And WithoutFilterExtension("notes.log", "Text|*.txt;*.log") = "notes", _
        "Suggested name keeps a filter's extension"
    SaveSetting APP_NAME, "Files", "Folder_test", root & "\gone"
    Check LastFolder("test") = "", "Missing remembered folder offered"
    DeleteSetting APP_NAME, "Files"
    ' Back-end links and version checks (modKitBackend).
    Dim be1 As String, be2 As String, beDb As DAO.Database, beErr As Long, beFile As Integer
    be1 = root & "\be_one.accdb": be2 = root & "\be_two.accdb"
    Set beDb = DBEngine.CreateDatabase(be1, dbLangGeneral)
    beDb.Execute "CREATE TABLE be_items (id LONG PRIMARY KEY, name TEXT(60))", dbFailOnError
    beDb.Execute "INSERT INTO be_items VALUES (1, 'one')", dbFailOnError
    beDb.Execute "CREATE TABLE be_notes (id LONG PRIMARY KEY)", dbFailOnError
    beDb.Close
    FileCopy be1, be2
    Set beDb = DBEngine.OpenDatabase(be2)
    beDb.Execute "UPDATE be_items SET name = 'two'", dbFailOnError
    beDb.Close
    Check KitLinkTables(be1, Array("be_items", "be_notes")) = 2, "Tables not linked"
    Check KitBackendPath("be_items") = be1 And KitBackendPath("test_rows") = "" And KitBackendPath("no_table") = "", "Linked data file not read"
    Check KitLinkedTables().Count = 2 And KitBackendReachable("be_items"), "Linked tables not listed or not reachable"
    Check KitRelinkAll(be2) = 2 And KitBackendPath("be_notes") = be2 And DLookup("name", "be_items", "id=1") = "two", "Relink did not move every table"
    Check KitEnsureBackend("be_items", "be_two.accdb"), "Reachable data file not accepted"
    Check KitEnsureBackend("test_rows", "be_two.accdb"), "Local table taken for a lost data file"
    ' A second, encrypted data file keeps its own links and password when the first moves.
    Dim be3 As String
    be3 = root & "\be_pw.accdb"
    Set beDb = DBEngine.CreateDatabase(be3, dbLangGeneral & ";pwd=kit test")
    beDb.Execute "CREATE TABLE be_people (id LONG PRIMARY KEY)", dbFailOnError
    beDb.Close
    Check KitLinkTables(be3, Array("be_people"), "kit test") = 1 And KitBackendPath("be_people") = be3, "Encrypted data file not linked"
    Check KitLinkedTables(be2).Count = 2 And KitLinkedTables().Count = 3, "Linked tables not listed per file"
    Check KitRelinkAll(be1, be2) = 2 And KitBackendPath("be_people") = be3 And DCount("*", "be_people") = 0, "Relink moved another file's tables"
    Check KitRelinkAll(be2, be1) = 2 And KitRelinkAll(be3, be3) = 1 And DCount("*", "be_people") = 0, "Relink lost the link's password"
    CurrentDb.TableDefs.Delete "be_people"
    beErr = 0: On Error Resume Next: KitCheckBackendPath "be_one.accdb": beErr = Err.Number: On Error GoTo 0
    Check beErr = 5, "Relative data file path accepted"
    beErr = 0: On Error Resume Next: KitCheckBackendPath CurrentDb.Name: beErr = Err.Number: On Error GoTo 0
    Check beErr = 5, "Front end accepted as its own data file"
    beErr = 0: On Error Resume Next: KitCheckBackendPath root & "\missing.accdb": beErr = Err.Number: On Error GoTo 0
    Check beErr = 5, "Missing data file accepted"
    Set beDb = DBEngine.OpenDatabase(be1)
    beDb.Execute "DROP TABLE be_notes", dbFailOnError
    beDb.Close
    beErr = 0: On Error Resume Next: KitRelinkAll be1: beErr = Err.Number: On Error GoTo 0
    Check beErr = 5 And KitBackendPath("be_items") = be2, "Relink to a file missing a table changed the links"
    beFile = FreeFile: Open root & "\junk.accdb" For Output As #beFile: Print #beFile, "not a database": Close #beFile
    beErr = 0: On Error Resume Next: KitRelinkAll root & "\junk.accdb": beErr = Err.Number: On Error GoTo 0
    Kill root & "\junk.accdb"
    Check beErr = 5 And KitBackendPath("be_items") = be2, "A file that won't open wasn't refused as the data file"
    CurrentDb.TableDefs.Delete "be_items": CurrentDb.TableDefs.Delete "be_notes"
    Check KitVersionProblem(3, 3, "", "2026-10-16 10:00") = "", "Matching versions reported a problem"
    Check InStr(KitVersionProblem(4, 3, "", "x", "Copy it from S:\App."), "newer version") > 0 And _
        InStr(KitVersionProblem(4, 3, "", "x", "Copy it from S:\App."), "S:\App.") > 0, "Newer data file not reported with the advice"
    Check InStr(KitVersionProblem(2, 3, "", "x"), "older version") > 0, "Older data file not reported"
    Check InStr(KitVersionProblem(3, 3, "2026-10-16 12:00", "2026-10-16 10:00"), "built on or after") > 0 And _
        KitVersionProblem(3, 3, "2026-10-16 12:00", "2026-10-17 09:00") = "", "Minimum build not compared"
    ' Personnel directory (modPersonnel): create, link, sign-in.
    Dim pePath As String, peId As Long, pe2 As Long, peErr As Long
    On Error Resume Next                        ' links left by an earlier run
    CurrentDb.TableDefs.Delete PERSON_TABLE: CurrentDb.TableDefs.Delete PERSON_APP_TABLE
    On Error GoTo 0
    PersonnelSignOut
    KitLinkTables be2, Array("be_items")
    Check PersonnelDefaultPath("be_items") = Left$(root, InStrRev(root, "\") - 1) & "\Personnel\" & PERSONNEL_FILE, "Default Personnel place not beside the data folder"
    CurrentDb.TableDefs.Delete "be_items"
    TestPersonnelNotSetUp root, be2
    pePath = root & "\pe\" & PERSONNEL_FILE
    peId = CreatePersonnelFile(pePath, "TEST", "Pat", "O'Neil", "LT")
    Check PersonnelPath() = pePath And PersonnelStatus("no_table") = PERSONNEL_READY And PersonnelFileProblem(pePath) = "", "New Personnel file not linked"
    Check InStr(CurrentDb.TableDefs(PERSON_TABLE).Connect, ";PWD=") > 0 And PersonnelFileProblem(be2) <> "", "Personnel file not encrypted, or another file accepted"
    ' Site key: a new random one, saved beside the file and read from there.
    Dim peKey As String, peKeyFile As String
    peKeyFile = root & "\pe\" & PERSONNEL_KEY_FILE
    peKey = PersonnelKeyFor(pePath)
    Check FileExists(peKeyFile) And Len(peKey) = PERSONNEL_KEY_MAX And PersonnelKeyProblem(peKey) = "" And _
        InStr(ReadTextFile(peKeyFile), peKey) = 1 And InStr(CurrentDb.TableDefs(PERSON_TABLE).Connect, "PWD=" & peKey) > 0, "Site key not made, saved or used"
    Check NewPersonnelKey() <> peKey, "Site keys not random"
    Name peKeyFile As peKeyFile & ".bak"
    Check PersonnelFileProblem(pePath) <> "", "Personnel file opened without its key file"
    Name peKeyFile & ".bak" As peKeyFile
    Check PersonnelKeyProblem("short1") <> "" And PersonnelKeyProblem("abcdef;123456") <> "" And PersonnelKeyProblem("abc def 123456") <> "" And _
        PersonnelKeyProblem(String(21, "a")) <> "" And PersonnelKeyProblem("Site-Key.2026!") = "", "Site key rules wrong"
    Check DLookup("display_name", PERSON_TABLE, "person_id=" & peId) = "LT P. O'Neil" And DLookup("is_admin", PERSON_TABLE, "person_id=" & peId), "First admin not added"
    peErr = 0: On Error Resume Next: CreatePersonnelFile root & "\pe2.accdb", "TEST", "A", "B": peErr = Err.Number: On Error GoTo 0
    Check peErr = 5 And Not FileExists(root & "\pe2.accdb"), "Second Personnel file created while one is linked"
    Check Not FileExists(root & "\" & PERSONNEL_KEY_FILE), "Refused Personnel file left a key file"
    Check PersonnelSignInWindows("TEST", True) And SignedInPersonID() = peId And SignedInIsAdmin(), "Windows sign-in did not find the creator"
    Check DLookup("windows_sid", PERSON_TABLE, "person_id=" & peId) = WinSid(), "SID not saved"
    PersonnelSignOut
    Check SignedInPersonID() = 0 And Not PersonnelSignInWindows("OTHER", True), "Sign-in let a person into an app they lack"
    Check PersonnelWindowsPersonId() = peId, "Windows account's person not found without app access"
    Check PersonnelAdminNames() = " (LT P. O'Neil)", "Admin to ask not named: " & PersonnelAdminNames()
    Check PersonnelWindowsWhy("OTHER") = "noaccess" And PersonnelWindowsWhy("TEST") = "" And _
        InStr(PersonnelSignInWhyText("noaccess"), "access to " & APP_NAME) > 0 And PersonnelSignInWhyText("signout") = "", "Why the Windows sign-in failed not told apart"
    pe2 = AddPerson("Ann", "Lee", "YN2", "", "alee")
    SetPersonApp pe2, "TEST", True
    peErr = 0: On Error Resume Next: AddPerson "Ann", "Lee", "", "", "alee": peErr = Err.Number: On Error GoTo 0
    Check peErr = 5, "Duplicate sign-in name accepted"
    Check PersonnelNeedsPassword("alee") And PersonnelSignIn("alee", "anything1", "TEST") <> "", "Sign-in without a password set"
    Check Not PersonnelNeedsPassword("nobody"), "Unknown username asked for a first password"
    Check PersonnelSetPassword(pe2, "short") <> "" And PersonnelSetPassword(pe2, "long enough 1") = "", "Password length not checked"
    Check PersonnelSignIn("alee", "wrong pass", "TEST") <> "" And SignedInPersonID() = 0, "Wrong password accepted"
    Check PersonnelSignIn("alee", "long enough 1", "TEST") = "" And SignedInPersonID() = pe2 And Not SignedInIsAdmin() And KitUser() = "alee", "Password sign-in failed"
    PersonnelSignOut
    SetPersonApp pe2, "TEST", False
    Dim peWhy As String
    peWhy = PersonnelSignIn("alee", "long enough 1", "TEST")
    Check InStr(peWhy, "access to " & APP_NAME) > 0 And InStr(peWhy, "LT P. O'Neil") > 0, "Sign-in ignored the app access, or didn't say who to ask: " & peWhy
    peErr = 0: On Error Resume Next: PersonnelSaveWindowsAccount: peErr = Err.Number: On Error GoTo 0
    Check peErr = 5, "Windows account saved with nobody signed in"
    ' Personnel admin: changing people, the list, import.
    Dim peRs As DAO.Recordset, peMsg As String, peFile As Integer
    Check PersonnelSignInWindows("TEST", True) And SignedInIsAdmin(), "Admin not signed in for the admin checks"
    peErr = 0: On Error Resume Next: UpdatePerson peId, "Pat", "O'Neil", "LT", WinAccount(), "poneil", False, True: peErr = Err.Number: On Error GoTo 0
    Check peErr = 5 And DLookup("is_admin", PERSON_TABLE, "person_id=" & peId), "Admin took away their own admin access"
    peErr = 0: On Error Resume Next: SetPersonApp peId, "TEST", False: peErr = Err.Number: On Error GoTo 0
    Check peErr = 5 And PersonCanUse(peId, "TEST"), "Admin took away their own access to the app"
    peErr = 0: On Error Resume Next: UpdatePerson pe2, "Ann", "Lee", "YN1", "", DLookup("username", PERSON_TABLE, "person_id=" & peId), False, True: peErr = Err.Number: On Error GoTo 0
    Check peErr = 5, "A person was given someone else's sign-in name"
    UpdatePerson pe2, "Ann", "Lee", "YN1", "TESTDOM\alee", "alee", False, True, "B", "USS Test", "ann.lee@test.mil"
    Check DLookup("display_name", PERSON_TABLE, "person_id=" & pe2) = "YN1 A. Lee" And DLookup("windows_account", PERSON_TABLE, "person_id=" & pe2) = "TESTDOM\alee" And _
        DLookup("command", PERSON_TABLE, "person_id=" & pe2) = "USS Test", "A person's changes not saved"
    PersonnelSignOut
    peErr = 0: On Error Resume Next: UpdatePerson peId, "Pat", "O'Neil", "LT", WinAccount(), "poneil", False, True: peErr = Err.Number: On Error GoTo 0
    Check peErr = 5 And DLookup("is_admin", PERSON_TABLE, "person_id=" & peId), "People changed with no admin signed in"
    peErr = 0: On Error Resume Next: PersonnelImport root & "\none.csv", "TEST": peErr = Err.Number: On Error GoTo 0
    Check peErr = 5, "Import allowed with no admin signed in"
    Check PersonnelSignInWindows("TEST", True) And PersonnelIsAdminNow(), "Admin not signed back in"
    Set peRs = CurrentDb.OpenRecordset(PersonnelPeopleSql("lee", "TEST"), dbOpenSnapshot)
    Check Not peRs.EOF And peRs.Fields(0) = pe2 And peRs.Fields(4) = "No" And peRs.Fields(5) = "Active", "Personnel list wrong for a search"
    peRs.MoveNext: Check peRs.EOF, "Personnel search matched too many"
    peRs.Close
    Set peRs = CurrentDb.OpenRecordset(PersonnelPeopleSql("[*", "TEST"), dbOpenSnapshot)
    Check peRs.EOF, "Personnel search treated [ or * as a wildcard"
    peRs.Close
    peFile = FreeFile
    Open root & "\people.csv" For Output As #peFile
    Print #peFile, "First name,Last name,Rank,Sign-in name,Windows account"
    Print #peFile, "Bo,Chen,LTJG,bchen,"
    Print #peFile, ",NoFirst,,,"
    Print #peFile, "Ann,Lee,,alee,"
    Close #peFile
    peMsg = PersonnelImport(root & "\people.csv", "TEST")
    Check InStr(peMsg, "Added 1 person") > 0 And InStr(peMsg, "Skipped 2") > 0 And InStr(peMsg, "Row 4") > 0, "CSV import summary wrong: " & peMsg
    Check PersonCanUse(PersonIdFor("bchen"), "TEST"), "Imported person can't open the app"
    CurrentDb.Execute "CREATE TABLE zz_pe_people (first_name TEXT(60), last_name TEXT(60), username TEXT(64))", dbFailOnError
    CurrentDb.Execute "INSERT INTO zz_pe_people VALUES ('Cy', 'Diaz', 'cdiaz')", dbFailOnError
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "zz_pe_people", root & "\people.xlsx", True
    CurrentDb.TableDefs.Delete "zz_pe_people"
    peMsg = PersonnelImport(root & "\people.xlsx", "TEST")
    Check InStr(peMsg, "Added 1 person") > 0 And PersonIdFor("cdiaz") > 0 And Not HasTableNamed("zz_personnel_import"), "Excel import failed: " & peMsg
    peErr = 0: On Error Resume Next: PersonnelImport root & "\people.txt", "TEST": peErr = Err.Number: On Error GoTo 0
    Check peErr = 5, "Import accepted a file type it can't read"
    Kill root & "\people.csv": Kill root & "\people.xlsx"
    ' Personnel screens (modPersonnelUI): built, and open in their first state.
    Dim peFrm As Form, peCtl As Control, peLow As Long
    Check PersonnelCanSetFirstPassword("alee") = False And PersonIdFor("alee") = pe2 And WindowsAccountKnown(), "First-password or account rules wrong"
    Check InStr(PersonnelAdminNote(), WinAccount()) > 0 And InStr(PersonnelAdminNote(), APP_NAME) > 0, "Admin note lacks the account or app"
    Check KitCopyText("kit copy test"), "Text not put on the clipboard"
    BuildPersonnelDialogs
    DoCmd.OpenForm SETUP_FORM
    Set peFrm = Forms(SETUP_FORM)
    Check peFrm.btnFind.Visible And peFrm.btnCreate.Visible And Not peFrm.txtFirst.Visible And Not peFrm.txtAsk.Visible, "Setup did not open on its choices"
    peLow = 0
    For Each peCtl In peFrm.Controls
        If peCtl.Visible And peCtl.Top + peCtl.Height > peLow Then peLow = peCtl.Top + peCtl.Height
    Next
    Check peFrm.Section(acDetail).Height = peLow + peFrm.lblTitle.Top, "Setup not fitted to its choices"
    DoCmd.Close acForm, SETUP_FORM, acSaveNo
    DoCmd.OpenForm SIGNIN_FORM
    Set peFrm = Forms(SIGNIN_FORM)
    Check Not peFrm.txtAccount.Visible And peFrm.txtUser.Visible And peFrm.lblUser.Top < peFrm.txtAccount.Top, "Sign-in kept the not-set-up row after sign out"
    DoCmd.Close acForm, SIGNIN_FORM, acSaveNo
    DoCmd.OpenForm PERSONNEL_FORM
    Set peFrm = Forms(PERSONNEL_FORM)
    Check peFrm.OnOpen = "[Event Procedure]" And peFrm.lstPeople.Visible, "Personnel page's admin check not wired to Open"
    Check peFrm.KeyPreview And peFrm.OnKeyDown = "[Event Procedure]", "Personnel page doesn't catch F1 itself"
    Check peFrm.lstPeople.ListCount = DCount("*", PERSON_TABLE) + 1 And Not peFrm.lblNone.Visible, "Personnel page doesn't list everyone"
    DoCmd.Close acForm, PERSONNEL_FORM, acSaveNo
    DoCmd.OpenForm PERSON_FORM, , , , , , "new"
    Set peFrm = Forms(PERSON_FORM)
    Check peFrm.OnOpen = "[Event Procedure]", "Add person's admin check not wired to Open"
    Check peFrm.lblTitle.Caption = "Add person" And peFrm.chkActive And peFrm.chkApp And Not peFrm.chkAdmin And IsNull(peFrm.txtFirst), "Add person didn't open empty"
    DoCmd.Close acForm, PERSON_FORM, acSaveNo
    DoCmd.OpenForm PERSON_FORM, , , , , , CStr(pe2)
    Set peFrm = Forms(PERSON_FORM)
    Check peFrm.txtUser = "alee" And peFrm.txtAccount = "TESTDOM\alee" And peFrm.txtMiddle = "B" And Not peFrm.lblUserHint.Visible, "Change person didn't show the person"
    DoCmd.Close acForm, PERSON_FORM, acSaveNo
    TestMore
    PersonnelSignOut
    CurrentDb.TableDefs.Delete PERSON_TABLE: CurrentDb.TableDefs.Delete PERSON_APP_TABLE
    g_Silent = False
    TestKit = checks & " kit checks passed"
End Function

' Dropdown menu (KitMenu, KitMenuAt). Its own procedure: TestKit is at VBA's
' procedure size limit.
Private Sub TestKitMenu(host As Form)
    ' Dropdown menu: placed under its button (right-aligned or above when it would
    ' run off), sized to its items, separators and disabled items skipped by the
    ' keys, a pick returns the key and closes it, Esc returns nothing.
    Dim mx As Long, my As Long, key As Integer, badName As Long, badScale As Long
    MenuSpot 100, 100, 130, 130, 200, 150, 0, 0, 1000, 800, 4, mx, my
    Check mx = 100 And my = 134, "Menu is not under its button"
    MenuSpot 900, 100, 930, 130, 200, 150, 0, 0, 1000, 800, 4, mx, my
    Check mx = 730 And my = 134, "Menu at the right edge is not right-aligned"
    MenuSpot 100, 700, 130, 730, 200, 150, 0, 0, 1000, 800, 4, mx, my
    Check mx = 100 And my = 546, "Menu at the bottom is not above its button"
    Check MenuOpen(host.tglTest, Array("edit|Edit R&D|E70F", "-", "~arch|Archive", "del|!Delete|E74D")), "Menu did not open"
    Check IsFormOpen("dlgMenu"), "Menu form is not open"
    Check Not MenuOpen(host.tglTest, Array("x|Other")), "A second menu opened while one was open"
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
    host.tglTest.ControlTipText = "Actions"
    Check MenuOpen(host.tglTest, Array("a|Alpha", "b|Beta")), "Menu did not open again"
    Check host.tglTest.ControlTipText = "", "Anchor tooltip not cleared while the menu shows"
    key = vbKeyEscape: MenuKey Forms("dlgMenu"), key
    Check Not IsFormOpen("dlgMenu") And MenuResult() = "", "Esc did not close the menu with nothing picked"
    Check host.tglTest.ControlTipText = "Actions", "Anchor tooltip not restored when the menu closed"
    host.tglTest.ControlTipText = ""
    On Error Resume Next
    MenuOpen host.tglTest, Array("a|A", "b|B", "c|C", "d|D", "e|E", "f|F", "g|G", "h|H", "i|I", "j|J", "k|K"): badName = Err.Number: Err.Clear
    MenuOpen host.tglTest, Array("nokey"): badScale = Err.Number: Err.Clear
    On Error GoTo 0
    Check badName = 5 And badScale = 5 And Not IsFormOpen("dlgMenu"), "Menu took too many items or a bad one"
    ' The same menu at the pointer (a row's right-click).
    Check MenuOpenAt(Array("a|Alpha", "-", "b|Beta")), "Menu at the pointer did not open"
    Check IsFormOpen("dlgMenu"), "Menu at the pointer is not open"
    Check Forms("dlgMenu").InsideHeight = px(4 + 32 + 9 + 32 + 4), "Menu at the pointer is not sized to its items"
    MenuPick Forms("dlgMenu"), 2
    Check Not IsFormOpen("dlgMenu") And MenuResult() = "b", "Menu at the pointer did not return its pick"
    On Error Resume Next
    MenuOpenAt Array("nokey"): badName = Err.Number: Err.Clear
    On Error GoTo 0
    Check badName = 5 And Not IsFormOpen("dlgMenu"), "Menu at the pointer took a bad item"
    DoCmd.SelectObject acForm, "frmTest"        ' the menu was the active window
    ' The picker (modMenu too): Esc cancelling its query is not an error.
    Check PickCanceled(3059) And Not PickCanceled(5), "Picker took Esc cancelling its query as an error"
End Sub

' Personnel not set up: asked for a data file whose neighbour has no Personnel
' folder, since the front end's own neighbour may have one (a Personnel folder
' beside the run's output). Its own procedure: TestKit is at VBA's procedure
' size limit.
Private Sub TestPersonnelNotSetUp(ByVal root As String, ByVal be2 As String)
    On Error Resume Next: MkDir root & "\pe0": MkDir root & "\pe0\data": On Error GoTo 0
    FileCopy be2, root & "\pe0\data\be_two.accdb"
    KitLinkTables root & "\pe0\data\be_two.accdb", Array("be_items")
    Check PersonnelPath() = "" And PersonnelStatus("be_items") = PERSONNEL_NOT_SET_UP, "Personnel reported set up before it was"
    CurrentDb.TableDefs.Delete "be_items"
End Sub

' Trust check: matching Trusted Locations, a dlgTrust that runs no code or
' expressions (it opens with the code off), isn't modal and quits Access from
' Close, the AutoExec macro that opens it with the code off, and the page as
' ShowTrustHelp shows it with the code on. Its own procedure: TestKit is at
' VBA's procedure size limit.
Private Sub TestTrustCheck()
    Dim frm As Form, c As Control, text As String, fitted As String, found As String
    Check PathCovers("C:\Apps\", False, "c:\apps") And Not PathCovers("C:\Apps", False, "C:\Apps\Sub"), "Trusted Location matching differs"
    Check PathCovers("C:\", True, "C:\Apps\Sub") And Not PathCovers("C:\Apps", True, "C:\AppsOld"), "Trusted Location subfolder matching differs"
    Check RecordPath("%USERPROFILE%/My%20Apps/App.accdb") = LCase$(Environ$("USERPROFILE") & "\my apps\app.accdb"), "Trusted document name not decoded"
    BuildTrustCheck "frmTest"
    DoCmd.OpenForm "dlgTrust", acDesign
    Set frm = Forms("dlgTrust")
    Check Not frm.HasModule, "dlgTrust has a code module"
    Check Not frm.Modal And frm.PopUp, "dlgTrust is modal, so Enable Content doesn't work while it is open"
    Check Not frm.CloseButton, "dlgTrust's X would leave an empty Access open with the code off"
    Check frm.btnClose.OnClick = "mcrTrustQuit", "dlgTrust Close doesn't quit Access with the code off"
    ' Every event and expression would be blocked with the code off.
    found = SetEventNames(frm.Properties)
    For Each c In frm.Controls
        text = SetEventNames(c.Properties)
        If c.Name = "btnClose" Then text = Replace(text, " OnClick", "")
        found = found & text
        If c.ControlType = acTextBox Then
            If Left$(c.ControlSource, 1) = "=" Then found = found & " " & c.Name & ".ControlSource"
        End If
    Next
    Check Len(found) = 0, "dlgTrust runs code or expressions:" & found
    Check Not frm.txtFolder.Visible And frm.lblFolderHint.Visible, "dlgTrust's folder hint doesn't show with the code off"
    Set frm = Nothing
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
    Application.SaveAsText acMacro, "mcrTrustQuit", text
    fitted = Replace(Replace(Replace(Replace(ReadTextFile(text), vbNullChar, ""), " ", ""), vbCr, ""), vbLf, "")
    Kill text
    Check InStr(fitted, "Quit") > 0 And InStr(fitted, "Argument=""1""") = 0 And InStr(fitted, "Argument=""2""") = 0, _
        "mcrTrustQuit doesn't quit with Options at Prompt"
    ' With the code on: Close only closes the page, and the box shows the folder.
    ShowTrustHelp
    Set frm = Forms("dlgTrust")
    Check frm.btnClose.OnClick = "mcrTrustClose", "dlgTrust Close quits Access with the code on"
    Check frm.txtFolder.Visible And frm.txtFolder.Value = CurrentProject.Path And Not frm.lblFolderHint.Visible, _
        "dlgTrust doesn't show the folder with the code on"
    Set frm = Nothing
    DoCmd.Close acForm, "dlgTrust", acSaveNo
End Sub

' The names of props' event properties (On...) that are set, each after a space.
Private Function SetEventNames(props As Object) As String
    Dim prp As Object
    On Error Resume Next
    For Each prp In props
        If Left$(prp.Name, 2) = "On" Then
            If Len(Nz(prp.Value, "")) > 0 Then SetEventNames = SetEventNames & " " & prp.Name
        End If
    Next
End Function

' The Windows Event Log (modKitEventLog): events are written, and LogError writes
' one per entry. Its own procedure: TestKit is at VBA's procedure size limit.
' Checks with their own procedures, called from one line of TestKit (it is
' at VBA's procedure size limit).
Private Sub TestMore()
    TestEventLog
    TestDesigner
    TestDesignerKinds
    TestDesignerData
    TestDesignerMore
End Sub

' Form designer (BuildDesigner, modKitDesigner): elements dropped from the
' palette land snapped on the canvas, the canvas's mouse selects the topmost
' one, drags it and resizes it from its corner, and a layout saves, clears
' and reopens the same. Generate writes the layout's build code and stub and
' builds the same form as a preview.
Private Sub TestDesigner()
    Const DZ_FORM As String = "pgDesignerTest"
    Dim frm As Form, g As Control, k As Long, id As Long, code As String
    BuildDesigner DZ_FORM, px(32), px(1056), px(560)
    DoCmd.OpenForm DZ_FORM, WindowMode:=acHidden
    Set frm = Forms(DZ_FORM)
    Set g = frm("dzGlass")
    Check HasTableNamed("kit_layout") And HasTableNamed("kit_layout_item"), "Designer layout tables not made"
    Check DesignerCount() = 0 And frm("dzHint").Visible And Not frm("dzSel").Visible, "Designer didn't open empty"
    ' Dropped centred on the pointer, snapped to 8px and kept on the canvas.
    Check DesignerPlace(frm, "card", 100, 100) = 1, "Designer card not placed"
    Check DesignerItem(1) = "card 0,32 240x136 Card title", "Designer card landed at " & DesignerItem(1)
    Check frm("dzCard1").Visible And frm("dzCard1").Left = g.Left And frm("dzCard1").Top = g.Top + px(32) _
        And frm("dzCard1_t").Caption = "Card title" And Not frm("dzHint").Visible, "Designer card preview not drawn"
    Check DesignerPlace(frm, "button", 400, 300) = 2, "Designer button not placed"
    Check DesignerItem(2) = "button 352,288 96x32 Button" And frm("dzSel").Visible And frm("dzGrip").Visible, _
        "Designer button landed at " & DesignerItem(2)
    ' Drag the button by the canvas's mouse, then resize it from its corner (a
    ' button is at most 48px tall).
    DesignerMouse frm, 1, acLeftButton, px(360), px(290)
    DesignerMouse frm, 2, acLeftButton, px(381), px(330)
    DesignerMouse frm, 3, acLeftButton, px(381), px(330)
    Check DesignerItem(2) = "button 376,328 96x32 Button" And frm("dzBtn1").Left = g.Left + px(376), _
        "Designer drag moved the button to " & DesignerItem(2)
    DesignerMouse frm, 1, acLeftButton, px(472), px(360)
    DesignerMouse frm, 2, acLeftButton, px(500), px(380)
    DesignerMouse frm, 3, acLeftButton, px(500), px(380)
    Check DesignerItem(2) = "button 376,328 128x48 Button" And frm("dzBtn1").Width = px(128), _
        "Designer resize gave " & DesignerItem(2)
    ' A click on the card alone selects it; its title follows its settings.
    DesignerMouse frm, 1, acLeftButton, px(10), px(40)
    DesignerMouse frm, 3, acLeftButton, px(10), px(40)
    Check frm("lblDzKind").Caption = "Card" And frm("btnDzEdit").Enabled, "Designer click didn't select the card"
    Check DesignerSetProps(1, "Readiness", "", "", "") = "", "Designer card settings refused"
    DesignerRedraw frm
    Check frm("dzCard1_t").Caption = "Readiness", "Designer card title not changed"
    ' Save, clear, reopen.
    frm("txtDzName").Value = "Test layout"
    Check DesignerSave(frm), "Designer layout not saved"
    id = Nz(DLookup("id", "kit_layout", "layout_name='Test layout'"), 0)
    Check id > 0 And DCount("*", "kit_layout_item", "layout_id=" & id) = 2 _
        And DLookup("item_name", "kit_layout_item", "kind='button'") = "btn1", "Designer layout rows not written"
    DesignerNew frm
    Check DesignerCount() = 0 And Not frm("dzCard1").Visible And Not frm("dzBtn1").Visible, "Designer New left elements"
    frm("cboDzOpen").Value = id
    DesignerOpen frm
    Check DesignerCount() = 2 And DesignerItem(1) = "card 0,32 240x136 Readiness" _
        And DesignerItem(2) = "button 376,328 128x48 Button" And frm("txtDzName").Value = "Test layout", "Designer layout didn't reopen the same"
    ' Generate: the build procedure and stub, then the preview built from the same calls.
    Check DesignerFormName("  review   navy 2!") = "pgReviewNavy2" And DesignerFormName("--") = "pgLayout", "Designer form name wrong"
    code = DesignerCode(id)
    Check InStr(code, "Private Sub Build_pgTestLayout()") > 0 And InStr(code, "NewForm """", 0, px(776), px(408)") > 0 _
        And InStr(code, "Card ""card1"", ""Readiness"", px(0), px(32), px(240), px(136)") > 0 _
        And InStr(code, "Btn ""btn1"", ""Button"", px(376), px(328), px(128), px(48)") > 0 _
        And InStr(code, "SaveForm ""pgTestLayout""") > 0, "Designer build code wrong: " & code
    Check InStr(code, "Card ""card1""") < InStr(code, "Btn ""btn1"""), "Designer build code not back to front"
    code = DesignerStub(id)
    Check InStr(code, "Private Sub btn1_Click()") > 0 And InStr(code, "HandleError ""pgTestLayout.btn1_Click""") > 0, _
        "Designer code-behind stub wrong: " & code
    DesignerGenerate frm
    Check FileExists(CurrentProject.Path & "\designer\Build_pgTestLayout.txt") _
        And FileExists(CurrentProject.Path & "\designer\pgTestLayout.vba"), "Designer Generate didn't write its files"
    Check IsFormOpen("dzPreview"), "Designer preview not open"
    With Forms("dzPreview")
        Check .Controls("boxcard1").Left = 0 And .Controls("boxcard1").Top = px(32) And .Controls("lblcard1Hdr").Caption = "Readiness" _
            And .Controls("btn1").Left = px(376) And .Controls("btn1").Width = px(128) And .Controls("btn1").Caption = "Button", _
            "Designer preview doesn't match the layout"
    End With
    DoCmd.Close acForm, "dzPreview", acSaveNo
    DoCmd.DeleteObject acForm, "dzPreview"
    Kill CurrentProject.Path & "\designer\*.*"
    RmDir CurrentProject.Path & "\designer"
    ' Delete, and a full pool refuses another element.
    DesignerSelect frm, 2
    DesignerDelete frm
    Check DesignerCount() = 1 And Not frm("dzBtn1").Visible And Not frm("dzSel").Visible, "Designer delete left the button"
    For k = 1 To 5
        DesignerPlace frm, "card", 400, 200
    Next
    Check DesignerCount() = 6 And DesignerPlace(frm, "card", 400, 200) = 0, "Designer card pool not limited to 6"
    DoCmd.Close acForm, DZ_FORM, acSaveNo
    DoCmd.DeleteObject acForm, DZ_FORM
    CurrentDb.TableDefs.Delete "kit_layout_item": CurrentDb.TableDefs.Delete "kit_layout"
End Sub

' Form designer elements beyond the first four: each lands at its default
' size in the outlined-box pool, settings are checked against the kit
' builder's rules and resize the element they constrain, and Generate writes
' and builds every kind with the handlers its stub declares.
Private Sub TestDesignerKinds()
    Const DZ_FORM As String = "pgDesignerKinds"
    Dim frm As Form, kinds() As String, k As Long, id As Long, code As String
    BuildDesigner DZ_FORM, px(32), px(1056), px(560)
    DoCmd.OpenForm DZ_FORM, WindowMode:=acHidden
    Set frm = Forms(DZ_FORM)
    kinds = Split("text;date;secret;toggle;segmented;slider;rating;stepper;search;stat;alert;avatar;empty", ";")
    For k = 0 To UBound(kinds)
        Check DesignerPlace(frm, kinds(k), 300, 200) = k + 1, "Designer " & kinds(k) & " not placed"
        Check frm("dzGen" & (k + 1)).Visible And frm("dzGen" & (k + 1) & "_t").Visible, "Designer " & kinds(k) & " preview not drawn"
    Next
    Check DesignerItem(5) = "segmented 184,184 240x32 All;Open;Done" And frm("dzGen5_t").Caption = "All  |  Open  |  Done", _
        "Designer segmented landed as " & DesignerItem(5)
    Check DesignerItem(12) = "avatar 288,184 32x32 Jane Doe" And frm("dzGen12_t").Caption = "JD", "Designer avatar landed as " & DesignerItem(12)
    ' The builders' rules, refused with a reason and nothing changed.
    Check DesignerSetProps(5, "Only", "", "", "") <> "" And DesignerSetProps(5, "A;;B", "", "", "") <> "", "Designer segmented took bad options"
    Check DesignerSetProps(6, "0", "100", "7", "") <> "" And DesignerSetProps(6, "0", "1,5", "1", "") <> "", "Designer slider took bad stops"
    Check DesignerSetProps(7, "2", "", "", "") <> "" And DesignerSetProps(8, "A;B", "4", "", "") <> "", "Designer rating or stepper took bad settings"
    Check DesignerSetProps(10, "Open", "many", "", "") <> "" And DesignerSetProps(11, "T", "M", "", "loud") <> "" _
        And DesignerSetProps(13, "", "", "", "") <> "", "Designer stat, alert or empty state took bad settings"
    Check DesignerItem(5) = "segmented 184,184 240x32 All;Open;Done", "Designer refused settings still changed " & DesignerItem(5)
    ' Settings that change what fits: more options widen, more stars widen.
    Check DesignerSetProps(5, "Day;Week;Month;Quarter;Year;All", "", "", "") = "", "Designer segmented options refused"
    Check DesignerItem(5) = "segmented 184,184 246x32 Day;Week;Month;Quarter;Year;All", "Designer segmented not widened: " & DesignerItem(5)
    Check DesignerSetProps(7, "8", "", "", "") = "" And DesignerItem(7) = "rating 240,192 192x24 8", "Designer rating gave " & DesignerItem(7)
    Check DesignerSetProps(11, "Saved", "All changes are saved.", "", "success") = "", "Designer alert settings refused"
    DesignerRedraw frm
    Check frm("dzGen11_t").Caption = "success: Saved" And frm("dzGen7_t").Caption = String$(8, ChrW(9733)), "Designer settings not redrawn"
    ' Save and Generate: every kind's builder call, its handler stub, and the preview.
    frm("txtDzName").Value = "Kinds test"
    Check DesignerSave(frm), "Designer kinds layout not saved"
    id = Nz(DLookup("id", "kit_layout", "layout_name='Kinds test'"), 0)
    Check DLookup("item_choice", "kit_layout_item", "layout_id=" & id & " AND kind='alert'") = "success" _
        And DLookup("item_text2", "kit_layout_item", "layout_id=" & id & " AND kind='slider'") = "100", "Designer settings columns not written"
    code = DesignerCode(id)
    Check InStr(code, "Segmented ""seg1"", """", px(184), px(184), px(246), ""day|Day;week|Week;month|Month;quarter|Quarter;year|Year;all|All"", px(32), ""seg1Changed""") > 0 _
        And InStr(code, "Slider ""sld1"", """", px(184), px(192), px(240), 0, 100, 10, ""sld1Changed""") > 0 _
        And InStr(code, "Stepper ""step1"", px(") > 0 And InStr(code, ", 1, ""step1Step"", ""Details"", ""Owner"", ""Review""") > 0 _
        And InStr(code, "Alert ""alert1"", px(") > 0 And InStr(code, """success"", ""Saved"", ""All changes are saved.""") > 0 _
        And InStr(code, "Avatar ""av1"", ""=""""Jane Doe"""""", px(288)") > 0 _
        And InStr(code, "Ev F, ""OnLoad""") > 0, "Designer kinds build code wrong: " & code
    code = DesignerStub(id)
    Check InStr(code, "Public Sub seg1Changed()") > 0 And InStr(code, "Public Sub step1Step(ByVal stepNo As Long)") > 0 _
        And InStr(code, "Public Sub search1Search(ByVal text As String)") > 0 And InStr(code, "Private Sub empty1_b_Click()") > 0 _
        And InStr(code, "tgl1Changed") > 0 And InStr(code, "rate1Changed") > 0 _
        And InStr(code, "SetStep Me, ""step1"", CurrentStep(Me, ""step1"")") > 0, "Designer kinds stub wrong: " & code
    DesignerPreview id
    Check IsFormOpen("dzPreview"), "Designer kinds preview not open"
    With Forms("dzPreview")
        Check .Controls("seg1_bg").Left = px(184) And .Controls("step1").Width = px(360) And .Controls("empty1_b").OnClick = "" _
            And .Controls("av1_bg").Width = px(32) And .OnLoad = "=DesignerPreviewLoad([Form],""step1"")" _
            And .Controls("step1_l1").FontWeight = 600, "Designer kinds preview doesn't match the layout"
    End With
    DoCmd.Close acForm, "dzPreview", acSaveNo
    DoCmd.DeleteObject acForm, "dzPreview"
    DoCmd.Close acForm, DZ_FORM, acSaveNo
    DoCmd.DeleteObject acForm, DZ_FORM
    CurrentDb.TableDefs.Delete "kit_layout_item": CurrentDb.TableDefs.Delete "kit_layout"
End Sub

' Form designer data and behaviour: the form's settings (page or dialog,
' size, record source) are checked and kept, elements bind to its fields
' and anchor on a page, the tab order is the reading order until moved, and
' Generate writes and builds all of it; a layout reopens with all of it.
Private Sub TestDesignerData()
    Const DZ_FORM As String = "pgDesignerData"
    Dim frm As Form, dlg As Form, id As Long, code As String
    CurrentDb.Execute "CREATE TABLE dz_item (id COUNTER PRIMARY KEY, item_name TEXT(60), done YESNO, pct LONG)", dbFailOnError
    CurrentDb.Execute "INSERT INTO dz_item (item_name, done, pct) VALUES ('Pump', True, 40)", dbFailOnError
    BuildDesigner DZ_FORM, px(32), px(1056), px(560)
    DoCmd.OpenForm DZ_FORM, WindowMode:=acHidden
    Set frm = Forms(DZ_FORM)
    Check DesignerFormData() = "page 776x408 |", "Designer form didn't start as a page the canvas's size: " & DesignerFormData()
    ' Placed out of reading order: a button, a text box above it, a toggle left of the text box, a card.
    Check DesignerPlace(frm, "button", 400, 300) = 1 And DesignerPlace(frm, "text", 400, 100) = 2 _
        And DesignerPlace(frm, "toggle", 100, 100) = 3 And DesignerPlace(frm, "card", 600, 100) = 4, "Designer data elements not placed"
    Check DesignerItem(2) = "text 304,88 200x30 " And DesignerItem(3) = "toggle 80,88 44x24 ", "Designer data elements landed at " & _
        DesignerItem(2) & " and " & DesignerItem(3)
    ' The form's settings: refused with the reason, nothing changed.
    Check DesignerSetForm("", "100", "300", "page", "") <> "" And DesignerSetForm("", "400", "300", "page", "") <> "" _
        And DesignerSetForm("", "776", "408", "page", "dz_missing") <> "" And DesignerSetForm("", "776", "408", "sheet", "") <> "" _
        And DesignerSetForm("", "800", "408", "page", "") <> "", "Designer form took bad settings"
    Check DesignerFormData() = "page 776x408 |", "Designer refused form settings still changed " & DesignerFormData()
    Check DesignerSetForm("", "776", "408", "page", "dz_item") = "", "Designer record source refused"
    ' Fields: only from the record source, only on a kind that holds a value.
    Check DesignerSetProps(2, "", "", "", "", "no_such_field") <> "" And DesignerSetProps(4, "Card title", "", "", "", "item_name") <> "", _
        "Designer bound a missing field or a card"
    Check DesignerSetProps(2, "", "", "", "", "item_name") = "" And DesignerSetProps(3, "", "", "", "", "done") = "", _
        "Designer field binding refused"
    Check DesignerSetForm("", "776", "408", "page", "") <> "", "Designer dropped the record source of bound elements"
    ' Anchors: any element moves with an edge; only one control or a card stretches.
    Check DesignerSetProps(1, "Button", "", "", "primary", , "tr") = "" And DesignerSetProps(4, "Card title", "", "", "", , "both") = "", _
        "Designer anchor refused"
    Check DesignerSetProps(3, "", "", "", "", , "both") <> "" And DesignerSetProps(3, "", "", "", "", , "nowhere") <> "", _
        "Designer stretched a toggle, or took an unknown anchor"
    Check DesignerItemData(1) = "|tr|0" And DesignerItemData(2) = "item_name||0" And DesignerItemData(3) = "done||0" _
        And DesignerItemData(4) = "|both|0", "Designer settings kept as " & DesignerItemData(1) & ", " & DesignerItemData(2)
    ' Tab order: the reading order (toggle, text box, button); Move up puts the button second.
    DoCmd.OpenForm "dlgDzTabs"
    Set dlg = Forms("dlgDzTabs")
    Check dlg("lstTabs").ListCount = 3 And InStr(dlg("lstTabs").Column(1, 0), "Toggle") > 0 _
        And InStr(dlg("lstTabs").Column(1, 2), "Button") > 0 And Not dlg("btnReading").Enabled, "Designer tab order isn't the reading order"
    dlg("lstTabs").Value = "2"
    DesignerTabsMove dlg, -1
    Check InStr(dlg("lstTabs").Column(1, 1), "Button") > 0 And dlg("lstTabs").Value = "1" And dlg("btnReading").Enabled, _
        "Designer tab order didn't move the button up"
    DesignerTabsSave dlg
    DoCmd.Close acForm, "dlgDzTabs", acSaveNo
    Check DesignerItemData(1) = "|tr|2" And DesignerItemData(2) = "item_name||3" And DesignerItemData(3) = "done||1", _
        "Designer tab order not kept: " & DesignerItemData(1) & ", " & DesignerItemData(2) & ", " & DesignerItemData(3)
    ' Generate: bound builders, anchors, the tab order, and the same in the preview.
    frm("txtDzName").Value = "Dz data"
    Check DesignerSave(frm), "Designer data layout not saved"
    id = Nz(DLookup("id", "kit_layout", "layout_name='Dz data'"), 0)
    code = DesignerCode(id)
    Check InStr(code, "NewForm ""dz_item"", 0, px(776), px(408)") > 0 _
        And InStr(code, "Txt ""txt1"", ""[item_name]"", px(304), px(88), px(200), px(30)") > 0 _
        And InStr(code, "Toggle ""tgl1"", ""[done]"", px(80), px(88)") > 0 _
        And InStr(code, "Card ""card1"", ""Card title"", px(480), px(32), px(240), px(136), """", 2, 2") > 0 _
        And InStr(code, "    Dim k As String") > 0 _
        And InStr(code, "k = BuiltNames()" & vbCrLf & "    Btn ""btn1""") > 0 And InStr(code, "AnchorBuilt k, 1, 0") > 0 _
        And InStr(code, "TabOrder Array(""tgl1"", ""btn1"", ""txt1"")") > 0 _
        And InStr(code, "    F.Cycle = 1") > 0, "Designer data build code wrong: " & code
    DesignerPreview id
    Check IsFormOpen("dzPreview"), "Designer data preview not open"
    With Forms("dzPreview")
        Check .RecordSource = "dz_item" And Not .AllowEdits And .txt1.Value = "Pump" And .tgl1_val.Value = True _
            And .Cycle = 1, "Designer preview not bound to the record source, read-only, Tab staying on the record"
        Check .btn1.HorizontalAnchor = 1 And .btn1.VerticalAnchor = 0 And .boxcard1.HorizontalAnchor = 2 _
            And .boxcard1.VerticalAnchor = 2 And .txt1.HorizontalAnchor = 0, "Designer preview not anchored"
        Check .tgl1.TabIndex < .btn1.TabIndex And .btn1.TabIndex < .txt1.TabIndex And .txtFocusSink.TabIndex = 0, _
            "Designer preview tab order wrong"
    End With
    DoCmd.Close acForm, "dzPreview", acSaveNo
    ' A dialog: no anchors, its title on the title bar, and a stub that says how to report OK.
    DesignerSelect frm, 4
    DesignerDelete frm
    Check DesignerSetForm("Edit item", "520", "336", "dialog", "dz_item") = "", "Designer dialog settings refused"
    Check DesignerFormData() = "dialog 520x336 dz_item|Edit item", "Designer dialog settings kept as " & DesignerFormData()
    Check DesignerSave(frm), "Designer dialog layout not saved"
    code = DesignerCode(id)
    Check InStr(code, "Private Sub Build_dlgDzData()") > 0 And InStr(code, "NewDialog ""Edit item"", px(520), px(336)") > 0 _
        And InStr(code, "F.RecordSource = ""dz_item""") > 0 And InStr(code, "F.Cycle = 1") > 0 _
        And InStr(code, "AnchorBuilt") = 0, "Designer dialog build code wrong: " & code
    Check InStr(DesignerStub(id), "RunDialog") > 0, "Designer dialog stub doesn't say how to report OK"
    DesignerPreview id
    Check Forms("dzPreview").Caption = "Preview: Edit item" And Forms("dzPreview").btn1.HorizontalAnchor = 0, _
        "Designer dialog preview wrong"
    DoCmd.Close acForm, "dzPreview", acSaveNo
    DoCmd.DeleteObject acForm, "dzPreview"
    ' Reopened: the form's settings, bindings, anchors and tab places come back.
    DesignerNew frm
    Check DesignerFormData() = "page 776x408 |", "Designer New kept the form's settings"
    frm("cboDzOpen").Value = id
    DesignerOpen frm
    Check DesignerFormData() = "dialog 520x336 dz_item|Edit item" And DesignerItemData(1) = "item_name||3" _
        And DesignerItemData(2) = "done||1" And DesignerItemData(3) = "|tr|2", "Designer layout didn't reopen its data settings: " & _
        DesignerFormData() & ", " & DesignerItemData(1) & ", " & DesignerItemData(3)
    DoCmd.Close acForm, DZ_FORM, acSaveNo
    DoCmd.DeleteObject acForm, DZ_FORM
    CurrentDb.TableDefs.Delete "kit_layout_item": CurrentDb.TableDefs.Delete "kit_layout": CurrentDb.TableDefs.Delete "dz_item"
End Sub

' Form designer tables, chips, tab strips, breadcrumbs and accordions: each
' kind's settings are checked (a table's source, its id field and columns),
' a form takes one tab strip and one breadcrumb, and Generate writes the
' table's rows form first, the handlers and Form_Load, and builds a preview
' whose table lists its source and whose chips and tabs light up.
Private Sub TestDesignerMore()
    Const DZ_FORM As String = "pgDesignerMore"
    Dim frm As Form, id As Long, code As String
    CurrentDb.Execute "CREATE TABLE dz_crew (id COUNTER PRIMARY KEY, crew_name TEXT(60), rate TEXT(10), qual TEXT(20))", dbFailOnError
    CurrentDb.Execute "INSERT INTO dz_crew (crew_name, rate, qual) VALUES ('Smith', 'ET1', 'EOOW')", dbFailOnError
    CurrentDb.Execute "CREATE TABLE dz_noid (code TEXT(10))", dbFailOnError
    BuildDesigner DZ_FORM, px(32), px(1056), px(560)
    DoCmd.OpenForm DZ_FORM, WindowMode:=acHidden
    Set frm = Forms(DZ_FORM)
    ' An element that can't be made small enough for the form is refused.
    Check DesignerSetForm("", "200", "120", "page", "") = "" And DesignerPlace(frm, "empty", 50, 50) = 0 _
        And DesignerCount() = 0, "Designer placed an element bigger than the form"
    Check DesignerPlace(frm, "rating", 50, 50) = 1 And DesignerSetProps(1, "10", "", "", "") <> "", "Designer took a rating wider than the form"
    DesignerDelete frm
    Check DesignerSetForm("", "776", "408", "page", "") = "", "Designer form not set back to the canvas's size"
    Check DesignerPlace(frm, "tabs", 200, 16) = 1 And DesignerPlace(frm, "chips", 150, 60) = 2 _
        And DesignerPlace(frm, "breadcrumb", 600, 16) = 3 And DesignerPlace(frm, "table", 200, 200) = 4 _
        And DesignerPlace(frm, "accordion", 600, 250) = 5, "Designer tables, chips, tabs, breadcrumb or accordion not placed"
    Check DesignerItem(1) = "tabs 24,0 360x32 Overview;Details;History" And DesignerItem(2) = "chips 24,48 254x26 All;Open;Done" _
        And DesignerItem(3) = "breadcrumb 440,8 320x20 Home;Reports;This report" And DesignerItem(4) = "table 24,112 360x180 " _
        And DesignerItem(5) = "accordion 416,120 360x258 General;Details;Notes", "Designer new kinds landed as " & _
        DesignerItem(1) & " / " & DesignerItem(2) & " / " & DesignerItem(5)
    Check DesignerPlace(frm, "tabs", 300, 300) = 0 And DesignerPlace(frm, "breadcrumb", 300, 300) = 0 And DesignerCount() = 5, _
        "Designer took a second tab strip or breadcrumb"
    ' Settings refused with the reason: too few or many, a body too short, a form too small.
    Check DesignerSetProps(1, "Only", "", "", "") <> "" And DesignerSetProps(2, "A;B;C;D;E;F;G;H;I", "", "", "") <> "" _
        And DesignerSetProps(3, "Home", "", "", "") <> "", "Designer took a bad tab strip, chip row or breadcrumb"
    Check DesignerSetProps(5, "A;B", "20", "", "one at a time") <> "" And DesignerSetProps(5, "A;B", "120", "", "all") <> "" _
        And DesignerSetProps(5, "A;B;C;D;E;F", "600", "", "any number") <> "", "Designer took a bad accordion"
    ' A table lists a source with an id, its own fields; empty columns are its first fields.
    Check DesignerSetProps(4, "", "", "", "", "dz_noid") <> "" And DesignerSetProps(4, "crew_name;missing", "", "", "", "dz_crew") <> "" _
        And DesignerSetProps(4, "crew_name;rate", "Name", "", "", "dz_crew") <> "" And DesignerSetProps(4, "", "", "", "", "") <> "", _
        "Designer took a table without an id, a field its source lacks, or captions that don't match"
    Check DesignerSetProps(4, "", "", "", "", "dz_crew") = "" And DesignerItem(4) = "table 24,112 360x180 crew_name;rate;qual" _
        And DesignerItemData(4) = "dz_crew||0", "Designer table didn't take its source's fields: " & DesignerItem(4)
    Check DesignerSetProps(4, "crew_name;rate", "Name;Rate", "", "", , "across") = "" And DesignerSetProps(4, "crew_name", "", "", "", , "br") <> "", _
        "Designer table anchors wrong"
    Check DesignerSetForm("", "776", "408", "page", "") = "", "Designer took a table's source for a bound field"
    ' Settings that change what fits: chips are as wide as their captions, an accordion as tall as its bodies.
    Check DesignerSetProps(2, "Mine;Everyone", "", "", "") = "" And DesignerItem(2) = "chips 24,48 200x26 Mine;Everyone", _
        "Designer chips not resized: " & DesignerItem(2)
    Check DesignerSetProps(5, "General;Notes", "80", "", "any number") = "" And DesignerItem(5) = "accordion 416,120 360x272 General;Notes", _
        "Designer accordion not resized: " & DesignerItem(5)
    DesignerRedraw frm
    Check frm("dzGen4_t").Caption = "Table of dz_crew: crew_name, rate" And frm("dzGen5_t").Caption = "Accordion: General, Notes", _
        "Designer new kinds not redrawn"
    ' Generate: the rows form before NewForm, each builder, Form_Load and the handlers.
    frm("txtDzName").Value = "Dz more"
    Check DesignerSave(frm), "Designer new kinds layout not saved"
    id = Nz(DLookup("id", "kit_layout", "layout_name='Dz more'"), 0)
    code = DesignerCode(id)
    Check InStr(code, "BuildGrid ""subPgDzMore_tbl1"", Array(""crew_name|Name|172"", ""rate|Rate|172""), px(360)") > 0 _
        And InStr(code, "BuildGrid ") < InStr(code, "NewForm ") _
        And InStr(code, "GridHeaders ""tbl1_h"", Array(""crew_name|Name|172"", ""rate|Rate|172""), px(24), px(112), , , px(360)") > 0 _
        And InStr(code, "SubF(""tbl1"", ""subPgDzMore_tbl1"", px(24), px(134), px(360), px(158)).TabStop = False" & vbCrLf & _
            "    Anchor F.Controls(""tbl1""), 2, 0") > 0, _
        "Designer table build code wrong: " & code
    Check InStr(code, "ChipRow ""chips1"", px(24), px(48), ""chips1Chip"", ""Mine"", ""Everyone""") > 0 _
        And InStr(code, "TabStrip px(24), px(0), px(360), ""tabs1Tab"", ""overview|Overview"", ""details|Details"", ""history|History""") > 0 _
        And InStr(code, "Breadcrumb ""crumb1"", px(440), px(8), px(320), 3, ""crumb1Crumb"", ""home|Home"", ""reports|Reports"", ""thisreport|This report""") > 0 _
        And InStr(code, "Accordion ""acc1"", px(416), px(120), px(360), False, ""General|subPgDzMore_acc1|80|open"", ""Notes|subPgDzMore_acc1|80""") > 0 _
        And InStr(code, "BuildBlank ""subPgDzMore_acc1"", px(360), px(80)" & vbCrLf) > 0 And InStr(code, "BuildBlank ") < InStr(code, "NewForm ") _
        And InStr(code, "Ev F, ""OnLoad""") > 0 _
        And InStr(code, "TabOrder Array(""crumb10_k"", ""crumb11_k"", ""acc1_h1"", ""acc1_b1"", ""acc1_h2"", ""acc1_b2"")") > 0, _
        "Designer chips, tabs, breadcrumb or accordion build code wrong: " & code
    code = DesignerStub(id)
    Check InStr(code, "Public Function chips1Chip(ByVal i As Long)") > 0 And InStr(code, "    ChipOn Me, ""chips1"", i") > 0 _
        And InStr(code, "Public Function tabs1Tab(ByVal key As String)") > 0 And InStr(code, "    SelectTab Me, key") > 0 _
        And InStr(code, "Exit Function") > 0 And InStr(code, "Public Sub crumb1Crumb(ByVal key As String)") > 0 _
        And InStr(code, "Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)") > 0, "Designer new kinds' handlers wrong: " & code
    Check InStr(code, "Private Sub Form_Load()") > 0 And InStr(code, "    ChipOn Me, ""chips1"", 0") > 0 _
        And InStr(code, "    SelectTab Me, ""overview""") > 0 _
        And InStr(code, "    SetBreadcrumb Me, ""crumb1"", ""home|Home"", ""reports|Reports"", ""thisreport|This report""") > 0 _
        And InStr(code, "    Me(""tbl1"").Form.RecordSource = ""SELECT * FROM [dz_crew]""") > 0, "Designer new kinds' Form_Load wrong: " & code
    ' The preview: the table lists its source, the chips and tabs paint and click, the crumbs call nothing.
    DesignerPreview id
    Check IsFormOpen("dzPreview"), "Designer new kinds preview not open"
    With Forms("dzPreview")
        Check .tbl1.SourceObject = "dzPreview_tbl1" And .tbl1.Form.RecordSource Like "SELECT [*] FROM*dz_crew*" _
            And .tbl1.Form!crew_name = "Smith" And .tbl1.Form.txtCol0.OnClick = "" And .tbl1.HorizontalAnchor = 2 _
            And .tbl1_h0.Caption = "NAME", "Designer preview table wrong"
        Check .OnLoad = "=DesignerPreviewLoad([Form],"""",""chips1"",""overview"",""crumb1"")" And .chips1_0.BackStyle = 1 And .chips1_1.BackStyle = 0 _
            And .ul_overview.Visible And Not .ul_details.Visible And .crumb10.OnClick = "" And .crumb10_k.OnClick = "" _
            And .crumb12.ForeColor = Th(CLR_FG) And .crumb10.ForeColor = Th(CLR_MUTED_FG) _
            And .acc1_b1.Visible And Not .acc1_b2.Visible, "Designer preview chips, tabs, breadcrumb or accordion wrong"
        ' An open body shows the blank form in the palette's page color, not Access's white.
        Check .acc1_b1.SourceObject = "dzPreview_acc1" And .acc1_b2.SourceObject = "dzPreview_acc1" _
            And .acc1_b1.Form.Section(acDetail).BackColor = Th(CLR_BG), "Designer preview accordion body not the blank form"
        Check .crumb10_k.TabIndex < .crumb11_k.TabIndex And .crumb11_k.TabIndex < .acc1_h1.TabIndex _
            And .acc1_h1.TabIndex < .acc1_b1.TabIndex And Not .tbl1.TabStop, _
            "Designer preview tab order wrong for the new kinds"
    End With
    DesignerChipClick Forms("dzPreview"), "chips1", 1
    DesignerTabClick Forms("dzPreview"), "history"
    With Forms("dzPreview")
        Check .chips1_1.BackStyle = 1 And .chips1_0.BackStyle = 0 And .ul_history.Visible And Not .ul_overview.Visible, _
            "Designer preview chip or tab click didn't light it"
    End With
    DoCmd.Close acForm, "dzPreview", acSaveNo
    ' A dialog builds its table's rows first too.
    Check DesignerSetForm("", "776", "408", "dialog", "") = "" And DesignerSave(frm), "Designer new kinds dialog not saved"
    code = DesignerCode(id)
    Check InStr(code, "BuildGrid ""subDlgDzMore_tbl1""") > 0 And InStr(code, "BuildGrid ") < InStr(code, "NewDialog ") _
        And InStr(code, "SubF(""tbl1"", ""subDlgDzMore_tbl1""") > 0 _
        And InStr(code, "BuildBlank ""subDlgDzMore_acc1"", px(360), px(80), CLR_DIALOG" & vbCrLf) > 0, _
        "Designer dialog lost its table's rows or its accordion's blank body: " & code
    DoCmd.DeleteObject acForm, "dzPreview"
    DoCmd.DeleteObject acForm, "dzPreview_tbl1"
    DoCmd.DeleteObject acForm, "dzPreview_acc1"
    DoCmd.Close acForm, DZ_FORM, acSaveNo
    DoCmd.DeleteObject acForm, DZ_FORM
    CurrentDb.TableDefs.Delete "kit_layout_item": CurrentDb.TableDefs.Delete "kit_layout"
    CurrentDb.TableDefs.Delete "dz_crew": CurrentDb.TableDefs.Delete "dz_noid"
End Sub

Private Sub TestEventLog()
    Dim before As Long, evErr As Long
    before = KitEventsWritten()
    Check KitEventLog("Kit test event " & RandomHex(4)) And KitEventsWritten() = before + 1, "Event not written to the Application log"
    Check KitEventLogError(0, 5, "Kit test error", "modKitTests.TestEventLog", "", "", "", "", ""), "Error event not written"
    ' A scratch error_log, so the entry doesn't go to the user's errors.log.
    CurrentDb.Execute "CREATE TABLE error_log (id COUNTER PRIMARY KEY, logged_at DATETIME, error_number LONG, " & _
        "description LONGTEXT, err_source TEXT(255), source TEXT(255), call_trail LONGTEXT, context LONGTEXT, " & _
        "page TEXT(32), active_form TEXT(64), active_control TEXT(64), username TEXT(150), windows_user TEXT(64), " & _
        "computer TEXT(64), fe_build TEXT(32), access_version TEXT(32))", dbFailOnError
    before = KitEventsWritten()
    On Error Resume Next
    Err.Raise 5, , "Kit event log test"
    evErr = Err.Number
    LogError "modKitTests.TestEventLog"
    On Error GoTo 0
    Check evErr = 5 And KitEventsWritten() = before + 1 And DCount("*", "error_log") = 1, "LogError did not write an event"
    CurrentDb.TableDefs.Delete "error_log"
    TestPurge                                   ' here, not in TestKit, which is at VBA's size limit
End Sub

' The purge (modKitPurge): the typed word, dlgConfirm, the user roster, the
' checked backup copy, and emptying tables all or none.
Private Sub TestPurge()
    Dim root As String, be As String, db As DAO.Database, bak As String, n As Long, pErr As Long, frm As Form
    root = CurrentProject.Path & "\purge_test"
    EnsureFolder root
    SetKitConfirm "Delete all test data", "Detail", PURGE_WORD, "Delete all data"
    Check KitConfirmMatches("DELETE") And Not KitConfirmMatches("delete") And Not KitConfirmMatches(" DELETE") _
        And Not KitConfirmMatches("DELETE ") And Not KitConfirmMatches(""), "Typed word not matched exactly"
    BuildConfirmDialog
    DoCmd.OpenForm "dlgConfirm"
    Set frm = Forms("dlgConfirm")
    Check Not frm.btnConfirm.Enabled And Not frm.btnConfirm.Default And Not frm.btnCancel.Default And frm.btnCancel.Cancel, _
        "Confirm button enabled before the word is typed, or Enter confirms"
    Check frm.btnConfirm.Caption = "Delete all data" And InStr(frm.lblAsk.Caption, PURGE_WORD) > 0 And _
        frm.lblHeading.Caption = "Delete all test data", "Confirm dialog doesn't show its heading, word and action"
    Check Not frm.chkAlso.Visible And frm.lblAsk.Top = frm.lblAlso.Top, "Confirm dialog without an option keeps its check box's space"
    DoCmd.Close acForm, "dlgConfirm", acSaveNo
    SetKitConfirm "Delete all test data", "Detail", PURGE_WORD, "Delete all data", "", "Also delete the people"
    DoCmd.OpenForm "dlgConfirm"
    Set frm = Forms("dlgConfirm")
    Check frm.chkAlso.Visible And Not frm.chkAlso And frm.lblAlso.Caption = "Also delete the people", _
        "Confirm dialog's option not shown unticked"
    DoCmd.Close acForm, "dlgConfirm", acSaveNo
    Check Not KitConfirmOptionChosen(), "Option chosen without a click"
    SetKitConfirm "", "", "", ""
    Check Not KitConfirmMatches(""), "An empty word confirmed"
    ' A data file with a parent and child table.
    be = root & "\purge_be.accdb"
    If FileExists(be) Then Kill be
    Set db = DBEngine.CreateDatabase(be, dbLangGeneral)
    db.Execute "CREATE TABLE pg_parent (id LONG PRIMARY KEY, name TEXT(60))", dbFailOnError
    db.Execute "CREATE TABLE pg_child (id LONG PRIMARY KEY, parent_id LONG REFERENCES pg_parent (id))", dbFailOnError
    db.Execute "INSERT INTO pg_parent VALUES (1, 'one')", dbFailOnError
    db.Execute "INSERT INTO pg_parent VALUES (2, 'two')", dbFailOnError
    db.Execute "INSERT INTO pg_child VALUES (1, 1)", dbFailOnError
    db.Close
    KitLinkTables be, Array("pg_parent", "pg_child")
    Check KitDataFileUsers(be) = "", "This computer listed as another user of the data file"
    bak = KitBackupDataFile(be, root & "\" & BACKUP_FOLDER)
    Check FileExists(bak) And InStr(bak, "\" & BACKUP_FOLDER & "\purge_be_backup_") > 0, "Backup copy not saved in the Backups folder"
    Set db = DBEngine.OpenDatabase(bak, False, True)
    n = db.OpenRecordset("SELECT COUNT(*) FROM pg_parent").Fields(0)
    db.Close
    Check n = 2, "Backup copy lacks the data"
    ' A missing table after two that emptied: the transaction puts both back.
    pErr = 0: On Error Resume Next: KitPurgeTables be, Array("pg_child", "pg_parent", "pg_missing"): pErr = Err.Number: On Error GoTo 0
    Check pErr <> 0 And DCount("*", "pg_parent") = 2 And DCount("*", "pg_child") = 1, "A failed purge left some tables emptied"
    Check KitPurgeTables(be, Array("pg_child", "pg_parent")) = 3, "Purge didn't count the deleted rows"
    Check DCount("*", "pg_parent") = 0 And DCount("*", "pg_child") = 0, "Purge left rows behind"
    pErr = 0: On Error Resume Next: KitBackupDataFile be, root & "\" & BACKUP_FOLDER & "\" & Chr$(1): pErr = Err.Number: On Error GoTo 0
    Check pErr <> 0, "Backup into a folder that can't be made reported success"
    CurrentDb.TableDefs.Delete "pg_child": CurrentDb.TableDefs.Delete "pg_parent"
    Kill bak: RmDir root & "\" & BACKUP_FOLDER
    Kill be: RmDir root
    ' Everyone deleted from Personnel: no admin until the first one is made again.
    pErr = 0
    If Len(PersonnelPath()) > 0 Then
        KitPurgeTables PersonnelPath(), Array(PERSON_APP_TABLE, PERSON_TABLE), PersonnelKey()
        Check PersonnelStatus("no_table") = PERSONNEL_NO_ADMIN, "Personnel with nobody in it reported ready"
        n = AddFirstAdmin("TEST", "Pat", "O'Neil", "LT")
        Check n > 0 And PersonnelStatus("no_table") = PERSONNEL_READY And PersonCanUse(n, "TEST"), "First admin not made again"
        On Error Resume Next: AddFirstAdmin "TEST", "Sam", "Lee": pErr = Err.Number: On Error GoTo 0
        Check pErr = 5, "A second first admin was made while one exists"
    Else
        Check False, "No Personnel file linked for the Personnel purge checks"
    End If
End Sub

' Rows a form shows (its recordset clone read to the end).
Private Function RowCount(g As Form) As Long
    With g.RecordsetClone
        If .BOF And .EOF Then Exit Function
        .MoveLast
        RowCount = .RecordCount
    End With
End Function

' True when the current database has a table named nm.
Private Function HasTableNamed(ByVal nm As String) As Boolean
    Dim td As DAO.TableDef
    For Each td In CurrentDb.TableDefs
        If StrComp(td.Name, nm, vbTextCompare) = 0 Then HasTableNamed = True: Exit Function
    Next
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
