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
    Dim db As DAO.Database, badName As Long, badScale As Long
    g_SrcRoot = root & "\kit": UseBuildPalette
    CurrentDb.Execute "CREATE TABLE test_rows (id LONG PRIMARY KEY, category TEXT(60), description TEXT(255))"
    For i = 1 To 23
        CurrentDb.Execute "INSERT INTO test_rows VALUES (" & i & ",'Long category','Description " & i & "')"
    Next
    cols = Array("category|Category|160|c", "description|Description|300|t")
    BuildGrid "subTest", cols, px(600)
    NewForm "", 0, px(200), px(60)
    Txt "ringC", "", 0, 0, px(100), px(30)
    Btn "ringSave", "Save", px(110), 0, px(80), px(30)
    SaveForm "subRing"
    NewForm "", 0, px(700), px(700)
    GridHeaders "head", cols, px(24), px(60)
    GridHeaders "sortHead", cols, px(24), px(40), sortSf:="sfRows"
    SubF "sfRows", "Form.subTest", px(24), px(82), px(600), px(180)
    SubF "sfReuse", "Form.subTest", px(24), px(320), px(600), px(180)
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
    SaveForm "frmTest", "testHost"
    DoCmd.OpenForm "frmTest"
    Set host = Forms("frmTest"): Set grid = host.sfRows.Form
    Check FitSingleLineText(Null, px(100)) = "", "Null presentation text differs"
    Check FitSingleLineText("A" & vbCrLf & "B" & vbLf & "C", px(100)) = "A B C", "Presentation text retains line breaks"
    measured = MeasureText("Exact fit", 9, 400)
    Check FitSingleLineText("Exact fit", measured) = "Exact fit", "Exact-fit text was truncated"
    Check FitSingleLineText("Text", 0) = "", "Zero-width text overflows"
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
    grid.RecordSource = "SELECT * FROM test_rows ORDER BY id"
    ' Click-to-sort captions: ascending with an up arrow, then reversed; the arrow moves with the column.
    Check host.head1.OnClick = "" And Right$(host.head1.Tag, 1) = "|", "Unsortable caption has a sort click"
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
    host.reqA = "   "
    host.reqB = Null
    text = RequiredMissing(host, "reqA=First;reqB=Second", fitted)
    Check text = ChrW(&H2022) & " First" & vbCrLf & ChrW(&H2022) & " Second", "Required list differs"
    Check fitted = "reqA", "First missing required field differs"
    Check host.reqA.BorderColor = Th(CLR_DESTRUCTIVE), "Missing required field not outlined"
    host.reqA = "x"
    Check RequiredMissing(host, "reqA=First", fitted) = "" And fitted = "", "Filled required field reported"
    Check host.reqA.BorderColor = Th(CLR_INPUT_BORDER), "Filled required field kept its outline"
    Check SaveDetail(host, "sfRows"), "Unchanged detail panel reported a refused save"
    DoCmd.Close acForm, "frmTest", acSaveNo
    BuildMessageDialog
    DoCmd.OpenForm "dlgMessage"
    Set host = Forms("dlgMessage")
    Check host.btn1.Caption = "OK" And host.btn1.Visible And host.btn1.Default And host.btn1.Cancel, "Message dialog lacks its OK button"
    Check Not host.btn2.Visible And Not host.btn3.Visible, "Message dialog shows unused buttons"
    Check host.btn1.Left + host.btn1.Width = host.btn3.Left + host.btn3.Width, "Message OK button is not right-aligned"
    DoCmd.Close acForm, "dlgMessage", acSaveNo
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
    TestKit = checks & " kit checks passed"
End Function
