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
    NewForm "", 0, px(700), px(700)
    GridHeaders "head", cols, px(24), px(60)
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
    grid.txtCol0.SetFocus: grid.RowClick
    Check Screen.ActiveControl.Name = "txtFocusSink", "Row click left text focused"
    Check host.callbackGrid = "sfRows", "Grid callback used wrong parent"
    Check CLng(grid.txtSel) = CLng(grid!id), "Row click lost record selection"
    host.sfReuse.Form.RecordSource = "SELECT * FROM test_rows ORDER BY id"
    host.sfReuse.Form.txtCol0.SetFocus: host.sfReuse.Form.RowClick
    Check host.callbackGrid = "sfReuse", "Reused grid callback selected first instance"
    ResizeGrid grid, px(200), cols
    Check grid.ScrollBars = 3, "Narrow grid clips minimum columns instead of scrolling"
    Check grid.txtCol1.Width >= px(292), "Narrow grid discarded Description width"
    DoCmd.Close acForm, "frmTest", acSaveNo
    BuildMessageDialog
    DoCmd.OpenForm "dlgMessage"
    Set host = Forms("dlgMessage")
    Check host.btn1.Caption = "OK" And host.btn1.Visible And host.btn1.Default And host.btn1.Cancel, "Message dialog lacks its OK button"
    Check Not host.btn2.Visible And Not host.btn3.Visible, "Message dialog shows unused buttons"
    Check host.btn1.Left + host.btn1.Width = host.btn3.Left + host.btn3.Width, "Message OK button is not right-aligned"
    DoCmd.Close acForm, "dlgMessage", acSaveNo
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
    TestKit = checks & " kit checks passed"
End Function
