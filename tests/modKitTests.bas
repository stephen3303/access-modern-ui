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
    Dim order() As Long, shown() As Boolean, sf As Control
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
    host.fit2.Tag = host.fit2.Tag & "|id"                     ' as a sortable caption's
    SetGridColumns sf, "fit", "2,0,1h"
    Set grid = sf.Form
    Check Abs(grid.txtCol0.Left + grid.txtCol0.Width - height) <= px(2), "Shown columns did not fill the table as far as all of them"
    Check grid.txtCol2.Left < grid.txtCol0.Left, "Moved column did not move"
    Check Not grid.txtCol1.Visible And Not host.fit1.Visible, "Hidden column or caption still shows"
    Check Abs(host.fit2.Left - (sf.Left + grid.txtCol2.Left)) <= 15, "Caption did not follow its column"
    Check grid.txtCol0.Width > px(160), "Shown columns did not widen into the hidden one's space"
    Check grid.txtCol1.Width <= 15, "Hidden column kept its width (can push the form past its limit)"
    Check Right$(host.fit2.Tag, 3) = "|id", "Caption lost a later Tag part (sort field)"
    On Error Resume Next
    SetGridColumns sf, "fit", "0,1"
    Check Err.Number = 5, "Bad column layout was applied"
    On Error GoTo 0
    SetGridColumns sf, "fit", DefaultLayout(3)
    Check grid.txtCol0.Left < grid.txtCol1.Left And grid.txtCol1.Left < grid.txtCol2.Left, "Default layout lost design order"
    Check grid.txtCol1.Visible And host.fit1.Visible, "Default layout left a column hidden"
    DoCmd.Close acForm, "frmTestFit", acSaveNo
    TestKit = checks & " kit checks passed"
End Function
