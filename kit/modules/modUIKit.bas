Option Compare Database
Option Explicit

' Access UI kit - build-time helpers that create forms and styled
' controls (a small stand-in for the Tailwind/shadcn components of the web UI).
' Units are twips; px(n) converts CSS pixels (15 twips per px at 96 dpi).

Public F As Object            ' form or report under construction
Public g_SrcRoot As String
Private m_isReport As Boolean
Private m_trace As String

' CSS pixels to twips.
' Printable width of a report page: 8.5 in less two 0.5 in margins (twips).
Public Const REPORT_W As Long = 10800
' Corner radius of rounded cards (px).
Public Const CARD_RADIUS As Long = 12
' DateBox: the text box (120 px, fits "dd mmm yyyy" at 10pt) and the whole
' field with its 30px calendar button 4px to the right (154 px).
Public Const DATEBOX_W As Long = 1800
Public Const DATEFIELD_W As Long = 2310
' Badge variants a status pill can take (BadgeColors).
Private Const BADGE_VARIANTS As String = "success,warning,info,destructive,default,outline,secondary"
' Spacing tokens (CSS px, guide 6.4) for Sp(), so a layout written with them
' follows the spacing the designer picked with SetSpacing.
Public Const SP_PAGE As Long = 32         ' page padding, left and right
Public Const SP_DIALOG As Long = 24       ' dialog padding
Public Const SP_CARD_INSET As Long = 16   ' controls inside a card, from its edge
Public Const SP_CARD_GAP As Long = 12     ' between cards and KPI tiles
Public Const SP_ROW_GAP As Long = 8       ' input bottom to the next caption or row
Public Const SP_ITEM_GAP As Long = 4      ' between chips, filter pills and tabs
' Database property that keeps the spacing chosen at build time for run-time layouts.
Private Const SPACING_PROP As String = "KitSpacing"
Private m_spacing As Double               ' 0 until read or set

Public Function px(ByVal n As Double) As Long
    px = CLng(n * 15)
End Function

' ---------------------------------------------------------------------------
' Spacing
' ---------------------------------------------------------------------------

' The designer picks how roomy layouts are, once, before building forms:
'   SetSpacing "compact"      0.75  (tighter, for dense data entry)
'   SetSpacing "default"      1     (the guide's 6.4 values)
'   SetSpacing "comfortable"  1.25
' or a scale from 0.5 to 2 (SetSpacing 1.1). Without a call the scale is 1,
' so existing layouts (SPM, SABER) build exactly as before. Only the space
' between and around controls scales; control sizes (30px inputs, 32px
' buttons) and font sizes stay as the guide sets them.
' The choice is saved in the database, so run-time layouts such as
' LayoutFilterRow match the build. Raises error 5 for an unknown name or a
' scale out of range.
Public Sub SetSpacing(ByVal spacing As Variant)
    Dim factor As Double, db As DAO.Database, prop As DAO.Property, n As Long, d As String
    Select Case LCase$(Trim$(CStr(spacing)))
        Case "compact": factor = 0.75
        Case "default", "": factor = 1
        Case "comfortable": factor = 1.25
        Case Else
            If Not IsNumeric(spacing) Then Err.Raise 5, , "Spacing must be compact, default, comfortable or a scale: " & spacing
            factor = CDbl(spacing)
            If factor < 0.5 Or factor > 2 Then Err.Raise 5, , "Spacing scale must be from 0.5 to 2: " & spacing
    End Select
    Set db = CurrentDb
    On Error Resume Next
    db.Properties(SPACING_PROP) = factor
    If Err.Number = 3270 Then         ' property not found: create it
        Err.Clear
        Set prop = db.CreateProperty(SPACING_PROP, dbDouble, factor)
        db.Properties.Append prop
    End If
    If Err.Number <> 0 Then
        n = Err.Number: d = Err.Description
        On Error GoTo 0
        Err.Raise n, , "Could not save the spacing: " & d
    End If
    On Error GoTo 0
    m_spacing = factor
End Sub

' The spacing scale in use: the one SetSpacing saved, else 1.
Public Function SpacingScale() As Double
    Dim db As DAO.Database
    If m_spacing = 0 Then
        m_spacing = 1
        Set db = CurrentDb
        On Error Resume Next
        m_spacing = CDbl(db.Properties(SPACING_PROP))
        If m_spacing < 0.5 Or m_spacing > 2 Then m_spacing = 1
        On Error GoTo 0
    End If
    SpacingScale = m_spacing
End Function

' A gap or padding of n CSS px at the chosen spacing, in twips, rounded to
' whole px so edges stay crisp. Use it for space, px() for control sizes:
'   x = x + w + Sp(SP_ITEM_GAP)
Public Function Sp(ByVal n As Double) As Long
    Sp = px(Int(n * SpacingScale() + 0.5))
End Function

' Last build step recorded with Trace, reported when the build fails.
Public Function BuildTrace() As String
    BuildTrace = m_trace
End Function

Public Sub Trace(ByVal s As String)
    m_trace = s
End Sub

' ---------------------------------------------------------------------------
' Forms
' ---------------------------------------------------------------------------

' view: 0 = single, 1 = continuous
Public Sub NewForm(ByVal recSrc As String, ByVal view As Integer, ByVal w As Long, ByVal h As Long)
    Set F = CreateForm()
    m_isReport = False
    With F
        .RecordSource = recSrc
        .DefaultView = view
        .AllowDatasheetView = False
        .AllowLayoutView = False
        .RecordSelectors = False
        .NavigationButtons = False
        .DividingLines = False
        .ScrollBars = IIf(view = 1, 2, 0)
        .AllowAdditions = False
        .AllowDeletions = False
        .Width = w
        .Section(acDetail).Height = h
        .Section(acDetail).BackColor = CLR_BG
        .Section(acDetail).AlternateBackColor = CLR_BG
        .AutoResize = False
        .HasModule = True
        .OnOpen = "=ThemeForm([Form])"     ' color palette (modDarkMode)
    End With
    ' Focus sink: takes the initial focus so Access does not highlight the
    ' first text control (which hides its text, most visibly in the dark theme).
    Dim sink As Control
    Set sink = CreateControl(F.Name, acTextBox, acDetail, "", "", 0, 0, 15, 15)
    sink.Name = "txtFocusSink"
    sink.BackStyle = 0
    sink.BorderStyle = 0
    sink.TabIndex = 0
End Sub

' Read-only text in list rows that has no click action can't take focus, so
' clicking it doesn't highlight the text or paint the box white. (Disabled +
' locked keeps the normal look.) Clickable row text moves focus to
' txtFocusSink in its click handler instead.
Private Sub LockRowText()
    Dim c As Control
    For Each c In F.Controls
        If c.ControlType = acTextBox Then
            If c.Name <> "txtFocusSink" And c.Locked And Not c.TabStop And Len(c.OnClick) = 0 Then c.Enabled = False
        End If
    Next
End Sub

' Modal pop-up dialog form that sizes to its design.
Public Sub NewDialog(ByVal caption As String, ByVal w As Long, ByVal h As Long)
    NewForm "", 0, w, h
    With F
        .Caption = caption
        .PopUp = True
        .Modal = True
        .BorderStyle = 3
        .AutoCenter = True
        .AutoResize = True
        .MinMaxButtons = 0
        .ControlBox = True
        .CloseButton = True
        .Section(acDetail).BackColor = CLR_DIALOG
        .Section(acDetail).AlternateBackColor = CLR_DIALOG
    End With
End Sub

' Save the form under its final name, attaching code-behind from src\forms\<name>.vba.
Public Sub SaveForm(ByVal finalName As String, Optional ByVal codeName As String = "")
    Dim tmp As String, codeFile As String, code As String
    Trace finalName
    codeFile = g_SrcRoot & "\forms\" & IIf(Len(codeName) > 0, codeName, finalName) & ".vba"
    If FileExists(codeFile) Then
        code = Replace(Replace(ReadTextFile(codeFile), vbCrLf, vbLf), vbLf, vbCrLf)
        With F.Module
            If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
            .AddFromString code
        End With
    End If
    If Not m_isReport Then
        If F.DefaultView = 1 Then LockRowText
    End If
    tmp = F.Name
    If m_isReport Then
        DoCmd.Close acReport, tmp, acSaveYes
        On Error Resume Next
        DoCmd.DeleteObject acReport, finalName
        On Error GoTo 0
        DoCmd.Rename finalName, acReport, tmp
    Else
        DoCmd.Close acForm, tmp, acSaveYes
        On Error Resume Next
        DoCmd.DeleteObject acForm, finalName
        On Error GoTo 0
        DoCmd.Rename finalName, acForm, tmp
    End If
    Set F = Nothing
End Sub

' Start a report; controls added with the helpers below go onto it.
Public Sub NewReport(ByVal recSrc As String, ByVal w As Long)
    Set F = CreateReport()
    m_isReport = True
    With F
        .RecordSource = recSrc
        .Width = w
        .HasModule = True
        ' Previews open in their own window with a title bar and Close button:
        ' the ribbon and document tabs are hidden, so a docked preview has no
        ' visible way to close it.
        .PopUp = True
        .AutoCenter = True
        ' Letter portrait with equal half-inch margins; reports are built
        ' REPORT_W wide so the body fills the page between them.
        .Printer.PaperSize = acPRPSLetter
        .Printer.Orientation = acPRORPortrait
        .Printer.LeftMargin = 720
        .Printer.RightMargin = 720
        .Printer.TopMargin = 720
        .Printer.BottomMargin = 720
    End With
End Sub

' ---------------------------------------------------------------------------
' Controls
' ---------------------------------------------------------------------------

' Create a control on the form or report under construction.
Private Function Mk(ByVal kind As Integer, ByVal nm As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, _
                    ByVal h As Long, ByVal sec As Integer, ByVal parentCtl As String) As Control
    Dim c As Control
    If m_isReport Then
        Set c = CreateReportControl(F.Name, kind, sec, parentCtl, "", x, y, w, h)
    Else
        Set c = CreateControl(F.Name, kind, sec, parentCtl, "", x, y, w, h)
    End If
    c.Name = nm
    Set Mk = c
End Function

' Vertically center single-line text: Access has no vertical alignment, so pad
' the top by half the free space. Taller (multi-line) controls stay top-aligned.
Public Sub VCenter(c As Control, ByVal h As Long, ByVal size As Integer, Optional ByVal force As Boolean = False)
    Dim lineH As Long, m As Long
    lineH = CLng(size * 20 * 1.36)
    If h >= lineH * 1.9 And Not force Then Exit Sub
    m = (h - lineH) \ 2 - 10
    If m < 0 Then m = 0
    On Error Resume Next
    c.TopMargin = m
    c.BottomMargin = 0
End Sub

' UI font, size, weight and color.
Private Sub Font(c As Control, ByVal size As Integer, ByVal bold As Boolean, ByVal fore As Long)
    c.FontName = FONT_UI
    c.FontSize = size
    c.FontWeight = IIf(bold, 600, 400)
    c.ForeColor = fore
End Sub

' Label (transparent background).
Public Function Lbl(ByVal nm As String, ByVal cap As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
        Optional ByVal size As Integer = 10, Optional ByVal bold As Boolean = False, Optional ByVal fore As Long = CLR_FG, _
        Optional ByVal sec As Integer = acDetail, Optional ByVal parentCtl As String = "") As Control
    Dim c As Control
    Set c = Mk(acLabel, nm, x, y, w, h, sec, parentCtl)
    c.Caption = cap
    c.BackStyle = 0
    c.BorderStyle = 0
    Font c, size, bold, fore
    VCenter c, h, size
    Set Lbl = c
End Function

' Text box. src is a field name or an expression starting with "=".
Public Function Txt(ByVal nm As String, ByVal src As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
        Optional ByVal size As Integer = 10, Optional ByVal sec As Integer = acDetail, Optional ByVal parentCtl As String = "") As Control
    Dim c As Control
    Set c = Mk(acTextBox, nm, x, y, w, h, sec, parentCtl)
    If Len(src) > 0 Then c.ControlSource = src
    Font c, size, False, CLR_FG
    c.BorderStyle = 1
    c.BorderColor = CLR_INPUT_BORDER
    c.SpecialEffect = 0
    c.BackStyle = 1
    c.BackColor = CLR_INPUT_BG
    c.LeftMargin = 60
    c.TopMargin = 30
    VCenter c, h, size
    Set Txt = c
End Function

' Date entry: a DATEBOX_W text box nm bound to src (dd mmm yyyy, left-aligned
' like text, Access's own picker icon off) and a calendar button cal_<nm>
' 4px to its right, h tall like the box, that opens the kit's date picker for
' it (PickDate, dlgDatePick: build it with BuildDatePicker). The pair takes
' DATEFIELD_W. Returns the text box.
Public Function DateBox(ByVal nm As String, ByVal src As String, ByVal x As Long, ByVal y As Long, ByVal h As Long, _
        Optional ByVal sec As Integer = acDetail, Optional ByVal parentCtl As String = "") As Control
    Dim c As Control, b As Control
    Set c = Txt(nm, src, x, y, DATEBOX_W, h, 10, sec, parentCtl)
    c.Format = "dd mmm yyyy"
    c.TextAlign = 1
    c.ShowDatePicker = 0
    Set b = Btn("cal_" & nm, ChrW(&HE787), x + DATEBOX_W + px(4), y, px(30), h, "outline", sec, parentCtl)
    b.FontName = "Segoe MDL2 Assets"
    b.FontSize = 10
    b.FontWeight = 400
    b.TabStop = False
    b.ControlTipText = "Choose a date"
    b.OnClick = "=PickDate([Form],""" & nm & """)"
    Set DateBox = c
End Function

' The date picker a DateBox's calendar button opens (dlgDatePick, code in
' forms\dlgDatePick.vba): previous and next month arrows around the month's
' name, the weekday initials, six weeks of day cells (34 x 30px on a 34 x 32
' grid), and Today, Clear and Cancel. Build it with the app's other dialogs.
Public Sub BuildDatePicker()
    Dim c As Control, i As Long
    NewDialog "Choose a date", px(270), px(316)
    Ev F, "OnLoad"
    Set c = Link("lnkPrev", ChrW(&HE76B), px(16), px(12), px(28), px(28), "=MoveMonth(-1)", 10, CLR_FG)
    c.FontName = "Segoe MDL2 Assets": c.TextAlign = 2: c.ControlTipText = "Previous month"
    Set c = Lbl("lblMonth", "Month", px(44), px(12), px(182), px(28), 11, True)
    c.TextAlign = 2
    Set c = Link("lnkNext", ChrW(&HE76C), px(226), px(12), px(28), px(28), "=MoveMonth(1)", 10, CLR_FG)
    c.FontName = "Segoe MDL2 Assets": c.TextAlign = 2: c.ControlTipText = "Next month"
    For i = 0 To 6
        Set c = Lbl("wd" & i, Mid$("SuMoTuWeThFrSa", i * 2 + 1, 2), px(16) + i * px(34), px(48), px(34), px(18), 8, True, CLR_MUTED_FG)
        c.TextAlign = 2
    Next
    For i = 0 To 41
        Set c = Link("d" & i, "1", px(16) + (i Mod 7) * px(34), px(70) + (i \ 7) * px(32), px(34), px(30), _
            "=PickDay(" & i & ")", 10, CLR_FG)
        c.TextAlign = 2
    Next
    Link "lnkToday", "Today", px(16), px(273), px(56), px(24), "=PickToday()", 9
    Link "lnkClear", "Clear", px(76), px(273), px(56), px(24), "=PickClear()", 9
    Set c = Btn("btnCancel", "Cancel", px(164), px(270), px(90), px(30), "outline"): c.Cancel = True
    SaveForm "dlgDatePick"
End Sub

' The message dialog KitMsgBox opens (dlgMessage, code in forms\dlgMessage.vba):
' an icon, the message in 10pt text wrapping across a 336px column (room for
' 14 lines, scrolling for more), and up to three 88px buttons along the bottom
' right. dlgMessage sets the icon and buttons for each message, and shrinks the
' dialog to the lines the message takes. Build it with the app's other dialogs.
Public Sub BuildMessageDialog()
    Dim c As Control, msgH As Long, btnY As Long, i As Long
    msgH = 14 * CLng(10 * 20 * 1.36)                ' 14 lines of 10pt text
    btnY = px(22) + msgH + px(20)
    NewDialog "Message", px(420), btnY + px(30) + px(16)
    F.CloseButton = False                           ' Esc or a button closes it, like MsgBox
    Ev F, "OnLoad"
    Set c = Lbl("lblIcon", ChrW(&HE946), px(20), px(20), px(32), px(32), 20, False, CLR_PRIMARY)
    c.FontName = "Segoe MDL2 Assets": c.FontWeight = 400: c.TextAlign = 2
    ' Read-only text box rather than a label, so a longer message can scroll.
    Set c = RO("txtMsg", "", px(64), px(22), px(336), msgH, 10, False, CLR_FG)
    c.TopMargin = 0: c.LeftMargin = 0: c.RightMargin = 0: c.BottomMargin = 0
    c.ScrollBars = 2                                ' shows only while it has the focus
    c.BackColor = CLR_DIALOG                        ' painted while it has the focus
    For i = 1 To 3
        Set c = Btn("btn" & i, "OK", px(420) - px(20) - (4 - i) * px(88) - (3 - i) * px(8), btnY, px(88), px(30), _
            IIf(i = 1, "primary", "outline"))
    Next
    SaveForm "dlgMessage"
End Sub

' Splash screen frmSplash, shown while the app starts: the app's icon glyph,
' APP_NAME, a subtitle, its version (database property AppVersion, else the
' AppBuild stamp) and a status line (SplashStatus), on the sidebar color of
' the user's palette (ThemeForm repaints it). Make it the StartUpForm. Its
' code-behind (kit\forms\frmSplash.vba) opens nextForm, whose Load does the
' app's start-up work under the splash, then closes. Call it once with the
' app's dialogs, and HideAccessSplash for Access's own splash.
' glyph: a Segoe MDL2 Assets character; default ChrW(&HE80F), home.
Public Sub BuildSplash(Optional ByVal nextForm As String = "frmMain", Optional ByVal subtitle As String = "", _
        Optional ByVal glyph As String = "")
    Dim c As Control, w As Long, h As Long, pad As Long, y As Long
    ' 480 x 280px at the default spacing. Padding and gaps follow SetSpacing,
    ' so a roomier scale makes the splash taller rather than overlapping the
    ' subtitle and the footer.
    w = px(480): pad = Sp(SP_PAGE)
    h = pad + px(48) + Sp(16) + px(36) + Sp(SP_ITEM_GAP) + px(20) _
        + Sp(SP_DIALOG) + Sp(SP_ROW_GAP) + px(1) + px(18) + pad
    If h < px(280) Then h = px(280)
    NewForm "", 0, w, h
    With F
        .Caption = APP_NAME
        .Tag = nextForm                             ' read by frmSplash.vba
        .PopUp = True                               ' stays over the window nextForm opens in
        .Modal = False
        .BorderStyle = 0
        .AutoCenter = True
        .AutoResize = True
        .ControlBox = False
        .MinMaxButtons = 0
        .CloseButton = False
        .ShortcutMenu = False
        .Section(acDetail).BackColor = CLR_SIDEBAR
        .Section(acDetail).AlternateBackColor = CLR_SIDEBAR
    End With
    Ev F, "OnLoad", "OnTimer"
    If Len(glyph) = 0 Then glyph = ChrW(&HE80F)
    Set c = Lbl("lblSplashIcon", glyph, pad, pad, px(48), px(48), 28, False, CLR_WHITE)
    c.FontName = "Segoe MDL2 Assets"
    y = pad + px(48) + Sp(16)
    Lbl "lblSplashName", APP_NAME, pad, y, w - 2 * pad, px(36), 20, True, CLR_WHITE
    y = y + px(36) + Sp(SP_ITEM_GAP)
    Lbl "lblSplashSub", subtitle, pad, y, w - 2 * pad, px(20), 10, False, CLR_NAV_FG
    ' Footer: a hairline, then the status line left and the version right.
    y = h - pad - px(18)
    Rule "lnSplash", pad, y - Sp(SP_ROW_GAP) - px(1), w - 2 * pad, 0, CLR_SIDEBAR_LINE
    Lbl "lblSplashStatus", "Starting" & ChrW(8230), pad, y, w - 2 * pad - px(180) - Sp(SP_ROW_GAP), px(18), 9, False, CLR_NAV_MUTED
    Set c = Lbl("lblSplashVersion", "", w - pad - px(180), y, px(180), px(18), 9, False, CLR_NAV_MUTED)
    c.TextAlign = 3
    SaveForm "frmSplash"
End Sub

' Hide Access's own splash screen. When a database is opened, Access shows
' <database name>.bmp from the database's folder in place of its logo, so
' this writes a 1 x 1 pixel bitmap there (58 bytes, rewritten each build).
' Call it from BuildAll's startup step and ship the .bmp with the front end.
Public Sub HideAccessSplash()
    Dim path As String, b(57) As Byte, i As Long, fh As Integer
    path = CurrentProject.Name
    If InStrRev(path, ".") > 0 Then path = Left$(path, InStrRev(path, ".") - 1)
    path = CurrentProject.Path & "\" & path & ".bmp"
    ' BITMAPFILEHEADER (14 bytes) + BITMAPINFOHEADER (40) + one 24-bit pixel
    ' padded to 4 bytes. Unset bytes stay 0; the pixel is black.
    b(0) = 66: b(1) = 77                ' "BM"
    b(2) = 58                           ' file size
    b(10) = 54                          ' offset of the pixel data
    b(14) = 40                          ' info header size
    b(18) = 1: b(22) = 1                ' width 1, height 1
    b(26) = 1                           ' planes
    b(28) = 24                          ' bits per pixel
    b(34) = 4                           ' pixel data size
    For i = 38 To 42 Step 4
        b(i) = &H13: b(i + 1) = &HB     ' 2835 pixels per metre (72 dpi)
    Next
    If FileExists(path) Then Kill path
    fh = FreeFile
    Open path For Binary Access Write As #fh
    Put #fh, , b
    Close #fh
End Sub

' Background box for a continuous-form row: paints (with conditional colors)
' but can never take focus. Disabled + locked keeps the normal appearance.
Public Function RowBack(ByVal nm As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As Control
    Dim c As Control
    Set c = RO(nm, "", x, y, w, h)
    c.Enabled = False
    c.Locked = True
    c.BackStyle = 1
    c.BackColor = CLR_SURFACE
    Set RowBack = c
End Function

' Read-only text: no border, transparent, not a tab stop.
Public Function RO(ByVal nm As String, ByVal src As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
        Optional ByVal size As Integer = 10, Optional ByVal bold As Boolean = False, Optional ByVal fore As Long = CLR_FG, _
        Optional ByVal sec As Integer = acDetail, Optional ByVal parentCtl As String = "") As Control
    Dim c As Control
    Set c = Mk(acTextBox, nm, x, y, w, h, sec, parentCtl)
    If Len(src) > 0 Then c.ControlSource = src
    Font c, size, bold, fore
    c.BorderStyle = 0
    c.BackStyle = 0
    c.SpecialEffect = 0
    If Not m_isReport Then
        c.Locked = True
        c.TabStop = False
    End If
    VCenter c, h, size
    Set RO = c
End Function

' Combo box with a row source (SQL or value list).
Public Function Cbo(ByVal nm As String, ByVal src As String, ByVal rowSrc As String, ByVal x As Long, ByVal y As Long, _
        ByVal w As Long, ByVal h As Long, Optional ByVal cols As Integer = 1, Optional ByVal widths As String = "", _
        Optional ByVal sec As Integer = acDetail, Optional ByVal parentCtl As String = "") As Control
    Dim c As Control
    Set c = Mk(acComboBox, nm, x, y, w, h, sec, parentCtl)
    If Len(src) > 0 Then c.ControlSource = src
    If Left$(rowSrc, 6) = "SELECT" Or Len(rowSrc) = 0 Then
        c.RowSourceType = "Table/Query"
    Else
        c.RowSourceType = "Value List"
    End If
    c.RowSource = rowSrc
    c.ColumnCount = cols
    If Len(widths) > 0 Then c.ColumnWidths = widths
    c.BoundColumn = 1
    c.LimitToList = True
    c.ListRows = 16
    c.AllowValueListEdits = False
    Font c, 10, False, CLR_FG
    c.BorderStyle = 1
    c.BorderColor = CLR_INPUT_BORDER
    c.SpecialEffect = 0
    c.BackColor = CLR_INPUT_BG
    VCenter c, h, 10
    Set Cbo = c
End Function

' List box.
Public Function Lst(ByVal nm As String, ByVal rowSrc As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
        ByVal cols As Integer, ByVal widths As String, Optional ByVal sec As Integer = acDetail, Optional ByVal parentCtl As String = "") As Control
    Dim c As Control
    Set c = Mk(acListBox, nm, x, y, w, h, sec, parentCtl)
    c.RowSourceType = IIf(Left$(rowSrc, 6) = "SELECT" Or Len(rowSrc) = 0, "Table/Query", "Value List")
    c.RowSource = rowSrc
    c.ColumnCount = cols
    c.ColumnWidths = widths
    c.ColumnHeads = True
    c.BoundColumn = 1
    Font c, 9, False, CLR_FG
    c.BorderStyle = 1
    c.BorderColor = CLR_BORDER
    c.SpecialEffect = 0
    c.BackColor = CLR_INPUT_BG
    Set Lst = c
End Function

' Check box with its caption label.
Public Function Chk(ByVal nm As String, ByVal src As String, ByVal x As Long, ByVal y As Long, _
        Optional ByVal sec As Integer = acDetail, Optional ByVal parentCtl As String = "") As Control
    Dim c As Control
    Set c = Mk(acCheckBox, nm, x, y, 260, 260, sec, parentCtl)
    If Len(src) > 0 Then c.ControlSource = src Else c.DefaultValue = "False"
    c.SpecialEffect = 0
    Set Chk = c
End Function

' kind: primary | outline | ghost | destructive | nav
Public Function Btn(ByVal nm As String, ByVal cap As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
        Optional ByVal kind As String = "primary", Optional ByVal sec As Integer = acDetail, Optional ByVal parentCtl As String = "") As Control
    Dim c As Control
    Set c = Mk(acCommandButton, nm, x, y, w, h, sec, parentCtl)
    c.Caption = cap
    c.UseTheme = True
    c.FontName = FONT_UI
    c.FontSize = 9
    c.FontWeight = 600
    c.BorderStyle = 1
    Select Case kind
        Case "primary"
            c.BackColor = CLR_BTN_BG: c.HoverColor = CLR_BTN_HOVER: c.PressedColor = CLR_BTN_PRESSED
            c.ForeColor = CLR_BTN_FG: c.HoverForeColor = CLR_BTN_FG: c.PressedForeColor = CLR_BTN_FG
            c.BorderColor = CLR_BTN_BG
        Case "destructive"
            c.BackColor = CLR_DESTRUCTIVE: c.HoverColor = CLR_DESTRUCTIVE_HOVER: c.PressedColor = CLR_DESTRUCTIVE_PRESSED
            c.ForeColor = CLR_WHITE: c.HoverForeColor = CLR_WHITE: c.PressedForeColor = CLR_WHITE
            c.BorderColor = CLR_DESTRUCTIVE
        Case "ghost"
            c.BackColor = CLR_OUT_BG: c.HoverColor = CLR_OUT_HOVER: c.PressedColor = CLR_BORDER
            c.ForeColor = CLR_OUT_FG: c.HoverForeColor = CLR_OUT_FG: c.PressedForeColor = CLR_OUT_FG
            c.BorderStyle = 0
        Case Else ' outline
            c.BackColor = CLR_OUT_BG: c.HoverColor = CLR_OUT_HOVER: c.PressedColor = CLR_BORDER
            c.ForeColor = CLR_OUT_FG: c.HoverForeColor = CLR_OUT_FG: c.PressedForeColor = CLR_OUT_FG
            c.BorderColor = CLR_OUT_BORDER
    End Select
    ' Rounded rectangle (Access 2010+ themed buttons).
    On Error Resume Next
    c.Properties("Shape") = 1
    On Error GoTo 0
    c.OnClick = "[Event Procedure]"
    Set Btn = c
End Function

' Rectangle. back / borderClr = -1 for none.
Public Function Box(ByVal nm As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
        ByVal back As Long, Optional ByVal borderClr As Long = -1, Optional ByVal sec As Integer = acDetail, _
        Optional ByVal parentCtl As String = "") As Control
    Dim c As Control
    Set c = Mk(acRectangle, nm, x, y, w, h, sec, parentCtl)
    c.SpecialEffect = 0
    If back = -1 Then
        c.BackStyle = 0
    Else
        c.BackStyle = 1
        c.BackColor = back
    End If
    If borderClr = -1 Then
        c.BorderStyle = 0
    Else
        c.BorderStyle = 1
        c.BorderColor = borderClr
    End If
    Set Box = c
End Function

' A 1px rule (hairline line control) from (x, y), w long and flat (h = 0),
' or h long and upright (w = 0), in color. Use it, not a 1px filled Box, for
' dividers and edges: Access paints a filled 15-twip rectangle 2 device
' pixels thick, a hairline always 1.
Public Function Rule(ByVal nm As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
        ByVal color As Long, Optional ByVal sec As Integer = acDetail, Optional ByVal parentCtl As String = "") As Control
    Dim c As Control
    Set c = Mk(acLine, nm, x, y, w, h, sec, parentCtl)
    c.BorderStyle = 1
    c.BorderWidth = 0
    c.BorderColor = color
    Set Rule = c
End Function

' Subform control showing the form srcObj.
Public Function SubF(ByVal nm As String, ByVal srcObj As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
        Optional ByVal sec As Integer = acDetail, Optional ByVal parentCtl As String = "") As Control
    Dim c As Control
    Set c = Mk(acSubform, nm, x, y, w, h, sec, parentCtl)
    c.SourceObject = srcObj
    c.BorderStyle = 0
    c.SpecialEffect = 0
    Set SubF = c
End Function

' ---------------------------------------------------------------------------
' Styled tables
' ---------------------------------------------------------------------------

' A table in the look of the Team Members list: a continuous form of
' read-only columns with row lines and a highlighted selected row, in place
' of a list box. cols holds one "field|caption|width px|options" per column;
' options: "r" right-aligned, "c" centered, "b" bold, "m" muted, "p" a
' rounded status pill colored by the app's StatusVariant(value), "w" wraps
' onto rowLines lines (rather than cutting text off mid-word; text longer
' than that stops at the end of the last line). The last
' column takes the rest of the width (and stretches with the form); with
' rowLines > 1 it always wraps. back is the color of what the
' table sits on. zebra fills every other row with CLR_ROW2. rowPx sets the
' row height (default: from rowLines). fit = True keeps the last column at
' its own width, for tables the host scales with GridFit. The record source
' needs an id column. The rows share forms\subGrid.vba; GridHeaders puts the
' captions on the parent form. Each column's Tag holds "caption|left|width"
' for GridFit and modGridOut.
Public Sub BuildGrid(ByVal finalName As String, ByVal cols As Variant, ByVal w As Long, Optional ByVal rowLines As Long = 1, _
                     Optional ByVal back As Long = CLR_SURFACE, Optional ByVal zebra As Boolean = False, _
                     Optional ByVal rowPx As Long = 0, Optional ByVal fit As Boolean = False)
    Dim c As Control, i As Long, p() As String, x As Long, cw As Long, h As Long, last As Boolean, ty As Long, th As Long
    Dim opts As String, wrap As Boolean
    h = IIf(rowPx > 0, px(rowPx), px(12) + rowLines * px(18))
    NewForm "", 1, w, h
    F.Section(acDetail).BackColor = back
    F.Section(acDetail).AlternateBackColor = IIf(zebra, CLR_ROW2, back)
    F.Tag = "grid:" & w
    F.ScrollBars = 0                        ' GridScroll shows it when the rows overflow
    F.OnCurrent = "=GridScroll([Form])"
    ' A fit table's host sizes these (GridFit); anchored, they would stretch
    ' the form past the columns and stop GridFit narrowing it.
    Set c = RowBack("txtBg", 0, 0, w, h): If Not fit Then Anchor c, 2, 0
    c.BackColor = back
    CF c, "[id]=[txtSel]", CLR_PRIMARY_SOFT, -1
    ' Zebra rows show the section's alternating colors through the background;
    ' the selected row's condition still fills it.
    If zebra Then c.BackStyle = 0
    Set c = RO("txtSel", "", 0, 0, px(4), px(4)): c.Visible = False
    x = px(8)
    For i = 0 To UBound(cols)
        p = Split(cols(i), "|")
        opts = ""
        If UBound(p) >= 3 Then opts = p(3)
        last = (i = UBound(cols))
        cw = IIf(last And Not fit, w - x - px(8), px(Val(p(2))))
        If InStr(opts, "p") > 0 Then
            ' A pill fills its cell, so in the last column it stops 8px further
            ' in, clear of the table's scrollbar.
            Set c = RO("txtCol" & i, p(0), x, (h - px(20)) \ 2, cw - px(8) - IIf(last And Not fit, px(8), 0), px(20), 8, True, CLR_FG)
            c.TextAlign = 2
        Else
            wrap = rowLines > 1 And (last Or InStr(opts, "w") > 0)
            ty = IIf(wrap, px(6), (h - px(18)) \ 2)
            ' A wrapped cell is exactly rowLines lines of 9pt text (VCenter's
            ' line height) plus 2px for descenders, so longer text stops at a
            ' line's end instead of showing the top of the next line.
            th = IIf(wrap, rowLines * CLng(9 * 20 * 1.36) + px(2), px(18))
            Set c = RO("txtCol" & i, p(0), x, ty, cw - px(8), th, 9, InStr(opts, "b") > 0, _
                IIf(InStr(opts, "m") > 0, CLR_MUTED_FG, CLR_FG))
            If wrap Then c.TopMargin = 0
            If InStr(opts, "r") > 0 Then c.TextAlign = 3
            If InStr(opts, "c") > 0 Then c.TextAlign = 2
        End If
        c.Tag = p(1) & "|" & x & "|" & cw
        c.OnClick = "=RowClick()"
        c.OnDblClick = "=RowDblClick()"
        If last And Not fit Then Anchor c, 2, 0
        ' After the click and anchor: the pill image copies the anchoring.
        If InStr(opts, "p") > 0 Then PillCol c, p(0)
        x = x + cw
    Next
    Set c = Rule("boxRowLine", 0, h - px(1), w, 0, CLR_BORDER): If Not fit Then Anchor c, 2, 0
    SaveForm finalName, "subGrid"
End Sub

' Status pill colors for a grid column: one condition per badge variant,
' picked by the app's StatusVariant(value), drawn as a rounded pill.
Private Sub PillCol(c As Control, ByVal field As String)
    Dim v As Variant, col As Variant, first As Integer, ref As String
    If Left$(field, 1) = "=" Then
        ref = "(" & Mid$(field, 2) & ")"
    ElseIf Left$(field, 1) = "[" Then
        ref = field
    Else
        ref = "[" & field & "]"
    End If
    first = c.FormatConditions.Count
    For Each v In Split(BADGE_VARIANTS, ",")
        col = BadgeColors(CStr(v))
        CF c, "StatusVariant(" & ref & ")=""" & v & """", col(0), col(1)
    Next
    RoundCF c, first
End Sub

' Column captions for a BuildGrid table whose subform sits at (x, y + px(22)).
' band = True draws them on a CLR_INSET strip w wide (default: the columns'
' widths), which stretches with the window like the table. Each caption's
' Tag holds "x|offset|width" for GridFit.
Public Sub GridHeaders(ByVal prefix As String, ByVal cols As Variant, ByVal x As Long, ByVal y As Long, _
                       Optional ByVal parentCtl As String = "", Optional ByVal band As Boolean = False, _
                       Optional ByVal w As Long = 0)
    Dim c As Control, i As Long, p() As String, cx As Long, cy As Long, cw As Long, opts As String
    If band Then
        If w = 0 Then
            w = px(16)
            For i = 0 To UBound(cols)
                w = w + px(Val(Split(cols(i), "|")(2)))
            Next
        End If
        Set c = Box(prefix & "Band", x, y, w, px(22), CLR_INSET, -1, acDetail, parentCtl): Anchor c, 2, 0
    End If
    cx = x + px(8)
    cy = IIf(band, y + px(3), y)
    For i = 0 To UBound(cols)
        p = Split(cols(i), "|")
        opts = ""
        If UBound(p) >= 3 Then opts = p(3)
        cw = IIf(i = UBound(cols), px(200), px(Val(p(2))) - px(8))
        Set c = Lbl(prefix & i, UCase$(p(1)), cx, cy, cw, px(16), 8, True, CLR_MUTED_FG, acDetail, parentCtl)
        If InStr(opts, "r") > 0 Then c.TextAlign = 3
        If InStr(opts, "c") > 0 Or InStr(opts, "p") > 0 Then c.TextAlign = 2
        c.Tag = x & "|" & (cx - x) & "|" & cw
        cx = cx + px(Val(p(2)))
    Next
End Sub

' Tab control; each page is "name|caption".
Public Function TabCtl(ByVal nm As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ParamArray pages() As Variant) As Control
    Dim c As Control, i As Long
    Set c = Mk(acTabCtl, nm, x, y, w, h, acDetail, "")
    Do While c.Pages.Count < UBound(pages) + 1
        c.Pages.Add
    Loop
    For i = 0 To UBound(pages)
        c.Pages(i).Name = Split(pages(i), "|")(0)
        c.Pages(i).Caption = Split(pages(i), "|")(1)
    Next
    c.FontName = FONT_UI
    c.FontSize = 10
    c.BackStyle = 0
    c.UseTheme = True
    On Error Resume Next
    c.BackColor = CLR_SURFACE
    c.BorderColor = CLR_BORDER
    c.ForeColor = CLR_MUTED_FG
    c.HoverColor = CLR_OUT_HOVER
    c.HoverForeColor = CLR_FG
    c.PressedColor = CLR_PRIMARY_SOFT
    c.PressedForeColor = CLR_FG
    On Error GoTo 0
    Set TabCtl = c
End Function

' Clickable text "link" or tab label.
Public Function Link(ByVal nm As String, ByVal cap As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
        ByVal onClickExpr As String, Optional ByVal size As Integer = 10, Optional ByVal fore As Long = CLR_PRIMARY, _
        Optional ByVal parentCtl As String = "") As Control
    Dim c As Control
    Set c = Lbl(nm, cap, x, y, w, h, size, False, fore, acDetail, parentCtl)
    c.OnClick = onClickExpr
    Set Link = c
End Function

' ---------------------------------------------------------------------------
' Behaviour
' ---------------------------------------------------------------------------

' Point event properties (e.g. "OnClick") at the form module's handlers.
Public Sub Ev(c As Object, ParamArray props() As Variant)
    Dim i As Long
    For i = 0 To UBound(props)
        c.Properties(props(i)) = "[Event Procedure]"
    Next
End Sub

' h/v: 0 = left/top, 1 = right/bottom, 2 = both
Public Sub Anchor(c As Control, ByVal h As Integer, ByVal v As Integer)
    c.HorizontalAnchor = h
    c.VerticalAnchor = v
End Sub

' Add an expression-based conditional format. back / fore = -1 keep the
' control's own color.
Public Function CF(c As Control, ByVal expr As String, ByVal back As Long, ByVal fore As Long, _
                   Optional ByVal bold As Boolean = False) As Object
    Dim fc As Object
    Set fc = c.FormatConditions.Add(acExpression, , expr)
    If back <> -1 Then
        ' A conditional back color only paints on an opaque control.
        If c.BackStyle = 0 Then c.BackStyle = 1: c.BackColor = CLR_SURFACE
        fc.BackColor = back
    ElseIf c.BackStyle = 1 Then
        fc.BackColor = c.BackColor
    End If
    ' An unset conditional fore color falls back to black, so carry the control's own.
    If fore <> -1 Then fc.ForeColor = fore Else fc.ForeColor = c.ForeColor
    fc.FontBold = bold
    Set CF = fc
End Function

' Band color conditions on a control whose value is "<score> <label>" or a label,
' drawn as a rounded pill. guard/dflt: see RoundCF.
Public Sub BandCF(c As Control, ByVal fieldExpr As String, Optional ByVal guard As String = "", Optional ByVal dflt As Long = -1)
    Dim b As Variant, first As Integer
    first = c.FormatConditions.Count
    For Each b In BandLabels()
        CF c, fieldExpr & "=""" & b & """", BandBack(CStr(b)), CLR_WHITE
    Next
    RoundCF c, first, -1, guard, dflt
End Sub

' Badge colors for a status-like field, pairs of (value, variant), drawn as a rounded pill.
Public Sub BadgeCF(c As Control, ByVal fieldExpr As String, ParamArray pairs() As Variant)
    Dim i As Long, col As Variant, first As Integer
    first = c.FormatConditions.Count
    For i = 0 To UBound(pairs) Step 2
        col = BadgeColors(CStr(pairs(i + 1)))
        CF c, fieldExpr & "=""" & pairs(i) & """", col(0), col(1)
    Next
    RoundCF c, first
End Sub

' ---------------------------------------------------------------------------
' Rounded pills. Access only draws square corners, so a pill is an image
' control (<name>_bg) showing a capsule PNG (modPng) behind a transparent
' text control. Call RoundCF last: CF() would make the control opaque again.
' ---------------------------------------------------------------------------

' Image behind c for conditional colors: conditions firstCond..lastCond pick the
' pill color; dflt (if set) when none match; guard (if set) must hold for any pill.
Public Sub RoundCF(c As Control, Optional ByVal firstCond As Integer = 0, Optional ByVal lastCond As Integer = -1, _
                   Optional ByVal guard As String = "", Optional ByVal dflt As Long = -1)
    Dim i As Integer, expr As String, im As Control
    If m_isReport Then Exit Sub
    If lastCond < 0 Then lastCond = c.FormatConditions.Count - 1
    For i = firstCond To lastCond
        With c.FormatConditions(i)
            expr = expr & "," & .Expression1 & "," & .BackColor
        End With
    Next
    If dflt <> -1 Then expr = expr & ",True," & dflt
    If Len(expr) = 0 Then Exit Sub
    If Len(guard) > 0 Then expr = ",Not (" & guard & "),Null" & expr
    Set im = PillImage(c)
    im.Properties("ControlSource") = "=PillPng(Switch(" & Mid$(expr, 2) & ")," & c.Width \ 15 & "," & c.Height \ 15 & ")"
    c.BackStyle = 0
    ' A focused text box paints a square over its pill: unclickable pills never
    ' take focus (disabled + locked keeps the normal look); click handlers move
    ' focus to txtFocusSink.
    If Len(c.OnClick) = 0 Then
        c.Enabled = False
        c.Locked = True
    End If
End Sub

' Image behind a label painted at run time (PaintBadge / PaintPill / band picker).
Public Sub PillLabel(c As Control)
    If m_isReport Then Exit Sub
    PillImage c
End Sub

' Image control <name>_bg behind c, the same size, for a rounded pill.
Private Function PillImage(c As Control) As Control
    Dim im As Control, parentCtl As String, x As Control
    On Error Resume Next
    If TypeOf c.Parent Is Page Then parentCtl = c.Parent.Name
    On Error GoTo 0
    Set im = Mk(acImage, c.Name & "_bg", c.Left, c.Top, c.Width, c.Height, c.Section, parentCtl)
    im.Properties("PictureType") = 1      ' linked file
    im.Properties("SizeMode") = 3         ' zoom (images match the control's aspect)
    im.Properties("BackStyle") = 0
    im.Properties("BorderStyle") = 0
    im.Properties("HorizontalAnchor") = c.HorizontalAnchor
    im.Properties("VerticalAnchor") = c.VerticalAnchor
    ' Created last, the image is on top: bring the text back in front of it.
    For Each x In F.Controls
        x.InSelection = False
    Next
    c.InSelection = True
    DoCmd.RunCommand acCmdBringToFront
    c.InSelection = False
    Set PillImage = im
End Function

' Blue left accent bar used beside section titles in the dark theme.
Public Function AccentBar(ByVal nm As String, ByVal x As Long, ByVal y As Long, ByVal h As Long, _
        Optional ByVal parentCtl As String = "") As Control
    Dim c As Control
    Set c = Box(nm, x, y, px(3), h, CLR_PRIMARY, -1, acDetail, parentCtl)
    DarkOnly c
    Set AccentBar = c
End Function

' AccentBar beside a title label whose top is at labelTop, in size-pt text:
' it spans the text itself (0.3 to 1.1 of the font's pixel height below the
' label's top, cap tops to baseline), not the label or a subtitle under it.
Public Function TitleBar(ByVal nm As String, ByVal x As Long, ByVal labelTop As Long, ByVal size As Integer, _
        Optional ByVal parentCtl As String = "") As Control
    Dim fpx As Long
    fpx = (size * 4 + 1) \ 3                  ' points to pixels at 96 dpi
    Set TitleBar = AccentBar(nm, x, labelTop + px((fpx * 3) \ 10), px((fpx * 8) \ 10), parentCtl)
End Function

' Hidden in light palettes; ThemeForm shows it in dark ones.
Public Sub DarkOnly(c As Control)
    c.Tag = "darkonly"
    c.Visible = False
End Sub

' Unbound image control (picture set at run time).
Public Function Img(ByVal nm As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As Control
    Dim c As Control
    Set c = Mk(acImage, nm, x, y, w, h, acDetail, "")
    c.Properties("SizeMode") = 3
    c.Properties("PictureType") = 1
    c.Properties("BorderStyle") = 0
    c.Properties("BackStyle") = 0
    Set Img = c
End Function

' Image control settings shared by the rounded images (pills, cards, progress).
Private Sub ImageProps(c As Control)
    c.Properties("PictureType") = 1      ' linked file
    c.Properties("SizeMode") = 3         ' zoom (images match the control's aspect)
    c.Properties("BackStyle") = 0
    c.Properties("BorderStyle") = 0
End Sub

' Card: a CLR_SURFACE panel with a 1px CLR_BORDER border and rounded corners,
' with a 10pt title and a dark-only accent bar beside it (title "" draws
' neither). Access only draws square corners, so each corner is a small image
' (box<nm>_c1 to _c4) showing CardPng; ThemeForm repaints it with the
' palette. The fill (box<nm>) has no border of its own: Access paints a
' rectangle's right and bottom border a pixel past its corner images, which
' showed as a thin square line along those edges. Instead the four straight
' edges are 1px rules (box<nm>_t, _b, _l, _r) that stop at the corner images.
' back is the color the card sits on. hAnchor / vAnchor anchor the card (see
' Anchor) and the corners and edges follow. square = True leaves the corners
' square (and keeps the box's own border), for a card where a subform or
' another control that repaints often would cover a corner (images under
' them flicker).
' Controls: box<nm>, box<nm>_c1..4, box<nm>_t/_b/_l/_r, bar<nm>, lbl<nm>Hdr.
Public Sub Card(ByVal nm As String, ByVal title As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
        Optional ByVal parentCtl As String = "", Optional ByVal hAnchor As Integer = 0, Optional ByVal vAnchor As Integer = 0, _
        Optional ByVal back As Long = CLR_BG, Optional ByVal square As Boolean = False)
    Dim c As Control, k As Long, r As Long, farX As Boolean, farY As Boolean
    If square Then
        Set c = Box("box" & nm, x, y, w, h, CLR_SURFACE, CLR_BORDER, acDetail, parentCtl): Anchor c, hAnchor, vAnchor
    Else
        r = px(CARD_RADIUS)
        ' The fill stops a pixel short on the right and bottom, inside the edges.
        Set c = Box("box" & nm, x, y, w - px(1), h - px(1), CLR_SURFACE, -1, acDetail, parentCtl): Anchor c, hAnchor, vAnchor
        Set c = Rule("box" & nm & "_t", x + r, y, w - 2 * r, 0, CLR_BORDER, acDetail, parentCtl)
        Anchor c, hAnchor, EdgeAnchor(vAnchor, False)
        Set c = Rule("box" & nm & "_b", x + r, y + h - px(1), w - 2 * r, 0, CLR_BORDER, acDetail, parentCtl)
        Anchor c, hAnchor, EdgeAnchor(vAnchor, True)
        Set c = Rule("box" & nm & "_l", x, y + r, 0, h - 2 * r, CLR_BORDER, acDetail, parentCtl)
        Anchor c, EdgeAnchor(hAnchor, False), vAnchor
        Set c = Rule("box" & nm & "_r", x + w - px(1), y + r, 0, h - 2 * r, CLR_BORDER, acDetail, parentCtl)
        Anchor c, EdgeAnchor(hAnchor, True), vAnchor
        For k = 1 To 4
            farX = (k = 2 Or k = 3)
            farY = (k >= 3)
            Set c = Mk(acImage, "box" & nm & "_c" & k, IIf(farX, x + w - r, x), IIf(farY, y + h - r, y), r, r, acDetail, parentCtl)
            ImageProps c
            c.Properties("ControlSource") = "=CardPng(" & CLR_SURFACE & "," & CLR_BORDER & "," & back & "," & CARD_RADIUS & "," & k & ")"
            Anchor c, EdgeAnchor(hAnchor, farX), EdgeAnchor(vAnchor, farY)
        Next
    End If
    If Len(title) > 0 Then
        Set c = TitleBar("bar" & nm, x + px(8), y + px(10), 10, parentCtl)
        Anchor c, IIf(hAnchor = 1, 1, 0), IIf(vAnchor = 1, 1, 0)
        Set c = Lbl("lbl" & nm & "Hdr", title, x + px(14), y + px(10), w - px(28), px(20), 10, True, CLR_FG, acDetail, parentCtl)
        Anchor c, hAnchor, IIf(vAnchor = 1, 1, 0)
    End If
End Sub

' Anchor for a part on one edge of something anchored a: on the far side
' (right or bottom) it follows a "both" anchor, on the near side it stays.
Private Function EdgeAnchor(ByVal a As Integer, ByVal farSide As Boolean) As Integer
    If a = 2 Then EdgeAnchor = IIf(farSide, 1, 0) Else EdgeAnchor = a
End Function

' Progress capsule: a CLR_BORDER track filled with CLR_PRIMARY to the
' percentage expr gives (0-100; Null shows an empty track). expr is a field
' or an expression, e.g. "[percent_done]" or "37". Bound, so it follows the
' record in a continuous form, and ThemeForm repaints it. Don't anchor it to
' stretch: the image keeps its shape.
Public Function Progress(ByVal nm As String, ByVal expr As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
        Optional ByVal parentCtl As String = "") As Control
    Dim c As Control
    If Left$(expr, 1) = "=" Then expr = Mid$(expr, 2)
    Set c = Mk(acImage, nm, x, y, w, h, acDetail, parentCtl)
    ImageProps c
    c.Properties("ControlSource") = "=ProgressPng(" & expr & "," & CLR_PRIMARY & "," & CLR_BORDER & "," & w \ 15 & "," & h \ 15 & ")"
    Set Progress = c
End Function

' Page title + optional subtitle at the standard position.
' tip: explanation shown by a "?" beside the title (TitleTip). lblSubtitle is
' filled at run time by pages that show context under the title.
Public Sub PageHeader(ByVal title As String, ByVal tip As String, Optional ByVal x As Long = 480)
    TitleBar "barTitle", x - px(14), px(22), 15
    Lbl "lblTitle", title, x, px(22), px(560), px(30), 15, True
    Lbl "lblSubtitle", "", x, px(52), px(700), px(20), 9, False, CLR_MUTED_FG
    If Len(tip) > 0 Then TitleTip "tipTitle", tip, Replace(title, "&&", "&"), x, px(22), px(30), 15, ""
End Sub

' "?" help mark placed just after a title's text (title drawn at x, y with
' height h and the given font size, bold). Hovering shows tip; clicking shows
' it in a message (ShowTip).
Public Function TitleTip(ByVal nm As String, ByVal tip As String, ByVal title As String, ByVal x As Long, ByVal y As Long, _
        ByVal h As Long, ByVal size As Integer, Optional ByVal parentCtl As String = "") As Control
    Dim c As Control
    Set c = Lbl(nm, ChrW(&HE9CE), x + TextW(title, size), y, px(20), h, IIf(size > 12, 12, 10), False, CLR_MUTED_FG, acDetail, parentCtl)
    c.FontName = "Segoe MDL2 Assets"
    c.TextAlign = 2
    c.ControlTipText = Left$(tip, 255)
    c.Tag = tip
    c.OnClick = "=ShowTip([Form],""" & nm & """)"
    Set TitleTip = c
End Function

' Width (twips) of a bold title as drawn by Lbl, plus a small gap.
Private Function TextW(ByVal s As String, ByVal size As Integer) As Long
    TextW = MeasureText(s, size, 600) + 30
End Function

' "Radio" context banner (bg-muted/50 border) at y with given width.
Public Sub ContextBanner(ByVal x As Long, ByVal y As Long, ByVal w As Long, Optional ByVal withPill As Boolean = True)
    Dim c As Control
    Set c = Box("boxCtx", x, y, w, px(46), CLR_INSET, CLR_BORDER)
    Anchor c, 2, 0
    Set c = Lbl("lblCtxIcon", ChrW(&HE81D), x + px(10), y + px(14), px(18), px(18), 10, False, CLR_PRIMARY)
    c.FontName = "Segoe MDL2 Assets"
    Set c = Lbl("lblCtxName", "", x + px(34), y + px(5), w - px(260), px(19), 10, True)
    Anchor c, 2, 0
    Set c = Lbl("lblCtxSub", "", x + px(34), y + px(24), w - px(260), px(17), 8, False, CLR_MUTED_FG)
    Anchor c, 2, 0
    If withPill Then
        Set c = Lbl("lblCtxPill", "", x + w - px(220), y + px(13), px(210), px(20), 9, True, CLR_WHITE)
        c.TextAlign = 2
        Anchor c, 1, 0
        PillLabel c
    End If
End Sub

' Tab strip of labels with underline boxes. tabs are "key|Caption", or
' "key|Caption|n" to show a count, "Caption (n)", which SetCount updates at
' run time (the tab is sized for a three-digit count).
Public Sub TabStrip(ByVal x As Long, ByVal y As Long, ByVal totalW As Long, ByVal handler As String, ParamArray tabs() As Variant)
    Dim i As Long, k As String, cap As String, w As Long, c As Control, x0 As Long, p() As String
    x0 = x
    For i = 0 To UBound(tabs)
        p = Split(tabs(i), "|")
        k = p(0)
        cap = p(1)
        If UBound(p) >= 2 Then
            w = px(Len(cap & " (000)") * 7.5 + 28)
            Set c = Link("tab_" & k, cap & " (" & p(2) & ")", x, y, w, px(26), "=" & handler & "(""" & k & """)", 10, CLR_MUTED_FG)
            c.Tag = cap
        Else
            w = px(Len(cap) * 7.5 + 28)
            Set c = Link("tab_" & k, cap, x, y, w, px(26), "=" & handler & "(""" & k & """)", 10, CLR_MUTED_FG)
        End If
        c.TextAlign = 2
        Box "ul_" & k, x, y + px(28), w, px(2), CLR_PRIMARY
        x = x + w + Sp(SP_ITEM_GAP)
    Next
    Set c = Rule("boxTabLine", x0, y + px(30), totalW, 0, CLR_BORDER)
    Anchor c, 2, 0
End Sub

' Footer controls live in the parent, below the table; pagination defaults off.
Public Sub GridFooter(ByVal subName As String, ByVal prefix As String, ByVal keyField As String, Optional ByVal enabled As Boolean = False)
    Dim c As Control, i As Long, suffix As String, delta As Long, slot As Long
    SetTablePagination F.Controls(subName), enabled
    Set c = Lbl(prefix & "Count", "", 0, 0, px(200), px(26), 9)
    c.Tag = "TableFooter:" & prefix: c.Visible = enabled
    For i = 0 To 4
        suffix = Choose(i + 1, "Previous", "Page0", "Page1", "Page2", "Next")
        delta = 0: slot = -1
        If i = 0 Then delta = -1
        If i = 4 Then delta = 1
        If i > 0 And i < 4 Then slot = i - 1
        Set c = Btn(prefix & suffix, Choose(i + 1, "<", "1", "2", "3", ">"), 0, 0, px(30), px(26), "outline")
        c.Tag = "TableFooter:" & prefix: c.Visible = enabled
        c.OnClick = "=MoveTablePage([Form],""" & subName & """,""" & prefix & """,""" & keyField & """," & delta & "," & slot & ")"
    Next
End Sub
' Responsive rounded card surface, with an invisible native image border.
Public Function RoundedCard(ByVal nm As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As Control
    Dim c As Control
    Set c = Img(nm, x, y, w, h)
    c.Tag = "RoundedCard": c.Picture = RoundedCardPng(w \ 15, h \ 15, CLR_SURFACE, CLR_BORDER)
    Set RoundedCard = c
End Function
Public Sub ResizeRoundedCard(ByVal c As Control, ByVal w As Long, ByVal h As Long)
    c.Width = w: c.Height = h: c.BorderStyle = 0
    c.Picture = RoundedCardPng(w \ 15, h \ 15, Th(CLR_SURFACE), Th(CLR_BORDER))
End Sub
