Option Compare Database
Option Explicit

' Kit Demo - BuildAll and the Build_<form> procedures. Each page shows a
' group of kit features: pgCards (Card, Progress), pgTables (BuildGrid
' options, GridFit, the column chooser, tabs and chips with counts),
' pgList (the list/detail pattern, PrintGrid and ExportGrid, Alert, Breadcrumb),
' pgSteps (Stepper, Accordion) and pgDesigner (the kit's form designer,
' BuildDesigner). frmMain is the shell; Help in its
' sidebar and F1 open the kit's help window on the page's topic
' (src\help\help.txt).

Public Const PAGE_W As Long = 15840             ' 1056 px
Public Const PAGE_H As Long = 8400              ' 560 px
Private Const CX As Long = 480                  ' page padding, 32 px
Private Const DETAIL_W As Long = 5400           ' detail card, 360 px

Private m_step As String

' Content width inside the page padding.
Private Function CW() As Long
    CW = PAGE_W - 2 * CX
End Function

' Build entry point (build.ps1). Returns "OK" or the step that failed.
Public Function BuildAll(ByVal srcRoot As String, ByVal outDir As String, ByVal uiOnly As Boolean) As String
    On Error GoTo Fail
    g_Silent = True
    UseBuildPalette
    g_SrcRoot = srcRoot
    If Not uiOnly Then
        m_step = "create tables": CreateTables
        m_step = "seed": SeedItems
    End If
    m_step = "user interface": BuildUserInterface
    m_step = "help": LoadHelpTopics srcRoot & "\help\help.txt"
    m_step = "startup": SetStartupOptions
    g_Silent = False
    BuildAll = "OK"
    Exit Function
Fail:
    BuildAll = "ERROR during " & m_step & ": " & Err.Number & " " & Err.Description & " [" & BuildTrace() & "]"
    g_Silent = False
End Function

' ---------------------------------------------------------------------------
' Data
' ---------------------------------------------------------------------------

Private Sub CreateTables()
    Dim db As DAO.Database
    Set db = CurrentDb
    db.Execute "CREATE TABLE demo_item (id COUNTER PRIMARY KEY, name TEXT(120), owner TEXT(60), " & _
        "category TEXT(30), status TEXT(30), percent_done LONG, due_date DATETIME, notes LONGTEXT, flagged YESNO, access_code TEXT(20))", dbFailOnError
    db.Execute "CREATE TABLE error_log (id COUNTER PRIMARY KEY, logged_at DATETIME, error_number LONG, " & _
        "description LONGTEXT, err_source TEXT(255), source TEXT(255), call_trail LONGTEXT, context LONGTEXT, " & _
        "page TEXT(32), active_form TEXT(64), active_control TEXT(64), username TEXT(150), windows_user TEXT(64), " & _
        "computer TEXT(64), fe_build TEXT(32), access_version TEXT(32))", dbFailOnError
End Sub

' 24 fictional items with every status and a spread of progress; every
' fourth is flagged for review.
Private Sub SeedItems()
    Dim names As Variant, owners As Variant, cats As Variant, sts As Variant, i As Long, pct As Long, st As String
    names = Array("Fire main pump overhaul", "Sonar dome inspection", "Galley exhaust fan repair", _
        "Navigation chart corrections", "Medical locker inventory", "Torpedo tube 2 maintenance", _
        "Battery well cleaning", "Periscope optics check", "Hydraulic plant sampling", _
        "Spare parts audit", "Damage control locker restock", "Radio room crypto inventory", _
        "Main ballast tank survey", "Oxygen generator service", "Fresh water still repair", _
        "Ship control panel lamp test", "Training records review", "Emergency diesel run", _
        "Weapons handling drill", "Hull valve lineup", "Atmosphere monitor calibration", _
        "Chart table lighting fix", "Supply requisition backlog", "Annual hearing tests")
    owners = Array("ETN1 Rivera", "STS2 Okafor", "CS1 Brandt", "QM1 Holloway", "HM1 Castillo", "TM2 Pruitt")
    cats = Array("Engineering", "Weapons", "Supply", "Navigation", "Medical", "Weapons")
    sts = Array("Complete", "In progress", "Blocked", "Not started", "In progress", "Cancelled")
    For i = 0 To UBound(names)
        st = sts(i Mod 6)
        Select Case st
            Case "Complete": pct = 100
            Case "Not started": pct = 0
            Case "Cancelled": pct = 10
            Case Else: pct = (i * 37) Mod 95 + 5
        End Select
        CurrentDb.Execute "INSERT INTO demo_item (name, owner, category, status, percent_done, due_date, notes, flagged, access_code) VALUES (" & _
            Q(names(i)) & "," & Q(owners(i Mod 6)) & "," & Q(cats((i \ 2) Mod 6)) & "," & Q(st) & "," & pct & "," & _
            "#" & Format$(DateSerial(2026, 10, 1) + i * 3, "mm/dd/yyyy") & "#," & _
            Q("Demonstration item " & (i + 1) & ". Fictional data for testing the kit.") & "," & _
            IIf(i Mod 4 = 0, "True", "False") & "," & _
            Q("KD-" & Format$(4100 + i * 37, "0000")) & ")", dbFailOnError
    Next
End Sub

' ---------------------------------------------------------------------------
' User interface
' ---------------------------------------------------------------------------

' Subforms before the pages that host them, the shell last.
Private Sub BuildUserInterface()
    BuildGrid "subGridCards", CardCols(), CW() - px(32)
    BuildGrid "subGridOptions", OptionCols(), CW(), 1, CLR_BG, True, 38, True
    BuildGrid "subGridPlain", PlainCols(), CW(), 1, CLR_BG
    BuildGrid "subGridList", ListCols(), ListW() - px(32)
    Build_subItem
    Build_pgCards
    Build_pgTables
    Build_pgList
    Build_subAbout
    Build_subPrefs
    Build_pgSteps
    Build_pgDash
    BuildDesigner "pgDesigner", CX, PAGE_W, PAGE_H   ' the kit's form designer page
    Build_dlgAppearance
    BuildMessageDialog             ' the kit's dlgMessage, for KitMsgBox
    BuildInputDialog               ' the kit's dlgInput, for KitInputBox
    BuildDatePicker                ' the kit's dlgDatePick, for DateBox
    BuildToast                     ' the kit's dlgToast, for KitToast
    BuildBusy                      ' the kit's dlgBusy, for KitBusy (relinks, lock waits, exports)
    BuildColumnsDialog             ' the kit's dlgColumns, for ChooseGridColumns
    BuildFilterDialog              ' the kit's dlgFilter, for the header funnels (GridFilter)
    BuildMenuDialog                ' the kit's dlgMenu, for KitMenu (the detail card's More button)
    BuildHelpDialog                ' the kit's dlgHelp, for KitHelp (Help in the sidebar and F1)
    BuildPopoverDialog             ' the kit's dlgPopover, for KitPopover (the Preferences section's help)
    BuildPickDialog                ' the kit's dlgPick, for PickBox (searchable pickers)
    BuildPaletteDialog             ' the kit's dlgPalette, for Ctrl+K (KitPalette)
    BuildSplash "frmMain", "access-modern-ui demo"   ' the kit's frmSplash, the startup form
    Build_frmMain
End Sub

' Styled tables (BuildGrid): "field|caption|width px|options[|sort field]".
' Due shows formatted text, so its captions sort by due_date (GridHeaders).
Private Function CardCols() As Variant
    CardCols = Array("name|Item|320|b", "owner|Owner|170|", "category|Category|140|m", "status|Status|120|p", "due|Due|100|")
End Function

Private Function OptionCols() As Variant
    OptionCols = Array("name|Item|300|bs", "owner|Owner|170|s", "category|Category|140|m", "status|Status|120|p", "due|Due|110|r|due_date")
End Function

Private Function PlainCols() As Variant
    PlainCols = Array("name|Item|340|bs", "owner|Owner|180|s", "status|Status|140|", "due|Due|120|r|due_date")
End Function

Private Function ListCols() As Variant
    ListCols = Array("name|Item|230|bs", "owner|Owner|140|s", "status|Status|110|p", "due|Due|80||due_date")
End Function

' Width of the list card on pgList.
Private Function ListW() As Long
    ListW = CW() - DETAIL_W - px(12)
End Function

' 8pt bold muted uppercase caption above a field.
Private Sub FieldLbl(ByVal nm As String, ByVal cap As String, ByVal x As Long, ByVal y As Long, ByVal w As Long)
    Lbl nm, cap, x, y, w, px(16), 8, True, CLR_MUTED_FG
End Sub

' Status filter chips chip0..chip4 at (x, y), each with a count; a click calls
' =Chip(i). Returns the x just after the last chip.
Private Function Chips(ByVal x As Long, ByVal y As Long) As Long
    Dim caps As Variant, i As Long, w As Long, c As Control
    caps = Array("All", "In progress", "Blocked", "Not started", "Complete")
    For i = 0 To UBound(caps)
        w = px(Len(caps(i) & " (00)") * 7 + 20)
        Set c = Link("chip" & i, caps(i) & " (0)", x, y, w, px(26), "=Chip(" & i & ")", 9, CLR_MUTED_FG)
        c.TextAlign = 2
        c.Tag = caps(i)
        x = x + w + px(6)
    Next
    Chips = x
End Function

' The kit's search box (magnifier, placeholder, clear mark); the page's
' SearchFor filters as you type.
Private Sub PageSearch(ByVal x As Long, ByVal y As Long, ByVal w As Long)
    SearchBox "txtSearch", x, y, w, "SearchFor", "Search items or owners...", px(28)
End Sub

' Detail panel form for pgList: one item, editable. The Flagged switch
' (Toggle) saves the record as soon as it changes.
Private Sub Build_subItem()
    Dim c As Control, x2 As Long, hw As Long
    NewForm "demo_item", 0, px(328), px(362)
    Ev F, "OnCurrent", "AfterUpdate"
    F.Section(acDetail).BackColor = CLR_SURFACE
    F.Section(acDetail).AlternateBackColor = CLR_SURFACE
    Ev F, "BeforeUpdate"
    hw = px(156): x2 = px(172)
    FieldLbl "lblName", "ITEM", 0, 0, px(328)
    Txt "txtName", "name", 0, px(18), px(328), px(30)
    FieldLbl "lblOwner", "OWNER", 0, px(56), px(328)
    Txt "txtOwner", "owner", 0, px(74), px(288), px(30)
    Avatar "avOwner", "owner", px(297), px(73), px(31)   ' the owner's initials, beside the field
    FieldLbl "lblStatus", "STATUS", 0, px(112), hw
    Cbo "cboStatus", "status", "Not started;In progress;Blocked;Complete;Cancelled", 0, px(130), hw, px(30)
    FieldLbl "lblPct", "PERCENT DONE", x2, px(112), hw
    Set c = Txt("txtPct", "percent_done", x2, px(130), hw, px(30))
    Ev c, "AfterUpdate"
    Progress "imgPct", "[percent_done]", 0, px(170), px(328), px(10)
    FieldLbl "lblDue", "DUE", 0, px(190), hw
    DateBox "txtDue", "due_date", 0, px(208), px(30)
    FieldLbl "lblCode", "ACCESS CODE", x2, px(190), hw
    SecretBox "txtCode", "access_code", x2, px(208), hw, px(30)
    FieldLbl "lblNotes", "NOTES", 0, px(246), px(328)
    Set c = Txt("txtNotes", "notes", 0, px(264), px(328), px(60))
    c.EnterKeyBehavior = True
    ' Compact 36 x 20px switch centered on the 28px Save row, 4px in so its focus
    ' ring stays inside the form; its label toggles it too.
    Toggle "tglFlag", "flagged", px(4), px(338), px(36), px(20), "FlagChanged"
    Link "lblFlag", "Flagged for review", px(48), px(334), px(176), px(28), "=ToggleSwitch([Form],""tglFlag"")", 9, CLR_FG
    Btn "btnSave", "Save", px(232), px(334), px(96), px(28)
    SaveForm "subItem"
End Sub

' Cards and progress: rounded cards, the square fallback, progress capsules
' at 0, 37 and 100 percent, and a table inside a rounded card. The rounded
' card's Primary button shows a message (KitMsgBox) and Outline a toast
' (KitToast), a different variant each click. A segmented control (All, Open,
' Done) beside the table's title filters it.
Private Sub Build_pgCards()
    Dim c As Control, x As Long, y As Long, w As Long, h As Long, i As Long, pct As Variant
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Cards and progress", "Rounded cards (Card), the square fallback, and progress capsules (Progress). " & _
        "Move the mouse over the cards and buttons: nothing should flicker.", CX
    w = px(320): h = px(150): y = px(84)

    x = CX
    Card "Round", "Rounded card", x, y, w, h
    Lbl "lblRoundBody", "A surface with a 1px border and rounded corners. The buttons inside should hover without the card flickering.", _
        x + px(14), y + px(38), w - px(28), px(54), 9, False, CLR_MUTED_FG
    Btn "btnRoundA", "Primary", x + px(14), y + h - px(42), px(96), px(28)
    Btn "btnRoundB", "Outline", x + px(118), y + h - px(42), px(96), px(28), "outline"

    x = CX + w + px(16)
    Card "Prog", "Progress", x, y, w, h
    For i = 0 To 2
        pct = Array(0, 37, 100)(i)
        Lbl "lblPct" & i, pct & "%", x + px(14), y + px(44) + i * px(30), px(44), px(18), 9, False, CLR_MUTED_FG
        Progress "imgPct" & i, CStr(pct), x + px(62), y + px(48) + i * px(30), w - px(80), px(10)
    Next

    x = CX + 2 * (w + px(16))
    Card "Square", "Square card", x, y, w, h, square:=True
    Lbl "lblSquareBody", "Square corners: the fallback for a card where a subform or a control that repaints often covers a corner. " & _
        "Its buttons ask a question (KitInputBox).", x + px(14), y + px(38), w - px(28), px(54), 9, False, CLR_MUTED_FG
    Btn "btnSquareA", "Ask", x + px(14), y + h - px(42), px(96), px(28)
    Btn "btnSquareB", "Ask masked", x + px(118), y + h - px(42), px(96), px(28), "outline"

    y = px(250): h = PAGE_H - y - px(16)
    Card "Items", "Items in a rounded card", CX, y, CW(), h, , 2, 2
    Segmented "segShow", "", CX + px(220), y + px(6), px(210), "all|All;open|Open;done|Done", px(28), "ShowChanged"
    GridHeaders "hdrItems", CardCols(), CX + px(16), y + px(40), w:=CW() - px(32)
    Set c = SubF("sfItems", "subGridCards", CX + px(16), y + px(62), CW() - px(32), h - px(78)): Anchor c, 2, 2
    SaveForm "pgCards"
End Sub

' Tables: one table with every BuildGrid option (header band, zebra rows,
' status pills, 38px rows, proportional fit, filter funnels) and one built the default way,
' switched by a tab strip with counts; status chips with counts and a search
' box and category filter to their right. Columns, top right, opens the
' column chooser for the first table.
Private Sub Build_pgTables()
    Dim c As Control, x As Long, y As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad", "OnResize"
    PageHeader "Tables", "The first tab shows every table option; the second shows a table built the default way. " & _
        "Resize the window: the first table's columns grow in proportion. Columns hides and reorders them; " & _
        "a caption's funnel filters by its values.", CX
    Set c = Btn("btnColumns", "Columns", PAGE_W - CX - px(96), px(24), px(96), px(32), "outline"): Anchor c, 1, 0
    TabStrip CX, px(84), CW(), "TabClick", "options|Table options|0", "plain|Default table|0"
    y = px(128)
    x = Chips(CX, y + px(1))
    PageSearch x + px(10), y, px(220)
    Set c = Cbo("cboCategory", "", "(All categories);Engineering;Weapons;Supply;Navigation;Medical", x + px(240), y, px(170), px(28))
    Ev c, "AfterUpdate"
    y = px(170)
    GridHeaders "hdrOpt", OptionCols(), CX, y, "", True, CW(), "sfOptions", "sfOptions", "HeaderFiltered", True, "sfOptions"
    Set c = SubF("sfOptions", "subGridOptions", CX, y + px(22), CW(), PAGE_H - y - px(38)): Anchor c, 2, 2
    GridHeaders "hdrPlain", PlainCols(), CX, y, w:=CW(), sortSf:="sfPlain"
    Set c = SubF("sfPlain", "subGridPlain", CX, y + px(22), CW(), PAGE_H - y - px(38)): Anchor c, 2, 2
    c.Visible = False
    Set c = Lbl("lblEmpty", "No items match.", CX, y + px(60), CW(), px(20), 10, False, CLR_MUTED_FG)
    c.TextAlign = 2: c.Visible = False: Anchor c, 2, 0
    SaveForm "pgTables"
End Sub

' List and detail: the list card fills the width, the 360px detail card is
' anchored right. Print and export the list as filtered. A warning alert in
' the header row counts the items past their due date, and a breadcrumb
' below the title shows where the open item sits.
Private Sub Build_pgList()
    Dim c As Control, x As Long, y As Long, h As Long, lw As Long
    lw = ListW()
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "List and detail", "Click a row to show it on the right, or a column caption to sort. Print and Export use the rows the list shows.", CX
    Set c = Btn("btnPrint", "Print", PAGE_W - CX - px(96), px(24), px(96), px(32), "outline"): Anchor c, 1, 0
    Set c = Btn("btnXlsx", "Export Excel", PAGE_W - CX - px(222), px(24), px(118), px(32), "outline"): Anchor c, 1, 0
    Set c = Btn("btnCsv", "Export CSV", PAGE_W - CX - px(338), px(24), px(108), px(32), "outline"): Anchor c, 1, 0
    ' Compact (36px) alert in the header row, between the title and the buttons.
    Alert "alOverdue", CX + px(200), px(22), px(440), px(36), "warning", dismiss:=True
    ' Breadcrumb in the subtitle's place: All items > the chip's status > the open item.
    Breadcrumb "crumb", CX, px(54), px(620), 3, "CrumbGo", "c0|All items"
    y = px(84)
    x = Chips(CX, y + px(1))
    PageSearch x + px(10), y, px(220)
    y = px(124): h = PAGE_H - y - px(16)
    Card "List", "Items", CX, y, lw, h, , 2, 2
    GridHeaders "hdrList", ListCols(), CX + px(16), y + px(40), w:=ListW() - px(32), sortSf:="sfList"
    Set c = SubF("sfList", "subGridList", CX + px(16), y + px(62), lw - px(32), h - px(78)): Anchor c, 2, 2
    ' Shown in place of the list when nothing matches (EmptyState).
    EmptyState "emptyList", CX + px(16), y + px(90), lw - px(32), "E721", "No items match", _
        "Nothing matches the status and search you picked. Clear them to see every item.", "Clear filters"
    x = CX + lw + px(12)
    Card "Detail", "Details", x, y, DETAIL_W, h, , 1, 2
    ' More: the open item's actions in a dropdown menu (KitMenu).
    Set c = Btn("btnMore", ChrW(&HE712), x + DETAIL_W - px(44), y + px(6), px(32), px(28), "ghost")
    c.FontName = "Segoe MDL2 Assets": c.FontWeight = 400: c.ControlTipText = "More actions"
    Anchor c, 1, 0
    Set c = SubF("sfDetail", "", x + px(16), y + px(40), DETAIL_W - px(32), h - px(56)): Anchor c, 1, 2
    c.Visible = False
    Set c = Lbl("lblDetailEmpty", "Select an item to see its details.", x + px(16), y + px(120), DETAIL_W - px(32), px(20), _
        10, False, CLR_MUTED_FG)
    c.TextAlign = 2: Anchor c, 1, 0
    SaveForm "pgList"
End Sub

' A short text body for the accordion's closed-by-default sections.
Private Sub Build_subAbout()
    NewForm "", 0, CW() - px(32), px(56)
    F.Section(acDetail).BackColor = CLR_BG
    F.Section(acDetail).AlternateBackColor = CLR_BG
    Lbl "lblAbout", "A section body is a subform, so it can hold anything a form can. Opening a section pushes the " & _
        "ones below it down; closing it lets them back up.", px(4), 0, CW() - px(40), px(48), 9, False, CLR_MUTED_FG
    SaveForm "subAbout"
End Sub

' The accordion's Preferences section: a stepped Slider with its value, a
' star Rating, a help mark that opens a KitPopover, and a searchable picker
' (PickBox) of the items. Unbound: the demo keeps no settings.
Private Sub Build_subPrefs()
    Dim c As Control, x2 As Long, x3 As Long
    NewForm "", 0, CW() - px(32), px(112)
    F.KeyPreview = True                    ' arrow keys move the slider (Form_KeyDown)
    Ev F, "OnKeyDown"
    F.Section(acDetail).BackColor = CLR_BG
    F.Section(acDetail).AlternateBackColor = CLR_BG
    FieldLbl "lblRows", "ROWS PER PAGE", px(4), px(4), px(140)
    Set c = Link("lnkRowsHelp", ChrW(&HE9CE), px(112), px(2), px(20), px(20), "=RowsHelp()", 10, CLR_MUTED_FG)
    c.FontName = "Segoe MDL2 Assets": c.TextAlign = 2: c.ControlTipText = "What is this?"
    KeyButton "lnkRowsHelp_k", "lnkRowsHelp", c.Left, c.Top, c.Width, c.Height, "=RowsHelp()", "What is rows per page?"
    Slider "sldRows", "", px(4), px(30), px(360), 10, 50, 5, "RowsChanged", True
    x2 = px(420)
    FieldLbl "lblUseful", "HOW USEFUL IS THIS DEMO?", x2, px(4), px(220)
    Rating "rtUseful", "", x2, px(28), onChange:="UsefulChanged"
    Lbl "lblUseful2", "", x2, px(64), px(260), px(20), 9, False, CLR_MUTED_FG
    x3 = px(700)
    FieldLbl "lblWatch", "ITEM TO WATCH", x3, px(4), px(240)
    Set c = PickBox("txtWatch", "", "SELECT id, name, owner FROM demo_item", x3, px(28), px(240), , "Type an item or owner")
    Ev c, "OnKeyDown"
    SaveForm "subPrefs"
End Sub

' Steps and sections: a step indicator (Stepper) in a card with Back and
' Next, and an accordion of three sections below it, the first holding the
' items table.
Private Sub Build_pgSteps()
    Dim c As Control, y As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Steps and sections", "A step indicator (Stepper) and an accordion (Accordion). " & _
        "Back and Next move the steps; click a check mark to go back. Click a section's header to open or close it.", CX
    y = px(84)
    Card "Wizard", "New item", CX, y, CW(), px(160), , 2, 0
    Stepper "steps", CX + px(16), y + px(44), CW() - px(32), 1, "StepBack", "Details", "Owner", "Review", "Submit"
    Lbl "lblStepBody", "", CX + px(16), y + px(112), CW() - px(260), px(36), 9, False, CLR_MUTED_FG
    Set c = Btn("btnBack", "Back", CX + CW() - px(224), y + px(116), px(96), px(28), "outline")
    Set c = Btn("btnNext", "Next", CX + CW() - px(112), y + px(116), px(96), px(28))
    Accordion "acc", CX, y + px(176), CW(), True, "Items|subGridCards|150|open", "About sections|subAbout|56", _
        "Preferences|subPrefs|112"
    SaveForm "pgSteps"
End Sub

' Dashboard: four stat tiles (StatTile) over the items, and the flagged items
' in a card, with an empty state (EmptyState) when none are flagged. The
' tiles' changes are fixed demo numbers: the demo keeps no history.
Private Sub Build_pgDash()
    Dim c As Control, tw As Long, y As Long, h As Long, i As Long
    Const OPEN_ITEMS As String = "status Not In ('Complete','Cancelled')"
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Dashboard", "Stat tiles (StatTile) with change badges, and an empty state (EmptyState) where a list has nothing to show.", CX
    y = px(84)
    tw = (CW() - 3 * Sp(SP_CARD_GAP)) \ 4
    StatTile "statOpen", "Open items", "=DCount(""*"",""demo_item"",""" & OPEN_ITEMS & """)", CX, y, tw, , "=8", _
        "vs last week (demo)", "E8FD", "#,##0"
    StatTile "statLate", "Past due", "=DCount(""*"",""demo_item"",""due_date<Date() And " & OPEN_ITEMS & """)", _
        CX + (tw + Sp(SP_CARD_GAP)), y, tw, , "=-25", "vs last week (demo)", "E823", "#,##0", False
    StatTile "statDone", "Complete", "=DCount(""*"",""demo_item"",""status='Complete'"")", CX + 2 * (tw + Sp(SP_CARD_GAP)), y, tw, , _
        "=12", "vs last week (demo)", "E73E", "#,##0"
    StatTile "statPct", "Average done", "=DAvg(""percent_done"",""demo_item"")/100", CX + 3 * (tw + Sp(SP_CARD_GAP)), y, tw, , _
        "=0", "no change (demo)", "E8EF", "0%"
    y = y + px(112) + Sp(SP_CARD_GAP): h = PAGE_H - y - px(16)
    Card "Flagged", "Flagged for review", CX, y, CW(), h, , 2, 2
    GridHeaders "hdrFlagged", CardCols(), CX + px(16), y + px(40), w:=CW() - px(32)
    Set c = SubF("sfFlagged", "subGridCards", CX + px(16), y + px(62), CW() - px(32), h - px(78)): Anchor c, 2, 2
    EmptyState "emptyFlagged", CX + px(16), y + px(56), CW() - px(32), "E7C1", "Nothing is flagged", _
        "Flag an item for review on the List and detail page, and it shows here.", "Flag the first item"
    SaveForm "pgDash"
End Sub

' Shell: sidebar with the pages, Folder trust (the kit's trust page) and
' Appearance, and the page host sfPage.
Private Sub Build_frmMain()
    ShellBuild APP_NAME, "access-modern-ui", _
        Array("cards|Cards and progress|pgCards|E8A9", "tables|Tables|pgTables|F0E2", "list|List and detail|pgList|E8FD", _
            "steps|Steps and sections|pgSteps|E9D5", "dash|Dashboard|pgDash|E9D2", _
            "design|Form designer|pgDesigner|E70F"), _
        Array("help|Help|E897|=KitHelp()", "trust|Folder trust|E72E|=ShowTrustHelp()", _
            "appearance|Appearance|E790|=DemoAppearance()"), _
        PAGE_W, PAGE_H
    SaveForm "frmMain"
End Sub

' ---------------------------------------------------------------------------
' Startup
' ---------------------------------------------------------------------------

Private Sub SetProp(ByVal name As String, ByVal kind As Integer, ByVal value As Variant)
    Dim db As DAO.Database
    Set db = CurrentDb
    On Error Resume Next
    db.Properties(name) = value
    If Err.Number = 3270 Then
        Err.Clear
        db.Properties.Append db.CreateProperty(name, kind, value)
    End If
End Sub

' Open on the splash screen, which opens frmMain, without Access's own
' splash, the navigation pane or document tabs, and record when this front
' end was built (error-log entries and the splash show it). The kit's
' BuildTrustCheck opens the splash at startup (in place of StartUpForm) and
' shows its trust page when Access would ask to enable content.
Private Sub SetStartupOptions()
    SetProp "AppTitle", dbText, APP_NAME
    SetProp "StartUpShowDBWindow", dbBoolean, False
    SetProp "ShowDocumentTabs", dbBoolean, False
    SetProp "Perform Name AutoCorrect", dbLong, 0
    SetProp "Track Name AutoCorrect Info", dbLong, 0
    SetProp "AppBuild", dbText, Format$(Now(), "yyyy-mm-dd hh:nn")
    BuildShortcutKeys               ' Ctrl+S, Ctrl+N, Ctrl+F, Ctrl+K and F1 (the kit's help), the AutoKeys macro
    HideAccessSplash                ' it writes a file, so a failure leaves the settings above in place
    BuildTrustCheck "frmSplash"     ' after any StartUpForm setting: it removes StartUpForm
End Sub
