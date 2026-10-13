Option Compare Database
Option Explicit

' Kit Demo - BuildAll and the Build_<form> procedures. Each page shows a
' group of kit features: pgCards (Card, Progress), pgTables (BuildGrid
' options, GridFit, tabs and chips with counts) and pgList (the list/detail
' pattern, PrintGrid and ExportGrid). frmMain is the shell.

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
        "category TEXT(30), status TEXT(30), percent_done LONG, due_date DATETIME, notes LONGTEXT, flagged YESNO)", dbFailOnError
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
        CurrentDb.Execute "INSERT INTO demo_item (name, owner, category, status, percent_done, due_date, notes, flagged) VALUES (" & _
            Q(names(i)) & "," & Q(owners(i Mod 6)) & "," & Q(cats((i \ 2) Mod 6)) & "," & Q(st) & "," & pct & "," & _
            "#" & Format$(DateSerial(2026, 10, 1) + i * 3, "mm/dd/yyyy") & "#," & _
            Q("Demonstration item " & (i + 1) & ". Fictional data for testing the kit.") & "," & _
            IIf(i Mod 4 = 0, "True", "False") & ")", dbFailOnError
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
    Build_dlgAppearance
    BuildMessageDialog             ' the kit's dlgMessage, for KitMsgBox
    Build_frmMain
End Sub

' Styled tables (BuildGrid): "field|caption|width px|options".
Private Function CardCols() As Variant
    CardCols = Array("name|Item|320|b", "owner|Owner|170|", "category|Category|140|m", "status|Status|120|p", "due|Due|100|")
End Function

Private Function OptionCols() As Variant
    OptionCols = Array("name|Item|300|b", "owner|Owner|170|", "category|Category|140|m", "status|Status|120|p", "due|Due|110|r")
End Function

Private Function PlainCols() As Variant
    PlainCols = Array("name|Item|340|b", "owner|Owner|180|", "status|Status|140|", "due|Due|120|r")
End Function

Private Function ListCols() As Variant
    ListCols = Array("name|Item|230|b", "owner|Owner|140|", "status|Status|110|p", "due|Due|80|")
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

' Search box with a placeholder; its Change event filters as you type.
Private Sub SearchBox(ByVal x As Long, ByVal y As Long, ByVal w As Long)
    Dim c As Control
    Set c = Txt("txtSearch", "", x, y, w, px(28), 9)
    c.Format = "@;""Search items or owners..."""
    Ev c, "OnChange"
End Sub

' Detail panel form for pgList: one item, editable. The Flagged switch
' (Toggle) saves as soon as it changes.
Private Sub Build_subItem()
    Dim c As Control, x2 As Long, hw As Long
    NewForm "demo_item", 0, px(328), px(362)
    F.Section(acDetail).BackColor = CLR_SURFACE
    F.Section(acDetail).AlternateBackColor = CLR_SURFACE
    hw = px(156): x2 = px(172)
    FieldLbl "lblName", "ITEM", 0, 0, px(328)
    Txt "txtName", "name", 0, px(18), px(328), px(30)
    FieldLbl "lblOwner", "OWNER", 0, px(56), px(328)
    Txt "txtOwner", "owner", 0, px(74), px(328), px(30)
    FieldLbl "lblStatus", "STATUS", 0, px(112), hw
    Cbo "cboStatus", "status", "Not started;In progress;Blocked;Complete;Cancelled", 0, px(130), hw, px(30)
    FieldLbl "lblPct", "PERCENT DONE", x2, px(112), hw
    Set c = Txt("txtPct", "percent_done", x2, px(130), hw, px(30))
    Ev c, "AfterUpdate"
    Progress "imgPct", "[percent_done]", 0, px(170), px(328), px(10)
    FieldLbl "lblDue", "DUE", 0, px(190), hw
    Set c = Txt("txtDue", "due_date", 0, px(208), hw, px(30))
    c.Format = "dd mmm yyyy"
    FieldLbl "lblFlag", "FLAGGED FOR REVIEW", x2, px(190), hw
    Toggle "tglFlag", "flagged", x2, px(211), onChange:="FlagChanged"   ' 24px track centred on the 30px row
    FieldLbl "lblNotes", "NOTES", 0, px(246), px(328)
    Set c = Txt("txtNotes", "notes", 0, px(264), px(328), px(60))
    c.EnterKeyBehavior = True
    Btn "btnSave", "Save", px(232), px(334), px(96), px(28)
    SaveForm "subItem"
End Sub

' Cards and progress: rounded cards, the square fallback, progress capsules
' at 0, 37 and 100 percent, and a table inside a rounded card.
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
    Lbl "lblSquareBody", "Square corners: the fallback for a card where a subform or a control that repaints often covers a corner.", _
        x + px(14), y + px(38), w - px(28), px(54), 9, False, CLR_MUTED_FG

    y = px(250): h = PAGE_H - y - px(16)
    Card "Items", "Items in a rounded card", CX, y, CW(), h, , 2, 2
    GridHeaders "hdrItems", CardCols(), CX + px(16), y + px(40)
    Set c = SubF("sfItems", "subGridCards", CX + px(16), y + px(62), CW() - px(32), h - px(78)): Anchor c, 2, 2
    SaveForm "pgCards"
End Sub

' Tables: one table with every BuildGrid option (header band, zebra rows,
' status pills, 38px rows, proportional fit) and one built the default way,
' switched by a tab strip with counts; status chips with counts and a search
' box and category filter to their right.
Private Sub Build_pgTables()
    Dim c As Control, x As Long, y As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad", "OnResize"
    PageHeader "Tables", "The first tab shows every table option; the second shows a table built the default way. " & _
        "Resize the window: the first table's columns grow in proportion.", CX
    TabStrip CX, px(84), CW(), "TabClick", "options|Table options|0", "plain|Default table|0"
    y = px(128)
    x = Chips(CX, y + px(1))
    SearchBox x + px(10), y, px(220)
    Set c = Cbo("cboCategory", "", "(All categories);Engineering;Weapons;Supply;Navigation;Medical", x + px(240), y, px(170), px(28))
    Ev c, "AfterUpdate"
    y = px(170)
    GridHeaders "hdrOpt", OptionCols(), CX, y, "", True, CW()
    Set c = SubF("sfOptions", "subGridOptions", CX, y + px(22), CW(), PAGE_H - y - px(38)): Anchor c, 2, 2
    GridHeaders "hdrPlain", PlainCols(), CX, y
    Set c = SubF("sfPlain", "subGridPlain", CX, y + px(22), CW(), PAGE_H - y - px(38)): Anchor c, 2, 2
    c.Visible = False
    Set c = Lbl("lblEmpty", "No items match.", CX, y + px(60), CW(), px(20), 10, False, CLR_MUTED_FG)
    c.TextAlign = 2: c.Visible = False: Anchor c, 2, 0
    SaveForm "pgTables"
End Sub

' List and detail: the list card fills the width, the 360px detail card is
' anchored right. Print and export the list as filtered.
Private Sub Build_pgList()
    Dim c As Control, x As Long, y As Long, h As Long, lw As Long
    lw = ListW()
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "List and detail", "Click a row to show it on the right. Print and Export use the rows the list shows.", CX
    Set c = Btn("btnPrint", "Print", PAGE_W - CX - px(96), px(24), px(96), px(32), "outline"): Anchor c, 1, 0
    Set c = Btn("btnXlsx", "Export Excel", PAGE_W - CX - px(222), px(24), px(118), px(32), "outline"): Anchor c, 1, 0
    Set c = Btn("btnCsv", "Export CSV", PAGE_W - CX - px(338), px(24), px(108), px(32), "outline"): Anchor c, 1, 0
    y = px(84)
    x = Chips(CX, y + px(1))
    SearchBox x + px(10), y, px(220)
    y = px(124): h = PAGE_H - y - px(16)
    Card "List", "Items", CX, y, lw, h, , 2, 2
    GridHeaders "hdrList", ListCols(), CX + px(16), y + px(40)
    Set c = SubF("sfList", "subGridList", CX + px(16), y + px(62), lw - px(32), h - px(78)): Anchor c, 2, 2
    Set c = Lbl("lblEmpty", "No items match.", CX + px(16), y + px(100), lw - px(32), px(20), 10, False, CLR_MUTED_FG)
    c.TextAlign = 2: c.Visible = False: Anchor c, 2, 0
    x = CX + lw + px(12)
    Card "Detail", "Details", x, y, DETAIL_W, h, , 1, 2
    Set c = SubF("sfDetail", "", x + px(16), y + px(40), DETAIL_W - px(32), h - px(56)): Anchor c, 1, 2
    c.Visible = False
    Set c = Lbl("lblDetailEmpty", "Select an item to see its details.", x + px(16), y + px(120), DETAIL_W - px(32), px(20), _
        10, False, CLR_MUTED_FG)
    c.TextAlign = 2: Anchor c, 1, 0
    SaveForm "pgList"
End Sub

' Shell: sidebar with the pages and Appearance, and the page host sfPage.
Private Sub Build_frmMain()
    Dim c As Control, i As Long, pages As Variant, y As Long
    NewForm "", 0, px(200) + PAGE_W, PAGE_H
    Ev F, "OnLoad", "OnTimer", "OnResize"
    F.Caption = APP_NAME
    Set c = Box("boxSide", 0, 0, px(200), PAGE_H, CLR_SIDEBAR): Anchor c, 0, 2
    Lbl "lblBrand", APP_NAME, px(20), px(20), px(160), px(26), 13, True, CLR_WHITE
    Lbl "lblBrandSub", "access-modern-ui", px(20), px(46), px(160), px(18), 8, False, CLR_NAV_MUTED
    Box "boxSideLine", px(16), px(76), px(168), px(1), CLR_SIDEBAR_LINE
    pages = Array("cards|Cards and progress", "tables|Tables", "list|List and detail")
    y = px(90)
    For i = 0 To UBound(pages)
        Set c = Link("nt_" & Split(pages(i), "|")(0), Split(pages(i), "|")(1), px(12), y, px(176), px(34), _
            "=NavTo(""" & Split(pages(i), "|")(0) & """)", 10, CLR_NAV_FG)
        c.BackColor = CLR_NAV_ACTIVE_BG
        c.LeftMargin = px(12)
        VCenter c, px(34), 10, True
        y = y + px(38)
    Next
    Set c = Link("lnkAppearance", "Appearance", px(12), PAGE_H - px(52), px(176), px(34), "=ShowAppearance()", 10, CLR_NAV_FG)
    c.LeftMargin = px(12)
    VCenter c, px(34), 10, True
    Anchor c, 0, 1
    Set c = SubF("sfPage", "pgCards", px(200), 0, PAGE_W, PAGE_H): Anchor c, 2, 2
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

' Open on frmMain without the navigation pane or document tabs, and record
' when this front end was built (error-log entries show it).
Private Sub SetStartupOptions()
    SetProp "AppTitle", dbText, APP_NAME
    SetProp "StartUpForm", dbText, "frmMain"
    SetProp "StartUpShowDBWindow", dbBoolean, False
    SetProp "ShowDocumentTabs", dbBoolean, False
    SetProp "Perform Name AutoCorrect", dbLong, 0
    SetProp "Track Name AutoCorrect Info", dbLong, 0
    SetProp "AppBuild", dbText, Format$(Now(), "yyyy-mm-dd hh:nn")
End Sub
