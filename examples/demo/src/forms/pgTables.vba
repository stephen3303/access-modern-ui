Option Compare Database
Option Explicit

' Tables (see Build_pgTables). The status chips, search box and category
' filter apply to whichever table the tab strip shows; the chips and tabs
' show how many items each would list. Columns (on the first tab) opens the
' column chooser for its table; the choice is put back when the page opens.
' The first table's captions also filter it (the funnels, modGridFilter): its
' chips, search and category go through SetGridFilter so the header filters
' stay, and its counts include them.

Private Const ITEMS_FROM As String = "SELECT id, name, owner, category, status, Format(due_date,'dd mmm yyyy') AS due, due_date " & _
    "FROM demo_item"
Private Const ITEMS_SQL As String = ITEMS_FROM & " ORDER BY due_date"

Private m_tab As String
Private m_chip As Long
Private m_search As String

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.sfOptions.Form.RecordSource = ITEMS_SQL
    Me.sfPlain.Form.RecordSource = ITEMS_SQL
    Me.cboCategory = "(All categories)"
    ShowTab "options"
    Exit Sub
Err_Handler:
    HandleError "pgTables.Form_Load"
End Sub

Private Sub Form_Resize()
    On Error GoTo Err_Handler
    GridFit Me.sfOptions, "hdrOpt"
    GridScrollAll Me
    Exit Sub
Err_Handler:
    HandleError "pgTables.Form_Resize"
End Sub

' Hide, restore or reorder the first table's columns (modGridCols).
Private Sub btnColumns_Click()
    On Error GoTo Err_Handler
    ChooseGridColumns Me.sfOptions, "hdrOpt"
    Exit Sub
Err_Handler:
    HandleError "pgTables.btnColumns_Click"
End Sub

' Tab strip click.
Public Function TabClick(ByVal key As String)
    On Error GoTo Err_Handler
    ShowTab key
    Exit Function
Err_Handler:
    HandleError "pgTables.TabClick"
End Function

' Status chip click: 0 is All.
Public Function Chip(ByVal i As Long)
    On Error GoTo Err_Handler
    Me.txtFocusSink.SetFocus
    m_chip = i
    RefreshPage
    Exit Function
Err_Handler:
    HandleError "pgTables.Chip"
End Function

Private Sub txtSearch_Change()
    On Error GoTo Err_Handler
    m_search = Nz(Me.txtSearch.Text, "")
    RefreshPage
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(m_search)
    Exit Sub
Err_Handler:
    HandleError "pgTables.txtSearch_Change"
End Sub

Private Sub cboCategory_AfterUpdate()
    On Error GoTo Err_Handler
    RefreshPage
    Exit Sub
Err_Handler:
    HandleError "pgTables.cboCategory_AfterUpdate"
End Sub

' Row click in either table: nothing to open on this page.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
End Sub

' A header filter on the first table changed (GridHeaders onFilter): update
' the counts and the empty state.
Public Sub HeaderFiltered(ByVal grid As String)
    RefreshPage
End Sub

' Show one table and its captions, hide the other.
Private Sub ShowTab(ByVal key As String)
    Dim c As Control, isOpt As Boolean
    m_tab = key
    isOpt = (key = "options")
    For Each c In Me.Controls
        If c.Name Like "hdrOpt*" Then SetVisible c, isOpt
        If c.Name Like "hdrPlain*" Then SetVisible c, Not isOpt
    Next
    SetVisible Me.sfOptions, isOpt
    SetVisible Me.sfPlain, Not isOpt
    SetVisible Me.btnColumns, isOpt
    ' After the captions: hides the ones the user hid, and puts back their layout.
    If isOpt Then RestoreGridColumns Me.sfOptions, "hdrOpt"
    Me.ul_options.Visible = isOpt
    Me.ul_plain.Visible = Not isOpt
    Me.tab_options.ForeColor = IIf(isOpt, Th(CLR_FG), Th(CLR_MUTED_FG))
    Me.tab_plain.ForeColor = IIf(isOpt, Th(CLR_MUTED_FG), Th(CLR_FG))
    RefreshPage
End Sub

' Apply the chips, search and category to the table shown, mark the search
' in both tables' rows, and update the counts. The first table's header
' filters count on its tab, and in the chips while it shows.
Private Sub RefreshPage()
    Dim statuses As Variant, i As Long, where As String, cat As String, hdr As String, nOpt As Long, nPlain As Long
    statuses = Array("", "In progress", "Blocked", "Not started", "Complete")
    cat = Nz(Me.cboCategory, "")
    If Left$(cat, 1) = "(" Then cat = ""
    where = ItemFilter(CStr(statuses(m_chip)), m_search, cat)
    hdr = GridFilterWhere(Me.sfOptions)
    If m_tab = "options" Then
        SetGridFilter Me.sfOptions, where
        Me.sfOptions.Form.txtSel = Null
    Else
        FilterGrid Me.sfPlain, where
    End If
    GridHighlight Me.sfOptions, m_search
    GridHighlight Me.sfPlain, m_search
    For i = 0 To UBound(statuses)
        StyleChip Me("chip" & i), i = m_chip, _
            CountItems(Both(ItemFilter(CStr(statuses(i)), m_search, cat), IIf(m_tab = "options", hdr, "")))
    Next
    nOpt = CountItems(Both(where, hdr))
    nPlain = CountItems(where)
    SetCount Me.tab_options, nOpt
    SetCount Me.tab_plain, nPlain
    SetVisible Me.lblEmpty, IIf(m_tab = "options", nOpt, nPlain) = 0
    ' The empty state shows in place of the table: a subform always covers a
    ' label, so the table shown is hidden while nothing matches.
    SetVisible Me(IIf(m_tab = "options", "sfOptions", "sfPlain")), Not Me.lblEmpty.Visible
End Sub

' Items matching where, counted over the tables' record source (the header
' filters name its fields, such as the formatted due).
Private Function CountItems(ByVal where As String) As Long
    Dim db As DAO.Database, rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT Count(*) FROM (" & ITEMS_FROM & ") AS q" & _
        IIf(Len(where) > 0, " WHERE " & where, ""), dbOpenSnapshot)
    CountItems = rs(0)
    rs.Close
End Function

' a AND b, either of which may be "".
Private Function Both(ByVal a As String, ByVal b As String) As String
    If Len(a) > 0 And Len(b) > 0 Then
        Both = "(" & a & ") AND (" & b & ")"
    Else
        Both = a & b
    End If
End Function
