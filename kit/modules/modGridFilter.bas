Option Compare Database
Option Explicit

' Access UI kit - click-to-filter table headers. GridHeaders(..., filterSf)
' puts a funnel at the right end of each caption whose column shows a field.
' Clicking it opens dlgFilter (GridFilter), which lists the column's values;
' the user picks the ones to show. The funnel turns CLR_PRIMARY while its
' column is filtered. Filters work alongside click-to-sort and the column
' chooser.
'
' The table's Filter is the page's own filter (search box, chips) AND each
' header filter, so PrintGrid, ExportGrid and pagination (GridFooter) follow
' it. A page that filters such a table itself must set its filter with
' SetGridFilter, never the table's Filter: setting Filter directly replaces
' the header filters, and the next header click clears them to match.
'
' A header filter keeps the rows whose value is one of those chosen:
' "[status] IN ('Blocked','In progress')", with "(Blanks)" for Null (and ""
' in a text field). A value that wasn't listed when the filter was set (a row
' added since, or past the list's limit) is hidden. Filters last while the
' page is open; a page opened again starts unfiltered. Memo, OLE and
' attachment columns can't be filtered.

' Funnel glyph (Segoe MDL2 Assets "Filter").
Public Const FILTER_GLYPH As Long = &HE71C&
' Width of a funnel at the right end of its caption, and the gap between the
' caption's text area and the funnel (twips: 20px and 2px). GridHeaders and
' GridFit lay them out.
Public Const FILTER_W As Long = 300
Public Const FILTER_GAP As Long = 30
' Most values the filter dialog lists, and most characters of them: a value
' list's row source holds at most 32,750 characters.
Private Const MAX_VALUES As Long = 500
Private Const MAX_LIST_CHARS As Long = 30000
' Longest value text the dialog shows; longer ones end in an ellipsis.
Private Const MAX_SHOWN As Long = 80
' Longest Filter this module sets. Access SQL allows 64K characters; the
' rest is left for the page's filter and pagination's id list.
Private Const MAX_FILTER As Long = 30000
' What "(Blanks)" is keyed as in a column's kept values.
Private Const BLANK_KEY As String = "Null"

' Header filter state per table: key "<host name>.<subform control>" ->
' Scripting.Dictionary with
'   Hwnd    the host's window: a page opened again starts with no filters
'   Base    the page's own filter (SetGridFilter)
'   Last    the Filter this module last set (to spot one the page set itself)
'   Prefix  the captions' prefix (for the funnels' colors)
'   Cols    Scripting.Dictionary: column index -> Array(clause, kept keys
'           Scripting.Dictionary, kept count)
Private m_tables As Object

' The filter dialog's question (GridFilter): open, title, hint, values shown,
' which are selected, and the choice made ("" until OK: a "1" or "0" per value).
Private m_open As Boolean
Private m_title As String
Private m_hint As String
Private m_items As Variant
Private m_checked As Variant
Private m_choice As String

' ---------------------------------------------------------------------------
' Funnel click
' ---------------------------------------------------------------------------

' A GridHeaders funnel's click (set by GridHeaders ..., filterSf): list the
' values of column col of the table in subform control sfName on host in
' dlgFilter, then show only the rows whose value the user leaves selected.
' Selecting every listed value removes the column's filter; Cancel changes
' nothing.
' The values listed are those the other filters (the page's and the other
' columns') leave, in the column's sort order. Afterwards the host's onFilter
' sub (GridHeaders), if any, is called with sfName, so it can update counts or
' its empty state. A failure is shown (HandleError), as for any click.
Public Function GridFilter(host As Form, ByVal sfName As String, ByVal prefix As String, ByVal col As Long)
    Dim sf As Control, vals As Variant, st As Object, kept As Object, i As Long, keep() As Variant, n As Long
    Dim shown() As String, checked() As Boolean, chars As Long, more As Boolean, t() As String, entry As Variant
    On Error GoTo Fail
    If m_open Then Exit Function
    Set sf = host.Controls(sfName)
    If Not CanFilter(sf, prefix, col) Then
        KitMsgBox "This column can't be filtered: it holds long text or files.", vbInformation, "Filter"
        Exit Function
    End If
    vals = GridColumnValues(sf, prefix, col, more)
    If UBound(vals) < 0 Then
        KitMsgBox "No rows to filter: the table's other filters hide every row.", vbInformation, "Filter"
        Exit Function
    End If
    Set st = TableState(sf, prefix)
    If st("Cols").Exists(col) Then
        entry = st("Cols")(col)
        Set kept = entry(1)
    End If
    ReDim shown(UBound(vals)): ReDim checked(UBound(vals))
    For i = 0 To UBound(vals)
        shown(i) = ValueText(vals(i))
        chars = chars + Len(shown(i)) + 8
        If chars > MAX_LIST_CHARS Then
            ReDim Preserve shown(i - 1): ReDim Preserve checked(i - 1): ReDim Preserve vals(i - 1)
            more = True
            Exit For
        End If
        If kept Is Nothing Then checked(i) = True Else checked(i) = kept.Exists(ValueKey(vals(i)))
    Next
    m_title = "Filter by " & ColumnCaption(sf, col)
    m_hint = "Choose the values to show. Click a value to select or clear it."
    If more Then m_hint = "Choose the values to show. Only the first " & (UBound(shown) + 1) & _
        " values are listed: clearing any hides the unlisted ones too."
    m_items = shown
    m_checked = checked
    m_choice = ""
    m_open = True
    DoCmd.OpenForm "dlgFilter", WindowMode:=acDialog
    m_open = False
    If Len(m_choice) = 0 Then Exit Function
    If InStr(m_choice, "0") = 0 Then
        SetColumnFilter sf, prefix, col, Empty
    Else
        ReDim keep(UBound(vals))
        For i = 0 To UBound(vals)
            If Mid$(m_choice, i + 1, 1) = "1" Then keep(n) = vals(i): n = n + 1
        Next
        If n = 0 Then Exit Function      ' dlgFilter's OK needs a value selected
        ReDim Preserve keep(n - 1)
        SetColumnFilter sf, prefix, col, keep
    End If
    t = FunnelTag(host, prefix, col)
    If Len(t(5)) > 0 Then CallByName host, t(5), VbMethod, sfName
    Exit Function
Fail:
    m_open = False
    HandleError "modGridFilter.GridFilter", prefix & col
End Function

' What dlgFilter shows and the choice it hands back.
Public Function GridFilterTitle() As String
    GridFilterTitle = m_title
End Function

Public Function GridFilterHint() As String
    GridFilterHint = m_hint
End Function

' The values' text, in order ("(Blanks)" for blanks).
Public Function GridFilterItems() As Variant
    GridFilterItems = m_items
End Function

' Whether each value starts selected.
Public Function GridFilterChecked() As Variant
    GridFilterChecked = m_checked
End Function

' dlgFilter's OK: "1" for each value selected, "0" for each cleared.
Public Sub SetGridFilterChoice(ByVal choice As String)
    m_choice = choice
End Sub

' ---------------------------------------------------------------------------
' Filtering in code
' ---------------------------------------------------------------------------

' The page's own filter for a table with header filters (search box, chips):
' the table shows the rows matching where and every header filter. Use it in
' place of setting the table's Filter. "" shows every row the header filters
' leave. A paginated table goes back to page 1.
Public Sub SetGridFilter(sf As Control, ByVal where As String)
    Dim st As Object
    Set st = TableState(sf)
    st("Base") = where
    ApplyFilters sf, st
End Sub

' Filter column col of the table in subform control sf (captions prefix0,
' prefix1, ... on its host) to the rows whose value is in keep, an array of
' values (Null or "" keeps blanks), as if chosen in dlgFilter; keep Empty
' removes the column's filter. The funnel shows the state. The host's
' onFilter isn't called. Raises error 5 for a column with no funnel, or one
' that can't be filtered, and when the filter would be too long.
Public Sub SetColumnFilter(sf As Control, ByVal prefix As String, ByVal col As Long, ByVal keep As Variant)
    Dim st As Object, t() As String, fld As String, isText As Boolean, clause As String, kept As Object, v As Variant
    Dim lits As String, blank As Boolean, k As String, n As Long, cols As Object
    t = FunnelTag(sf.Parent, prefix, col)
    If Not CanFilter(sf, prefix, col) Then
        Err.Raise 5, "modGridFilter.SetColumnFilter", "Column " & col & " of " & sf.Name & " can't be filtered."
    End If
    Set st = TableState(sf, prefix)
    Set cols = st("Cols")
    If IsEmpty(keep) Then
        If cols.Exists(col) Then cols.Remove col
    Else
        If Not IsArray(keep) Then keep = Array(keep)
        fld = Bracket(t(3))
        isText = (FieldType(sf.Form, t(3)) = dbText)
        Set kept = CreateObject("Scripting.Dictionary")
        For Each v In keep
            If IsNull(v) Then
                k = BLANK_KEY
            ElseIf VarType(v) = vbString And Len(v) = 0 Then
                k = BLANK_KEY
            Else
                k = ValueKey(v)
            End If
            If Not kept.Exists(k) Then
                kept.Add k, True
                n = n + 1
                If k = BLANK_KEY Then blank = True Else lits = lits & "," & k
            End If
        Next
        If Len(lits) > 0 Then clause = fld & " IN (" & Mid$(lits, 2) & ")"
        If blank Then
            clause = clause & IIf(Len(clause) > 0, " OR ", "") & fld & " Is Null"
            If isText Then clause = clause & " OR " & fld & "=''"
        End If
        If Len(clause) = 0 Then clause = "False"       ' keep is an empty array: no row
        If Len(clause) > MAX_FILTER Then
            Err.Raise 5, "modGridFilter.SetColumnFilter", "Too many values to filter by. Choose fewer values, or select all of them."
        End If
        cols(col) = Array(clause, kept, n)
    End If
    ApplyFilters sf, st
End Sub

' Remove every header filter from the table in subform control sf, keeping
' the page's own filter (SetGridFilter). The funnels go back to muted.
Public Sub ClearGridFilters(sf As Control)
    Dim st As Object
    Set st = TableState(sf)
    st("Cols").RemoveAll
    ApplyFilters sf, st
End Sub

' The header filters of the table in subform control sf, ANDed ("" when no
' column is filtered), without the page's own filter: for a page that counts
' the rows its chips or tabs would show. It names the table's fields, so use it
' against the table's record source.
Public Function GridFilterWhere(sf As Control) As String
    GridFilterWhere = Combined(TableState(sf), -1, False)
End Function

' True when column col of the table in subform control sf is filtered.
Public Function GridColumnFiltered(sf As Control, ByVal col As Long) As Boolean
    Dim st As Object
    Set st = TableState(sf)
    GridColumnFiltered = st("Cols").Exists(col)
End Function

' The distinct values of column col (caption prefix & col) among the rows the
' page's filter and the other columns' filters leave, in the column's sort
' order (its sort field when GridHeaders has one), blanks (Null or "") last
' as one Null. At most MAX_VALUES; more is set True when there were others.
Public Function GridColumnValues(sf As Control, ByVal prefix As String, ByVal col As Long, Optional ByRef more As Boolean) As Variant
    Dim db As DAO.Database, rs As DAO.Recordset, t() As String, fld As String, srt As String, where As String
    Dim src As String, sel As String, vals() As Variant, n As Long, seen As Object, v As Variant, k As String
    Dim blank As Boolean
    On Error GoTo Fail
    more = False
    t = FunnelTag(sf.Parent, prefix, col)
    fld = "q." & Bracket(t(3))
    If Len(t(4)) > 0 Then srt = "q." & Bracket(t(4))
    where = Combined(TableState(sf, prefix), CLng(col), True)
    src = SourceSql(sf.Form)
    sel = "SELECT DISTINCT " & fld
    Set db = CurrentDb
    ' Sorted by the column's sort field when the record source returns it;
    ' else by the value. DISTINCT allows only ORDER BY terms it selects.
    If Len(srt) > 0 And srt <> fld Then
        On Error Resume Next
        Set rs = OpenRows(db, sel & ", " & srt & " FROM " & src & IIf(Len(where) > 0, " WHERE " & where, "") & " ORDER BY " & srt)
        On Error GoTo Fail
    End If
    If rs Is Nothing Then Set rs = OpenRows(db, sel & " FROM " & src & IIf(Len(where) > 0, " WHERE " & where, "") & " ORDER BY " & fld)
    Set seen = CreateObject("Scripting.Dictionary")
    ReDim vals(MAX_VALUES)
    Do While Not rs.EOF
        v = rs.Fields(0).Value
        If IsNull(v) Then
            blank = True
        ElseIf VarType(v) = vbString And Len(v) = 0 Then
            blank = True
        Else
            k = ValueKey(v)
            If Not seen.Exists(k) Then
                If n = MAX_VALUES Then more = True: Exit Do
                seen.Add k, True
                vals(n) = v
                n = n + 1
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    If blank Then vals(n) = Null: n = n + 1        ' the extra slot: MAX_VALUES + 1 in all
    If n = 0 Then
        GridColumnValues = Array()
    Else
        ReDim Preserve vals(n - 1)
        GridColumnValues = vals
    End If
    Exit Function
Fail:
    Dim num As Long, desc As String
    num = Err.Number: desc = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise num, "modGridFilter.GridColumnValues", desc
End Function

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

' The state for the table in subform control sf, made on first use with the
' table's current filter as the page's filter. A page opened again (another
' window) starts afresh. When the page has set the table's Filter itself since
' this module last did, the header filters are gone from the table: they are
' dropped here too, and the funnels go back to muted.
Private Function TableState(sf As Control, Optional ByVal prefix As String = "") As Object
    Dim host As Form, key As String, st As Object, full As String
    Set host = sf.Parent
    If m_tables Is Nothing Then Set m_tables = CreateObject("Scripting.Dictionary")
    key = host.Name & "." & sf.Name
    full = TableFullFilter(sf.Form)
    If m_tables.Exists(key) Then
        Set st = m_tables(key)
        If st("Hwnd") <> host.hWnd Then
            m_tables.Remove key
            Set st = Nothing
        End If
    End If
    If st Is Nothing Then
        Set st = CreateObject("Scripting.Dictionary")
        st("Hwnd") = host.hWnd
        st("Base") = full
        st("Last") = full
        st("Prefix") = prefix
        Set st("Cols") = CreateObject("Scripting.Dictionary")
        m_tables.Add key, st
    ElseIf full <> st("Last") Then
        st("Base") = full
        st("Last") = full
        st("Cols").RemoveAll
        StyleFunnels host, st
    End If
    If Len(prefix) > 0 Then st("Prefix") = prefix
    Set TableState = st
End Function

' Set the table's Filter to the page's filter AND the header filters, color
' the funnels, and send a paginated table back to page 1.
Private Sub ApplyFilters(sf As Control, st As Object)
    Dim g As Form, where As String
    where = Combined(st, -1, True)
    Set g = sf.Form
    g.Filter = where
    g.FilterOn = Len(where) > 0
    st("Last") = where
    StyleFunnels sf.Parent, st
    RestartTablePages sf.Parent, sf.Name
End Sub

' The page's filter (withBase) and every column's filter but skip's, each in
' brackets, joined by AND.
Private Function Combined(st As Object, ByVal skip As Long, ByVal withBase As Boolean) As String
    Dim k As Variant, s As String, cols As Object, entry As Variant
    If withBase And Len(st("Base")) > 0 Then s = "(" & st("Base") & ")"
    Set cols = st("Cols")
    For Each k In cols.Keys
        If k <> skip Then
            entry = cols(k)
            s = s & IIf(Len(s) > 0, " AND ", "") & "(" & entry(0) & ")"
        End If
    Next
    Combined = s
End Function

' Color each funnel of the table: CLR_PRIMARY when its column is filtered,
' else CLR_MUTED_FG, with a tip saying which.
Private Sub StyleFunnels(host As Form, st As Object)
    Dim i As Long, c As Control, nm As String, prefix As String, cols As Object, entry As Variant, fore As Long
    prefix = st("Prefix")
    If Len(prefix) = 0 Then Exit Sub
    Set cols = st("Cols")
    Do While CtlExists(host, prefix & i)
        nm = prefix & "Filter" & i
        If CtlExists(host, nm) Then
            Set c = host.Controls(nm)
            fore = Th(IIf(cols.Exists(i), CLR_PRIMARY, CLR_MUTED_FG))
            If c.ForeColor <> fore Then c.ForeColor = fore
            If cols.Exists(i) Then
                entry = cols(i)
                c.ControlTipText = "Filtered: showing " & entry(2) & IIf(entry(2) = 1, " value", " values") & ". Click to change."
            Else
                c.ControlTipText = "Filter"
            End If
        End If
        i = i + 1
    Loop
End Sub

' The funnel's Tag parts: "x|offset|width|field|sort field|onFilter".
' Raises error 5 when the caption has no funnel.
Private Function FunnelTag(host As Form, ByVal prefix As String, ByVal col As Long) As String()
    Dim nm As String
    nm = prefix & "Filter" & col
    If Not CtlExists(host, nm) Then Err.Raise 5, "modGridFilter", "Column " & col & " (" & prefix & ") has no filter funnel."
    FunnelTag = Split(host.Controls(nm).Tag & "||||||", "|")
End Function

' False for a column whose field is memo, OLE or an attachment (DISTINCT and
' IN don't work on them). Raises error 5 when the column has no funnel.
Private Function CanFilter(sf As Control, ByVal prefix As String, ByVal col As Long) As Boolean
    Select Case FieldType(sf.Form, FunnelTag(sf.Parent, prefix, col)(3))
        Case dbMemo, dbLongBinary, dbAttachment, dbComplexByte To dbComplexGUID
        Case Else: CanFilter = True
    End Select
End Function

' DAO type of the table's field fld (bracketed or not).
Private Function FieldType(g As Form, ByVal fld As String) As Long
    FieldType = g.RecordsetClone.Fields(Unbracket(fld)).Type
End Function

' The column's caption as BuildGrid wrote it (its Tag "caption|left|width").
Private Function ColumnCaption(sf As Control, ByVal col As Long) As String
    On Error Resume Next
    ColumnCaption = Split(sf.Form.Controls("txtCol" & col).Tag & "|", "|")(0)
End Function

' The table's record source as a FROM item named q: "(SELECT ...) AS q", its
' ORDER BY dropped, or "[name] AS q" for a table or query.
Private Function SourceSql(g As Form) As String
    Dim src As String, pos As Long
    src = Trim$(Replace(Replace(Replace(g.RecordSource, vbCr, " "), vbLf, " "), vbTab, " "))
    If Right$(src, 1) = ";" Then src = Trim$(Left$(src, Len(src) - 1))
    If UCase$(Left$(src, 7)) = "SELECT " Then
        pos = InStrRev(UCase$(src), " ORDER BY ")
        If pos > 0 Then src = Left$(src, pos - 1)
        SourceSql = "(" & src & ") AS q"
    Else
        SourceSql = Bracket(src) & " AS q"
    End If
End Function

' A snapshot of sql. Parameters (a record source naming a form's control, say)
' take their values from Eval, as the form's own recordset does.
Private Function OpenRows(db As DAO.Database, ByVal sql As String) As DAO.Recordset
    Dim qd As DAO.QueryDef, p As DAO.Parameter
    Set qd = db.CreateQueryDef("", sql)
    For Each p In qd.Parameters
        p.Value = Eval(p.Name)
    Next
    Set OpenRows = qd.OpenRecordset(dbOpenSnapshot)
End Function

' A value as Access SQL: 'text' (quotes doubled), #yyyy-mm-dd hh:nn:ss#, True
' or False, or a number with a "." decimal point whatever the locale. The
' same value always gives the same text, so it also keys the kept values.
Private Function ValueKey(ByVal v As Variant) As String
    Select Case VarType(v)
        Case vbNull: ValueKey = BLANK_KEY
        Case vbString: ValueKey = "'" & Replace(v, "'", "''") & "'"
        Case vbDate: ValueKey = "#" & Format$(v, "yyyy\-mm\-dd hh\:nn\:ss") & "#"
        Case vbBoolean: ValueKey = IIf(v, "True", "False")
        Case Else: ValueKey = Trim$(Str$(v))
    End Select
End Function

' A value as the dialog lists it.
Private Function ValueText(ByVal v As Variant) As String
    Dim s As String
    If IsNull(v) Then ValueText = "(Blanks)": Exit Function
    s = Replace(Replace(CStr(v), vbCr, " "), vbLf, " ")
    If Len(s) > MAX_SHOWN Then s = Left$(s, MAX_SHOWN - 1) & ChrW(&H2026)
    ValueText = s
End Function

' fld as "[fld]" (a name given as "[SHA-256]" keeps its brackets once).
Private Function Bracket(ByVal fld As String) As String
    Bracket = "[" & Unbracket(fld) & "]"
End Function

Private Function Unbracket(ByVal fld As String) As String
    fld = Trim$(fld)
    If Left$(fld, 1) = "[" And Right$(fld, 1) = "]" Then fld = Mid$(fld, 2, Len(fld) - 2)
    Unbracket = fld
End Function

' True when frm has a control named nm.
Private Function CtlExists(frm As Form, ByVal nm As String) As Boolean
    Dim c As Control
    On Error Resume Next
    Set c = frm.Controls(nm)
    CtlExists = Not c Is Nothing
End Function

' ---------------------------------------------------------------------------
' Build
' ---------------------------------------------------------------------------

' The filter dialog GridFilter opens (dlgFilter, code in forms\dlgFilter.vba):
' a hint, a list of the column's values where the selected ones show, the
' count selected and Select all under it, and OK and Cancel along the bottom
' right. Build it with the app's other dialogs.
Public Sub BuildFilterDialog()
    Dim c As Control, w As Long, listY As Long, listH As Long, rowY As Long, btnY As Long
    w = px(320)
    listY = px(64): listH = px(240)
    rowY = listY + listH + px(12)
    btnY = rowY + px(30) + px(20)
    NewDialog "Filter", w, btnY + px(30) + px(SP_DIALOG)
    Ev F, "OnLoad"
    Lbl "lblHint", "", px(SP_DIALOG), px(20), w - 2 * px(SP_DIALOG), px(36), 9, False, CLR_MUTED_FG
    ' Hidden position and the value; 20px left for the scroll bar.
    Set c = Lst("lstValues", "", px(SP_DIALOG), listY, w - 2 * px(SP_DIALOG), listH, 2, "0;" & (w - 2 * px(SP_DIALOG) - px(20)))
    c.RowSourceType = "Value List"
    c.ColumnHeads = False
    c.MultiSelect = 1                        ' simple: each click selects or clears a value
    Ev c, "AfterUpdate"
    Lbl "lblCount", "", px(SP_DIALOG), rowY + px(6), px(160), px(18), 9, False, CLR_MUTED_FG
    Btn "btnAll", "Select all", w - px(SP_DIALOG) - px(88), rowY, px(88), px(30), "ghost"
    Set c = Btn("btnOK", "OK", w - px(SP_DIALOG) - px(88) - px(8) - px(88), btnY, px(88), px(30), "primary")
    c.Default = True
    Set c = Btn("btnCancel", "Cancel", w - px(SP_DIALOG) - px(88), btnY, px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgFilter"
End Sub
