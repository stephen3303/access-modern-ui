Option Compare Database
Option Explicit

' Access UI kit - column chooser for BuildGrid tables built with fit:=True
' (the tables GridFit scales). From a Columns button the user hides, restores
' and reorders the table's columns in dlgColumns (ChooseGridColumns); the
' captions GridHeaders drew follow them, and so do PrintGrid and ExportGrid.
'
' The choice is kept per Windows user and table (registry, HKCU\Software\VB
' and VBA Program Settings\<APP_NAME>\Columns, value = the table's form name)
' and put back when the page opens (RestoreGridColumns).
'
' A layout lists the table's columns in screen order by their index in the
' BuildGrid cols, "h" marking a hidden one: "2,0,1h,3" shows column 2, then
' 0, then 3, and hides 1. A layout must name every column once and show at
' least one; anything else (say, saved before a rebuild added a column) is
' ignored and the table keeps its design.

Private Const SETTINGS_KEY As String = "Columns"

' The table the open dlgColumns is for (ChooseGridColumns), and the layout
' chosen in it ("" while none is).
Private m_sf As Control
Private m_choice As String

' ---------------------------------------------------------------------------
' Applying a layout
' ---------------------------------------------------------------------------

' Put back the user's saved layout for the table in subform control sf, with
' its captions prefix0, prefix1, ... on the host. Call it once the page has
' set the table's record source, and again wherever the page shows the
' captions again (SetVisible on every caption would show hidden ones). A
' missing or out-of-date layout leaves the design; a failure is logged, not
' shown, so the page still opens.
Public Sub RestoreGridColumns(sf As Control, Optional ByVal prefix As String = "")
    Dim layout As String
    On Error GoTo Fail
    layout = SavedLayout(sf)
    If Len(layout) > 0 Then SetGridColumns sf, prefix, layout
    Exit Sub
Fail:
    LogError "modGridCols.RestoreGridColumns", sf.Name
End Sub

' Show the table's columns in the order and visibility layout gives (see the
' top of this module), and move its captions with them. The columns keep
' their design widths and GridFit scales them to fill the table, so hiding
' one widens the rest. Raises error 5 for a layout that doesn't fit the table
' and for a table not built with fit:=True.
Public Sub SetGridColumns(sf As Control, ByVal prefix As String, ByVal layout As String)
    Dim g As Form, host As Form, order() As Long, shown() As Boolean, n As Long, k As Long
    Dim c As Control, t() As String, x As Long, w As Long, hiddenW As Long, design As Long
    Set g = sf.Form
    Set host = sf.Parent
    CheckFitGrid g
    n = ColumnCount(g)
    If Not ParseLayout(layout, n, order, shown) Then
        Err.Raise 5, "modGridCols.SetGridColumns", "The layout """ & layout & """ doesn't fit the " & n & " columns of " & g.Name & "."
    End If
    ' Rewrite each column's design position (Tag "caption|left|width", which
    ' GridFit scales) in the new order. Hidden columns keep a position at the
    ' left edge so they never push the form wider.
    x = px(8)
    For k = 0 To n - 1
        Set c = g.Controls("txtCol" & order(k))
        t = Split(c.Tag, "|")
        w = Val(t(2))
        If shown(k) Then
            c.Tag = t(0) & "|" & x & "|" & w
            SetCaption host, prefix, order(k), x, w, True
            x = x + w
        Else
            c.Tag = t(0) & "|" & px(8) & "|" & w
            SetCaption host, prefix, order(k), px(8), w, False
            hiddenW = hiddenW + w
        End If
        SetVisible c, shown(k)
    Next
    ' The table's Tag "grid:<width>" is the width GridFit scales from. Taking
    ' the hidden columns off it lets the shown ones fill the table; the design
    ' width is kept after a "|" (GridFit reads only the number before it).
    design = DesignWidth(g)
    g.Tag = "grid:" & (design - hiddenW) & "|" & design
    GridFit sf, prefix
End Sub

' Caption prefix<i> (GridHeaders, Tag "x|offset|width") moved to sit over a
' column at offset, w wide, and shown or hidden with it.
Private Sub SetCaption(host As Form, ByVal prefix As String, ByVal i As Long, ByVal offset As Long, _
                       ByVal w As Long, ByVal visible As Boolean)
    Dim c As Control
    If Len(prefix) = 0 Then Exit Sub
    If Not CtlExists(host, prefix & i) Then Exit Sub
    Set c = host.Controls(prefix & i)
    c.Tag = Split(c.Tag, "|")(0) & "|" & offset & "|" & (w - px(8))
    SetVisible c, visible
End Sub

' ---------------------------------------------------------------------------
' The chooser
' ---------------------------------------------------------------------------

' A Columns button's click: open the column chooser (dlgColumns) for the
' table in subform control sf, then show and save the layout the user
' applies. Cancel changes nothing.
Public Sub ChooseGridColumns(sf As Control, Optional ByVal prefix As String = "")
    On Error GoTo Fail
    If Not m_sf Is Nothing Then Exit Sub        ' already open
    CheckFitGrid sf.Form
    Set m_sf = sf
    m_choice = ""
    DoCmd.OpenForm "dlgColumns", WindowMode:=acDialog
    Set m_sf = Nothing
    If Len(m_choice) > 0 Then
        SetGridColumns sf, prefix, m_choice
        SaveSetting APP_NAME, SETTINGS_KEY, sf.Form.Name, m_choice
    End If
    Exit Sub
Fail:
    Set m_sf = Nothing
    HandleError "modGridCols.ChooseGridColumns", sf.Name
End Sub

' What dlgColumns shows: the open table's column captions in design order
' (from each column's Tag), and its current layout.
Public Function GridColumnCaptions() As Variant
    Dim g As Form, n As Long, i As Long, caps() As String
    GridColumnCaptions = Array()
    If m_sf Is Nothing Then Exit Function
    Set g = m_sf.Form
    n = ColumnCount(g)
    If n = 0 Then Exit Function
    ReDim caps(n - 1)
    For i = 0 To n - 1
        caps(i) = Split(g.Controls("txtCol" & i).Tag & "|", "|")(0)
    Next
    GridColumnCaptions = caps
End Function

Public Function GridColumnLayout() As String
    If m_sf Is Nothing Then Exit Function
    GridColumnLayout = SavedLayout(m_sf)
    If Len(GridColumnLayout) = 0 Then GridColumnLayout = DefaultLayout(ColumnCount(m_sf.Form))
End Function

' dlgColumns' Apply.
Public Sub SetGridColumnChoice(ByVal layout As String)
    m_choice = layout
End Sub

' Every column shown, in design order: "0,1,2,...".
Public Function DefaultLayout(ByVal n As Long) As String
    Dim i As Long
    For i = 0 To n - 1
        DefaultLayout = DefaultLayout & IIf(i > 0, ",", "") & i
    Next
End Function

' Split a layout into the column order and whether each is shown. False when
' it doesn't name each of the n columns exactly once, or shows none.
Public Function ParseLayout(ByVal layout As String, ByVal n As Long, order() As Long, shown() As Boolean) As Boolean
    Dim parts() As String, k As Long, s As String, i As Long, seen() As Boolean, anyShown As Boolean
    If n = 0 Or Len(layout) = 0 Then Exit Function
    parts = Split(layout, ",")
    If UBound(parts) <> n - 1 Then Exit Function
    ReDim order(n - 1): ReDim shown(n - 1): ReDim seen(n - 1)
    For k = 0 To n - 1
        s = LCase$(Trim$(parts(k)))
        shown(k) = Right$(s, 1) <> "h"
        If Not shown(k) Then s = Left$(s, Len(s) - 1)
        If Len(s) = 0 Or Not s Like String(Len(s), "#") Then Exit Function
        i = CLng(s)
        If i >= n Then Exit Function
        If seen(i) Then Exit Function
        seen(i) = True
        order(k) = i
        If shown(k) Then anyShown = True
    Next
    ParseLayout = anyShown
End Function

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

' The layout saved for the table in sf, if it still fits it; else "".
Private Function SavedLayout(sf As Control) As String
    Dim layout As String, order() As Long, shown() As Boolean
    layout = GetSetting(APP_NAME, SETTINGS_KEY, sf.Form.Name, "")
    If ParseLayout(layout, ColumnCount(sf.Form), order, shown) Then SavedLayout = layout
End Function

' Number of columns BuildGrid made (txtCol0, txtCol1, ...).
Private Function ColumnCount(g As Form) As Long
    Do While CtlExists(g, "txtCol" & ColumnCount)
        ColumnCount = ColumnCount + 1
    Loop
End Function

' The width BuildGrid gave the table, kept in its Tag "grid:<width>[|<design>]".
Private Function DesignWidth(g As Form) As Long
    Dim t() As String
    t = Split(Mid$(g.Tag, 6), "|")
    DesignWidth = Val(t(UBound(t)))
End Function

' Raise error 5 unless g is a BuildGrid table built with fit:=True. Only those
' can move columns: in any other the last column stretches with the window,
' so it can't sit anywhere but last. BuildGrid anchors that table's row
' background to both sides; a fit table's isn't anchored.
Private Sub CheckFitGrid(g As Form)
    Dim ok As Boolean
    If Left$(Nz(g.Tag, ""), 5) = "grid:" And CtlExists(g, "txtBg") And CtlExists(g, "txtCol0") Then ok = g.Controls("txtBg").HorizontalAnchor <> 2
    If Not ok Then Err.Raise 5, "modGridCols", "The table " & g.Name & " can't choose its columns: build it with BuildGrid fit:=True and size it with GridFit."
End Sub

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

' The column chooser ChooseGridColumns opens (dlgColumns, code in
' forms\dlgColumns.vba): a hint, a list of the table's columns with whether
' each shows, Move up, Move down, Hide/Show and Reset beside it, and Apply and
' Cancel along the bottom right. Build it with the app's other dialogs.
Public Sub BuildColumnsDialog()
    Dim c As Control, listY As Long, listH As Long, sideX As Long, btnY As Long
    listY = px(64): listH = px(240)
    sideX = SP_DIALOG + px(220) + px(12)
    btnY = listY + listH + px(20)
    NewDialog "Columns", px(360), btnY + px(30) + SP_DIALOG
    Ev F, "OnLoad"
    Set c = Lbl("lblHint", "Choose the columns this table shows and their order. Double-click a column to hide or show it.", _
        SP_DIALOG, px(20), px(312), px(36), 9, False, CLR_MUTED_FG)
    ' Hidden index, caption, state. 200 + 80px leaves room for the scroll bar.
    Set c = Lst("lstCols", "", SP_DIALOG, listY, px(220), listH, 3, "0;" & px(130) & ";" & px(70))
    c.RowSourceType = "Value List"
    c.ColumnHeads = False
    Ev c, "AfterUpdate", "OnDblClick"
    Btn "btnUp", "Move up", sideX, listY, px(80), px(30), "outline"
    Btn "btnDown", "Move down", sideX, listY + px(38), px(80), px(30), "outline"
    Btn "btnToggle", "Hide", sideX, listY + px(76), px(80), px(30), "outline"
    Btn "btnReset", "Reset", sideX, listY + px(126), px(80), px(30), "ghost"
    Set c = Btn("btnApply", "Apply", px(360) - SP_DIALOG - px(88) - px(8) - px(88), btnY, px(88), px(30), "primary")
    c.Default = True
    Set c = Btn("btnCancel", "Cancel", px(360) - SP_DIALOG - px(88), btnY, px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgColumns"
End Sub
