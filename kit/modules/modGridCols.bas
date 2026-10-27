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
'
' Widths: GridHeaders(..., resizeSf:=) puts a grip in the gap after each
' caption, and dragging it (GridSizeDown) changes that column's width. The
' widths go into the same layout, as ":<px>" after a column
' ("2:180,0:96,1h,3:120"), in the table's design pixels; a column without one
' keeps its design width. GridFit scales user widths like design ones, and
' also below them when the window is too narrow, so they never overflow the
' table (its captions sit on the host, which can't scroll sideways with the
' rows). Widening a column takes the width from the shown columns to its
' right, so the table's right edge stays. GridColumnsMenu offers the chooser
' and Reset widths.

Private Const SETTINGS_KEY As String = "Columns"
' Width GridFit leaves for the table's vertical scroll bar (modKitUI SCROLL_W).
Private Const SCROLL_W As Long = 300
' Narrowest a dragged column gets (px), and the widest width a layout keeps.
Private Const MIN_COL_PX As Long = 40
Private Const MAX_COL_PX As Long = 2000

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As LongPtr, ByVal lpCursorName As LongPtr) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private Const VK_LBUTTON As Long = 1
Private Const IDC_SIZEWE As Long = 32644
Private Const LOGPIXELSX As Long = 88

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

' Show the table's columns in the order, visibility and widths layout gives
' (see the top of this module), and move its captions with them. Without
' widths the columns keep their design widths and GridFit scales them to
' fill the table, so hiding one widens the rest. Raises error 5 for a layout
' that doesn't fit the table and for a table not built with fit:=True.
Public Sub SetGridColumns(sf As Control, ByVal prefix As String, ByVal layout As String)
    Dim g As Form, host As Form, order() As Long, shown() As Boolean, n As Long, k As Long
    Dim c As Control, t() As String, x As Long, w As Long, allW As Long, design As Long, widths As Variant, userW As Boolean
    Set g = sf.Form
    Set host = OwnerForm(sf)
    CheckFitGrid g
    n = ColumnCount(g)
    If Not ParseLayout(layout, n, order, shown) Then
        Err.Raise 5, "modGridCols.SetGridColumns", "The layout """ & layout & """ doesn't fit the " & n & " columns of " & g.Name & "."
    End If
    widths = LayoutWidths(layout, n)
    ' Rewrite each column's design position (Tag "caption|left|width", which
    ' GridFit scales) in the new order, keeping its design width as a fourth
    ' part. A hidden column gets no width at the left edge: GridFit scales
    ' hidden columns too, and at a large ratio their design width would push
    ' the form past its 22-inch limit (error 2100).
    x = px(8)
    For k = 0 To n - 1
        Set c = g.Controls("txtCol" & order(k))
        t = Split(c.Tag & "|||", "|")
        w = Val(t(2))
        If Val(t(3)) > 0 Then w = Val(t(3))
        allW = allW + w                     ' the design widths, shown or not
        ' Part 3 keeps the design width; part 2 is the width shown.
        If shown(k) Then
            If widths(order(k)) > 0 Then
                c.Tag = SetTagParts(c.Tag, 1, x, px(widths(order(k))), w)
                w = px(widths(order(k)))
                userW = True
            Else
                c.Tag = SetTagParts(c.Tag, 1, x, w, w)
            End If
            SetCaption host, prefix, order(k), x, w, True
            x = x + w
        Else
            c.Tag = SetTagParts(c.Tag, 1, 0, 0, w)
            SetCaption host, prefix, order(k), px(8), px(8), False     ' no width, as the column
        End If
        SetVisible c, shown(k)
    Next
    ' The table's Tag "grid:<width>" is the width GridFit scales from. Shrink
    ' it so the shown columns (x - px(8) wide) take the share of the table
    ' that all the columns take in the design, scroll bar aside: then they
    ' fill the table as far as the full set would. The design width is kept
    ' after a "|" (GridFit reads only the number before it). User widths
    ' that add up to the design's fill the same space, so a drag that moves
    ' width between columns doesn't move the table's right edge; "|u" lets
    ' GridFit scale them below their widths when they wouldn't fit.
    design = DesignWidth(g)
    g.Tag = "grid:" & (SCROLL_W + CLng(CDbl(design - SCROLL_W) * x / (allW + px(8)))) & "|" & design & IIf(userW, "|u", "")
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
    c.Tag = SetTagParts(c.Tag, 1, offset, w - px(8))   ' later parts (a sort field) stay
    SetVisible c, visible
    ' Its filter funnel (GridHeaders filterSf) and width grip go with it; GridFit places them.
    If CtlExists(host, prefix & "Filter" & i) Then SetVisible host.Controls(prefix & "Filter" & i), visible
    If CtlExists(host, prefix & "Size" & i) Then SetVisible host.Controls(prefix & "Size" & i), visible
End Sub

' tag with its "|"-separated parts from index first on replaced by vals; the
' other parts stay as they were.
Private Function SetTagParts(ByVal tag As String, ByVal first As Long, ParamArray vals() As Variant) As String
    Dim t() As String, i As Long
    t = Split(tag, "|")
    If UBound(t) < first + UBound(vals) Then ReDim Preserve t(first + UBound(vals))
    For i = 0 To UBound(vals)
        t(first + i) = CStr(vals(i))
    Next
    SetTagParts = Join(t, "|")
End Function

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
        ' The chooser sets order and visibility; widths dragged earlier stay.
        m_choice = WithWidths(m_choice, LayoutWidths(SavedLayout(sf), ColumnCount(sf.Form)))
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
    GridColumnLayout = StripWidths(SavedLayout(m_sf))
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
        If InStr(s, ":") > 0 Then               ' a width (see the top of the module)
            If Not WidthPart(Mid$(s, InStr(s, ":") + 1)) Then Exit Function
            s = Left$(s, InStr(s, ":") - 1)
        End If
        shown(k) = Right$(s, 1) <> "h"
        If Not shown(k) Then s = Left$(s, Len(s) - 1)
        If Len(s) = 0 Or Len(s) > 4 Or Not s Like String(Len(s), "#") Then Exit Function
        i = CLng(s)
        If i >= n Then Exit Function
        If seen(i) Then Exit Function
        seen(i) = True
        order(k) = i
        If shown(k) Then anyShown = True
    Next
    ParseLayout = anyShown
End Function

' True when s is a layout width: 1 to 4 digits, MIN_COL_PX to MAX_COL_PX.
Private Function WidthPart(ByVal s As String) As Boolean
    If Len(s) = 0 Or Len(s) > 4 Or Not s Like String(Len(s), "#") Then Exit Function
    WidthPart = CLng(s) >= MIN_COL_PX And CLng(s) <= MAX_COL_PX
End Function

' The widths a layout gives, by column index (0 to n - 1, in px; 0 where a
' column has none). All zeros for a layout that doesn't fit n columns.
Public Function LayoutWidths(ByVal layout As String, ByVal n As Long) As Variant
    Dim w() As Long, order() As Long, shown() As Boolean, parts() As String, k As Long, s As String
    If n < 1 Then LayoutWidths = Array(): Exit Function
    ReDim w(n - 1)
    If ParseLayout(layout, n, order, shown) Then
        parts = Split(layout, ",")
        For k = 0 To n - 1
            s = Trim$(parts(k))
            If InStr(s, ":") > 0 Then w(order(k)) = CLng(Mid$(s, InStr(s, ":") + 1))
        Next
    End If
    LayoutWidths = w
End Function

' layout without its widths: "2:180,0,1h:96" gives "2,0,1h".
Public Function StripWidths(ByVal layout As String) As String
    Dim parts() As String, k As Long, s As String
    If Len(layout) = 0 Then Exit Function
    parts = Split(layout, ",")
    For k = 0 To UBound(parts)
        s = Trim$(parts(k))
        If InStr(s, ":") > 0 Then s = Left$(s, InStr(s, ":") - 1)
        StripWidths = StripWidths & IIf(k > 0, ",", "") & s
    Next
End Function

' layout (order and visibility) with widths (by column index, 0 for none)
' added to its columns.
Private Function WithWidths(ByVal layout As String, widths As Variant) As String
    Dim parts() As String, k As Long, s As String, i As Long
    parts = Split(StripWidths(layout), ",")
    For k = 0 To UBound(parts)
        s = parts(k)
        i = Val(s)
        If i <= UBound(widths) Then
            If widths(i) > 0 Then s = s & ":" & widths(i)
        End If
        WithWidths = WithWidths & IIf(k > 0, ",", "") & s
    Next
End Function

' ---------------------------------------------------------------------------
' Column widths
' ---------------------------------------------------------------------------

' A width grip's mouse down (GridHeaders resizeSf): follow the pointer while
' the left button is held, the grip showing where the column will end, then
' set the width (ResizeGridColumn). frm is the form holding the captions,
' sfName the table's subform control on it, i the column.
Public Function GridSizeDown(frm As Form, ByVal sfName As String, ByVal prefix As String, ByVal i As Long)
    Dim grip As Control, pt As POINTAPI, x0 As Long, left0 As Long, minLeft As Long, maxLeft As Long
    Dim tpp As Double, newLeft As Long, oldBack As Integer
    On Error GoTo Fail
    If (GetAsyncKeyState(VK_LBUTTON) And &H8000) = 0 Then Exit Function    ' another button
    Set grip = frm.Controls(prefix & "Size" & i)
    tpp = TwipsPerPixel()
    GetCursorPos pt
    x0 = pt.x
    left0 = grip.Left
    ' The column can't get narrower than MIN_COL_PX, and the grip stays on the form.
    minLeft = frm.Controls(prefix & i).Left + px(MIN_COL_PX) - px(8)
    maxLeft = frm.Width - grip.Width
    oldBack = grip.BackStyle
    grip.BackStyle = 1
    grip.BackColor = Th(CLR_PRIMARY_SOFT)
    newLeft = left0
    Do While (GetAsyncKeyState(VK_LBUTTON) And &H8000) <> 0
        SetCursor LoadCursor(0, IDC_SIZEWE)
        GetCursorPos pt
        newLeft = left0 + CLng((pt.x - x0) * tpp)
        If newLeft < minLeft Then newLeft = minLeft
        If newLeft > maxLeft Then newLeft = maxLeft
        If grip.Left <> newLeft Then grip.Left = newLeft
        DoEvents
        Sleep 15
    Loop
    grip.BackStyle = oldBack
    If Abs(newLeft - left0) >= px(2) Then ResizeGridColumn frm.Controls(sfName), prefix, i, newLeft - left0 Else grip.Left = left0
    Exit Function
Fail:
    On Error Resume Next
    If Not grip Is Nothing Then grip.BackStyle = 0
    On Error GoTo 0
    HandleError "modGridCols.GridSizeDown", sfName & " column " & i
End Function

' The pointer over a width grip shows the left-right sizing arrows.
Public Function GridSizeHover()
    SetCursor LoadCursor(0, IDC_SIZEWE)
End Function

' Make column i of the table in sf dx twips wider (narrower when negative),
' as the user sees it, and save the widths. The shown columns to its right
' give up the difference in proportion to what they have above MIN_COL_PX
' (or share what it gives up, in proportion to their widths), so the
' columns to its left and the table's width stay and the divider lands
' where it was dropped. A column can grow only by what the columns to its
' right can give up. The last shown column has none: its right edge is the
' table's, so dragging it changes nothing. Every shown column's width goes
' into the layout, so the table now fills by these widths; a hidden
' column keeps the width saved for it.
Public Sub ResizeGridColumn(sf As Control, ByVal prefix As String, ByVal i As Long, ByVal dx As Long)
    Dim g As Form, n As Long, layout As String, order() As Long, shown() As Boolean, w() As Double, saved As Variant
    Dim k As Long, pos As Long, ratio As Double, d As Double, spare As Double, rightW As Double, donors As Long, out As String
    Set g = sf.Form
    CheckFitGrid g
    n = ColumnCount(g)
    layout = SavedLayout(sf)
    If Len(layout) = 0 Then layout = DefaultLayout(n)
    ParseLayout layout, n, order, shown
    saved = LayoutWidths(layout, n)
    ' Each shown column's width now, in design px (part 2 of its Tag, which
    ' GridFit scales by ratio to what shows).
    ReDim w(n - 1)
    For k = 0 To n - 1
        w(k) = Val(Split(g.Controls("txtCol" & k).Tag & "|||", "|")(2)) / px(1)
        If w(k) <= 0 Then w(k) = MIN_COL_PX
        If order(k) = i Then pos = k
    Next
    For k = pos + 1 To n - 1
        If shown(k) Then
            donors = donors + 1
            rightW = rightW + w(order(k))
            spare = spare + w(order(k)) - MIN_COL_PX
        End If
    Next
    If donors = 0 Then Exit Sub                     ' the last shown column
    If spare < 0 Then spare = 0
    ratio = (g.Controls("txtCol" & i).Width + px(8)) / px(w(i))
    If ratio <= 0 Then ratio = 1
    d = dx / ratio / px(1)                          ' design px
    If w(i) + d < MIN_COL_PX Then d = MIN_COL_PX - w(i)
    If d > spare Then d = spare                     ' the columns to the right stop at MIN_COL_PX
    If d = 0 Then Exit Sub
    For k = pos + 1 To n - 1
        If shown(k) Then
            If d > 0 Then
                w(order(k)) = w(order(k)) - d * (w(order(k)) - MIN_COL_PX) / spare
            Else
                w(order(k)) = w(order(k)) - d * w(order(k)) / rightW
            End If
        End If
    Next
    w(i) = w(i) + d
    For k = 0 To n - 1
        out = out & IIf(k > 0, ",", "") & order(k) & IIf(shown(k), "", "h")
        If shown(k) Then
            out = out & ":" & ClampW(w(order(k)))
        ElseIf saved(order(k)) > 0 Then
            out = out & ":" & saved(order(k))
        End If
    Next
    SetGridColumns sf, prefix, out
    SaveSetting APP_NAME, SETTINGS_KEY, g.Name, out
End Sub

' A width for a layout: whole px within MIN_COL_PX to MAX_COL_PX.
Private Function ClampW(ByVal w As Double) As Long
    ClampW = CLng(w)
    If ClampW < MIN_COL_PX Then ClampW = MIN_COL_PX
    If ClampW > MAX_COL_PX Then ClampW = MAX_COL_PX
End Function

' Back to the design widths, keeping the chosen order and hidden columns.
Public Sub ResetGridWidths(sf As Control, Optional ByVal prefix As String = "")
    Dim layout As String
    On Error GoTo Fail
    CheckFitGrid sf.Form
    layout = StripWidths(SavedLayout(sf))
    If Len(layout) = 0 Then layout = DefaultLayout(ColumnCount(sf.Form))
    SetGridColumns sf, prefix, layout
    If layout = DefaultLayout(ColumnCount(sf.Form)) Then
        DeleteLayout sf.Form.Name
    Else
        SaveSetting APP_NAME, SETTINGS_KEY, sf.Form.Name, layout
    End If
    Exit Sub
Fail:
    HandleError "modGridCols.ResetGridWidths", sf.Name
End Sub

' True when the user has set widths on the table in sf.
Public Function GridHasWidths(sf As Control) As Boolean
    GridHasWidths = InStr(SavedLayout(sf), ":") > 0
End Function

' A Columns button's click when the table's widths can be dragged: a menu
' under the button with Choose columns (ChooseGridColumns) and Reset widths
' (ResetGridWidths, shown once widths were changed).
Public Sub GridColumnsMenu(btn As Control, sf As Control, Optional ByVal prefix As String = "")
    Dim pick As String
    On Error GoTo Fail
    If GridHasWidths(sf) Then
        pick = KitMenu(btn, "cols|Choose columns" & ChrW(8230) & "|E8A9", "widths|Reset widths|E72C")
    Else
        pick = KitMenu(btn, "cols|Choose columns" & ChrW(8230) & "|E8A9")
    End If
    Select Case pick
        Case "cols": ChooseGridColumns sf, prefix
        Case "widths": ResetGridWidths sf, prefix
    End Select
    Exit Sub
Fail:
    HandleError "modGridCols.GridColumnsMenu", sf.Name
End Sub

' Remove a table's saved layout (none saved is fine).
Private Sub DeleteLayout(ByVal formName As String)
    On Error Resume Next
    DeleteSetting APP_NAME, SETTINGS_KEY, formName
End Sub

' Twips in one screen pixel (15 at 96 dpi).
Private Function TwipsPerPixel() As Double
    Dim hdc As LongPtr, dpi As Long
    hdc = GetDC(0)
    dpi = GetDeviceCaps(hdc, LOGPIXELSX)
    ReleaseDC 0, hdc
    If dpi <= 0 Then dpi = 96
    TwipsPerPixel = 1440 / dpi
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

' The width BuildGrid gave the table, kept in its Tag "grid:<width>[|<design>
' [|u]]".
Private Function DesignWidth(g As Form) As Long
    Dim t() As String
    t = Split(Mid$(g.Tag, 6), "|")
    DesignWidth = Val(t(IIf(UBound(t) >= 1, 1, 0)))
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
    sideX = px(SP_DIALOG) + px(220) + px(12)
    btnY = listY + listH + px(20)
    NewDialog "Columns", px(360), btnY + px(30) + px(SP_DIALOG)
    Ev F, "OnLoad"
    Set c = Lbl("lblHint", "Choose the columns this table shows and their order. Double-click a column to hide or show it.", _
        px(SP_DIALOG), px(20), px(312), px(36), 9, False, CLR_MUTED_FG)
    ' Hidden position, caption (130px) and state (70px): 20px left for the scroll bar.
    Set c = Lst("lstCols", "", px(SP_DIALOG), listY, px(220), listH, 3, "0;" & px(130) & ";" & px(70))
    c.RowSourceType = "Value List"
    c.ColumnHeads = False
    Ev c, "AfterUpdate", "OnDblClick"
    Btn "btnUp", "Move up", sideX, listY, px(80), px(30), "outline"
    Btn "btnDown", "Move down", sideX, listY + px(38), px(80), px(30), "outline"
    Btn "btnToggle", "Hide", sideX, listY + px(76), px(80), px(30), "outline"
    Btn "btnReset", "Reset", sideX, listY + px(126), px(80), px(30), "ghost"
    Set c = Btn("btnApply", "Apply", px(360) - px(SP_DIALOG) - px(88) - px(8) - px(88), btnY, px(88), px(30), "primary")
    c.Default = True
    Set c = Btn("btnCancel", "Cancel", px(360) - px(SP_DIALOG) - px(88), btnY, px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgColumns"
End Sub
