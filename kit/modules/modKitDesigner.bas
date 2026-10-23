Option Compare Database
Option Explicit

' Access UI kit - the form designer (prototype): a page where kit elements
' are dragged from a palette onto a canvas, moved, resized, and saved as a
' layout. Layouts are rows in kit_layout and kit_layout_item (made the first
' time the page opens), so a later step can build a real form from them with
' the kit's builders (Card, Btn, Lbl, Progress).
'
' Access can't create controls on a form that is open in Form view, so the
' page is built with a fixed pool of hidden preview controls per element
' (DZ_CARDS cards, DZ_PROGS progress bars, DZ_LBLS labels, DZ_BTNS
' buttons), and placing an element shows the next free one. The pools are
' built back to front (cards, progress, labels, buttons), so content always
' draws over the cards it sits on. A transparent button over the canvas
' (dzGlass) takes every mouse event and finds the element under the pointer;
' a transparent button over each palette tile (dzTool0..3) does the same for
' dragging a new element in. A pressed button keeps getting MouseMove until
' the mouse is released, wherever the pointer goes.
'
' Build the page with BuildDesigner (it attaches kit\forms\pgDesigner.vba);
' its code-behind hands every event to the Designer* routines here.
' Positions and sizes are in px from the canvas's top-left corner, on an
' 8px grid (DZ_GRID). Esc cancels a palette click; Delete removes the
' selected element; the arrow keys move it 8px (1px with Shift).

Private Const DZ_GRID As Long = 8               ' snap, px
Private Const DZ_GRIP As Long = 8               ' resize handle, px
Private Const DZ_CARDS As Long = 6
Private Const DZ_PROGS As Long = 8
Private Const DZ_LBLS As Long = 16
Private Const DZ_BTNS As Long = 12
Private Const DZ_MAX As Long = 42               ' all four pools
Private Const DZ_TOOLS As String = "card;button;label;progress"   ' palette order (dzTool0..3)

Private Const MODE_IDLE As Integer = 0
Private Const MODE_MOVE As Integer = 1
Private Const MODE_SIZE As Integer = 2
Private Const MODE_DRAG As Integer = 3          ' dragging a new element in from the palette
Private Const MODE_ARMED As Integer = 4         ' palette element clicked; the next canvas click places it

' The layout being edited: items 1..m_n, each in slot m_slot of its kind's pool.
Private m_kind(1 To DZ_MAX) As String
Private m_slot(1 To DZ_MAX) As Long
Private m_x(1 To DZ_MAX) As Long
Private m_y(1 To DZ_MAX) As Long
Private m_w(1 To DZ_MAX) As Long
Private m_h(1 To DZ_MAX) As Long
Private m_text(1 To DZ_MAX) As String
Private m_n As Long
Private m_sel As Long                           ' selected item, 0 for none
Private m_mode As Integer
Private m_tool As String                        ' kind being dragged in or armed
Private m_offX As Long                          ' pointer to the item's corner, px
Private m_offY As Long
Private m_gx As Long                            ' where a dragged-in element would land, px
Private m_gy As Long
Private m_gOn As Boolean                        ' ... and whether the pointer is over the canvas
Private m_cw As Long                            ' canvas size, px
Private m_ch As Long
Private m_layoutId As Long                      ' kit_layout.id, 0 until saved
Private m_dirty As Boolean

' ---------------------------------------------------------------------------
' Build
' ---------------------------------------------------------------------------

' The designer page, w x h twips with x twips of padding on each side:
' Elements (the palette) and Selected (the selected element's text, position
' and Delete) on the left, the layout list, name, New and Save above the
' canvas on the right. Saved as formName with pgDesigner.vba's code.
Public Sub BuildDesigner(ByVal formName As String, ByVal x As Long, ByVal w As Long, ByVal h As Long)
    Dim c As Control, i As Long, ty As Long, py As Long, lw As Long, cx As Long, cy As Long, cw As Long, ch As Long
    Dim kinds() As String
    NewForm "", 0, w, h
    F.KeyPreview = True                          ' Delete, Esc and the arrow keys (Form_KeyDown)
    Ev F, "OnLoad", "OnKeyDown"
    PageHeader "Form designer", "Drag an element from Elements onto the canvas, or click one and then click the canvas. " & _
        "Drag an element to move it, and its corner handle to resize it. Everything snaps to an 8px grid. " & _
        "Save keeps the layout under its name; open it again from the list.", x
    lw = px(200)
    cx = x + lw + px(16)
    cw = w - x - cx
    cy = px(136)
    ch = h - cy - px(16)

    ' Palette: a tile per element, with a transparent button over it for the mouse.
    ty = px(76)
    Card "DzTools", "Elements", x, ty, lw, px(204)
    kinds = Split(DZ_TOOLS, ";")
    For i = 0 To 3
        Box "dzTile" & i, x + px(16), ty + px(44) + i * px(38), lw - px(32), px(32), CLR_BG, CLR_BORDER
        Set c = Lbl("dzTile" & i & "_g", ChrW(CLng("&H" & Choose(i + 1, "E8A9", "E73E", "E8D2", "E9D2"))), _
            x + px(26), ty + px(44) + i * px(38), px(20), px(32), 11, False, CLR_PRIMARY)
        c.FontName = "Segoe MDL2 Assets"
        VCenter c, px(32), 11, True
        Set c = Lbl("dzTile" & i & "_n", KindCaption(kinds(i)), x + px(54), ty + px(44) + i * px(38), lw - px(86), px(32), 9, True)
        VCenter c, px(32), 9, True                ' centred on the tile, as the glyph is
        Set c = Glass("dzTool" & i, x + px(16), ty + px(44) + i * px(38), lw - px(32), px(32))
        c.ControlTipText = "Drag onto the canvas, or click and then click the canvas"
    Next

    ' The selected element.
    py = ty + px(204) + px(12)
    Card "DzProps", "Selected", x, py, lw, h - py - px(16)
    Lbl "lblDzKind", "Nothing selected", x + px(16), py + px(40), lw - px(32), px(20), 9, False, CLR_MUTED_FG
    Lbl "lblDzTextCap", "TEXT", x + px(16), py + px(66), lw - px(32), px(16), 8, True, CLR_MUTED_FG
    Set c = Txt("txtDzText", "", x + px(16), py + px(84), lw - px(32), px(30))
    c.Enabled = False
    Ev c, "AfterUpdate"
    Lbl "lblDzPos", "", x + px(16), py + px(122), lw - px(32), px(36), 9, False, CLR_MUTED_FG
    Set c = Btn("btnDzDelete", "Delete", x + px(16), py + px(166), px(84), px(28), "outline")
    c.Enabled = False

    ' Layout list, name, New and Save.
    Lbl "lblDzOpen", "OPEN A LAYOUT", cx, px(76), px(220), px(16), 8, True, CLR_MUTED_FG
    Set c = Cbo("cboDzOpen", "", "", cx, px(94), px(220), px(30), 2, "0;" & px(220))
    Ev c, "AfterUpdate"
    Lbl "lblDzName", "LAYOUT NAME", cx + px(232), px(76), px(220), px(16), 8, True, CLR_MUTED_FG
    Set c = Txt("txtDzName", "", cx + px(232), px(94), px(220), px(30))
    Ev c, "AfterUpdate"
    Btn "btnDzNew", "New", cx + cw - px(176), px(95), px(84), px(28), "outline"
    Btn "btnDzSave", "Save", cx + cw - px(84), px(95), px(84), px(28)

    ' Canvas, its empty-state hint, then the pools back to front.
    Box "dzCanvas", cx, cy, cw, ch, CLR_BG, CLR_BORDER
    Set c = Lbl("dzHint", "Drag an element from Elements onto this canvas, or click one and then click here.", _
        cx + px(40), cy + ch \ 2 - px(20), cw - px(80), px(40), 10, False, CLR_MUTED_FG)
    c.TextAlign = 2
    For i = 1 To DZ_CARDS
        Set c = Box("dzCard" & i, cx, cy, px(240), px(140), CLR_SURFACE, CLR_BORDER): c.Visible = False
        Set c = Lbl("dzCard" & i & "_t", "", cx + px(14), cy + px(10), px(212), px(20), 10, True): c.Visible = False
    Next
    For i = 1 To DZ_PROGS
        Set c = Box("dzProg" & i, cx, cy, px(160), px(10), CLR_BORDER): c.Visible = False
        Set c = Box("dzProg" & i & "_f", cx, cy, px(64), px(10), CLR_PRIMARY): c.Visible = False
    Next
    For i = 1 To DZ_LBLS
        Set c = Lbl("dzLbl" & i, "", cx, cy, px(120), px(20), 10): c.Visible = False
    Next
    For i = 1 To DZ_BTNS
        Set c = Lbl("dzBtn" & i, "", cx, cy, px(96), px(32), 9, True, CLR_BTN_FG)
        c.BackStyle = 1
        c.BackColor = CLR_BTN_BG
        c.TextAlign = 2
        c.Visible = False
    Next
    ' Selection outline and its resize handle, over every element.
    Set c = Box("dzSel", cx, cy, px(16), px(16), -1, CLR_PRIMARY): c.BorderWidth = 1: c.Visible = False
    Set c = Box("dzGrip", cx, cy, px(DZ_GRIP), px(DZ_GRIP), CLR_PRIMARY, CLR_SURFACE): c.Visible = False
    ' Last, so it is on top: the canvas's mouse.
    Glass "dzGlass", cx, cy, cw, ch
    SaveForm formName, "pgDesigner"
End Sub

' A transparent button that takes the mouse for what is drawn under it.
Private Function Glass(ByVal nm As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As Control
    Dim c As Control
    Set c = Btn(nm, "", x, y, w, h, "ghost")
    c.Transparent = True
    c.TabStop = False
    c.OnClick = ""
    Ev c, "OnMouseDown", "OnMouseMove", "OnMouseUp"
    Set Glass = c
End Function

' ---------------------------------------------------------------------------
' Page events (pgDesigner's code-behind)
' ---------------------------------------------------------------------------

' Form_Load: make the layout tables if they are missing, and start empty.
Public Sub DesignerLoad(frm As Form)
    DesignerTables
    m_cw = frm("dzGlass").Width \ 15
    m_ch = frm("dzGlass").Height \ 15
    frm("cboDzOpen").RowSource = "SELECT id, layout_name FROM kit_layout ORDER BY layout_name"
    ClearLayout frm
    Status frm, ""
End Sub

' The canvas's MouseDown (phase 1), MouseMove (2) and MouseUp (3). x and y
' are twips from the canvas's corner, as Access passes them.
Public Sub DesignerMouse(frm As Form, ByVal phase As Integer, ByVal btn As Integer, ByVal x As Single, ByVal y As Single)
    Dim p As Long, q As Long, i As Long, nx As Long, ny As Long
    p = CLng(x / 15): q = CLng(y / 15)
    Select Case phase
        Case 1
            If btn <> acLeftButton Then Exit Sub
            If m_mode = MODE_ARMED Then
                Place frm, m_tool, p, q
                Exit Sub
            End If
            m_mode = MODE_IDLE
            If m_sel > 0 Then
                If Abs(p - (m_x(m_sel) + m_w(m_sel))) <= DZ_GRIP And Abs(q - (m_y(m_sel) + m_h(m_sel))) <= DZ_GRIP Then
                    m_mode = MODE_SIZE
                    Exit Sub
                End If
            End If
            i = HitTest(p, q)
            m_sel = i
            If i > 0 Then
                m_mode = MODE_MOVE
                m_offX = p - m_x(i)
                m_offY = q - m_y(i)
            End If
            ShowSelection frm
        Case 2
            If (btn And acLeftButton) = 0 Or m_sel = 0 Then Exit Sub
            i = m_sel
            If m_mode = MODE_MOVE Then
                nx = Snap(p - m_offX): ny = Snap(q - m_offY)
                If nx <> m_x(i) Or ny <> m_y(i) Then SetRect frm, i, nx, ny, m_w(i), m_h(i)
            ElseIf m_mode = MODE_SIZE Then
                nx = Snap(p) - m_x(i): ny = Snap(q) - m_y(i)
                If nx <> m_w(i) Or ny <> m_h(i) Then SetRect frm, i, m_x(i), m_y(i), nx, ny
            End If
        Case 3
            If m_mode = MODE_MOVE Or m_mode = MODE_SIZE Then m_mode = MODE_IDLE
    End Select
End Sub

' A palette tile's MouseDown (phase 1), MouseMove (2) and MouseUp (3); x and
' y are twips from the tile's corner. Released over the canvas, the element
' lands where its outline was shown; released on the tile itself (a click),
' the next click on the canvas places it.
Public Sub DesignerToolMouse(frm As Form, ByVal idx As Integer, ByVal phase As Integer, ByVal btn As Integer, _
        ByVal x As Single, ByVal y As Single)
    Dim t As Control, g As Control, p As Long, q As Long
    Set t = frm("dzTool" & idx)
    Set g = frm("dzGlass")
    p = CLng((t.Left + x - g.Left) / 15)
    q = CLng((t.Top + y - g.Top) / 15)
    Select Case phase
        Case 1
            If btn <> acLeftButton Then Exit Sub
            m_mode = MODE_DRAG
            m_tool = Split(DZ_TOOLS, ";")(idx)
            m_gOn = False
        Case 2
            If m_mode <> MODE_DRAG Or (btn And acLeftButton) = 0 Then Exit Sub
            If p >= 0 And q >= 0 And p < m_cw And q < m_ch Then
                Landing m_tool, p, q, m_gx, m_gy
                m_gOn = True
                Outline frm, m_gx, m_gy, DefaultW(m_tool), DefaultH(m_tool), False
            ElseIf m_gOn Then
                m_gOn = False
                ShowSelection frm
            End If
        Case 3
            If m_mode <> MODE_DRAG Then Exit Sub
            If m_gOn Then
                Place frm, m_tool, p, q
            ElseIf x >= 0 And y >= 0 And x < t.Width And y < t.Height Then
                m_mode = MODE_ARMED
                ShowSelection frm
                Status frm, "Click the canvas to place a " & LCase$(KindCaption(m_tool)) & ". Esc cancels."
            Else
                m_mode = MODE_IDLE
                m_tool = ""
                ShowSelection frm
            End If
    End Select
End Sub

' Form_KeyDown: Esc cancels a palette click or clears the selection; with the
' canvas or a palette tile focused, Delete removes the selected element and the arrow keys
' move it (8px, or 1px with Shift). Returns the key code to pass on (0 when
' the designer used the key).
Public Function DesignerKeyDown(frm As Form, ByVal keyCode As Integer, ByVal shift As Integer) As Integer
    Dim nm As String, stp As Long, i As Long
    DesignerKeyDown = keyCode
    On Error Resume Next
    nm = frm.ActiveControl.Name
    On Error GoTo 0
    If keyCode = vbKeyEscape And m_mode = MODE_ARMED Then
        m_mode = MODE_IDLE
        m_tool = ""
        Status frm, ""
        DesignerKeyDown = 0
        Exit Function
    End If
    If m_sel = 0 Or Not (nm = "dzGlass" Or nm Like "dzTool#") Then Exit Function
    i = m_sel
    stp = IIf((shift And acShiftMask) <> 0, 1, DZ_GRID)
    Select Case keyCode
        Case vbKeyEscape
            m_sel = 0
            ShowSelection frm
        Case vbKeyDelete
            DesignerDelete frm
        Case vbKeyLeft: SetRect frm, i, m_x(i) - stp, m_y(i), m_w(i), m_h(i)
        Case vbKeyRight: SetRect frm, i, m_x(i) + stp, m_y(i), m_w(i), m_h(i)
        Case vbKeyUp: SetRect frm, i, m_x(i), m_y(i) - stp, m_w(i), m_h(i)
        Case vbKeyDown: SetRect frm, i, m_x(i), m_y(i) + stp, m_w(i), m_h(i)
        Case Else: Exit Function
    End Select
    DesignerKeyDown = 0
End Function

' txtDzText's AfterUpdate: the selected element's title, text or percent.
Public Sub DesignerText(frm As Form)
    Dim v As String
    If m_sel = 0 Then Exit Sub
    v = Trim$(Nz(frm("txtDzText").Value, ""))
    If m_kind(m_sel) = "progress" Then
        If Not IsWholePercent(v) Then
            KitMsgBox "Percent done is a whole number from 0 to 100.", vbExclamation, "Form designer"
            frm("txtDzText").Value = m_text(m_sel)
            Exit Sub
        End If
        v = CStr(CLng(v))
    End If
    m_text(m_sel) = Left$(v, 255)
    Render frm, m_sel
    m_dirty = True
End Sub

' txtDzName's AfterUpdate: a new name is a change to save.
Public Sub DesignerRenamed()
    m_dirty = True
End Sub

' btnDzDelete, and Delete on the canvas: remove the selected element.
Public Sub DesignerDelete(frm As Form)
    Dim i As Long, j As Long
    i = m_sel
    If i = 0 Then Exit Sub
    ShowSlot frm, m_kind(i), m_slot(i), False
    For j = i To m_n - 1
        m_kind(j) = m_kind(j + 1): m_slot(j) = m_slot(j + 1)
        m_x(j) = m_x(j + 1): m_y(j) = m_y(j + 1): m_w(j) = m_w(j + 1): m_h(j) = m_h(j + 1)
        m_text(j) = m_text(j + 1)
    Next
    m_n = m_n - 1
    m_sel = 0
    m_dirty = True
    ShowSelection frm
End Sub

' btnDzNew: an empty canvas, after asking to save unsaved changes.
Public Sub DesignerNew(frm As Form)
    If Not OkToDiscard(frm) Then Exit Sub
    ClearLayout frm
End Sub

' cboDzOpen's AfterUpdate: open the layout picked, after asking to save
' unsaved changes (No keeps the list on the open layout).
Public Sub DesignerOpen(frm As Form)
    Dim id As Long
    id = Nz(frm("cboDzOpen").Value, 0)
    If id = m_layoutId Then Exit Sub
    If Not OkToDiscard(frm) Then
        frm("cboDzOpen").Value = IIf(m_layoutId = 0, Null, m_layoutId)
        Exit Sub
    End If
    LoadLayout frm, id
End Sub

' btnDzSave: save the layout under its name. False when it needs a name, the
' name is taken by another layout, or the save failed (the error is raised).
Public Function DesignerSave(frm As Form) As Boolean
    Dim db As DAO.Database, rs As DAO.Recordset, ws As DAO.Workspace, nm As String, id As Long
    Dim i As Long, k As Long, layer As Long, n(0 To 3) As Long, inTrans As Boolean
    nm = Trim$(Nz(frm("txtDzName").Value, ""))
    If Len(nm) = 0 Then
        KitMsgBox "Give the layout a name first.", vbExclamation, "Form designer"
        frm("txtDzName").SetFocus
        Exit Function
    End If
    nm = Left$(nm, 64)
    If DCount("*", "kit_layout", "layout_name=" & SqlText(nm) & " AND id<>" & m_layoutId) > 0 Then
        KitMsgBox "A layout called """ & nm & """ already exists. Give this one another name, or open that one from the list.", _
            vbExclamation, "Form designer"
        frm("txtDzName").SetFocus
        Exit Function
    End If
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    On Error GoTo Fail
    ws.BeginTrans
    inTrans = True
    Set rs = db.OpenRecordset("SELECT * FROM kit_layout WHERE id=" & m_layoutId, dbOpenDynaset)
    If rs.EOF Then rs.AddNew Else rs.Edit
    rs!layout_name = nm
    rs!form_width = m_cw
    rs!form_height = m_ch
    rs!saved_at = Now()
    rs!saved_by = KitUser()
    rs.Update
    rs.Bookmark = rs.LastModified
    id = rs!id
    rs.Close
    db.Execute "DELETE FROM kit_layout_item WHERE layout_id=" & id, dbFailOnError
    ' Back to front, as the pools draw: cards, progress, labels, buttons.
    Set rs = db.OpenRecordset("kit_layout_item", dbOpenDynaset, dbAppendOnly)
    For layer = 0 To 3
        For i = 1 To m_n
            If KindLayer(m_kind(i)) = layer Then
                k = k + 1
                n(layer) = n(layer) + 1
                rs.AddNew
                rs!layout_id = id
                rs!item_order = k
                rs!kind = m_kind(i)
                rs!item_name = KindPrefix(m_kind(i)) & n(layer)
                rs!item_text = m_text(i)
                rs!x = m_x(i): rs!y = m_y(i): rs!w = m_w(i): rs!h = m_h(i)
                rs.Update
            End If
        Next
    Next
    rs.Close
    ws.CommitTrans
    inTrans = False
    On Error GoTo 0
    m_layoutId = id
    m_dirty = False
    frm("txtDzName").Value = nm
    frm("cboDzOpen").Requery
    frm("cboDzOpen").Value = id
    KitToast "Layout """ & nm & """ saved.", "success"
    DesignerSave = True
    Exit Function
Fail:
    Dim errNo As Long, errSrc As String, errDesc As String
    errNo = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    On Error Resume Next
    If inTrans Then ws.Rollback
    On Error GoTo 0
    Err.Raise errNo, errSrc, errDesc
End Function

' The page's CanLeave (ShellCanLeave): ask to save unsaved changes first.
Public Function DesignerCanLeave(frm As Form) As Boolean
    DesignerCanLeave = OkToDiscard(frm)
End Function

' ---------------------------------------------------------------------------
' Layout tables and for tests
' ---------------------------------------------------------------------------

' Make kit_layout (one row per saved layout) and kit_layout_item (its
' elements, px from the canvas corner) when they are missing.
Public Sub DesignerTables()
    Dim db As DAO.Database
    Set db = CurrentDb
    If Not TableExists(db, "kit_layout") Then
        db.Execute "CREATE TABLE kit_layout (id COUNTER PRIMARY KEY, layout_name TEXT(64) NOT NULL, " & _
            "form_width LONG, form_height LONG, saved_at DATETIME, saved_by TEXT(150))", dbFailOnError
        db.Execute "CREATE UNIQUE INDEX ux_kit_layout_name ON kit_layout (layout_name)", dbFailOnError
    End If
    If Not TableExists(db, "kit_layout_item") Then
        db.Execute "CREATE TABLE kit_layout_item (id COUNTER PRIMARY KEY, layout_id LONG NOT NULL, item_order LONG, " & _
            "kind TEXT(20), item_name TEXT(64), item_text TEXT(255), x LONG, y LONG, w LONG, h LONG)", dbFailOnError
        db.Execute "CREATE INDEX ix_kit_layout_item_layout ON kit_layout_item (layout_id)", dbFailOnError
    End If
    db.TableDefs.Refresh
End Sub

' Place a new kind element centred on (x, y) px, as dropping it there does.
' Returns its item number, 0 when that kind's pool is full.
Public Function DesignerPlace(frm As Form, ByVal kind As String, ByVal x As Long, ByVal y As Long) As Long
    DesignerPlace = Place(frm, kind, x, y)
End Function

' Items on the canvas, and item i as "kind x,y wxh text" (smoke tests).
Public Function DesignerCount() As Long
    DesignerCount = m_n
End Function

Public Function DesignerItem(ByVal i As Long) As String
    DesignerItem = m_kind(i) & " " & m_x(i) & "," & m_y(i) & " " & m_w(i) & "x" & m_h(i) & " " & m_text(i)
End Function

' Select item i (0 for none), as clicking it does.
Public Sub DesignerSelect(frm As Form, ByVal i As Long)
    m_sel = i
    ShowSelection frm
End Sub

' ---------------------------------------------------------------------------
' Layout
' ---------------------------------------------------------------------------

' New item of kind centred on (x, y) px, snapped and kept on the canvas, then
' selected. 0 (with a toast) when that kind's pool is full.
Private Function Place(frm As Form, ByVal kind As String, ByVal x As Long, ByVal y As Long) As Long
    Dim i As Long, lx As Long, ly As Long
    m_mode = MODE_IDLE
    m_tool = ""
    Status frm, ""
    Landing kind, x, y, lx, ly
    i = AddItem(kind, lx, ly, DefaultW(kind), DefaultH(kind), DefaultText(kind))
    If i = 0 Then
        KitToast "The designer holds up to " & PoolSize(kind) & " " & LCase$(KindCaption(kind)) & " elements.", "warning"
        ShowSelection frm
        Exit Function
    End If
    Render frm, i
    m_sel = i
    m_dirty = True
    ShowSelection frm
    Place = i
End Function

' Where a kind element centred on (x, y) px lands: snapped, on the canvas.
Private Sub Landing(ByVal kind As String, ByVal x As Long, ByVal y As Long, ByRef lx As Long, ByRef ly As Long)
    lx = Clamp(Snap(x - DefaultW(kind) \ 2), 0, m_cw - DefaultW(kind))
    ly = Clamp(Snap(y - DefaultH(kind) \ 2), 0, m_ch - DefaultH(kind))
End Sub

' Append an item in the first free slot of its pool; 0 when the pool is full.
' The rectangle is kept to the kind's smallest size and inside the canvas.
Private Function AddItem(ByVal kind As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
        ByVal txt As String) As Long
    Dim s As Long, i As Long, used As Boolean
    For s = 1 To PoolSize(kind)
        used = False
        For i = 1 To m_n
            If m_kind(i) = kind And m_slot(i) = s Then used = True: Exit For
        Next
        If Not used Then Exit For
    Next
    If s > PoolSize(kind) Or m_n >= DZ_MAX Then Exit Function
    m_n = m_n + 1
    m_kind(m_n) = kind
    m_slot(m_n) = s
    m_text(m_n) = txt
    FitRect m_n, x, y, w, h
    AddItem = m_n
End Function

' Move or resize item i, keeping it on the canvas, and redraw it.
Private Sub SetRect(frm As Form, ByVal i As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    FitRect i, x, y, w, h
    Render frm, i
    m_dirty = True
    ShowSelection frm
End Sub

' Item i's rectangle: no smaller than its kind allows, no larger than the
' canvas, and inside it.
Private Sub FitRect(ByVal i As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    w = Clamp(w, MinW(m_kind(i)), m_cw)
    h = Clamp(h, MinH(m_kind(i)), MaxH(m_kind(i)))
    m_w(i) = w
    m_h(i) = h
    m_x(i) = Clamp(x, 0, m_cw - w)
    m_y(i) = Clamp(y, 0, m_ch - h)
End Sub

' The topmost item under (x, y) px: the pools draw cards, progress, labels,
' then buttons, and within a pool the later slot on top. 0 for none.
Private Function HitTest(ByVal x As Long, ByVal y As Long) As Long
    Dim i As Long, rank As Long, best As Long
    best = -1
    For i = 1 To m_n
        If x >= m_x(i) And x < m_x(i) + m_w(i) And y >= m_y(i) And y < m_y(i) + m_h(i) Then
            rank = KindLayer(m_kind(i)) * 100 + m_slot(i)
            If rank > best Then best = rank: HitTest = i
        End If
    Next
End Function

' Hide every element and start a new, unsaved layout.
Private Sub ClearLayout(frm As Form)
    Dim i As Long
    For i = 1 To m_n
        ShowSlot frm, m_kind(i), m_slot(i), False
    Next
    m_n = 0
    m_sel = 0
    m_mode = MODE_IDLE
    m_tool = ""
    m_layoutId = 0
    m_dirty = False
    frm("cboDzOpen").Value = Null
    frm("txtDzName").Value = Null
    ShowSelection frm
End Sub

' Show the saved layout id (elements past a pool's size are left out, with a toast).
Private Sub LoadLayout(frm As Form, ByVal id As Long)
    Dim db As DAO.Database, rs As DAO.Recordset, nm As String, skipped As Long, i As Long
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT layout_name FROM kit_layout WHERE id=" & id, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        ClearLayout frm
        Exit Sub
    End If
    nm = rs!layout_name
    rs.Close
    ClearLayout frm
    Set rs = db.OpenRecordset("SELECT kind, item_text, x, y, w, h FROM kit_layout_item WHERE layout_id=" & id & _
        " ORDER BY item_order", dbOpenSnapshot)
    Do Until rs.EOF
        If PoolSize(Nz(rs!kind, "")) = 0 Then
            skipped = skipped + 1
        Else
            i = AddItem(rs!kind, Nz(rs!x, 0), Nz(rs!y, 0), Nz(rs!w, 0), Nz(rs!h, 0), Nz(rs!item_text, ""))
            If i = 0 Then skipped = skipped + 1 Else Render frm, i
        End If
        rs.MoveNext
    Loop
    rs.Close
    m_layoutId = id
    frm("cboDzOpen").Value = id
    frm("txtDzName").Value = nm
    ShowSelection frm
    If skipped > 0 Then KitToast skipped & " element(s) didn't fit the designer and were left out.", "warning"
End Sub

' True when there is nothing unsaved, or the user saved it or chose to drop it.
Private Function OkToDiscard(frm As Form) As Boolean
    If Not m_dirty Then OkToDiscard = True: Exit Function
    Select Case KitMsgBox("Save the changes to this layout first?", vbYesNoCancel + vbQuestion, "Form designer")
        Case vbYes: OkToDiscard = DesignerSave(frm)
        Case vbNo: OkToDiscard = True
    End Select
End Function

' ---------------------------------------------------------------------------
' Drawing
' ---------------------------------------------------------------------------

' Put item i's preview controls at its rectangle, with its text.
Private Sub Render(frm As Form, ByVal i As Long)
    Dim l As Long, t As Long, w As Long, h As Long, nm As String, fw As Long
    l = frm("dzGlass").Left + px(m_x(i))
    t = frm("dzGlass").Top + px(m_y(i))
    w = px(m_w(i)): h = px(m_h(i))
    nm = SlotName(m_kind(i), m_slot(i))
    Select Case m_kind(i)
        Case "card"
            MoveCtl frm(nm), l, t, w, h
            MoveCtl frm(nm & "_t"), l + px(14), t + px(10), w - px(28), px(20)
            frm(nm & "_t").Caption = m_text(i)
            frm(nm).Visible = True
            frm(nm & "_t").Visible = (m_h(i) >= 40)
        Case "progress"
            MoveCtl frm(nm), l, t, w, h
            frm(nm).Visible = True
            fw = (w * Val(m_text(i))) \ 100
            If fw >= px(1) Then MoveCtl frm(nm & "_f"), l, t, fw, h
            frm(nm & "_f").Visible = (fw >= px(1))
        Case "label", "button"
            MoveCtl frm(nm), l, t, w, h
            frm(nm).Caption = m_text(i)
            frm(nm).TopMargin = 0
            VCenter frm(nm), h, IIf(m_kind(i) = "button", 9, 10), (m_kind(i) = "button")
            frm(nm).Visible = True
    End Select
End Sub

' Set a control's rectangle without it ever reaching past both the old and
' the new one: a control moved past the form's edge widens the form for good.
Private Sub MoveCtl(c As Control, ByVal l As Long, ByVal t As Long, ByVal w As Long, ByVal h As Long)
    If w < c.Width Then c.Width = w
    If h < c.Height Then c.Height = h
    c.Left = l
    c.Top = t
    c.Width = w
    c.Height = h
End Sub

Private Sub ShowSlot(frm As Form, ByVal kind As String, ByVal s As Long, ByVal v As Boolean)
    Dim nm As String
    nm = SlotName(kind, s)
    frm(nm).Visible = v
    Select Case kind
        Case "card": frm(nm & "_t").Visible = v
        Case "progress": frm(nm & "_f").Visible = v
    End Select
End Sub

' The outline and resize handle on the selected item (none when nothing is
' selected), the Selected card, and the canvas hint when it is empty.
Private Sub ShowSelection(frm As Form)
    Dim i As Long
    i = m_sel
    If i > 0 Then
        Outline frm, m_x(i), m_y(i), m_w(i), m_h(i), True
    Else
        frm("dzSel").Visible = False
        frm("dzGrip").Visible = False
    End If
    frm("dzHint").Visible = (m_n = 0)
    SetEnabled frm("txtDzText"), i > 0
    SetEnabled frm("btnDzDelete"), i > 0
    If i = 0 Then
        frm("lblDzKind").Caption = "Nothing selected"
        frm("lblDzTextCap").Caption = "TEXT"
        frm("txtDzText").Value = Null
        frm("lblDzPos").Caption = ""
    Else
        frm("lblDzKind").Caption = KindCaption(m_kind(i))
        frm("lblDzTextCap").Caption = Choose(KindLayer(m_kind(i)) + 1, "TITLE", "PERCENT DONE (0-100)", "TEXT", "CAPTION")
        frm("txtDzText").Value = m_text(i)
        frm("lblDzPos").Caption = "Left " & m_x(i) & ", top " & m_y(i) & vbCrLf & m_w(i) & " x " & m_h(i) & " px"
    End If
End Sub

' The selection outline 2px outside (x, y, w, h) px, with the resize handle
' on its bottom-right corner when withGrip.
Private Sub Outline(frm As Form, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal withGrip As Boolean)
    Dim l As Long, t As Long
    l = frm("dzGlass").Left: t = frm("dzGlass").Top
    MoveCtl frm("dzSel"), l + px(x - 2), t + px(y - 2), px(w + 4), px(h + 4)
    frm("dzSel").Visible = True
    If withGrip Then
        MoveCtl frm("dzGrip"), l + px(x + w - DZ_GRIP \ 2), t + px(y + h - DZ_GRIP \ 2), px(DZ_GRIP), px(DZ_GRIP)
    End If
    frm("dzGrip").Visible = withGrip
End Sub

' The page subtitle: what to do next, or the default how-to.
Private Sub Status(frm As Form, ByVal msg As String)
    If Len(msg) = 0 Then msg = "Drag elements onto the canvas. Drag to move, drag the corner to resize; everything snaps to 8px."
    frm("lblSubtitle").Caption = msg
End Sub

' ---------------------------------------------------------------------------
' Element kinds
' ---------------------------------------------------------------------------

' Back-to-front order of the pools: 0 card, 1 progress, 2 label, 3 button.
Private Function KindLayer(ByVal kind As String) As Long
    Select Case kind
        Case "card": KindLayer = 0
        Case "progress": KindLayer = 1
        Case "label": KindLayer = 2
        Case Else: KindLayer = 3
    End Select
End Function

Private Function PoolSize(ByVal kind As String) As Long
    Select Case kind
        Case "card": PoolSize = DZ_CARDS
        Case "progress": PoolSize = DZ_PROGS
        Case "label": PoolSize = DZ_LBLS
        Case "button": PoolSize = DZ_BTNS
    End Select
End Function

Private Function KindCaption(ByVal kind As String) As String
    KindCaption = Choose(KindLayer(kind) + 1, "Card", "Progress", "Label", "Button")
End Function

' Name prefix of a kind's saved elements (card1, prog1, lbl1, btn1).
Private Function KindPrefix(ByVal kind As String) As String
    KindPrefix = Choose(KindLayer(kind) + 1, "card", "prog", "lbl", "btn")
End Function

' The first preview control of slot s in a kind's pool.
Private Function SlotName(ByVal kind As String, ByVal s As Long) As String
    SlotName = Choose(KindLayer(kind) + 1, "dzCard", "dzProg", "dzLbl", "dzBtn") & s
End Function

' Sizes (px) a new element gets, the smallest it can be, and the tallest.
Private Function DefaultW(ByVal kind As String) As Long
    DefaultW = Choose(KindLayer(kind) + 1, 240, 160, 120, 96)
End Function

Private Function DefaultH(ByVal kind As String) As Long
    DefaultH = Choose(KindLayer(kind) + 1, 136, 8, 24, 32)
End Function

Private Function MinW(ByVal kind As String) As Long
    MinW = Choose(KindLayer(kind) + 1, 80, 24, 24, 48)
End Function

Private Function MinH(ByVal kind As String) As Long
    MinH = Choose(KindLayer(kind) + 1, 48, 4, 16, 24)
End Function

Private Function MaxH(ByVal kind As String) As Long
    MaxH = Choose(KindLayer(kind) + 1, m_ch, 16, m_ch, 48)
End Function

Private Function DefaultText(ByVal kind As String) As String
    DefaultText = Choose(KindLayer(kind) + 1, "Card title", "40", "Label text", "Button")
End Function

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

' Nearest multiple of DZ_GRID (negative values too).
Private Function Snap(ByVal v As Long) As Long
    Snap = CLng(Int((v + DZ_GRID / 2) / DZ_GRID)) * DZ_GRID
End Function

Private Function Clamp(ByVal v As Long, ByVal lo As Long, ByVal hi As Long) As Long
    If hi < lo Then hi = lo
    If v < lo Then v = lo
    If v > hi Then v = hi
    Clamp = v
End Function

' "0" to "100", digits only.
Private Function IsWholePercent(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Or Len(s) > 3 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next
    IsWholePercent = (CLng(s) <= 100)
End Function

Private Function SqlText(ByVal s As String) As String
    SqlText = "'" & Replace(s, "'", "''") & "'"
End Function

Private Function TableExists(db As DAO.Database, ByVal nm As String) As Boolean
    Dim td As DAO.TableDef
    For Each td In db.TableDefs
        If td.Name = nm Then TableExists = True: Exit Function
    Next
End Function
