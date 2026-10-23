Option Compare Database
Option Explicit

' Access UI kit - the form designer (prototype): a page where kit elements
' are dragged from a palette onto a canvas, moved, resized, given their
' settings, and saved as a layout. Layouts are rows in kit_layout and
' kit_layout_item (made, or brought up to date, when the page opens).
' Generate builds a real form from the saved layout with the kit's builders
' (Card, Btn, Segmented, Stepper, StatTile and the rest), shows it
' (dzPreview), and writes the same builder calls as a Build_<form>
' procedure, with a code-behind stub, to the designer folder beside the
' database, for an app's build. Generating needs full Access and the
' editable .accdb: a locked .accde can't create forms.
'
' Access can't create controls on a form that is open in Form view, so the
' page is built with a fixed pool of hidden preview controls (DZ_CARDS
' cards, DZ_PROGS progress bars, DZ_GENS outlined boxes for every other kind
' of element, DZ_LBLS labels, DZ_BTNS buttons), and placing an element shows
' the next free one of its pool. The pools are built back to front in that
' order, so content always draws over the cards it sits on. Elements other
' than cards, buttons, labels and progress bars show on the canvas as an
' outlined box naming what they will be; Generate shows them for real. A
' transparent button over the canvas (dzGlass) takes every mouse event and
' finds the element under the pointer; one over the palette (dzPal) finds
' the tile and drags a new element in. A pressed button keeps getting
' MouseMove until the mouse is released, wherever the pointer goes.
'
' Build the page with BuildDesigner (it attaches kit\forms\pgDesigner.vba
' and builds the settings dialog dlgDzProps, kit\forms\dlgDzProps.vba);
' their code-behind hands every event to the Designer* routines here.
' Positions and sizes are in px from the canvas's top-left corner, on an
' 8px grid (DZ_GRID). Double-click an element, or Edit, for its settings;
' Esc cancels a palette click; Delete removes the selected element; the
' arrow keys move it 8px (1px with Shift).

Private Const DZ_GRID As Long = 8               ' snap, px
Private Const DZ_GRIP As Long = 8               ' resize handle, px
Private Const DZ_CARDS As Long = 6
Private Const DZ_PROGS As Long = 8
Private Const DZ_GENS As Long = 24
Private Const DZ_LBLS As Long = 16
Private Const DZ_BTNS As Long = 12
Private Const DZ_MAX As Long = 66               ' all five pools
' Palette order: tile i (two columns, row by row) places kind i.
Private Const DZ_KINDS As String = "card;button;label;progress;text;date;secret;toggle;segmented;slider;" & _
    "rating;stepper;search;stat;alert;avatar;empty"
Private Const DZ_TILE_W As Long = 80            ' palette tile, px
Private Const DZ_TILE_H As Long = 26
Private Const DZ_TILE_STEP As Long = 30         ' tile to tile, down
Private Const DZ_PREVIEW As String = "dzPreview"   ' the form Generate builds to show a layout
Private Const DZ_PROPS As String = "dlgDzProps"    ' the settings dialog
Private Const DZ_EMPTY_GLYPH As String = "E7C1"    ' an empty state's icon

Private Const MODE_IDLE As Integer = 0
Private Const MODE_MOVE As Integer = 1
Private Const MODE_SIZE As Integer = 2
Private Const MODE_DRAG As Integer = 3          ' dragging a new element in from the palette
Private Const MODE_ARMED As Integer = 4         ' palette element clicked; the next canvas click places it

' The layout being edited: items 1..m_n, each in slot m_slot of its kind's
' pool. m_text, m_text2 and m_text3 are the settings in the dialog's first
' three rows, m_choice the fourth (a pick from a list).
Private m_kind(1 To DZ_MAX) As String
Private m_slot(1 To DZ_MAX) As Long
Private m_x(1 To DZ_MAX) As Long
Private m_y(1 To DZ_MAX) As Long
Private m_w(1 To DZ_MAX) As Long
Private m_h(1 To DZ_MAX) As Long
Private m_text(1 To DZ_MAX) As String
Private m_text2(1 To DZ_MAX) As String
Private m_text3(1 To DZ_MAX) As String
Private m_choice(1 To DZ_MAX) As String
Private m_n As Long
Private m_sel As Long                           ' selected item, 0 for none
Private m_mode As Integer
Private m_tool As String                        ' kind being dragged in or armed
Private m_toolIdx As Long                       ' its palette tile
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
' Elements (the palette) and Selected (the selected element, Edit and
' Delete) on the left, the layout list, name, New, Generate and Save above
' the canvas on the right. Saved as formName with pgDesigner.vba's code.
' Builds its settings dialog dlgDzProps first.
Public Sub BuildDesigner(ByVal formName As String, ByVal x As Long, ByVal w As Long, ByVal h As Long)
    Dim c As Control, i As Long, ty As Long, py As Long, lw As Long, cx As Long, cy As Long, cw As Long, ch As Long
    Dim kinds() As String, rows As Long
    BuildDesignerDialog
    NewForm "", 0, w, h
    F.KeyPreview = True                          ' Delete, Esc and the arrow keys (Form_KeyDown)
    Ev F, "OnLoad", "OnKeyDown"
    PageHeader "Form designer", "Drag an element from Elements onto the canvas, or click one and then click the canvas. " & _
        "Drag an element to move it, and its corner handle to resize it. Everything snaps to an 8px grid. " & _
        "Double-click an element for its settings. Save keeps the layout under its name; Generate builds it as a real form.", x
    lw = px(200)
    cx = x + lw + px(16)
    cw = w - x - cx
    cy = px(136)
    ch = h - cy - px(16)

    ' Palette: a tile per kind in two columns, with one transparent button over
    ' them all for the mouse.
    kinds = Split(DZ_KINDS, ";")
    rows = (UBound(kinds) + 2) \ 2
    ty = px(76)
    Card "DzTools", "Elements", x, ty, lw, px(44 + rows * DZ_TILE_STEP + 12)
    For i = 0 To UBound(kinds)
        Box "dzTile" & i, x + px(16 + (i Mod 2) * (DZ_TILE_W + 8)), ty + px(44 + (i \ 2) * DZ_TILE_STEP), _
            px(DZ_TILE_W), px(DZ_TILE_H), CLR_BG, CLR_BORDER
        Set c = Lbl("dzTile" & i & "_n", KindCaption(kinds(i)), x + px(16 + (i Mod 2) * (DZ_TILE_W + 8)), _
            ty + px(44 + (i \ 2) * DZ_TILE_STEP), px(DZ_TILE_W), px(DZ_TILE_H), 9)
        c.TextAlign = 2
        VCenter c, px(DZ_TILE_H), 9, True
    Next
    Set c = Glass("dzPal", x + px(16), ty + px(44), px(2 * DZ_TILE_W + 8), px(rows * DZ_TILE_STEP))
    c.ControlTipText = "Drag onto the canvas, or click and then click the canvas"

    ' The selected element.
    py = ty + px(44 + rows * DZ_TILE_STEP + 12) + px(12)
    Card "DzProps", "Selected", x, py, lw, h - py - px(16)
    Lbl "lblDzKind", "Nothing selected", x + px(16), py + px(40), lw - px(32), px(20), 9, True
    Lbl "lblDzPos", "", x + px(16), py + px(62), lw - px(32), px(20), 9, False, CLR_MUTED_FG
    Set c = Btn("btnDzEdit", "Edit" & ChrW(8230), x + px(16), py + px(90), px(84), px(28))
    c.Enabled = False
    Set c = Btn("btnDzDelete", "Delete", x + px(108), py + px(90), px(76), px(28), "outline")
    c.Enabled = False

    ' Layout list, name, New, Generate and Save.
    Lbl "lblDzOpen", "OPEN A LAYOUT", cx, px(76), px(220), px(16), 8, True, CLR_MUTED_FG
    Set c = Cbo("cboDzOpen", "", "", cx, px(94), px(220), px(30), 2, "0;" & px(220))
    Ev c, "AfterUpdate"
    Lbl "lblDzName", "LAYOUT NAME", cx + px(232), px(76), px(220), px(16), 8, True, CLR_MUTED_FG
    Set c = Txt("txtDzName", "", cx + px(232), px(94), px(220), px(30))
    Ev c, "AfterUpdate"
    Btn "btnDzNew", "New", cx + cw - px(268), px(95), px(84), px(28), "outline"
    Btn "btnDzGenerate", "Generate", cx + cw - px(176), px(95), px(84), px(28), "outline"
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
    For i = 1 To DZ_GENS
        Set c = Box("dzGen" & i, cx, cy, px(160), px(32), CLR_SURFACE, CLR_INPUT_BORDER): c.Visible = False
        Set c = Lbl("dzGen" & i & "_t", "", cx, cy, px(160), px(32), 9, False, CLR_MUTED_FG)
        c.TextAlign = 2
        c.Visible = False
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
    Set c = Glass("dzGlass", cx, cy, cw, ch)
    Ev c, "OnDblClick"
    SaveForm formName, "pgDesigner"
End Sub

' The settings dialog dlgDzProps: a title and four rows (three text boxes
' and a list), each with its caption; dlgDzProps.vba shows the rows the
' selected kind uses (DesignerPropsLoad) and checks them on OK
' (DesignerPropsSave).
Public Sub BuildDesignerDialog()
    Dim c As Control, r As Long, y As Long
    NewDialog "Element settings", px(400), px(56 + 4 * 64 + 48 + 30 + 20)
    Ev F, "OnLoad"
    Lbl "lblPTitle", "", px(20), px(18), px(360), px(26), 12, True
    For r = 1 To 4
        y = px(56 + (r - 1) * 64)
        Lbl "lblP" & r, "", px(20), y, px(360), px(16), 8, True, CLR_MUTED_FG
        If r < 4 Then
            Txt "txtP" & r, "", px(20), y + px(18), px(360), px(30)
        Else
            Set c = Cbo("cboP4", "", "", px(20), y + px(18), px(360), px(30))
            c.RowSourceType = "Value List"
        End If
    Next
    Lbl "lblPHint", "", px(20), px(56 + 4 * 64), px(360), px(40), 9, False, CLR_MUTED_FG
    y = px(56 + 4 * 64 + 48)
    Set c = Btn("btnOK", "OK", px(400) - px(20) - 2 * px(88) - px(8), y, px(88), px(30))
    c.Default = True
    Set c = Btn("btnCancel", "Cancel", px(400) - px(20) - px(88), y, px(88), px(30), "outline")
    c.Cancel = True
    SaveForm DZ_PROPS
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

' Form_Load: make or update the layout tables, and start empty.
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

' The palette's MouseDown (phase 1), MouseMove (2) and MouseUp (3); x and y
' are twips from the palette's corner. Pressed on a tile and released over
' the canvas, the element lands where its outline was shown; released on the
' same tile (a click), the next click on the canvas places it.
Public Sub DesignerToolMouse(frm As Form, ByVal phase As Integer, ByVal btn As Integer, ByVal x As Single, ByVal y As Single)
    Dim t As Control, g As Control, p As Long, q As Long
    Set t = frm("dzPal")
    Set g = frm("dzGlass")
    p = CLng((t.Left + x - g.Left) / 15)
    q = CLng((t.Top + y - g.Top) / 15)
    Select Case phase
        Case 1
            m_mode = MODE_IDLE
            If btn <> acLeftButton Then Exit Sub
            m_toolIdx = TileAt(x / 15, y / 15)
            If m_toolIdx < 0 Then Exit Sub
            m_mode = MODE_DRAG
            m_tool = Split(DZ_KINDS, ";")(m_toolIdx)
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
            ElseIf TileAt(x / 15, y / 15) = m_toolIdx Then
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
' canvas or the palette focused, Delete removes the selected element and the
' arrow keys move it (8px, or 1px with Shift). Returns the key code to pass
' on (0 when the designer used the key).
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
    If m_sel = 0 Or Not (nm = "dzGlass" Or nm = "dzPal") Then Exit Function
    i = m_sel
    stp = IIf((shift And acShiftMask) <> 0, 1, DZ_GRID)
    Select Case keyCode
        Case vbKeyEscape
            m_sel = 0
            ShowSelection frm
        Case vbKeyDelete
            DesignerDelete frm
        Case vbKeyReturn
            DesignerEdit frm
        Case vbKeyLeft: SetRect frm, i, m_x(i) - stp, m_y(i), m_w(i), m_h(i)
        Case vbKeyRight: SetRect frm, i, m_x(i) + stp, m_y(i), m_w(i), m_h(i)
        Case vbKeyUp: SetRect frm, i, m_x(i), m_y(i) - stp, m_w(i), m_h(i)
        Case vbKeyDown: SetRect frm, i, m_x(i), m_y(i) + stp, m_w(i), m_h(i)
        Case Else: Exit Function
    End Select
    DesignerKeyDown = 0
End Function

' btnDzEdit, a double-click on an element, and Enter on the canvas: the
' selected element's settings (dlgDzProps), redrawn when they change.
Public Sub DesignerEdit(frm As Form)
    If m_sel = 0 Then Exit Sub
    If Len(RowCaption(m_kind(m_sel), 1)) = 0 And Len(RowCaption(m_kind(m_sel), 4)) = 0 Then
        KitToast "A " & LCase$(KindCaption(m_kind(m_sel))) & " has no settings yet: size it on the canvas.", "info"
        Exit Sub
    End If
    If RunDialog(DZ_PROPS) Then
        Render frm, m_sel
        ShowSelection frm
    End If
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
        m_text(j) = m_text(j + 1): m_text2(j) = m_text2(j + 1): m_text3(j) = m_text3(j + 1): m_choice(j) = m_choice(j + 1)
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
    Dim i As Long, k As Long, layer As Long, n(0 To 4) As Long, inTrans As Boolean
    Dim kinds() As String, j As Long, cnt() As Long
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
    kinds = Split(DZ_KINDS, ";")
    ReDim cnt(0 To UBound(kinds))
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
    ' Back to front, as the pools draw; each element named by its kind and count (card1, seg2).
    Set rs = db.OpenRecordset("kit_layout_item", dbOpenDynaset, dbAppendOnly)
    For layer = 0 To 4
        For i = 1 To m_n
            If KindLayer(m_kind(i)) = layer Then
                k = k + 1
                For j = 0 To UBound(kinds)
                    If kinds(j) = m_kind(i) Then cnt(j) = cnt(j) + 1: Exit For
                Next
                rs.AddNew
                rs!layout_id = id
                rs!item_order = k
                rs!kind = m_kind(i)
                rs!item_name = KindPrefix(m_kind(i)) & cnt(j)
                rs!item_text = m_text(i)
                rs!item_text2 = m_text2(i)
                rs!item_text3 = m_text3(i)
                rs!item_choice = m_choice(i)
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
' Settings dialog (dlgDzProps's code-behind)
' ---------------------------------------------------------------------------

' Form_Load: the selected element's kind as the title, the rows it uses
' (moved up under each other, the dialog shrunk to fit) with its settings,
' and what the rows take.
Public Sub DesignerPropsLoad(dlg As Form)
    Dim i As Long, r As Long, y As Long, cap As String, inp As Control, dy As Long
    i = m_sel
    If i = 0 Then Exit Sub
    dlg("lblPTitle").Caption = KindCaption(m_kind(i)) & " settings"
    y = dlg("lblP1").Top
    For r = 1 To 4
        cap = RowCaption(m_kind(i), r)
        If r < 4 Then Set inp = dlg("txtP" & r) Else Set inp = dlg("cboP4")
        SetVisible dlg("lblP" & r), Len(cap) > 0
        SetVisible inp, Len(cap) > 0
        If Len(cap) > 0 Then
            dlg("lblP" & r).Caption = cap
            dlg("lblP" & r).Top = y
            inp.Top = y + px(18)
            y = y + px(64)
        End If
    Next
    dlg("txtP1").Value = m_text(i)
    dlg("txtP2").Value = m_text2(i)
    dlg("txtP3").Value = m_text3(i)
    dlg("cboP4").RowSource = ChoiceList(m_kind(i))
    dlg("cboP4").Value = IIf(Len(m_choice(i)) = 0, Null, m_choice(i))
    dlg("lblPHint").Caption = KindHint(m_kind(i))
    dlg("lblPHint").Top = y
    dy = dlg("btnOK").Top - (y + px(48))
    If dy > 0 Then
        dlg("btnOK").Top = dlg("btnOK").Top - dy
        dlg("btnCancel").Top = dlg("btnOK").Top
        dlg.Section(acDetail).Height = dlg.Section(acDetail).Height - dy
        dlg.InsideHeight = dlg.Section(acDetail).Height
    End If
    CenterOnApp dlg
    If Len(RowCaption(m_kind(i), 1)) > 0 Then dlg("txtP1").SetFocus Else dlg("cboP4").SetFocus
End Sub

' OK: apply the dialog's settings to the selected element (DesignerSetProps).
' False, after saying what is wrong, when they don't fit its kind.
Public Function DesignerPropsSave(dlg As Form) As Boolean
    Dim msg As String
    If m_sel = 0 Then DesignerPropsSave = True: Exit Function
    msg = DesignerSetProps(m_sel, Nz(dlg("txtP1").Value, ""), Nz(dlg("txtP2").Value, ""), Nz(dlg("txtP3").Value, ""), _
        Nz(dlg("cboP4").Value, ""))
    If Len(msg) > 0 Then
        KitMsgBox msg, vbExclamation, "Form designer"
        Exit Function
    End If
    DesignerPropsSave = True
End Function

' Give item i these settings (the dialog's four rows), then keep it inside
' the sizes they allow (a segmented control's options, a rating's stars).
' Returns what is wrong with them instead, changing nothing; "" when applied.
' The caller redraws the item (DesignerRedraw).
Public Function DesignerSetProps(ByVal i As Long, ByVal t1 As String, ByVal t2 As String, ByVal t3 As String, _
        ByVal choice As String) As String
    Dim msg As String
    t1 = Left$(Trim$(t1), 255): t2 = Left$(Trim$(t2), 255): t3 = Left$(Trim$(t3), 255): choice = Trim$(choice)
    msg = PropsProblem(m_kind(i), t1, t2, t3, choice)
    If Len(msg) > 0 Then DesignerSetProps = msg: Exit Function
    m_text(i) = t1: m_text2(i) = t2: m_text3(i) = t3: m_choice(i) = choice
    FitRect i, m_x(i), m_y(i), m_w(i), m_h(i)
    m_dirty = True
End Function

' ---------------------------------------------------------------------------
' Generate
' ---------------------------------------------------------------------------

' btnDzGenerate: save the layout if it has changes, build it as a real form
' and show it (dzPreview), and write its build procedure and code-behind stub
' to <database folder>\designer as Build_<form>.txt and <form>.vba.
Public Sub DesignerGenerate(frm As Form)
    Dim nm As String, folder As String, id As Long
    If m_n = 0 Then
        KitMsgBox "Place at least one element on the canvas first.", vbExclamation, "Form designer"
        Exit Sub
    End If
    If Not CanBuildForms() Then
        KitMsgBox "Generate builds forms, which a locked (.accde) copy or the Access Runtime can't do. " & _
            "Open the editable .accdb in full Access to generate.", vbExclamation, "Form designer"
        Exit Sub
    End If
    If m_dirty Or m_layoutId = 0 Then
        If Not DesignerSave(frm) Then Exit Sub
    End If
    id = m_layoutId
    nm = DesignerFormName(Nz(DLookup("layout_name", "kit_layout", "id=" & id), ""))
    folder = CurrentProject.Path & "\designer"
    EnsureFolder folder
    WriteTextFile folder & "\Build_" & nm & ".txt", DesignerCode(id)
    WriteTextFile folder & "\" & nm & ".vba", DesignerStub(id)
    DesignerPreview id
    KitToast "Preview built. Its code is in the designer folder: Build_" & nm & ".txt and " & nm & ".vba.", "success", 6000
End Sub

' The form a layout generates: "pg" and the name's letters and digits, each
' word capitalised ("review navy 2" gives pgReviewNavy2; no letters, pgLayout).
Public Function DesignerFormName(ByVal layoutName As String) As String
    Dim i As Long, ch As Long, out As String, up As Boolean
    up = True
    For i = 1 To Len(layoutName)
        ch = AscW(Mid$(layoutName, i, 1))
        If (ch >= 48 And ch <= 57) Or (ch >= 65 And ch <= 90) Or (ch >= 97 And ch <= 122) Then
            If up Then out = out & UCase$(ChrW(ch)) Else out = out & ChrW(ch)
            up = False
        Else
            up = True
        End If
    Next
    If Len(out) = 0 Then out = "Layout"
    DesignerFormName = Left$("pg" & out, 64)
End Function

' Layout id as a Build_<form> procedure for an app's build module: NewForm
' at the canvas's size, one kit builder call per element back to front,
' SaveForm. The elements that call the form's code name a handler after
' themselves (seg1Changed, step1Step, search1Search: see DesignerStub).
Public Function DesignerCode(ByVal id As Long) As String
    Dim db As DAO.Database, rs As DAO.Recordset, nm As String, ln As String, s As String, title As String
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT layout_name, form_width, form_height FROM kit_layout WHERE id=" & id, dbOpenSnapshot)
    If rs.EOF Then Err.Raise 5, , "No form designer layout " & id
    title = OneLine(rs!layout_name)
    nm = DesignerFormName(title)
    s = "' " & nm & ", built from the form designer layout """ & title & """ (" & Format$(Now(), "yyyy-mm-dd hh:nn") & ")." & vbCrLf & _
        "' Call it from the app's build, and keep " & nm & ".vba, its code-behind, in src\forms." & vbCrLf & _
        "Private Sub Build_" & nm & "()" & vbCrLf & _
        "    NewForm """", 0, px(" & Nz(rs!form_width, 0) & "), px(" & Nz(rs!form_height, 0) & ")" & vbCrLf
    rs.Close
    Set rs = db.OpenRecordset(ItemsSql(id), dbOpenSnapshot)
    Do Until rs.EOF
        ln = ItemCall(rs)
        If Len(ln) > 0 Then s = s & "    " & ln & vbCrLf
        rs.MoveNext
    Loop
    rs.Close
    DesignerCode = s & "    SaveForm """ & nm & """" & vbCrLf & "End Sub" & vbCrLf
End Function

' Layout id's code-behind to start from: an empty handler, with the kit's
' error handling, for each button's click, each empty state's button and
' each segmented control, toggle, slider, rating, stepper and search box
' that DesignerCode names, for the app to fill in.
Public Function DesignerStub(ByVal id As Long) As String
    Dim db As DAO.Database, rs As DAO.Recordset, nm As String, s As String, title As String, b As String
    Dim sig As String, proc As String, what As String
    Set db = CurrentDb
    title = OneLine(Nz(DLookup("layout_name", "kit_layout", "id=" & id), ""))
    nm = DesignerFormName(title)
    s = "Option Compare Database" & vbCrLf & "Option Explicit" & vbCrLf & vbCrLf & _
        "' " & nm & " (Build_" & nm & ", from the form designer layout """ & title & """)." & vbCrLf
    Set rs = db.OpenRecordset(ItemsSql(id), dbOpenSnapshot)
    Do Until rs.EOF
        b = ItemName(rs)
        sig = ""
        Select Case Nz(rs!kind, "")
            Case "button"
                sig = "Private Sub " & b & "_Click()": proc = b & "_Click"
                what = "What " & OneLine(Nz(rs!item_text, "")) & " does goes here."
            Case "empty"
                If Len(Nz(rs!item_text3, "")) > 0 Then
                    sig = "Private Sub " & b & "_b_Click()": proc = b & "_b_Click"
                    what = "What " & OneLine(rs!item_text3) & " does goes here."
                End If
            Case "toggle", "segmented", "slider", "rating"
                sig = "Public Sub " & b & "Changed()": proc = b & "Changed"
                what = "The user changed " & b & "."
            Case "stepper"
                sig = "Public Sub " & b & "Step(ByVal stepNo As Long)": proc = b & "Step"
                what = "The user clicked done step stepNo, to go back to it (SetStep)."
            Case "search"
                sig = "Public Sub " & b & "Search(ByVal text As String)": proc = b & "Search"
                what = "Filter the page by text, as the user types (and """" when cleared)."
        End Select
        If Len(sig) > 0 Then
            s = s & vbCrLf & sig & vbCrLf & _
                "    On Error GoTo Err_Handler" & vbCrLf & _
                "    ' " & what & vbCrLf & _
                "    Exit Sub" & vbCrLf & _
                "Err_Handler:" & vbCrLf & _
                "    HandleError """ & nm & "." & proc & """" & vbCrLf & _
                "End Sub" & vbCrLf
        End If
        rs.MoveNext
    Loop
    rs.Close
    DesignerStub = s
End Function

' Build layout id as the pop-up form dzPreview, with the builder calls
' DesignerCode writes but no handlers, and open it. Its controls work (a
' toggle switches, a slider slides) but call no code.
Public Sub DesignerPreview(ByVal id As Long)
    Dim db As DAO.Database, rs As DAO.Recordset, title As String, fw As Long, fh As Long
    If IsFormOpen(DZ_PREVIEW) Then DoCmd.Close acForm, DZ_PREVIEW, acSaveNo
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT layout_name, form_width, form_height FROM kit_layout WHERE id=" & id, dbOpenSnapshot)
    If rs.EOF Then Err.Raise 5, , "No form designer layout " & id
    title = OneLine(rs!layout_name): fw = Nz(rs!form_width, 0): fh = Nz(rs!form_height, 0)
    rs.Close
    On Error GoTo Fail
    NewForm "", 0, px(fw), px(fh)
    F.Caption = "Preview: " & title
    F.PopUp = True
    F.Modal = False
    F.AutoCenter = True
    F.AutoResize = True
    F.BorderStyle = 3
    F.MinMaxButtons = 0
    F.ControlBox = True
    F.CloseButton = True
    Set rs = db.OpenRecordset(ItemsSql(id), dbOpenSnapshot)
    Do Until rs.EOF
        BuildItem rs
        rs.MoveNext
    Loop
    rs.Close
    SaveForm DZ_PREVIEW
    On Error GoTo 0
    DoCmd.OpenForm DZ_PREVIEW
    Exit Sub
Fail:
    Dim errNo As Long, errSrc As String, errDesc As String
    errNo = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    On Error Resume Next
    If Not F Is Nothing Then DoCmd.Close acForm, F.Name, acSaveNo
    Set F = Nothing
    On Error GoTo 0
    Err.Raise errNo, errSrc, errDesc
End Sub

Private Function ItemsSql(ByVal id As Long) As String
    ItemsSql = "SELECT item_name, kind, item_text, item_text2, item_text3, item_choice, x, y, w, h " & _
        "FROM kit_layout_item WHERE layout_id=" & id & " ORDER BY item_order"
End Function

' The row's element as a kit builder call (DesignerCode); "" for a kind the
' designer doesn't know. BuildItem makes the same call.
Private Function ItemCall(rs As DAO.Recordset) As String
    Dim nm As String, t1 As String, t2 As String, t3 As String, ch As String, n As String
    Dim x As String, y As String, w As String, h As String, at As String, parts() As String, i As Long
    nm = ItemName(rs): n = VbText(nm)
    t1 = Nz(rs!item_text, ""): t2 = Nz(rs!item_text2, ""): t3 = Nz(rs!item_text3, ""): ch = Nz(rs!item_choice, "")
    x = "px(" & Nz(rs!x, 0) & ")": y = "px(" & Nz(rs!y, 0) & ")": w = "px(" & Nz(rs!w, 0) & ")": h = "px(" & Nz(rs!h, 0) & ")"
    at = ", " & x & ", " & y & ", " & w & ", " & h
    Select Case Nz(rs!kind, "")
        Case "card": ItemCall = "Card " & n & ", " & VbText(t1) & at
        Case "button": ItemCall = "Btn " & n & ", " & VbText(t1) & at & ", " & VbText(ButtonStyle(ch))
        Case "label": ItemCall = "Lbl " & n & ", " & VbText(t1) & at
        Case "progress": ItemCall = "Progress " & n & ", " & VbText(PercentText(t1)) & at
        Case "text": ItemCall = "Txt " & n & ", """"" & at
        Case "date": ItemCall = "DateBox " & n & ", """", " & x & ", " & y & ", " & h
        Case "secret": ItemCall = "SecretBox " & n & ", """"" & at
        Case "toggle": ItemCall = "Toggle " & n & ", """"" & at & ", " & VbText(nm & "Changed")
        Case "segmented"
            ItemCall = "Segmented " & n & ", """", " & x & ", " & y & ", " & w & ", " & VbText(SegmentItems(t1)) & ", " & h & ", " & _
                VbText(nm & "Changed")
        Case "slider"
            ItemCall = "Slider " & n & ", """", " & x & ", " & y & ", " & w & ", " & NumText(t1) & ", " & NumText(t2) & ", " & _
                NumText(t3) & ", " & VbText(nm & "Changed")
        Case "rating": ItemCall = "Rating " & n & ", """", " & x & ", " & y & ", " & CLng(Val(t1)) & ", px(24), False, " & VbText(nm & "Changed")
        Case "stepper"
            ItemCall = "Stepper " & n & ", " & x & ", " & y & ", " & w & ", " & CLng(Val(t2)) & ", " & VbText(nm & "Step")
            parts = Split(t1, ";")
            For i = 0 To UBound(parts)
                ItemCall = ItemCall & ", " & VbText(Trim$(parts(i)))
            Next
        Case "search": ItemCall = "SearchBox " & n & ", " & x & ", " & y & ", " & w & ", " & VbText(nm & "Search") & ", " & VbText(t1) & ", " & h
        Case "stat"
            ItemCall = "StatTile " & n & ", " & VbText(t1) & ", " & VbText("=" & NumText(t2)) & ", " & x & ", " & y & ", " & w & ", " & h & _
                ", """", " & VbText(t3)
        Case "alert": ItemCall = "Alert " & n & at & ", " & VbText(AlertVariant(ch)) & ", " & VbText(t1) & ", " & VbText(t2)
        Case "avatar": ItemCall = "Avatar " & n & ", " & VbText(AvatarSource(t1)) & ", " & x & ", " & y & ", " & w
        Case "empty"
            ItemCall = "EmptyState " & n & ", " & x & ", " & y & ", " & w & ", " & VbText(DZ_EMPTY_GLYPH) & ", " & VbText(t1) & ", " & _
                VbText(t2) & ", " & VbText(t3) & ", """", True"
    End Select
End Function

' The row's element built on the form under construction (DesignerPreview):
' the call ItemCall writes, with no handlers, and buttons that call no code.
Private Sub BuildItem(rs As DAO.Recordset)
    Dim nm As String, t1 As String, t2 As String, t3 As String, ch As String, c As Control, parts() As String
    Dim x As Long, y As Long, w As Long, h As Long
    nm = ItemName(rs)
    t1 = Nz(rs!item_text, ""): t2 = Nz(rs!item_text2, ""): t3 = Nz(rs!item_text3, ""): ch = Nz(rs!item_choice, "")
    x = px(Nz(rs!x, 0)): y = px(Nz(rs!y, 0)): w = px(Nz(rs!w, 0)): h = px(Nz(rs!h, 0))
    Select Case Nz(rs!kind, "")
        Case "card": Card nm, t1, x, y, w, h
        Case "button"
            Set c = Btn(nm, t1, x, y, w, h, ButtonStyle(ch))
            c.OnClick = ""
        Case "label": Lbl nm, t1, x, y, w, h
        Case "progress": Progress nm, PercentText(t1), x, y, w, h
        Case "text": Txt nm, "", x, y, w, h
        Case "date": DateBox nm, "", x, y, h
        Case "secret": SecretBox nm, "", x, y, w, h
        Case "toggle": Toggle nm, "", x, y, w, h
        Case "segmented": Segmented nm, "", x, y, w, SegmentItems(t1), h
        Case "slider": Slider nm, "", x, y, w, Val(t1), Val(t2), Val(t3)
        Case "rating": Rating nm, "", x, y, CLng(Val(t1)), px(24)
        Case "stepper"
            parts = Split(t1, ";")
            StepperFrom nm, x, y, w, CLng(Val(t2)), parts
        Case "search": SearchBox nm, x, y, w, "", t1, h
        Case "stat": StatTile nm, t1, "=" & NumText(t2), x, y, w, h, "", t3
        Case "alert": Alert nm, x, y, w, h, AlertVariant(ch), t1, t2
        Case "avatar": Avatar nm, AvatarSource(t1), x, y, w
        Case "empty"
            EmptyState nm, x, y, w, DZ_EMPTY_GLYPH, t1, t2, t3, "", True
            On Error Resume Next                 ' no button without a caption
            F.Controls(nm & "_b").OnClick = ""
            On Error GoTo 0
    End Select
End Sub

' Stepper takes its captions as a ParamArray: pass up to 8 of them.
Private Sub StepperFrom(ByVal nm As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal current As Long, parts() As String)
    Dim c(0 To 7) As String, i As Long
    For i = 0 To UBound(parts)
        c(i) = Trim$(parts(i))
    Next
    Select Case UBound(parts) + 1
        Case 2: Stepper nm, x, y, w, current, "", c(0), c(1)
        Case 3: Stepper nm, x, y, w, current, "", c(0), c(1), c(2)
        Case 4: Stepper nm, x, y, w, current, "", c(0), c(1), c(2), c(3)
        Case 5: Stepper nm, x, y, w, current, "", c(0), c(1), c(2), c(3), c(4)
        Case 6: Stepper nm, x, y, w, current, "", c(0), c(1), c(2), c(3), c(4), c(5)
        Case 7: Stepper nm, x, y, w, current, "", c(0), c(1), c(2), c(3), c(4), c(5), c(6)
        Case 8: Stepper nm, x, y, w, current, "", c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7)
    End Select
End Sub

' The row's item_name when it is a plain control name (letters, digits and
' _, starting with a letter), else its kind's prefix and order (btn7).
Private Function ItemName(rs As DAO.Recordset) As String
    Dim nm As String, i As Long, ch As Long, ok As Boolean
    nm = Nz(rs!item_name, "")
    ok = Len(nm) > 0 And Len(nm) <= 48
    For i = 1 To Len(nm)
        ch = AscW(Mid$(nm, i, 1))
        If Not ((ch >= 65 And ch <= 90) Or (ch >= 97 And ch <= 122) Or (i > 1 And ((ch >= 48 And ch <= 57) Or ch = 95))) Then ok = False
    Next
    If Not ok Then nm = KindPrefix(Nz(rs!kind, "")) & rs.AbsolutePosition + 1
    ItemName = nm
End Function

' A VBA string literal for s, on one line.
Private Function VbText(ByVal s As String) As String
    VbText = """" & Replace(OneLine(s), """", """""") & """"
End Function

Private Function OneLine(ByVal s As Variant) As String
    OneLine = Replace(Replace(Nz(s, ""), vbCr, " "), vbLf, " ")
End Function

' A progress element's percent, "0" when it isn't one.
Private Function PercentText(ByVal s As String) As String
    If IsWholePercent(s) Then PercentText = CStr(CLng(s)) Else PercentText = "0"
End Function

' A number as VBA source: digits, at most one ".", and a leading "-"; "0"
' when it isn't one. (Str$ always writes ".", whatever the Windows locale.)
Private Function NumText(ByVal s As String) As String
    If IsNumberText(s) Then NumText = Trim$(Str$(Val(s))) Else NumText = "0"
End Function

' Segmented's items from the options "All;Open;Done": "all|All;open|Open;done|Done".
' A key is the caption's letters and digits in lower case, or opt<n> when
' that is empty or taken.
Private Function SegmentItems(ByVal options As String) As String
    Dim parts() As String, keys() As String, i As Long, j As Long, k As String, out As String, ch As Long, cap As String
    parts = Split(options, ";")
    ReDim keys(UBound(parts))
    For i = 0 To UBound(parts)
        cap = Trim$(parts(i))
        k = ""
        For j = 1 To Len(cap)
            ch = AscW(Mid$(cap, j, 1))
            If (ch >= 48 And ch <= 57) Or (ch >= 97 And ch <= 122) Then k = k & ChrW(ch)
            If ch >= 65 And ch <= 90 Then k = k & ChrW(ch + 32)
        Next
        For j = 0 To i - 1
            If keys(j) = k Then k = ""
        Next
        If Len(k) = 0 Then k = "opt" & (i + 1)
        keys(i) = k
        out = out & IIf(i > 0, ";", "") & k & "|" & Replace(Replace(cap, "|", "/"), ";", ",")
    Next
    SegmentItems = out
End Function

Private Function ButtonStyle(ByVal ch As String) As String
    If ch = "outline" Or ch = "destructive" Or ch = "ghost" Then ButtonStyle = ch Else ButtonStyle = "primary"
End Function

Private Function AlertVariant(ByVal ch As String) As String
    If ch = "success" Or ch = "warning" Or ch = "destructive" Then AlertVariant = ch Else AlertVariant = "info"
End Function

' An avatar shows the initials of a fixed name: the expression ="Jane Doe".
Private Function AvatarSource(ByVal nameText As String) As String
    AvatarSource = "=""" & Replace(OneLine(nameText), """", """""") & """"
End Function

' False in a locked (.accde) copy and in the Access Runtime, which can't create forms.
Private Function CanBuildForms() As Boolean
    Dim v As String
    If SysCmd(acSysCmdRuntime) Then Exit Function
    On Error Resume Next
    v = CurrentDb.Properties("MDE")
    On Error GoTo 0
    CanBuildForms = (v <> "T")
End Function

' ---------------------------------------------------------------------------
' Layout tables and for tests
' ---------------------------------------------------------------------------

' Make kit_layout (one row per saved layout) and kit_layout_item (its
' elements, px from the canvas corner) when they are missing, and add the
' settings columns a table from the first version lacks.
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
    AddColumn db, "kit_layout_item", "item_text2", "TEXT(255)"
    AddColumn db, "kit_layout_item", "item_text3", "TEXT(255)"
    AddColumn db, "kit_layout_item", "item_choice", "TEXT(30)"
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

' Draw every item again (after DesignerSetProps).
Public Sub DesignerRedraw(frm As Form)
    Dim i As Long
    For i = 1 To m_n
        Render frm, i
    Next
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
    i = AddItem(kind, lx, ly, DefaultW(kind), DefaultH(kind), KindDefault(kind, 1), KindDefault(kind, 2), _
        KindDefault(kind, 3), KindDefault(kind, 4))
    If i = 0 Then
        KitToast "The canvas has room for " & PoolSize(KindPool(kind)) & " " & PoolCaption(KindPool(kind)) & ".", "warning"
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
' The rectangle is kept to the sizes its kind and settings allow, inside the canvas.
Private Function AddItem(ByVal kind As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
        ByVal t1 As String, ByVal t2 As String, ByVal t3 As String, ByVal choice As String) As Long
    Dim s As Long, i As Long, used As Boolean, pool As String
    pool = KindPool(kind)
    For s = 1 To PoolSize(pool)
        used = False
        For i = 1 To m_n
            If KindPool(m_kind(i)) = pool And m_slot(i) = s Then used = True: Exit For
        Next
        If Not used Then Exit For
    Next
    If s > PoolSize(pool) Or m_n >= DZ_MAX Then Exit Function
    m_n = m_n + 1
    m_kind(m_n) = kind
    m_slot(m_n) = s
    m_text(m_n) = t1: m_text2(m_n) = t2: m_text3(m_n) = t3: m_choice(m_n) = choice
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

' Item i's rectangle: within the sizes its kind and settings allow (an avatar
' stays round), and inside the canvas.
Private Sub FitRect(ByVal i As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    w = Clamp(w, MinW(i), MaxW(i))
    If m_kind(i) = "avatar" Then h = w Else h = Clamp(h, MinH(i), MaxH(i))
    m_w(i) = w
    m_h(i) = h
    m_x(i) = Clamp(x, 0, m_cw - w)
    m_y(i) = Clamp(y, 0, m_ch - h)
End Sub

' The topmost item under (x, y) px: the pools draw back to front (KindLayer),
' and within a pool the later slot on top. 0 for none.
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

' The palette tile at (x, y) px from the palette's corner, -1 for a gap or
' past the last tile.
Private Function TileAt(ByVal x As Double, ByVal y As Double) As Long
    Dim col As Long, row As Long, idx As Long
    TileAt = -1
    If x < 0 Or y < 0 Then Exit Function
    col = Int(x / (DZ_TILE_W + 8))
    row = Int(y / DZ_TILE_STEP)
    If col > 1 Then Exit Function
    If x - col * (DZ_TILE_W + 8) >= DZ_TILE_W Or y - row * DZ_TILE_STEP >= DZ_TILE_H Then Exit Function
    idx = row * 2 + col
    If idx <= UBound(Split(DZ_KINDS, ";")) Then TileAt = idx
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

' Show the saved layout id (elements of a kind this version doesn't know, or
' past a pool's size, are left out, with a toast).
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
    Set rs = db.OpenRecordset("SELECT kind, item_text, item_text2, item_text3, item_choice, x, y, w, h " & _
        "FROM kit_layout_item WHERE layout_id=" & id & " ORDER BY item_order", dbOpenSnapshot)
    Do Until rs.EOF
        If Len(KindPool(Nz(rs!kind, ""))) = 0 Then
            skipped = skipped + 1
        Else
            i = AddItem(rs!kind, Nz(rs!x, 0), Nz(rs!y, 0), Nz(rs!w, 0), Nz(rs!h, 0), Nz(rs!item_text, ""), _
                Nz(rs!item_text2, ""), Nz(rs!item_text3, ""), Nz(rs!item_choice, ""))
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
    Select Case KindPool(m_kind(i))
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
        Case "gen"
            MoveCtl frm(nm), l, t, w, h
            MoveCtl frm(nm & "_t"), l + px(4), t, w - px(8), h
            frm(nm & "_t").Caption = DescText(i)
            frm(nm & "_t").TopMargin = 0
            VCenter frm(nm & "_t"), h, 9
            frm(nm).Visible = True
            frm(nm & "_t").Visible = True
        Case "label", "button"
            MoveCtl frm(nm), l, t, w, h
            frm(nm).Caption = m_text(i)
            frm(nm).TopMargin = 0
            VCenter frm(nm), h, IIf(m_kind(i) = "button", 9, 10), (m_kind(i) = "button")
            frm(nm).Visible = True
    End Select
End Sub

' What an outlined box on the canvas says it will be.
Private Function DescText(ByVal i As Long) As String
    Select Case m_kind(i)
        Case "segmented": DescText = Replace(m_text(i), ";", "  |  ")
        Case "slider": DescText = "Slider, " & m_text(i) & " to " & m_text2(i)
        Case "rating": DescText = String$(CLng(Val(m_text(i))), ChrW(9733))
        Case "stepper": DescText = Replace(m_text(i), ";", "  " & ChrW(8250) & "  ")
        Case "search": DescText = m_text(i)
        Case "stat": DescText = m_text(i) & ": " & m_text2(i)
        Case "alert": DescText = AlertVariant(m_choice(i)) & ": " & m_text(i)
        Case "avatar": DescText = AvatarInitials(m_text(i))
        Case "empty": DescText = m_text(i)
        Case Else: DescText = KindCaption(m_kind(i))
    End Select
End Function

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
    Select Case KindPool(kind)
        Case "card", "gen": frm(nm & "_t").Visible = v
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
    SetEnabled frm("btnDzEdit"), i > 0
    SetEnabled frm("btnDzDelete"), i > 0
    If i = 0 Then
        frm("lblDzKind").Caption = "Nothing selected"
        frm("lblDzPos").Caption = ""
    Else
        frm("lblDzKind").Caption = KindCaption(m_kind(i))
        frm("lblDzPos").Caption = m_x(i) & ", " & m_y(i) & "   " & m_w(i) & " x " & m_h(i) & " px"
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
    If Len(msg) = 0 Then msg = "Drag elements onto the canvas; drag to move, drag the corner to resize, double-click for settings."
    frm("lblSubtitle").Caption = msg
End Sub

' ---------------------------------------------------------------------------
' Element kinds
' ---------------------------------------------------------------------------

' The pool a kind's canvas preview comes from: card, progress, gen (an
' outlined box), label or button; "" for a kind this version doesn't know.
Private Function KindPool(ByVal kind As String) As String
    Select Case kind
        Case "card", "progress", "label", "button": KindPool = kind
        Case "text", "date", "secret", "toggle", "segmented", "slider", "rating", "stepper", "search", "stat", _
            "alert", "avatar", "empty": KindPool = "gen"
    End Select
End Function

' Back-to-front order of the pools: 0 card, 1 progress, 2 gen, 3 label, 4 button.
Private Function KindLayer(ByVal kind As String) As Long
    Select Case KindPool(kind)
        Case "card": KindLayer = 0
        Case "progress": KindLayer = 1
        Case "gen": KindLayer = 2
        Case "label": KindLayer = 3
        Case Else: KindLayer = 4
    End Select
End Function

Private Function PoolSize(ByVal pool As String) As Long
    Select Case pool
        Case "card": PoolSize = DZ_CARDS
        Case "progress": PoolSize = DZ_PROGS
        Case "gen": PoolSize = DZ_GENS
        Case "label": PoolSize = DZ_LBLS
        Case "button": PoolSize = DZ_BTNS
    End Select
End Function

Private Function PoolCaption(ByVal pool As String) As String
    Select Case pool
        Case "card": PoolCaption = "cards"
        Case "progress": PoolCaption = "progress bars"
        Case "label": PoolCaption = "labels"
        Case "button": PoolCaption = "buttons"
        Case Else: PoolCaption = "elements other than cards, buttons, labels and progress bars"
    End Select
End Function

' The preview control of slot s in a kind's pool.
Private Function SlotName(ByVal kind As String, ByVal s As Long) As String
    Select Case KindPool(kind)
        Case "card": SlotName = "dzCard" & s
        Case "progress": SlotName = "dzProg" & s
        Case "gen": SlotName = "dzGen" & s
        Case "label": SlotName = "dzLbl" & s
        Case Else: SlotName = "dzBtn" & s
    End Select
End Function

Private Function KindCaption(ByVal kind As String) As String
    Select Case kind
        Case "card": KindCaption = "Card"
        Case "button": KindCaption = "Button"
        Case "label": KindCaption = "Label"
        Case "progress": KindCaption = "Progress"
        Case "text": KindCaption = "Text box"
        Case "date": KindCaption = "Date"
        Case "secret": KindCaption = "Secret"
        Case "toggle": KindCaption = "Toggle"
        Case "segmented": KindCaption = "Segmented"
        Case "slider": KindCaption = "Slider"
        Case "rating": KindCaption = "Rating"
        Case "stepper": KindCaption = "Stepper"
        Case "search": KindCaption = "Search"
        Case "stat": KindCaption = "Stat tile"
        Case "alert": KindCaption = "Alert"
        Case "avatar": KindCaption = "Avatar"
        Case "empty": KindCaption = "Empty state"
    End Select
End Function

' Name prefix of a kind's saved elements (card1, btn1, seg1).
Private Function KindPrefix(ByVal kind As String) As String
    Select Case kind
        Case "card": KindPrefix = "card"
        Case "button": KindPrefix = "btn"
        Case "label": KindPrefix = "lbl"
        Case "progress": KindPrefix = "prog"
        Case "text": KindPrefix = "txt"
        Case "date": KindPrefix = "date"
        Case "secret": KindPrefix = "secret"
        Case "toggle": KindPrefix = "tgl"
        Case "segmented": KindPrefix = "seg"
        Case "slider": KindPrefix = "sld"
        Case "rating": KindPrefix = "rate"
        Case "stepper": KindPrefix = "step"
        Case "search": KindPrefix = "search"
        Case "stat": KindPrefix = "stat"
        Case "alert": KindPrefix = "alert"
        Case "avatar": KindPrefix = "av"
        Case Else: KindPrefix = "empty"
    End Select
End Function

' The caption of the settings dialog's row r (1-3 text, 4 a list) for a
' kind, "" when it doesn't use that row.
Private Function RowCaption(ByVal kind As String, ByVal r As Long) As String
    Dim caps As Variant
    Select Case kind
        Case "card": caps = Array("TITLE", "", "", "")
        Case "button": caps = Array("CAPTION", "", "", "STYLE")
        Case "label": caps = Array("TEXT", "", "", "")
        Case "progress": caps = Array("PERCENT DONE (0 TO 100)", "", "", "")
        Case "segmented": caps = Array("OPTIONS, SEPARATED BY ;", "", "", "")
        Case "slider": caps = Array("FROM", "TO", "STEP", "")
        Case "rating": caps = Array("STARS (3 TO 10)", "", "", "")
        Case "stepper": caps = Array("STEPS, SEPARATED BY ;", "CURRENT STEP", "", "")
        Case "search": caps = Array("PLACEHOLDER", "", "", "")
        Case "stat": caps = Array("LABEL", "VALUE", "CAPTION", "")
        Case "alert": caps = Array("TITLE", "MESSAGE", "", "VARIANT")
        Case "avatar": caps = Array("NAME", "", "", "")
        Case "empty": caps = Array("TITLE", "TEXT", "BUTTON (OPTIONAL)", "")
        Case Else: caps = Array("", "", "", "")
    End Select
    RowCaption = caps(r - 1)
End Function

' What a new element's settings start as: row r, as RowCaption.
Private Function KindDefault(ByVal kind As String, ByVal r As Long) As String
    Dim v As Variant
    Select Case kind
        Case "card": v = Array("Card title", "", "", "")
        Case "button": v = Array("Button", "", "", "primary")
        Case "label": v = Array("Label text", "", "", "")
        Case "progress": v = Array("40", "", "", "")
        Case "segmented": v = Array("All;Open;Done", "", "", "")
        Case "slider": v = Array("0", "100", "10", "")
        Case "rating": v = Array("5", "", "", "")
        Case "stepper": v = Array("Details;Owner;Review", "1", "", "")
        Case "search": v = Array("Search" & ChrW(8230), "", "", "")
        Case "stat": v = Array("Open items", "42", "vs last week", "")
        Case "alert": v = Array("Heads up", "Something here needs your attention.", "", "info")
        Case "avatar": v = Array("Jane Doe", "", "", "")
        Case "empty": v = Array("Nothing here yet", "Items you add show here.", "Add an item", "")
        Case Else: v = Array("", "", "", "")
    End Select
    KindDefault = v(r - 1)
End Function

' The list in row 4 (a value list, so no commas).
Private Function ChoiceList(ByVal kind As String) As String
    Select Case kind
        Case "button": ChoiceList = "primary;outline;destructive;ghost"
        Case "alert": ChoiceList = "info;success;warning;destructive"
    End Select
End Function

' What the settings take, under the rows.
Private Function KindHint(ByVal kind As String) As String
    Select Case kind
        Case "segmented": KindHint = "2 to 6 options. Each needs 40px of the control's width."
        Case "slider": KindHint = "From, to and step make 2 to 41 stops; the step must divide the range."
        Case "stepper": KindHint = "2 to 8 steps, each 64px wide or more. The current step is 1 to one past the last (all done)."
        Case "stat": KindHint = "The value is a number here; in the app it is usually a field or a count."
        Case "empty": KindHint = "Leave the button empty for none."
        Case "progress": KindHint = "A whole number."
        Case Else: KindHint = ""
    End Select
End Function

' What is wrong with these settings for a kind, "" when nothing is.
Private Function PropsProblem(ByVal kind As String, ByVal t1 As String, ByVal t2 As String, ByVal t3 As String, _
        ByVal choice As String) As String
    Dim n As Long, a As Double, b As Double, s As Double, stops As Double
    Select Case kind
        Case "button"
            If Len(ChoiceList(kind)) > 0 And InStr(";" & ChoiceList(kind) & ";", ";" & choice & ";") = 0 Then _
                PropsProblem = "Pick a style from the list."
        Case "alert"
            If InStr(";" & ChoiceList(kind) & ";", ";" & choice & ";") = 0 Then PropsProblem = "Pick a variant from the list."
        Case "progress"
            If Not IsWholePercent(t1) Then PropsProblem = "Percent done is a whole number from 0 to 100."
        Case "segmented"
            n = PartCount(t1)
            If n < 2 Or n > 6 Then PropsProblem = "A segmented control takes 2 to 6 options, each with a caption, separated by ;."
        Case "slider"
            If Not (IsNumberText(t1) And IsNumberText(t2) And IsNumberText(t3)) Then
                PropsProblem = "From, to and step are numbers."
            Else
                a = Val(t1): b = Val(t2): s = Val(t3)
                If b <= a Or s <= 0 Then
                    PropsProblem = "To must be more than from, and the step more than 0."
                Else
                    stops = (b - a) / s + 1
                    If Abs(stops - CLng(stops)) > 0.000001 Or stops < 2 Or stops > 41 Then _
                        PropsProblem = "The step must divide the range into 2 to 41 stops."
                End If
            End If
        Case "rating"
            If Not IsWholeNumber(t1) Then
                PropsProblem = "Stars is a whole number from 3 to 10."
            ElseIf Val(t1) < 3 Or Val(t1) > 10 Then
                PropsProblem = "Stars is a whole number from 3 to 10."
            End If
        Case "stepper"
            n = PartCount(t1)
            If n < 2 Or n > 8 Then
                PropsProblem = "A stepper takes 2 to 8 steps, each with a caption, separated by ;."
            ElseIf Not IsWholeNumber(t2) Then
                PropsProblem = "The current step is a whole number from 1 to " & n + 1 & "."
            ElseIf Val(t2) < 1 Or Val(t2) > n + 1 Then
                PropsProblem = "The current step is a whole number from 1 to " & n + 1 & "."
            End If
        Case "stat"
            If Not IsNumberText(t2) Then PropsProblem = "The value is a number."
        Case "empty"
            If Len(t1) = 0 Then PropsProblem = "An empty state needs a title."
    End Select
End Function

' Sizes (px) a new element gets.
Private Function DefaultW(ByVal kind As String) As Long
    Select Case kind
        Case "card": DefaultW = 240
        Case "button": DefaultW = 96
        Case "label": DefaultW = 120
        Case "progress": DefaultW = 160
        Case "text", "secret": DefaultW = 200
        Case "date": DefaultW = DATEFIELD_W \ 15
        Case "toggle": DefaultW = 44
        Case "segmented", "slider", "search", "stat": DefaultW = 240
        Case "rating": DefaultW = 5 * 24
        Case "stepper", "alert": DefaultW = 360
        Case "avatar": DefaultW = 32
        Case Else: DefaultW = 320
    End Select
End Function

Private Function DefaultH(ByVal kind As String) As Long
    Select Case kind
        Case "card": DefaultH = 136
        Case "button", "segmented", "search", "avatar": DefaultH = 32
        Case "label", "toggle", "rating": DefaultH = 24
        Case "progress": DefaultH = 8
        Case "text", "date", "secret": DefaultH = 30
        Case "slider": DefaultH = 20
        Case "stepper": DefaultH = 56
        Case "stat": DefaultH = 112
        Case "alert": DefaultH = 64
        Case Else: DefaultH = 180
    End Select
End Function

' The smallest and largest item i can be, from its kind and settings.
Private Function MinW(ByVal i As Long) As Long
    Select Case m_kind(i)
        Case "card", "slider": MinW = 80
        Case "button": MinW = 48
        Case "label", "progress": MinW = 24
        Case "text": MinW = 60
        Case "secret", "search": MinW = 120
        Case "date", "toggle", "rating": MinW = DefaultW(m_kind(i))
        Case "segmented": MinW = 6 + 40 * PartCount(m_text(i))
        Case "stepper": MinW = 64 * PartCount(m_text(i))
        Case "stat", "alert": MinW = 160
        Case "avatar": MinW = 24
        Case Else: MinW = 240
    End Select
    If m_kind(i) = "rating" Then MinW = 24 * CLng(Val(m_text(i)))
End Function

Private Function MaxW(ByVal i As Long) As Long
    Select Case m_kind(i)
        Case "date", "toggle": MaxW = DefaultW(m_kind(i))
        Case "rating": MaxW = 24 * CLng(Val(m_text(i)))
        Case "avatar": MaxW = 64
        Case Else: MaxW = m_cw
    End Select
End Function

Private Function MinH(ByVal i As Long) As Long
    Select Case m_kind(i)
        Case "card": MinH = 48
        Case "button", "text", "secret", "date", "segmented": MinH = 24
        Case "label": MinH = 16
        Case "progress": MinH = 4
        Case "search": MinH = 28
        Case "alert": MinH = 40
        Case Else: MinH = DefaultH(m_kind(i))
    End Select
End Function

Private Function MaxH(ByVal i As Long) As Long
    Select Case m_kind(i)
        Case "card", "label": MaxH = m_ch
        Case "button", "segmented": MaxH = 48
        Case "progress": MaxH = 16
        Case "text": MaxH = 120
        Case "date", "secret", "search": MaxH = 40
        Case "alert": MaxH = 160
        Case Else: MaxH = DefaultH(m_kind(i))
    End Select
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

' How many ;-separated parts s has, 0 when any is empty.
Private Function PartCount(ByVal s As String) As Long
    Dim parts() As String, i As Long
    If Len(Trim$(s)) = 0 Then Exit Function
    parts = Split(s, ";")
    For i = 0 To UBound(parts)
        If Len(Trim$(parts(i))) = 0 Then Exit Function
    Next
    PartCount = UBound(parts) + 1
End Function

' "0" to "100", digits only.
Private Function IsWholePercent(ByVal s As String) As Boolean
    If Not IsWholeNumber(s) Or Len(s) > 3 Then Exit Function
    IsWholePercent = (CLng(s) <= 100)
End Function

' Digits only, 1 to 9 of them.
Private Function IsWholeNumber(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Or Len(s) > 9 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next
    IsWholeNumber = True
End Function

' A number written with digits, at most one "." and an optional leading "-"
' (what Val reads the same in every Windows locale).
Private Function IsNumberText(ByVal s As String) As Boolean
    Dim i As Long, ch As String, dots As Long, digits As Long
    If Len(s) = 0 Or Len(s) > 15 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "-" And i = 1 Then
        ElseIf ch = "." Then
            dots = dots + 1
        ElseIf ch >= "0" And ch <= "9" Then
            digits = digits + 1
        Else
            Exit Function
        End If
    Next
    IsNumberText = (dots <= 1 And digits > 0)
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

' Add column col (of SQL type sqlType) to table tbl when it lacks it.
Private Sub AddColumn(db As DAO.Database, ByVal tbl As String, ByVal col As String, ByVal sqlType As String)
    Dim fld As DAO.Field
    For Each fld In db.TableDefs(tbl).Fields
        If fld.Name = col Then Exit Sub
    Next
    db.Execute "ALTER TABLE " & tbl & " ADD COLUMN " & col & " " & sqlType, dbFailOnError
    db.TableDefs.Refresh
End Sub
