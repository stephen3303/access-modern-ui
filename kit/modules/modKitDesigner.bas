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
' and builds the settings dialog dlgDzProps, kit\forms\dlgDzProps.vba, and
' the tab order dialog dlgDzTabs, kit\forms\dlgDzTabs.vba); their
' code-behind hands every event to the Designer* routines here. Positions
' and sizes are in px from the canvas's top-left corner, on an 8px grid
' (DZ_GRID). Double-click an element, or Edit, for its settings; Esc cancels
' a palette click; Delete removes the selected element; the arrow keys move
' it 8px (1px with Shift).
'
' Form (the same settings dialog) sets what the layout builds: a page or a
' dialog, its size (the outlined area of the canvas) and its record source,
' a table or query whose fields the elements that hold a value can be bound
' to. On a page each element can be anchored to the right or bottom edge, or
' stretch with the window. Tab order lists the elements that take the focus:
' in reading order (top to bottom, then left to right) until moved.

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
Private Const DZ_TABS As String = "dlgDzTabs"      ' the tab order dialog
Private Const DZ_EMPTY_GLYPH As String = "E7C1"    ' an empty state's icon
Private Const DZ_ROWS As Long = 6               ' settings dialog rows: 3 text, then 3 lists
Private Const DZ_MIN_FW As Long = 160           ' smallest form, px
Private Const DZ_MIN_FH As Long = 120

Private Const MODE_IDLE As Integer = 0
Private Const MODE_MOVE As Integer = 1
Private Const MODE_SIZE As Integer = 2
Private Const MODE_DRAG As Integer = 3          ' dragging a new element in from the palette
Private Const MODE_ARMED As Integer = 4         ' palette element clicked; the next canvas click places it

' The layout being edited: items 1..m_n, each in slot m_slot of its kind's
' pool. m_text, m_text2 and m_text3 are the settings in the dialog's first
' three rows, m_choice the fourth (a pick from a list), m_field the field
' it is bound to ("" for none), m_anchor its anchor (AnchorOf; "" stays
' top left) and m_tab its place in the tab order (0 for reading order).
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
Private m_field(1 To DZ_MAX) As String
Private m_anchor(1 To DZ_MAX) As String
Private m_tab(1 To DZ_MAX) As Long
Private m_n As Long
' The form the layout builds: "page" or "dialog", its size in px (inside the
' canvas), record source ("" for none) and a dialog's title ("" for the
' layout's name).
Private m_formKind As String
Private m_fw As Long
Private m_fh As Long
Private m_recSrc As String
Private m_title As String
Private m_propsForm As Boolean                  ' the settings dialog is showing Form, not an element
Private m_tabSeq() As Long                      ' the tab order dialog's list: item numbers in order
Private m_tabReading As Boolean                 ' ... and whether it is the reading order
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
' Delete) on the left, the layout list, name, Form, Tab order, New, Generate
' and Save above the canvas on the right. Saved as formName with
' pgDesigner.vba's code. Builds its dialogs dlgDzProps and dlgDzTabs first.
Public Sub BuildDesigner(ByVal formName As String, ByVal x As Long, ByVal w As Long, ByVal h As Long)
    Dim c As Control, i As Long, ty As Long, py As Long, lw As Long, cx As Long, cy As Long, cw As Long, ch As Long
    Dim kinds() As String, rows As Long
    BuildDesignerDialog
    BuildDesignerTabsDialog
    NewForm "", 0, w, h
    F.KeyPreview = True                          ' Delete, Esc and the arrow keys (Form_KeyDown)
    Ev F, "OnLoad", "OnKeyDown"
    PageHeader "Form designer", "Drag an element from Elements onto the canvas, or click one and then click the canvas. " & _
        "Drag an element to move it, and its corner handle to resize it. Everything snaps to an 8px grid. " & _
        "Double-click an element for its settings, and click Form for the form's: page or dialog, size and record source. " & _
        "Save keeps the layout under its name; Generate builds it as a real form.", x
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

    ' Layout list, name, Form, Tab order, New, Generate and Save.
    Lbl "lblDzOpen", "OPEN A LAYOUT", cx, px(76), px(168), px(16), 8, True, CLR_MUTED_FG
    Set c = Cbo("cboDzOpen", "", "", cx, px(94), px(168), px(30), 2, "0;" & px(168))
    Ev c, "AfterUpdate"
    Lbl "lblDzName", "LAYOUT NAME", cx + px(180), px(76), px(168), px(16), 8, True, CLR_MUTED_FG
    Set c = Txt("txtDzName", "", cx + px(180), px(94), px(168), px(30))
    Ev c, "AfterUpdate"
    Set c = Btn("btnDzForm", "Form" & ChrW(8230), cx + cw - px(416), px(95), px(72), px(28), "outline")
    c.ControlTipText = "Page or dialog, size and record source"
    Btn "btnDzTabs", "Tab order" & ChrW(8230), cx + cw - px(336), px(95), px(92), px(28), "outline"
    Btn "btnDzNew", "New", cx + cw - px(236), px(95), px(64), px(28), "outline"
    Btn "btnDzGenerate", "Generate", cx + cw - px(164), px(95), px(84), px(28), "outline"
    Btn "btnDzSave", "Save", cx + cw - px(72), px(95), px(72), px(28)

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

' The settings dialog dlgDzProps: a title and DZ_ROWS rows (three text
' boxes, then three lists), each with its caption; dlgDzProps.vba shows the
' rows the selected element's kind uses, or Form's (DesignerPropsLoad), and
' checks them on OK (DesignerPropsSave). The last list's first column is
' the setting, its second what the list shows.
Public Sub BuildDesignerDialog()
    Dim c As Control, r As Long, y As Long
    y = px(56 + DZ_ROWS * 64 + 60)
    NewDialog "Element settings", px(400), y + px(30 + 20)
    Ev F, "OnLoad"
    Lbl "lblPTitle", "", px(20), px(18), px(360), px(26), 12, True
    For r = 1 To DZ_ROWS
        Lbl "lblP" & r, "", px(20), px(56 + (r - 1) * 64), px(360), px(16), 8, True, CLR_MUTED_FG
        If r <= 3 Then
            Txt "txtP" & r, "", px(20), px(56 + (r - 1) * 64 + 18), px(360), px(30)
        ElseIf r < DZ_ROWS Then
            Set c = Cbo("cboP" & r, "", "", px(20), px(56 + (r - 1) * 64 + 18), px(360), px(30))
            c.RowSourceType = "Value List"
        Else
            Set c = Cbo("cboP" & r, "", "", px(20), px(56 + (r - 1) * 64 + 18), px(360), px(30), 2, "0;" & px(340))
            c.RowSourceType = "Value List"
        End If
    Next
    Lbl "lblPHint", "", px(20), px(56 + DZ_ROWS * 64), px(360), px(52), 9, False, CLR_MUTED_FG
    Set c = Btn("btnOK", "OK", px(400) - px(20) - 2 * px(88) - px(8), y, px(88), px(30))
    c.Default = True
    Set c = Btn("btnCancel", "Cancel", px(400) - px(20) - px(88), y, px(88), px(30), "outline")
    c.Cancel = True
    SaveForm DZ_PROPS
End Sub

' The tab order dialog dlgDzTabs: the elements that take the focus, in tab
' order (DesignerTabsLoad), with Move up, Move down and Reading order beside
' them, and OK and Cancel; dlgDzTabs.vba hands its events to DesignerTabs*.
Public Sub BuildDesignerTabsDialog()
    Dim c As Control, listY As Long, listH As Long, sideX As Long, btnY As Long
    listY = px(64): listH = px(260)
    sideX = px(SP_DIALOG) + px(300) + px(12)
    btnY = listY + listH + px(20)
    NewDialog "Tab order", px(460), btnY + px(30) + px(SP_DIALOG)
    Ev F, "OnLoad"
    Lbl "lblHint", "The order the Tab key moves through the form. Until you move one, it is the reading order: " & _
        "top to bottom, then left to right.", px(SP_DIALOG), px(18), px(412), px(36), 9, False, CLR_MUTED_FG
    ' Hidden position, then what the element is: 20px left for the scroll bar.
    Set c = Lst("lstTabs", "", px(SP_DIALOG), listY, px(300), listH, 2, "0;" & px(280))
    c.RowSourceType = "Value List"
    c.ColumnHeads = False
    Ev c, "AfterUpdate"
    Btn "btnUp", "Move up", sideX, listY, px(100), px(30), "outline"
    Btn "btnDown", "Move down", sideX, listY + px(38), px(100), px(30), "outline"
    Btn "btnReading", "Reading order", sideX, listY + px(88), px(100), px(30), "ghost"
    Set c = Btn("btnOK", "OK", px(460) - px(SP_DIALOG) - 2 * px(88) - px(8), btnY, px(88), px(30))
    c.Default = True
    Set c = Btn("btnCancel", "Cancel", px(460) - px(SP_DIALOG) - px(88), btnY, px(88), px(30), "outline")
    c.Cancel = True
    SaveForm DZ_TABS
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
                If p < m_fw And q < m_fh Then Place frm, m_tool, p, q     ' past the form's edge: still armed
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
            If p >= 0 And q >= 0 And p < m_fw And q < m_fh Then
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
    Dim r As Long, used As Boolean
    If m_sel = 0 Then Exit Sub
    For r = 1 To DZ_ROWS
        If Len(RowCaption(m_kind(m_sel), r)) > 0 Then used = True
    Next
    If Not used Then
        KitToast "A " & LCase$(KindCaption(m_kind(m_sel))) & " has no settings on a dialog: size it on the canvas.", "info"
        Exit Sub
    End If
    m_propsForm = False
    If RunDialog(DZ_PROPS) Then
        Render frm, m_sel
        ShowSelection frm
    End If
End Sub

' btnDzForm: the form's settings (page or dialog, size, record source) in
' dlgDzProps; the canvas shows the new size.
Public Sub DesignerForm(frm As Form)
    m_propsForm = True
    If RunDialog(DZ_PROPS) Then
        ShowCanvas frm
        ShowSelection frm
    End If
    m_propsForm = False
End Sub

' btnDzTabs: the tab order of the elements that take the focus (dlgDzTabs).
Public Sub DesignerTabs(frm As Form)
    Dim seq() As Long
    If TabSequence(seq) = 0 Then
        KitToast "Nothing on the canvas takes the focus yet: add a text box, a button or another control first.", "info"
        Exit Sub
    End If
    RunDialog DZ_TABS
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
        m_field(j) = m_field(j + 1): m_anchor(j) = m_anchor(j + 1): m_tab(j) = m_tab(j + 1)
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
    rs!form_width = m_fw
    rs!form_height = m_fh
    rs!form_kind = m_formKind
    rs!record_source = IIf(Len(m_recSrc) = 0, Null, m_recSrc)
    rs!form_title = IIf(Len(m_title) = 0, Null, m_title)
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
                rs!item_field = IIf(Len(m_field(i)) = 0, Null, m_field(i))
                rs!item_anchor = IIf(Len(m_anchor(i)) = 0, Null, m_anchor(i))
                rs!item_tab = m_tab(i)
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

' Form_Load: the selected element's kind as the title (or "Form settings"),
' the rows it uses (moved up under each other, the dialog shrunk to fit)
' with its settings, and what the rows take.
Public Sub DesignerPropsLoad(dlg As Form)
    Dim i As Long, r As Long, y As Long, cap As String, inp As Control, dy As Long, first As String
    i = m_sel
    If i = 0 And Not m_propsForm Then Exit Sub
    If m_propsForm Then
        dlg.Caption = "Form settings"
        dlg("lblPTitle").Caption = "Form settings"
    Else
        dlg("lblPTitle").Caption = KindCaption(m_kind(i)) & " settings"
    End If
    y = dlg("lblP1").Top
    For r = 1 To DZ_ROWS
        cap = PropsCaption(r)
        If r <= 3 Then Set inp = dlg("txtP" & r) Else Set inp = dlg("cboP" & r)
        SetVisible dlg("lblP" & r), Len(cap) > 0
        SetVisible inp, Len(cap) > 0
        If Len(cap) > 0 Then
            dlg("lblP" & r).Caption = cap
            dlg("lblP" & r).Top = y
            inp.Top = y + px(18)
            y = y + px(64)
            If Len(first) = 0 Then first = inp.Name
        End If
    Next
    If m_propsForm Then
        dlg("txtP1").Value = m_title
        dlg("txtP2").Value = m_fw
        dlg("txtP3").Value = m_fh
        dlg("cboP4").RowSource = "page;dialog"
        dlg("cboP4").Value = m_formKind
        dlg("cboP5").RowSource = SourceList()
        dlg("cboP5").Value = IIf(Len(m_recSrc) = 0, Null, m_recSrc)
        dlg("lblPHint").Caption = "A page fills the app's window; a dialog opens over it. The form shows the first record " & _
            "of its record source: one with no records shows nothing. Elements' fields come from it."
    Else
        dlg("txtP1").Value = m_text(i)
        dlg("txtP2").Value = m_text2(i)
        dlg("txtP3").Value = m_text3(i)
        dlg("cboP4").RowSource = ChoiceList(m_kind(i))
        dlg("cboP4").Value = IIf(Len(m_choice(i)) = 0, Null, m_choice(i))
        dlg("cboP5").RowSource = """""" & FieldList(m_recSrc)  ' an empty first row unbinds
        dlg("cboP5").Value = IIf(Len(m_field(i)) = 0, Null, m_field(i))
        dlg("cboP6").RowSource = AnchorList(m_kind(i))
        dlg("cboP6").Value = IIf(Len(m_anchor(i)) = 0, "tl", m_anchor(i))
        dlg("lblPHint").Caption = Trim$(KindHint(m_kind(i)) & " " & FieldHint(m_kind(i)))
    End If
    dlg("lblPHint").Top = y
    dy = dlg("btnOK").Top - (y + px(60))
    If dy > 0 Then
        dlg("btnOK").Top = dlg("btnOK").Top - dy
        dlg("btnCancel").Top = dlg("btnOK").Top
        dlg.Section(acDetail).Height = dlg.Section(acDetail).Height - dy
        dlg.InsideHeight = dlg.Section(acDetail).Height
    End If
    CenterOnApp dlg
    dlg(first).SetFocus
End Sub

' OK: apply the dialog's settings to the selected element (DesignerSetProps)
' or the form (DesignerSetForm). False, after saying what is wrong, when they
' don't fit.
Public Function DesignerPropsSave(dlg As Form) As Boolean
    Dim msg As String
    If m_propsForm Then
        msg = DesignerSetForm(Nz(dlg("txtP1").Value, ""), Nz(dlg("txtP2").Value, ""), Nz(dlg("txtP3").Value, ""), _
            Nz(dlg("cboP4").Value, ""), Nz(dlg("cboP5").Value, ""))
    ElseIf m_sel = 0 Then
        DesignerPropsSave = True
        Exit Function
    Else
        msg = DesignerSetProps(m_sel, Nz(dlg("txtP1").Value, ""), Nz(dlg("txtP2").Value, ""), Nz(dlg("txtP3").Value, ""), _
            Nz(dlg("cboP4").Value, ""), Nz(dlg("cboP5").Value, ""), Nz(dlg("cboP6").Value, ""))
    End If
    If Len(msg) > 0 Then
        KitMsgBox msg, vbExclamation, "Form designer"
        Exit Function
    End If
    DesignerPropsSave = True
End Function

' Give item i these settings (the dialog's rows: three texts, a choice, the
' field it is bound to and its anchor; a field or anchor left out stays as
' it is), then keep it inside the sizes they allow (a segmented control's
' options, a rating's stars). Returns what is wrong with them instead,
' changing nothing; "" when applied. The caller redraws the item
' (DesignerRedraw).
Public Function DesignerSetProps(ByVal i As Long, ByVal t1 As String, ByVal t2 As String, ByVal t3 As String, _
        ByVal choice As String, Optional ByVal fld As Variant, Optional ByVal anc As Variant) As String
    Dim msg As String
    t1 = Left$(Trim$(t1), 255): t2 = Left$(Trim$(t2), 255): t3 = Left$(Trim$(t3), 255): choice = Trim$(choice)
    If IsMissing(fld) Then fld = m_field(i) Else fld = Trim$(Nz(fld, ""))
    If IsMissing(anc) Then anc = m_anchor(i) Else anc = Trim$(Nz(anc, ""))
    If anc = "tl" Then anc = ""
    msg = PropsProblem(m_kind(i), t1, t2, t3, choice)
    If Len(msg) = 0 Then msg = FieldProblem(m_kind(i), CStr(fld))
    If Len(msg) = 0 And Len(anc) > 0 And InStr(";" & AnchorList(m_kind(i)) & ";", ";" & anc & ";") = 0 Then _
        msg = "Pick an anchor from the list."
    If Len(msg) > 0 Then DesignerSetProps = msg: Exit Function
    m_text(i) = t1: m_text2(i) = t2: m_text3(i) = t3: m_choice(i) = choice
    m_field(i) = fld: m_anchor(i) = anc
    FitRect i, m_x(i), m_y(i), m_w(i), m_h(i)
    m_dirty = True
End Function

' Give the form these settings: a dialog's title ("" for the layout's name),
' its width and height in px, "page" or "dialog", and its record source (a
' table or query, "" for none). Returns what is wrong with them instead,
' changing nothing: a size outside DZ_MIN_FW x DZ_MIN_FH to the canvas, or
' smaller than the elements reach; a record source that isn't here, or
' lacks a field an element is bound to. "" when applied; the caller redraws
' the canvas (ShowCanvas).
Public Function DesignerSetForm(ByVal title As String, ByVal w As String, ByVal h As String, ByVal kind As String, _
        ByVal src As String) As String
    Dim i As Long, reachX As Long, reachY As Long, msg As String
    title = Left$(Trim$(title), 255): w = Trim$(w): h = Trim$(h): kind = Trim$(kind): src = Trim$(src)
    For i = 1 To m_n
        If m_x(i) + m_w(i) > reachX Then reachX = m_x(i) + m_w(i)
        If m_y(i) + m_h(i) > reachY Then reachY = m_y(i) + m_h(i)
    Next
    If kind <> "page" And kind <> "dialog" Then
        msg = "Pick page or dialog."
    ElseIf Not IsWholeNumber(w) Or Not IsWholeNumber(h) Then
        msg = "The width and height are whole numbers of px."
    ElseIf Val(w) < DZ_MIN_FW Or Val(h) < DZ_MIN_FH Or Val(w) > m_cw Or Val(h) > m_ch Then
        msg = "The form can be " & DZ_MIN_FW & " x " & DZ_MIN_FH & " px to " & m_cw & " x " & m_ch & " px, the canvas."
    ElseIf Val(w) < reachX Or Val(h) < reachY Then
        msg = "The elements reach " & reachX & " px across and " & reachY & " px down. Make the form at least that big, " & _
            "or move them first."
    ElseIf Len(src) > 0 And Not SourceExists(src) Then
        msg = "There is no table or query called """ & src & """ here."
    Else
        For i = 1 To m_n
            If Len(m_field(i)) > 0 Then
                If Not FieldExists(src, m_field(i)) Then
                    msg = IIf(Len(src) = 0, "Without a record source", """" & src & """ has no field """ & m_field(i) & """, and") & _
                        " a " & LCase$(KindCaption(m_kind(i))) & " is bound to """ & m_field(i) & """. " & _
                        "Change the field it is bound to first."
                    Exit For
                End If
            End If
        Next
    End If
    If Len(msg) > 0 Then DesignerSetForm = msg: Exit Function
    m_title = title: m_fw = CLng(w): m_fh = CLng(h): m_formKind = kind: m_recSrc = src
    m_dirty = True
End Function

' The caption of the settings dialog's row r for what it is showing (Form,
' or the selected element's kind), "" when it doesn't use that row.
Private Function PropsCaption(ByVal r As Long) As String
    If m_propsForm Then
        PropsCaption = Choose(r, "DIALOG TITLE (EMPTY FOR THE LAYOUT'S NAME)", "WIDTH (PX)", "HEIGHT (PX)", _
            "PAGE OR DIALOG", "RECORD SOURCE (EMPTY FOR NONE)", "")
    ElseIf m_sel > 0 Then
        PropsCaption = RowCaption(m_kind(m_sel), r)
    End If
End Function

' ---------------------------------------------------------------------------
' Tab order dialog (dlgDzTabs's code-behind)
' ---------------------------------------------------------------------------

' Form_Load: the elements that take the focus, in their tab order.
Public Sub DesignerTabsLoad(dlg As Form)
    Dim k As Long
    TabSequence m_tabSeq
    m_tabReading = True
    For k = 0 To UBound(m_tabSeq)
        If m_tab(m_tabSeq(k)) > 0 Then m_tabReading = False
    Next
    ShowTabs dlg, 0
    CenterOnApp dlg
    dlg("lstTabs").SetFocus
End Sub

' lstTabs's AfterUpdate: the buttons that fit the row picked.
Public Sub DesignerTabsPicked(dlg As Form)
    TabButtons dlg
End Sub

' Move up (delta -1) and Move down (1): swap the picked element with its
' neighbour. The order is no longer the reading order.
Public Sub DesignerTabsMove(dlg As Form, ByVal delta As Long)
    Dim k As Long, i As Long
    k = TabPicked(dlg)
    If k < 0 Or k + delta < 0 Or k + delta > UBound(m_tabSeq) Then Exit Sub
    i = m_tabSeq(k): m_tabSeq(k) = m_tabSeq(k + delta): m_tabSeq(k + delta) = i
    m_tabReading = False
    ShowTabs dlg, k + delta
End Sub

' Reading order: top to bottom, then left to right, as new elements follow.
Public Sub DesignerTabsReading(dlg As Form)
    TabSequence m_tabSeq, True
    m_tabReading = True
    ShowTabs dlg, 0
End Sub

' OK: give each element its place in the list (0, the reading order, when
' it is the reading order, so elements placed later fall into it too).
Public Sub DesignerTabsSave(dlg As Form)
    Dim k As Long
    For k = 0 To UBound(m_tabSeq)
        m_tab(m_tabSeq(k)) = IIf(m_tabReading, 0, k + 1)
    Next
    m_dirty = True
End Sub

' Fill the list from m_tabSeq and pick the row at position k.
Private Sub ShowTabs(dlg As Form, ByVal k As Long)
    Dim i As Long, rows As String
    For i = 0 To UBound(m_tabSeq)
        rows = rows & IIf(i > 0, ";", "") & i & ";" & ValueItem((i + 1) & ".  " & TabCaption(m_tabSeq(i)))
    Next
    dlg("lstTabs").RowSource = rows
    dlg("lstTabs").Value = CStr(k)                 ' the bound column is the position
    TabButtons dlg
End Sub

Private Sub TabButtons(dlg As Form)
    Dim k As Long
    k = TabPicked(dlg)
    SetEnabled dlg("btnUp"), k > 0
    SetEnabled dlg("btnDown"), k >= 0 And k < UBound(m_tabSeq)
    SetEnabled dlg("btnReading"), Not m_tabReading
End Sub

' The position of the row picked in the tab order list, -1 for none.
Private Function TabPicked(dlg As Form) As Long
    TabPicked = -1
    If Not IsNull(dlg("lstTabs").Value) Then TabPicked = CLng(dlg("lstTabs").Value)
End Function

' Item i in the tab order list: its kind, what it says or is bound to, and where it is.
Private Function TabCaption(ByVal i As Long) As String
    Dim what As String
    what = m_field(i)
    If Len(what) = 0 Then
        Select Case m_kind(i)
            Case "empty": what = m_text3(i)
            Case "button", "segmented", "stepper", "search": what = m_text(i)
        End Select
    End If
    TabCaption = KindCaption(m_kind(i)) & IIf(Len(what) > 0, ": " & OneLine(what), "") & "   (" & m_x(i) & ", " & m_y(i) & ")"
End Function

' Item numbers of the elements that take the focus, in tab order (TabRank),
' into seq; returns how many. reading = True ignores the places given.
Private Function TabSequence(seq() As Long, Optional ByVal reading As Boolean = False) As Long
    Dim i As Long, n As Long, keys() As Double, idx() As Long, items() As Long
    ReDim keys(0 To DZ_MAX): ReDim idx(0 To DZ_MAX): ReDim items(0 To DZ_MAX)
    For i = 1 To m_n
        If Focusable(m_kind(i), m_text3(i)) Then
            items(n) = i
            keys(n) = TabRank(IIf(reading, 0, m_tab(i)), m_x(i), m_y(i))
            idx(n) = n
            n = n + 1
        End If
    Next
    SortByKey idx, keys, n
    If n > 0 Then
        ReDim seq(0 To n - 1)
        For i = 0 To n - 1
            seq(i) = items(idx(i))
        Next
    Else
        Erase seq
    End If
    TabSequence = n
End Function

' Where an element falls in the tab order: by the place it was given
' (tabNo > 0), then by reading order, top to bottom, then left to right.
Private Function TabRank(ByVal tabNo As Long, ByVal x As Long, ByVal y As Long) As Double
    If tabNo > 0 Then TabRank = tabNo Else TabRank = 1000000# + y * 100000# + x
End Function

' Sort idx(0..n-1) by keys(idx(i)), keeping equal keys in their order.
Private Sub SortByKey(idx() As Long, keys() As Double, ByVal n As Long)
    Dim i As Long, j As Long, t As Long
    For i = 1 To n - 1
        t = idx(i)
        j = i
        Do While j > 0
            If keys(idx(j - 1)) <= keys(t) Then Exit Do
            idx(j) = idx(j - 1)
            j = j - 1
        Loop
        idx(j) = t
    Next
End Sub

' ---------------------------------------------------------------------------
' Generate
' ---------------------------------------------------------------------------

' btnDzGenerate: save the layout if it has changes, build it as a real form
' and show it (dzPreview), and write its build procedure and code-behind stub
' to <database folder>\designer as Build_<form>.txt and <form>.vba. Refused,
' with the reason, when its record source or a bound field is gone.
Public Sub DesignerGenerate(frm As Form)
    Dim nm As String, folder As String, id As Long, msg As String
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
    msg = LayoutProblem(id)
    If Len(msg) > 0 Then
        KitMsgBox msg, vbExclamation, "Form designer"
        Exit Sub
    End If
    nm = LayoutFormName(id)
    folder = CurrentProject.Path & "\designer"
    EnsureFolder folder
    WriteTextFile folder & "\Build_" & nm & ".txt", DesignerCode(id)
    WriteTextFile folder & "\" & nm & ".vba", DesignerStub(id)
    DesignerPreview id
    KitToast "Preview built. Its code is in the designer folder: Build_" & nm & ".txt and " & nm & ".vba.", "success", 6000
End Sub

' The form a layout generates: "pg" (or "dlg" for a dialog) and the name's
' letters and digits, each word capitalised ("review navy 2" gives
' pgReviewNavy2; no letters, pgLayout).
Public Function DesignerFormName(ByVal layoutName As String, Optional ByVal kind As String = "page") As String
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
    DesignerFormName = Left$(IIf(kind = "dialog", "dlg", "pg") & out, 64)
End Function

' Layout id as a Build_<form> procedure for an app's build module: NewForm
' (a page) or NewDialog at the form's size, with its record source; one kit
' builder call per element back to front, each anchored on a page; the tab
' order; SaveForm. The elements that call the form's code name a handler
' after themselves (seg1Changed, step1Step, search1Search: see
' DesignerStub). Bound elements take their field as "[field]".
Public Function DesignerCode(ByVal id As Long) As String
    Dim db As DAO.Database, rs As DAO.Recordset, nm As String, ln As String, s As String, body As String, title As String
    Dim isPage As Boolean, src As String, anc As String, h As Integer, v As Integer, anchored As Boolean
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM kit_layout WHERE id=" & id, dbOpenSnapshot)
    If rs.EOF Then Err.Raise 5, , "No form designer layout " & id
    title = OneLine(rs!layout_name)
    isPage = (Nz(rs!form_kind, "page") <> "dialog")
    nm = DesignerFormName(title, IIf(isPage, "page", "dialog"))
    src = Nz(rs!record_source, "")
    s = "' " & nm & ", built from the form designer layout """ & title & """ (" & Format$(Now(), "yyyy-mm-dd hh:nn") & ")." & vbCrLf & _
        "' Call it from the app's build, and keep " & nm & ".vba, its code-behind, in src\forms." & vbCrLf & _
        "Private Sub Build_" & nm & "()" & vbCrLf
    If isPage Then
        body = "    NewForm " & VbText(src) & ", 0, px(" & Nz(rs!form_width, 0) & "), px(" & Nz(rs!form_height, 0) & ")" & vbCrLf
    Else
        body = "    NewDialog " & VbText(DialogTitle(rs)) & ", px(" & Nz(rs!form_width, 0) & "), px(" & Nz(rs!form_height, 0) & ")" & vbCrLf
        If Len(src) > 0 Then body = body & "    F.RecordSource = " & VbText(src) & vbCrLf
    End If
    ' Tab past the last control stays on the record, not on to the next one.
    If Len(src) > 0 Then body = body & "    F.Cycle = 1" & vbCrLf
    rs.Close
    If Len(StepperNames(id)) > 0 Then body = body & "    Ev F, ""OnLoad""" & vbCrLf     ' Form_Load paints the steppers (DesignerStub)
    Set rs = db.OpenRecordset(ItemsSql(id), dbOpenSnapshot)
    Do Until rs.EOF
        anc = IIf(isPage, Nz(rs!item_anchor, ""), "")
        ln = ItemCall(rs, anc)
        If Len(ln) > 0 Then
            If Len(anc) > 0 And rs!kind <> "card" Then
                ' Every part of the element moves (or stretches) with the window.
                AnchorOf anc, h, v
                body = body & "    k = BuiltNames()" & vbCrLf & "    " & ln & vbCrLf & "    AnchorBuilt k, " & h & ", " & v & vbCrLf
                anchored = True
            Else
                body = body & "    " & ln & vbCrLf
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    ln = TabOrderCall(id)
    If Len(ln) > 0 Then body = body & "    " & ln & vbCrLf
    If anchored Then s = s & "    Dim k As String" & vbCrLf
    DesignerCode = s & body & "    SaveForm """ & nm & """" & vbCrLf & "End Sub" & vbCrLf
End Function

' Layout id's code-behind to start from: an empty handler, with the kit's
' error handling, for each button's click, each empty state's button and
' each segmented control, toggle, slider, rating, stepper and search box
' that DesignerCode names, for the app to fill in.
Public Function DesignerStub(ByVal id As Long) As String
    Dim db As DAO.Database, rs As DAO.Recordset, nm As String, s As String, title As String, b As Variant
    Dim sig As String, proc As String, what As String, isDialog As Boolean
    Set db = CurrentDb
    title = OneLine(Nz(DLookup("layout_name", "kit_layout", "id=" & id), ""))
    isDialog = (Nz(DLookup("form_kind", "kit_layout", "id=" & id), "page") = "dialog")
    nm = LayoutFormName(id)
    s = "Option Compare Database" & vbCrLf & "Option Explicit" & vbCrLf & vbCrLf & _
        "' " & nm & " (Build_" & nm & ", from the form designer layout """ & title & """)." & vbCrLf
    If isDialog Then
        s = s & "' A dialog: open it with RunDialog, which is True when the dialog sets" & vbCrLf & _
            "' TempVars(""DialogResult"") = ""ok"" before it closes." & vbCrLf
    End If
    If Len(StepperNames(id)) > 0 Then
        s = s & vbCrLf & "' A stepper is built in the build's palette: SetStep paints it in the user's." & vbCrLf & _
            "Private Sub Form_Load()" & vbCrLf & "    On Error GoTo Err_Handler" & vbCrLf
        For Each b In Split(StepperNames(id), ";")
            s = s & "    SetStep Me, " & VbText(CStr(b)) & ", CurrentStep(Me, " & VbText(CStr(b)) & ")" & vbCrLf
        Next
        s = s & "    Exit Sub" & vbCrLf & "Err_Handler:" & vbCrLf & "    HandleError """ & nm & ".Form_Load""" & vbCrLf & _
            "End Sub" & vbCrLf
    End If
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
' toggle switches, a slider slides) but call no code. A page's preview can
' be resized, to try its anchors; a dialog's can't. A bound preview shows
' the first record read-only, so trying it never changes the data.
Public Sub DesignerPreview(ByVal id As Long)
    Dim db As DAO.Database, rs As DAO.Recordset, title As String, fw As Long, fh As Long, isPage As Boolean, src As String
    Dim names As String
    If IsFormOpen(DZ_PREVIEW) Then DoCmd.Close acForm, DZ_PREVIEW, acSaveNo
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM kit_layout WHERE id=" & id, dbOpenSnapshot)
    If rs.EOF Then Err.Raise 5, , "No form designer layout " & id
    isPage = (Nz(rs!form_kind, "page") <> "dialog")
    title = IIf(isPage, OneLine(rs!layout_name), DialogTitle(rs))
    fw = Nz(rs!form_width, 0): fh = Nz(rs!form_height, 0): src = Nz(rs!record_source, "")
    rs.Close
    On Error GoTo Fail
    If isPage Then NewForm src, 0, px(fw), px(fh) Else NewDialog title, px(fw), px(fh)
    F.RecordSource = src
    F.Caption = "Preview: " & title
    F.PopUp = True
    F.Modal = False
    F.AutoCenter = True
    F.AutoResize = True
    F.BorderStyle = IIf(isPage, 2, 3)          ' a page's preview resizes, to try its anchors
    F.MinMaxButtons = 0
    F.ControlBox = True
    F.CloseButton = True
    If Len(src) > 0 Then
        F.AllowEdits = False
        F.Cycle = 1                                ' Tab stays on the record, as DesignerCode's form
    End If
    Set rs = db.OpenRecordset(ItemsSql(id), dbOpenSnapshot)
    Do Until rs.EOF
        BuildItem rs, IIf(isPage, Nz(rs!item_anchor, ""), "")
        rs.MoveNext
    Loop
    rs.Close
    names = TabNames(id)
    If Len(names) > 0 Then TabOrder Split(names, ";")
    If Len(StepperNames(id)) > 0 Then F.OnLoad = "=DesignerPreviewLoad([Form]," & VbText(StepperNames(id)) & ")"
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

' dzPreview's Form_Load (its OnLoad expression; the preview has no code):
' paint each stepper (names, ;-separated) in the current palette, as an
' app's Form_Load does with SetStep. The builder painted them in the
' palette it ran under, which leaves dark text on a dark palette.
Public Function DesignerPreviewLoad(frm As Form, ByVal names As String)
    Dim nm As Variant
    On Error GoTo Fail
    For Each nm In Split(names, ";")
        SetStep frm, CStr(nm), CurrentStep(frm, CStr(nm))
    Next
    Exit Function
Fail:
    HandleError "modKitDesigner.DesignerPreviewLoad", names
End Function

' What stops layout id generating: its record source gone, or a field an
' element is bound to gone from it ("" when nothing does).
Private Function LayoutProblem(ByVal id As Long) As String
    Dim rs As DAO.Recordset, src As String
    src = Nz(DLookup("record_source", "kit_layout", "id=" & id), "")
    If Len(src) > 0 And Not SourceExists(src) Then
        LayoutProblem = "The record source """ & src & """ isn't here any more. Pick another under Form."
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset(ItemsSql(id), dbOpenSnapshot)
    Do Until rs.EOF
        If Len(Nz(rs!item_field, "")) > 0 Then
            If Not FieldExists(src, rs!item_field) Then
                LayoutProblem = "A " & LCase$(KindCaption(rs!kind)) & " is bound to """ & rs!item_field & """, which " & _
                    IIf(Len(src) = 0, "needs a record source.", """" & src & """ doesn't have. Bind it to another field.")
                Exit Do
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
End Function

' The form layout id generates (DesignerFormName of its name and kind).
Private Function LayoutFormName(ByVal id As Long) As String
    LayoutFormName = DesignerFormName(OneLine(Nz(DLookup("layout_name", "kit_layout", "id=" & id), "")), _
        Nz(DLookup("form_kind", "kit_layout", "id=" & id), "page"))
End Function

' A dialog layout's title bar: its title, or else its name.
Private Function DialogTitle(rs As DAO.Recordset) As String
    DialogTitle = OneLine(Nz(rs!form_title, ""))
    If Len(DialogTitle) = 0 Then DialogTitle = OneLine(rs!layout_name)
End Function

' Layout id's steppers' control names, ;-separated ("" for none).
Private Function StepperNames(ByVal id As Long) As String
    Dim rs As DAO.Recordset, out As String
    Set rs = CurrentDb.OpenRecordset(ItemsSql(id), dbOpenSnapshot)
    Do Until rs.EOF
        If rs!kind = "stepper" Then out = out & IIf(Len(out) > 0, ";", "") & ItemName(rs)
        rs.MoveNext
    Loop
    rs.Close
    StepperNames = out
End Function

' Layout id's elements that take the focus, in tab order (TabRank), as
' control names ;-separated ("" for none).
Private Function TabNames(ByVal id As Long) As String
    Dim rs As DAO.Recordset, nms() As String, keys() As Double, idx() As Long, n As Long, i As Long, out As String
    Set rs = CurrentDb.OpenRecordset(ItemsSql(id), dbOpenSnapshot)
    ReDim nms(0 To DZ_MAX): ReDim keys(0 To DZ_MAX): ReDim idx(0 To DZ_MAX)
    Do Until rs.EOF
        If Focusable(Nz(rs!kind, ""), Nz(rs!item_text3, "")) And n <= DZ_MAX Then
            nms(n) = ItemName(rs)
            keys(n) = TabRank(Nz(rs!item_tab, 0), Nz(rs!x, 0), Nz(rs!y, 0))
            idx(n) = n
            n = n + 1
        End If
        rs.MoveNext
    Loop
    rs.Close
    SortByKey idx, keys, n
    For i = 0 To n - 1
        out = out & IIf(i > 0, ";", "") & nms(idx(i))
    Next
    TabNames = out
End Function

' The TabOrder call for layout id, 8 names to a line ("" when nothing takes
' the focus).
Private Function TabOrderCall(ByVal id As Long) As String
    Dim names() As String, i As Long, s As String
    If Len(TabNames(id)) = 0 Then Exit Function
    names = Split(TabNames(id), ";")
    s = "TabOrder Array("
    For i = 0 To UBound(names)
        If i > 0 Then s = s & IIf(i Mod 8 = 0, ", _" & vbCrLf & "        ", ", ")
        s = s & VbText(names(i))
    Next
    TabOrderCall = s & ")"
End Function

Private Function ItemsSql(ByVal id As Long) As String
    ItemsSql = "SELECT item_name, kind, item_text, item_text2, item_text3, item_choice, item_field, item_anchor, item_tab, " & _
        "x, y, w, h FROM kit_layout_item WHERE layout_id=" & id & " ORDER BY item_order"
End Function

' The row's element as a kit builder call (DesignerCode); "" for a kind the
' designer doesn't know. A card takes anc (AnchorOf, "" for none) itself.
' BuildItem makes the same call.
Private Function ItemCall(rs As DAO.Recordset, ByVal anc As String) As String
    Dim nm As String, t1 As String, t2 As String, t3 As String, ch As String, n As String, src As String
    Dim x As String, y As String, w As String, h As String, at As String, parts() As String, i As Long
    Dim ah As Integer, av As Integer
    nm = ItemName(rs): n = VbText(nm)
    t1 = Nz(rs!item_text, ""): t2 = Nz(rs!item_text2, ""): t3 = Nz(rs!item_text3, ""): ch = Nz(rs!item_choice, "")
    src = FieldSource(rs)
    x = "px(" & Nz(rs!x, 0) & ")": y = "px(" & Nz(rs!y, 0) & ")": w = "px(" & Nz(rs!w, 0) & ")": h = "px(" & Nz(rs!h, 0) & ")"
    at = ", " & x & ", " & y & ", " & w & ", " & h
    Select Case Nz(rs!kind, "")
        Case "card"
            ItemCall = "Card " & n & ", " & VbText(t1) & at
            If Len(anc) > 0 Then
                AnchorOf anc, ah, av
                ItemCall = ItemCall & ", """", " & ah & ", " & av
            End If
        Case "button": ItemCall = "Btn " & n & ", " & VbText(t1) & at & ", " & VbText(ButtonStyle(ch))
        Case "label": ItemCall = "Lbl " & n & ", " & VbText(t1) & at
        Case "progress": ItemCall = "Progress " & n & ", " & VbText(IIf(Len(src) > 0, src, PercentText(t1))) & at
        Case "text": ItemCall = "Txt " & n & ", " & VbText(src) & at
        Case "date": ItemCall = "DateBox " & n & ", " & VbText(src) & ", " & x & ", " & y & ", " & h
        Case "secret": ItemCall = "SecretBox " & n & ", " & VbText(src) & at
        Case "toggle": ItemCall = "Toggle " & n & ", " & VbText(src) & at & ", " & VbText(nm & "Changed")
        Case "segmented"
            ItemCall = "Segmented " & n & ", " & VbText(src) & ", " & x & ", " & y & ", " & w & ", " & VbText(SegmentItems(t1)) & _
                ", " & h & ", " & VbText(nm & "Changed")
        Case "slider"
            ItemCall = "Slider " & n & ", " & VbText(src) & ", " & x & ", " & y & ", " & w & ", " & NumText(t1) & ", " & _
                NumText(t2) & ", " & NumText(t3) & ", " & VbText(nm & "Changed")
        Case "rating"
            ItemCall = "Rating " & n & ", " & VbText(src) & ", " & x & ", " & y & ", " & CLng(Val(t1)) & ", px(24), False, " & _
                VbText(nm & "Changed")
        Case "stepper"
            ItemCall = "Stepper " & n & ", " & x & ", " & y & ", " & w & ", " & CLng(Val(t2)) & ", " & VbText(nm & "Step")
            parts = Split(t1, ";")
            For i = 0 To UBound(parts)
                ItemCall = ItemCall & ", " & VbText(Trim$(parts(i)))
            Next
        Case "search": ItemCall = "SearchBox " & n & ", " & x & ", " & y & ", " & w & ", " & VbText(nm & "Search") & ", " & VbText(t1) & ", " & h
        Case "stat"
            ItemCall = "StatTile " & n & ", " & VbText(t1) & ", " & VbText(IIf(Len(src) > 0, src, "=" & NumText(t2))) & ", " & _
                x & ", " & y & ", " & w & ", " & h & ", """", " & VbText(t3)
        Case "alert": ItemCall = "Alert " & n & at & ", " & VbText(AlertVariant(ch)) & ", " & VbText(t1) & ", " & VbText(t2)
        Case "avatar": ItemCall = "Avatar " & n & ", " & VbText(IIf(Len(src) > 0, src, AvatarSource(t1))) & ", " & x & ", " & y & ", " & w
        Case "empty"
            ItemCall = "EmptyState " & n & ", " & x & ", " & y & ", " & w & ", " & VbText(DZ_EMPTY_GLYPH) & ", " & VbText(t1) & ", " & _
                VbText(t2) & ", " & VbText(t3) & ", """", True"
    End Select
End Function

' The row's element built on the form under construction (DesignerPreview):
' the call ItemCall writes, with no handlers, buttons that call no code, and
' anchored by anc (AnchorOf, "" for none) as DesignerCode anchors it.
Private Sub BuildItem(rs As DAO.Recordset, ByVal anc As String)
    Dim nm As String, t1 As String, t2 As String, t3 As String, ch As String, c As Control, parts() As String, src As String
    Dim x As Long, y As Long, w As Long, h As Long, ah As Integer, av As Integer, before As String
    nm = ItemName(rs)
    t1 = Nz(rs!item_text, ""): t2 = Nz(rs!item_text2, ""): t3 = Nz(rs!item_text3, ""): ch = Nz(rs!item_choice, "")
    src = FieldSource(rs)
    x = px(Nz(rs!x, 0)): y = px(Nz(rs!y, 0)): w = px(Nz(rs!w, 0)): h = px(Nz(rs!h, 0))
    If Len(anc) > 0 Then AnchorOf anc, ah, av
    If Len(anc) > 0 And rs!kind <> "card" Then before = BuiltNames()
    Select Case Nz(rs!kind, "")
        Case "card": Card nm, t1, x, y, w, h, "", ah, av
        Case "button"
            Set c = Btn(nm, t1, x, y, w, h, ButtonStyle(ch))
            c.OnClick = ""
        Case "label": Lbl nm, t1, x, y, w, h
        Case "progress": Progress nm, IIf(Len(src) > 0, src, PercentText(t1)), x, y, w, h
        Case "text": Txt nm, src, x, y, w, h
        Case "date": DateBox nm, src, x, y, h
        Case "secret": SecretBox nm, src, x, y, w, h
        Case "toggle": Toggle nm, src, x, y, w, h
        Case "segmented": Segmented nm, src, x, y, w, SegmentItems(t1), h
        Case "slider": Slider nm, src, x, y, w, Val(t1), Val(t2), Val(t3)
        Case "rating": Rating nm, src, x, y, CLng(Val(t1)), px(24)
        Case "stepper"
            parts = Split(t1, ";")
            StepperFrom nm, x, y, w, CLng(Val(t2)), parts
        Case "search": SearchBox nm, x, y, w, "", t1, h
        Case "stat": StatTile nm, t1, IIf(Len(src) > 0, src, "=" & NumText(t2)), x, y, w, h, "", t3
        Case "alert": Alert nm, x, y, w, h, AlertVariant(ch), t1, t2
        Case "avatar": Avatar nm, IIf(Len(src) > 0, src, AvatarSource(t1)), x, y, w
        Case "empty"
            EmptyState nm, x, y, w, DZ_EMPTY_GLYPH, t1, t2, t3, "", True
            On Error Resume Next                 ' no button without a caption
            F.Controls(nm & "_b").OnClick = ""
            On Error GoTo 0
    End Select
    If Len(before) > 0 Then AnchorBuilt before, ah, av
End Sub

' The row's bound field as a control source, "[field]" ("" when unbound).
Private Function FieldSource(rs As DAO.Recordset) As String
    If Len(Nz(rs!item_field, "")) > 0 Then FieldSource = "[" & rs!item_field & "]"
End Function

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
' settings, form and binding columns a table from an earlier version lacks.
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
    AddColumn db, "kit_layout_item", "item_field", "TEXT(64)"
    AddColumn db, "kit_layout_item", "item_anchor", "TEXT(10)"
    AddColumn db, "kit_layout_item", "item_tab", "LONG"
    AddColumn db, "kit_layout", "form_kind", "TEXT(10)"
    AddColumn db, "kit_layout", "record_source", "TEXT(64)"
    AddColumn db, "kit_layout", "form_title", "TEXT(255)"
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

' Item i's binding, anchor and tab place as "field|anchor|tab", and the
' form as "kind wxh source|title" (smoke tests).
Public Function DesignerItemData(ByVal i As Long) As String
    DesignerItemData = m_field(i) & "|" & m_anchor(i) & "|" & m_tab(i)
End Function

Public Function DesignerFormData() As String
    DesignerFormData = m_formKind & " " & m_fw & "x" & m_fh & " " & m_recSrc & "|" & m_title
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
        KindDefault(kind, 3), KindDefault(kind, 4), "", "", 0)
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

' Where a kind element centred on (x, y) px lands: snapped, on the form.
Private Sub Landing(ByVal kind As String, ByVal x As Long, ByVal y As Long, ByRef lx As Long, ByRef ly As Long)
    lx = Clamp(Snap(x - DefaultW(kind) \ 2), 0, m_fw - DefaultW(kind))
    ly = Clamp(Snap(y - DefaultH(kind) \ 2), 0, m_fh - DefaultH(kind))
End Sub

' Append an item in the first free slot of its pool; 0 when the pool is full.
' The rectangle is kept to the sizes its kind and settings allow, inside the form.
Private Function AddItem(ByVal kind As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
        ByVal t1 As String, ByVal t2 As String, ByVal t3 As String, ByVal choice As String, ByVal fld As String, _
        ByVal anc As String, ByVal tabNo As Long) As Long
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
    m_field(m_n) = fld: m_anchor(m_n) = anc: m_tab(m_n) = tabNo
    FitRect m_n, x, y, w, h
    AddItem = m_n
End Function

' Move or resize item i, keeping it on the form, and redraw it.
Private Sub SetRect(frm As Form, ByVal i As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    FitRect i, x, y, w, h
    Render frm, i
    m_dirty = True
    ShowSelection frm
End Sub

' Item i's rectangle: within the sizes its kind and settings allow (an avatar
' stays round), and inside the form.
Private Sub FitRect(ByVal i As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    w = Clamp(w, MinW(i), MaxW(i))
    If m_kind(i) = "avatar" Then h = w Else h = Clamp(h, MinH(i), MaxH(i))
    m_w(i) = w
    m_h(i) = h
    m_x(i) = Clamp(x, 0, m_fw - w)
    m_y(i) = Clamp(y, 0, m_fh - h)
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
    m_formKind = "page"
    m_fw = m_cw: m_fh = m_ch
    m_recSrc = "": m_title = ""
    frm("cboDzOpen").Value = Null
    frm("txtDzName").Value = Null
    ShowCanvas frm
    ShowSelection frm
End Sub

' Show the saved layout id (elements of a kind this version doesn't know, or
' past a pool's size, are left out, with a toast).
Private Sub LoadLayout(frm As Form, ByVal id As Long)
    Dim db As DAO.Database, rs As DAO.Recordset, nm As String, skipped As Long, i As Long
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM kit_layout WHERE id=" & id, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        ClearLayout frm
        Exit Sub
    End If
    nm = rs!layout_name
    ClearLayout frm
    ' A layout saved before the form settings is a page the canvas's size.
    m_formKind = IIf(Nz(rs!form_kind, "page") = "dialog", "dialog", "page")
    m_fw = Clamp(Nz(rs!form_width, m_cw), DZ_MIN_FW, m_cw)
    m_fh = Clamp(Nz(rs!form_height, m_ch), DZ_MIN_FH, m_ch)
    m_recSrc = Nz(rs!record_source, "")
    m_title = Nz(rs!form_title, "")
    rs.Close
    ShowCanvas frm
    Set rs = db.OpenRecordset("SELECT kind, item_text, item_text2, item_text3, item_choice, item_field, item_anchor, item_tab, " & _
        "x, y, w, h FROM kit_layout_item WHERE layout_id=" & id & " ORDER BY item_order", dbOpenSnapshot)
    Do Until rs.EOF
        If Len(KindPool(Nz(rs!kind, ""))) = 0 Then
            skipped = skipped + 1
        Else
            i = AddItem(rs!kind, Nz(rs!x, 0), Nz(rs!y, 0), Nz(rs!w, 0), Nz(rs!h, 0), Nz(rs!item_text, ""), _
                Nz(rs!item_text2, ""), Nz(rs!item_text3, ""), Nz(rs!item_choice, ""), Nz(rs!item_field, ""), _
                Nz(rs!item_anchor, ""), Nz(rs!item_tab, 0))
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

' What an outlined box on the canvas says it will be (a bound box, its field).
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
        Case "toggle": DescText = "On"           ' "Toggle" doesn't fit its 44px
        Case "text", "date", "secret"
            If Len(m_field(i)) > 0 Then DescText = "[" & m_field(i) & "]" Else DescText = KindCaption(m_kind(i))
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

' The form's outline on the canvas: its size, in a dialog's colour or a page's.
Private Sub ShowCanvas(frm As Form)
    Dim g As Control
    Set g = frm("dzGlass")
    MoveCtl frm("dzCanvas"), g.Left, g.Top, px(m_fw), px(m_fh)
    frm("dzCanvas").BackColor = Th(IIf(m_formKind = "dialog", CLR_DIALOG, CLR_BG))
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

' The caption of the settings dialog's row r (1-3 text, 4 a list, 5 the
' field, 6 the anchor) for a kind, "" when it doesn't use that row. Anchors
' are for a page: a dialog doesn't resize.
Private Function RowCaption(ByVal kind As String, ByVal r As Long) As String
    Dim caps As Variant
    If r = 5 Then
        If BindKind(kind) Then RowCaption = "FIELD (EMPTY FOR NONE)"
        Exit Function
    ElseIf r = 6 Then
        If m_formKind <> "dialog" And Len(KindPool(kind)) > 0 Then RowCaption = "ANCHOR (AS THE WINDOW RESIZES)"
        Exit Function
    End If
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

' True for a kind that shows or takes a value, so it can be bound to a field.
Private Function BindKind(ByVal kind As String) As Boolean
    Select Case kind
        Case "progress", "text", "date", "secret", "toggle", "segmented", "slider", "rating", "stat", "avatar": BindKind = True
    End Select
End Function

' True for a kind that can stretch with the window: one control, or a card
' (which anchors its own parts).
Private Function StretchKind(ByVal kind As String) As Boolean
    Select Case kind
        Case "card", "label", "text", "button": StretchKind = True
    End Select
End Function

' True for an element that takes the focus (an empty state only with its
' button, text3).
Private Function Focusable(ByVal kind As String, ByVal text3 As String) As Boolean
    Select Case kind
        Case "button", "text", "date", "secret", "toggle", "segmented", "slider", "rating", "stepper", "search": Focusable = True
        Case "empty": Focusable = (Len(text3) > 0)
    End Select
End Function

' The anchors a kind can take, as row 6's list: the key, then what it shows.
Private Function AnchorList(ByVal kind As String) As String
    AnchorList = "tl;Top left (stays put);tr;Top right;bl;Bottom left;br;Bottom right"
    If StretchKind(kind) Then AnchorList = AnchorList & ";across;Top - stretches across;both;Stretches across and down"
End Function

' An anchor key as Anchor's h and v (0 left or top, 1 right or bottom, 2 both).
Private Sub AnchorOf(ByVal key As String, ByRef h As Integer, ByRef v As Integer)
    Select Case key
        Case "tr": h = 1: v = 0
        Case "bl": h = 0: v = 1
        Case "br": h = 1: v = 1
        Case "across": h = 2: v = 0
        Case "both": h = 2: v = 2
        Case Else: h = 0: v = 0
    End Select
End Sub

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

' What binding a kind to a field does, under KindHint ("" for a kind that
' can't be bound).
Private Function FieldHint(ByVal kind As String) As String
    If Not BindKind(kind) Then Exit Function
    If Len(m_recSrc) = 0 Then
        FieldHint = "To bind a field, pick the form's record source under Form first."
        Exit Function
    End If
    Select Case kind
        Case "segmented": FieldHint = "Bound, the field holds the option's key: its caption's letters and digits, in lower case."
        Case "progress": FieldHint = "Bound, the field gives the percent."
        Case "stat": FieldHint = "Bound, the field gives the value."
        Case "avatar": FieldHint = "Bound, the initials come from the field."
        Case Else: FieldHint = "Bound, it shows the field and saves what is entered."
    End Select
End Function

' What is wrong with binding a kind to field fld ("" for none) on this
' form, "" when nothing is.
Private Function FieldProblem(ByVal kind As String, ByVal fld As String) As String
    If Len(fld) = 0 Then Exit Function
    If Not BindKind(kind) Then
        FieldProblem = "A " & LCase$(KindCaption(kind)) & " can't be bound to a field."
    ElseIf Len(m_recSrc) = 0 Then
        FieldProblem = "Pick the form's record source under Form first."
    ElseIf Not FieldExists(m_recSrc, fld) Then
        FieldProblem = """" & m_recSrc & """ has no field """ & fld & """."
    End If
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
        Case Else: MaxW = m_fw
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
        Case "card", "label": MaxH = m_fh
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

' A value-list item: quoted, so it may hold ; , or ".
Private Function ValueItem(ByVal s As String) As String
    ValueItem = """" & Replace(s, """", """""") & """"
End Function

' The tables and queries a form can take as its record source, as a value
' list: not the system's, temporary ones, action queries or the designer's
' own tables.
Private Function SourceList() As String
    Dim db As DAO.Database, td As DAO.TableDef, qd As DAO.QueryDef, out As String
    Set db = CurrentDb
    For Each td In db.TableDefs
        If SourceShown(td.Name) And (td.Attributes And dbSystemObject) = 0 Then out = out & ";" & ValueItem(td.Name)
    Next
    For Each qd In db.QueryDefs
        If SourceShown(qd.Name) And (qd.Type = dbQSelect Or qd.Type = dbQSetOperation) Then out = out & ";" & ValueItem(qd.Name)
    Next
    SourceList = Mid$(out, 2)
End Function

Private Function SourceShown(ByVal nm As String) As Boolean
    SourceShown = Not (nm Like "MSys*" Or nm Like "USys*" Or nm Like "~*" Or nm Like "kit_layout*")
End Function

' True when src is a table or query SourceList offers.
Private Function SourceExists(ByVal src As String) As Boolean
    Dim db As DAO.Database, td As DAO.TableDef, qd As DAO.QueryDef
    If Len(src) = 0 Or Not SourceShown(src) Then Exit Function
    Set db = CurrentDb
    For Each td In db.TableDefs
        If StrComp(td.Name, src, vbTextCompare) = 0 Then SourceExists = ((td.Attributes And dbSystemObject) = 0): Exit Function
    Next
    For Each qd In db.QueryDefs
        If StrComp(qd.Name, src, vbTextCompare) = 0 Then SourceExists = (qd.Type = dbQSelect Or qd.Type = dbQSetOperation): Exit Function
    Next
End Function

' The fields of table or query src, ;-separated after a ; each, as items of
' a value list ("" when it isn't here).
Private Function FieldList(ByVal src As String) As String
    Dim f As Variant
    For Each f In SourceFields(src)
        FieldList = FieldList & ";" & ValueItem(CStr(f))
    Next
End Function

Private Function FieldExists(ByVal src As String, ByVal fld As String) As Boolean
    Dim f As Variant
    For Each f In SourceFields(src)
        If StrComp(CStr(f), fld, vbTextCompare) = 0 Then FieldExists = True: Exit Function
    Next
End Function

' The names of table or query src's fields (none when it isn't here, or a
' query can't be read).
Private Function SourceFields(ByVal src As String) As Variant
    Dim db As DAO.Database, td As DAO.TableDef, qd As DAO.QueryDef, f As DAO.Field, out As String
    SourceFields = Array()
    If Not SourceExists(src) Then Exit Function
    Set db = CurrentDb
    On Error GoTo Done                          ' a query whose table is gone has no fields to read
    If TableExists(db, src) Then
        Set td = db.TableDefs(src)
        For Each f In td.Fields
            out = out & vbNullChar & f.Name
        Next
    Else
        Set qd = db.QueryDefs(src)
        For Each f In qd.Fields
            out = out & vbNullChar & f.Name
        Next
    End If
    If Len(out) > 0 Then SourceFields = Split(Mid$(out, 2), vbNullChar)
Done:
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
