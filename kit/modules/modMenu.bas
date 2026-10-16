Option Compare Database
Option Explicit

' Access UI kit - dropdown action menu (a web "dropdown menu" or kebab menu):
' KitMenu opens a short list of actions under the button that was clicked and
' returns the key of the one the user picks. The menu (dlgMenu, built by
' BuildMenuDialog) is a borderless pop-up with a hairline edge, sized to its
' items, placed under the button (above it when there's no room below, and
' kept on the button's monitor). A row lights up under the mouse; Up, Down,
' Home, End and Enter work from the keyboard. A click anywhere else, Esc or
' Tab closes it with nothing picked.
'
' Items are "key|Caption" or "key|Caption|glyph", glyph a Segoe MDL2 Assets
' code point in hex (E70F is Edit, E74D Delete). "-" is a separator. A caption
' starting with "!" is destructive (CLR_DESTRUCTIVE, for Delete); a key
' starting with "~" is disabled (CLR_MUTED_FG, can't be picked).

Public Const MENU_ROWS As Long = 10        ' most items a menu takes, separators included
Public Const MENU_ROW_H As Long = 32       ' an item's height (px)
Public Const MENU_SEP_H As Long = 9        ' a separator's height (px)
Public Const MENU_PAD As Long = 4          ' inside the edge, around the rows (px)
Public Const MENU_MIN_W As Long = 160      ' narrowest menu (px)
Public Const MENU_MAX_W As Long = 320      ' widest menu (px), its build width
Private Const MENU_GAP As Long = 4         ' between the button and the menu (px)

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Private Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hWnd As LongPtr, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function MonitorFromWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" ( _
    ByVal hMonitor As LongPtr, lpmi As MONITORINFO) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private Const LOGPIXELSY As Long = 90
Private Const MONITOR_DEFAULTTONEAREST As Long = 2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4

Private m_items As Variant      ' the open menu's items
Private m_pick As String        ' the key picked, "" for none
Private m_hot As Long           ' the lit row, -1 for none
Private m_anchor As RECT        ' the button, screen pixels
Private m_hAnchor As LongPtr    ' the button's form window (for its monitor)
Private m_closing As Boolean

' Show a menu of items under button anchor and wait for the user: returns the
' picked item's key, or "" when the menu was closed with nothing picked.
' Raises error 5 for no items, more than MENU_ROWS, or a bad item. Without
' dlgMenu (not built) it returns "" and logs why.
'   Select Case KitMenu(Me.btnMore, "edit|Edit|E70F", "copy|Duplicate|E8C8", "-", "del|!Delete|E74D")
Public Function KitMenu(anchor As Control, ParamArray items() As Variant) As String
    Dim list As Variant
    list = items
    If Not MenuOpen(anchor, list) Then Exit Function
    Do While IsFormOpen("dlgMenu")
        DoEvents
        Sleep 15
    Loop
    KitMenu = m_pick
End Function

' Open the menu without waiting (KitMenu waits; tests drive it this way).
' Returns False, logging why, when dlgMenu isn't built.
Public Function MenuOpen(anchor As Control, items As Variant) As Boolean
    Dim i As Long, p() As String
    If Not IsArray(items) Then Err.Raise 5, , "KitMenu has no items"
    If UBound(items) < LBound(items) Then Err.Raise 5, , "KitMenu has no items"
    If UBound(items) - LBound(items) + 1 > MENU_ROWS Then Err.Raise 5, , "KitMenu takes at most " & MENU_ROWS & " items"
    For i = LBound(items) To UBound(items)
        If CStr(items(i)) <> "-" Then
            p = Split(CStr(items(i)), "|")
            If UBound(p) < 1 Or UBound(p) > 2 Then Err.Raise 5, , "KitMenu item is not key|Caption[|glyph]: " & items(i)
            If Len(p(0)) = 0 Then Err.Raise 5, , "KitMenu item has no key: " & items(i)
        End If
    Next
    m_items = items
    m_pick = ""
    m_hot = -1
    m_closing = False
    AnchorRect anchor
    If IsFormOpen("dlgMenu") Then DoCmd.Close acForm, "dlgMenu"
    On Error GoTo NotBuilt
    DoCmd.OpenForm "dlgMenu"
    MenuOpen = True
    Exit Function
NotBuilt:
    LogError "modMenu.MenuOpen", "dlgMenu could not open: " & Err.Description
End Function

' The key the last menu returned ("" for none).
Public Function MenuResult() As String
    MenuResult = m_pick
End Function

' dlgMenu's Load: fill the rows, size the window to them and place it.
Public Sub MenuShow(frm As Form)
    Dim i As Long, r As Long, y As Long, w As Long, p() As String, cap As String, ok As Boolean
    Dim capW As Long, n As Long
    n = UBound(m_items) - LBound(m_items) + 1
    ' As wide as the widest caption, within MENU_MIN_W and MENU_MAX_W.
    w = px(MENU_MIN_W)
    For i = 0 To n - 1
        If ItemText(i) <> "-" Then
            capW = px(40) + MeasureText(ItemCaption(i), 9, 400) + px(24) + 2 * px(MENU_PAD)
            If capW > w Then w = capW
        End If
    Next
    If w > px(MENU_MAX_W) Then w = px(MENU_MAX_W)
    y = px(MENU_PAD)
    For r = 0 To MENU_ROWS - 1
        If r < n Then
            If ItemText(r) = "-" Then
                RowParts frm, r, False
                With frm.Controls("ruleSep" & r)
                    .Top = y + px(MENU_SEP_H \ 2) + 15: .Width = w - 2 * px(MENU_PAD): .Visible = True
                End With
                y = y + px(MENU_SEP_H)
            Else
                p = Split(ItemText(r), "|")
                cap = ItemCaption(r)
                frm.Controls("ruleSep" & r).Visible = False
                frm.Controls("boxRow" & r).Top = y
                frm.Controls("boxRow" & r).Width = w - 2 * px(MENU_PAD)
                frm.Controls("lblIcon" & r).Top = y
                frm.Controls("lblCap" & r).Top = y
                frm.Controls("lblCap" & r).Width = w - 2 * px(MENU_PAD)
                frm.Controls("lblCap" & r).Caption = Replace(cap, "&", "&&")
                frm.Controls("lblIcon" & r).Caption = ""
                If UBound(p) >= 2 Then
                    If Len(p(2)) > 0 Then frm.Controls("lblIcon" & r).Caption = ChrW(CLng("&H" & p(2)))
                End If
                RowColors frm, r
                RowParts frm, r, True
                frm.Controls("boxRow" & r).Visible = False
                y = y + px(MENU_ROW_H)
            End If
        Else
            RowParts frm, r, False
            frm.Controls("ruleSep" & r).Visible = False
        End If
    Next
    y = y + px(MENU_PAD)
    ' Hairline edges: each line paints 1px up and left of where it is placed.
    frm.Controls("ruleTop").Width = w
    frm.Controls("ruleBottom").Top = y: frm.Controls("ruleBottom").Width = w
    frm.Controls("ruleLeft").Height = y
    frm.Controls("ruleRight").Left = w: frm.Controls("ruleRight").Height = y
    frm.InsideWidth = w
    frm.InsideHeight = y
    Place frm
End Sub

' Where a menu w x h (pixels) goes for a button at aL..aR, aT..aB on a screen
' work area wL..wR, wT..wB: under the button, left edges lined up, gap pixels
' below it; right edges lined up when it would run off the right; above the
' button when it would run off the bottom; never off the top or left.
Public Sub MenuSpot(ByVal aL As Long, ByVal aT As Long, ByVal aR As Long, ByVal aB As Long, ByVal w As Long, ByVal h As Long, _
        ByVal wL As Long, ByVal wT As Long, ByVal wR As Long, ByVal wB As Long, ByVal gap As Long, x As Long, y As Long)
    x = aL
    If x + w > wR Then x = aR - w
    If x < wL Then x = wL
    y = aB + gap
    If y + h > wB Then y = aT - gap - h
    If y < wT Then y = wT
End Sub

' A row's click (set by BuildMenuDialog): pick its item, unless disabled.
Public Function MenuPick(frm As Form, ByVal r As Long)
    On Error GoTo Fail
    If Not Pickable(r) Then Exit Function
    m_pick = Split(ItemText(r), "|")(0)
    MenuClose frm
    Exit Function
Fail:
    HandleError "modMenu.MenuPick", "row " & r
End Function

' A row's mouse move (set by BuildMenuDialog): light row r (-1 for none).
Public Function MenuHover(frm As Form, ByVal r As Long)
    On Error GoTo Fail
    If r = m_hot Then Exit Function
    If r >= 0 Then If Not Pickable(r) Then r = -1
    If m_hot >= 0 Then frm.Controls("boxRow" & m_hot).Visible = False
    m_hot = r
    If r >= 0 Then frm.Controls("boxRow" & r).Visible = True
    Exit Function
Fail:
    LogError "modMenu.MenuHover", "row " & r
End Function

' dlgMenu's KeyDown: Up, Down, Home and End move the lit row (skipping
' separators and disabled items), Enter or Space picks it, Esc and Tab close.
Public Sub MenuKey(frm As Form, KeyCode As Integer)
    Dim n As Long, r As Long, stp As Long
    n = UBound(m_items) - LBound(m_items) + 1
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp, vbKeyHome, vbKeyEnd
            Select Case KeyCode
                Case vbKeyDown: stp = 1: r = m_hot + 1
                Case vbKeyUp: stp = -1: r = IIf(m_hot < 0, n - 1, m_hot - 1)
                Case vbKeyHome: stp = 1: r = 0
                Case vbKeyEnd: stp = -1: r = n - 1
            End Select
            Do While r >= 0 And r < n
                If Pickable(r) Then MenuHover frm, r: Exit Do
                r = r + stp
            Loop
        Case vbKeyReturn, vbKeySpace
            If m_hot >= 0 Then MenuPick frm, m_hot
        Case vbKeyEscape, vbKeyTab
            MenuClose frm
        Case Else
            Exit Sub
    End Select
    KeyCode = 0
End Sub

' Close the menu (once: dlgMenu's Deactivate also asks, as it closes).
Public Sub MenuClose(frm As Form)
    If m_closing Then Exit Sub
    m_closing = True
    DoCmd.Close acForm, frm.Name
End Sub

' dlgMenu's Deactivate: a click outside the menu. Closing during the event is
' refused, so the form's timer closes it a moment later.
Public Sub MenuLostFocus(frm As Form)
    If Not m_closing Then frm.TimerInterval = 1
End Sub

' Build dlgMenu, the pop-up KitMenu opens. Call it once with the app's dialogs.
Public Sub BuildMenuDialog()
    Dim c As Control, r As Long, w As Long, h As Long, y As Long, nm As Variant
    w = px(MENU_MAX_W)
    h = MENU_ROWS * px(MENU_ROW_H) + 2 * px(MENU_PAD)
    NewForm "", 0, w, h
    With F
        .Caption = "Menu"
        .PopUp = True
        .Modal = False
        .BorderStyle = 0
        .AutoCenter = False
        .AutoResize = False
        .ControlBox = False
        .MinMaxButtons = 0
        .CloseButton = False
        .ShortcutMenu = False
        .ScrollBars = 0
        .RecordSelectors = False
        .NavigationButtons = False
        .Section(acDetail).BackColor = CLR_DIALOG
        .Section(acDetail).AlternateBackColor = CLR_DIALOG
        .Section(acDetail).OnMouseMove = "=MenuHover([Form],-1)"
    End With
    Ev F, "OnLoad", "OnDeactivate", "OnTimer"
    ' Takes the keys: a 1px box with no border, the only Tab stop.
    Set c = Txt("txtKeys", "", 0, 0, 15, 15)
    c.BorderStyle = 0: c.BackStyle = 0: c.OnGotFocus = "": c.OnLostFocus = ""
    Ev c, "OnKeyDown"
    For r = 0 To MENU_ROWS - 1
        y = px(MENU_PAD) + r * px(MENU_ROW_H)
        Rule "ruleSep" & r, px(MENU_PAD), y, w - 2 * px(MENU_PAD), 0, CLR_BORDER
        Set c = Box("boxRow" & r, px(MENU_PAD), y, w - 2 * px(MENU_PAD), px(MENU_ROW_H), CLR_OUT_HOVER)
        c.Visible = False
        ' The caption spans the row (its left margin clears the icon), so the
        ' mouse is over one of the row's parts anywhere in it.
        Set c = Lbl("lblCap" & r, "", px(MENU_PAD), y, w - 2 * px(MENU_PAD), px(MENU_ROW_H), 9, False, CLR_FG)
        c.LeftMargin = px(40)
        VCenter c, px(MENU_ROW_H), 9, True
        Set c = Lbl("lblIcon" & r, "", px(MENU_PAD + 10), y, px(20), px(MENU_ROW_H), 10, False, CLR_MUTED_FG)
        c.FontName = "Segoe MDL2 Assets": c.TextAlign = 2
        VCenter c, px(MENU_ROW_H), 10, True
        For Each nm In Array("boxRow", "lblIcon", "lblCap")
            F.Controls(nm & r).OnClick = "=MenuPick([Form]," & r & ")"
            F.Controls(nm & r).OnMouseMove = "=MenuHover([Form]," & r & ")"
        Next
    Next
    Rule "ruleTop", 0, 15, w, 0, CLR_INPUT_BORDER
    Rule "ruleBottom", 0, h, w, 0, CLR_INPUT_BORDER
    Rule "ruleLeft", 15, 0, 0, h, CLR_INPUT_BORDER
    Rule "ruleRight", w, 0, 0, h, CLR_INPUT_BORDER
    SaveForm "dlgMenu"
End Sub

Private Function ItemText(ByVal i As Long) As String
    ItemText = CStr(m_items(LBound(m_items) + i))
End Function

' An item's caption without its "!" (destructive) mark.
Private Function ItemCaption(ByVal i As Long) As String
    Dim p() As String
    p = Split(ItemText(i), "|")
    ItemCaption = p(1)
    If Left$(ItemCaption, 1) = "!" Then ItemCaption = Mid$(ItemCaption, 2)
End Function

Private Function Pickable(ByVal r As Long) As Boolean
    If r < 0 Or r > UBound(m_items) - LBound(m_items) Then Exit Function
    If ItemText(r) = "-" Then Exit Function
    Pickable = (Left$(ItemText(r), 1) <> "~")
End Function

Private Sub RowColors(frm As Form, ByVal r As Long)
    Dim fore As Long, p() As String
    p = Split(ItemText(r), "|")
    If Left$(p(0), 1) = "~" Then
        fore = Th(CLR_MUTED_FG)
    ElseIf Left$(p(1), 1) = "!" Then
        fore = Th(CLR_DESTRUCTIVE)
    Else
        fore = Th(CLR_FG)
    End If
    frm.Controls("lblCap" & r).ForeColor = fore
    frm.Controls("lblIcon" & r).ForeColor = IIf(fore = Th(CLR_FG), Th(CLR_MUTED_FG), fore)
End Sub

Private Sub RowParts(frm As Form, ByVal r As Long, ByVal v As Boolean)
    frm.Controls("lblIcon" & r).Visible = v
    frm.Controls("lblCap" & r).Visible = v
    If Not v Then frm.Controls("boxRow" & r).Visible = False
End Sub

' The button's place on screen, in pixels.
Private Sub AnchorRect(c As Control)
    Dim frm As Object, pt As POINTAPI, dpi As Long, hdc As LongPtr, top As Long
    Set frm = c.Parent
    Do While Not TypeOf frm Is Form
        Set frm = frm.Parent
    Loop
    top = c.Top
    If c.Section = acDetail Then top = top + HeaderHeight(frm)
    hdc = GetDC(0): dpi = GetDeviceCaps(hdc, LOGPIXELSY): ReleaseDC 0, hdc
    If dpi <= 0 Then dpi = 96
    m_hAnchor = frm.hWnd
    ClientToScreen m_hAnchor, pt
    m_anchor.Left = pt.x + CLng(c.Left * dpi / 1440)
    m_anchor.Top = pt.y + CLng(top * dpi / 1440)
    m_anchor.Right = pt.x + CLng((c.Left + c.Width) * dpi / 1440)
    m_anchor.Bottom = pt.y + CLng((top + c.Height) * dpi / 1440)
End Sub

Private Function HeaderHeight(frm As Object) As Long
    On Error Resume Next
    If frm.Section(acHeader).Visible Then HeaderHeight = frm.Section(acHeader).Height
End Function

Private Sub Place(frm As Form)
    Dim rc As RECT, mi As MONITORINFO, x As Long, y As Long, gap As Long
    GetWindowRect frm.hWnd, rc
    mi.cbSize = LenB(mi)
    If GetMonitorInfo(MonitorFromWindow(m_hAnchor, MONITOR_DEFAULTTONEAREST), mi) = 0 Then Exit Sub
    gap = CLng(MENU_GAP * (rc.Right - rc.Left) / (frm.InsideWidth / 15))
    MenuSpot m_anchor.Left, m_anchor.Top, m_anchor.Right, m_anchor.Bottom, rc.Right - rc.Left, rc.Bottom - rc.Top, _
        mi.rcWork.Left, mi.rcWork.Top, mi.rcWork.Right, mi.rcWork.Bottom, gap, x, y
    SetWindowPos frm.hWnd, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER
End Sub
