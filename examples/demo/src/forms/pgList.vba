Option Compare Database
Option Explicit

' List and detail (see Build_pgList). Clicking a row shows it in the detail
' panel; the first row opens whenever the list changes. Clicking a column
' caption sorts the list (GridSort). Print and Export use the rows the list
' shows, in its order.

Private m_chip As Long
Private m_search As String
Private m_auto As Boolean       ' RefreshPage is opening the first row (not a recent record)

Private Sub Form_Load()
    Dim id As Long
    On Error GoTo Err_Handler
    Me.sfList.Form.RecordSource = "SELECT id, name, owner, status, Format(due_date,'dd mmm yyyy') AS due, due_date " & _
        "FROM demo_item ORDER BY due_date"
    RefreshPage
    ShowOverdue
    ' A recent record picked in the Ctrl+K palette (DemoOpenItem).
    If Not IsNull(TempVars("OpenItem")) Then
        id = TempVars("OpenItem")
        TempVars("OpenItem") = Null
        OpenItem id
    End If
    Exit Sub
Err_Handler:
    HandleError "pgList.Form_Load"
End Sub

' Row click in the list (subGrid): show that item. When the open item can't
' be saved (a required field is empty) it stays, and so does its row.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    On Error GoTo Err_Handler
    If IsNull(Me.sfList.Form.txtSel) Then Exit Sub
    If Not ShowDetail(Me, "sfDetail", "subItem", "id=" & Me.sfList.Form.txtSel) Then
        Me.sfList.Form.txtSel = Me.sfDetail.Form!id
        Me.sfList.Form.Recalc
    ElseIf Me.sfDetail.Visible And Not m_auto Then
        KitRecent Nz(Me.sfDetail.Form!name, ""), "DemoOpenItem(" & Me.sfDetail.Form!id & ")", "E8A5"
    End If
    ShowCrumbs
    Exit Sub
Err_Handler:
    HandleError "pgList.GridClick"
End Sub

' frmMain asks before it swaps or closes this page: save the open item, or
' keep the page while a required field is empty (CheckRequired says which).
Public Function CanLeave() As Boolean
    CanLeave = SaveDetail(Me, "sfDetail")
End Function

' The header's alert (Alert): how many open items are past their due date,
' or nothing when none are.
Private Sub ShowOverdue()
    Dim n As Long
    n = DCount("*", "demo_item", "due_date < Date() And status Not In ('Complete','Cancelled')")
    If n = 0 Then
        HideAlert Me, "alOverdue"
    Else
        ShowAlert Me, "alOverdue", n & IIf(n = 1, " open item is", " open items are") & " past the due date.", "warning"
    End If
End Sub

' The breadcrumb: All items, the chip's status when one is picked, and the
' open item (the current page).
Private Sub ShowCrumbs()
    Dim nm As String
    If Me.sfDetail.Visible Then nm = Nz(Me.sfDetail.Form!name, "")
    If m_chip = 0 And Len(nm) = 0 Then
        SetBreadcrumb Me, "crumb", "c0|All items"
    ElseIf m_chip = 0 Then
        SetBreadcrumb Me, "crumb", "c0|All items", "item|" & nm
    ElseIf Len(nm) = 0 Then
        SetBreadcrumb Me, "crumb", "c0|All items", "c" & m_chip & "|" & Me("chip" & m_chip).Tag
    Else
        SetBreadcrumb Me, "crumb", "c0|All items", "c" & m_chip & "|" & Me("chip" & m_chip).Tag, "item|" & nm
    End If
End Sub

' A breadcrumb click: All items, or a status, as its chip would.
Public Sub CrumbGo(ByVal key As String)
    On Error GoTo Err_Handler
    If Left$(key, 1) = "c" Then Chip CLng(Mid$(key, 2))
    Exit Sub
Err_Handler:
    HandleError "pgList.CrumbGo"
End Sub

' More (the detail card's menu button): the open item's actions.
Private Sub btnMore_Click()
    On Error GoTo Err_Handler
    ItemActions False
    Exit Sub
Err_Handler:
    HandleError "pgList.btnMore_Click"
End Sub

' Right-click on a row (subGrid): the row is selected and shown, then the same
' actions open at the pointer.
Public Sub GridRightClick(ByVal grid As String)
    On Error GoTo Err_Handler
    If Not Me.sfDetail.Visible Then Exit Sub
    ' GridClick kept the open item (it can't be saved yet), and said why.
    If Nz(Me.sfList.Form.txtSel, 0) <> Nz(Me.sfDetail.Form!id, 0) Then Exit Sub
    ItemActions True
    Exit Sub
Err_Handler:
    HandleError "pgList.GridRightClick"
End Sub

' The open item's actions, under More or at the pointer.
Private Sub ItemActions(ByVal atPointer As Boolean)
    Dim id As Long, flagged As Boolean, k As String, a(3) As String
    If Not Me.sfDetail.Visible Then Exit Sub
    If Not SaveDetail(Me, "sfDetail") Then Exit Sub
    id = Me.sfDetail.Form!id
    flagged = Nz(Me.sfDetail.Form!flagged, False)
    a(0) = "done|Mark complete|E73E"
    a(1) = "flag|" & IIf(flagged, "Remove flag", "Flag for review") & "|E7C1"
    a(2) = "-"
    a(3) = "del|!Delete item|E74D"
    If atPointer Then
        k = KitMenuAt(a(0), a(1), a(2), a(3))
    Else
        k = KitMenu(Me.btnMore, a(0), a(1), a(2), a(3))
    End If
    Select Case k
        Case "done"
            CurrentDb.Execute "UPDATE demo_item SET status='Complete', percent_done=100 WHERE id=" & id, dbFailOnError
            Me.sfDetail.Form.Requery
            ItemSaved id
            KitToast "Marked complete.", "success"
        Case "flag"
            CurrentDb.Execute "UPDATE demo_item SET flagged=" & IIf(flagged, "False", "True") & " WHERE id=" & id, dbFailOnError
            Me.sfDetail.Form.Requery
            KitToast IIf(flagged, "Flag removed.", "Flagged for review."), "info"
        Case "del"
            If KitMsgBox("Delete """ & Me.sfDetail.Form!name & """? This can't be undone.", vbYesNo + vbExclamation + vbDefaultButton2, _
                "Delete item") = vbYes Then
                CurrentDb.Execute "DELETE FROM demo_item WHERE id=" & id, dbFailOnError
                RefreshPage
                ShowOverdue
                KitToast "Item deleted.", "destructive"
            End If
    End Select
End Sub

' The empty state's button: back to every item.
Private Sub emptyList_b_Click()
    On Error GoTo Err_Handler
    SetSearch Me, "txtSearch", ""
    Chip 0
    Exit Sub
Err_Handler:
    HandleError "pgList.emptyList_b_Click"
End Sub

' Show item id (a recent record from the Ctrl+K palette): every item is
' listed again, so a chip or search can't hide it. Nothing happens when the
' open item can't be saved yet, or the item is gone.
Public Sub OpenItem(ByVal id As Long)
    If Not SaveDetail(Me, "sfDetail") Then Exit Sub
    If Len(m_search) > 0 Then SetSearch Me, "txtSearch", ""     ' its handler refreshes the list
    If m_chip <> 0 Or Len(m_search) > 0 Then m_chip = 0: m_search = "": RefreshPage
    With Me.sfList.Form
        .RecordsetClone.FindFirst "id=" & id
        If .RecordsetClone.NoMatch Then Exit Sub
        .Bookmark = .RecordsetClone.Bookmark
        .txtSel = id
        .Recalc
    End With
    GridClick "sfList", False
End Sub

' The detail panel saved an item: show the change in the list.
Public Sub ItemSaved(ByVal id As Long)
    Me.sfList.Form.Requery
    Me.sfList.Form.txtSel = id
    RefreshCounts
    ShowOverdue
    ShowCrumbs
End Sub

' Status chip click: 0 is All.
Public Function Chip(ByVal i As Long)
    On Error GoTo Err_Handler
    If Not SaveDetail(Me, "sfDetail") Then Exit Function    ' open item can't be saved yet
    m_chip = i
    RefreshPage
    Exit Function
Err_Handler:
    HandleError "pgList.Chip"
End Function

' The search box's handler (SearchBox): filter as the user types.
Public Sub SearchFor(ByVal text As String)
    On Error GoTo Err_Handler
    m_search = text
    RefreshPage
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(m_search)
    Exit Sub
Err_Handler:
    HandleError "pgList.SearchFor"
End Sub

Private Sub btnPrint_Click()
    On Error GoTo Err_Handler
    PrintGrid Me.sfList.Form, "Items", FilterText()
    Exit Sub
Err_Handler:
    HandleError "pgList.btnPrint_Click"
End Sub

Private Sub btnCsv_Click()
    On Error GoTo Err_Handler
    Export "csv"
    Exit Sub
Err_Handler:
    HandleError "pgList.btnCsv_Click"
End Sub

Private Sub btnXlsx_Click()
    On Error GoTo Err_Handler
    Export "xlsx"
    Exit Sub
Err_Handler:
    HandleError "pgList.btnXlsx_Click"
End Sub

' Export the list as shown next to the front end and say where it went.
Private Sub Export(ByVal ext As String)
    Dim path As String, n As Long
    path = CurrentProject.Path & "\Kit Demo items " & Format$(Now(), "yyyy-mm-dd hhnnss") & "." & ext
    n = ExportGrid(Me.sfList.Form, path)
    Notify n & " item(s) exported to" & vbCrLf & path, vbInformation
End Sub

' What the list is filtered by, for the printout's subtitle.
Private Function FilterText() As String
    Dim s As String
    s = "Status: " & Me.chip0.Tag
    If m_chip > 0 Then s = "Status: " & Me("chip" & m_chip).Tag
    If Len(m_search) > 0 Then s = s & "   Search: " & m_search
    FilterText = s
End Function

' Apply the chip and search, mark the search in the rows, then open the
' first item (or show the empty state).
Private Sub RefreshPage()
    FilterGrid Me.sfList, ItemFilter(Status(m_chip), m_search)
    GridHighlight Me.sfList, m_search
    RefreshCounts
    ' ClearDetail's result isn't needed: every caller (Chip, Form_Load, and
    ' SearchFor, which only runs once the focus has left the panel)
    ' has already saved the open item.
    m_auto = True
    If IsNull(SelectFirstRow(Me.sfList)) Then ClearDetail Me, "sfDetail"
    m_auto = False
    ShowCrumbs
End Sub

Private Sub RefreshCounts()
    Dim i As Long, n As Long
    For i = 0 To 4
        StyleChip Me("chip" & i), i = m_chip, ItemCount(Status(i), m_search)
    Next
    n = ItemCount(Status(m_chip), m_search)
    ShowEmpty Me, "emptyList", n = 0
    ' The empty state shows in place of the list: a subform always covers a
    ' label.
    SetVisible Me.sfList, n > 0
End Sub

' Status a chip filters on ("" for All).
Private Function Status(ByVal i As Long) As String
    Status = Choose(i + 1, "", "In progress", "Blocked", "Not started", "Complete")
End Function
