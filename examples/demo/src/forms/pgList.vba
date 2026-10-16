Option Compare Database
Option Explicit

' List and detail (see Build_pgList). Clicking a row shows it in the detail
' panel; the first row opens whenever the list changes. Clicking a column
' caption sorts the list (GridSort). Print and Export use the rows the list
' shows, in its order.

Private m_chip As Long
Private m_search As String

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.sfList.Form.RecordSource = "SELECT id, name, owner, status, Format(due_date,'dd mmm yyyy') AS due, due_date " & _
        "FROM demo_item ORDER BY due_date"
    RefreshPage
    ShowOverdue
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
    End If
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

' The detail panel saved an item: show the change in the list.
Public Sub ItemSaved(ByVal id As Long)
    Me.sfList.Form.Requery
    Me.sfList.Form.txtSel = id
    RefreshCounts
    ShowOverdue
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

Private Sub txtSearch_Change()
    On Error GoTo Err_Handler
    m_search = Nz(Me.txtSearch.Text, "")
    RefreshPage
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(m_search)
    Exit Sub
Err_Handler:
    HandleError "pgList.txtSearch_Change"
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
    ' txtSearch_Change, which only runs once the focus has left the panel)
    ' has already saved the open item.
    If IsNull(SelectFirstRow(Me.sfList)) Then ClearDetail Me, "sfDetail"
End Sub

Private Sub RefreshCounts()
    Dim i As Long, n As Long
    For i = 0 To 4
        StyleChip Me("chip" & i), i = m_chip, ItemCount(Status(i), m_search)
    Next
    n = ItemCount(Status(m_chip), m_search)
    SetVisible Me.lblEmpty, n = 0
    ' The empty state shows in place of the list: a subform always covers a
    ' label.
    SetVisible Me.sfList, n > 0
End Sub

' Status a chip filters on ("" for All).
Private Function Status(ByVal i As Long) As String
    Status = Choose(i + 1, "", "In progress", "Blocked", "Not started", "Complete")
End Function
