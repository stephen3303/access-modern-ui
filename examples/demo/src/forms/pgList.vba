Option Compare Database
Option Explicit

' List and detail (see Build_pgList). Clicking a row shows it in the detail
' panel; the first row opens whenever the list changes. Print and Export use
' the rows the list shows.

Private m_chip As Long
Private m_search As String

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.sfList.Form.RecordSource = "SELECT id, name, owner, status, Format(due_date,'dd mmm yyyy') AS due " & _
        "FROM demo_item ORDER BY due_date"
    RefreshPage
    Exit Sub
Err_Handler:
    HandleError "pgList.Form_Load"
End Sub

' Row click in the list (subGrid): show that item.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    On Error GoTo Err_Handler
    If IsNull(Me.sfList.Form.txtSel) Then Exit Sub
    ShowDetail Me, "sfDetail", "subItem", "id=" & Me.sfList.Form.txtSel
    Exit Sub
Err_Handler:
    HandleError "pgList.GridClick"
End Sub

' The detail panel saved an item: show the change in the list.
Public Sub ItemSaved(ByVal id As Long)
    Me.sfList.Form.Requery
    Me.sfList.Form.txtSel = id
    RefreshCounts
End Sub

' Status chip click: 0 is All.
Public Function Chip(ByVal i As Long)
    On Error GoTo Err_Handler
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
    If IsNull(SelectFirstRow(Me.sfList)) Then ClearDetail Me, "sfDetail"
End Sub

Private Sub RefreshCounts()
    Dim i As Long, n As Long
    For i = 0 To 4
        StyleChip Me("chip" & i), i = m_chip, ItemCount(Status(i), m_search)
    Next
    n = ItemCount(Status(m_chip), m_search)
    SetVisible Me.lblEmpty, n = 0
End Sub

' Status a chip filters on ("" for All).
Private Function Status(ByVal i As Long) As String
    Status = Choose(i + 1, "", "In progress", "Blocked", "Not started", "Complete")
End Function
