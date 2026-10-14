Option Compare Database
Option Explicit

' Column filter for a table whose GridHeaders show funnels (GridFilter opens
' it; BuildFilterDialog builds it). Lists the column's values; the selected
' (highlighted) ones show. A click or Space selects or clears a value; Select
' all selects every value, or clears them all when every value is selected.
' OK hands the selection back (SetGridFilterChoice) and closes; Cancel or Esc
' closes without changing the table. OK needs at least one value selected.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Dim items As Variant, checked As Variant, i As Long, rows As String
    Me.Caption = GridFilterTitle()
    Me.lblHint.Caption = GridFilterHint()
    items = GridFilterItems()
    checked = GridFilterChecked()
    For i = 0 To UBound(items)
        rows = rows & IIf(i > 0, ";", "") & i & ";" & Quoted(items(i))
    Next
    Me.lstValues.RowSource = rows
    For i = 0 To UBound(items)
        Me.lstValues.Selected(i) = checked(i)
    Next
    ShowCount
    Me.lstValues.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgFilter.Form_Load"
End Sub

' A value-list item: quoted so a value may hold ; or ".
Private Function Quoted(ByVal s As String) As String
    Quoted = """" & Replace(s, """", """""") & """"
End Function

' Show how many values are selected, name Select all for what it will do, and
' enable OK only while a value is selected.
Private Sub ShowCount()
    Dim n As Long, total As Long
    n = Me.lstValues.ItemsSelected.Count
    total = Me.lstValues.ListCount
    Me.lblCount.Caption = n & " of " & total & IIf(total = 1, " value", " values") & " shown"
    Me.btnAll.Caption = IIf(n = total, "Clear all", "Select all")
    SetEnabled Me.btnOK, n > 0
End Sub

Private Sub lstValues_AfterUpdate()
    On Error GoTo Err_Handler
    ShowCount
    Exit Sub
Err_Handler:
    HandleError "dlgFilter.lstValues_AfterUpdate"
End Sub

' Select every value, or clear them all when every value is selected.
Private Sub btnAll_Click()
    On Error GoTo Err_Handler
    Dim i As Long, pick As Boolean
    pick = Me.lstValues.ItemsSelected.Count < Me.lstValues.ListCount
    For i = 0 To Me.lstValues.ListCount - 1
        Me.lstValues.Selected(i) = pick
    Next
    ShowCount                   ' setting Selected in code fires no AfterUpdate
    Exit Sub
Err_Handler:
    HandleError "dlgFilter.btnAll_Click"
End Sub

Private Sub btnOK_Click()
    On Error GoTo Err_Handler
    Dim i As Long, choice As String
    For i = 0 To Me.lstValues.ListCount - 1
        choice = choice & IIf(Me.lstValues.Selected(i), "1", "0")
    Next
    SetGridFilterChoice choice
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgFilter.btnOK_Click"
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgFilter.btnCancel_Click"
End Sub
