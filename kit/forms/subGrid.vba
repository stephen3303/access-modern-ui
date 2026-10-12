Option Compare Database
Option Explicit

' Rows of a styled table (modUIKit.BuildGrid; shared by every subGrid* form).
' Clicking a row selects it and tells the parent form, whose
' GridClick(grid, dbl) gets the subform control's name and whether it was a
' double click.

Public Function RowClick()
    SelectRow False
End Function

Public Function RowDblClick()
    SelectRow True
End Function

Private Sub SelectRow(ByVal dbl As Boolean)
    On Error GoTo Err_Handler
    Me.txtFocusSink.SetFocus   ' a focused cell is highlighted and painted white
    If IsNull(Me!id) Then Exit Sub
    Me.txtSel = Me!id
    Me.Recalc
    Me.Parent.GridClick GridName(), dbl
    Exit Sub
Err_Handler:
    HandleError Me.Name & ".SelectRow"
End Sub

' Name of the parent's subform control that shows this form.
Private Function GridName() As String
    Dim c As Control
    For Each c In Me.Parent.Controls
        If c.ControlType = acSubform Then
            If c.SourceObject = Me.Name Then
                GridName = c.Name
                Exit Function
            End If
        End If
    Next
End Function
