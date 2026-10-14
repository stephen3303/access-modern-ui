Option Compare Database
Option Explicit

' Cards and progress (see Build_pgCards).

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.sfItems.Form.RecordSource = "SELECT id, name, owner, category, status, Format(due_date,'dd mmm yyyy') AS due " & _
        "FROM demo_item ORDER BY due_date"
    Exit Sub
Err_Handler:
    HandleError "pgCards.Form_Load"
End Sub

' Row click in the table (subGrid): nothing to open on this page.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
End Sub

Private Sub btnRoundA_Click()
    On Error GoTo Err_Handler
    Notify "Primary button clicked.", vbInformation
    Exit Sub
Err_Handler:
    HandleError "pgCards.btnRoundA_Click"
End Sub

Private Sub btnRoundB_Click()
    On Error GoTo Err_Handler
    Notify "Outline button clicked.", vbInformation
    Exit Sub
Err_Handler:
    HandleError "pgCards.btnRoundB_Click"
End Sub

' KitInputBox: the answer, or that it was cancelled (Null).
Private Sub btnSquareA_Click()
    On Error GoTo Err_Handler
    Dim v As Variant
    v = KitInputBox("What name should the demo use? Leave it empty to see an empty answer.", APP_NAME, "Sailor")
    If IsNull(v) Then
        Notify "Cancelled.", vbInformation
    Else
        Notify "You typed """ & v & """.", vbInformation
    End If
    Exit Sub
Err_Handler:
    HandleError "pgCards.btnSquareA_Click"
End Sub

' Masked KitInputBox: only the answer's length is shown, never the text.
Private Sub btnSquareB_Click()
    On Error GoTo Err_Handler
    Dim v As Variant
    v = KitInputBox("Enter a PIN. Each character shows as * while you type.", APP_NAME, masked:=True)
    If IsNull(v) Then
        Notify "Cancelled.", vbInformation
    Else
        Notify "You typed " & Len(v) & " character(s).", vbInformation
    End If
    Exit Sub
Err_Handler:
    HandleError "pgCards.btnSquareB_Click"
End Sub
