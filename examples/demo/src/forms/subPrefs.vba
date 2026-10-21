Option Compare Database
Option Explicit

' Preferences section of the Steps and sections page (see Build_subPrefs).

' The help mark beside ROWS PER PAGE: a popover under it.
Public Function RowsHelp()
    On Error GoTo Err_Handler
    KitPopover Me.lnkRowsHelp, "Rows per page", "How many rows a long table shows before it pages. " & _
        "Fewer rows load faster; more rows scroll less. The demo keeps no settings, so this only moves the slider."
    Exit Function
Err_Handler:
    HandleError "subPrefs.RowsHelp"
End Function

' The slider's onChange.
Public Sub RowsChanged()
    On Error GoTo Err_Handler
    KitToast "Tables would show " & SliderValue(Me, "sldRows") & " rows per page.", "info"
    Exit Sub
Err_Handler:
    HandleError "subPrefs.RowsChanged"
End Sub

' The rating's onChange.
Public Sub UsefulChanged()
    On Error GoTo Err_Handler
    Dim n As Long
    n = RatingValue(Me, "rtUseful")
    Me.lblUseful2.Caption = IIf(n = 0, "", "You gave it " & n & " of 5. Click the same star to clear it.")
    Exit Sub
Err_Handler:
    HandleError "subPrefs.UsefulChanged"
End Sub

' The item picker's keys (PickBox): Up, Down, Enter and Esc work its list.
Private Sub txtWatch_KeyDown(KeyCode As Integer, Shift As Integer)
    KitPickKey Me.txtWatch, KeyCode
End Sub

' Left and Right on the slider move it a step (the kit leaves keys to the form).
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Handler
    If Left$(Screen.ActiveControl.Name, 8) <> "sldRows_" Then Exit Sub
    Select Case KeyCode
        Case vbKeyLeft: SliderStep Me, "sldRows", -1: KeyCode = 0
        Case vbKeyRight: SliderStep Me, "sldRows", 1: KeyCode = 0
    End Select
    Exit Sub
Err_Handler:
    HandleError "subPrefs.Form_KeyDown"
End Sub
