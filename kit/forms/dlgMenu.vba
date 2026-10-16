Option Compare Database
Option Explicit

' Dropdown menu for KitMenu (BuildMenuDialog builds it; modMenu does the
' work). It closes when it loses the focus, so a click outside it closes it.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    MenuShow Me
    Me.txtKeys.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgMenu.Form_Load"
End Sub

Private Sub Form_Deactivate()
    MenuLostFocus Me
End Sub

Private Sub Form_Timer()
    Me.TimerInterval = 0
    MenuClose Me
End Sub

Private Sub txtKeys_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Handler
    MenuKey Me, KeyCode
    Exit Sub
Err_Handler:
    HandleError "dlgMenu.txtKeys_KeyDown"
End Sub
