Option Compare Database
Option Explicit

' Kit Demo shell: the sidebar and the page host sfPage (ShellBuild). Its
' events go to modKitShell, which swaps pages from the timer (QueueNav),
' never while one of their own events is running.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    DoCmd.Maximize
    SplashStatus "Opening Cards and progress" & ChrW(8230)
    ShellLoad Me
    MarkDueDates            ' item due dates on the date picker
    Exit Sub
Err_Handler:
    HandleError "frmMain.Form_Load"
End Sub

' A resize doesn't fire a table's OnCurrent, so recheck which tables need a
' scroll bar on every page shown (GridScrollAll).
Private Sub Form_Resize()
    On Error GoTo Err_Handler
    GridScrollAll Me
    Exit Sub
Err_Handler:
    HandleError "frmMain.Form_Resize"
End Sub

' Key of the page being shown (SetPalette reopens on it; the error log records it).
Public Property Get PageKey() As String
    PageKey = ShellPageKey()
End Property

' Go to page key once the calling event has finished. Not while the current
' page refuses to be left (CanLeavePage).
Public Sub QueueNav(ByVal key As String)
    ShellQueue Me, key
End Sub

' Reopen the window on page pg (after a palette switch).
Public Sub Reopen(ByVal pg As String)
    ShellReopen Me, pg
End Sub

' True when the current page can be swapped or closed (pgList saves its
' open item, or says which required field is empty).
Public Function CanLeavePage() As Boolean
    CanLeavePage = ShellCanLeave(Me)
End Function

' Closing the window: not while the page refuses to be left, so Access
' doesn't drop the page's unsaved record or ask "close anyway?".
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Err_Handler
    Cancel = Not ShellCanLeave(Me)
    Exit Sub
Err_Handler:
    HandleError "frmMain.Form_Unload"
End Sub

Private Sub Form_Timer()
    On Error GoTo Err_Handler
    ShellTimer Me
    Exit Sub
Err_Handler:
    HandleError "frmMain.Form_Timer"
End Sub
