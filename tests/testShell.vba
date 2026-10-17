Option Compare Database
Option Explicit
' Test fixture (modKitTests): an app shell (ShellBuild) whose events go to
' modKitShell, with every optional hook.

Private Sub Form_Load()
    ShellLoad Me
End Sub

Private Sub Form_Resize()
    GridScrollAll Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = Not ShellCanLeave(Me)
End Sub

Private Sub Form_Timer()
    ShellTimer Me
End Sub

Public Property Get PageKey() As String
    PageKey = ShellPageKey()
End Property

Public Sub QueueNav(ByVal key As String)
    ShellQueue Me, key
End Sub

Public Sub Reopen(ByVal pg As String)
    ShellReopen Me, pg
End Sub

Public Function CanLeavePage() As Boolean
    CanLeavePage = ShellCanLeave(Me)
End Function

' "secret" goes to Gamma.
Public Function RoutePage(ByVal key As String) As String
    RoutePage = IIf(key = "secret", "c", key)
End Function

' "detail" is a page with no sidebar item.
Public Function PageForm(ByVal key As String) As String
    If key = "detail" Then PageForm = "subLeave"
End Function

Public Sub PageShown(ByVal key As String)
    TempVars("ShellShown") = key
End Sub
