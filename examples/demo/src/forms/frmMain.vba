Option Compare Database
Option Explicit

' Kit Demo shell: the sidebar and the page host sfPage. Pages are swapped
' from the timer (QueueNav), never while one of their own events is running.

Private m_page As String
Private m_pending As String

Private Sub Form_Load()
    On Error GoTo Err_Handler
    DoCmd.Maximize
    SplashStatus "Opening Cards and progress" & ChrW(8230)
    ShowPage Nz(TempVars("StartPage"), "cards")
    TempVars("StartPage") = Null
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
    PageKey = m_page
End Property

' Go to page key once the calling event has finished.
Public Sub QueueNav(ByVal key As String)
    m_pending = key
    Me.TimerInterval = 10
End Sub

' Reopen the window on page pg (after a palette switch).
Public Sub Reopen(ByVal pg As String)
    TempVars("StartPage") = pg
    QueueNav "*reopen"
End Sub

Private Sub Form_Timer()
    On Error GoTo Err_Handler
    Me.TimerInterval = 0
    If m_pending = "*reopen" Then
        m_pending = ""
        DoCmd.Close acForm, Me.Name, acSaveNo
        DoCmd.OpenForm "frmMain"
        Exit Sub
    End If
    If Len(m_pending) > 0 Then
        ShowPage m_pending
        m_pending = ""
    End If
    Exit Sub
Err_Handler:
    HandleError "frmMain.Form_Timer"
End Sub

Private Sub ShowPage(ByVal key As String)
    Dim formName As String
    Select Case key
        Case "tables": formName = "pgTables"
        Case "list": formName = "pgList"
        Case Else: key = "cards": formName = "pgCards"
    End Select
    m_page = key
    If Me.sfPage.SourceObject <> formName Then Me.sfPage.SourceObject = formName
    RefreshNav
End Sub

' Highlight the current page in the sidebar.
Private Sub RefreshNav()
    Dim k As Variant
    For Each k In Array("cards", "tables", "list")
        With Me("nt_" & k)
            .BackStyle = IIf(k = m_page, 1, 0)
            .BackColor = Th(CLR_NAV_ACTIVE_BG)
            .ForeColor = IIf(k = m_page, Th(CLR_NAV_ACTIVE_FG), Th(CLR_NAV_FG))
        End With
    Next
End Sub
