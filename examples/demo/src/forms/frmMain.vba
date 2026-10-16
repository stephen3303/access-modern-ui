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
    MarkDueDates            ' item due dates on the date picker
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

' Go to page key once the calling event has finished. Not while the current
' page refuses to be left (CanLeavePage).
Public Sub QueueNav(ByVal key As String)
    If Not CanLeavePage() Then Exit Sub
    m_pending = key
    Me.TimerInterval = 10
End Sub

' Reopen the window on page pg (after a palette switch).
Public Sub Reopen(ByVal pg As String)
    TempVars("StartPage") = pg
    QueueNav "*reopen"
End Sub

' True when the current page can be swapped or closed: a page with unsaved
' work it can't save yet (pgList with an empty required field) defines a
' public CanLeave that saves it or says why not, and returns False. Pages
' without CanLeave can always be left. The sidebar and Appearance are labels,
' so clicking them never takes the focus out of the page, which is what
' would otherwise make Access save it.
Public Function CanLeavePage() As Boolean
    Dim ok As Boolean, errNo As Long, errDesc As String, errSrc As String
    CanLeavePage = True
    If Len(Me.sfPage.SourceObject) = 0 Then Exit Function
    On Error Resume Next
    ok = Me.sfPage.Form.CanLeave()
    errNo = Err.Number: errDesc = Err.Description: errSrc = Err.Source
    On Error GoTo 0
    Select Case errNo
        Case 0: CanLeavePage = ok
        Case 438, 2465          ' the page has no CanLeave
        Case Else: Err.Raise errNo, errSrc, errDesc
    End Select
End Function

' Closing the window: not while the page refuses to be left, so Access
' doesn't drop the page's unsaved record or ask "close anyway?".
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Err_Handler
    Cancel = Not CanLeavePage()
    Exit Sub
Err_Handler:
    HandleError "frmMain.Form_Unload"
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
        Case "steps": formName = "pgSteps"
        Case Else: key = "cards": formName = "pgCards"
    End Select
    m_page = key
    If Me.sfPage.SourceObject <> formName Then Me.sfPage.SourceObject = formName
    RefreshNav
End Sub

' Highlight the current page in the sidebar.
Private Sub RefreshNav()
    Dim k As Variant
    For Each k In Array("cards", "tables", "list", "steps")
        With Me("nt_" & k)
            .BackStyle = IIf(k = m_page, 1, 0)
            .BackColor = Th(CLR_NAV_ACTIVE_BG)
            .ForeColor = IIf(k = m_page, Th(CLR_NAV_ACTIVE_FG), Th(CLR_NAV_FG))
        End With
    Next
End Sub
