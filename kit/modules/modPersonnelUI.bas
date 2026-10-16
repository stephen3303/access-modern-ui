Option Compare Database
Option Explicit

' Access UI kit - the Personnel directory's screens (data and sign-in rules
' are in modPersonnel).
'
' Startup, straight after the app's own KitEnsureBackend:
'   If Not KitPersonnelStart("item", APP_CODE) Then DoCmd.Quit
' finds or sets up the Personnel file (dlgPersonnelSetup), then signs the
' Windows user in with no questions, or shows the sign-in screen (dlgSignIn)
' when their account isn't set up. The sidebar's Sign out button:
'   If Not KitPersonnelSignOut(APP_CODE) Then DoCmd.Quit
' clears the session and shows the sign-in screen, so another person can
' sign in on the same copy.
'
' Build both dialogs with the app's others: BuildPersonnelDialogs.

Public Const SETUP_FORM As String = "dlgPersonnelSetup"
Public Const SIGNIN_FORM As String = "dlgSignIn"

' What the open Personnel dialog is for (the dialogs read these).
Private m_appTable As String
Private m_appCode As String
Private m_setupMode As Long
Private m_signInWhy As String

' ---------------------------------------------------------------------------
' Startup and sign out
' ---------------------------------------------------------------------------

' Make sure the Personnel file is linked and someone is signed in to appCode
' (see the top of the module). False when the user closed the app instead.
Public Function KitPersonnelStart(ByVal appTable As String, ByVal appCode As String) As Boolean
    Dim status As Long
    On Error GoTo Fail
    m_appTable = appTable: m_appCode = appCode
    status = PersonnelStatus(appTable)
    If status <> PERSONNEL_READY Then
        m_setupMode = status
        CloseSplash
        If Not RunDialog(SETUP_FORM) Then Exit Function
        If PersonnelStatus(appTable) <> PERSONNEL_READY Then Exit Function
    End If
    If SignedInPersonID() > 0 Then KitPersonnelStart = True: Exit Function    ' the setup's creator
    If PersonnelSignInWindows(appCode) Then KitPersonnelStart = True: Exit Function
    m_signInWhy = "notset"
    CloseSplash
    KitPersonnelStart = RunDialog(SIGNIN_FORM)
    Exit Function
Fail:
    HandleError "modPersonnelUI.KitPersonnelStart", appCode, "The Personnel directory couldn't be opened."
End Function

' Sign out and show the sign-in screen. False when the user closed the app
' instead of signing in.
Public Function KitPersonnelSignOut(ByVal appCode As String) As Boolean
    On Error GoTo Fail
    m_appCode = appCode
    PersonnelSignOut
    m_signInWhy = "signout"
    KitPersonnelSignOut = RunDialog(SIGNIN_FORM)
    Exit Function
Fail:
    HandleError "modPersonnelUI.KitPersonnelSignOut", appCode
End Function

' For the dialogs: the app's code and linked table, why the setup screen is
' open (PERSONNEL_NOT_SET_UP or PERSONNEL_LOST), why the sign-in screen is
' ("notset": this Windows account isn't set up; "signout").
Public Function PersonnelUiAppCode() As String
    PersonnelUiAppCode = m_appCode
End Function

Public Function PersonnelUiAppTable() As String
    PersonnelUiAppTable = m_appTable
End Function

Public Function PersonnelUiSetupMode() As Long
    PersonnelUiSetupMode = m_setupMode
End Function

Public Function PersonnelUiSignInWhy() As String
    PersonnelUiSignInWhy = m_signInWhy
End Function

' What a person who isn't the admin sends to their admin.
Public Function PersonnelAdminNote() As String
    Dim dataFile As String
    dataFile = KitBackendPath(m_appTable)
    If Len(dataFile) = 0 Then dataFile = CurrentProject.FullName
    PersonnelAdminNote = APP_NAME & " can't find its Personnel file." & vbCrLf & _
        "App data file: " & dataFile & vbCrLf & _
        "Personnel file looked for: " & IIf(Len(PersonnelPath()) > 0, PersonnelPath(), PersonnelDefaultPath(m_appTable)) & vbCrLf & _
        "Computer: " & WinComputer() & vbCrLf & _
        "Windows account: " & WinAccount()
End Function

' ---------------------------------------------------------------------------
' Build
' ---------------------------------------------------------------------------

' Build step: dlgPersonnelSetup and dlgSignIn (code in kit\forms).
Public Sub BuildPersonnelDialogs()
    Build_dlgPersonnelSetup
    Build_dlgSignIn
End Sub

' dlgPersonnelSetup: three panels in one dialog, shown one at a time by Tag
' (the dialog sizes itself to the panel):
'   choice  why it shows and three choices, each with a one-line hint:
'           Find the Personnel file, Create a new Personnel file (never while
'           one is linked), I'm not the admin
'   create  the first admin's name and rank, the folder (Change opens the
'           folder picker) and their app password twice
'   ask     what to send the admin, with Copy, and Close app
Private Sub Build_dlgPersonnelSetup()
    Const DLG_W As Long = 7200          ' 480 px
    Dim c As Control, pad As Long, w As Long, x As Long, y As Long, half As Long, gap As Long, btnW As Long
    pad = Sp(SP_DIALOG): gap = Sp(SP_ROW_GAP)
    x = pad: w = DLG_W - 2 * pad: half = (w - Sp(SP_ITEM_GAP)) \ 2: btnW = px(96)
    NewDialog "Personnel", DLG_W, px(480)
    Ev F, "OnLoad"
    y = pad
    Lbl "lblTitle", "Set up the Personnel directory", x, y, w, px(26), 13, True
    y = y + px(26) + gap
    Set c = Lbl("lblWhy", "", x, y, w, px(56))            ' filled in by Form_Load
    c.Tag = "choice"
    ' choice
    y = y + px(56) + gap
    ChoiceRow "Find", "Find the Personnel file", _
        "When Personnel is already set up and you know where its file is.", x, y, w
    ChoiceRow "Create", "Create a new Personnel file", _
        "For the first app set up here. You become the Personnel admin.", x, y, w
    ChoiceRow "NotAdmin", "I'm not the admin", _
        "Shows what to send your admin, then closes " & APP_NAME & ".", x, y, w
    ' create: starts under the title
    y = pad + px(26) + gap
    FieldLabel "lblFirst", "First name", x, y, half, "create"
    FieldLabel "lblLast", "Last name", x + half + Sp(SP_ITEM_GAP), y, half, "create"
    y = y + px(18)
    Txt("txtFirst", "", x, y, half, px(30)).Tag = "create"
    Txt("txtLast", "", x + half + Sp(SP_ITEM_GAP), y, half, px(30)).Tag = "create"
    y = y + px(30) + gap
    FieldLabel "lblRank", "Rank (optional)", x, y, half, "create"
    y = y + px(18)
    Txt("txtRank", "", x, y, half, px(30)).Tag = "create"
    y = y + px(30) + gap
    FieldLabel "lblFolder", "Folder for the Personnel file", x, y, w, "create"
    y = y + px(18)
    Set c = RO("txtFolder", "", x, y, w - btnW - Sp(SP_ITEM_GAP), px(30), 10)
    c.BackStyle = 1: c.BackColor = CLR_SURFACE: c.BorderStyle = 1: c.BorderColor = CLR_BORDER
    c.LeftMargin = px(8)
    c.Tag = "create"
    Btn("btnFolder", "Change...", x + w - btnW, y, btnW, px(30), "outline").Tag = "create"
    y = y + px(30) + px(4)
    Lbl("lblFolderHint", "Only Personnel admins should be able to change this folder. Ask IT to restrict it.", _
        x, y, w, px(18), 9, False, CLR_MUTED_FG).Tag = "create"
    y = y + px(18) + gap
    FieldLabel "lblPwd", "Your app password (8 or more characters)", x, y, half, "create"
    FieldLabel "lblPwd2", "Type it again", x + half + Sp(SP_ITEM_GAP), y, half, "create"
    y = y + px(18)
    Set c = Txt("txtPwd", "", x, y, half, px(30)): c.InputMask = SECRET_MASK: c.Tag = "create"
    Set c = Txt("txtPwd2", "", x + half + Sp(SP_ITEM_GAP), y, half, px(30)): c.InputMask = SECRET_MASK: c.Tag = "create"
    y = y + px(30) + Sp(SP_DIALOG)
    Set c = Btn("btnCreateGo", "Create", x + w - 2 * btnW - Sp(SP_ITEM_GAP), y, btnW, px(32)): c.Tag = "create"
    Set c = Btn("btnBack", "Back", x + w - btnW, y, btnW, px(32), "outline"): c.Tag = "create"
    ' ask: starts under the title
    y = pad + px(26) + gap
    Lbl("lblAsk", "Send this to your Personnel admin. They can find the file, or set Personnel up.", _
        x, y, w, px(36)).Tag = "ask"
    y = y + px(36) + gap
    Set c = RO("txtAsk", "", x, y, w, px(110), 9)
    c.BackStyle = 1: c.BackColor = CLR_SURFACE: c.BorderStyle = 1: c.BorderColor = CLR_BORDER
    c.LeftMargin = px(8): c.TopMargin = px(6): c.RightMargin = px(8)
    c.Tag = "ask"
    y = y + px(110) + Sp(SP_DIALOG)
    Set c = Btn("btnCopy", "Copy", x + w - 2 * btnW - Sp(SP_ITEM_GAP), y, btnW, px(32), "outline"): c.Tag = "ask"
    Set c = Btn("btnCloseApp", "Close app", x + w - btnW, y, btnW, px(32)): c.Tag = "ask"
    ' Tall enough for the tallest panel; Form_Load fits it to the one shown.
    For Each c In F.Controls
        If c.Top + c.Height > y Then y = c.Top + c.Height
    Next
    F.Section(acDetail).Height = y + pad
    SaveForm SETUP_FORM
End Sub

' One choice: an outline button the width of the dialog and a muted hint
' under it; moves y below them.
Private Sub ChoiceRow(ByVal key As String, ByVal cap As String, ByVal hint As String, ByVal x As Long, y As Long, ByVal w As Long)
    Dim c As Control
    Set c = Btn("btn" & key, cap, x, y, w, px(36), "outline")
    c.Tag = "choice"
    y = y + px(36) + px(4)
    Lbl("lbl" & key & "Hint", hint, x, y, w, px(18), 9, False, CLR_MUTED_FG).Tag = "choice"
    y = y + px(18) + Sp(SP_ROW_GAP)
End Sub

Private Sub FieldLabel(ByVal nm As String, ByVal cap As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal tg As String)
    Lbl(nm, cap, x, y, w, px(16), 9, True, CLR_MUTED_FG).Tag = tg
End Sub

' dlgSignIn: why it shows (this Windows account isn't set up, with the
' account and Copy; or signed out), sign-in name and app password, "Use my
' Windows sign-in next time" (only while this Windows account isn't saved
' for anyone), a line for what went wrong, and Sign in / Close app. Rows
' that don't apply are hidden and the rows under them move up (Form_Load).
Private Sub Build_dlgSignIn()
    Const DLG_W As Long = 6000          ' 400 px
    Dim c As Control, pad As Long, w As Long, x As Long, y As Long, gap As Long, btnW As Long
    pad = Sp(SP_DIALOG): gap = Sp(SP_ROW_GAP)
    x = pad: w = DLG_W - 2 * pad: btnW = px(96)
    NewDialog "Sign in", DLG_W, px(400)
    F.CloseButton = False                           ' Close app or Sign in ends it
    Ev F, "OnLoad"
    y = pad
    Lbl "lblTitle", "Sign in to " & APP_NAME, x, y, w, px(26), 13, True
    y = y + px(26) + gap
    ' Row "why": this Windows account isn't set up.
    Lbl("lblWhy", "This Windows account isn't set up in Personnel yet. Ask a Personnel admin to add it, " & _
        "or sign in with your sign-in name and app password.", x, y, w, px(36), 9).Tag = "why"
    y = y + px(36) + px(4)
    Set c = RO("txtAccount", "", x, y, w - px(72) - Sp(SP_ITEM_GAP), px(28), 9)
    c.BackStyle = 1: c.BackColor = CLR_SURFACE: c.BorderStyle = 1: c.BorderColor = CLR_BORDER
    c.LeftMargin = px(8): c.Tag = "why"
    Btn("btnCopyAccount", "Copy", x + w - px(72), y, px(72), px(28), "outline").Tag = "why"
    y = y + px(28) + gap
    Lbl "lblUser", "Sign-in name", x, y, w, px(16), 9, True, CLR_MUTED_FG
    y = y + px(18)
    Txt "txtUser", "", x, y, w, px(30)
    y = y + px(30) + gap
    Lbl "lblPwd", "App password", x, y, w, px(16), 9, True, CLR_MUTED_FG
    y = y + px(18)
    Set c = Txt("txtPwd", "", x, y, w, px(30)): c.InputMask = SECRET_MASK
    y = y + px(30) + gap
    ' Row "win": save this Windows account to the person signing in.
    Chk("chkWindows", "", x, y + px(2)).Tag = "win"
    Lbl("lblWindows", "Use my Windows sign-in next time", x + px(24), y, w - px(24), px(20), 9).Tag = "win"
    y = y + px(20) + gap
    Set c = Lbl("lblError", "", x, y, w, px(36), 9, False, CLR_ERROR)
    y = y + px(36) + gap
    Set c = Btn("btnSignIn", "Sign in", x + w - 2 * btnW - Sp(SP_ITEM_GAP), y, btnW, px(32)): c.Default = True
    Set c = Btn("btnClose", "Close app", x + w - btnW, y, btnW, px(32), "outline"): c.Cancel = True
    F.Section(acDetail).Height = y + px(32) + pad
    SaveForm SIGNIN_FORM
End Sub
