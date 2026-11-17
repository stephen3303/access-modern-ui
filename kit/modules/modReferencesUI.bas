Option Compare Database
Option Explicit

' Access UI kit - the shared Reference library's screens (the data and its
' rules are in modReferences).
'
' Startup, after KitPersonnelStart:
'   KitReferencesStart APP_CODE
' links the library file beside the Personnel file when it is there. It asks
' nothing: an app that cites references works without the library until it
' is set up, and the library page says what is missing.
'
' A "Reference library" link in the sidebar or a menu:
'   =KitReferencesManage()
' opens the library page (dlgReferences): search, Show retired, Add reference
' (anyone signed in), Change (Library editors and Personnel admins: fields,
' status, superseded by, history), Open, and Library editors (Personnel admins,
' dlgRefEditors). When the library isn't set up yet, a Personnel admin is
' offered to make it beside the Personnel file, or to find a moved one.
'
' Citing a reference on an app's form (the app keeps ref_id and the page in
' its own data file, as the web's comment_citation table does):
'   RefPickBox "txtRef", "ref_id", x, y, w      a searchable picker of the active references
'   id = KitRefAdd("typed title")                Add a missing one (returns its ref_id, or 0)
'   =KitRefOpen([ref_id], [page_number])         "Open p. N": the document at that page
'   RefOpenCaption(page)                         the button's caption
' Wire the picker's KeyDown in the form's code as PickBox (modMenu) says.
'
' Build the dialogs with the app's others: BuildReferenceDialogs, after
' BuildPickDialog (the pickers) and BuildDatePicker (the date box). Help
' topics for them are in kit\help\references.txt, under the key reflib.

Public Const REFS_FORM As String = "dlgReferences"
Public Const REF_FORM As String = "dlgReference"
Public Const REF_EDITORS_FORM As String = "dlgRefEditors"

Private m_appCode As String

' ---------------------------------------------------------------------------
' Startup and opening
' ---------------------------------------------------------------------------

' Link the library when it is beside the Personnel file (ReferencesStatus)
' and remember the app for the audit. Returns the status; never asks
' anything and never stops the app.
Public Function KitReferencesStart(ByVal appCode As String) As Long
    m_appCode = appCode
    On Error GoTo Fail
    KitReferencesStart = ReferencesStatus()
    Exit Function
Fail:
    LogError "modReferencesUI.KitReferencesStart"
    KitReferencesStart = REFS_LOST
End Function

' The app the library is used from (KitReferencesStart), for the audit.
Public Function RefUiAppCode() As String
    RefUiAppCode = m_appCode
End Function

' Open the library page, setting the library up first when it isn't (see the
' top of the module). For a sidebar link or a button's OnClick.
Public Function KitReferencesManage()
    On Error GoTo Fail
    If Not KitReferencesEnsure() Then Exit Function
    RunDialog REFS_FORM
    Exit Function
Fail:
    HandleError "modReferencesUI.KitReferencesManage"
End Function

' True when the library can be used now. Otherwise says why, and offers a
' Personnel admin to make the file (not set up) or find it (moved); True
' once that worked.
Public Function KitReferencesEnsure() As Boolean
    Dim st As Long, def As String, path As String, problem As String
    On Error GoTo Fail
    st = ReferencesStatus()
    Select Case st
        Case REFS_READY
            KitReferencesEnsure = True
        Case REFS_NO_PERSONNEL
            KitMsgBox "The Reference library is kept beside the Personnel file, and " & APP_NAME & _
                " isn't linked to a Personnel file.", vbExclamation, "Reference library"
        Case REFS_NOT_SET_UP
            def = ReferencesDefaultPath()
            If Not PersonnelIsAdminNow() Then
                KitMsgBox "The shared Reference library isn't set up yet. Ask a Personnel admin to open it once: " & _
                    "it is made beside the Personnel file, at" & vbCrLf & def, vbInformation, "Reference library"
                Exit Function
            End If
            If KitMsgBox("The shared Reference library isn't set up yet. Make it now, beside the Personnel file?" & vbCrLf & vbCrLf & _
                    def & vbCrLf & vbCrLf & "Every app (Access and web) uses this one list.", vbQuestion + vbYesNo, _
                    "Reference library") <> vbYes Then Exit Function
            CreateReferencesFile def
            KitToast "Reference library made", "success"
            KitReferencesEnsure = True
        Case REFS_LOST
            If Not PersonnelIsAdminNow() Then
                KitMsgBox "The Reference library can't be reached at" & vbCrLf & ReferencesPath() & vbCrLf & vbCrLf & _
                    "The drive may not be connected on this PC. If it was moved, ask a Personnel admin.", vbExclamation, "Reference library"
                Exit Function
            End If
            If KitMsgBox("The Reference library can't be reached at" & vbCrLf & ReferencesPath() & vbCrLf & vbCrLf & _
                    "Find the file?", vbQuestion + vbYesNo, "Reference library") <> vbYes Then Exit Function
            path = KitPickFile("Find the Reference library (" & REF_FILE & ")", "Reference library|" & REF_FILE & "|Access database|*.accdb", "references")
            If Len(path) = 0 Then Exit Function
            problem = ReferencesFileProblem(path)
            If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, "Reference library": Exit Function
            LinkReferencesFile path
            KitReferencesEnsure = True
    End Select
    Exit Function
Fail:
    If Err.Number = 5 Then KitMsgBox Err.Description, vbExclamation, "Reference library": Exit Function
    HandleError "modReferencesUI.KitReferencesEnsure", , "The Reference library couldn't be opened."
End Function

' ---------------------------------------------------------------------------
' For an app's citations
' ---------------------------------------------------------------------------

' A searchable picker (PickBox, modMenu) of the active references: the box
' nm shows "NUMBER Title (REV)", with Document in muted text when one is
' linked; nm & "Key" holds the ref_id, bound to keyField ("" for none).
' Build dlgPick with the app's dialogs (BuildPickDialog), and wire the box in
' the form's code as PickBox says.
Public Function RefPickBox(ByVal nm As String, ByVal keyField As String, ByVal x As Long, ByVal y As Long, _
        ByVal w As Long, Optional ByVal h As Long = 450) As Control
    Set RefPickBox = PickBox(nm, keyField, RefPickSql(), x, y, w, h, "Type a number or title")
End Function

' Add a missing reference (dlgReference), with title typed in already.
' Returns the new ref_id, or the one already in the library when it is the
' same document, or 0 when the person closed it without saving.
Public Function KitRefAdd(Optional ByVal title As String = "") As Long
    On Error GoTo Fail
    If Not KitReferencesEnsure() Then Exit Function
    TempVars("RefId") = 0
    If RunDialog(REF_FORM, "new|" & title) Then KitRefAdd = Nz(TempVars("RefId"), 0)
    Exit Function
Fail:
    HandleError "modReferencesUI.KitRefAdd"
End Function

' Open reference refId's document at page (blank or 0 for none), telling the
' person when it can't, or that it opened at page 1. For a button's OnClick:
'   =KitRefOpen([ref_id],[page_number])
Public Function KitRefOpen(ByVal refId As Variant, Optional ByVal page As Variant = Null)
    Dim note As String, pg As Long
    On Error GoTo Fail
    If IsNull(refId) Then KitMsgBox "Choose a reference first.", vbInformation: Exit Function
    If Not ReferencesReady() Then
        If Not KitReferencesEnsure() Then Exit Function
    End If
    If Len(RefPageProblem(page)) = 0 And Not IsNull(page) Then
        If Len(Trim$(CStr(page))) > 0 Then pg = CLng(page)
    End If
    note = RefOpen(CLng(refId), pg)
    If Len(note) > 0 Then KitMsgBox note, vbInformation, "Open reference"
    Exit Function
Fail:
    HandleError "modReferencesUI.KitRefOpen", Nz(refId, "") & " p." & Nz(page, "")
End Function

' F1 in a library dialog (its Form_KeyDown) opens the kit help on its topic.
' KitHelp is called by name, so an app without modKitHelp still compiles.
Public Sub RefF1(KeyCode As Integer)
    If KeyCode <> vbKeyF1 Then Exit Sub
    KeyCode = 0
    On Error Resume Next                        ' no modKitHelp: F1 does nothing
    Application.Run "KitHelp"
End Sub

' ---------------------------------------------------------------------------
' Build
' ---------------------------------------------------------------------------

' Build step: dlgReferences, dlgReference and dlgRefEditors (code in kit\forms).
Public Sub BuildReferenceDialogs()
    Build_dlgReferences
    Build_dlgReference
    Build_dlgRefEditors
End Sub

' dlgReferences, the library page: a search box, Show retired, Library
' editors (Personnel admins), Add reference; the list (double-click or Change
' to change one); Open, Change and Close.
Private Sub Build_dlgReferences()
    Const DLG_W As Long = 13200         ' 880 px
    Dim c As Control, pad As Long, w As Long, x As Long, y As Long, btnW As Long, listH As Long, colW As String
    pad = Sp(SP_DIALOG)
    x = pad: w = DLG_W - 2 * pad: btnW = px(120): listH = px(360)
    NewDialog "Reference library", DLG_W, px(560)
    Ev F, "OnLoad", "OnKeyDown": F.KeyPreview = True   ' F1 (RefF1)
    y = pad
    Lbl "lblTitle", "Reference library", x, y, w - px(300), px(26), 13, True
    y = y + px(26) + px(4)
    Lbl "lblSub", "One list for every app. Anyone can add a publication; Library editors change or retire one.", _
        x, y, w, px(18), 9, False, CLR_MUTED_FG
    y = y + px(18) + Sp(16)
    Set c = Txt("txtSearch", "", x, y, px(260), px(30), 9)
    c.Format = "@;""Search number, title or notes..."""
    Ev c, "OnChange"
    Set c = Chk("chkRetired", "", x + px(276), y + px(6))
    Ev c, "AfterUpdate"
    Lbl "lblRetired", "Show retired", x + px(300), y + px(4), px(110), px(22), 9
    Btn "btnEditors", "Library editors...", x + w - 2 * btnW - Sp(SP_ITEM_GAP), y, btnW, px(30), "outline"
    Btn "btnAdd", "Add reference", x + w - btnW, y, btnW, px(30)
    y = y + px(30) + Sp(SP_ROW_GAP)
    ' ref_id, Number, Title, Revision, Date, Status, Document
    colW = "0;" & px(170) & ";" & px(330) & ";" & px(70) & ";" & px(90) & ";" & px(90) & ";" & px(70)
    Set c = Lst("lstRefs", "", x, y, w, listH, 7, colW)
    Ev c, "OnDblClick"
    Set c = Lbl("lblNone", "No references match your search.", x + px(10), y + px(30), px(360), px(20), 9, False, CLR_MUTED_FG)
    c.Visible = False
    y = y + listH + Sp(16)
    Lbl "lblTip", "Double-click a reference to change it. Open shows its document.", _
        x, y, w - 3 * px(96) - 2 * Sp(SP_ITEM_GAP) - Sp(16), px(30), 8, False, CLR_MUTED_FG
    Btn "btnOpen", "Open", x + w - 3 * px(96) - 2 * Sp(SP_ITEM_GAP), y, px(96), px(30), "outline"
    Btn "btnEdit", "Change", x + w - 2 * px(96) - Sp(SP_ITEM_GAP), y, px(96), px(30), "outline"
    Set c = Btn("btnClose", "Close", x + w - px(96), y, px(96), px(30), "outline"): c.Cancel = True
    F.Section(acDetail).Height = y + px(30) + pad
    SaveForm REFS_FORM
End Sub

' dlgReference: add or change one reference. OpenArgs "new", "new|<title>"
' or a ref_id. Number and revision; title; date and marking; the document
' link with Browse... and Open; notes; for a change, status and Superseded by
' (a reference picker) and History; a line for what went wrong; Save and Close.
Private Sub Build_dlgReference()
    Const DLG_W As Long = 9000          ' 600 px
    Dim c As Control, pad As Long, w As Long, x As Long, x2 As Long, y As Long, half As Long, gap As Long, btnW As Long
    pad = Sp(SP_DIALOG): gap = Sp(SP_ROW_GAP)
    x = pad: w = DLG_W - 2 * pad: half = (w - Sp(16)) \ 2: x2 = x + half + Sp(16): btnW = px(96)
    NewDialog "Reference", DLG_W, px(700)
    Ev F, "OnLoad", "OnOpen", "OnKeyDown": F.KeyPreview = True   ' Form_Open checks who may change; F1 (RefF1)
    y = pad
    Lbl "lblTitle", "Add reference", x, y, w, px(26), 13, True
    y = y + px(26) + Sp(16)
    RefFieldLabel "lblNumber", "Number (such as OPNAVINST 3120.32D)", x, y, half
    RefFieldLabel "lblRevision", "Revision or change (optional)", x2, y, half
    y = y + px(18)
    Txt "txtNumber", "", x, y, half, px(30)
    Txt "txtRevision", "", x2, y, half, px(30)
    y = y + px(30) + gap
    RefFieldLabel "lblRefTitle", "Title", x, y, w
    y = y + px(18)
    Txt "txtTitle", "", x, y, w, px(30)
    y = y + px(30) + gap
    RefFieldLabel "lblDate", "Date (optional)", x, y, half
    RefFieldLabel "lblMarking", "Marking", x2, y, half
    y = y + px(18)
    DateBox "txtDate", "", x, y, px(30)
    Set c = Cbo("cboMarking", "", REF_MARKING_LEVELS, x2, y, half, px(30))
    c.LimitToList = False                           ' a REL TO marking is typed
    y = y + px(30) + gap
    RefFieldLabel "lblLink", "Document (a file on the shared drive, or a web address)", x, y, w
    y = y + px(18)
    Txt "txtLink", "", x, y, w - 2 * btnW - 2 * Sp(SP_ITEM_GAP), px(30)
    Btn "btnBrowse", "Browse...", x + w - 2 * btnW - Sp(SP_ITEM_GAP), y, btnW, px(30), "outline"
    Btn "btnOpenDoc", "Open", x + w - btnW, y, btnW, px(30), "outline"
    y = y + px(30) + px(4)
    Lbl "lblLinkHint", "Choose the copy on the shared drive, so everyone can open it. A PDF opens at a citation's page.", _
        x, y, w, px(16), 8, False, CLR_MUTED_FG
    y = y + px(16) + gap
    RefFieldLabel "lblNotes", "Notes (optional)", x, y, w
    y = y + px(18)
    Set c = Txt("txtNotes", "", x, y, w, px(60))
    c.EnterKeyBehavior = True: c.ScrollBars = 2
    y = y + px(60) + gap
    ' Status row: only when changing (Form_Load hides it while adding).
    RefFieldLabel("lblStatus", "Status", x, y, half).Tag = "change"
    RefFieldLabel("lblSup", "Superseded by", x2, y, half).Tag = "change"
    y = y + px(18)
    Set c = Cbo("cboStatus", "", REF_ACTIVE & ";" & REF_SUPERSEDED & ";" & REF_CANCELLED, x, y, half, px(30))
    c.Tag = "change"
    Ev c, "AfterUpdate"
    Set c = RefPickBox("txtSup", "", x2, y, half)
    c.Tag = "change"
    Ev c, "OnKeyDown"
    F.Controls("txtSupKey").Tag = "change"
    y = y + px(30) + px(4)
    Lbl("lblStatusHint", "Retire a reference rather than deleting it: citations to it keep working.", _
        x, y, w, px(16), 8, False, CLR_MUTED_FG).Tag = "change"
    y = y + px(16) + gap
    Lbl "lblError", "", x, y, w, px(36), 9, False, CLR_ERROR
    y = y + px(36) + gap
    Btn("btnHistory", "History", x, y, btnW, px(32), "outline").Tag = "change"
    Set c = Btn("btnSave", "Save", x + w - 2 * btnW - Sp(SP_ITEM_GAP), y, btnW, px(32)): c.Default = True
    Set c = Btn("btnClose", "Close", x + w - btnW, y, btnW, px(32), "outline"): c.Cancel = True
    F.Section(acDetail).Height = y + px(32) + pad
    SaveForm REF_FORM
End Sub

' dlgRefEditors: who may change and retire references, for Personnel admins.
' A person picker and Add; the list of editors with Remove; Close.
Private Sub Build_dlgRefEditors()
    Const DLG_W As Long = 9000          ' 600 px
    Dim c As Control, pad As Long, w As Long, x As Long, y As Long, btnW As Long, listH As Long, colW As String
    pad = Sp(SP_DIALOG)
    x = pad: w = DLG_W - 2 * pad: btnW = px(96): listH = px(240)
    NewDialog "Library editors", DLG_W, px(480)
    Ev F, "OnLoad", "OnOpen", "OnKeyDown": F.KeyPreview = True   ' Form_Open checks Personnel admin; F1 (RefF1)
    y = pad
    Lbl "lblTitle", "Library editors", x, y, w, px(26), 13, True
    y = y + px(26) + px(4)
    Lbl "lblSub", "Library editors change, retire and supersede references in every app. Personnel admins always can.", _
        x, y, w, px(32), 9, False, CLR_MUTED_FG
    y = y + px(32) + Sp(16)
    Set c = PersonPickBox("txtPerson", "", x, y, w - btnW - Sp(SP_ITEM_GAP))
    Ev c, "OnKeyDown"
    Btn "btnAdd", "Add", x + w - btnW, y, btnW, px(30)
    y = y + px(30) + Sp(SP_ROW_GAP)
    ' person_id, Name, Granted by, Granted
    colW = "0;" & px(240) & ";" & px(200) & ";" & px(90)
    Set c = Lst("lstEditors", "", x, y, w, listH, 4, colW)
    Set c = Lbl("lblNone", "No Library editors yet: only Personnel admins can change references.", x + px(10), y + px(30), w - px(20), px(20), 9, False, CLR_MUTED_FG)
    c.Visible = False
    y = y + listH + Sp(16)
    Lbl "lblError", "", x, y, w - 2 * btnW - 2 * Sp(SP_ITEM_GAP), px(30), 9, False, CLR_ERROR
    Btn "btnRemove", "Remove", x + w - 2 * btnW - Sp(SP_ITEM_GAP), y, btnW, px(30), "outline"
    Set c = Btn("btnClose", "Close", x + w - btnW, y, btnW, px(30), "outline"): c.Cancel = True
    F.Section(acDetail).Height = y + px(30) + pad
    SaveForm REF_EDITORS_FORM
End Sub

Private Function RefFieldLabel(ByVal nm As String, ByVal cap As String, ByVal x As Long, ByVal y As Long, ByVal w As Long) As Control
    Set RefFieldLabel = Lbl(nm, cap, x, y, w, px(16), 9, True, CLR_MUTED_FG)
End Function
