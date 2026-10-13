Option Compare Database
Option Explicit

' Access UI kit - trusted-folder check at startup. When the app's folder is
' not a Trusted Location, Access asks the user to enable content every time
' the app opens (or opens it with its code turned off). BuildTrustCheck gives
' the app a startup that notices and shows dlgTrust, a plain page explaining
' how to add the folder in the Trust Center. Nothing here changes Trust Center
' settings: the registry is only read.
'
' How startup works (BuildTrustCheck sets it up):
'   AutoExec macro   runs instead of a StartUpForm, because a macro runs even
'                    when the code is turned off.
'     code off       (CurrentProject.IsTrusted False): opens dlgTrust and stops.
'     code on        runs TrustStartup: opens the main form, then dlgTrust over
'                    it when the folder still isn't a Trusted Location (the
'                    user chose Enable Content, which only trusts this one file).
'   dlgTrust         has no code and no Open hook, so it works with the code
'                    off. Its Close button runs the macro mcrTrustClose.

Private Const TRUST_FORM As String = "dlgTrust"
Private Const CLOSE_MACRO As String = "mcrTrustClose"
Private Const MAIN_FORM_PROP As String = "KitMainForm"    ' database property: the form TrustStartup opens

Private Const HKEY_CURRENT_USER As LongPtr = &H80000001
Private Const KEY_READ As Long = &H20019
Private Const ERROR_SUCCESS As Long = 0
Private Const REG_DWORD As Long = 4
Private Const DRIVE_REMOTE As Long = 4

Private Declare PtrSafe Function RegOpenKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal subKey As LongPtr, _
    ByVal options As Long, ByVal samDesired As Long, phkResult As LongPtr) As Long
Private Declare PtrSafe Function RegEnumKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal index As Long, _
    ByVal name As LongPtr, cchName As Long, ByVal reserved As LongPtr, ByVal cls As LongPtr, _
    ByVal cchClass As LongPtr, ByVal lastWrite As LongPtr) As Long
Private Declare PtrSafe Function RegQueryValueExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal valueName As LongPtr, _
    ByVal reserved As LongPtr, valueType As Long, ByVal data As LongPtr, cbData As Long) As Long
Private Declare PtrSafe Function RegCloseKey Lib "advapi32" (ByVal hKey As LongPtr) As Long
Private Declare PtrSafe Function ExpandEnvironmentStringsW Lib "kernel32" (ByVal src As LongPtr, _
    ByVal dst As LongPtr, ByVal size As Long) As Long
Private Declare PtrSafe Function GetDriveTypeW Lib "kernel32" (ByVal root As LongPtr) As Long

' ---------------------------------------------------------------------------
' Build
' ---------------------------------------------------------------------------

' Build step, in place of setting StartUpForm: builds dlgTrust and the macros
' AutoExec and mcrTrustClose, and remembers mainForm for TrustStartup. Any
' StartUpForm set earlier is removed, so mainForm doesn't open twice.
Public Sub BuildTrustCheck(ByVal mainForm As String)
    Dim db As DAO.Database, n As Long, d As String
    Build_dlgTrust
    LoadMacro CLOSE_MACRO, MacroRow("", "Close", "2", TRUST_FORM, "2")      ' form dlgTrust, don't save
    LoadMacro "AutoExec", _
        MacroRow("Not [CurrentProject].[IsTrusted]", "OpenForm", TRUST_FORM, "0") & _
        MacroRow("Not [CurrentProject].[IsTrusted]", "StopMacro") & _
        MacroRow("", "RunCode", "TrustStartup()")
    Set db = CurrentDb
    On Error Resume Next
    db.Properties(MAIN_FORM_PROP) = mainForm
    If Err.Number = 3270 Then           ' property not found: create it
        Err.Clear
        db.Properties.Append db.CreateProperty(MAIN_FORM_PROP, dbText, mainForm)
    End If
    If Err.Number <> 0 Then
        n = Err.Number: d = Err.Description
        On Error GoTo 0
        Err.Raise n, , "Could not save the main form name: " & d
    End If
    Err.Clear
    db.Properties.Delete "StartUpForm"  ' error 3265 when there is none
    On Error GoTo 0
End Sub

' dlgTrust: heading, why it shows, the folder (selectable, to copy), the Trust
' Center steps, and a note for computers where IT manages the settings. Text
' heights come from MeasureText, so each block sits Sp(SP_ROW_GAP) under the
' one above. Built with no code module and no Open hook (it must open with the
' code turned off); ShowTrustHelp themes it when the code is on.
Private Sub Build_dlgTrust()
    Dim c As Control, pad As Long, w As Long, x As Long, y As Long, h As Long, s As String
    Const DLG_W As Long = 7200          ' 480 px
    pad = Sp(SP_DIALOG)
    x = pad: w = DLG_W - 2 * pad
    NewDialog APP_NAME, DLG_W, px(400)
    F.OnOpen = ""
    F.HasModule = False
    y = pad
    Lbl "lblTitle", "Trust this app's folder", x, y, w, px(26), 13, True
    y = y + px(26) + Sp(SP_ROW_GAP)
    s = "Access doesn't trust the folder this app is in, so it asks you to enable content, or turns the " & _
        "app's code off, each time the app opens. Trust the folder once and the app opens without asking."
    h = TextH(s, 10, w)
    Lbl "lblWhy", s, x, y, w, h
    y = y + h + Sp(SP_ROW_GAP) * 2
    Lbl "lblFolderCap", "FOLDER", x, y, w, px(16), 8, True, CLR_MUTED_FG
    y = y + px(18)
    Set c = RO("txtFolder", "=[CurrentProject].[Path]", x, y, w, px(44), 10)
    c.BackStyle = 1: c.BackColor = CLR_SURFACE
    c.BorderStyle = 1: c.BorderColor = CLR_BORDER
    c.LeftMargin = px(8): c.TopMargin = px(4): c.RightMargin = px(8)
    y = y + px(44) + Sp(SP_ROW_GAP) * 2
    Lbl "lblStepsCap", "TO TRUST IT", x, y, w, px(16), 8, True, CLR_MUTED_FG
    y = y + px(18)
    s = "1. In Access, choose File, then Options, then Trust Center, then Trust Center Settings." & vbCrLf & _
        "2. Choose Trusted Locations, then Add new location." & vbCrLf & _
        "3. Browse to the folder above and choose OK." & vbCrLf & _
        "4. If the folder is on a network drive, first tick Allow Trusted Locations on my network." & vbCrLf & _
        "5. Close Access and open the app again."
    h = TextH(s, 10, w)
    Lbl "lblSteps", s, x, y, w, h
    y = y + h + Sp(SP_ROW_GAP)
    s = "If these settings are locked on your computer, ask your IT support to add the folder."
    h = TextH(s, 9, w)
    Lbl "lblIT", s, x, y, w, h, 9, False, CLR_MUTED_FG
    y = y + h + Sp(SP_DIALOG)
    Set c = Btn("btnClose", "Close", DLG_W - pad - px(96), y, px(96), px(32))
    c.OnClick = CLOSE_MACRO             ' a macro, not code: works with the code off
    c.Default = True: c.Cancel = True
    F.Section(acDetail).Height = y + px(32) + pad
    SaveForm TRUST_FORM
End Sub

' Height in twips that s takes in a label w wide at size points: each line
' (split at vbCrLf) wraps to as many lines as its measured width needs, with
' 10% spare for word breaks, as dlgMessage measures its message.
Private Function TextH(ByVal s As String, ByVal size As Integer, ByVal w As Long) As Long
    Dim para As Variant, lines As Long
    For Each para In Split(s, vbCrLf)
        If Len(para) = 0 Then
            lines = lines + 1
        Else
            lines = lines - Int(-(MeasureText(CStr(para), size, 400) * 1.1) / w)
        End If
    Next
    If lines < 1 Then lines = 1
    TextH = lines * CLng(size * 20 * 1.36) + 30
End Function

' One row of a macro in Access's text format (LoadFromText acMacro), with an
' optional condition.
Private Function MacroRow(ByVal cond As String, ByVal action As String, ParamArray args() As Variant) As String
    Dim s As String, i As Long
    s = "Begin" & vbCrLf
    If Len(cond) > 0 Then s = s & "    Condition =""" & cond & """" & vbCrLf
    s = s & "    Action =""" & action & """" & vbCrLf
    For i = 0 To UBound(args)
        s = s & "    Argument =""" & args(i) & """" & vbCrLf
    Next
    MacroRow = s & "End" & vbCrLf
End Function

' Replace macro name with the given rows (a temporary file, as LoadFromText
' needs one).
Private Sub LoadMacro(ByVal name As String, ByVal rows As String)
    Dim path As String, f As Integer
    path = Environ$("TEMP") & "\kit_macro_" & name & ".txt"
    f = FreeFile
    Open path For Output As #f
    Print #f, "Version =196611" & vbCrLf & "ColumnsShown =2" & vbCrLf & rows;
    Close #f
    On Error Resume Next
    DoCmd.DeleteObject acMacro, name
    On Error GoTo 0
    Application.LoadFromText acMacro, name, path
    Kill path
End Sub

' ---------------------------------------------------------------------------
' Run time
' ---------------------------------------------------------------------------

' AutoExec, with the code on: open the main form, then the trust page over it
' when the folder isn't a Trusted Location. Does nothing under automation
' (a build script opening the file), so nothing opens or waits there.
Public Function TrustStartup()
    Dim mainForm As String
    On Error GoTo Err_Handler
    If Not Application.UserControl Then Exit Function
    On Error Resume Next
    mainForm = CurrentDb.Properties(MAIN_FORM_PROP)
    On Error GoTo Err_Handler
    If Len(mainForm) > 0 Then DoCmd.OpenForm mainForm
    If Not InTrustedLocation() Then ShowTrustHelp
    Exit Function
Err_Handler:
    HandleError "modTrust.TrustStartup", "main form " & mainForm
End Function

' Open the trust page (dlgTrust) in the current palette. It has no Open hook,
' so it opens hidden, is themed, then shown.
Public Function ShowTrustHelp()
    On Error GoTo Err_Handler
    DoCmd.OpenForm TRUST_FORM, WindowMode:=acHidden
    ThemeForm Forms(TRUST_FORM)
    Forms(TRUST_FORM).Visible = True
    Exit Function
Err_Handler:
    HandleError "modTrust.ShowTrustHelp"
End Function

' True when folder (default: the app's own) is inside a Trusted Location
' Access will honor: listed in the user's or the policy Trusted Locations,
' with "subfolders" ticked when folder is below it, network locations allowed
' when it is on a network drive, and trusted locations not turned off by
' policy. False if the settings can't be read, so the worst case is showing
' the trust page when it wasn't needed.
Public Function InTrustedLocation(Optional ByVal folder As String = "") As Boolean
    Dim roots As Variant, root As Variant, allowNet As Boolean, isNet As Boolean, onList As Boolean
    On Error GoTo Fail
    If Len(folder) = 0 Then folder = CurrentProject.Path
    folder = NormPath(folder)
    roots = Array("Software\Policies\Microsoft\Office\" & Application.Version & "\Access\Security\Trusted Locations", _
                  "Software\Microsoft\Office\" & Application.Version & "\Access\Security\Trusted Locations")
    If RegDword(roots(0), "", "AllLocationsDisabled", 0) = 1 Then Exit Function
    ' A policy setting wins over the user's own.
    allowNet = (RegDword(roots(0), "", "AllowNetworkLocations", RegDword(roots(1), "", "AllowNetworkLocations", 0)) = 1)
    isNet = (Left$(folder, 2) = "\\")
    If Not isNet And Mid$(folder, 2, 1) = ":" Then isNet = (GetDriveTypeW(StrPtr(Left$(folder, 3))) = DRIVE_REMOTE)
    If isNet And Not allowNet Then Exit Function
    For Each root In roots
        If ListHas(CStr(root), folder) Then onList = True: Exit For
    Next
    InTrustedLocation = onList
    Exit Function
Fail:
    LogError "modTrust.InTrustedLocation", "folder " & folder
End Function

' True when one of the Location keys under root covers folder.
Private Function ListHas(ByVal root As String, ByVal folder As String) As Boolean
    Dim hKey As LongPtr, i As Long, nm As String, n As Long, p As String, subs As Boolean
    If RegOpenKeyExW(HKEY_CURRENT_USER, StrPtr(root), 0, KEY_READ, hKey) <> ERROR_SUCCESS Then Exit Function
    Do
        nm = String$(256, vbNullChar): n = 256
        If RegEnumKeyExW(hKey, i, StrPtr(nm), n, 0, 0, 0, 0) <> ERROR_SUCCESS Then Exit Do
        nm = Left$(nm, n)
        p = RegString(root, nm, "Path")
        If Len(p) > 0 Then
            p = NormPath(p)
            subs = (RegDword(root, nm, "AllowSubfolders", 0) = 1)
            If folder = p Or (subs And Left$(folder, Len(p) + 1) = p & "\") Then ListHas = True: Exit Do
        End If
        i = i + 1
    Loop
    RegCloseKey hKey
End Function

' Lower case, environment variables expanded, no trailing backslash.
Private Function NormPath(ByVal p As String) As String
    Dim buf As String, n As Long
    If InStr(p, "%") > 0 Then
        buf = String$(1024, vbNullChar)
        n = ExpandEnvironmentStringsW(StrPtr(p), StrPtr(buf), 1024)
        If n > 0 And n <= 1024 Then p = Left$(buf, n - 1)
    End If
    p = LCase$(Trim$(p))
    Do While Len(p) > 3 And Right$(p, 1) = "\"
        p = Left$(p, Len(p) - 1)
    Loop
    NormPath = p
End Function

' HKCU\root\subKey value name as text ("" when missing).
Private Function RegString(ByVal root As String, ByVal subKey As String, ByVal name As String) As String
    Dim hKey As LongPtr, kind As Long, cb As Long, buf As String
    If RegOpenKeyExW(HKEY_CURRENT_USER, StrPtr(KeyPath(root, subKey)), 0, KEY_READ, hKey) <> ERROR_SUCCESS Then Exit Function
    If RegQueryValueExW(hKey, StrPtr(name), 0, kind, 0, cb) = ERROR_SUCCESS And cb > 0 And kind <> REG_DWORD Then
        buf = String$(cb \ 2 + 1, vbNullChar)
        cb = Len(buf) * 2
        If RegQueryValueExW(hKey, StrPtr(name), 0, kind, StrPtr(buf), cb) = ERROR_SUCCESS Then
            CutAtNull buf
            RegString = buf
        End If
    End If
    RegCloseKey hKey
End Function

' Cut s at its first null character.
Private Sub CutAtNull(s As String)
    Dim i As Long
    i = InStr(s, vbNullChar)
    If i > 0 Then s = Left$(s, i - 1)
End Sub

' HKCU\root\subKey value name as a number (dflt when missing or not a DWORD).
Private Function RegDword(ByVal root As String, ByVal subKey As String, ByVal name As String, ByVal dflt As Long) As Long
    Dim hKey As LongPtr, kind As Long, cb As Long, v As Long
    RegDword = dflt
    If RegOpenKeyExW(HKEY_CURRENT_USER, StrPtr(KeyPath(root, subKey)), 0, KEY_READ, hKey) <> ERROR_SUCCESS Then Exit Function
    cb = 4
    If RegQueryValueExW(hKey, StrPtr(name), 0, kind, VarPtr(v), cb) = ERROR_SUCCESS Then
        If kind = REG_DWORD Then RegDword = v
    End If
    RegCloseKey hKey
End Function

Private Function KeyPath(ByVal root As String, ByVal subKey As String) As String
    KeyPath = root & IIf(Len(subKey) > 0, "\" & subKey, "")
End Function
