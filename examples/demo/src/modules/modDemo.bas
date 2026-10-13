Option Compare Database
Option Explicit

' Kit Demo - the names the kit calls (README, "What the app must provide")
' and the small run-time helpers the demo pages share. The demo exists to
' show and test the kit's builders; it has one local table, demo_item.

Public Const APP_NAME As String = "Kit Demo"
Public g_Silent As Boolean          ' True during the build: nothing may block

Private Type SIZEL
    cx As Long
    cy As Long
End Type

Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function CreateFontW Lib "gdi32" (ByVal h As Long, ByVal w As Long, ByVal e As Long, _
    ByVal o As Long, ByVal wt As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal cs As Long, _
    ByVal op As Long, ByVal cp As Long, ByVal q As Long, ByVal pf As Long, ByVal face As LongPtr) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal obj As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal obj As LongPtr) As Long
Private Declare PtrSafe Function GetTextExtentPoint32W Lib "gdi32" (ByVal hdc As LongPtr, ByVal s As LongPtr, _
    ByVal n As Long, sz As SIZEL) As Long

Private Const LOGPIXELSX As Long = 88

' ---------------------------------------------------------------------------
' Names the kit calls
' ---------------------------------------------------------------------------

' Width in twips of s in the UI font at size points and the given weight.
Public Function MeasureText(ByVal s As String, ByVal size As Long, ByVal weight As Long) As Long
    Dim hdc As LongPtr, fnt As LongPtr, old As LongPtr, dpi As Long, sz As SIZEL
    hdc = GetDC(0)
    dpi = GetDeviceCaps(hdc, LOGPIXELSX)
    fnt = CreateFontW(-CLng(size * dpi / 72), 0, 0, 0, weight, 0, 0, 0, 1, 0, 0, 5, 0, StrPtr(FONT_UI))
    old = SelectObject(hdc, fnt)
    If Len(s) > 0 Then GetTextExtentPoint32W hdc, StrPtr(s), Len(s), sz
    SelectObject hdc, old
    DeleteObject fnt
    ReleaseDC 0, hdc
    MeasureText = CLng(sz.cx * 1440& / dpi)
End Function

' Click on a "?" help mark (TitleTip): show its full explanation.
Public Function ShowTip(frm As Form, ByVal nm As String)
    On Error GoTo Err_Handler
    Notify frm(nm).Tag, vbInformation
    Exit Function
Err_Handler:
    HandleError "modDemo.ShowTip"
End Function

' Badge variant -> Array(back, fore)
Public Function BadgeColors(ByVal vrnt As String) As Variant
    Select Case vrnt
        Case "success": BadgeColors = Array(Th(CLR_SUCCESS_BG), Th(CLR_SUCCESS_FG))
        Case "warning": BadgeColors = Array(Th(CLR_WARNING_BG), Th(CLR_WARNING_FG))
        Case "info": BadgeColors = Array(Th(CLR_INFO_BG), Th(CLR_INFO_FG))
        Case "destructive": BadgeColors = Array(Th(CLR_DESTRUCTIVE), Th(CLR_WHITE))
        Case "default": BadgeColors = Array(Th(CLR_PRIMARY), Th(CLR_WHITE))
        Case "outline": BadgeColors = Array(Th(CLR_SURFACE), Th(CLR_FG))
        Case Else: BadgeColors = Array(Th(CLR_SECONDARY_BG), Th(CLR_SECONDARY_FG))
    End Select
End Function

' Badge variant for an item status (status pill columns in BuildGrid).
Public Function StatusVariant(ByVal status As Variant) As String
    Select Case Nz(status, "")
        Case "Complete": StatusVariant = "success"
        Case "In progress": StatusVariant = "info"
        Case "Blocked": StatusVariant = "destructive"
        Case "Not started": StatusVariant = "warning"
        Case "Cancelled": StatusVariant = "secondary"
    End Select
End Function

' Band pill colors (BandCF). The demo shows no bands; these keep the kit's
' BandCF compiling.
Public Function BandLabels() As Variant
    BandLabels = Array("Above", "At", "Below")
End Function

Public Function BandBack(ByVal label As String) As Long
    Select Case label
        Case "Above": BandBack = CLR_BAND_ABOVE
        Case "At": BandBack = CLR_BAND_AT
        Case "Below": BandBack = CLR_BAND_BELOW
        Case Else: BandBack = CLR_BAND_NE
    End Select
End Function

' ---------------------------------------------------------------------------
' Shared helpers
' ---------------------------------------------------------------------------

' Message to the user. During the build it is printed instead; warnings raise.
Public Function Notify(ByVal msg As String, Optional ByVal style As VbMsgBoxStyle = vbInformation) As VbMsgBoxResult
    If g_Silent Then
        Debug.Print msg
        If (style And vbExclamation) = vbExclamation Or (style And vbCritical) = vbCritical Then
            Err.Raise vbObjectError + 999, APP_NAME, msg
        End If
        If (style And vbYesNo) = vbYesNo Then Notify = vbYes Else Notify = vbOK
    Else
        Notify = KitMsgBox(msg, style, APP_NAME)    ' centred over Access (modKitUI)
    End If
End Function

' Text as a quoted SQL string literal.
Public Function Q(ByVal s As Variant) As String
    Q = "'" & Replace(Nz(s, ""), "'", "''") & "'"
End Function

' Go to a page (deferred until the calling event finishes; frmMain.QueueNav).
Public Function NavTo(ByVal key As String)
    On Error GoTo Err_Handler
    If IsFormOpen("frmMain") Then Forms("frmMain").QueueNav key
    Exit Function
Err_Handler:
    HandleError "modDemo.NavTo"
End Function

' Items filter for the demo pages: a status chip ("" for all), search text
' and a category ("" for all).
Public Function ItemFilter(ByVal status As String, ByVal search As String, Optional ByVal category As String = "") As String
    Dim s As String
    If Len(status) > 0 Then s = "status=" & Q(status)
    If Len(category) > 0 Then s = s & IIf(Len(s) > 0, " AND ", "") & "category=" & Q(category)
    If Len(search) > 0 Then
        s = s & IIf(Len(s) > 0, " AND ", "") & "(name LIKE " & Q("*" & search & "*") & _
            " OR owner LIKE " & Q("*" & search & "*") & ")"
    End If
    ItemFilter = s
End Function

' Number of items ItemFilter would show.
Public Function ItemCount(ByVal status As String, ByVal search As String, Optional ByVal category As String = "") As Long
    Dim w As String
    w = ItemFilter(status, search, category)
    ItemCount = DCount("*", "demo_item", IIf(Len(w) > 0, w, "True"))
End Function

' Mark every item's due date on the date picker ("Due: <name>"), replacing
' the marks made before. Called when frmMain opens and after an item is saved.
' An error goes to HandleError (logged and shown) and leaves the marks made
' so far; the recordset closes when it goes out of scope.
Public Sub MarkDueDates()
    Dim db As DAO.Database, rs As DAO.Recordset
    On Error GoTo Err_Handler
    ClearMarkedDates
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT name, due_date FROM demo_item WHERE due_date IS NOT NULL", dbOpenSnapshot)
    Do Until rs.EOF
        MarkDate rs!due_date, "Due: " & Nz(rs!name, "(no name)")
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
Err_Handler:
    HandleError "modDemo.MarkDueDates"
End Sub

' Apply a filter to a table (subform control): "" shows every row.
Public Sub FilterGrid(sf As Control, ByVal where As String)
    sf.Form.Filter = where
    sf.Form.FilterOn = Len(where) > 0
    sf.Form.txtSel = Null
End Sub
