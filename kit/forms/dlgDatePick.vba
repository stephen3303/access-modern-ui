Option Compare Database
Option Explicit

' Date picker for a kit DateBox (PickDate opens it as a dialog; BuildDatePicker
' builds it). It shows one month as day cells d0..d41, Sunday first, with the
' box's date highlighted and today in bold. A day click, Today or Clear writes
' into the box (SetDateTarget) and closes; Cancel or Esc closes unchanged.

Private m_month As Date     ' first day of the month shown

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Dim v As Variant
    v = DateTargetValue()
    If Not IsDate(v) Then v = Date
    m_month = DateSerial(Year(v), Month(v), 1)
    ShowMonth
    Exit Sub
Err_Handler:
    HandleError "dlgDatePick.Form_Load"
End Sub

' Caption and paint the 42 day cells for m_month. Each cell's Tag holds its
' date as a serial number.
Private Sub ShowMonth()
    Dim first As Date, d As Date, i As Long, c As Control, v As Variant, sel As Date, hasSel As Boolean
    v = DateTargetValue()
    hasSel = IsDate(v)
    If hasSel Then sel = DateValue(v)
    Me.lblMonth.Caption = Format$(m_month, "mmmm yyyy")
    first = m_month - (Weekday(m_month, vbSunday) - 1)
    For i = 0 To 41
        d = first + i
        Set c = Me("d" & i)
        c.Caption = CStr(Day(d))
        c.Tag = CStr(CLng(d))
        c.ForeColor = Th(IIf(Month(d) = Month(m_month), CLR_FG, CLR_MUTED_FG))
        c.FontWeight = IIf(d = Date, 700, 400)
        If hasSel And d = sel Then
            c.BackStyle = 1
            c.BackColor = Th(CLR_PRIMARY_SOFT)
        Else
            c.BackStyle = 0
        End If
    Next
End Sub

' Previous (-1) or next (1) month.
Public Function MoveMonth(ByVal n As Long)
    On Error GoTo Err_Handler
    m_month = DateAdd("m", n, m_month)
    ShowMonth
    Exit Function
Err_Handler:
    HandleError "dlgDatePick.MoveMonth"
End Function

' Day cell i clicked: that date.
Public Function PickDay(ByVal i As Long)
    On Error GoTo Err_Handler
    Finish CDate(CLng(Me("d" & i).Tag))
    Exit Function
Err_Handler:
    HandleError "dlgDatePick.PickDay"
End Function

Public Function PickToday()
    On Error GoTo Err_Handler
    Finish Date
    Exit Function
Err_Handler:
    HandleError "dlgDatePick.PickToday"
End Function

Public Function PickClear()
    On Error GoTo Err_Handler
    Finish Null
    Exit Function
Err_Handler:
    HandleError "dlgDatePick.PickClear"
End Function

' Write v into the box and close. Errors go to the caller's handler.
Private Sub Finish(ByVal v As Variant)
    SetDateTarget v
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgDatePick.btnCancel_Click"
End Sub
