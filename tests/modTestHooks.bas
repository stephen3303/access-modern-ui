Option Compare Database
Option Explicit
Public Const APP_NAME As String = "AccessModernUI-Tests"
Public g_Silent As Boolean
' Test-only stand-in for the host application's GDI text measurement.
Public Function MeasureText(ByVal text As String, ByVal size As Long, ByVal weight As Long) As Long
    MeasureText = Len(text) * size * 12
End Function
Public Function ShowTip(ByVal frm As Object, ByVal name As String)
End Function
Public Function BandLabels() As Variant
    BandLabels = Array("Ready")
End Function
Public Function BandBack(ByVal label As String) As Long
    BandBack = CLR_PRIMARY
End Function
Public Function BadgeColors(ByVal variant As String) As Variant
    BadgeColors = Array(CLR_PRIMARY, CLR_WHITE)
End Function
