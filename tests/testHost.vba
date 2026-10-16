Option Compare Database
Option Explicit
Public callbackGrid As String
Public toggleChanges As Long
Public segmentChanges As Long
Public lastCrumb As String
' A Toggle's onChange (modKitTests).
Public Sub ToggleChanged()
    toggleChanges = toggleChanges + 1
End Sub
' A Segmented control's onChange (modKitTests).
Public Sub SegmentChanged()
    segmentChanges = segmentChanges + 1
End Sub
' A Breadcrumb's handler (modKitTests).
Public Sub CrumbPicked(ByVal key As String)
    lastCrumb = key
End Sub
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    callbackGrid = grid
End Sub
