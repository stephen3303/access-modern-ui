Option Compare Database
Option Explicit
Public callbackGrid As String
Public toggleChanges As Long
Public segmentChanges As Long
Public lastCrumb As String
Public lastStep As Long
Public lastSearch As String
Public ratingChanges As Long
Public sliderChanges As Long
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
' A Stepper's handler (modKitTests).
Public Sub StepPicked(ByVal i As Long)
    lastStep = i
End Sub
' A SearchBox's handler (modKitTests).
Public Sub SearchTyped(ByVal text As String)
    lastSearch = text
End Sub
' A Rating's and a Slider's onChange (modKitTests).
Public Sub RatingChanged()
    ratingChanges = ratingChanges + 1
End Sub
Public Sub SliderChanged()
    sliderChanges = sliderChanges + 1
End Sub
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    callbackGrid = grid
End Sub
