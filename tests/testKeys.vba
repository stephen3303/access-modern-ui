Option Compare Database
Option Explicit
' A form that does the "new" shortcut itself (modKitTests: KitShortcut).
Public lastAction As String
Public Function Shortcut(ByVal action As String) As Boolean
    lastAction = action
    Shortcut = (action = "new")
End Function
