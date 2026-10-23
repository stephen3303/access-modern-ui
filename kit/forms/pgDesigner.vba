Option Compare Database
Option Explicit

' Form designer page (BuildDesigner, modKitDesigner): every event goes to the
' Designer* routines there.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    DesignerLoad Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Handler
    KeyCode = DesignerKeyDown(Me, KeyCode, Shift)
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.Form_KeyDown"
End Sub

' The shell asks before showing another page: unsaved changes are saved or dropped first.
Public Function CanLeave() As Boolean
    On Error GoTo Err_Handler
    CanLeave = DesignerCanLeave(Me)
    Exit Function
Err_Handler:
    HandleError "pgDesigner.CanLeave"
End Function

' The canvas.
Private Sub dzGlass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Handler
    DesignerMouse Me, 1, Button, X, Y
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.dzGlass_MouseDown"
End Sub

Private Sub dzGlass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Handler
    DesignerMouse Me, 2, Button, X, Y
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.dzGlass_MouseMove"
End Sub

Private Sub dzGlass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Handler
    DesignerMouse Me, 3, Button, X, Y
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.dzGlass_MouseUp"
End Sub

' The palette tiles: Card, Button, Label, Progress.
Private Sub dzTool0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToolMouse 0, 1, Button, X, Y
End Sub

Private Sub dzTool0_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToolMouse 0, 2, Button, X, Y
End Sub

Private Sub dzTool0_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToolMouse 0, 3, Button, X, Y
End Sub

Private Sub dzTool1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToolMouse 1, 1, Button, X, Y
End Sub

Private Sub dzTool1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToolMouse 1, 2, Button, X, Y
End Sub

Private Sub dzTool1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToolMouse 1, 3, Button, X, Y
End Sub

Private Sub dzTool2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToolMouse 2, 1, Button, X, Y
End Sub

Private Sub dzTool2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToolMouse 2, 2, Button, X, Y
End Sub

Private Sub dzTool2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToolMouse 2, 3, Button, X, Y
End Sub

Private Sub dzTool3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToolMouse 3, 1, Button, X, Y
End Sub

Private Sub dzTool3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToolMouse 3, 2, Button, X, Y
End Sub

Private Sub dzTool3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToolMouse 3, 3, Button, X, Y
End Sub

Private Sub ToolMouse(ByVal idx As Integer, ByVal phase As Integer, ByVal btn As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo Err_Handler
    DesignerToolMouse Me, idx, phase, btn, X, Y
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.dzTool" & idx & "_Mouse"
End Sub

' The selected element and the layout.
Private Sub txtDzText_AfterUpdate()
    On Error GoTo Err_Handler
    DesignerText Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.txtDzText_AfterUpdate"
End Sub

Private Sub btnDzDelete_Click()
    On Error GoTo Err_Handler
    DesignerDelete Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.btnDzDelete_Click"
End Sub

Private Sub cboDzOpen_AfterUpdate()
    On Error GoTo Err_Handler
    DesignerOpen Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.cboDzOpen_AfterUpdate"
End Sub

Private Sub txtDzName_AfterUpdate()
    On Error GoTo Err_Handler
    DesignerRenamed
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.txtDzName_AfterUpdate"
End Sub

Private Sub btnDzNew_Click()
    On Error GoTo Err_Handler
    DesignerNew Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.btnDzNew_Click"
End Sub

Private Sub btnDzSave_Click()
    On Error GoTo Err_Handler
    DesignerSave Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.btnDzSave_Click"
End Sub

Private Sub btnDzGenerate_Click()
    On Error GoTo Err_Handler
    DesignerGenerate Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.btnDzGenerate_Click"
End Sub
