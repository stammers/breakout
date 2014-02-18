Attribute VB_Name = "Module1"
Public score As Integer
Public lives As Integer
Public Sub reset()
Unload BreakoutFrm
Load BreakoutFrm
BreakoutFrm.Visible = True
End Sub
