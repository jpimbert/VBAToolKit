VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VtkEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents cbe As CommandBarEvents
Attribute cbe.VB_VarHelpID = -1

Private Sub cbe_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Application.Run CommandBarControl.OnAction
    Handled = True
    CancelDefault = True
End Sub
