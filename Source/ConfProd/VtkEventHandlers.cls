VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VtkEventHandlers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private colEventHandlers As New Collection

Public Sub AddNew(OnAction As String, CmdBarCtl As CommandBarControl)
    Dim evh As New VtkEventHandler

    CmdBarCtl.OnAction = OnAction
    Set evh.cbe = Application.VBE.Events.CommandBarEvents(CmdBarCtl)
    colEventHandlers.Add evh
    Set evh = Nothing
End Sub

Public Sub Clear()
    Do Until colEventHandlers.Count = 0: colEventHandlers.Remove 1: Loop
End Sub
