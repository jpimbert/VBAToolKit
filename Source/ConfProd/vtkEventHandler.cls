VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class     : vtkEventHandler
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Simple object used to route a click event on a commandbar control
'             - the use of event object is the only way to proceed VBE commandbar events
'             - one instance of this object must be created for each control
'             - these instances are managed with the vtkEventHandlers module
'---------------------------------------------------------------------------------------
Option Explicit
Public WithEvents cbe As CommandBarEvents
Attribute cbe.VB_VarHelpID = -1

Private Sub cbe_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    '   Run the action registered on the command Bar control
    Application.Run CommandBarControl.OnAction
    '   Tell VBA that the Event was proceeded
    Handled = True
    CancelDefault = True
End Sub

