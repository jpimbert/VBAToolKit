Attribute VB_Name = "vtkEventHandlers"
'---------------------------------------------------------------------------------------
' Module    : vtkEventHandlers
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Manage the event handlers used
'---------------------------------------------------------------------------------------
Option Explicit
Private eventHandlers As Collection

'---------------------------------------------------------------------------------------
' Procedure : vtkAddEventHandler
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Create a new managed eventHandler for the commandbar control given as parameter
'             - the new event handler is included in the event handlers list
'---------------------------------------------------------------------------------------
'
Public Sub vtkAddEventHandler(CmdBarCtl As CommandBarControl)
    Dim evh As New vtkEventHandler

    Set evh.cbe = Application.VBE.Events.CommandBarEvents(CmdBarCtl)
    If eventHandlers Is Nothing Then Set eventHandlers = New Collection
    eventHandlers.Add evh
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkDeleteEventHandlers
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Delete all the event handlers and the list
'---------------------------------------------------------------------------------------
'
Public Sub vtkDeleteEventHandlers()
    If Not eventHandlers Is Nothing Then
        Do Until eventHandlers.Count = 0: colEventHandlers.Remove 1: Loop
        eventHandlers = Nothing
    End If
End Sub

