Attribute VB_Name = "vtkApplicationEventHandlerHook"
Option Explicit

Dim applicationEventHandler As New vtkApplicationEventHandler
 
Sub vtkInitializeApplicationEventHandler()
    Set applicationEventHandler.App = Application
End Sub

