Attribute VB_Name = "vtkProjects"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkProjects
' Author    : Jean-Pierre Imbert
' Date      : 04/06/2013
' Purpose   : Manage the vtkProject objects (class vtkProject) for open projects
'
' Usage:
'   - Each instance of vtkProject is attached to an VB Tool Kit Project
'       - the method vtkProjectForName give the instance attached to VTK project, or create it
'
'---------------------------------------------------------------------------------------

'   collection of instances indexed by project names
Private projects As Collection

'---------------------------------------------------------------------------------------
' Procedure : vtkProjectForName
' Author    : Jean-Pierre Imbert
' Date      : 04/06/2013
' Purpose   : Return the vtkProject given its name
'               - if the object doesn't exist, it is created
'               - if the objects collection doesn't exist, it is created
'---------------------------------------------------------------------------------------
'
Public Function vtkProjectForName(projectName As String) As vtkProject
    ' Create the collection if it doesn't exist
    If projects Is Nothing Then
        Set projects = New Collection
        End If
    ' search for the configuration manager in the collection
    Dim cm As vtkProject
    On Error Resume Next
    Set cm = projects(projectName)
    If Err <> 0 Then
        Set cm = New vtkProject
        cm.projectName = projectName
        projects.Add Item:=cm, Key:=projectName
        End If
    On Error GoTo 0
    ' return the configuration manager
    Set vtkProjectForName = cm
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkResetProjects
' Author    : Jean-Pierre Imbert
' Date      : 04/06/2013
' Purpose   : Reset all vtkProjects (used during tests)
'---------------------------------------------------------------------------------------
'
Public Sub vtkResetConfigurationManagers()
    Set projects = Nothing
End Sub

