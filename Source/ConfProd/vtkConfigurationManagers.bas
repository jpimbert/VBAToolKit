Attribute VB_Name = "vtkConfigurationManagers"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkConfigurationManagers
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Manage the configuration managers (class vtkConfigurationManager) for open projects
'
' Usage:
'   - Each instance of Configuration Manager is attached to a VBA project (supposed to be a VTK project)
'       - the method configurationManagerForProject give the instance attached to a project, or create it
'
'---------------------------------------------------------------------------------------

'   collection of instances indexed by project names
Private configurationManagers As Collection

Public Function configurationManagerForProject(projectName As String) As vtkConfigurationManager

End Function

