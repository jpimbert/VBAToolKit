Attribute VB_Name = "vtkProjects"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkProjects
' Author    : Jean-Pierre Imbert
' Date      : 04/06/2013
' Purpose   : Manage the vtkProject objects (class vtkProject) for open projects
'
' Usage:
'   - Each instance of vtkProject is attached to a VBA Tool Kit Project
'       - the method vtkProjectForName give the instance attached to a VTK project, or create it
'
'
' Copyright 2013 Skwal-Soft (http://skwalsoft.com)
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.
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
Public Sub vtkResetProjects()
    Set projects = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : getCurrentProjectName
' Author    : Lucas Vitorino
' Purpose   : Get the project name associated with the active DEV workbook.
' Notes     : Temporary and not tested
'---------------------------------------------------------------------------------------
'
Public Function getCurrentProjectName() As String
    getCurrentProjectName = vtkStripPathOrNameOfVtkExtension(ActiveWorkbook.FullName, "DEV")
End Function
