Attribute VB_Name = "vtkConfigurationManagers"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkConfigurationManagers
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Manage the configuration managers (class vtkConfigurationManager) for open projects
'
' Usage:
'   - Each instance of Configuration Manager is attached to the DEV Excel Workbook of a project
'       - the method vtkConfigurationManagerForProject give the instance attached to a workbook, or create it
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
Private m_configurationManagers As Collection

'---------------------------------------------------------------------------------------
' Procedure : vtkConfigurationManagerForProject
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Return the configuration manager attached to the DEV Excel file given its project name
'               - if the configuration doesn't exist, it is created
'               - if the m_configurationManagers collection doesn't exist, it is created
'---------------------------------------------------------------------------------------
'
Public Function vtkConfigurationManagerForProject(projectName As String) As vtkConfigurationManager
    ' Create the collection if it doesn't exist
    If m_configurationManagers Is Nothing Then
        Set m_configurationManagers = New Collection
        End If
    ' search for the configuration manager in the collection
    Dim cm As vtkConfigurationManager
    On Error Resume Next
    Set cm = m_configurationManagers(projectName)
    If Err <> 0 Then
        Set cm = New vtkConfigurationManager
        cm.projectName = projectName
        If cm.projectName Like projectName Then     ' The initialization could fail (if the Workbook is closed)
            m_configurationManagers.Add Item:=cm, Key:=projectName
           Else
            Set cm = Nothing
        End If
    End If
   On Error GoTo 0
    ' return the configuration manager
    Set vtkConfigurationManagerForProject = cm
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkResetConfigurationManagers
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Reset all configuration managers (used during tests)
'---------------------------------------------------------------------------------------
'
Public Sub vtkResetConfigurationManagers()
    Set m_configurationManagers = Nothing
End Sub

