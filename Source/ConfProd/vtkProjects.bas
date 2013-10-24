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

Private xmlRememberedProjectsFolder As String
Private Const xmlFileDefaultName As String = "VBAToolKitProjects.xml"

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
' Procedure : vtkXmlConfigurationsFullPath
' Author    : Lucas Vitorino
' Purpose   : Returns the full path of the list of the projects remembered by VBAToolKit.
'             - If the folder in which the list should be located has not been set, it will
'               be located in the same folder as the current workbook.
'---------------------------------------------------------------------------------------
'
Private Function xmlRememberedProjectsFullPath() As String
    Dim fso As New FileSystemObject
    If xmlRememberedProjectsFolder <> "" Then
        xmlRememberedProjectsFullPath = fso.BuildPath(xmlRememberedProjectsFolder, xmlFileDefaultName)
    Else
        xmlRememberedProjectsFullPath = fso.BuildPath(fso.GetParentFolderName(ThisWorkbook.FullName), xmlFileDefaultName)
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkSetRememberedProjectsFolder
' Author    : Lucas Vitorino
' Purpose   : Set the folder in which the list of the remembered projects will be located.
'             This path is supposed to be absolute.
'---------------------------------------------------------------------------------------
'
Public Sub vtkSetRememberedProjectsFolder(ByVal folderPath As String)
    xmlRememberedProjectsFolder = folderPath
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkResetRememberedProjectsFolder
' Author    : Lucas Vitorino
' Purpose   : Reset the private variable xmlRememberedProjectsFolder
'---------------------------------------------------------------------------------------
'
Public Sub vtkResetRememberedProjectsFolder()
    xmlRememberedProjectsFolder = ""
End Sub

'---------------------------------------------------------------------------------------
' Procedure : loadProjectsFromList
' Author    : Lucas Vitorino
' Purpose   : Load projets from the xml list in the private Collection.
'---------------------------------------------------------------------------------------
'
Private Sub loadProjectsFromList()
    
    On Error GoTo loadProjectsFromList_Error
    
    ' Check the existence of the file
    ' If it doesn't exist, raise an error
    Dim fso As New FileSystemObject
    If fso.FileExists(xmlRememberedProjectsFullPath) = False Then
        Err.Raise VTK_NO_PROJECT_LIST
    End If
    
    ' Load the dom
    Dim dom As MSXML2.DOMDocument
    dom.Load xmlRememberedProjetsFullPath
    
    ' Loop in the dom
    Dim tmpProj As vtkProject
    Dim tmpNode As MSXML2.IXMLDOMNode
    For Each tmpNode In dom.getElementsByTagName("project")
        Set tmpProj = New vtkProject
        tmpProj.projectName = tmpNode.ChildNodes.Item(0).Text
        tmpProj.projectBeforeRootFolderPath = tmpNode.ChildNodes.Item(1).Text
        tmpProj.xmlRelativeFolderPath = tmpNode.ChildNodes.Item(2).Text
        
        projects.Add Item:=tmpProj, Key:=tmpProj.projectName
        
        Set tmpProj = Nothing
    Next

    On Error GoTo 0
    Exit Sub

loadProjectsFromList_Error:
    Err.Source = "loadProjectsFromList in vtkProjects"
    
    Select Case Err.Number
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Sub
End Sub


'---------------------------------------------------------------------------------------
' Procedure : saveProjectsInList
' Author    : Lucas Vitorino
' Purpose   : Save projects in the private collection in the xml list.
' Notes     : Overwrites the list.
'---------------------------------------------------------------------------------------
'
Private Sub saveProjectsInList()
    
    On Error GoTo saveProjectsInList_Error

    Dim fso As New FileSystemObject

    ' Save in a tmp file with a random name, if all goes well, we remove the old file
    ' and rename the new.
    Dim tmpPath As String
    tmpPath = fso.BuildPath(fso.GetParentFolderName(xmlRememberedProjectsFullPath), _
              vtkStripFilePathOrNameOfExtension(xmlRememberedProjectsFullPath) & _
              CStr(Round((99999 - 10000 + 1) * Rnd(), 0)) + 10000 & _
              "." & fso.GetExtensionName(xmlRememberedProjectsFullPath))
    
    vtkCreateListOfRememberedProjects tmpPath
    
    ' Loop in the collection
    Dim tmpProj As vtkProject
    For Each tmpProj In projects
        vtkAddProjectToListOfRememberedProjects tmpPath, _
                                                tmpProj.projectName, _
                                                tmpProj.projectRootFolderPath, _
                                                tmpProj.xmlRelativeFolderPath
    Next

    ' All went well, remove the old file, rename the new one
    Kill xmlRememberedProjectsFullPath
    fso.GetFile(tmpPath).name = fso.GetFileName(xmlRememberedProjectsFullPath)

    On Error GoTo 0
    Exit Sub

saveProjectsInList_Error:
    Err.Source = "saveProjectsInList in vtkProjects"
    
    Select Case Err.Number
        Case 53
            ' Error raised by Kill because there was no file
            Resume Next
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Sub
End Sub
