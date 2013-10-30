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
Private m_projects As Collection

' collection of Strings indexed by project names
Private m_rootPathsCol As Collection
Private m_xmlRelPathsCol As Collection

Private m_xmlRememberedProjectsFullPath As String
Private Const m_xmlFileDefaultName As String = "VBAToolKitProjects.xml"

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
    
    ' Create and intialize the collections if they don't exist
    initFromList
    
    ' search for the vtkProject in the collection
    Dim project As vtkProject
    On Error Resume Next
    Set project = m_projects(projectName)
    If Err <> 0 Then
        Set project = New vtkProject
        project.projectName = projectName
        m_projects.Add Item:=project, Key:=projectName
    End If
    
    On Error GoTo 0
    
    ' return the vtkProject
    Set vtkProjectForName = project
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkRootPathForProject
' Author    : Lucas Vitorino
' Purpose   : Gives the rootpah of a project given its name. Returns "" if the project
'             doesn't have a known root path.
' Notes     : No project should have "" as its root path
'---------------------------------------------------------------------------------------
'
Public Function vtkRootPathForProject(projectName As String) As String
    ' Create and intialize the collections if they don't exist
    initFromList

    On Error Resume Next
    Dim tmpStr As String
    vtkRootPathForProject = m_rootPathsCol(projectName)
    If Err <> 0 Then
        vtkRootPathForProject = ""
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkXmlRelPathForProject
' Author    : Lucas Vitorino
' Purpose   : Gives the relative path of the xml vtkConfigurations sheet of a project relatively
'             too its root path.
' Notes     : No project should have "" as its xmlRelPath. The minimum value is the name of the xmm file.
'---------------------------------------------------------------------------------------
'
Public Function vtkXmlRelPathForProject(projectName As String) As String
    ' Create and intialize the collections if they don't exist
    initFromList
    
    On Error Resume Next
    
    vtkXmlRelPathForProject = m_xmlRelPathsCol(projectName)
    If Err <> 0 Then
        vtkXmlRelPathForProject = ""
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkResetProjects
' Author    : Jean-Pierre Imbert
' Date      : 04/06/2013
' Purpose   : Reset all vtkProjects (used during tests)
'---------------------------------------------------------------------------------------
'
Public Sub vtkResetProjects()
    Set m_projects = Nothing
    Set m_rootPathsCol = Nothing
    Set m_xmlRelPathsCol = Nothing
    m_xmlRememberedProjectsFullPath = ""
End Sub

'---------------------------------------------------------------------------------------
' Procedure : xmlRememberedProjectsFullPath
' Author    : Lucas Vitorino
' Purpose   : Returns the full path of the list of the projects remembered by VBAToolKit.
'             - If the folder in which the list should be located has not been set, it will
'               be located in the same folder as the current workbook.
'---------------------------------------------------------------------------------------
'
Private Property Get xmlRememberedProjectsFullPath() As String
    Dim fso As New FileSystemObject
    If m_xmlRememberedProjectsFullPath <> "" Then
        xmlRememberedProjectsFullPath = m_xmlRememberedProjectsFullPath
    Else
        xmlRememberedProjectsFullPath = fso.BuildPath(fso.GetParentFolderName(ThisWorkbook.FullName), m_xmlFileDefaultName)
    End If
End Property

'---------------------------------------------------------------------------------------
' Procedure : xmlRememberedProjectsFullPath
' Author    : Lucas Vitorino
' Purpose   : Set the full path of the list of the remembered projects.
'             This path is supposed to be absolute.
'---------------------------------------------------------------------------------------
'
Public Property Let xmlRememberedProjectsFullPath(ByVal filePath As String)
    m_xmlRememberedProjectsFullPath = filePath
End Property

'---------------------------------------------------------------------------------------
' Procedure : initFromList
' Author    : Lucas Vitorino
' Purpose   : Load projets from the xml list to the private Collections.
'---------------------------------------------------------------------------------------
'
Public Sub initFromList()
    
    On Error GoTo initFromList_Error
    
    ' If the collections are not yet initialized
    If (m_projects Is Nothing) Or (m_xmlRelPathsCol Is Nothing) Or (m_xmlRelPathsCol Is Nothing) Then
    
        ' Create the collections
        Set m_projects = New Collection
        Set m_rootPathsCol = New Collection
        Set m_xmlRelPathsCol = New Collection
        
        ' Check the existence of the file
        Dim fso As New FileSystemObject
        If fso.FileExists(xmlRememberedProjectsFullPath) = False Then
            Exit Sub
        End If

        ' Check the validity of the list format
        If isXmlSheetValid <> True Then
            Exit Sub
        End If
        
        ' Load the dom
        Dim dom As New MSXML2.DOMDocument
        dom.Load xmlRememberedProjectsFullPath
        
        ' Parse the dom according to the version
        If dom.getElementsByTagName("version").Item(0).Text = "1.0" Then
            ' Loop in the dom
            Dim projCount As Integer
            projCount = 0
            Dim tmpProj As vtkProject
            Dim tmpNode As MSXML2.IXMLDOMNode
            For Each tmpNode In dom.getElementsByTagName("project")
                Set tmpProj = New vtkProject
                tmpProj.projectName = dom.getElementsByTagName("name").Item(projCount).Text
        
                ' Add the relevant informations in the collections
                m_projects.Add Item:=tmpProj, Key:=tmpProj.projectName
                m_rootPathsCol.Add Item:=dom.getElementsByTagName("rootFolder").Item(projCount).Text, Key:=tmpProj.projectName
                m_xmlRelPathsCol.Add Item:=dom.getElementsByTagName("xmlRelativePath").Item(projCount).Text, Key:=tmpProj.projectName
                
                ' Prepare the next roll of the loop
                Set tmpProj = Nothing
                projCount = projCount + 1
            Next
            
        Else
            ' Version format is not supported
            Err.Raise VTK_WRONG_FORMAT
        End If

    End If

    On Error GoTo 0
    Exit Sub

initFromList_Error:
    Err.Source = "initFromList in vtkProjects"
    
    Select Case Err.Number
        Case VTK_WRONG_FORMAT
            Err.Description = "This version of the project list is not supported."
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Sub
End Sub


'---------------------------------------------------------------------------------------
' Procedure : isXmlSheetValid
' Author    : Lucas Vitorino
' Purpose   : Check the validity of the xml sheet containing all the projects.
'---------------------------------------------------------------------------------------
'
Private Function isXmlSheetValid() As Boolean
        
        Dim fso As New FileSystemObject
        Dim tmpNode As MSXML2.IXMLDOMNode
        
        On Error GoTo isXmlSheetValid_Error
        
        ' Make the verifications only if the file exists
        If fso.FileExists(xmlRememberedProjectsFullPath) Then
        
            ' Load the dom from the file
            Dim dom As New MSXML2.DOMDocument
            dom.Load xmlRememberedProjectsFullPath
            
            ' The root node must be called "rememberedProjects"
            Dim rootNode As MSXML2.IXMLDOMNode: Set rootNode = dom.ChildNodes.Item(1)
            If rootNode.BaseName <> "rememberedProjects" Then
                GoTo xmlSheetNotValid
            End If
            
            ' There must be only one "info" node with a only one "version" subnode
            If Not (dom.getElementsByTagName("version").Length = 1 And _
                    countElementsInNode("info", rootNode) = 1 And _
                    dom.getElementsByTagName("version").Length = 1 And _
                    countElementsInNode("version", getFirstChildNodeByName("info", rootNode)) = 1) Then
                GoTo xmlSheetNotValid
            End If
            
            ' All subnodes should have an accepted tag name
            If checkValidTags(rootNode) = False Then
                GoTo xmlSheetNotValid
            End If
            
            ' For now, the only accepted version of "version" is "1.0"
            If dom.getElementsByTagName("version").Item(0).Text = "1.0" Then
                ' Each "project" node must have 3 subnodes, in this order : name, folderPath, et xmlRelativePath
                ' These subnodes must not be empty (NB : content is not cheked)
                For Each tmpNode In dom.getElementsByTagName("project")
                    If (tmpNode.ChildNodes.Length <> 3 Or _
                       tmpNode.ChildNodes(0).BaseName <> "name" Or _
                       tmpNode.ChildNodes(0).Text = "" Or _
                       tmpNode.ChildNodes(1).BaseName <> "rootFolder" Or _
                       tmpNode.ChildNodes(1).Text = "" Or _
                       tmpNode.ChildNodes(2).BaseName <> "xmlRelativePath" Or _
                       tmpNode.ChildNodes(2).Text = "") _
                       Then
                        GoTo xmlSheetNotValid
                    End If
                Next
            Else
                GoTo xmlSheetNotValid
            End If
          
        Else
            ' If the file doesn't exist, return nothing
            isXmlSheetValid = False
        End If

    On Error GoTo 0
    isXmlSheetValid = True
    Exit Function

xmlSheetNotValid:
    isXmlSheetValid = False
    Exit Function

isXmlSheetValid_Error:
    Err.Source = "Function isXmlSheetValid in module vtkProjects"
    
    Select Case Err.Number
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
    
    Debug.Print "Error " & Err.Number & " : " & Err.Description & " in " & Err.Source
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Function
       
End Function


Public Function checkValidTags(node As MSXML2.IXMLDOMNode) As Boolean
    
    ' Check the name of the node
    If node.BaseName = "info" Or _
       node.BaseName = "version" Or _
       node.BaseName = "rememberedProjects" Or _
       node.BaseName = "project" Or _
       node.BaseName = "name" Or _
       node.BaseName = "rootFolder" Or _
       node.BaseName = "xmlRelativePath" Or _
       node.BaseName = "" _
        Then
       checkValidTags = True
    Else
        checkValidTags = False
    End If
    
    ' launch the sub for every child node
    Dim tmpNode As MSXML2.IXMLDOMNode
    If node.ChildNodes.Length <> 0 Then
        For Each tmpNode In node.ChildNodes
            If checkValidTags(tmpNode) = False Then checkValidTags = False
        Next
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : saveProjectsInList
' Author    : Lucas Vitorino
' Purpose   : Save projects in the private collection in the xml list.
'             Projects that are not added via the "add remembered projects" function are not saved.
' Notes     : - Overwrites the existing list.
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
    For Each tmpProj In m_projects
        ' Save only the projects that have an entry in the 3 collections.
        If vtkRootPathForProject(tmpProj.projectName) <> "" And vtkXmlRelPathForProject(tmpProj.projectName) <> "" Then
            vtkAddProjectToListOfRememberedProjects tmpPath, _
                                                    tmpProj.projectName, _
                                                    vtkRootPathForProject(tmpProj.projectName), _
                                                    vtkXmlRelPathForProject(tmpProj.projectName)
        End If
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

'---------------------------------------------------------------------------------------
' Procedure : vtkAddRememberedProject
' Author    : Lucas Vitorino
' Purpose   : Add a project to the collections and save it to the list.
' Notes     : This function does not filter the data, it will be done by the form that calls the function.
'---------------------------------------------------------------------------------------
'
Public Sub vtkAddRememberedProject(projectName As String, rootFolder As String, xmlRelativePath As String)

    On Error GoTo vtkAddRememberedProject_Error

    ' Load from the list if it has't been done yet
    initFromList
    
    ' Add the project to the collections
    vtkProjectForName projectName
    m_rootPathsCol.Add Item:=rootFolder, Key:=projectName
    m_xmlRelPathsCol.Add Item:=xmlRelativePath, Key:=projectName

    ' Save the list
    saveProjectsInList

    On Error GoTo 0
    Exit Sub

vtkAddRememberedProject_Error:
    Err.Source = "Sub vtkAddRememberedProject in module vtkProjects"
    
    Select Case Err.Number
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
    
    Err.Raise Err.Number, Err.Source, Err.Description

    Exit Sub

End Sub


'---------------------------------------------------------------------------------------
' Procedure : vtkRemoveRememberedProjects
' Author    : Lucas Vitorino
' Purpose   : Remove a project from the collections and from the list.
'---------------------------------------------------------------------------------------
'
Public Sub vtkRemoveRememberedProjects(projectName As String)

    On Error GoTo vtkRemoveRememberedProjects_Error

    ' Load the list if it hasn't been done yet
    initFromList
    
    ' Remove the projects from the collection
    m_rootPathsCol.Remove (projectName)
    m_xmlRelPathsCol.Remove (projectName)
    
    ' Save the list
    saveProjectsInList

    On Error GoTo 0
    Exit Sub

vtkRemoveRememberedProjects_Error:
    Err.Source = "Sub vtkAddRememberedProject in module vtkProjects"

    Select Case Err.Number
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
    
    Err.Raise Err.Number, Err.Source, Err.Description

    Exit Sub

End Sub
