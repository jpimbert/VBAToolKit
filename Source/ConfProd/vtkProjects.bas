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

' collection of Strings indexed by project names
Private rootPathsCol As Collection
Private xmlRelPathsCol As Collection

Private xmlRememberedProjectsFullPath As String
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
    
    ' Create and intialize the collections if they don't exist
    initFromList
    
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
' Procedure : vtkRootPathForProject
' Author    : Lucas Vitorino
' Purpose   : Gives the rootpah of a project given its name. Returns "" if the project
'             doesn't have a known root path.
'---------------------------------------------------------------------------------------
'
Public Function vtkRootPathForProject(projectName As String) As String
    
    ' Create and intialize the collections if they don't exist
    initFromList
    
    On Error Resume Next
    Dim tmpStr As String
    tmpStr = rootPathsCol(projectName)
    If Err <> 0 Then
        vtkRootPathForProject = ""
        Exit Function
    End If
    
    On Error GoTo 0
    vtkRootPathForProject = tmpStr
    Exit Function
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkXmlRelPathForProject
' Author    : Lucas Vitorino
' Purpose   : Gives the relative path of the xml vtkConfigurations sheet of a project relatively
'             too its root path. This path can be "".
'---------------------------------------------------------------------------------------
'
Public Function vtkXmlRelPathForProject(projectName As String) As String
    
    ' Create and intialize the collections if they don't exist
    initFromList
    
    On Error Resume Next
    Dim tmpStr As String
    tmpStr = xmlRelPathsCol(projectName)
    If Err <> 0 Then
        vtkXmlRelPathForProject = ""
        Exit Function
    End If
    
    On Error GoTo 0
    vtkXmlRelPathForProject = tmpStr
    Exit Function
    
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
    Set rootPathsCol = Nothing
    Set xmlRelPathsCol = Nothing
    xmlRememberedProjectsFullPath = ""
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getXmlRememberedProjectsFullPath
' Author    : Lucas Vitorino
' Purpose   : Returns the full path of the list of the projects remembered by VBAToolKit.
'             - If the folder in which the list should be located has not been set, it will
'               be located in the same folder as the current workbook.
'---------------------------------------------------------------------------------------
'
Private Function getXmlRememberedProjectsFullPath() As String
    Dim fso As New FileSystemObject
    If xmlRememberedProjectsFullPath <> "" Then
        getXmlRememberedProjectsFullPath = xmlRememberedProjectsFullPath
    Else
        getXmlRememberedProjectsFullPath = fso.BuildPath(fso.GetParentFolderName(ThisWorkbook.FullName), xmlFileDefaultName)
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkSetRememberedProjectsFullPath
' Author    : Lucas Vitorino
' Purpose   : Set the full path of the list of the remembered projects.
'             This path is supposed to be absolute.
'---------------------------------------------------------------------------------------
'
Public Sub vtkSetRememberedProjectsFullPath(ByVal filePath As String)
    xmlRememberedProjectsFullPath = filePath
End Sub

'---------------------------------------------------------------------------------------
' Procedure : initFromList
' Author    : Lucas Vitorino
' Purpose   : Load projets from the xml list in the private Collection.
'---------------------------------------------------------------------------------------
'
Public Sub initFromList()
    
    On Error GoTo initFromList_Error
    
    ' If the collections are not yet initialized
    If (projects Is Nothing) Or (xmlRelPathsCol Is Nothing) Or (xmlRelPathsCol Is Nothing) Then
    
        ' Create the collections
        Set projects = New Collection
        Set rootPathsCol = New Collection
        Set xmlRelPathsCol = New Collection
        
        ' Check the existence of the file
        ' If it doesn't exist, exit function without further processing
        Dim fso As New FileSystemObject
        If fso.FileExists(xmlRememberedProjectsFullPath) = False Then
            Exit Sub
        End If

        ' Load the dom
        Dim dom As New MSXML2.DOMDocument
        dom.Load xmlRememberedProjectsFullPath
        
        ' Loop in the dom
        Dim projCount As Integer
        projCount = 0
        Dim tmpProj As vtkProject
        Dim tmpNode As MSXML2.IXMLDOMNode
        For Each tmpNode In dom.getElementsByTagName("project")
            Set tmpProj = New vtkProject
            tmpProj.projectName = dom.getElementsByTagName("name").Item(projCount).Text
    
            ' Add the relevant informations in the collections
            projects.Add Item:=tmpProj, Key:=tmpProj.projectName
            rootPathsCol.Add Item:=dom.getElementsByTagName("rootFolder").Item(projCount).Text, Key:=tmpProj.projectName
            xmlRelPathsCol.Add Item:=dom.getElementsByTagName("xmlRelativePath").Item(projCount).Text, Key:=tmpProj.projectName
            
            ' Prepare the next roll of the loop
            Set tmpProj = Nothing
            projCount = projCount + 1
        Next

    End If

    On Error GoTo 0
    Exit Sub

initFromList_Error:
    Err.Source = "initFromList in vtkProjects"
    
    Select Case Err.Number
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Sub
End Sub


'---------------------------------------------------------------------------------------
' Procedure : isXmlSheetValid
' Author    : Lucas Vitorino
' Purpose   : Check the validity of xml sheet
'---------------------------------------------------------------------------------------
'
Private Function isXmlSheetValid() As Boolean
        
        Dim fso As New FileSystemObject
        Dim tmpNode As MSXML2.IXMLDOMNode
        
        On Error GoTo isXmlSheetValid_Error
        
        ' Make the verifications only if the file exists
        If fso.FileExists(xmlRememberedProjectsPath) Then
        
            ' Load the dom from the file
            Dim dom As New MSXML2.DOMDocument
            dom.Load xmlRememberedProjectsFullPath
            
            ' The root node must be called "rememberedProjects"
            Dim rootNode As MSXML2.IXMLDOMNode: rootNode = dom.ChildNodes.Item(1)
            If rootNode.BaseName <> "rememberedProjects" Then
                GoTo xmlSheetNotValid
            End If
            
            ' There must be one "info" node with a "version" subnode
            ' For now, the only accepted version of "version" is "1.0"
            
            
            
            ' All subnodes should have an accepted tag name
            If checkValidTags(dom.ChildNodes.Item(1)) = False Then
                GoTo xmlSheetNotValid
            End If
            
            ' Each "project" node must have 3 subnodes, in this order : name, folderPath, et xmlRelativePath
            ' These subnodes must not be empty
            For Each tmpNode In dom.getElementsByTagName("project")
                If tmpNode.ChildNodes.Length <> 3 Or _
                   tmpNode.ChildNodes(0).BaseName <> "name" Or _
                   tmpNode.ChildNodes(0).Text = "" Or _
                   tmpNode.ChildNodes(1).BaseName <> "folderPath" Or _
                   tmpNode.ChildNodes(1).Text = "" Or _
                   tmpNode.ChildNodes(2).BaseName <> "xmlRelativePath" Or _
                   tmpNode.ChildNodes(2).Text = "" _
                   Then
                    GoTo xmlSheetNotValid
                End If
            Next
        
        
        End If

    On Error GoTo 0
    Exit Function

xmlSheetNotValid:
    isXmlSheetValid = False
    Exit Function

isXmlSheetValid_Error:
    Err.Source = "Function isXmlSheetValid in module vtkProjects"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Function
       
End Function


Public Function checkValidTags(node As MSXML2.IXMLDOMNode) As Boolean
    
    ' Check the name of the node
    If node.BaseName = "version" Or _
       node.BaseName = "rememberedProjects" Or _
       node.BaseName = "name" Or _
       node.BaseName = "rootFolder" Or _
       node.BaseName = "xmlRelativePath" _
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
' Notes     : Overwrites the list.
'---------------------------------------------------------------------------------------
'
'Private Sub saveProjectsInList()
'
'    On Error GoTo saveProjectsInList_Error
'
'    Dim fso As New FileSystemObject
'
'    ' Save in a tmp file with a random name, if all goes well, we remove the old file
'    ' and rename the new.
'    Dim tmpPath As String
'    tmpPath = fso.BuildPath(fso.GetParentFolderName(xmlRememberedProjectsFullPath), _
'              vtkStripFilePathOrNameOfExtension(xmlRememberedProjectsFullPath) & _
'              CStr(Round((99999 - 10000 + 1) * Rnd(), 0)) + 10000 & _
'              "." & fso.GetExtensionName(xmlRememberedProjectsFullPath))
'
'    vtkCreateListOfRememberedProjects tmpPath
'
'    ' Loop in the collection
'    Dim tmpProj As vtkProject
'    For Each tmpProj In projects
'        vtkAddProjectToListOfRememberedProjects tmpPath, _
'                                                tmpProj.projectName, _
'                                                tmpProj.projectRootFolderPath, _
'                                                tmpProj.xmlRelativeFolderPath
'    Next
'
'    ' All went well, remove the old file, rename the new one
'    Kill xmlRememberedProjectsFullPath
'    fso.GetFile(tmpPath).name = fso.GetFileName(xmlRememberedProjectsFullPath)
'
'    On Error GoTo 0
'    Exit Sub
'
'saveProjectsInList_Error:
'    Err.Source = "saveProjectsInList in vtkProjects"
'
'    Select Case Err.Number
'        Case 53
'            ' Error raised by Kill because there was no file
'            Resume Next
'        Case Else
'            Err.Number = VTK_UNEXPECTED_ERROR
'    End Select
'
'    Err.Raise Err.Number, Err.Source, Err.Description
'
'    Exit Sub
'End Sub
