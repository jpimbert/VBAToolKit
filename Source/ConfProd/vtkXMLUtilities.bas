Attribute VB_Name = "vtkXMLUtilities"
'---------------------------------------------------------------------------------------
' Module    : vtkXMLUtilities
' Author    : Lucas Vitorino
' Purpose   : Provide utilities to support XML.
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


'---------------------------------------------------------------------------------------
' Procedure : vtkExportAsXMLDOM
' Author    : Lucas Vitorino
' Purpose   : Export a DEV workbook as an XML DOM.
'
' Returns   : MSXML2.DOMDocument
' Raises    : - VTK_WORKBOOK_NOT_OPEN
'             - VTK_WORKBOOK_NOT_INITIALIZED
'             - VTK_PROJECT_NOT_INITIALIZED
'             - VTK_UNEXPECTED_ERROR
'---------------------------------------------------------------------------------------
'
Public Function vtkExportAsXMLDOM(projectName As String) As MSXML2.DOMDocument
    Dim dom As MSXML2.DOMDocument
    Dim node As MSXML2.IXMLDOMNode
    Dim rootNode As MSXML2.IXMLDOMElement
    Dim tmpEl As MSXML2.IXMLDOMElement
    Dim attr As MSXML2.IXMLDOMAttribute
    
    On Error GoTo vtkExportAsXMLDOM_Error
    
    ' If the project is not initialized
    Dim cm As vtkConfigurationManager
    Dim conf As vtkConfiguration
    Set cm = vtkConfigurationManagerForProject(projectName)
    If cm Is Nothing Then
        Err.Raise VTK_PROJECT_NOT_INITIALIZED
    End If
    
    Set dom = New MSXML2.DOMDocument
    Set node = dom.createProcessingInstruction("xml", "version=""1.0"" encoding=""ISO-8859-1""")
    dom.appendChild node

    ' The root node is an DOMElement, not a DOMNode
    Set rootNode = dom.createElement("vtkConf")
    
    With dom.appendChild(rootNode)
        
        ' The info element
        With .appendChild(dom.createElement("info"))
        
            ' Project name
            With .appendChild(dom.createElement("projectName"))
                .Text = projectName
            End With
            
            ' Version of vtkConfigurations
            With .appendChild(dom.createElement("vtkConfigurationsVersion"))
                .Text = "1.0"
            End With
            
        End With
        
        
        'The configuration elements
        For Each conf In cm.configurations
            With .appendChild(dom.createElement("configuration"))
                
                ' The name
                With .appendChild(dom.createElement("name"))
                    .Text = conf.name
                End With
                
                'The path
                With .appendChild(dom.createElement("path"))
                    .Text = conf.path
                End With
                
            End With
        Next
        
        
        ' The module elements
        Dim mo As vtkModule
        For Each mo In cm.modules
            With .appendChild(dom.createElement("module"))
                
                ' The name
                With .appendChild(dom.createElement("name"))
                    .Text = mo.name
                End With
                
                ' The path for each configuration
                For Each conf In cm.configurations
                    Set attr = dom.createAttribute("confName")
                    attr.NodeValue = conf.name
                    Set tmpEl = dom.createElement("path")
                    tmpEl.setAttributeNode attr
                    With .appendChild(tmpEl)
                        .Text = mo.getPathForConfiguration(conf.name)
                    End With
                Next
            
            End With
        Next
        
    End With

    On Error GoTo 0
    Set vtkExportAsXMLDOM = dom
    Exit Function

vtkExportAsXMLDOM_Error:
    Err.Source = "function vtkExportAsDOMXML of module vtkXMLutilities"
    
    Select Case Err.Number
        Case VTK_WORKBOOK_NOT_OPEN
            Err.Description = "Workbook should be open."
        Case VTK_WORKBOOK_NOT_INITIALIZED
            Err.Description = "Workbook not initialized."
        Case -2147221080 ' Automation error undocumented by Microsoft
            Err.Number = VTK_WORKBOOK_NOT_OPEN
            Err.Description = "Workbook should be open."
        Case VTK_PROJECT_NOT_INITIALIZED
            Err.Description = "Project " & projectName & " has not been initialized. The name might be wrong."
        Case Else
            Debug.Print "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select

    Err.Raise Err.Number

    Exit Function
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkWriteXMLDOMToFile
' Author    : Lucas Vitorino
' Purpose   : - Write an XML DOM to a xml text file.
'             - The content of the file is nicely indented, to be human-readable.
'             - Overwrite the output file if it exists.
' Raises    : - VTK_DOM_NOT_INITIALIZED
'             - VTK_UNEXPECTED_ERRROR
' Notes     : Heavily based on code from Baptiste Wicht, http://baptiste-wicht.developpez.com/
'---------------------------------------------------------------------------------------
'
Public Sub vtkWriteXMLDOMToFile(dom As MSXML2.DOMDocument, filePath As String)

    Dim rdr As MSXML2.SAXXMLReader
    Dim wrt As MSXML2.MXXMLWriter
    
    On Error GoTo vtkWriteXMLDOMToFile_Error
    
    ' Check DOM intialization
    If dom Is Nothing Then
        Err.Raise VTK_DOM_NOT_INITIALIZED
    End If
    
    Set rdr = CreateObject("MSXML2.SAXXMLReader")
    Set wrt = CreateObject("MSXML2.MXXMLWriter")
    
    Dim oStream As ADODB.STREAM
    Set oStream = CreateObject("ADODB.STREAM")
    oStream.Open
    oStream.Charset = "ISO-8859-1"

    wrt.indent = True
    wrt.Encoding = "ISO-8859-1"
    wrt.output = oStream
    Set rdr.contentHandler = wrt
    Set rdr.errorHandler = wrt

    rdr.Parse dom
    wrt.flush

    oStream.SaveToFile filePath, adSaveCreateOverWrite

    On Error GoTo 0
    Exit Sub

vtkWriteXMLDOMToFile_Error:
    Err.Source = "function vtkWriteXMLDOMToFile of module vtkXMLutilities"
    
    Select Case Err.Number
        Case VTK_DOM_NOT_INITIALIZED
            Err.Description = "Dom object is not initialized."
        Case 3004 ' ADODB.Stream.SaveToFile failed because it couldn't find the path
            Err.Number = VTK_WRONG_FILE_PATH
            Err.Description = "File path is wrong. Make sure the folder tree is valid."
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select

    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Sub

End Sub


'---------------------------------------------------------------------------------------
' Procedure : vtkCreateListOfRememberedProjects
' Author    : Lucas Vitorino
' Purpose   : Create the xml file containing the list of the remembered projects.
' Raises    : - VTK_UNEXPECTED_ERROR
'             - VTK_DOM_NOT_INITIALIZED
'             - VTK_WRONG_FILE_PATH
'---------------------------------------------------------------------------------------
'
Public Sub vtkCreateListOfRememberedProjects(filePath As String)
    
    Dim dom As MSXML2.DOMDocument
    Dim rootNode As MSXML2.IXMLDOMNode
    
    On Error GoTo vtkCreateListOfRememberedProjects_Error

    ' Create the processing instruction
    Set dom = New MSXML2.DOMDocument
    dom.appendChild dom.createProcessingInstruction("xml", "version=""1.0"" encoding=""ISO-8859-1""")

    ' Create the root node
    dom.appendChild dom.createElement("rememberedProjects")

    vtkWriteXMLDOMToFile dom, filePath

    On Error GoTo 0
    Exit Sub

vtkCreateListOfRememberedProjects_Error:
    Err.Source = "vtkCreateXMLListOfRememberedProjects of module vtkXMLUtilities"
    
    Select Case Err.Number
        Case VTK_WRONG_FILE_PATH ' ADODB.Stream.SaveToFile failed because it couldn't find the path
            ' Do nothing but forward the error
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select

    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Sub
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : vtkAddProjectToListOfRememberedProjects
' Author    : Lucas Vitorino
' Purpose   : Add a project to a list of remembered projects
' Raises    : - VTK_PROJECT_ALREADY_IN_LIST
'             - VTK_WRONG_FILE_PATH
'             - VTK_UNEXEPECTED_ERROR
'---------------------------------------------------------------------------------------
'
Public Sub vtkAddProjectToListOfRememberedProjects(listPath As String, _
                                                   projectName As String, _
                                                   projectBeforeRootFolder As String, _
                                                   projectXMLRelativePath As String)
                                                                                         
    On Error GoTo vtkAddProjectToListOfRememberedProjects_Error

    ' Check existence of the file
    Dim fso As New FileSystemObject
    If fso.FileExists(listPath) = False Then Err.Raise VTK_WRONG_FILE_PATH

    ' Load the list
    Dim dom As New MSXML2.DOMDocument
    dom.Load listPath
    
    ' Filter projects with the same name
    Dim tmpNode As MSXML2.IXMLDOMNode
    For Each tmpNode In dom.ChildNodes.Item(1).ChildNodes
        If tmpNode.ChildNodes.Item(0).Text Like projectName Then Err.Raise VTK_PROJECT_ALREADY_IN_LIST
    Next

    ' Insert a project node in the root node
    With dom.ChildNodes.Item(1).appendChild(dom.createElement("project"))
        
        'Project name
        With .appendChild(dom.createElement("name"))
            .Text = projectName
        End With
        
        ' Project root folder
        With .appendChild(dom.createElement("beforeRootFolder"))
            .Text = projectRootFolder
        End With
        
        ' Relative path of the xml file
        With .appendChild(dom.createElement("xmlRelativePath"))
            .Text = projectXMLRelativePath
        End With
        
    End With
    
    ' Save changes to the list
    vtkWriteXMLDOMToFile dom, listPath

    On Error GoTo 0
    Exit Sub

vtkAddProjectToListOfRememberedProjects_Error:
    Err.Source = "vtkAddProjectToListOfRememberedProjects of module vtkXMLUtilities"
    
    Select Case Err.Number
        Case VTK_PROJECT_ALREADY_IN_LIST
            Err.Description = "There is already a project with that name in the list."
        Case VTK_WRONG_FILE_PATH
            Err.Description = "The file path you specified is wrong. Make sure the folder tree is valid."
        Case Else
            Err.Number = VTK_UNEXEPECTED_ERROR
    End Select
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Sub

End Sub


'---------------------------------------------------------------------------------------
' Procedure : vtkModifyProjectInList
' Author    : Lucas Vitorino
' Purpose   : Modify the field of a given project in the project list.
' Raises    : VTK_WRONG_FILE_PATH
'             VTK_UNEXEPECTED_ERROR
'             VTK_NO_SUCH_PROJECT
' Notes     : It's impossible to modify the name of a project.
'---------------------------------------------------------------------------------------
'
Public Sub vtkModifyProjectInList(listPath As String, _
                                  projectName As String, _
                                  Optional projectBeforeRootFolder, _
                                  Optional projectXMLRelativePath)
                                               
    On Error GoTo vtkModifyProjectInList_Error

    ' Check existence of the file
    Dim fso As New FileSystemObject
    If fso.FileExists(listPath) = False Then Err.Raise VTK_WRONG_FILE_PATH

    ' Load the list
    Dim dom As New MSXML2.DOMDocument
    dom.Load listPath
    
    ' For all the childnodes of the root node
    Dim projectFound As Boolean: projectFound = False
    Dim tmpNode As MSXML2.IXMLDOMNode
    For Each tmpNode In dom.ChildNodes.Item(1).ChildNodes
        ' If the name of the node is the one given as a parameter
        If tmpNode.ChildNodes.Item(0).Text Like projectName Then
            projectFound = True
            ' Update projectRootFolder if needed
            If Not IsEmpty(projectBeforeRootFolder) Then tmpNode.ChildNodes.Item(1).Text = projectBeforeRootFolder
            ' Update projectXMLRelativePath if needed
            If Not IsEmpty(projectXMLRelativePath) Then tmpNode.ChildNodes.Item(2).Text = projectXMLRelativePath
        End If
    Next
    
    ' Raise error if the project has not been found in the list
    If Not projectFound Then Err.Raise VTK_NO_SUCH_PROJECT
    
    ' Save changes to the list
    vtkWriteXMLDOMToFile dom, listPath

    On Error GoTo 0
    Exit Sub

vtkModifyProjectInList_Error:
    Err.Source = "vtkModifyProjectInList of module vtkXMLUtilities"
    
    Select Case Err.Number
        Case VTK_WRONG_FILE_PATH
            Err.Description = "The file path you specified is wrong. Make sure the folder tree is valid."
        Case VTK_NO_SUCH_PROJECT
            Err.Description = "No project with this name has been found in the list."
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Sub
End Sub
                                  

'---------------------------------------------------------------------------------------
' Procedure : vtkRemoveProjectFromList
' Author    : Lucas Vitorino
' Purpose   : Removes a project from the list of projects
' Raises    : VTK_UNEXPECTED_ERROR
'             VTK_WRONG_FILE_PATH
'             VTK_NO_SUCH_PROJECT
'---------------------------------------------------------------------------------------
'
Public Sub vtkRemoveProjectFromList(listPath As String, projectName As String)

    On Error GoTo vtkRemoveProjectFromList_Error

    Dim tmpNode As MSXML2.IXMLDOMNode

    ' Check existence of the file
    Dim fso As New FileSystemObject
    If fso.FileExists(listPath) = False Then Err.Raise VTK_WRONG_FILE_PATH

    ' Load the list
    Dim dom As New MSXML2.DOMDocument
    dom.Load listPath
    
    ' Main loop
    Dim index As Integer: index = 0
    Dim projectFound As Boolean: projectFound = False
    For Each tmpNode In dom.ChildNodes.Item(1).ChildNodes
        ' If the name of the node is the one given as a parameter
        If tmpNode.ChildNodes.Item(0).Text Like projectName Then
            ' Remove this node
            dom.ChildNodes.Item(1).RemoveChild dom.ChildNodes.Item(1).ChildNodes.Item(index)
            projectFound = True
        End If
        index = index + 1
    Next
    
    ' Raise error if the project has not been found in the list
    If Not projectFound Then Err.Raise VTK_NO_SUCH_PROJECT
    
    ' Save changes to the list
    vtkWriteXMLDOMToFile dom, listPath

    On Error GoTo 0
    Exit Sub

vtkRemoveProjectFromList_Error:
    Err.Source = "vtkModifyProjectInList of module vtkXMLUtilities"
    
    Select Case Err.Number
        Case VTK_WRONG_FILE_PATH
            Err.Description = "The file path you specified is wrong. Make sure the folder tree is valid."
        Case VTK_NO_SUCH_PROJECT
            Err.Description = "No project with this name has been found in the list."
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Sub
End Sub



'---------------------------------------------------------------------------------------
' UNTESTED UTILITY FUNCTIONS
'---------------------------------------------------------------------------------------
'
Public Function getProjectBeforeRootPathInList(listPath As String, projectName As String) As String
    
    ' Load the list
    Dim dom As New MSXML2.DOMDocument
    dom.Load listPath
    
    ' For all the childnodes of the root node
    Dim tmpNode As MSXML2.IXMLDOMNode
    For Each tmpNode In dom.ChildNodes.Item(1).ChildNodes
        ' If the name of the node is the one given as a parameter
        If tmpNode.ChildNodes.Item(0).Text Like projectName Then
            getProjectBeforeRootPathInList = tmpNode.ChildNodes.Item(1).Text
            Exit Function
        End If
    Next

End Function

Public Function getProjectXMLRelativePathInList(listPath As String, projectName As String) As String

    ' Load the list
    Dim dom As New MSXML2.DOMDocument
    dom.Load listPath
    
    ' For all the childnodes of the root node
    Dim tmpNode As MSXML2.IXMLDOMNode
    For Each tmpNode In dom.ChildNodes.Item(1).ChildNodes
        ' If the name of the node is the one given as a parameter
        If tmpNode.ChildNodes.Item(0).Text Like projectName Then
            getProjectXMLRelativePathInList = tmpNode.ChildNodes.Item(2).Text
            Exit Function
        End If
    Next

End Function
