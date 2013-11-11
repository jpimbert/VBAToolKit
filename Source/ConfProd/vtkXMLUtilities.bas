Attribute VB_Name = "vtkXMLUtilities"
Option Explicit
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
' Procedure : vtkExportConfigurationsAsXML
' Author    : Jean-Pierre IMBERT
' Date      : 07/11/2013
' Purpose   : Export the configurations of a project as a new XML file
'             - the "new" XML file format is the one designed for configuration management as XML
'             - this subroutine is temporary, dedicated to prepare the migration from
'               Excel sheet management of configurations to XML file management
' Raises    : - VTK_WORKBOOK_NOT_OPEN if the _DEV workbook containing the configuration sheet is not opened
'             - VTK_WRONG_FILE_PATH if the file path couldn't be created
'---------------------------------------------------------------------------------------
'
Public Sub vtkExportConfigurationsAsXML(projectName As String, filePath As String)

   On Error GoTo vtkExportConfigurationsAsXML_Error

    ' Get the configurationManager of the project to export
    Dim cm As vtkConfigurationManager
    Set cm = vtkConfigurationManagerForProject(projectName)
    If cm Is Nothing Then
        Err.Raise Number:=VTK_WORKBOOK_NOT_OPEN
    End If
    
    ' Create a new XML configuration file
    Dim fso As New FileSystemObject
    Dim xmlFile As TextStream
    Set xmlFile = fso.CreateTextFile(fileName:=filePath, Overwrite:=True)

    ' Create the XML preamble
    xmlFile.WriteLine Text:="<?xml version=""1.0"" encoding=""ISO-8859-1"" standalone=""no""?>"
    xmlFile.WriteLine Text:="<!DOCTYPE vtkConf SYSTEM ""vtkConfigurationsDTD.dtd"">"
    xmlFile.WriteLine Text:="<vtkConf>"
    xmlFile.WriteBlankLines Lines:=1
    xmlFile.WriteLine Text:="    <info>"
    xmlFile.WriteLine Text:="        <vtkConfigurationsVersion>1.0</vtkConfigurationsVersion>"
    xmlFile.WriteLine Text:="        <projectName>" & projectName & "</projectName>"
    xmlFile.WriteLine Text:="    </info>"
    xmlFile.WriteBlankLines Lines:=1
    
    ' Create Configuration elements
    Dim cf As vtkConfiguration
    For Each cf In cm.configurations
        ' create the configuration element
        xmlFile.WriteLine Text:="    <configuration cID=""" & cf.ID & """>"
        xmlFile.WriteLine Text:="        <name>" & cf.name & "</name>"
        xmlFile.WriteLine Text:="        <path>" & cf.path & "</path>"
        xmlFile.WriteLine Text:="        <title>VBAToolKit</title>" ' must be initialized in Workbook with Wb.BuiltinDocumentProperties("Title").Value
        xmlFile.WriteLine Text:="        <comment>Toolkit improving IDE for VBA projects</comment>" ' must be initialized in Workbook with Wb.BuiltinDocumentProperties("Comments").Value
        xmlFile.WriteLine Text:="    </configuration>"
    Next
    
    ' Create Module elements
    Dim md As vtkModule, modulepath As String
    For Each md In cm.modules
        xmlFile.WriteLine Text:="    <module mID=""" & md.ID & """>"
        xmlFile.WriteLine Text:="        <name>" & md.name & "</name>"
        For Each cf In cm.configurations
            modulepath = md.getPathForConfiguration(confName:=cf.name)
            If Not modulepath Like "" Then xmlFile.WriteLine Text:="        <modulePath confId=""" & cf.ID & """>" & modulepath & "</modulePath>"
        Next
        xmlFile.WriteLine Text:="    </module>"
    Next
    
    ' Create References elements
    '   Designed only to be used with an active Workbook being a _DEV VBA project
    '   All references are exported to the XML configuration file
    '   except the VBAToolKit reference that is exported only for the _DEV configuration
    Dim r As VBIDE.Reference, allConfIDs As String, confIDsOnlyDEV As String
    ' Build ConfIDs lists
    For Each cf In cm.configurations
        If Not allConfIDs Like "" Then
            allConfIDs = allConfIDs & " " & cf.ID
           Else
            allConfIDs = """" & cf.ID
        End If
        If cf.isDEV Then
            If Not confIDsOnlyDEV Like "" Then
                confIDsOnlyDEV = confIDsOnlyDEV & " " & cf.ID
               Else
                confIDsOnlyDEV = """" & cf.ID
            End If
        End If
    Next
    allConfIDs = allConfIDs & """"
    confIDsOnlyDEV = confIDsOnlyDEV & """"
    ' Add reference elements to the XML configuration file
    For Each r In ActiveWorkbook.VBProject.references
        If r.name Like "VBAToolKit" Then
            xmlFile.WriteLine Text:="    <reference confIDs=" & confIDsOnlyDEV & ">"
           Else
            xmlFile.WriteLine Text:="    <reference confIDs=" & allConfIDs & ">"
        End If
        xmlFile.WriteLine Text:="        <name>" & r.name & "</name>"
        If r.guid Like "" Then
            xmlFile.WriteLine Text:="        <path>" & r.fullPath & "</path>"
           Else
            xmlFile.WriteLine Text:="        <guid>" & r.guid & "</guid>"
        End If
        xmlFile.WriteLine Text:="    </reference>"
    Next
    
    ' Close the file
    xmlFile.WriteLine Text:="</vtkConf>"
    xmlFile.Close
    
   On Error GoTo 0
   Exit Sub

vtkExportConfigurationsAsXML_Error:
    Dim s As String
    s = "vtkXMLutilities::exportConfigurationsAsXML"
    
    Select Case Err.Number
        Case VTK_WORKBOOK_NOT_OPEN
            Err.Description = "The " & projectName & "_DEV workbook is not opened"
        Case 76
            Err.Number = VTK_WRONG_FILE_PATH
            Err.Description = "The " & filePath & " path is unreachable"
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
            s = s & " -> " & Err.Source
    End Select

    Err.Raise Err.Number, s, Err.Description
End Sub

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
    
    Dim oStream As ADODB.Stream
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
' Notes     : Useless now
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
                                                   projectRootFolder As String, _
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
        With .appendChild(dom.createElement("rootFolder"))
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
            Err.Number = VTK_UNEXPECTED_ERROR
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
                                  Optional projectRootFolder As String, _
                                  Optional projectXMLRelativePath As String)
                                               
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
            If Not (IsEmpty(projectRootFolder)) Then tmpNode.ChildNodes.Item(1).Text = projectRootFolder
            ' Update projectXMLRelativePath if needed
            If Not (IsEmpty(projectXMLRelativePath)) Then tmpNode.ChildNodes.Item(2).Text = projectXMLRelativePath
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

Public Function countElementsInDom(elementName As String, dom As MSXML2.DOMDocument) As Integer

    On Error GoTo countElementsInDom_Error

    If dom Is Nothing Then
        countElementsInDom = -1
        Exit Function
    End If
    
    Dim rootNode As MSXML2.IXMLDOMNode
    Set rootNode = dom.ChildNodes.Item(1)
    
    countElementsInDom = countElementsInNode(elementName, rootNode)

    On Error GoTo 0
    Exit Function

    On Error GoTo 0
    Exit Function

countElementsInDom_Error:
    Err.Source = "Function countElementsInDom in module vtkXMLUtilities"
    Err.Raise Err.Number, Err.Description, Err.Source
    Exit Function

End Function

Public Function countElementsInNode(elementName As String, node As MSXML2.IXMLDOMNode) As Integer
    
    Dim Count As Integer: Count = 0
    
    On Error GoTo countElementsInNode_Error
    
    If node Is Nothing Then
        countElementsInNode = -1
        Exit Function
    End If

    Dim subNode As MSXML2.IXMLDOMNode
    For Each subNode In node.ChildNodes
        If StrComp(subNode.BaseName, elementName) = 0 Then Count = Count + 1
    Next
        
    countElementsInNode = Count

    On Error GoTo 0
    Exit Function

countElementsInNode_Error:
    Err.Source = "Function countElementsInDom in module vtkXMLUtilities"
    Debug.Print "Error " & Err.Number & " : " & Err.Description & " in " & Err.Source
    Err.Raise Err.Number, Err.Description, Err.Source
    Exit Function
End Function


Public Function getFirstChildNodeByName(nodeName As String, node As MSXML2.IXMLDOMNode) As MSXML2.IXMLDOMNode
    
    On Error GoTo getFirstChildNodeByName_Error

    Dim subNode As MSXML2.IXMLDOMNode
    For Each subNode In node.ChildNodes
        If subNode.BaseName = nodeName Then
            Set getFirstChildNodeByName = subNode
            Exit Function
        End If
    Next
    
    Set getFirstChildNodeByName = Nothing

    On Error GoTo 0
    Exit Function

getFirstChildNodeByName_Error:
    Err.Source = "Function getFirstChildNodeByName in module vtkXMLUtilities"
    Debug.Print "Error " & Err.Number & " : " & Err.Description & " in " & Err.Source
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
End Function

Public Function getProjectRootPathInList(listPath As String, projectName As String) As String
    
    ' Load the list
    Dim dom As New MSXML2.DOMDocument
    dom.Load listPath
    
    ' For all the childnodes of the root node
    Dim tmpNode As MSXML2.IXMLDOMNode
    For Each tmpNode In dom.ChildNodes.Item(1).ChildNodes
        ' If the name of the node is the one given as a parameter
        If tmpNode.ChildNodes.Item(0).Text Like projectName Then
            getProjectRootPathInList = tmpNode.ChildNodes.Item(1).Text
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
