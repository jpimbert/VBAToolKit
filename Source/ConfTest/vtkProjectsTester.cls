VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkProjectsTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkProjectsTester
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Test the vtkProjects module
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

Option Explicit
Implements ITest
Implements ITestCase

Private mManager As TestCaseManager
Private mAssert As IAssert

Private projectName As String      ' default workbook Name for test

Private fso As New FileSystemObject
Private Const xmlFileName As String = "VBAToolKitTestProjects.xml"
Private xmlFileFullPath As String
Private Const testProject1Name As String = "TestProject1"
Private Const testProject1XmlRelPath As String = "TestProject1.xml"
Private Const testProject2Name As String = "TestProject2"
Private Const testProject2XmlRelPath As String = "Project\TestProject2.xml"
Private testProject1RootPath As String
Private testProject2RootPath As String

Private Sub Class_Initialize()
    Set mManager = New TestCaseManager
End Sub

Private Property Get ITestCase_Manager() As TestCaseManager
    Set ITestCase_Manager = mManager
End Property

Private Property Get ITest_Manager() As ITestManager
    Set ITest_Manager = mManager
End Property

Private Sub createTestXML()

    ' Create the processing instruction
    On Error GoTo createTestXML_Error

    Dim dom As New MSXML2.DOMDocument
    dom.appendChild dom.createProcessingInstruction("xml", "version=""1.0"" encoding=""ISO-8859-1""")

    ' Create the root node
    With dom.appendChild(dom.createElement("rememberedProjects"))
    
        With .appendChild(dom.createElement("info"))
            With .appendChild(dom.createElement("version"))
                .Text = "1.0"
            End With
        End With
        
        With .appendChild(dom.createElement("project"))
           With .appendChild(dom.createElement("name"))
                .Text = testProject1Name
            End With
            With .appendChild(dom.createElement("rootFolder"))
                .Text = testProject1RootPath
            End With
            With .appendChild(dom.createElement("xmlRelativePath"))
                .Text = testProject1XmlRelPath
            End With
        End With

        With .appendChild(dom.createElement("project"))
            With .appendChild(dom.createElement("name"))
                .Text = testProject2Name
            End With
            
            With .appendChild(dom.createElement("rootFolder"))
                .Text = testProject2RootPath
            End With
                
            With .appendChild(dom.createElement("xmlRelativePath"))
                .Text = testProject2XmlRelPath
            End With
        End With

    End With
    
    vtkWriteXMLDOMToFile dom, xmlFileFullPath

    On Error GoTo 0
    Exit Sub

createTestXML_Error:
    Err.Source = "Sub createTestXML in module vtkProjectsTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub



Private Sub ITestCase_SetUp(Assert As IAssert)
    Set mAssert = Assert
    projectName = "TestProject"
    
    VBAToolKit.vtkResetProjects
    
    ' set the missing consts (calling function so can't be set out of a function/sub)
    testProject1RootPath = fso.BuildPath(VBAToolKit.vtkTestPath, testProject1Name)
    testProject2RootPath = fso.BuildPath(VBAToolKit.vtkTestPath, testProject2Name)
    xmlFileFullPath = fso.BuildPath(VBAToolKit.vtkTestPath, xmlFileName)
    
    ' Indicate to the current workbook where is the list of projects
    xmlRememberedProjectsFullPath = xmlFileFullPath
    
    ' Create the file
    createTestXML
    
    
End Sub

Private Sub ITestCase_TearDown()
    vtkResetProjects
    VBAToolKit.resetTestFolder
End Sub

Public Sub TestProjectIsNotNull()
'       Verify that the vtkProject is not null
    mAssert.Should Not vtkProjectForName(projectName) Is Nothing, "A project must not be null"
End Sub

Public Sub TestProjectsAreDifferentForTwoNames()
'       Verify that the vtkProject-s for two projects are different
    mAssert.Should Not vtkProjectForName(projectName) Is vtkProjectForName("SecondProject"), "Two projects must not be equal"
End Sub

Public Sub TestProjectIsInvariant()
'       Verify that the vtkProject for a project is always the same
    Dim project As vtkProject
    Set project = vtkProjectForName(projectName)
    mAssert.Should vtkProjectForName(projectName) Is project, "The project for a project name must be always the same"
End Sub


' Update of vtkProjects
' --------------------
'
' The new functionalities of vtkProjects must be tested
'
' - projects in the file must be read
' - a sheet with a wrong format is not read and raises an error
' - modifying or deleting the xml file after it has been loaded does not have any effect
' - the export of the sheet must be readable, contain the information and be conform
' - add a project via the "add remembered project" function must also add it to the sheet
' - remove a project must do the same
'

Public Sub Test_vtkProjects_listIsRead()
' projects in the file must be read

    On Error GoTo Test_vtkProjects_listIsRead_Error

    mAssert.Equals vtkRootPathForProject("TestProject1"), testProject1RootPath, _
                   "vtkProjects should access to the information in the sheet."

    mAssert.Equals vtkXmlRelPathForProject("TestProject1"), testProject1XmlRelPath, _
                    "vtkProjects should access to the information in the sheet."

    On Error GoTo 0
    Exit Sub

Test_vtkProjects_listIsRead_Error:
    Err.Source = "Sub Test_vtkProjects_listIsRead in module vtkProjectsTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

' will do later with method from JPI
'Public Sub Test_vtkProjects_listHasWrongFormat()
'' a sheet with a wrong format is not read and raises an error
'
'On Error GoTo Test_vtkProjects_listHasWrongFormat_Error
'
'    Dim error As Integer: error = 0
'
'
'
'    On Error GoTo 0
'    Exit Sub
'
'Test_vtkProjects_listHasWrongFormat_Error:
'    Err.Source = "Sub Test_vtkProjects_listHasWrongFormat in module vtkProjectsTester"
'    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
'    Exit Sub
'
'End Sub


Public Sub Test_vtkProjects_modifyXmlFileAfterLoading()
' modifying or deleting the xml file after it has been loaded should not have any effect

    On Error GoTo Test_vtkProjects_modifyXmlFileAfterLoading_Error

    Dim beforeModPathStr As String, afterModPathStr As String

    beforeModPathStr = vtkRootPathForProject(testProject1Name)

    vtkModifyProjectInList listPath:=xmlFileFullPath, _
                           projectName:=testProject1Name, _
                           projectRootFolder:=fso.BuildPath(VBAToolKit.vtkTestPath, "\dummy")

    afterModPathStr = vtkRootPathForProject(testProject1Name)

    mAssert.Equals beforeModPathStr, afterModPathStr, _
                   "The root path should not have been modified in the data structures of the program, only in the file."


    On Error GoTo 0
    Exit Sub

Test_vtkProjects_modifyXmlFileAfterLoading_Error:
    Err.Source = "Sub Test_vtkProjects_modifyXmlFileAfterLoading in module vtkProjectsTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub Test_vtkAddRememberedProject_projectIsAddedToTheCollections()
' after using the "vtkAddRememberedProject" function, collections must be updated

    On Error GoTo Test_vtkAddRememberedProject_projectIsAddedToTheCollections_Error

    ' Generate the names for the new project
    Dim newProjectName As String: newProjectName = "NewProject"
    Dim newProjectRootFolder As String: newProjectRootFolder = fso.BuildPath(VBAToolKit.vtkTestPath, newProjectName)
    Dim newProjectXmlRelPath As String: newProjectXmlRelPath = newProjectName & ".xml"

    ' Add it
    vtkAddRememberedProject projectName:=newProjectName, _
                            rootFolder:=newProjectRootFolder, _
                            xmlRelativePath:=newProjectXmlRelPath

    mAssert.Equals vtkRootPathForProject(newProjectName), newProjectRootFolder, _
                   "The collection cotaining the root paths should have been updated."

    mAssert.Equals vtkXmlRelPathForProject(newProjectName), newProjectXmlRelPath, _
                   "The collection cotaining the xml relative paths should have been updated."

    On Error GoTo 0
    Exit Sub

Test_vtkAddRememberedProject_projectIsAddedToTheCollections_Error:
    Err.Source = "Sub Test_vtkAddRememberedProject_projectIsAddedToTheCollections in module vtkProjectsTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub Test_vtkAddRememberedProject_projectIsAddedToTheList()
' after using the "vtkAddRememberedProject" function, list must be updated

    On Error GoTo Test_vtkAddRememberedProject_projectIsAddedToTheList_Error

    ' Generate the names for the new project
    Dim newProjectName As String: newProjectName = "NewProject"
    Dim newProjectRootFolder As String: newProjectRootFolder = fso.BuildPath(VBAToolKit.vtkTestPath, newProjectName)
    Dim newProjectXmlRelPath As String: newProjectXmlRelPath = newProjectName & ".xml"
    
    ' Add it
    vtkAddRememberedProject projectName:=newProjectName, _
                            rootFolder:=newProjectRootFolder, _
                            xmlRelativePath:=newProjectXmlRelPath
                            
    ' Load the dom from the updated file
    Dim dom As New MSXML2.DOMDocument
    dom.Load xmlFileFullPath
    
    ' Another rain of mAssert, is it justified here ?
    mAssert.Equals dom.getElementsByTagName("project").Length, 3, "The project must have been added to the list"
    
    mAssert.Equals dom.getElementsByTagName("name").Item(2).Text, _
                   newProjectName, _
                   "The project must have been added to the list with its name tag"
    
    mAssert.Equals dom.getElementsByTagName("rootFolder").Item(2).Text, _
                   newProjectRootFolder, _
                   "The project must have been added to the list with its rootFolder tag"
                   
    mAssert.Equals dom.getElementsByTagName("xmlRelativePath").Item(2).Text, _
                    newProjectXmlRelPath, _
                    "The project must have been added to the list with its xmlRelativePath tag"
    
    On Error GoTo 0
    Exit Sub

Test_vtkAddRememberedProject_projectIsAddedToTheList_Error:
    Err.Source = "Sub Test_vtkAddRememberedProject_projectIsAddedToTheList in module vtkProjectsTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub Test_vtkAddRememberdProject_notRememberedProjectIsNotAdded()
' a project added with vtkProjectForName instead of vtkAddRememberedProject should not be written in the list
    
    ' Generate the names for the new project
    Dim newProjectName As String: newProjectName = "NewProject"
    Dim newProjectRootFolder As String: newProjectRootFolder = fso.BuildPath(VBAToolKit.vtkTestPath, newProjectName)
    Dim newProjectXmlRelPath As String: newProjectXmlRelPath = newProjectName & ".xml"

    ' Adding a not remembered project to the collection
    On Error GoTo Test_vtkAddRememberdProject_notRememberedProjectIsNotAdded_Error

    vtkProjectForName ("notRememberedProject")

    ' Add the remembered project
    vtkAddRememberedProject projectName:=newProjectName, _
                            rootFolder:=newProjectRootFolder, _
                            xmlRelativePath:=newProjectXmlRelPath
                            
    ' Load the dom from the updated file
    Dim dom As New MSXML2.DOMDocument
    dom.Load xmlFileFullPath
    
    ' Make sure the not remembered project has not been remembered
    mAssert.Equals dom.getElementsByTagName("project").Length, 3, "The not remembered project must not have been added to the list."

    On Error GoTo 0
    Exit Sub

Test_vtkAddRememberdProject_notRememberedProjectIsNotAdded_Error:
    Err.Source = "Sub Test_vtkAddRememberdProject_notRememberedProjectIsNotAdded in module vtkProjectsTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub Test_vtkRemoveRememberedProject_projectIsRemovedFromTheCollections()
' vtkRemoveRememberedProject should remove a given project from the collections

    On Error GoTo Test_vtkRemoveRememberedProject_projectIsRemovedFromTheCollections_Error

    vtkRemoveRememberedProject testProject1Name
    
    mAssert.Equals vtkRootPathForProject(projectName), "", _
                   "Project must have been deleted from the root paths collection"

    mAssert.Equals vtkXmlRelPathForProject(projectName), "", _
                   "Project must have been deleted from the xml relative paths collection"

    On Error GoTo 0
    Exit Sub

Test_vtkRemoveRememberedProject_projectIsRemovedFromTheCollections_Error:
    Err.Source = "Sub Test_vtkRemoveRememberedProject_projectIsRemovedFromTheCollections in module vtkProjectsTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub Test_vtkRemoveRememberedProject_projectIsRemovedFromTheList()
' vtkRemoveRememberedProject should remove a given project from the collections

    On Error GoTo Test_vtkRemoveRememberedProject_projectIsRemovedFromTheList_Error

    vtkRemoveRememberedProject testProject1Name

    ' Load the dom from the updated file
    Dim dom As New MSXML2.DOMDocument
    dom.Load xmlFileFullPath
    
    ' Are these tests enough ?
    mAssert.Equals dom.getElementsByTagName("project").Length, 1, "Only one project should remain in the list."
    mAssert.Equals dom.getElementsByTagName("name").Item(0).Text, testProject2Name, "The only remaining project should be TestProject2"
    
    On Error GoTo 0
    Exit Sub

Test_vtkRemoveRememberedProject_projectIsRemovedFromTheList_Error:
    Err.Source = "Sub Test_vtkRemoveRememberedProject_projectIsRemovedFromTheList in module vtkProjectsTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Public Sub Test_vtkModifyRememberedProject_projectIsModifiedInTheCollections()
' vtkModifyRememberedProject should modify the given attribute of a project (folderPath or xmlRelPath, not name) in a collection

    On Error GoTo Test_vtkModifyRememberedProject_projectIsModifiedInTheCollections_Error

    vtkModifyRememberedProject testProject1Name, "dummy", "dummy2"

    mAssert.Equals vtkRootPathForProject(testProject1Name), "dummy"
    mAssert.Equals vtkXmlRelPathForProject(testProject1Name), "dummy2"

    On Error GoTo 0
    Exit Sub

Test_vtkModifyRememberedProject_projectIsModifiedInTheCollections_Error:
    Err.Source = "Sub Test_vtkModifyRememberedProject_projectIsModifiedInTheCollections in module vtkProjectsTester"
    Debug.Print "Error " & Err.Number & " : " & Err.Description & " in " & Err.Source
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Public Sub Test_vtkModifyRememberedProject_projectIsModifiedInTheList()
' vtkModifyRememberedProject should modify the given attribute of a project (folderPath or xmlRelPath, not name) in the list

    On Error GoTo Test_vtkModifyRememberedProject_projectIsModifiedInTheList_Error

    vtkModifyRememberedProject testProject1Name, "dummy", "dummy2"

    ' Load the dom from the updated file
    Dim dom As New MSXML2.DOMDocument
    dom.Load xmlFileFullPath

    ' Are these tests enough ?
    mAssert.Equals dom.getElementsByTagName("project").Length, 2
    mAssert.Equals dom.getElementsByTagName("rootFolder").Item(0).Text, "dummy"
    mAssert.Equals dom.getElementsByTagName("xmlRelativePath").Item(0).Text, "dummy2"

    On Error GoTo 0
    Exit Sub

Test_vtkModifyRememberedProject_projectIsModifiedInTheList_Error:
    Err.Source = "Sub Test_vtkModifyRememberedProject_projectIsModifiedInTheList in module vtkProjectsTester"
    Debug.Print "Error " & Err.Number & " : " & Err.Description & " in " & Err.Source
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectIsNotNull"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectsAreDifferentForTwoNames"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectIsInvariant"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkProjects_listIsRead"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkProjects_modifyXmlFileAfterLoading"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkAddRememberedProject_projectIsAddedToTheCollections"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkAddRememberedProject_projectIsAddedToTheList"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkAddRememberdProject_notRememberedProjectIsNotAdded"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRemoveRememberedProject_projectIsRemovedFromTheCollections"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRemoveRememberedProject_projectIsRemovedFromTheList"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkModifyRememberedProject_projectIsModifiedInTheCollections"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkModifyRememberedProject_projectIsModifiedInTheList"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestProjectIsNotNull": TestProjectIsNotNull
        Case "TestProjectsAreDifferentForTwoNames": TestProjectsAreDifferentForTwoNames
        Case "TestProjectIsInvariant": TestProjectIsInvariant
        Case "Test_vtkProjects_listIsRead": Test_vtkProjects_listIsRead
        Case "Test_vtkProjects_modifyXmlFileAfterLoading": Test_vtkProjects_modifyXmlFileAfterLoading
        Case "Test_vtkAddRememberedProject_projectIsAddedToTheCollections": Test_vtkAddRememberedProject_projectIsAddedToTheCollections
        Case "Test_vtkAddRememberedProject_projectIsAddedToTheList": Test_vtkAddRememberedProject_projectIsAddedToTheList
        Case "Test_vtkAddRememberdProject_notRememberedProjectIsNotAdded": Test_vtkAddRememberdProject_notRememberedProjectIsNotAdded
        Case "Test_vtkRemoveRememberedProject_projectIsRemovedFromTheCollections": Test_vtkRemoveRememberedProject_projectIsRemovedFromTheCollections
        Case "Test_vtkRemoveRememberedProject_projectIsRemovedFromTheList": Test_vtkRemoveRememberedProject_projectIsRemovedFromTheList
        Case "Test_vtkModifyRememberedProject_projectIsModifiedInTheCollections": Test_vtkModifyRememberedProject_projectIsModifiedInTheCollections
        Case "Test_vtkModifyRememberedProject_projectIsModifiedInTheList": Test_vtkModifyRememberedProject_projectIsModifiedInTheList
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub



