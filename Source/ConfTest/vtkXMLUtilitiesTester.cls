VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkXMLUtilitiesTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkXMLUtilitiesTester
' Author    : Lucas Vitorino
' Purpose   : Test the utilities contained in the vtkXMLUtilities, except vtkExportAsXML.
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

Private fso As New FileSystemObject

Private Const listName As String = "projects_list.xml"
Private Const testProjectName As String = "TestProject"
Private listFullPath As String
Private testProjectRootPath As String
Private testProjectXMLPath As String

Private error As Integer

Private Sub Class_Initialize()
    Set mManager = New TestCaseManager
End Sub

Private Property Get ITestCase_Manager() As TestCaseManager
    Set ITestCase_Manager = mManager
End Property

Private Property Get ITest_Manager() As ITestManager
    Set ITest_Manager = mManager
End Property

Private Sub ITestCase_SetUp(Assert As IAssert)

    'TMP
    resetTestFolder

    Set mAssert = Assert
    
    vtkCreateTreeFolder fso.BuildPath(VBAToolKit.vtkTestPath, testProjectName)
    testProjectRootPath = fso.BuildPath(VBAToolKit.vtkTestPath, testProjectName)
    testProjectXMLPath = "Project\"
    
    listFullPath = fso.BuildPath(VBAToolKit.vtkTestPath, listName)
    
    
    error = 0

End Sub

Private Sub ITestCase_TearDown()
    VBAToolKit.resetTestFolder
End Sub

' Tests to do
'
' vtkCreateListOfRememberedProjects
'   - correct path : file exists
'   - correct path : file content is right
'       - XML version
'       - rootNode name
'   - wrong path : raise error VTK_WRONG_FILE_PATH
'
' vtkAddProjectToList
'   - correct path, list empty : file added normally
'   - correct path, list has already one project with a different name : file added normally
'   - wrong path : raise error VTK_WRONG_FILE_PATH
'   - correct path, list has already one project with the same name  : raise error VTK_PROJECT_ALREADY_IN_LIST
'
' vtkModifyProjectInList
'   - correct path : fields updated
'   - wrong path : VTK_WRONG_FILE_PATH
'   - correct path, no such project : VTK_NO_SUCH_PROJECT
'
' vtkRemoveProjectFromList
'   - correct path : project removed
'   - wrong path : VTK_WRONG_FILE_PATH
'   - correct path, no such project : VTK_NO_SUCH_PROJECT
'
'

Public Sub Test_vtkCreateListOfRememberedProjects_correctPathExistence()
' correct path : file exists

    On Error GoTo Test_vtkCreateListOfRememberedProjects_correctPathExistence_Error
    
    vtkCreateListOfRememberedProjects (listFullPath)
    
    mAssert.Should fso.FileExists(listFullPath), _
                   "File should have been created as the path is correct."

    On Error GoTo 0
    Exit Sub

Test_vtkCreateListOfRememberedProjects_correctPathExistence_Error:
    Err.Source = "Test_vtkCreateXMLListOfRememberedProjects_correctPathExistence of module vtkXMLUtilitiesTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkCreateListOfRememberedProjects_correctPathContent()
' correct path : file content is as expecte

    On Error GoTo Test_vtkCreateListOfRememberedProjects_correctPathContent_Error

    vtkCreateListOfRememberedProjects (listFullPath)

    mAssert.Equals vtkTextFileReader(listFullPath), _
                     Chr(10) & _
                     "<?xml version=""1.0"" encoding=""ISO-8859-1"" standalone=""no""?>" & Chr(10) & _
                     "<rememberedProjects/>", _
                     "List has not been initialized properly."

    On Error GoTo 0
    Exit Sub

Test_vtkCreateListOfRememberedProjects_correctPathContent_Error:
    Err.Source = "Test_vtkCreateListOfRememberedProjects_correctPathContent of module vtkXMLUtilitiesTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkCreateListOfRememberedProjects_wrongPathExistence()
' wrong path

    On Error GoTo Test_vtkCreateListOfRememberedProjects_wrongPathExistence_Error

    Dim error As Integer

    listFullPath = fso.BuildPath(VBAToolKit.vtkTestPath, "dummy_folder\list.xml")

    vtkCreateListOfRememberedProjects (listFullPath)

    mAssert.Equals error, VTK_WRONG_FILE_PATH, "An error should have been raised as the folder tree is not valid."

    mAssert.Should fso.FileExists(listFullPath) = False, "The file shouldn't have been created."

    On Error GoTo 0
    Exit Sub

Test_vtkCreateListOfRememberedProjects_wrongPathExistence_Error:
    error = Err.Number
    Exit Sub
End Sub



Public Sub Test_vtkAddProjectToListOfRememberedProjects_correctPath()
' correct path, list empty : file added normally

    On Error GoTo Test_vtkAddvtkAddProjectToListOfRememberedProjects_correctPath_Error

    vtkCreateListOfRememberedProjects (listFullPath)

    vtkAddProjectToListOfRememberedProjects listFullPath, testProjectName, testProjectRootPath, testProjectXMLPath

    Dim dom As New MSXML2.DOMDocument
    dom.Load listFullPath

    mAssert.Equals countElementsInDom("project", dom), 1, _
                   "The project has not been added properly."

    On Error GoTo 0
    Exit Sub

Test_vtkAddvtkAddProjectToListOfRememberedProjects_correctPath_Error:
    Err.Source = "Test_vtkAddProjectToListOfRememberedProjects_correctPath of module vtkXMLUtilitiesTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub Test_vtkAddProjectToListOfRememberedProjects_correctPathDifferentName()
' correct path, list has already one project with a different name : project added normally

    On Error GoTo Test_vtkAddProjectToListOfRememberedProjects_correctPathDifferentName_Error

    vtkCreateListOfRememberedProjects (listFullPath)

    ' Add 2 projects with different names
    vtkAddProjectToListOfRememberedProjects listFullPath, testProjectName, testProjectRootPath, testProjectXMLPath
    vtkAddProjectToListOfRememberedProjects listFullPath, testProjectName & "2", testProjectRootPath, testProjectXMLPath

    Dim dom As New MSXML2.DOMDocument
    dom.Load listFullPath

    mAssert.Equals countElementsInDom("project", dom), 2, _
                   "The project has not been added properly."


    mAssert.Equals error, 0, "No error should have occured as the project names are different"

    On Error GoTo 0
    Exit Sub

Test_vtkAddProjectToListOfRememberedProjects_correctPathDifferentName_Error:
    error = Err.Number
    Resume Next
End Sub



Public Sub Test_vtkAddProjectToListOfRememberedProjects_correctPathSameName()
' correct path, list has already one project with a same name : project not added and error raised

    On Error GoTo Test_vtkAddProjectToListOfRememberedProjects_correctPathSameName_Error

    vtkCreateListOfRememberedProjects (listFullPath)

    ' Add 2 projects with same names
    vtkAddProjectToListOfRememberedProjects listFullPath, testProjectName, testProjectRootPath, testProjectXMLPath
    vtkAddProjectToListOfRememberedProjects listFullPath, testProjectName, testProjectRootPath, testProjectXMLPath

    Dim dom As New MSXML2.DOMDocument
    dom.Load listFullPath

    mAssert.Equals countElementsInDom("project", dom), 1, _
                   "The project should not have been added."

    mAssert.Equals error, VTK_PROJECT_ALREADY_IN_LIST, _
                   "The project is already in the list, an error should have been raised."

    On Error GoTo 0
    Exit Sub

Test_vtkAddProjectToListOfRememberedProjects_correctPathSameName_Error:
    error = Err.Number
    Resume Next
End Sub



Public Sub Test_vtkModifyProjectInList_correctPath()
' correct path : fields updated

    On Error GoTo Test_vtkModifyProjectInList_correctPath_Error
    
    ' Create list and add project
    vtkCreateListOfRememberedProjects (listFullPath)
    vtkAddProjectToListOfRememberedProjects listFullPath, testProjectName, testProjectRootPath, testProjectXMLPath

    ' Modify the fields we want
    vtkModifyProjectInList listFullPath, testProjectName, "foo", "bar"
    
    ' Test
    mAssert.Equals getProjectRootPathInList(listFullPath, testProjectName), "foo", _
                   "Project root path should have been changed."
                   
    mAssert.Equals getProjectXMLRelativePathInList(listFullPath, testProjectName), "bar", _
                   "Project relative path should have been changed."
                    
    On Error GoTo 0
    Exit Sub

Test_vtkModifyProjectInList_correctPath_Error:
    Err.Source = "Test_vtkModifyProjectInList_correctPath of module vtkXMLUtilitiesTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkModifyProjectInList_wrongPath()
' wrong path : raise error VTK_WRONG_FILE_PATH

    On Error GoTo Test_vtkModifyProjectInList_wrongPath_Error
    
    vtkModifyProjectInList fso.BuildPath(VBAToolKit.vtkTestPath, "dummyFolder\dummyList.xml"), testProjectName, "foo", "bar"

    mAssert.Equals error, VTK_WRONG_FILE_PATH, _
                   "File path is wrong, the relevant error should have been raised."

    On Error GoTo 0
    Exit Sub

Test_vtkModifyProjectInList_wrongPath_Error:
    error = Err.Number
    Resume Next
End Sub



Public Sub Test_vtkModifyProjectInList_correctPathNoSuchProject()
' correct path, no such project : raise error VTK_NO_SUCH_PROJECT

    On Error GoTo Test_vtkModifyProjectInList_correctPathNoSuchProject_Error

    ' Create list and add project
    vtkCreateListOfRememberedProjects (listFullPath)
    vtkAddProjectToListOfRememberedProjects listFullPath, testProjectName, testProjectRootPath, testProjectXMLPath

    ' Modify the fields we want
    vtkModifyProjectInList listFullPath, testProjectName & "2", "foo", "bar"

    mAssert.Equals error, VTK_NO_SUCH_PROJECT, _
                   "No such project has been put in the list, the relevant error should be raised."

    On Error GoTo 0
    Exit Sub

Test_vtkModifyProjectInList_correctPathNoSuchProject_Error:
    error = Err.Number
    Exit Sub
End Sub


Public Sub Test_vtkRemoveProjectFromList_correctPath()

' correct path : project removed

    On Error GoTo Test_vtkRemoveProjectFromList_correctPath_Error
    
    ' Create list and add one project
    vtkCreateListOfRememberedProjects (listFullPath)
    vtkAddProjectToListOfRememberedProjects listFullPath, testProjectName, testProjectRootPath, testProjectXMLPath
    
    ' Remove the project
    vtkRemoveProjectFromList listFullPath, testProjectName
    
    ' Test
    Dim dom As New MSXML2.DOMDocument
    dom.Load listFullPath
    mAssert.Equals countElementsInDom("project", dom), 0, _
                   "The project should have been deleted."
                   
    On Error GoTo 0
    Exit Sub

Test_vtkRemoveProjectFromList_correctPath_Error:
    Err.Source = "Test_vtkRemoveProjectFromList_correctPath of module vtkXMLUtilitiesTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
Exit Sub
End Sub



Public Sub Test_vtkRemoveProjectFromList_wrongPath()
' wrong path : raise error VTK_WRONG_FILE_PATH

    On Error GoTo Test_vtkRemoveProjectFromList_wrongPath_Error
    
    vtkRemoveProjectFromList listFullPath, testProjectName

    mAssert.Equals error, VTK_WRONG_FILE_PATH, _
                   "File path is wrong, the relevant error should have been raised."

    On Error GoTo 0
    Exit Sub

Test_vtkRemoveProjectFromList_wrongPath_Error:
    error = Err.Number
    Resume Next
End Sub



Public Sub Test_vtkRemoveProjectFromList_correctPathNoSuchProject()
' correct path, no such project : raise error VTK_NO_SUCH_PROJECT

    On Error GoTo Test_vtkRemoveProjectFromList_correctPathNoSuchProject_Error

    ' Create list and add project
    vtkCreateListOfRememberedProjects (listFullPath)
    vtkAddProjectToListOfRememberedProjects listFullPath, testProjectName, testProjectRootPath, testProjectXMLPath

    ' Modify the fields we want
    vtkRemoveProjectFromList listFullPath, testProjectName & "2"

    mAssert.Equals error, VTK_NO_SUCH_PROJECT, _
                   "No such project has been put in the list, the relevant error should be raised."

    On Error GoTo 0
    Exit Sub

Test_vtkRemoveProjectFromList_correctPathNoSuchProject_Error:
    error = Err.Number
    Exit Sub
End Sub


Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateListOfRememberedProjects_correctPathExistence"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateListOfRememberedProjects_correctPathContent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateListOfRememberedProjects_wrongPathExistence"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkAddProjectToListOfRememberedProjects_correctPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkAddProjectToListOfRememberedProjects_correctPathDifferentName"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkAddProjectToListOfRememberedProjects_correctPathSameName"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkModifyProjectInList_correctPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkModifyProjectInList_wrongPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkModifyProjectInList_correctPathNoSuchProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRemoveProjectFromList_correctPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRemoveProjectFromList_wrongPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRemoveProjectFromList_correctPathNoSuchProject"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkCreateListOfRememberedProjects_correctPathExistence": Test_vtkCreateListOfRememberedProjects_correctPathExistence
        Case "Test_vtkCreateListOfRememberedProjects_correctPathContent": Test_vtkCreateListOfRememberedProjects_correctPathContent
        Case "Test_vtkCreateListOfRememberedProjects_wrongPathExistence": Test_vtkCreateListOfRememberedProjects_wrongPathExistence
        Case "Test_vtkAddProjectToListOfRememberedProjects_correctPath": Test_vtkAddProjectToListOfRememberedProjects_correctPath
        Case "Test_vtkAddProjectToListOfRememberedProjects_correctPathDifferentName": Test_vtkAddProjectToListOfRememberedProjects_correctPathDifferentName
        Case "Test_vtkAddProjectToListOfRememberedProjects_correctPathSameName": Test_vtkAddProjectToListOfRememberedProjects_correctPathSameName
        Case "Test_vtkModifyProjectInList_correctPath": Test_vtkModifyProjectInList_correctPath
        Case "Test_vtkModifyProjectInList_wrongPath": Test_vtkModifyProjectInList_wrongPath
        Case "Test_vtkModifyProjectInList_correctPathNoSuchProject": Test_vtkModifyProjectInList_correctPathNoSuchProject
        Case "Test_vtkRemoveProjectFromList_correctPath": Test_vtkRemoveProjectFromList_correctPath
        Case "Test_vtkRemoveProjectFromList_wrongPath": Test_vtkRemoveProjectFromList_wrongPath
        Case "Test_vtkRemoveProjectFromList_correctPathNoSuchProject": Test_vtkRemoveProjectFromList_correctPathNoSuchProject
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

