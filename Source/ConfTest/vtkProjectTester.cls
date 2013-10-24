VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkProjectTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkProjectTester
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Test the vtkProject class
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
Private project As vtkProject      ' Default vtkProject for test

Private fso As New FileSystemObject

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
    Set mAssert = Assert
    projectName = "TestProject"
    Set project = vtkProjectForName(projectName)
End Sub

Private Sub ITestCase_TearDown()
    vtkResetProjects
    Set project = Nothing
End Sub

Public Sub TestNullProject()
'       Verify that all properties are null when the object is not initialized
    Dim p As New vtkProject
    mAssert.Equals p.projectName, "", "The project must be null"
    mAssert.Equals p.projectDEVName, "", "The project development must be null"
    mAssert.Equals p.projectDEVStandardRelativePath, "", "The path for development must be null"
    mAssert.Equals p.projectStandardRelativePath, "", "The path for delivery must be null"
    mAssert.Equals p.workbookDEVName, "", "The workbook for development must be null"
    mAssert.Equals p.workbookName, "", "The workbook for delivery must be null"
End Sub

Public Sub TestProjectNameForDevelopment()
'       Verify the project name for development
    mAssert.Equals project.projectDEVName, "TestProject_DEV", "The project development is suffixed by _DEV"
End Sub

Public Sub TestProjectDevelopmentRelativePath()
'       Verify the project path for development
    mAssert.Equals project.projectDEVStandardRelativePath, "Project\" & "TestProject_DEV.xlsm", "The project development standard path"
End Sub

Public Sub TestProjectRelativePath()
'       Verify the project path
    mAssert.Equals project.projectStandardRelativePath, "Delivery\" & "TestProject.xlsm", "The project development standard path"
End Sub

Public Sub TestXMLConfigurationStandardRelativePath()
'       Verify the project path
    mAssert.Equals project.XMLConfigurationStandardRelativePath, "Project\TestProject.xml", "The project configuration standard path"
End Sub

Public Sub TestProjectDevelopmentWorkbookName()
'       Verify the project path for development
    mAssert.Equals project.workbookDEVName, "TestProject_DEV.xlsm", "The project development standard name"
End Sub

Public Sub TestProjectWorkbookName()
'       Verify the project path
    mAssert.Equals project.workbookName, "TestProject.xlsm", "The project delivery standard name"
End Sub

Public Sub TestXMLConfigurationName()
'       Verify the project path
    mAssert.Equals project.XMLConfigurationName, "TestProject.xml", "The project confighuration standard name"
End Sub


Public Sub Test_RootFolderPath_BeforeSet()
' verify the root folder path before the "before root" folder path has been set
    mAssert.Equals project.projectRootFolderPath, fso.BuildPath(VBAToolKit.vtkTestPath, "TestProject"), _
                   "The absolute path of the project's folder before it has been set."
End Sub

Public Sub Test_RootFolderPath_AfterSet()
' verify the root folder path after the "before root" folder path has been set
    project.projectBeforeRootFolderPath = fso.BuildPath(VBAToolKit.vtkTestPath, "folder1\")
    mAssert.Equals project.projectRootFolderPath, fso.BuildPath(VBAToolKit.vtkTestPath, "\folder1\TestProject"), _
                   "The absolute path of the project's folder after it has been set."
End Sub


Public Sub Test_beforeRootFolderPath_BeforeSet()
' verify the before root folder path after it has been set
    mAssert.Equals project.projectBeforeRootFolderPath, VBAToolKit.vtkTestPath, _
                   "The absolute path of the project's folder before it has been set."
End Sub

Public Sub Test_beforeRootFolderPath_afterSet()
' verify the before root folder path after it has been set
    project.projectBeforeRootFolderPath = fso.BuildPath(VBAToolKit.vtkTestPath, "folder1")
    mAssert.Equals project.projectBeforeRootFolderPath, fso.BuildPath(VBAToolKit.vtkTestPath, "\folder1"), _
                   "The absolute path of the project's folder after it has been set."
End Sub


Public Sub Test_xmlRelativeFolderPath_BeforeSet()
' verify the xml relative path before it has been set
    mAssert.Equals project.xmlRelativeFolderPath, "", "The xml relative folder path before it has been set."
End Sub

Public Sub Test_xmlRelativeFolderPath_afterSet()
' verify the xml relative path after it has been set
    project.xmlRelativeFolderPath = "\folder1"
    mAssert.Equals project.xmlRelativeFolderPath, "\folder1", "The xml relative folder path after it has been set."
End Sub

Public Sub Test_xmlSheetFullPath_beforeRootAndRelativePathAreSet()
' verify the full path of the xml sheet before the project root path and the relative path have been set
    mAssert.Equals project.xmlSheetFullPath, fso.BuildPath(VBAToolKit.vtkTestPath, "\TestProject\TestProject.xml")
End Sub

Public Sub Test_xmlSheetFullPath_rootPathSetNotRelativePath()
' verify the full path of the xml sheet before the relative path has been set
    project.projectBeforeRootFolderPath = fso.BuildPath(VBAToolKit.vtkTestPath, "\folder1")
    mAssert.Equals project.xmlSheetFullPath, fso.BuildPath(VBAToolKit.vtkTestPath, "\folder1\TestProject\TestProject.xml")
End Sub

Public Sub Test_xmlSheetFullPath_relativePathSetNotRootPath()
' verify the full path of the xml sheet before the project root path has been set
    project.xmlRelativeFolderPath = "\folder1"
    mAssert.Equals project.xmlSheetFullPath, fso.BuildPath(VBAToolKit.vtkTestPath, "\TestProject\folder1\TestProject.xml")
End Sub

Public Sub Test_xmlSheetFullPath_rootAndRelativePathSet()
' verify the full path of the xml sheet after the project root path and the relative path have been set
    project.projectBeforeRootFolderPath = fso.BuildPath(VBAToolKit.vtkTestPath, "\beforeRoot")
    project.xmlRelativeFolderPath = "\folder1"
    mAssert.Equals project.xmlSheetFullPath, fso.BuildPath(VBAToolKit.vtkTestPath, "beforeRoot\TestProject\folder1\TestProject.xml")
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestNullProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectNameForDevelopment"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectDevelopmentRelativePath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectRelativePath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestXMLConfigurationStandardRelativePath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectDevelopmentWorkbookName"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectWorkbookName"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestXMLConfigurationName"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_RootFolderPath_BeforeSet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_RootFolderPath_AfterSet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_beforeRootFolderPath_BeforeSet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_beforeRootFolderPath_afterSet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_xmlRelativeFolderPath_BeforeSet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_xmlRelativeFolderPath_afterSet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_xmlSheetFullPath_beforeRootAndRelativePathAreSet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_xmlSheetFullPath_rootPathSetNotRelativePath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_xmlSheetFullPath_relativePathSetNotRootPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_xmlSheetFullPath_rootAndRelativePathSet"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestNullProject": TestNullProject
        Case "TestProjectNameForDevelopment": TestProjectNameForDevelopment
        Case "TestProjectDevelopmentRelativePath": TestProjectDevelopmentRelativePath
        Case "TestProjectRelativePath": TestProjectRelativePath
        Case "TestXMLConfigurationStandardRelativePath": TestXMLConfigurationStandardRelativePath
        Case "TestProjectDevelopmentWorkbookName": TestProjectDevelopmentWorkbookName
        Case "TestProjectWorkbookName": TestProjectWorkbookName
        Case "TestXMLConfigurationName": TestXMLConfigurationName
        Case "Test_RootFolderPath_BeforeSet": Test_RootFolderPath_BeforeSet
        Case "Test_RootFolderPath_AfterSet": Test_RootFolderPath_AfterSet
        Case "Test_beforeRootFolderPath_BeforeSet": Test_beforeRootFolderPath_BeforeSet
        Case "Test_beforeRootFolderPath_afterSet": Test_beforeRootFolderPath_afterSet
        Case "Test_xmlRelativeFolderPath_BeforeSet": Test_xmlRelativeFolderPath_BeforeSet
        Case "Test_xmlRelativeFolderPath_afterSet": Test_xmlRelativeFolderPath_afterSet
        Case "Test_xmlSheetFullPath_beforeRootAndRelativePathAreSet": Test_xmlSheetFullPath_beforeRootAndRelativePathAreSet
        Case "Test_xmlSheetFullPath_rootPathSetNotRelativePath": Test_xmlSheetFullPath_rootPathSetNotRelativePath
        Case "Test_xmlSheetFullPath_relativePathSetNotRootPath": Test_xmlSheetFullPath_relativePathSetNotRootPath
        Case "Test_xmlSheetFullPath_rootAndRelativePathSet": Test_xmlSheetFullPath_rootAndRelativePathSet
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub




