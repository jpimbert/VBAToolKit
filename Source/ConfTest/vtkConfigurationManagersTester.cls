VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkConfigurationManagersTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkConfigurationManagersTester
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Test the vtkConfigurationManagers module capacities
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

Private wbForTest As Workbook       ' default WorkBook for test
Private workbookName As String      ' default workbook Name for test
Private Const projectName As String = "TestProject"     ' default project Name for test
Private Const uninitializedWorkbookName As String = "UninitializedProject.xlsm"     ' unititialized workbook Name for test
Private Const existingWorkbookName As String = "WorkBookForConfigurationsTests.xlsm"     ' unititialized workbook Name for test

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
    Set wbForTest = vtkCreateExcelWorkbookForTestWithProjectName(projectName)   ' create the default Excel project for tests
    workbookName = wbForTest.name
    FileCopy Source:=vtkPathToTemplateFolder & "\" & uninitializedWorkbookName, destination:=vtkPathToTestFolder & "\" & uninitializedWorkbookName
    FileCopy Source:=vtkPathToTemplateFolder & "\" & existingWorkbookName, destination:=vtkPathToTestFolder & "\" & existingWorkbookName
End Sub

Private Sub ITestCase_TearDown()
    vtkCloseAndKillWorkbook Wb:=wbForTest
    On Error Resume Next
    Kill PathName:=vtkPathToTestFolder & "\" & uninitializedWorkbookName
    Kill PathName:=vtkPathToTestFolder & "\" & existingWorkbookName
    On Error GoTo 0
    vtkResetConfigurationManagers               ' reset all configuration managers
End Sub

' TESTS TO WRITE
'
'



Public Sub TestUtilityFunctionCreateExcelProject()
'       Verify that the utility function createExcelProjectNamed correctly create a new Excel file with proper project name
    mAssert.Equals wbForTest.name, workbookName, "Expected name for Excel project"
End Sub

Public Sub TestConfigurationManagerIsNotNull()
'       Verify that the configuration Manager is not null
    mAssert.Should Not vtkConfigurationManagerForProject(projectName) Is Nothing, "A configuartion manager must not be null"
End Sub

Public Sub TestConfigurationManagersAreDifferentForTwoProjects()
'       Verify that the configuration Managers for two project are different
    Dim secondWb As Workbook
    Set secondWb = vtkCreateExcelWorkbookForTestWithProjectName("SecondTest")    ' create a second Excel project for test
    mAssert.Should Not vtkConfigurationManagerForProject(projectName) Is vtkConfigurationManagerForProject("SecondTest"), "Two configuration managers must not be equal"
    vtkCloseAndKillWorkbook Wb:=secondWb
End Sub

Public Sub TestConfigurationManagerIsInvariant()
'       Verify that the configuration Manager for a project is always the same
    Dim cm As vtkConfigurationManager
    Set cm = vtkConfigurationManagerForProject(projectName)
    mAssert.Should vtkConfigurationManagerForProject(projectName) Is cm, "The configuration manager for a project must be always the same"
End Sub

Public Sub TestInoperantProjectConfigurationSheetInitializationForExistingProject()
    '   Verifies that the configuration sheet initialization does nothing on existing project
    
    ' Active an existing VBAToolKit project
    vtkCloseAndKillWorkbook Wb:=wbForTest
    Set wbForTest = Workbooks.Open(fileName:=vtkPathToTestFolder & "\" & existingWorkbookName)
    
    ' Try to Initialize workbook
    vtkInitializeConfigurationForActiveWorkBook
    
    ' Verify the project name
    mAssert.Equals wbForTest.VBProject.name, "ExistingProject", "The project name must not be changed"
End Sub

Public Sub TestProjectNameAfterProjectConfigurationSheetInitialization()
    '   Verify the capacity to build a configuration sheet based on existing modules
    '   The test uses a template project containing different types of empty modules
    '   This specific test verifies the project name modification of the active workbook
    
    ' Open uninitialized workbook
    vtkCloseAndKillWorkbook Wb:=wbForTest
    Set wbForTest = Workbooks.Open(fileName:=vtkPathToTestFolder & "\" & uninitializedWorkbookName)
    
    ' Initialize workbook
    vtkInitializeConfigurationForActiveWorkBook
    
    ' Verify the project name
    mAssert.Equals wbForTest.VBProject.name, "TestProject_DEV", "The suffix _DEV must have been appended to the project name"
End Sub

Public Sub TestWorkbookNameAfterProjectConfigurationSheetInitialization()
    '   Verify the capacity to build a configuration sheet based on existing modules
    '   The test uses a template project containing different types of empty modules
    '   This specific test verifies the workbook name modification of the active workbook
    
    ' Open uninitialized workbook
    vtkCloseAndKillWorkbook Wb:=wbForTest
    Set wbForTest = Workbooks.Open(fileName:=vtkPathToTestFolder & "\" & uninitializedWorkbookName)
    
    ' Initialize workbook
    vtkInitializeConfigurationForActiveWorkBook
    
    ' Verify the workbook name
    mAssert.Equals wbForTest.name, "TestProject_DEV.xlsm", "The workbook name must be the project name"
End Sub

Public Sub TestConfigurationAfterProjectConfigurationSheetInitialization()
    '   Verify the capacity to build a configuration sheet based on existing modules
    '   The test uses a template project containing different types of empty modules
    '   This specific test verifies the project name modification of the active workbook
    
    ' Open uninitialized workbook
    Dim pName As String
    vtkCloseAndKillWorkbook Wb:=wbForTest
    Set wbForTest = Workbooks.Open(fileName:=vtkPathToTestFolder & "\" & uninitializedWorkbookName)
    pName = wbForTest.VBProject.name
    
    ' Initialize workbook
    vtkInitializeConfigurationForActiveWorkBook
    
    ' Verify the configuration
    Dim cm As vtkConfigurationManager, p As vtkProject
    Set p = vtkProjectForName(projectName:=pName)
    Set cm = vtkConfigurationManagerForProject(projectName:=pName)
    Dim nc_dev As Integer, nc_prod As Integer
    nc_dev = cm.getConfigurationNumber(configuration:=p.projectDEVName)
    nc_prod = cm.getConfigurationNumber(configuration:=p.projectName)
    
    mAssert.Equals cm.moduleCount, 5, "There are 5 modules to export, the Document modules are not exported"
    
    mAssert.Equals cm.getModulePathWithNumber(numModule:=cm.getModuleNumber("IAssert"), numConfiguration:=nc_dev), "Source\VbaUnit\IAssert.cls", "Path for IAssert module for Test configuration"
    mAssert.Equals cm.getModulePathWithNumber(numModule:=cm.getModuleNumber("UserForm"), numConfiguration:=nc_dev), "Source\ConfProd\UserForm.frm", "Path for UserForm module for Test configuration"
    mAssert.Equals cm.getModulePathWithNumber(numModule:=cm.getModuleNumber("Module"), numConfiguration:=nc_dev), "Source\ConfProd\Module.bas", "Path for Module module for Test configuration"
    mAssert.Equals cm.getModulePathWithNumber(numModule:=cm.getModuleNumber("Class"), numConfiguration:=nc_dev), "Source\ConfProd\Class.cls", "Path for Class module for Test configuration"
    mAssert.Equals cm.getModulePathWithNumber(numModule:=cm.getModuleNumber("ClassTester"), numConfiguration:=nc_dev), "Source\ConfTest\ClassTester.cls", "Path for ClassTester module for Test configuration"
    
    mAssert.Equals cm.getModulePathWithNumber(numModule:=cm.getModuleNumber("IAssert"), numConfiguration:=nc_prod), "", "Path for IAssert module for Prod configuration"
    mAssert.Equals cm.getModulePathWithNumber(numModule:=cm.getModuleNumber("UserForm"), numConfiguration:=nc_prod), "Source\ConfProd\UserForm.frm", "Path for UserForm module for Prod configuration"
    mAssert.Equals cm.getModulePathWithNumber(numModule:=cm.getModuleNumber("Module"), numConfiguration:=nc_prod), "Source\ConfProd\Module.bas", "Path for Module module for Prod configuration"
    mAssert.Equals cm.getModulePathWithNumber(numModule:=cm.getModuleNumber("Class"), numConfiguration:=nc_prod), "Source\ConfProd\Class.cls", "Path for Class module for Prod configuration"
    mAssert.Equals cm.getModulePathWithNumber(numModule:=cm.getModuleNumber("ClassTester"), numConfiguration:=nc_prod), "", "Path for ClassTester module for Prod configuration"
    
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestUtilityFunctionCreateExcelProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationManagerIsNotNull"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationManagersAreDifferentForTwoProjects"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationManagerIsInvariant"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestInoperantProjectConfigurationSheetInitializationForExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectNameAfterProjectConfigurationSheetInitialization"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestWorkbookNameAfterProjectConfigurationSheetInitialization"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationAfterProjectConfigurationSheetInitialization"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestUtilityFunctionCreateExcelProject": TestUtilityFunctionCreateExcelProject
        Case "TestConfigurationManagerIsNotNull": TestConfigurationManagerIsNotNull
        Case "TestConfigurationManagersAreDifferentForTwoProjects": TestConfigurationManagersAreDifferentForTwoProjects
        Case "TestConfigurationManagerIsInvariant": TestConfigurationManagerIsInvariant
        Case "TestInoperantProjectConfigurationSheetInitializationForExistingProject": TestInoperantProjectConfigurationSheetInitializationForExistingProject
        Case "TestProjectNameAfterProjectConfigurationSheetInitialization": TestProjectNameAfterProjectConfigurationSheetInitialization
        Case "TestWorkbookNameAfterProjectConfigurationSheetInitialization": TestWorkbookNameAfterProjectConfigurationSheetInitialization
        Case "TestConfigurationAfterProjectConfigurationSheetInitialization": TestConfigurationAfterProjectConfigurationSheetInitialization
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


