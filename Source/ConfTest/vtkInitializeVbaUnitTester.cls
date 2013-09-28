VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkInitializeVbaUnitTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkInitializeVbaUnitTester
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Test the vtkInitializeVBAUnit function
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

Private newWorkBook As Workbook
Private Const projectName As String = "NewProject"
Private newConfManager As vtkConfigurationManager

Private mManager As TestCaseManager
Private mAssert As IAssert

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
    Set newWorkBook = vtkCreateExcelWorkbookForTestWithProjectName(projectName)    ' create an Excel project for tests
    Set newConfManager = vtkConfigurationManagerForProject(projectName)
End Sub

Private Sub ITestCase_TearDown()
    vtkCloseAndKillWorkbook Wb:=newWorkBook     ' close the new Excel project
    vtkResetConfigurationManagers             ' reset all configuration managers
End Sub

Public Sub TestAddVBAUnitInNewProject()
'       Verify the count of added VBAUnit modules
    mAssert.Equals newConfManager.moduleCount, 0, "There is no module in the new project"
    vtkInitializeVbaUnitNamesAndPathes project:=projectName
    mAssert.Equals newConfManager.moduleCount, 18, "Number of VBAUnit modules included in new project"
End Sub

Public Sub TestPathOfModuleVBAUnitInNewProject()
'       Verify the path of VbaUnitMain
    vtkInitializeVbaUnitNamesAndPathes project:=projectName
    Dim nm As Integer
    nm = newConfManager.getModuleNumber("VbaUnitMain")
    mAssert.Should nm > 0, "The module VBA Unit must be added in the new project"
    Dim nc As Integer
    nc = newConfManager.getConfigurationNumber(vtkProjectForName(projectName).projectDEVName)
    mAssert.Should nc > 0, "The New Project configuration must be added in the new project"
    mAssert.Equals newConfManager.getModulePathWithNumber(nm, nc), "Source\VbaUnit\VbaUnitMain.bas", "Relative path of VBAUnitMain module"
End Sub

Public Sub TestPathOfModuleAssertInNewProject()
'       Verify the path of VbaUnitMain
    vtkInitializeVbaUnitNamesAndPathes project:=projectName
    Dim nm As Integer
    nm = newConfManager.getModuleNumber("Assert")
    mAssert.Should nm > 0, "The module VBA Unit must be added in the new project"
    Dim nc As Integer
    nc = newConfManager.getConfigurationNumber(vtkProjectForName(projectName).projectDEVName)
    mAssert.Should nc > 0, "The New Project configuration must be added in the new project"
    mAssert.Equals newConfManager.getModulePathWithNumber(nm, nc), "Source\VbaUnit\Assert.cls", "Relative path of Assert module"
End Sub

Public Sub TestPathOfModuleIAssertInNewProject()
'       Verify the path of VbaUnitMain
    vtkInitializeVbaUnitNamesAndPathes project:=projectName
    Dim nm As Integer
    nm = newConfManager.getModuleNumber("IAssert")
    mAssert.Should nm > 0, "The module VBA Unit must be added in the new project"
    Dim nc As Integer
    nc = newConfManager.getConfigurationNumber(vtkProjectForName(projectName).projectDEVName)
    mAssert.Should nc > 0, "The New Project configuration must be added in the new project"
    mAssert.Equals newConfManager.getModulePathWithNumber(nm, nc), "Source\VbaUnit\IAssert.cls", "Relative path of IAssert module"
End Sub

' Ajouter des tests pour Module absent, configuration absente, ...

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestAddVBAUnitInNewProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestPathOfModuleVBAUnitInNewProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestPathOfModuleAssertInNewProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestPathOfModuleIAssertInNewProject"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestAddVBAUnitInNewProject": TestAddVBAUnitInNewProject
        Case "TestPathOfModuleVBAUnitInNewProject": TestPathOfModuleVBAUnitInNewProject
        Case "TestPathOfModuleAssertInNewProject": TestPathOfModuleAssertInNewProject
        Case "TestPathOfModuleIAssertInNewProject": TestPathOfModuleIAssertInNewProject
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


