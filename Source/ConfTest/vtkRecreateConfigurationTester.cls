VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkRecreateConfigurationTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkRecreateConfigurationTester
' Author    : Lucas Vitorino
' Purpose   : Test the vtkRecreateConfiguration function
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

Private Wb As Workbook
Private recreatedWb As Workbook
Private testProject As New vtkProject
Private normalConf As vtkConfiguration
Private emptyConf As vtkConfiguration
Private wrongPathConf As vtkConfiguration
Private cm As vtkConfigurationManager

Private Const testProjectName = "TestProjectForRecreateConf"
Private Const normalConfName = "TestProject_normalConf"
Private Const emptyConfName = "TestProject_emptyConf"
Private Const wrongPathConfName = "TestProject_wrongPathConf"


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
    
    Set testProject = vtkProjectForName(testProjectName)
    
    Set Wb = getTestFileFromTemplate(testProject.workbookDEVName, openExcel:=True)

    Set cm = vtkConfigurationManagerForProject(testProjectName)
    Set normalConf = cm.configurations(normalConfName)
    Set emptyConf = cm.configurations(emptyConfName)
    Set wrongPathConf = cm.configurations(wrongPathConfName)
    
End Sub

Private Sub ITestCase_TearDown()
    On Error Resume Next
    vtkCloseAndKillWorkbook Wb
    vtkCloseAndKillWorkbook recreatedWb
    VBAToolKit.resetTestFolder
    vtkResetProjects
    vtkResetConfigurationManagers
End Sub

' TESTS TO BE WRITTEN
'
' normal configuration, normal path
'   - workbook is created in the good path
'   - all modules of the configuration are present
'   - references are activated
'   - name of the vbproject is correct
'   - filetype is consistent with the extension
'
' configuration particular cases
'   - configuration is empty
'   - different kind of modules (noraml, class, form)
'
' particular cases
'   - workbook with the same name is already open
'   - workbook with the same name already exists in the same path
'   - source files have not been yet exported
'

Public Sub Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath()
' workbook is created in the good path

    On Error GoTo Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_Error
    
    vtkRecreateConfiguration testProject.projectName, normalConf.name

    mAssert.Should fso.FileExists(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & normalConf.path), _
                   "The configuration should have been created in the designated path."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_Error:
    Err.Source = "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub

Public Sub Test_vtkRecreateConfiguration_allModulesArePresent()
' all modules of the configuration are present

    On Error GoTo Test_vtkRecreateConfiguration_allModulesArePresent_Error
    
    vtkExportConfiguration Wb.VBProject, testProject.projectName, normalConf.name

    vtkRecreateConfiguration testProject.projectName, normalConf.name

    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & normalConf.path)

    Dim mo As vtkModule
    For Each mo In normalConf.modules
        mAssert.Should (Not recreatedWb.VBProject.VBComponents(mo.name) Is Nothing), _
                        "Module " & mo.name & " missing in the recreated workbook."
    Next
    
    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_allModulesArePresent_Error:
    Err.Source = "Test_vtkRecreateConfiguration_allModulesArePresent of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_allModulesArePresent"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath": Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath
        Case "Test_vtkRecreateConfiguration_allModulesArePresent": Test_vtkRecreateConfiguration_allModulesArePresent
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

