VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkModuleTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkConfigurationTester
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : Test the vtkConfiguration object
'             The vtkConfiguration object is a wrapper to vtkConfigurationManager
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

Private initializedWorkbook As Workbook    ' Existing Workbook opened for each test
Private Const initializedWorkbookName As String = "InitializedProject_DEV.xlsm"
Private confManager As vtkConfigurationManager   ' Configuration Manager for the new workbook
Private Const initializedProjectName As String = "InitializedProject"

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
    
    Dim source As String, destination As String
    source = vtkPathToTemplateFolder & "\" & initializedWorkbookName
    destination = vtkPathToTestFolder & "\" & initializedWorkbookName
    Set initializedWorkbook = getTestFileFromTemplate(fileName:=initializedWorkbookName, openExcel:=True)
    Set confManager = vtkConfigurationManagerForProject(initializedProjectName)
End Sub

Private Sub ITestCase_TearDown()
    vtkCloseAndKillWorkbook wb:=initializedWorkbook ' close the existing Excel project
    vtkResetConfigurationManagers             ' reset all configuration managers
End Sub

Public Sub Test_ModuleCollection_Count()
    ' This test verifies the count of modules collection property of Configuration Management
    ' just initialized from an Excel worlbook
    
    mAssert.Equals confManager.modules.Count, 4, "Number of modules in the Test workbook"
End Sub

Public Sub Test_AddedConfigurationCollection_Count()
    ' This test verifies the count of modules collection property of Configuration Management
    ' Added module to the initialized from an Excel worlbook
    
    confManager.addModule ("NewModule")
    
    mAssert.Equals confManager.modules.Count, 5, "Number of modules in the Test workbook with added module"
End Sub

Public Sub Test_ModuleCollection_Parent()
    ' This test verifies the parent of each vtkModule object
    ' collection just initialized from an Excel worlbook
    
    mAssert.Should confManager.modules(1).parent Is confManager, "The parent of each vtkModule must be initialized"
End Sub

Public Sub Test_AddedModuleCollection_Parent()
    ' This test verifies the parent of each vtkModule object
    ' Added module to the collection initialized from an Excel worlbook
    
    confManager.addModule ("NewConfiguration")
    
    mAssert.Should confManager.modules(1).parent Is confManager, "The parent of each vtkModule must be initialized"
    mAssert.Should confManager.modules(2).parent Is confManager, "The parent of each vtkModule must be initialized"
End Sub

Public Sub Test_ModuleCollection_Number()
    ' This test verifies the number of each vtkModule object
    ' collection just initialized from an Excel worlbook
    Dim i As Integer
    For i = 1 To 4
        mAssert.Equals confManager.modules(i).number, i, "The number of vtkModule " & i
    Next i
End Sub

Public Sub Test_AddedModuleCollection_Number()
    ' This test verifies the number of each vtkModule object
    ' Added module to the collection initialized from an Excel worlbook
    
    confManager.addModule ("NewConfiguration")
    Dim i As Integer
    For i = 1 To 5
        mAssert.Equals confManager.modules(i).number, i, "The number of vtkModule " & i
    Next i
End Sub

Public Sub Test_ModuleGetName()
    ' This test verifies the name of each vtkModule object
    ' collection just initialized from an Excel worlbook
    
   On Error GoTo M_Error

    mAssert.Equals confManager.modules(3).name, "TestForm", "Name of the module in the Test workbook"
    mAssert.Equals confManager.modules("TestModule").name, "TestModule", "Name of the module in the Test workbook"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & err.number & "): " & err.Description
End Sub

Public Sub Test_GetPathForConfiguration_ExistingConfiguration()
    ' This test verifies the path for configuration of each vtkModule object
    ' collection just initialized from an Excel worlbook
    
   On Error GoTo M_Error

    mAssert.Equals confManager.modules(3).getPathForConfiguration("InitializedProject_DEV"), "Tests\TestForm.frm", "Path of the module in the Test workbook"
    mAssert.Equals confManager.modules("TestModule").getPathForConfiguration("InitializedProject_DEV"), "Tests\TestModule.bas", "Path of the module in the Test workbook"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & err.number & "): " & err.Description
End Sub

Public Sub Test_GetPathForConfiguration_InexistingConfigurationRaiseError()
    ' This test verifies the path for configuration of each vtkModule object
    ' collection just initialized from an Excel worlbook
    Dim module As vtkModule
    
   On Error GoTo M_Error
    Set module = confManager.modules(3)
    mAssert.Equals module.getPathForConfiguration("NonExistant Configuration"), "", "Path of the module for non existant configuration"
    mAssert.Should False, "An error must be raised"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Equals err.number, VTK_INEXISTANT_CONFIGURATION, "Error raised"
End Sub

Public Sub Test_GetPathForConfiguration_InexistingConfigurationNullPath()
    ' This test verifies the path for configuration of each vtkModule object
    ' collection just initialized from an Excel worlbook
    
   On Error Resume Next

    mAssert.Equals confManager.modules(3).getPathForConfiguration("NonExistant Configuration"), "", "Path of the module for non existant configuration"

   On Error GoTo 0
End Sub

Public Sub Test_SetPathForConfiguration_ExistingConfiguration()
    ' This test verifies the setting of path for configuration of each vtkModule object
    ' collection just initialized from an Excel worlbook
    
   On Error GoTo M_Error

    confManager.modules(3).setPathForConfiguration "InitializedProject_DEV", "new Path"
    mAssert.Equals confManager.modules("TestForm").getPathForConfiguration("InitializedProject_DEV"), "new Path", "Path of the module in the Test workbook"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & err.number & "): " & err.Description
End Sub

Public Sub Test_SetPathForConfiguration_InexistingConfigurationRaiseError()
    ' This test verifies the setting of path for configuration of each vtkModule object
    ' collection just initialized from an Excel worlbook
    
   On Error GoTo M_Error
    confManager.modules(3).setPathForConfiguration "NonExistant Configuration", "newPath"
    mAssert.Should False, "An error must be raised"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Equals err.number, VTK_INEXISTANT_CONFIGURATION, "Error raised"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ModuleCollection_Count"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_AddedConfigurationCollection_Count"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ModuleCollection_Parent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_AddedModuleCollection_Parent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ModuleCollection_Number"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_AddedModuleCollection_Number"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ModuleGetName"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_GetPathForConfiguration_ExistingConfiguration"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_GetPathForConfiguration_InexistingConfigurationRaiseError"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_GetPathForConfiguration_InexistingConfigurationNullPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_SetPathForConfiguration_ExistingConfiguration"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_SetPathForConfiguration_InexistingConfigurationRaiseError"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_ModuleCollection_Count": Test_ModuleCollection_Count
        Case "Test_AddedConfigurationCollection_Count": Test_AddedConfigurationCollection_Count
        Case "Test_ModuleCollection_Parent": Test_ModuleCollection_Parent
        Case "Test_AddedModuleCollection_Parent": Test_AddedModuleCollection_Parent
        Case "Test_ModuleCollection_Number": Test_ModuleCollection_Number
        Case "Test_AddedModuleCollection_Number": Test_AddedModuleCollection_Number
        Case "Test_ModuleGetName": Test_ModuleGetName
        Case "Test_GetPathForConfiguration_ExistingConfiguration": Test_GetPathForConfiguration_ExistingConfiguration
        Case "Test_GetPathForConfiguration_InexistingConfigurationRaiseError": Test_GetPathForConfiguration_InexistingConfigurationRaiseError
        Case "Test_GetPathForConfiguration_InexistingConfigurationNullPath": Test_GetPathForConfiguration_InexistingConfigurationNullPath
        Case "Test_SetPathForConfiguration_ExistingConfiguration": Test_SetPathForConfiguration_ExistingConfiguration
        Case "Test_SetPathForConfiguration_InexistingConfigurationRaiseError": Test_SetPathForConfiguration_InexistingConfigurationRaiseError
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

