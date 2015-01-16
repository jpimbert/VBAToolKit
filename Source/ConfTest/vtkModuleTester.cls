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
    
    Dim Source As String, destination As String
    Source = vtkPathToTemplateFolder & "\" & initializedWorkbookName
    destination = vtkPathToTestFolder & "\" & initializedWorkbookName
    Set initializedWorkbook = getTestFileFromTemplate(fileName:=initializedWorkbookName, openExcel:=True)
    Set confManager = vtkConfigurationManagerForProject(initializedProjectName)
End Sub

Private Sub ITestCase_TearDown()
    vtkCloseAndKillWorkbook Wb:=initializedWorkbook ' close the existing Excel project
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
        mAssert.Equals confManager.modules(i).Number, i, "The number of vtkModule " & i
    Next i
End Sub

Public Sub Test_ModuleCollection_ID()
    ' This test verifies the ID of each vtkModule object
    ' collection just initialized from an Excel worlbook
    Dim i As Integer
    For i = 1 To 4
        mAssert.Equals confManager.modules(i).id, "m" & i, "Expected ID for module " & i
    Next i
End Sub

Public Sub Test_AddedModuleCollection_Number()
    ' This test verifies the Number of each vtkModule object
    ' Added module to the collection initialized from an Excel worlbook
    
    confManager.addModule ("NewConfiguration")
    Dim i As Integer
    For i = 1 To 5
        mAssert.Equals confManager.modules(i).Number, i, "The number of vtkModule " & i
    Next i
End Sub

Public Sub Test_AddedModuleCollection_ID()
    ' This test verifies the ID of each vtkModule object
    ' Added module to the collection initialized from an Excel worlbook
    
    confManager.addModule ("NewConfiguration")
    Dim i As Integer
    For i = 1 To 5
        mAssert.Equals confManager.modules(i).id, "m" & i, "Expected ID of vtkModule " & i
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
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
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
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
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
    mAssert.Equals Err.Number, VTK_INEXISTANT_CONFIGURATION, "Error raised"
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
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
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
    mAssert.Equals Err.Number, VTK_INEXISTANT_CONFIGURATION, "Error raised"
End Sub

Public Sub Test_Configuration_Attached()
    ' This test verifies the getting of configuration for a module attached to a configuration
    ' collection just initialized from an Excel worlbook
    
   On Error GoTo M_Error
    mAssert.Equals confManager.configurations(1).modules(2).configuration, "InitializedProject_DEV", "Name of configuration"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_Configuration_NotAttachedErrorRaised()
    ' This test verifies the getting of configuration for a module not attached to a configuration
    ' collection just initialized from an Excel worlbook
    
   On Error GoTo M_Error
    mAssert.Equals confManager.modules(2).configuration, "", "Name of configuration"
    mAssert.Should False, "An error must be raised"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Equals Err.Number, VTK_MODULE_NOTATTACHED, "Error raised"
End Sub

Public Sub Test_Configuration_NotAttachedNoError()
    ' This test verifies the getting of configuration for a module not attached to a configuration
    ' collection just initialized from an Excel worlbook
    
   On Error Resume Next
    mAssert.Equals confManager.modules(2).configuration, "", "Name of configuration"
    mAssert.Equals Err.Number, VTK_MODULE_NOTATTACHED, "Error raised"
End Sub

Public Sub Test_Path_Attached()
    ' This test verifies the getting of the path for a module attached to a configuration
    ' collection just initialized from an Excel worlbook
    
   On Error GoTo M_Error
    mAssert.Equals confManager.configurations(1).modules(2).path, "Tests\TestClass.cls", "Path of module"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_Path_NotAttachedErrorRaised()
    ' This test verifies the getting of the path for a module not attached to a configuration
    ' collection just initialized from an Excel worlbook
    
   On Error GoTo M_Error
    mAssert.Equals confManager.modules(2).path, "", "Path of module"
    mAssert.Should False, "An error must be raised"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Equals Err.Number, VTK_MODULE_NOTATTACHED, "Error raised"
End Sub

Public Sub Test_Path_NotAttachedNoError()
    ' This test verifies the getting of the path for a module not attached to a configuration
    ' collection just initialized from an Excel worlbook
    
   On Error Resume Next
    mAssert.Equals confManager.modules(2).path, "", "Path of module"
    mAssert.Equals Err.Number, VTK_MODULE_NOTATTACHED, "Error raised"
End Sub

Public Sub Test_VBAModule_Attached()
    ' This test verifies the getting of the VBAModule for a module attached to a configuration
    ' collection just initialized from an Excel worlbook
    
   On Error GoTo M_Error
    mAssert.Equals confManager.configurations(1).modules(4).VBAModule.name, "ThisWorkbook", "Name of VBComponent"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_VBAModule_NotAttachedErrorRaised()
    ' This test verifies the getting of the VBAModule for a module not attached to a configuration
    ' collection just initialized from an Excel worlbook
    
   On Error GoTo M_Error
    mAssert.Equals confManager.modules(4).VBAModule.name, "", "Name of VBComponent"
    mAssert.Should False, "An error must be raised"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Equals Err.Number, VTK_MODULE_NOTATTACHED, "Error raised"
End Sub

Public Sub Test_VBAModule_NotAttachedNoError()
    ' This test verifies the getting of the VBAModule for a module not attached to a configuration
    ' collection just initialized from an Excel worlbook
    
   On Error Resume Next
    mAssert.Equals confManager.modules(4).VBAModule.name, "", "Name of VBComponent"
    mAssert.Equals Err.Number, VTK_MODULE_NOTATTACHED, "Error raised"
End Sub

Public Sub Test_VBAModule_NotOpen()
    ' This test verifies the getting of the VBAModule for a module attached to a configuration
    ' New configuration, Excel file not open
    confManager.addConfiguration configuration:="NewConfiguration", path:="path/toto.xlsm"
    confManager.modules(3).setPathForConfiguration "NewConfiguration", "Toto/Module.cls"
    
   On Error GoTo M_Error
    mAssert.Equals confManager.configurations("NewConfiguration").modules(1).VBAModule.name, "", "Name of VBComponent"
    mAssert.Should False, "An error must be raised"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Equals Err.Number, VTK_WORKBOOK_NOTOPEN, "Error raised"
End Sub

Public Sub Test_VBAModule_NoModule()
    ' This test verifies the getting of the VBAModule for a module attached to a configuration
    ' New module not in project, Excel file open
    confManager.addModule module:="NewModule"
    confManager.modules("NewModule").setPathForConfiguration "InitializedProject_DEV", "Toto/Module.cls"

   On Error GoTo M_Error
    mAssert.Should confManager.configurations(1).modules("NewModule").VBAModule Is Nothing, "VBComponent must ne Nothing"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ModuleCollection_Count"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_AddedConfigurationCollection_Count"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ModuleCollection_Parent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_AddedModuleCollection_Parent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ModuleCollection_Number"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ModuleCollection_ID"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_AddedModuleCollection_Number"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_AddedModuleCollection_ID"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ModuleGetName"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_GetPathForConfiguration_ExistingConfiguration"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_GetPathForConfiguration_InexistingConfigurationRaiseError"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_GetPathForConfiguration_InexistingConfigurationNullPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_SetPathForConfiguration_ExistingConfiguration"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_SetPathForConfiguration_InexistingConfigurationRaiseError"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_Configuration_Attached"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_Configuration_NotAttachedErrorRaised"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_Configuration_NotAttachedNoError"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_Path_Attached"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_Path_NotAttachedErrorRaised"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_Path_NotAttachedNoError"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_VBAModule_Attached"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_VBAModule_NotAttachedErrorRaised"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_VBAModule_NotAttachedNoError"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_VBAModule_NotOpen"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_VBAModule_NoModule"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_ModuleCollection_Count": Test_ModuleCollection_Count
        Case "Test_AddedConfigurationCollection_Count": Test_AddedConfigurationCollection_Count
        Case "Test_ModuleCollection_Parent": Test_ModuleCollection_Parent
        Case "Test_AddedModuleCollection_Parent": Test_AddedModuleCollection_Parent
        Case "Test_ModuleCollection_Number": Test_ModuleCollection_Number
        Case "Test_ModuleCollection_ID": Test_ModuleCollection_ID
        Case "Test_AddedModuleCollection_Number": Test_AddedModuleCollection_Number
        Case "Test_AddedModuleCollection_ID": Test_AddedModuleCollection_ID
        Case "Test_ModuleGetName": Test_ModuleGetName
        Case "Test_GetPathForConfiguration_ExistingConfiguration": Test_GetPathForConfiguration_ExistingConfiguration
        Case "Test_GetPathForConfiguration_InexistingConfigurationRaiseError": Test_GetPathForConfiguration_InexistingConfigurationRaiseError
        Case "Test_GetPathForConfiguration_InexistingConfigurationNullPath": Test_GetPathForConfiguration_InexistingConfigurationNullPath
        Case "Test_SetPathForConfiguration_ExistingConfiguration": Test_SetPathForConfiguration_ExistingConfiguration
        Case "Test_SetPathForConfiguration_InexistingConfigurationRaiseError": Test_SetPathForConfiguration_InexistingConfigurationRaiseError
        Case "Test_Configuration_Attached": Test_Configuration_Attached
        Case "Test_Configuration_NotAttachedErrorRaised": Test_Configuration_NotAttachedErrorRaised
        Case "Test_Configuration_NotAttachedNoError": Test_Configuration_NotAttachedNoError
        Case "Test_Path_Attached": Test_Path_Attached
        Case "Test_Path_NotAttachedErrorRaised": Test_Path_NotAttachedErrorRaised
        Case "Test_Path_NotAttachedNoError": Test_Path_NotAttachedNoError
        Case "Test_VBAModule_Attached": Test_VBAModule_Attached
        Case "Test_VBAModule_NotAttachedErrorRaised": Test_VBAModule_NotAttachedErrorRaised
        Case "Test_VBAModule_NotAttachedNoError": Test_VBAModule_NotAttachedNoError
        Case "Test_VBAModule_NotOpen": Test_VBAModule_NotOpen
        Case "Test_VBAModule_NoModule": Test_VBAModule_NoModule
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

