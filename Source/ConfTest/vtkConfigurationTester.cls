VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkConfigurationTester"
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

Public Sub Test_ConfigurationCollection_Count()
    ' This test verifies the count of configurations collection property of Configuration Management
    ' just initialized from an Excel worlbook
    
    mAssert.Equals confManager.configurations.count, 1, "Number of configurations in the Test workbook"
End Sub

Public Sub Test_AddedConfigurationCollection_Count()
    ' This test verifies the count of configurations collection property of Configuration Management
    ' Added configuration to the initialized from an Excel worlbook
    
    confManager.addConfiguration ("NewConfiguration")
    
    mAssert.Equals confManager.configurations.count, 2, "Number of configurations in the Test workbook with added configuration"
End Sub

Public Sub Test_ConfigurationCollection_Parent()
    ' This test verifies the parent of each vtkConfiguration object
    ' collection just initialized from an Excel worlbook
    
    mAssert.Should confManager.configurations(1).parent Is confManager, "The parent of each vtkConfiguration must be initialized"
End Sub

Public Sub Test_AddedConfigurationCollection_Parent()
    ' This test verifies the parent of each vtkConfiguration object
    ' Added configuration to the collection initialized from an Excel worlbook
    
    confManager.addConfiguration ("NewConfiguration")
    
    mAssert.Should confManager.configurations(1).parent Is confManager, "The parent of each vtkConfiguration must be initialized"
    mAssert.Should confManager.configurations(2).parent Is confManager, "The parent of each vtkConfiguration must be initialized"
End Sub

Public Sub Test_ConfigurationCollection_Number()
    ' This test verifies the number of each vtkConfiguration object
    ' collection just initialized from an Excel worlbook
    
    mAssert.Equals confManager.configurations(1).Number, 1, "The number of each vtkConfiguration must be initialized"
End Sub

Public Sub Test_AddedConfigurationCollection_Number()
    ' This test verifies the number of each vtkConfiguration object
    ' Added configuration to the collection initialized from an Excel worlbook
    
    confManager.addConfiguration ("NewConfiguration")
    
    mAssert.Equals confManager.configurations(1).Number, 1, "The number of each vtkConfiguration must be initialized"
    mAssert.Equals confManager.configurations(2).Number, 2, "The number of each vtkConfiguration must be initialized"
End Sub

Public Sub Test_ConfigurationGetName()
    ' This test verifies the name of each vtkConfiguration object
    ' collection just initialized from an Excel worlbook
    
   On Error GoTo M_Error

    mAssert.Equals confManager.configurations(1).name, "InitializedProject_DEV", "Name of the configuration in the Test workbook"
    mAssert.Equals confManager.configurations("InitializedProject_DEV").name, "InitializedProject_DEV", "Name of the configuration in the Test workbook"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationGetPath()
    ' This test verifies the path of each vtkConfiguration object
    ' collection just initialized from an Excel worlbook
    
   On Error GoTo M_Error

    mAssert.Equals confManager.configurations(1).path, "Tests\InitializedProject_DEV.xlsm", "Path of the vtkConfiguration object by Index"
    mAssert.Equals confManager.configurations("InitializedProject_DEV").path, "Tests\InitializedProject_DEV.xlsm", "Path of the vtkConfiguration object by Name"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationSetPath()
    ' This test verifies the path setting of each vtkConfiguration object
    ' collection just initialized from an Excel worlbook

   On Error GoTo M_Error
    confManager.configurations(1).path = "NewPath"
    mAssert.Equals confManager.configurations(1).path, "NewPath", "Path of the vtkConfiguration object by index"
    mAssert.Equals confManager.configurations("InitializedProject_DEV").path, "NewPath", "Path of the vtkConfiguration object by name"
    mAssert.Equals confManager.getConfigurationPathWithNumber(1), "NewPath", "Path of the modified configuration in the configuration manager"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationModules_Initialized()
    ' This test verifies the modules property of a vtkConfiguration object
    ' collection just initialized from an Excel worlbook

   On Error GoTo M_Error
    mAssert.Equals confManager.configurations(1).modules.count, 4, "Number of modules in the Configuration"
    mAssert.Equals confManager.configurations(1).modules(1).name, "TestModule", "Name of the first module in the configuration"
    mAssert.Equals confManager.configurations(1).modules(2).name, "TestClass", "Name of the second module in the configuration"
    mAssert.Equals confManager.configurations(1).modules(3).name, "TestForm", "Name of the third module in the configuration"
    mAssert.Equals confManager.configurations(1).modules(4).name, "ThisWorkbook", "Name of the fourth module in the configuration"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationModules_NewEmpty()
    ' This test verifies the modules property of a vtkConfiguration object
    ' New collection added in an Excel worlbook
    
    confManager.addConfiguration ("NewConfiguration")

   On Error GoTo M_Error
    mAssert.Equals confManager.configurations("NewConfiguration").modules.count, 0, "Number of modules in the Configuration"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationModules_NewNotEmpty()
    ' This test verifies the modules property of a vtkConfiguration object
    ' New collection initialized in an Excel worlbook
    
    confManager.addConfiguration ("NewConfiguration")
    confManager.setModulePathWithNumber "newPath", 3, 2

   On Error GoTo M_Error
    mAssert.Equals confManager.configurations("NewConfiguration").modules.count, 1, "Number of modules in the Configuration"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationCollection_Count"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_AddedConfigurationCollection_Count"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationCollection_Parent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_AddedConfigurationCollection_Parent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationCollection_Number"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_AddedConfigurationCollection_Number"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationGetName"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationGetPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationSetPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationModules_Initialized"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationModules_NewEmpty"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationModules_NewNotEmpty"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_ConfigurationCollection_Count": Test_ConfigurationCollection_Count
        Case "Test_AddedConfigurationCollection_Count": Test_AddedConfigurationCollection_Count
        Case "Test_ConfigurationCollection_Parent": Test_ConfigurationCollection_Parent
        Case "Test_AddedConfigurationCollection_Parent": Test_AddedConfigurationCollection_Parent
        Case "Test_ConfigurationCollection_Number": Test_ConfigurationCollection_Number
        Case "Test_AddedConfigurationCollection_Number": Test_AddedConfigurationCollection_Number
        Case "Test_ConfigurationGetName": Test_ConfigurationGetName
        Case "Test_ConfigurationGetPath": Test_ConfigurationGetPath
        Case "Test_ConfigurationSetPath": Test_ConfigurationSetPath
        Case "Test_ConfigurationModules_Initialized": Test_ConfigurationModules_Initialized
        Case "Test_ConfigurationModules_NewEmpty": Test_ConfigurationModules_NewEmpty
        Case "Test_ConfigurationModules_NewNotEmpty": Test_ConfigurationModules_NewNotEmpty
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

