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
Private confManager As vtkConfigurationManager   ' Configuration Manager for the existing workbook
Private Const initializedProjectName As String = "InitializedProject"

Private Const workbookNameWithConfigurationV10 As String = "WorkBookForConfigurationsTests10.xlsm"
Private Const workbookNameWithConfigurationV11 As String = "WorkBookForConfigurationsTests11.xlsm"
Private Const workbookNameWithConfigurationV12 As String = "WorkBookForConfigurationsTests.xlsm"

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
    
    Set initializedWorkbook = getTestFileFromTemplate(fileName:=initializedWorkbookName, openExcel:=True)
    Set confManager = vtkConfigurationManagerForProject(initializedProjectName)
End Sub

Private Sub ITestCase_TearDown()
    vtkCloseAndKillWorkbook Wb:=initializedWorkbook ' close the existing Excel project
    vtkResetConfigurationManagers             ' reset all configuration managers
    VBAToolKit.resetTestFolder
End Sub

Public Sub Test_ConfigurationCollection_Count()
    ' This test verifies the count of configurations collection property of Configuration Management
    ' just initialized from an Excel workbook
    
    mAssert.Equals confManager.configurations.Count, 1, "Number of configurations in the Test workbook"
End Sub

Public Sub Test_AddedConfigurationCollection_Count()
    ' This test verifies the count of configurations collection property of Configuration Management
    ' Added configuration to the initialized from an Excel workbook
    
    confManager.addConfiguration ("NewConfiguration")
    
    mAssert.Equals confManager.configurations.Count, 2, "Number of configurations in the Test workbook with added configuration"
End Sub

Public Sub Test_ConfigurationCollection_Parent()
    ' This test verifies the parent of each vtkConfiguration object
    ' collection just initialized from an Excel workbook
    
    mAssert.Should confManager.configurations(1).parent Is confManager, "The parent of each vtkConfiguration must be initialized"
End Sub

Public Sub Test_AddedConfigurationCollection_Parent()
    ' This test verifies the parent of each vtkConfiguration object
    ' Added configuration to the collection initialized from an Excel workbook
    
    confManager.addConfiguration ("NewConfiguration")
    
    mAssert.Should confManager.configurations(1).parent Is confManager, "The parent of each vtkConfiguration must be initialized"
    mAssert.Should confManager.configurations(2).parent Is confManager, "The parent of each vtkConfiguration must be initialized"
End Sub

Public Sub Test_ConfigurationCollection_Number()
    ' This test verifies the number of each vtkConfiguration object
    ' collection just initialized from an Excel workbook
    
    mAssert.Equals confManager.configurations(1).Number, 1, "The number of each vtkConfiguration must be initialized"
End Sub

Public Sub Test_ConfigurationCollection_ID()
    ' This test verifies the ID of each vtkConfiguration object
    ' collection just initialized from an Excel workbook
    
    mAssert.Equals confManager.configurations(1).id, "c1", "Expected ID for the first vtkConfiguration"
End Sub

Public Sub Test_AddedConfigurationCollection_Number()
    ' This test verifies the number of each vtkConfiguration object
    ' Added configuration to the collection initialized from an Excel workbook
    
    confManager.addConfiguration ("NewConfiguration")
    
    mAssert.Equals confManager.configurations(1).Number, 1, "The number of each vtkConfiguration must be initialized"
    mAssert.Equals confManager.configurations(2).Number, 2, "The number of each vtkConfiguration must be initialized"
End Sub

Public Sub Test_AddedConfigurationCollection_isDEV()
    ' This test verifies the number of each vtkConfiguration object
    ' Added configuration to the collection initialized from an Excel workbook
    
    confManager.addConfiguration ("NewConfiguration")
    
    mAssert.Should confManager.configurations(1).isDEV, "The first configuration is expected to be DEV"
    mAssert.Should Not confManager.configurations(2).isDEV, "The second configuration is expected to not be DEV"
End Sub

Public Sub Test_AddedConfigurationCollection_ID()
    ' This test verifies the ID of each vtkConfiguration object
    ' Added configuration to the collection initialized from an Excel workbook
    
    confManager.addConfiguration ("NewConfiguration")
    
    mAssert.Equals confManager.configurations(1).id, "c1", "Expected ID for the first vtkConfiguration"
    mAssert.Equals confManager.configurations(2).id, "c2", "Expected ID for the second vtkConfiguration"
End Sub

Public Sub Test_ConfigurationGetName()
    ' This test verifies the name of each vtkConfiguration object
    ' collection just initialized from an Excel workbook
    
   On Error GoTo M_Error

    mAssert.Equals confManager.configurations(1).name, "InitializedProject_DEV", "Name of the configuration in the Test workbook"
    mAssert.Equals confManager.configurations("InitializedProject_DEV").name, "InitializedProject_DEV", "Name of the configuration in the Test workbook"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationGetInitializedProjectNameFromVersion12()
    ' This test verifies the projectName of a vtkConfiguration object
    ' collection just initialized from an Excel workbook
    
   On Error GoTo M_Error

    mAssert.Equals confManager.configurations(1).projectName, "InitializedProjectName", "Project Name (Title) got by number of the configuration in the Test workbook"
    mAssert.Equals confManager.configurations("InitializedProject_DEV").projectName, "InitializedProjectName", "Project Name (Title) got by name of the configuration in the Test workbook"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationGetInitializedCommentFromVersion12()
    ' This test verifies the comment of a vtkConfiguration object
    ' collection just initialized from an Excel workbook
    
   On Error GoTo M_Error

    mAssert.Equals confManager.configurations(1).comment, "Comment for project", "Comment got by number of the configuration in the Test workbook"
    mAssert.Equals confManager.configurations("InitializedProject_DEV").comment, "Comment for project", "Comment got by name of the configuration in the Test workbook"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationGetInitializedTemplateFromVersion12()
    ' This test verifies the template of a vtkConfiguration object
    ' collection just initialized from an Excel workbook
    
   On Error GoTo M_Error

    mAssert.Equals confManager.configurations(1).template, "TemplatePathOfConfiguration", "Template got by number of the configuration in the Test workbook"
    mAssert.Equals confManager.configurations("InitializedProject_DEV").template, "TemplatePathOfConfiguration", "Template got by name of the configuration in the Test workbook"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationGetDefaultProjectNameFromVersion12()
    ' This test verifies the default projectName of a vtkConfiguration object
    ' Added configuration to the initialized from an Excel workbook
    
    confManager.addConfiguration ("NewConfiguration")
    
   On Error GoTo M_Error

    mAssert.Equals confManager.configurations(2).projectName, "NewConfiguration", "Project Name (Title) got by number of the configuration in the Test workbook"
    mAssert.Equals confManager.configurations("NewConfiguration").projectName, "NewConfiguration", "Project Name (Title) got by name of the configuration in the Test workbook"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationGetDefaultCommentFromVersion12()
    ' This test verifies the default comment of a vtkConfiguration object
    ' Added configuration to the initialized from an Excel workbook
    
    confManager.addConfiguration ("NewConfiguration")
    
   On Error GoTo M_Error

    mAssert.Equals confManager.configurations(2).comment, "Project NewConfiguration", "comment got by number of the configuration in the Test workbook"
    mAssert.Equals confManager.configurations("NewConfiguration").comment, "Project NewConfiguration", "comment got by name of the configuration in the Test workbook"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationGetDefaultTemplateFromVersion12()
    ' This test verifies the default template path of a vtkConfiguration object
    ' Added configuration to the initialized from an Excel workbook
    
    confManager.addConfiguration ("NewConfiguration")
    
   On Error GoTo M_Error

    mAssert.Equals confManager.configurations(2).template, "", "template got by number of the configuration in the Test workbook"
    mAssert.Equals confManager.configurations("NewConfiguration").template, "", "template got by name of the configuration in the Test workbook"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationIsDEV_True()
    ' This test verifies the property isDEV of vtkConfiguration object
    ' collection just initialized from an Excel workbook
    
   On Error GoTo M_Error

    mAssert.Should confManager.configurations(1).isDEV, "The unique configuration of the project is a DEV one, get by number"
    mAssert.Should confManager.configurations("InitializedProject_DEV").isDEV, "The unique configuration of the project is a DEV one, get by name"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationGetPath()
    ' This test verifies the path of each vtkConfiguration object
    ' collection just initialized from an Excel workbook
    
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
    ' collection just initialized from an Excel workbook

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

Public Sub Test_ConfigurationSetProjectNameInVersion12()
    ' This test verifies the setting of projectName in a vtkConfiguration object
    ' collection just initialized from an Excel workbook
    
   On Error GoTo M_Error

    confManager.configurations(1).projectName = "NewProjectName"
    mAssert.Equals confManager.configurations(1).projectName, "NewProjectName", "Set project name (title) of the vtkConfiguration object by index"
    mAssert.Equals confManager.configurations("InitializedProject_DEV").projectName, "NewProjectName", "Set project name (title) of the vtkConfiguration object by name"
    mAssert.Equals confManager.getConfigurationProjectNameWithNumber(1), "NewProjectName", "Set project name (title) of the modified configuration in the configuration manager"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationSetProjectNameInVersion11()
    ' This test verifies the setting of projectName in a vtkConfiguration object
    ' collection just initialized from an Excel workbook
    Dim workbookV11 As Workbook, confManagerV11 As vtkConfigurationManager

   On Error GoTo M_Error

    Set workbookV11 = getTestFileFromTemplate(fileName:=workbookNameWithConfigurationV11, destinationName:="ProjectV11_DEV.xlsm", openExcel:=True)
    Set confManagerV11 = vtkConfigurationManagerForProject("ProjectV11")

    confManagerV11.configurations(1).projectName = "NewProjectName"
    mAssert.Equals confManagerV11.configurations(1).projectName, "NewProjectName", "Set project name (title) of the vtkConfiguration object by index"
    mAssert.Equals confManagerV11.configurations("ExistingProject").projectName, "NewProjectName", "Set project name (title) of the vtkConfiguration object by name"
    mAssert.Equals confManagerV11.getConfigurationProjectNameWithNumber(1), "NewProjectName", "Set project name (title) of the modified configuration in the configuration manager"

   On Error GoTo 0
    vtkCloseAndKillWorkbook Wb:=workbookV11 ' close the V11 Excel project
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
    vtkCloseAndKillWorkbook Wb:=workbookV11 ' close the V11 Excel project
End Sub

Public Sub Test_ConfigurationSetProjectNameInVersion10()
    ' This test verifies the error on setting of projectName in a v1.0 vtkConfiguration object
    ' collection just initialized from an Excel workbook
    Dim workbookV10 As Workbook, confManagerV10 As vtkConfigurationManager
    
   On Error GoTo M_Error

    Set workbookV10 = getTestFileFromTemplate(fileName:=workbookNameWithConfigurationV10, destinationName:="ProjectV10_DEV.xlsm", openExcel:=True)
    Set confManagerV10 = vtkConfigurationManagerForProject("ProjectV10")
    
   On Error GoTo 0
    confManagerV10.configurations(1).projectName = "NewProjectName"
    mAssert.Equals Err.Number, VTK_OBSOLETE_CONFIGURATION_SHEET, "expected error when modifying a v1.0 project name configuration"
   On Error GoTo M_Error
    
    mAssert.Equals confManagerV10.configurations(1).projectName, "ExistingProject", "Set project name (title) of the vtkConfiguration object by index"
    mAssert.Equals confManagerV10.configurations("ExistingProject").projectName, "ExistingProject", "Set project name (title) of the vtkConfiguration object by name"
    mAssert.Equals confManagerV10.getConfigurationProjectNameWithNumber(1), "ExistingProject", "Set project name (title) of the modified configuration in the configuration manager"

   On Error GoTo 0
    vtkCloseAndKillWorkbook Wb:=workbookV10 ' close the V10 Excel project
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationSetCommentInVersion10()
    ' This test verifies the error on setting of comment in a v1.0 vtkConfiguration object
    ' collection just initialized from an Excel workbook
    Dim workbookV10 As Workbook, confManagerV10 As vtkConfigurationManager
    
   On Error GoTo M_Error

    Set workbookV10 = getTestFileFromTemplate(fileName:=workbookNameWithConfigurationV10, destinationName:="ProjectV10_DEV.xlsm", openExcel:=True)
    Set confManagerV10 = vtkConfigurationManagerForProject("ProjectV10")
    
   On Error GoTo 0
    confManagerV10.configurations(1).comment = "New comment for project ExistingProject"
    mAssert.Equals Err.Number, VTK_OBSOLETE_CONFIGURATION_SHEET, "expected error when modifying a v1.0 project name configuration"
   On Error GoTo M_Error
    
    mAssert.Equals confManagerV10.configurations(1).comment, "Project ExistingProject", "Set comment of the vtkConfiguration object by index"
    mAssert.Equals confManagerV10.configurations("ExistingProject").comment, "Project ExistingProject", "Set comment of the vtkConfiguration object by name"
    mAssert.Equals confManagerV10.getConfigurationCommentWithNumber(1), "Project ExistingProject", "Set comment of the modified configuration in the configuration manager"

   On Error GoTo 0
    vtkCloseAndKillWorkbook Wb:=workbookV10 ' close the V10 Excel project
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationSetCommentInVersion11()
    ' This test verifies the setting of comment in a vtkConfiguration object
    ' collection just initialized from an Excel workbook
    Dim workbookV11 As Workbook, confManagerV11 As vtkConfigurationManager

   On Error GoTo M_Error

    Set workbookV11 = getTestFileFromTemplate(fileName:=workbookNameWithConfigurationV11, destinationName:="ProjectV11_DEV.xlsm", openExcel:=True)
    Set confManagerV11 = vtkConfigurationManagerForProject("ProjectV11")

    confManagerV11.configurations(1).comment = "New comment of project"
    mAssert.Equals confManagerV11.configurations(1).comment, "New comment of project", "Set comment of the vtkConfiguration object by index"
    mAssert.Equals confManagerV11.configurations("ExistingProject").comment, "New comment of project", "Set comment of the vtkConfiguration object by name"
    mAssert.Equals confManagerV11.getConfigurationCommentWithNumber(1), "New comment of project", "Set comment of the modified configuration in the configuration manager"

   On Error GoTo 0
    vtkCloseAndKillWorkbook Wb:=workbookV11 ' close the V11 Excel project
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
    vtkCloseAndKillWorkbook Wb:=workbookV11 ' close the V11 Excel project
End Sub

Public Sub Test_ConfigurationSetCommentInVersion12()
    ' This test verifies the setting of comment in a vtkConfiguration object
    ' collection just initialized from an Excel workbook
    
   On Error GoTo M_Error

    confManager.configurations(1).comment = "New comment of project"
    mAssert.Equals confManager.configurations(1).comment, "New comment of project", "Set comment of the vtkConfiguration object by index"
    mAssert.Equals confManager.configurations("InitializedProject_DEV").comment, "New comment of project", "Set comment of the vtkConfiguration object by name"
    mAssert.Equals confManager.getConfigurationCommentWithNumber(1), "New comment of project", "Set comment of the modified configuration in the configuration manager"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationSetTemplateInVersion10()
    ' This test verifies the error on setting of template in a v1.0 vtkConfiguration object
    ' collection just initialized from an Excel workbook
    Dim workbookV10 As Workbook, confManagerV10 As vtkConfigurationManager
    
   On Error GoTo M_Error

    Set workbookV10 = getTestFileFromTemplate(fileName:=workbookNameWithConfigurationV10, destinationName:="ProjectV10_DEV.xlsm", openExcel:=True)
    Set confManagerV10 = vtkConfigurationManagerForProject("ProjectV10")
    
   On Error GoTo 0
    confManagerV10.configurations(1).template = "NewTemplatePath"
    mAssert.Equals Err.Number, VTK_OBSOLETE_CONFIGURATION_SHEET, "expected error when modifying a v1.0 project name configuration"
   On Error GoTo M_Error
    
    mAssert.Equals confManagerV10.configurations(1).template, "", "Set template of the vtkConfiguration object by index"
    mAssert.Equals confManagerV10.configurations("ExistingProject").template, "", "Set template of the vtkConfiguration object by name"
    mAssert.Equals confManagerV10.getConfigurationTemplateWithNumber(1), "", "Set template of the modified configuration in the configuration manager"

   On Error GoTo 0
    vtkCloseAndKillWorkbook Wb:=workbookV10 ' close the V10 Excel project
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationSetTemplateInVersion11()
    ' This test verifies the error on setting of template in a v1.1 vtkConfiguration object
    ' collection just initialized from an Excel workbook
    Dim workbookV11 As Workbook, confManagerV11 As vtkConfigurationManager

   On Error GoTo M_Error

    Set workbookV11 = getTestFileFromTemplate(fileName:=workbookNameWithConfigurationV11, destinationName:="ProjectV11_DEV.xlsm", openExcel:=True)
    Set confManagerV11 = vtkConfigurationManagerForProject("ProjectV11")

    confManagerV11.configurations(1).template = "NewTemplatePath"
    mAssert.Equals confManagerV11.configurations(1).template, "NewTemplatePath", "Set template path of the vtkConfiguration object by index"
    mAssert.Equals confManagerV11.configurations("ExistingProject").template, "NewTemplatePath", "Set template path of the vtkConfiguration object by name"
    mAssert.Equals confManagerV11.getConfigurationTemplateWithNumber(1), "NewTemplatePath", "Set template path of the modified configuration in the configuration manager"

   On Error GoTo 0
    vtkCloseAndKillWorkbook Wb:=workbookV11 ' close the V11 Excel project
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
    vtkCloseAndKillWorkbook Wb:=workbookV11 ' close the V11 Excel project
End Sub

Public Sub Test_ConfigurationSetTemplateInVersion12()
    ' This test verifies the setting of template path in a vtkConfiguration object
    ' collection just initialized from an Excel workbook
    
   On Error GoTo M_Error

    confManager.configurations(1).template = "NewTemplatePath"
    mAssert.Equals confManager.configurations(1).template, "NewTemplatePath", "Set template of the vtkConfiguration object by index"
    mAssert.Equals confManager.configurations("InitializedProject_DEV").template, "NewTemplatePath", "Set template of the vtkConfiguration object by name"
    mAssert.Equals confManager.getConfigurationTemplateWithNumber(1), "NewTemplatePath", "Set template of the modified configuration in the configuration manager"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationGetDefaultProjectNameFromVersion10()
    ' This test verifies the default project name of a V1.0 vtkConfiguration object
    ' Added configuration to the initialized from an Excel workbook
    Dim workbookV10 As Workbook, confManagerV10 As vtkConfigurationManager
    
   On Error GoTo M_Error

    Set workbookV10 = getTestFileFromTemplate(fileName:=workbookNameWithConfigurationV10, destinationName:="ProjectV10_DEV.xlsm", openExcel:=True)
    Set confManagerV10 = vtkConfigurationManagerForProject("ProjectV10")
    
    mAssert.Equals confManagerV10.configurations(1).projectName, "ExistingProject", "project name (Title) got by number of the configuration in the Test workbook"
    mAssert.Equals confManagerV10.configurations("ExistingProject").projectName, "ExistingProject", "project name (Title) got by name of the configuration in the Test workbook"

   On Error GoTo 0
    vtkCloseAndKillWorkbook Wb:=workbookV10 ' close the V10 Excel project
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
    vtkCloseAndKillWorkbook Wb:=workbookV10 ' close the V10 Excel project
End Sub

Public Sub Test_ConfigurationGetDefaultCommentFromVersion10()
    ' This test verifies the default comment of a V1.0 vtkConfiguration object
    ' Added configuration to the initialized from an Excel workbook
    Dim workbookV10 As Workbook, confManagerV10 As vtkConfigurationManager
    
   On Error GoTo M_Error

    Set workbookV10 = getTestFileFromTemplate(fileName:=workbookNameWithConfigurationV10, destinationName:="ProjectV10_DEV.xlsm", openExcel:=True)
    Set confManagerV10 = vtkConfigurationManagerForProject("ProjectV10")
    
    mAssert.Equals confManagerV10.configurations(1).comment, "Project ExistingProject", "comment got by number of the configuration in the Test workbook"
    mAssert.Equals confManagerV10.configurations("ExistingProject").comment, "Project ExistingProject", "comment got by name of the configuration in the Test workbook"

   On Error GoTo 0
    vtkCloseAndKillWorkbook Wb:=workbookV10 ' close the V10 Excel project
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
    vtkCloseAndKillWorkbook Wb:=workbookV10 ' close the V10 Excel project
End Sub

Public Sub Test_ConfigurationGetDefaultTemplateFromVersion10()
    ' This test verifies the default Excel template of a V1.0 vtkConfiguration object
    ' Added configuration to the initialized from an Excel workbook
    Dim workbookV10 As Workbook, confManagerV10 As vtkConfigurationManager
    
   On Error GoTo M_Error

    Set workbookV10 = getTestFileFromTemplate(fileName:=workbookNameWithConfigurationV10, destinationName:="ProjectV10_DEV.xlsm", openExcel:=True)
    Set confManagerV10 = vtkConfigurationManagerForProject("ProjectV10")
    
    mAssert.Equals confManagerV10.configurations(1).template, "", "template path got by number of the configuration in the Test workbook"
    mAssert.Equals confManagerV10.configurations("ExistingProject").template, "", "template path got by name of the configuration in the Test workbook"

   On Error GoTo 0
    vtkCloseAndKillWorkbook Wb:=workbookV10 ' close the V10 Excel project
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
    vtkCloseAndKillWorkbook Wb:=workbookV10 ' close the V10 Excel project
End Sub

Public Sub Test_ConfigurationModules_Initialized()
    ' This test verifies the modules property of a vtkConfiguration object
    ' collection just initialized from an Excel workbook

   On Error GoTo M_Error
    mAssert.Equals confManager.configurations(1).modules.Count, 4, "Number of modules in the Configuration"
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
    ' New collection added in an Excel workbook
    
    confManager.addConfiguration ("NewConfiguration")

   On Error GoTo M_Error
    mAssert.Equals confManager.configurations("NewConfiguration").modules.Count, 0, "Number of modules in the Configuration"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationModules_NewNotEmpty()
    ' This test verifies the modules property of a vtkConfiguration object
    ' New collection initialized in an Excel workbook
    
    confManager.addConfiguration ("NewConfiguration")
    confManager.setModulePathWithNumber "newPath", 3, 2

   On Error GoTo M_Error
    mAssert.Equals confManager.configurations("NewConfiguration").modules.Count, 1, "Number of modules in the Configuration"

   On Error GoTo 0
   Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationReferences_NoReferenceSheet()
    ' This test verifies the references property of a vtkConfiguration object
    ' no references are initialized when creating a reference sheet

    Dim c As Collection

   On Error GoTo M_Error
    Set c = confManager.configurations(1).references
    
    mAssert.Equals c.Count, 0, "Count of references of the first configuration"
   On Error GoTo 0
    Exit Sub

M_Error:
    mAssert.Should False, "Unexpected Error (" & Err.Number & "): " & Err.Description
End Sub

Public Sub Test_ConfigurationReferences_WithReferenceSheet()
    ' This test verifies the references property of a vtkConfiguration object
    ' collection just initialized from an Excel workbook

    Dim refNames(), i As Integer, c1 As Collection, c2 As Collection, r As vtkReference
    Dim Wb As Workbook, cf As vtkConfigurationManager
    refNames = Array("VBA", "Excel", "stdole", "Office", "MSForms", "Scripting", "VBIDE", "Shell32", "MSXML2")
       
    Set Wb = getTestFileFromTemplate(fileName:=workbookNameWithConfigurationV12, destinationName:="ProjectV12_DEV.xlsm", openExcel:=True)
    Set cf = vtkConfigurationManagerForProject("ProjectV12")

   On Error Resume Next
    ' Test for the main configuration
    Set c1 = cf.configurations(1).references
    ' Rearrange the collection by name
    Set c2 = New Collection
    For Each r In c1
        c2.Add r, r.name
    Next
    mAssert.Equals c2.Count, UBound(refNames) - LBound(refNames) + 1, "Count of all references of a new workbook"
    ' il faut boucler sur le tableau et rechercher dans la collection (si pas trouvé = erreur)
    For i = LBound(refNames) To UBound(refNames)
        Set r = c2(refNames(i))
        mAssert.Equals Err.Number, 0, "Error when getting " & refNames(i) & " reference"
    Next i
    
    ' Test for the Development configuration
    Set c1 = cf.configurations(2).references
    ' Rearrange the collection by name
    Set c2 = New Collection
    For Each r In c1
        c2.Add r, r.name
    Next
    mAssert.Equals c2.Count, UBound(refNames) - LBound(refNames) + 1, "Count of all references of a new workbook"
    ' il faut boucler sur le tableau et rechercher dans la collection (si pas trouvé = erreur)
    For i = LBound(refNames) To UBound(refNames)
        Set r = c2(refNames(i))
        mAssert.Equals Err.Number, 0, "Error when getting " & refNames(i) & " reference"
    Next i
   On Error GoTo 0
    vtkCloseAndKillWorkbook Wb:=Wb ' close the existing Excel project

End Sub

Public Sub Test_ConfigurationReferences_NewWorkbook()
    ' This test verifies the references property of a vtkConfiguration object
    ' references are initialized when creating a new workbook

    Dim newWorkbook As Workbook, newConfManager As vtkConfigurationManager
    
    Set newWorkbook = vtkCreateExcelWorkbookForTestWithProjectName("NewProject")    ' create an Excel project for tests
    Set newConfManager = vtkConfigurationManagerForProject(newWorkbook.VBProject.name)

   On Error GoTo M_Error
    mAssert.Equals newConfManager.configurations(1).references.Count, 10, "Count of references of the main configuration"
    mAssert.Equals newConfManager.configurations(2).references.Count, 11, "Count of references of the development configuration"
   
   On Error GoTo 0
    vtkCloseAndKillWorkbook Wb:=newWorkbook     ' close the new Excel project
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
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationCollection_ID"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_AddedConfigurationCollection_Number"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_AddedConfigurationCollection_isDEV"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_AddedConfigurationCollection_ID"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationGetName"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationGetInitializedProjectNameFromVersion12"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationGetInitializedCommentFromVersion12"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationGetInitializedTemplateFromVersion12"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationGetDefaultProjectNameFromVersion12"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationGetDefaultCommentFromVersion12"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationGetDefaultTemplateFromVersion12"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationIsDEV_True"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationGetPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationSetPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationSetProjectNameInVersion12"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationSetProjectNameInVersion11"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationSetProjectNameInVersion10"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationSetCommentInVersion10"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationSetCommentInVersion11"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationSetCommentInVersion12"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationSetTemplateInVersion10"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationSetTemplateInVersion12"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationGetDefaultProjectNameFromVersion10"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationGetDefaultCommentFromVersion10"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationGetDefaultTemplateFromVersion10"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationModules_Initialized"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationModules_NewEmpty"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationModules_NewNotEmpty"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationReferences_NoReferenceSheet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationReferences_WithReferenceSheet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ConfigurationReferences_NewWorkbook"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_ConfigurationCollection_Count": Test_ConfigurationCollection_Count
        Case "Test_AddedConfigurationCollection_Count": Test_AddedConfigurationCollection_Count
        Case "Test_ConfigurationCollection_Parent": Test_ConfigurationCollection_Parent
        Case "Test_AddedConfigurationCollection_Parent": Test_AddedConfigurationCollection_Parent
        Case "Test_ConfigurationCollection_Number": Test_ConfigurationCollection_Number
        Case "Test_ConfigurationCollection_ID": Test_ConfigurationCollection_ID
        Case "Test_AddedConfigurationCollection_Number": Test_AddedConfigurationCollection_Number
        Case "Test_AddedConfigurationCollection_isDEV": Test_AddedConfigurationCollection_isDEV
        Case "Test_AddedConfigurationCollection_ID": Test_AddedConfigurationCollection_ID
        Case "Test_ConfigurationGetName": Test_ConfigurationGetName
        Case "Test_ConfigurationGetInitializedProjectNameFromVersion12": Test_ConfigurationGetInitializedProjectNameFromVersion12
        Case "Test_ConfigurationGetInitializedCommentFromVersion12": Test_ConfigurationGetInitializedCommentFromVersion12
        Case "Test_ConfigurationGetInitializedTemplateFromVersion12": Test_ConfigurationGetInitializedTemplateFromVersion12
        Case "Test_ConfigurationGetDefaultProjectNameFromVersion12": Test_ConfigurationGetDefaultProjectNameFromVersion12
        Case "Test_ConfigurationGetDefaultCommentFromVersion12": Test_ConfigurationGetDefaultCommentFromVersion12
        Case "Test_ConfigurationGetDefaultTemplateFromVersion12": Test_ConfigurationGetDefaultTemplateFromVersion12
        Case "Test_ConfigurationIsDEV_True": Test_ConfigurationIsDEV_True
        Case "Test_ConfigurationGetPath": Test_ConfigurationGetPath
        Case "Test_ConfigurationSetPath": Test_ConfigurationSetPath
        Case "Test_ConfigurationSetProjectNameInVersion12": Test_ConfigurationSetProjectNameInVersion12
        Case "Test_ConfigurationSetProjectNameInVersion11": Test_ConfigurationSetProjectNameInVersion11
        Case "Test_ConfigurationSetProjectNameInVersion10": Test_ConfigurationSetProjectNameInVersion10
        Case "Test_ConfigurationSetCommentInVersion10": Test_ConfigurationSetCommentInVersion10
        Case "Test_ConfigurationSetCommentInVersion11": Test_ConfigurationSetCommentInVersion11
        Case "Test_ConfigurationSetCommentInVersion12": Test_ConfigurationSetCommentInVersion12
        Case "Test_ConfigurationSetTemplateInVersion10": Test_ConfigurationSetTemplateInVersion10
        Case "Test_ConfigurationSetTemplateInVersion12": Test_ConfigurationSetTemplateInVersion12
        Case "Test_ConfigurationGetDefaultProjectNameFromVersion10": Test_ConfigurationGetDefaultProjectNameFromVersion10
        Case "Test_ConfigurationGetDefaultCommentFromVersion10": Test_ConfigurationGetDefaultCommentFromVersion10
        Case "Test_ConfigurationGetDefaultTemplateFromVersion10": Test_ConfigurationGetDefaultTemplateFromVersion10
        Case "Test_ConfigurationModules_Initialized": Test_ConfigurationModules_Initialized
        Case "Test_ConfigurationModules_NewEmpty": Test_ConfigurationModules_NewEmpty
        Case "Test_ConfigurationModules_NewNotEmpty": Test_ConfigurationModules_NewNotEmpty
        Case "Test_ConfigurationReferences_NoReferenceSheet": Test_ConfigurationReferences_NoReferenceSheet
        Case "Test_ConfigurationReferences_WithReferenceSheet": Test_ConfigurationReferences_WithReferenceSheet
        Case "Test_ConfigurationReferences_NewWorkbook": Test_ConfigurationReferences_NewWorkbook
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

