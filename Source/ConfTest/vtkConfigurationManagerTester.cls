VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkConfigurationManagerTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkConfigurationManagerTester
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Test the vtkConfigurationManager class
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

Private newWorkBook As Workbook         ' New Workbook created for each test
Private newWorkbookName As String
Private newConfManager As vtkConfigurationManager   ' Configuration Manager for the new workbook
Private newProjectName As String

Private existingWorkbook As Workbook    ' Existing Workbook opened for each test
Private Const existingWorkbookNameForTest As String = "WorkBookForConfigurationsTests.xlsm"
Private existingConfManager As vtkConfigurationManager   ' Configuration Manager for the new workbook
Private Const existingProjectName As String = "ExistingProject"
Private existingWorkbookName As String

Private Sub SaveThenReOpenNewWorkbook()
' Utility Function for saving then re-open the new workbook and re-initialize the configuration managers
    Dim savedNewWorkbookPath As String
    savedNewWorkbookPath = newWorkBook.FullName
'    If newWorkbook.path = "" Then
'        newWorkbook.Close SaveChanges:=True, Filename:=savedNewWorkbookPath
'       Else
'        newWorkbook.Close SaveChanges:=True
'    End If
    newWorkBook.Close saveChanges:=True
    vtkResetConfigurationManagers           ' reset all configuration managers
    Set newWorkBook = Workbooks.Open(fileName:=savedNewWorkbookPath)
    Set newConfManager = vtkConfigurationManagerForProject("NewProject")
End Sub

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
    Set newWorkBook = vtkCreateExcelWorkbookForTestWithProjectName("NewProject")    ' create an Excel project for tests
    newWorkbookName = newWorkBook.name
    newProjectName = newWorkBook.VBProject.name
    Set newConfManager = vtkConfigurationManagerForProject(newProjectName)
    
    Set existingWorkbook = getTestFileFromTemplate(fileName:=existingWorkbookNameForTest, destinationName:=existingProjectName & "_DEV.xlsm", openExcel:=True)
    existingWorkbookName = existingWorkbook.name
    Set existingConfManager = vtkConfigurationManagerForProject(existingProjectName)
End Sub

Private Sub ITestCase_TearDown()
    vtkCloseAndKillWorkbook Wb:=newWorkBook     ' close the new Excel project
    vtkCloseAndKillWorkbook Wb:=existingWorkbook ' close the existing Excel project
    vtkResetConfigurationManagers             ' reset all configuration managers
End Sub

Public Sub Test_PropertyName_DefaultGet()
    '   Verify that the Property Name is the Default property for vtkConfigurationManager
    '   - In fact there is no need to run the test, just to compile it
    mAssert.Equals newConfManager, "NewProject", "The name property must be the default one for vtkConfigurationManager"
End Sub

Public Sub Test_PropertyName_DefaultLet()
    '   Verify that the Property Name is the Default property for vtkConfigurationManager
    '   - In fact there is no need to run the test, just to compile it
    mAssert.Equals existingConfManager, "ExistingProject", "The name property of existingConf before modification"
    existingConfManager = "NewProject"
    mAssert.Equals existingConfManager, "NewProject", "The name property of existingConf after modification"
End Sub

Public Sub TestConfigurationSheetExistsInExistingWorkbook()
    '   Verify that the configuration sheet presence is detected in existing workbook
    '   using a fresh configuration Manager (with no default sheet initialized)
    Dim cm As New vtkConfigurationManager
    mAssert.Should cm.isConfigurationInitializedForWorkbook(ExcelName:=existingWorkbookName), "The Configuration sheet must exist in existing workbook"
End Sub

Public Sub TestConfigurationSheetDoesntExistInNewWorkbook()
    '   Verify that the configuration sheet missing is detected in new workbook
    '   using a fresh configuration Manager (with no default sheet initialized)
    Dim cm As New vtkConfigurationManager, Wb As Workbook, wbFullName As String
    Set Wb = vtkCreateExcelWorkbookForTestWithProjectName("NewWorkbook")    ' create a fresh new Excel workbook
    wbFullName = Wb.FullName
    mAssert.Should Not cm.isConfigurationInitializedForWorkbook(ExcelName:=Wb.name), "The Configuration sheet must not exist in new workbook"
    Wb.Close saveChanges:=False
    Kill PathName:=wbFullName
End Sub

Public Sub TestConfigurationSheetCreationForNewProject()
'       Verify that a Configuration Sheet is created in a new project
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = newWorkBook.Sheets("vtkConfigurations")
    mAssert.Equals Err, 0, "A configuration manager must create a Configuration sheet"
    On Error GoTo 0
    mAssert.Should newWorkBook.Sheets("vtkConfigurations") Is newConfManager.configurationSheet, "The configurationSheet property of the conf manager must be equal to the configuration sheet of the workbook"
End Sub

Public Sub TestConfigurationSheetRetrievalForExistingProject()
'       Verify that a Configuration Sheet is retreived in an existing project
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = existingWorkbook.Sheets("vtkConfigurations")
    mAssert.Equals Err, 0, "A configuration manager must be accessible in an existing project"
    On Error GoTo 0
    mAssert.Should existingWorkbook.Sheets("vtkConfigurations") Is existingConfManager.configurationSheet, "The configurationSheet property of the conf manager must be equal to the configuration sheet of the workbook"
End Sub

Public Sub TestConfigurationSheetFormatForNewProjet()
'       Verify the newly created configuration sheet of a new project
    Dim ws As Worksheet
    Set ws = newConfManager.configurationSheet
    mAssert.Equals ws.Range("A1"), "vtkConfigurations v1.0", "Expected identification of the configuration sheet"
    mAssert.Equals ws.Range("A2"), "Module Name", "Expected Title for Modules column"
    mAssert.Equals ws.Range("B1"), newProjectName, "Expected Title for main project column"
    mAssert.Equals ws.Range("B2"), "Delivery\" & newProjectName & ".xlsm", "Expected related Path for new main workbook"
    mAssert.Equals ws.Range("C1"), newProjectName & "_DEV", "Expected Title for development project column"
    mAssert.Equals ws.Range("C2"), "Project\" & newWorkbookName, "Expected related Path for new development workbook"
End Sub

Public Sub TestConfigurationSheetFormatForExistingProjet()
'       Verify the retrieved configuration sheet from an existing project
    Dim ws As Worksheet
    Set ws = existingConfManager.configurationSheet
    mAssert.Equals ws.Range("A1"), "vtkConfigurations v1.0", "Expected identification of the configuration sheet"
    mAssert.Equals ws.Range("A2"), "Module Name", "Expected Title for Modules column"
    mAssert.Equals ws.Range("B1"), existingProjectName, "Expected Title for main project column"
    mAssert.Equals ws.Range("B2"), "Delivery\ExistingProject.xlsm", "Expected related Path for new main workbook"
    mAssert.Equals ws.Range("C1"), existingProjectName & "_DEV", "Expected Title for development project column"
    mAssert.Equals ws.Range("C2"), "Project\ExistingProject_DEV.xlsm", "Expected related Path for new development workbook"
End Sub

Public Sub TestGetConfigurationsFromNewProject()
'       Verify the list of the configurations of a new project
    mAssert.Equals newConfManager.configurationCount, 2, "There must be two configurations in a new project"
    mAssert.Equals newConfManager.configuration(0), "", "Inexistant configuration number 0"
    mAssert.Equals newConfManager.configuration(1), newProjectName, "Name of the first configuration"
    mAssert.Equals newConfManager.configuration(2), newProjectName & "_DEV", "Name of the second configuration"
    mAssert.Equals newConfManager.configuration(3), "", "Inexistant configuration number 3"
    mAssert.Equals newConfManager.configuration(-23), "", "Inexistant configuration number -23"
    mAssert.Equals newConfManager.configuration(150), "", "Inexistant configuration number 150"
End Sub

Public Sub TestGetConfigurationsFromExistingProject()
'       Verify the list of the configurations of an existing project
    mAssert.Equals existingConfManager.configurationCount, 2, "There must be two configurations in the existing template project"
    mAssert.Equals existingConfManager.configuration(0), "", "Inexistant configuration number 0"
    mAssert.Equals existingConfManager.configuration(1), existingProjectName, "Name of the first configuration"
    mAssert.Equals existingConfManager.configuration(2), existingProjectName & "_DEV", "Name of the second configuration"
    mAssert.Equals existingConfManager.configuration(3), "", "Inexistant configuration number 3"
    mAssert.Equals existingConfManager.configuration(-23), "", "Inexistant configuration number -23"
    mAssert.Equals existingConfManager.configuration(150), "", "Inexistant configuration number 150"
End Sub

Public Sub Test_AddConfigurationInExistingProject_Name()
'       Verify the add of configuration in an existing project
'       Verify the number and name of the added configuration
    Dim n As Integer
    n = existingConfManager.addConfiguration("NewConfiguration", "ConfigurationPath")

    mAssert.Equals existingConfManager.configurationCount, 3, "There must be two configurations in the existing template project"
    mAssert.Equals existingConfManager.configuration(0), "", "Inexistant configuration number 0"
    mAssert.Equals existingConfManager.configuration(1), existingProjectName, "Name of the first configuration"
    mAssert.Equals existingConfManager.configuration(2), existingProjectName & "_DEV", "Name of the second configuration"
    mAssert.Equals existingConfManager.configuration(3), "NewConfiguration", "Name of the new configuration"
    mAssert.Equals existingConfManager.getConfigurationPathWithNumber(3), "ConfigurationPath", "Path of new configuration given by number"
    mAssert.Equals existingConfManager.configuration(4), "", "Inexistant configuration number 4"
    mAssert.Equals existingConfManager.configuration(-23), "", "Inexistant configuration number -23"
    mAssert.Equals existingConfManager.configuration(150), "", "Inexistant configuration number 150"
End Sub

Public Sub Test_AddConfigurationInExistingProject_Cells()
'       Verify the add of configuration in an existing project
'       Verify the cells of the configuration sheet
    Dim ws As Worksheet
    Dim n As Integer
    n = existingConfManager.addConfiguration("NewConfiguration", "ConfigurationPath")
    
    Set ws = existingConfManager.configurationSheet
    mAssert.Equals ws.Range("A1"), "vtkConfigurations v1.0", "Expected identification of the configuration sheet"
    mAssert.Equals ws.Range("A2"), "Module Name", "Expected Title for Modules column"
    mAssert.Equals ws.Range("B1"), existingProjectName, "Expected Title for main project column"
    mAssert.Equals ws.Range("B2"), "Delivery\ExistingProject.xlsm", "Expected related Path for new main workbook"
    mAssert.Equals ws.Range("C1"), existingProjectName & "_DEV", "Expected Title for development project column"
    mAssert.Equals ws.Range("C2"), "Project\ExistingProject_DEV.xlsm", "Expected related Path for new development workbook"
    mAssert.Equals ws.Range("D1"), "NewConfiguration", "Expected Title for new configuration column"
    mAssert.Equals ws.Range("D2"), "ConfigurationPath", "Expected related Path for new configuration"
End Sub

Public Sub Test_AddConfigurationInExistingProject_NullPathes()
'       Verify the add of configuration in an existing project
'       Verify That the module pathes are initialized to null
    Dim n As Integer, i As Integer
    n = existingConfManager.addConfiguration("NewConfiguration", "ConfigurationPath")
    
    For i = 1 To existingConfManager.configurationCount
        mAssert.Equals existingConfManager.getModulePathWithNumber(i, n), "", "Path of module " & i & " must be null"
    Next
End Sub

Public Sub TestGetConfigurationPathWithNumberFromExistingProject()
'       Verify the capability to get the configuration path by number
    mAssert.Equals existingConfManager.getConfigurationPathWithNumber(0), "", "Inexistant configuration number 0"
    mAssert.Equals existingConfManager.getConfigurationPathWithNumber(1), "Delivery\ExistingProject.xlsm", "Path of first configuration given by number"
    mAssert.Equals existingConfManager.getConfigurationPathWithNumber(2), "Project\ExistingProject_DEV.xlsm", "Path of second configuration given by number"
    mAssert.Equals existingConfManager.getConfigurationPathWithNumber(3), "", "Inexistant configuration number 3"
End Sub

Public Sub TestSetConfigurationPathWithNumberToNewProject()
'       Verify the capability to set and retrieve the configuration path by number
    ' set new pathes
    newConfManager.setConfigurationPathWithNumber n:=0, path:="Path0"
    newConfManager.setConfigurationPathWithNumber n:=1, path:="Path1"
    newConfManager.setConfigurationPathWithNumber n:=2, path:="Path2"
    newConfManager.setConfigurationPathWithNumber n:=3, path:="Path3"
    ' verify pathes
    mAssert.Equals newConfManager.getConfigurationPathWithNumber(0), "", "Inexistant configuration number 0"
    mAssert.Equals newConfManager.getConfigurationPathWithNumber(1), "Path1", "Path of first configuration given by number"
    mAssert.Equals newConfManager.getConfigurationPathWithNumber(2), "Path2", "Path of second configuration given by number"
    mAssert.Equals newConfManager.getConfigurationPathWithNumber(3), "", "Inexistant configuration number 3"
End Sub

Public Sub TestSetConfigurationPathWithNumberToSavedProject()
'       Verify the capability to set and retrieve the configuration path by number
    ' set new pathes
    newConfManager.setConfigurationPathWithNumber n:=0, path:="Path0"
    newConfManager.setConfigurationPathWithNumber n:=1, path:="Path1"
    newConfManager.setConfigurationPathWithNumber n:=2, path:="Path2"
    newConfManager.setConfigurationPathWithNumber n:=3, path:="Path3"
    ' save and re-open file
    SaveThenReOpenNewWorkbook
    ' verify pathes
    mAssert.Equals newConfManager.getConfigurationPathWithNumber(0), "", "Inexistant configuration number 0"
    mAssert.Equals newConfManager.getConfigurationPathWithNumber(1), "Path1", "Path of first configuration given by number"
    mAssert.Equals newConfManager.getConfigurationPathWithNumber(2), "Path2", "Path of second configuration given by number"
    mAssert.Equals newConfManager.getConfigurationPathWithNumber(3), "", "Inexistant configuration number 3"
End Sub

Public Sub TestGetConfigurationNumbersFromNewProject()
'       Verify the capability to get the number of a configuration
    mAssert.Equals newConfManager.configurationCount, 2, "There must be two configurations in a new project"
    mAssert.Equals newConfManager.getConfigurationNumber(newProjectName), 1, "Number of the main configuration"
    mAssert.Equals newConfManager.getConfigurationNumber(newProjectName & "_DEV"), 2, "Number of the Development configuration"
    mAssert.Equals newConfManager.getConfigurationNumber("InexistantConfiguration"), 0, "Inexistant configuration"
End Sub

Public Sub TestGetConfigurationPathFromExistingProject()
'       Verify the capability to get a configutaion path given the configuration name
    mAssert.Equals existingConfManager.getConfigurationPath(existingProjectName), "Delivery\ExistingProject.xlsm", "Path of the main configuration"
    mAssert.Equals existingConfManager.getConfigurationPath(existingProjectName & "_DEV"), "Project\ExistingProject_DEV.xlsm", "Path of the Development configuration"
    mAssert.Equals existingConfManager.getConfigurationPath("InexistantConfiguration"), "", "Inexistant configuration"
End Sub

Public Sub TestSetConfigurationPathToSavedProject()
'       Verify the capability to set and retrieve the configuration path by configuration name
    ' set new pathes
    newConfManager.setConfigurationPath configuration:="InexistantConfiguration", path:="Path0"
    newConfManager.setConfigurationPath configuration:=newProjectName, path:="Path1"
    newConfManager.setConfigurationPath configuration:=newProjectName & "_DEV", path:="Path2"
    ' save and re-open file
    SaveThenReOpenNewWorkbook
    ' verify pathes
    mAssert.Equals newConfManager.getConfigurationPath("InexistantConfiguration"), "", "Inexistant configuration"
    mAssert.Equals newConfManager.getConfigurationPath(newProjectName), "Path1", "Path of first configuration given by name"
    mAssert.Equals newConfManager.getConfigurationPath(newProjectName & "_DEV"), "Path2", "Path of second configuration given by name"
End Sub

Public Sub TestGetModulesFromExistingProject()
'       Verify the capability to retrieve the list of Modules from an existing project
    mAssert.Equals existingConfManager.moduleCount, 5, "There must be five configurations in the existing project"
    mAssert.Equals existingConfManager.module(0), "", "Inexistant module number 0"
    mAssert.Equals existingConfManager.module(1), "Module1", "Name of the first module"
    mAssert.Equals existingConfManager.module(2), "Module2", "Name of the second module"
    mAssert.Equals existingConfManager.module(3), "Module3", "Name of the third module"
    mAssert.Equals existingConfManager.module(4), "Module4", "Name of the fourth module"
    mAssert.Equals existingConfManager.module(5), "Module5", "Name of the fifth module"
    mAssert.Equals existingConfManager.module(6), "", "Inexistant module number 6"
    mAssert.Equals existingConfManager.module(-23), "", "Inexistant module number -23"
    mAssert.Equals existingConfManager.module(150), "", "Inexistant module number 150"
End Sub

Public Sub TestGetModulesFromNewProject()
'       Verify the capability to retrieve the list of Modules from an existing project
    mAssert.Equals newConfManager.moduleCount, 0, "There must be no modules in a new project"
    mAssert.Equals newConfManager.module(0), "", "Inexistant module number 0"
    mAssert.Equals newConfManager.module(1), "", "Inexistant module number 1"
    mAssert.Equals newConfManager.module(6), "", "Inexistant module number 6"
    mAssert.Equals newConfManager.module(-23), "", "Inexistant module number -23"
    mAssert.Equals newConfManager.module(150), "", "Inexistant module number 150"
End Sub

Public Sub TestGetModuleNumbersFromExistingProject()
'       Verify the capability to get the number of a configuration
    mAssert.Equals existingConfManager.getModuleNumber("Module0"), 0, "Inexistant module"
    mAssert.Equals existingConfManager.getModuleNumber("Module1"), 1, "First Module"
    mAssert.Equals existingConfManager.getModuleNumber("Module2"), 2, "Second Module"
    mAssert.Equals existingConfManager.getModuleNumber("Module3"), 3, "Third module"
    mAssert.Equals existingConfManager.getModuleNumber("Module4"), 4, "Fourth module"
    mAssert.Equals existingConfManager.getModuleNumber("Module5"), 5, "Fifth module"
    mAssert.Equals existingConfManager.getModuleNumber("InexistantModule"), 0, "Inexistant module"
End Sub

Public Sub TestAddNonExistantModuleToSavedProject()
'       Verify the capability to add a new module, non existant, and retrieve it
    ' set new modules
    mAssert.Equals newConfManager.addModule(module:="NewModule1"), 1, "Number of the first module added"
    mAssert.Equals newConfManager.addModule(module:="NewModule2"), 2, "Number of the second module added"
    ' save and re-open file
    SaveThenReOpenNewWorkbook
    ' verify modules
    mAssert.Equals newConfManager.moduleCount, 2, "There must be two new modules in the saved project"
    mAssert.Equals newConfManager.module(0), "", "Inexistant module number 0"
    mAssert.Equals newConfManager.module(1), "NewModule1", "New module number 1"
    mAssert.Equals newConfManager.module(2), "NewModule2", "New module number 2"
    mAssert.Equals newConfManager.module(3), "", "Inexistant module number 3"
End Sub

Public Sub TestAddExistantModuleToExistingProject()
'       Verify the capability to not add an existing module in an existing project
    Dim n As Integer
    ' set new modules
    mAssert.Equals existingConfManager.addModule(module:="Module1"), -1, "Number of the first existing module"
    mAssert.Equals existingConfManager.addModule(module:="Module5"), -5, "Number of the fifth existing module"
    mAssert.Equals existingConfManager.moduleCount, 5, "There must be five modules, no change, in the existing project"
End Sub

Public Sub TestAddModuleWithExistingStringToExistingProject()
'       Verify the capability to add a new module whose name is included in existing module in an existing project
    Dim n As Integer
    ' set new modules
    mAssert.Equals existingConfManager.addModule(module:="Module"), 6, "Number for the new module"
    mAssert.Equals existingConfManager.moduleCount, 6, "There must be six modules, one more module, in the existing project"
End Sub

Public Sub TestGetModulePathWithNumberFromExistingProject()
'       Verify the capability to get the module path by number
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=0, numConfiguration:=2), "", "Inexistant module path number 0,2"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=3, numConfiguration:=3), "", "Inexistant module path number 3,3"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=1, numConfiguration:=1), "Path1Module1", "Module path number 1,1"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=1, numConfiguration:=2), "", "Module path number 1,2"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=2, numConfiguration:=1), "", "Module path number 2,1"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=2, numConfiguration:=2), "Path2Module2", "Module path number 2,2"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=3, numConfiguration:=1), "", "Module path number 3,1"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=3, numConfiguration:=2), "", "Module path number 3,2"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=4, numConfiguration:=1), "Path1Module4", "Module path number 4,1"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=4, numConfiguration:=2), "Path2Module4", "Module path number 4,2"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=5, numConfiguration:=1), "", "Module path number 5,1"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=5, numConfiguration:=2), "Path2Module5", "Module path number 5,2"
End Sub

Public Sub TestGetModulePathWithNumberForNewModule()
'       Verify the default module path define at module adding
    mAssert.Equals newConfManager.addModule(module:="NewModule1"), 1, "Number of the first module added"
    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=1, numConfiguration:=1), "", "Inexistant module path number 1,1"
    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=1, numConfiguration:=2), "", "Inexistant module path number 1,2"
End Sub

Public Sub TestAddModulePathToSavedProject()
'       Verify the capability to add a new module, non existant, and retrieve it
    ' set new modules
    mAssert.Equals newConfManager.addModule(module:="NewModule1"), 1, "Number of the first module added"
    mAssert.Equals newConfManager.addModule(module:="NewModule2"), 2, "Number of the second module added"
    mAssert.Equals newConfManager.addModule(module:="NewModule3"), 3, "Number of the third module added"
    mAssert.Equals newConfManager.addModule(module:="NewModule4"), 4, "Number of the fourth module added"
    ' save and re-open file
    SaveThenReOpenNewWorkbook
    ' set new pathes
    newConfManager.setModulePathWithNumber path:="Path1Module1", numModule:=1, numConfiguration:=1
    newConfManager.setModulePathWithNumber path:="Path2Module1", numModule:=1, numConfiguration:=2
    newConfManager.setModulePathWithNumber path:="Path1Module2", numModule:=2, numConfiguration:=1
    newConfManager.setModulePathWithNumber path:="Path2Module3", numModule:=3, numConfiguration:=2
    ' save and re-open file
    SaveThenReOpenNewWorkbook
    ' verify module pathes
    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=1, numConfiguration:=1), "Path1Module1", "Module path number 1,1"
    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=1, numConfiguration:=2), "Path2Module1", "Module path number 1,2"
    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=2, numConfiguration:=1), "Path1Module2", "Module path number 2,1"
    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=2, numConfiguration:=2), "", "Module path number 2,2"
    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=3, numConfiguration:=1), "", "Module path number 3,1"
    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=3, numConfiguration:=2), "Path2Module3", "Module path number 3,2"
    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=4, numConfiguration:=1), "", "Module path number 4,1"
    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=4, numConfiguration:=2), "", "Module path number 4,2"
End Sub

Public Sub TestRootPathForExistingProject()
    mAssert.Equals existingConfManager.rootPath, vtkPathOfCurrentProject, "The root Path is not initialized for a new Workbook"
    mAssert.Equals existingConfManager.rootPath, vtkPathOfCurrentProject, "The second call to rootPath give the same result as the previous one"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_PropertyName_DefaultGet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_PropertyName_DefaultLet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationSheetExistsInExistingWorkbook"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationSheetDoesntExistInNewWorkbook"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationSheetCreationForNewProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationSheetRetrievalForExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationSheetFormatForNewProjet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationSheetFormatForExistingProjet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetConfigurationsFromNewProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetConfigurationsFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_AddConfigurationInExistingProject_Name"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_AddConfigurationInExistingProject_Cells"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_AddConfigurationInExistingProject_NullPathes"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetConfigurationPathWithNumberFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestSetConfigurationPathWithNumberToNewProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestSetConfigurationPathWithNumberToSavedProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetConfigurationNumbersFromNewProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetConfigurationPathFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestSetConfigurationPathToSavedProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetModulesFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetModulesFromNewProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetModuleNumbersFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestAddNonExistantModuleToSavedProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestAddExistantModuleToExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestAddModuleWithExistingStringToExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetModulePathWithNumberFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetModulePathWithNumberForNewModule"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestAddModulePathToSavedProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestRootPathForExistingProject"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_PropertyName_DefaultGet": Test_PropertyName_DefaultGet
        Case "Test_PropertyName_DefaultLet": Test_PropertyName_DefaultLet
        Case "TestConfigurationSheetExistsInExistingWorkbook": TestConfigurationSheetExistsInExistingWorkbook
        Case "TestConfigurationSheetDoesntExistInNewWorkbook": TestConfigurationSheetDoesntExistInNewWorkbook
        Case "TestConfigurationSheetCreationForNewProject": TestConfigurationSheetCreationForNewProject
        Case "TestConfigurationSheetRetrievalForExistingProject": TestConfigurationSheetRetrievalForExistingProject
        Case "TestConfigurationSheetFormatForNewProjet": TestConfigurationSheetFormatForNewProjet
        Case "TestConfigurationSheetFormatForExistingProjet": TestConfigurationSheetFormatForExistingProjet
        Case "TestGetConfigurationsFromNewProject": TestGetConfigurationsFromNewProject
        Case "TestGetConfigurationsFromExistingProject": TestGetConfigurationsFromExistingProject
        Case "Test_AddConfigurationInExistingProject_Name": Test_AddConfigurationInExistingProject_Name
        Case "Test_AddConfigurationInExistingProject_Cells": Test_AddConfigurationInExistingProject_Cells
        Case "Test_AddConfigurationInExistingProject_NullPathes": Test_AddConfigurationInExistingProject_NullPathes
        Case "TestGetConfigurationPathWithNumberFromExistingProject": TestGetConfigurationPathWithNumberFromExistingProject
        Case "TestSetConfigurationPathWithNumberToNewProject": TestSetConfigurationPathWithNumberToNewProject
        Case "TestSetConfigurationPathWithNumberToSavedProject": TestSetConfigurationPathWithNumberToSavedProject
        Case "TestGetConfigurationNumbersFromNewProject": TestGetConfigurationNumbersFromNewProject
        Case "TestGetConfigurationPathFromExistingProject": TestGetConfigurationPathFromExistingProject
        Case "TestSetConfigurationPathToSavedProject": TestSetConfigurationPathToSavedProject
        Case "TestGetModulesFromExistingProject": TestGetModulesFromExistingProject
        Case "TestGetModulesFromNewProject": TestGetModulesFromNewProject
        Case "TestGetModuleNumbersFromExistingProject": TestGetModuleNumbersFromExistingProject
        Case "TestAddNonExistantModuleToSavedProject": TestAddNonExistantModuleToSavedProject
        Case "TestAddExistantModuleToExistingProject": TestAddExistantModuleToExistingProject
        Case "TestAddModuleWithExistingStringToExistingProject": TestAddModuleWithExistingStringToExistingProject
        Case "TestGetModulePathWithNumberFromExistingProject": TestGetModulePathWithNumberFromExistingProject
        Case "TestGetModulePathWithNumberForNewModule": TestGetModulePathWithNumberForNewModule
        Case "TestAddModulePathToSavedProject": TestAddModulePathToSavedProject
        Case "TestRootPathForExistingProject": TestRootPathForExistingProject
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

