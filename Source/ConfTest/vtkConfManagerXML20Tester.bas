VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkConfManagerXML20Tester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkConfManagerXML20Tester
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Test the vtkConfigurationManagerXML class
'             with vtkConfigurations XML version 2.0
'
' Copyright 2014 Skwal-Soft (http://skwalsoft.com)
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

Private Const existingXMLNameForTest As String = "XMLForConfigurationsTests.xml"
Private existingConfManager As vtkConfigurationManager   ' Configuration Manager for the existing workbook
Private Const existingProjectName As String = "ExistingProject"

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
    
    getTestFileFromTemplate fileName:=existingXMLNameForTest, destinationName:=existingProjectName & ".xml", openExcel:=False
    Set existingConfManager = New vtkConfigurationManagerXML
    existingConfManager.init vtkTestPath & "\" & existingProjectName & ".xml"
End Sub

Private Sub ITestCase_TearDown()
End Sub

'Public Sub Test_PropertyName_DefaultGet()
'    '   Verify that the Property Name is the Default property for vtkConfigurationManager
'    '   - In fact there is no need to run the test, just to compile it
'    mAssert.Equals newConfManager, "NewProject", "The name property must be the default one for vtkConfigurationManager"
'End Sub
'
'Public Sub Test_PropertyName_DefaultLet()
'    '   Verify that the Property Name is the Default property for vtkConfigurationManager
'    '   - In fact there is no need to run the test, just to compile it
'    '   - both existing project and new project worbooks must be opened
'    mAssert.Equals existingConfManager, "ExistingProject", "The name property of existingConf before modification"
'    existingConfManagerExcel = "NewProject"
'    mAssert.Equals existingConfManager, "NewProject", "The name property of existingConf after modification"
'End Sub
'
'Public Sub TestConfigurationSheetExistsInExistingWorkbook()
'    '   Verify that the configuration sheet presence is detected in existing workbook
'    '   using a fresh configuration Manager (with no default sheet initialized)
'    Dim cm As New vtkConfigurationManagerExcel
'    mAssert.Should cm.isConfigurationInitializedForWorkbook(ExcelName:=existingWorkbookName), "The Configuration sheet must exist in existing workbook"
'End Sub
'
'Public Sub TestConfigurationSheetDoesntExistInNewWorkbook()
'    '   Verify that the configuration sheet missing is created in new workbook
'    '   using a fresh configuration Manager (with no default sheet initialized)
'    Dim cm As New vtkConfigurationManagerExcel, wb As Workbook, wbFullName As String
'    Set wb = vtkCreateExcelWorkbookForTestWithProjectName("NewWorkbook")    ' create a fresh new Excel workbook
'    wbFullName = wb.FullName
'    mAssert.Should Not cm.isConfigurationInitializedForWorkbook(ExcelName:=wb.name), "The Configuration sheet must not exist in new workbook"
'    wb.Close saveChanges:=False
'    Kill PathName:=wbFullName
'End Sub
'
'Public Sub TestConfigurationSheetCreationForNewProject()
''       Verify that a Configuration Sheet is created in a new project
'    Dim ws As Worksheet
'    On Error Resume Next
'    Set ws = newWorkbook.Sheets("vtkConfigurations")
'    mAssert.Equals Err, 0, "A configuration manager must create a Configuration sheet"
'    On Error GoTo 0
'    mAssert.Should newWorkbook.Sheets("vtkConfigurations") Is newConfManagerExcel.configurationSheet, "The configurationSheet property of the conf manager must be equal to the configuration sheet of the workbook"
'End Sub
'
'Public Sub TestConfigurationSheetRetrievalForExistingProject()
''       Verify that a Configuration Sheet is retreived in an existing project
'    Dim ws As Worksheet
'    On Error Resume Next
'    Set ws = existingWorkbook.Sheets("vtkConfigurations")
'    mAssert.Equals Err, 0, "A configuration manager must be accessible in an existing project"
'    On Error GoTo 0
'    mAssert.Should existingWorkbook.Sheets("vtkConfigurations") Is existingConfManagerExcel.configurationSheet, "The configurationSheet property of the conf manager must be equal to the configuration sheet of the workbook"
'End Sub
'
'Public Sub TestConfigurationSheetFormatForNewProjet()
''       Verify the newly created configuration sheet of a new project
'    Dim ws As Worksheet
'    Set ws = newConfManagerExcel.configurationSheet
'    mAssert.Equals ws.Range("A1"), "vtkConfigurations v1.1", "Expected identification of the configuration sheet"
'    mAssert.Equals ws.Range("A2"), "", "Expected Title for Modules column"
'    mAssert.Equals ws.Range("A3"), "Path, template, name and comment", "Expected Title for Configurations columns"
'    mAssert.Equals ws.Range("A4"), "", "Expected Title for Modules column"
'    mAssert.Equals ws.Range("A5"), "Module Name", "Expected Title for Modules column"
'    mAssert.Equals ws.Range("B1"), newProjectName, "Expected Title for main project column"
'    mAssert.Equals ws.Range("B2"), "Delivery\" & newProjectName & ".xlsm", "Expected related Path for new main workbook"
'    mAssert.Equals ws.Range("B3"), "", "Expected Template path for main workbook"
'    mAssert.Equals ws.Range("B4"), "", "Expected Project name for new workbook"
'    mAssert.Equals ws.Range("B5"), "", "Expected Comment for new workbook"
'    mAssert.Equals ws.Range("C1"), newProjectName & "_DEV", "Expected Title for development project column"
'    mAssert.Equals ws.Range("C2"), "Project\" & newWorkbookName, "Expected related Path for new development workbook"
'    mAssert.Equals ws.Range("C3"), "", "Expected Template path for development workbook"
'    mAssert.Equals ws.Range("C4"), "", "Expected Project name for new workbook"
'    mAssert.Equals ws.Range("C5"), "", "Expected Comment for new workbook"
'End Sub
'
'Public Sub TestConfigurationSheetFormatForExistingProjet()
''       Verify the retrieved configuration sheet from an existing project
'    Dim ws As Worksheet
'    Set ws = existingConfManagerExcel.configurationSheet
'    mAssert.Equals ws.Range("A1"), "vtkConfigurations v1.1", "Expected identification of the configuration sheet"
'    mAssert.Equals ws.Range("A2"), "", "Expected empty cell"
'    mAssert.Equals ws.Range("A3"), "Path, template, name and comment", "Expected Title for Configurations columns"
'    mAssert.Equals ws.Range("A4"), "", "Expected Title for Modules column"
'    mAssert.Equals ws.Range("A5"), "Module Name", "Expected Title for Modules column"
'    mAssert.Equals ws.Range("B1"), existingProjectName, "Expected Title for main project column"
'    mAssert.Equals ws.Range("B2"), "Delivery\ExistingProject.xlsm", "Expected related Path for main workbook"
'    mAssert.Equals ws.Range("B3"), "", "Expected Template path for main workbook"
'    mAssert.Equals ws.Range("B4"), "ExistingProjectName", "Expected Project name for main workbook"
'    mAssert.Equals ws.Range("B5"), "Existing project for various tests of VBAToolKit", "Expected Comment for main workbook"
'    mAssert.Equals ws.Range("C1"), existingProjectName & "_DEV", "Expected Title for development project column"
'    mAssert.Equals ws.Range("C2"), "Project\ExistingProject_DEV.xlsm", "Expected related Path for development workbook"
'    mAssert.Equals ws.Range("C3"), "", "Expected Template path for development workbook"
'    mAssert.Equals ws.Range("C4"), "", "Expected empty projectName for development workbook"
'    mAssert.Equals ws.Range("C5"), "Existing project for development for various tests of VBAToolKit", "Expected Comment for development workbook"
'End Sub
'
'Public Sub TestGetConfigurationsFromNewProject()
''       Verify the list of the configurations of a new project
'    mAssert.Equals newConfManager.configurationCount, 2, "There must be two configurations in a new project"
'    mAssert.Equals newConfManager.configuration(0), "", "Inexistant configuration number 0"
'    mAssert.Equals newConfManager.configuration(1), newProjectName, "Name of the first configuration"
'    mAssert.Equals newConfManager.configuration(2), newProjectName & "_DEV", "Name of the second configuration"
'    mAssert.Equals newConfManager.configuration(3), "", "Inexistant configuration number 3"
'    mAssert.Equals newConfManager.configuration(-23), "", "Inexistant configuration number -23"
'    mAssert.Equals newConfManager.configuration(150), "", "Inexistant configuration number 150"
'End Sub
'
'Public Sub TestGetConfigurationsFromExistingProject()
''       Verify the list of the configurations of an existing project
'    mAssert.Equals existingConfManager.configurationCount, 2, "There must be two configurations in the existing template project"
'    mAssert.Equals existingConfManager.configuration(0), "", "Inexistant configuration number 0"
'    mAssert.Equals existingConfManager.configuration(1), existingProjectName, "Name of the first configuration"
'    mAssert.Equals existingConfManager.configuration(2), existingProjectName & "_DEV", "Name of the second configuration"
'    mAssert.Equals existingConfManager.configuration(3), "", "Inexistant configuration number 3"
'    mAssert.Equals existingConfManager.configuration(-23), "", "Inexistant configuration number -23"
'    mAssert.Equals existingConfManager.configuration(150), "", "Inexistant configuration number 150"
'End Sub
'
'Public Sub Test_AddConfigurationInExistingProject_Name()
''       Verify the add of configuration in an existing project
''       Verify the number and name of the added configuration
'    Dim n As Integer
'    n = existingConfManager.addConfiguration("NewConfiguration", "ConfigurationPath")
'
'    mAssert.Equals existingConfManager.configurationCount, 3, "There must be two configurations in the existing template project"
'    mAssert.Equals existingConfManager.configuration(0), "", "Inexistant configuration number 0"
'    mAssert.Equals existingConfManager.configuration(1), existingProjectName, "Name of the first configuration"
'    mAssert.Equals existingConfManager.configuration(2), existingProjectName & "_DEV", "Name of the second configuration"
'    mAssert.Equals existingConfManager.configuration(3), "NewConfiguration", "Name of the new configuration"
'    mAssert.Equals existingConfManager.getConfigurationPathWithNumber(3), "ConfigurationPath", "Path of new configuration given by number"
'    mAssert.Equals existingConfManager.configuration(4), "", "Inexistant configuration number 4"
'    mAssert.Equals existingConfManager.configuration(-23), "", "Inexistant configuration number -23"
'    mAssert.Equals existingConfManager.configuration(150), "", "Inexistant configuration number 150"
'End Sub
'
'Public Sub Test_AddConfigurationInExistingProject_Cells()
''       Verify the add of configuration in an existing project
''       Verify the cells of the configuration sheet
'    Dim ws As Worksheet
'    Dim n As Integer
'    n = existingConfManager.addConfiguration("NewConfiguration", "ConfigurationPath", comment:="New comment")
'
'    Set ws = existingConfManagerExcel.configurationSheet
'    mAssert.Equals ws.Range("A1"), "vtkConfigurations v1.1", "Expected identification of the configuration sheet"
'    mAssert.Equals ws.Range("A2"), "", "Expected Title for Modules column"
'    mAssert.Equals ws.Range("A3"), "Path, template, name and comment", "Expected Title for Configurations columns"
'    mAssert.Equals ws.Range("A4"), "", "Expected Title for Modules column"
'    mAssert.Equals ws.Range("A5"), "Module Name", "Expected Title for Modules column"
'    mAssert.Equals ws.Range("B1"), existingProjectName, "Expected Title for main project column"
'    mAssert.Equals ws.Range("B2"), "Delivery\ExistingProject.xlsm", "Expected related Path for new main workbook"
'    mAssert.Equals ws.Range("B3"), "", "Expected Template path for main workbook"
'    mAssert.Equals ws.Range("B4"), "ExistingProjectName", "Expected Project name for main workbook"
'    mAssert.Equals ws.Range("B5"), "Existing project for various tests of VBAToolKit", "Expected Comment for main workbook"
'    mAssert.Equals ws.Range("C1"), existingProjectName & "_DEV", "Expected Title for development project column"
'    mAssert.Equals ws.Range("C2"), "Project\ExistingProject_DEV.xlsm", "Expected related Path for new development workbook"
'    mAssert.Equals ws.Range("C3"), "", "Expected Template path for development workbook"
'    mAssert.Equals ws.Range("C4"), "", "Expected empty projectName for development workbook"
'    mAssert.Equals ws.Range("C5"), "Existing project for development for various tests of VBAToolKit", "Expected Comment for development workbook"
'    mAssert.Equals ws.Range("D1"), "NewConfiguration", "Expected Title for new configuration column"
'    mAssert.Equals ws.Range("D2"), "ConfigurationPath", "Expected related Path for new configuration"
'    mAssert.Equals ws.Range("D3"), "", "Expected Template path for new configuration"
'    mAssert.Equals ws.Range("D4"), "", "Expected empty projectName for new configuration"
'    mAssert.Equals ws.Range("D5"), "New comment", "Expected Comment for new configuration"
'End Sub
'
'Public Sub Test_AddConfigurationInExistingProject_NullPathes()
''       Verify the add of configuration in an existing project
''       Verify That the module pathes are initialized to null
'    Dim n As Integer, i As Integer
'    n = existingConfManager.addConfiguration("NewConfiguration", "ConfigurationPath")
'
'    For i = 1 To existingConfManager.configurationCount
'        mAssert.Equals existingConfManager.getModulePathWithNumber(i, n), "", "Path of module " & i & " must be null"
'    Next
'End Sub
'
'Public Sub TestGetConfigurationPathWithNumberFromExistingProject()
''       Verify the capability to get the configuration path by number
'    mAssert.Equals existingConfManager.getConfigurationPathWithNumber(0), "", "Inexistant configuration number 0"
'    mAssert.Equals existingConfManager.getConfigurationPathWithNumber(1), "Delivery\ExistingProject.xlsm", "Path of first configuration given by number"
'    mAssert.Equals existingConfManager.getConfigurationPathWithNumber(2), "Project\ExistingProject_DEV.xlsm", "Path of second configuration given by number"
'    mAssert.Equals existingConfManager.getConfigurationPathWithNumber(3), "", "Inexistant configuration number 3"
'End Sub
'
'Public Sub TestSetConfigurationPathWithNumberToNewProject()
''       Verify the capability to set and retrieve the configuration path by number
'    ' set new pathes
'    newConfManager.setConfigurationPathWithNumber n:=0, path:="Path0"
'    newConfManager.setConfigurationPathWithNumber n:=1, path:="Path1"
'    newConfManager.setConfigurationPathWithNumber n:=2, path:="Path2"
'    newConfManager.setConfigurationPathWithNumber n:=3, path:="Path3"
'    ' verify pathes
'    mAssert.Equals newConfManager.getConfigurationPathWithNumber(0), "", "Inexistant configuration number 0"
'    mAssert.Equals newConfManager.getConfigurationPathWithNumber(1), "Path1", "Path of first configuration given by number"
'    mAssert.Equals newConfManager.getConfigurationPathWithNumber(2), "Path2", "Path of second configuration given by number"
'    mAssert.Equals newConfManager.getConfigurationPathWithNumber(3), "", "Inexistant configuration number 3"
'End Sub
'
'Public Sub TestSetConfigurationPathWithNumberToSavedProject()
''       Verify the capability to set and retrieve the configuration path by number
'    ' set new pathes
'    newConfManager.setConfigurationPathWithNumber n:=0, path:="Path0"
'    newConfManager.setConfigurationPathWithNumber n:=1, path:="Path1"
'    newConfManager.setConfigurationPathWithNumber n:=2, path:="Path2"
'    newConfManager.setConfigurationPathWithNumber n:=3, path:="Path3"
'    ' save and re-open file
'    SaveThenReOpenNewWorkbook
'    ' verify pathes
'    mAssert.Equals newConfManager.getConfigurationPathWithNumber(0), "", "Inexistant configuration number 0"
'    mAssert.Equals newConfManager.getConfigurationPathWithNumber(1), "Path1", "Path of first configuration given by number"
'    mAssert.Equals newConfManager.getConfigurationPathWithNumber(2), "Path2", "Path of second configuration given by number"
'    mAssert.Equals newConfManager.getConfigurationPathWithNumber(3), "", "Inexistant configuration number 3"
'End Sub
'
'Public Sub TestGetConfigurationNumbersFromNewProject()
''       Verify the capability to get the number of a configuration
'    mAssert.Equals newConfManager.configurationCount, 2, "There must be two configurations in a new project"
'    mAssert.Equals newConfManager.getConfigurationNumber(newProjectName), 1, "Number of the main configuration"
'    mAssert.Equals newConfManager.getConfigurationNumber(newProjectName & "_DEV"), 2, "Number of the Development configuration"
'    mAssert.Equals newConfManager.getConfigurationNumber("InexistantConfiguration"), 0, "Inexistant configuration"
'End Sub
'
'Public Sub TestGetConfigurationPathFromExistingProject()
''       Verify the capability to get a configutaion path given the configuration name
'    mAssert.Equals existingConfManager.getConfigurationPath(existingProjectName), "Delivery\ExistingProject.xlsm", "Path of the main configuration"
'    mAssert.Equals existingConfManager.getConfigurationPath(existingProjectName & "_DEV"), "Project\ExistingProject_DEV.xlsm", "Path of the Development configuration"
'    mAssert.Equals existingConfManager.getConfigurationPath("InexistantConfiguration"), "", "Inexistant configuration"
'End Sub
'
'Public Sub TestSetConfigurationPathToSavedProject()
''       Verify the capability to set and retrieve the configuration path by configuration name
'    ' set new pathes
'    newConfManager.setConfigurationPath configuration:="InexistantConfiguration", path:="Path0"
'    newConfManager.setConfigurationPath configuration:=newProjectName, path:="Path1"
'    newConfManager.setConfigurationPath configuration:=newProjectName & "_DEV", path:="Path2"
'    ' save and re-open file
'    SaveThenReOpenNewWorkbook
'    ' verify pathes
'    mAssert.Equals newConfManager.getConfigurationPath("InexistantConfiguration"), "", "Inexistant configuration"
'    mAssert.Equals newConfManager.getConfigurationPath(newProjectName), "Path1", "Path of first configuration given by name"
'    mAssert.Equals newConfManager.getConfigurationPath(newProjectName & "_DEV"), "Path2", "Path of second configuration given by name"
'End Sub
'
'Public Sub TestGetModulesFromExistingProject()
''       Verify the capability to retrieve the list of Modules from an existing project
'    mAssert.Equals existingConfManager.moduleCount, 5, "There must be five configurations in the existing project"
'    mAssert.Equals existingConfManager.module(0), "", "Inexistant module number 0"
'    mAssert.Equals existingConfManager.module(1), "Module1", "Name of the first module"
'    mAssert.Equals existingConfManager.module(2), "Module2", "Name of the second module"
'    mAssert.Equals existingConfManager.module(3), "Module3", "Name of the third module"
'    mAssert.Equals existingConfManager.module(4), "Module4", "Name of the fourth module"
'    mAssert.Equals existingConfManager.module(5), "Module5", "Name of the fifth module"
'    mAssert.Equals existingConfManager.module(6), "", "Inexistant module number 6"
'    mAssert.Equals existingConfManager.module(-23), "", "Inexistant module number -23"
'    mAssert.Equals existingConfManager.module(150), "", "Inexistant module number 150"
'End Sub
'
'Public Sub TestGetModulesFromNewProject()
''       Verify the capability to retrieve the list of Modules from an existing project
'    mAssert.Equals newConfManager.moduleCount, 0, "There must be no modules in a new project"
'    mAssert.Equals newConfManager.module(0), "", "Inexistant module number 0"
'    mAssert.Equals newConfManager.module(1), "", "Inexistant module number 1"
'    mAssert.Equals newConfManager.module(6), "", "Inexistant module number 6"
'    mAssert.Equals newConfManager.module(-23), "", "Inexistant module number -23"
'    mAssert.Equals newConfManager.module(150), "", "Inexistant module number 150"
'End Sub
'
'Public Sub TestGetModuleNumbersFromExistingProject()
''       Verify the capability to get the number of a configuration
'    mAssert.Equals existingConfManager.getModuleNumber("Module0"), 0, "Inexistant module"
'    mAssert.Equals existingConfManager.getModuleNumber("Module1"), 1, "First Module"
'    mAssert.Equals existingConfManager.getModuleNumber("Module2"), 2, "Second Module"
'    mAssert.Equals existingConfManager.getModuleNumber("Module3"), 3, "Third module"
'    mAssert.Equals existingConfManager.getModuleNumber("Module4"), 4, "Fourth module"
'    mAssert.Equals existingConfManager.getModuleNumber("Module5"), 5, "Fifth module"
'    mAssert.Equals existingConfManager.getModuleNumber("InexistantModule"), 0, "Inexistant module"
'End Sub
'
'Public Sub TestAddNonExistantModuleToSavedProject()
''       Verify the capability to add a new module, non existant, and retrieve it
'    ' set new modules
'    mAssert.Equals newConfManager.addModule(module:="NewModule1"), 1, "Number of the first module added"
'    mAssert.Equals newConfManager.addModule(module:="NewModule2"), 2, "Number of the second module added"
'    ' save and re-open file
'    SaveThenReOpenNewWorkbook
'    ' verify modules
'    mAssert.Equals newConfManager.moduleCount, 2, "There must be two new modules in the saved project"
'    mAssert.Equals newConfManager.module(0), "", "Inexistant module number 0"
'    mAssert.Equals newConfManager.module(1), "NewModule1", "New module number 1"
'    mAssert.Equals newConfManager.module(2), "NewModule2", "New module number 2"
'    mAssert.Equals newConfManager.module(3), "", "Inexistant module number 3"
'End Sub
'
'Public Sub TestAddExistantModuleToExistingProject()
''       Verify the capability to not add an existing module in an existing project
'    Dim n As Integer
'    ' set new modules
'    mAssert.Equals existingConfManager.addModule(module:="Module1"), -1, "Number of the first existing module"
'    mAssert.Equals existingConfManager.addModule(module:="Module5"), -5, "Number of the fifth existing module"
'    mAssert.Equals existingConfManager.moduleCount, 5, "There must be five modules, no change, in the existing project"
'End Sub
'
'Public Sub TestAddModuleWithExistingStringToExistingProject()
''       Verify the capability to add a new module whose name is included in existing module in an existing project
'    Dim n As Integer
'    ' set new modules
'    mAssert.Equals existingConfManager.addModule(module:="Module"), 6, "Number for the new module"
'    mAssert.Equals existingConfManager.moduleCount, 6, "There must be six modules, one more module, in the existing project"
'End Sub
'
'Public Sub TestGetModulePathWithNumberFromExistingProject()
''       Verify the capability to get the module path by number
'    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=0, numConfiguration:=2), "", "Inexistant module path number 0,2"
'    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=3, numConfiguration:=3), "", "Inexistant module path number 3,3"
'    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=1, numConfiguration:=1), "Path1Module1", "Module path number 1,1"
'    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=1, numConfiguration:=2), "", "Module path number 1,2"
'    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=2, numConfiguration:=1), "", "Module path number 2,1"
'    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=2, numConfiguration:=2), "Path2Module2", "Module path number 2,2"
'    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=3, numConfiguration:=1), "", "Module path number 3,1"
'    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=3, numConfiguration:=2), "", "Module path number 3,2"
'    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=4, numConfiguration:=1), "Path1Module4", "Module path number 4,1"
'    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=4, numConfiguration:=2), "Path2Module4", "Module path number 4,2"
'    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=5, numConfiguration:=1), "", "Module path number 5,1"
'    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=5, numConfiguration:=2), "Path2Module5", "Module path number 5,2"
'End Sub
'
'Public Sub TestGetModulePathWithNumberForNewModule()
''       Verify the default module path define at module adding
'    mAssert.Equals newConfManager.addModule(module:="NewModule1"), 1, "Number of the first module added"
'    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=1, numConfiguration:=1), "", "Inexistant module path number 1,1"
'    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=1, numConfiguration:=2), "", "Inexistant module path number 1,2"
'End Sub
'
'Public Sub TestAddModulePathToSavedProject()
''       Verify the capability to add a new module, non existant, and retrieve it
'    ' set new modules
'    mAssert.Equals newConfManager.addModule(module:="NewModule1"), 1, "Number of the first module added"
'    mAssert.Equals newConfManager.addModule(module:="NewModule2"), 2, "Number of the second module added"
'    mAssert.Equals newConfManager.addModule(module:="NewModule3"), 3, "Number of the third module added"
'    mAssert.Equals newConfManager.addModule(module:="NewModule4"), 4, "Number of the fourth module added"
'    ' save and re-open file
'    SaveThenReOpenNewWorkbook
'    ' set new pathes
'    newConfManager.setModulePathWithNumber path:="Path1Module1", numModule:=1, numConfiguration:=1
'    newConfManager.setModulePathWithNumber path:="Path2Module1", numModule:=1, numConfiguration:=2
'    newConfManager.setModulePathWithNumber path:="Path1Module2", numModule:=2, numConfiguration:=1
'    newConfManager.setModulePathWithNumber path:="Path2Module3", numModule:=3, numConfiguration:=2
'    ' save and re-open file
'    SaveThenReOpenNewWorkbook
'    ' verify module pathes
'    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=1, numConfiguration:=1), "Path1Module1", "Module path number 1,1"
'    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=1, numConfiguration:=2), "Path2Module1", "Module path number 1,2"
'    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=2, numConfiguration:=1), "Path1Module2", "Module path number 2,1"
'    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=2, numConfiguration:=2), "", "Module path number 2,2"
'    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=3, numConfiguration:=1), "", "Module path number 3,1"
'    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=3, numConfiguration:=2), "Path2Module3", "Module path number 3,2"
'    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=4, numConfiguration:=1), "", "Module path number 4,1"
'    mAssert.Equals newConfManager.getModulePathWithNumber(numModule:=4, numConfiguration:=2), "", "Module path number 4,2"
'End Sub
'
'Public Sub TestRootPathForExistingProject()
'    mAssert.Equals existingConfManager.rootPath, vtkPathOfCurrentProject, "The root Path is not initialized for a new Workbook"
'    mAssert.Equals existingConfManager.rootPath, vtkPathOfCurrentProject, "The second call to rootPath give the same result as the previous one"
'End Sub
'
'Public Sub TestConfigurationSheetFormatAfterConversion()
''       Verify the newly converted configuration sheet of a new project
'    Dim ws As Worksheet, workbookV10 As Workbook, confManagerV10 As vtkConfigurationManagerExcel
'
'    Set workbookV10 = getTestFileFromTemplate(fileName:=workbookNameWithConfigurationV10, destinationName:="ProjectV10_DEV.xlsm", openExcel:=True)
'    Set confManagerV10 = vtkConfigurationManagerForProject("ProjectV10")
'    Set ws = confManagerV10.configurationSheet
'    mAssert.Equals ws.Range("A1"), "vtkConfigurations v1.0", "Expected identification of the configuration sheet"
'
'    confManagerV10.updateConfigurationSheetFormat
'
'    mAssert.Equals ws.Range("A1"), "vtkConfigurations v1.1", "Expected identification of the configuration sheet"
'    mAssert.Equals ws.Range("A2"), "", "Expected Title for Modules column"
'    mAssert.Equals ws.Range("A3"), "Path, template, name and comment", "Expected Title for Configurations columns"
'    mAssert.Equals ws.Range("A4"), "", "Expected Title for Modules column"
'    mAssert.Equals ws.Range("A5"), "Module Name", "Expected Title for Modules column"
'    mAssert.Equals ws.Range("B1"), "ExistingProject", "Expected Title for main project column"
'    mAssert.Equals ws.Range("B2"), "Delivery\ExistingProject.xlsm", "Expected related Path for new main workbook"
'    mAssert.Equals ws.Range("B3"), "", "Expected Template path for main workbook"
'    mAssert.Equals ws.Range("B4"), "", "Expected Project name for new workbook"
'    mAssert.Equals ws.Range("B5"), "", "Expected Comment for new workbook"
'    mAssert.Equals ws.Range("C1"), "ExistingProject_DEV", "Expected Title for development project column"
'    mAssert.Equals ws.Range("C2"), "Project\ExistingProject_DEV.xlsm", "Expected related Path for new development workbook"
'    mAssert.Equals ws.Range("C3"), "", "Expected Template path for development workbook"
'    mAssert.Equals ws.Range("C4"), "", "Expected Project name for new workbook"
'    mAssert.Equals ws.Range("C5"), "", "Expected Comment for new workbook"
'
'    mAssert.Equals confManagerV10.configurationCount, 2, "There must be two configurations in the converted project"
'    mAssert.Equals confManagerV10.moduleCount, 5, "There must be two modules in the converted project"
'
'    vtkCloseAndKillWorkbook wb:=workbookV10 ' close the V10 Excel project
'End Sub
'
'Public Sub TestConfigurationSheetFormatUpToDateConversion()
''       Verify that an up to date configuration sheet is not modified when converted
'    Dim ws As Worksheet
'
'    existingConfManagerExcel.updateConfigurationSheetFormat
'
'    Set ws = existingConfManagerExcel.configurationSheet
'    mAssert.Equals ws.Range("A1"), "vtkConfigurations v1.1", "Expected identification of the configuration sheet"
'    mAssert.Equals ws.Range("A2"), "", "Expected Title for Modules column"
'    mAssert.Equals ws.Range("A3"), "Path, template, name and comment", "Expected Title for Configurations columns"
'    mAssert.Equals ws.Range("A4"), "", "Expected Title for Modules column"
'    mAssert.Equals ws.Range("A5"), "Module Name", "Expected Title for Modules column"
'    mAssert.Equals ws.Range("B1"), "ExistingProject", "Expected Title for main project column"
'    mAssert.Equals ws.Range("B2"), "Delivery\ExistingProject.xlsm", "Expected related Path for new main workbook"
'    mAssert.Equals ws.Range("B3"), "", "Expected Template path for main workbook"
'    mAssert.Equals ws.Range("B4"), "ExistingProjectName", "Expected Project name for existing workbook"
'    mAssert.Equals ws.Range("B5"), "Existing project for various tests of VBAToolKit", "Expected Comment for existing workbook"
'    mAssert.Equals ws.Range("C1"), "ExistingProject_DEV", "Expected Title for development project column"
'    mAssert.Equals ws.Range("C2"), "Project\ExistingProject_DEV.xlsm", "Expected related Path for new development workbook"
'    mAssert.Equals ws.Range("C3"), "", "Expected Template path for development workbook"
'    mAssert.Equals ws.Range("C4"), "", "Expected Project name for existing workbook"
'    mAssert.Equals ws.Range("C5"), "Existing project for development for various tests of VBAToolKit", "Expected Comment for existing workbook"
'
'    mAssert.Equals existingConfManager.configurationCount, 2, "There must be two configurations in the up to date project"
'    mAssert.Equals existingConfManager.moduleCount, 5, "There must be two modules in the up to date project"
'
'End Sub
'
'
'Public Sub TestGetAllReferencesFromNewWorkbook()
''       Verify that all standard references are listed
'    Dim refNames(), i As Integer, c1 As Collection, c2 As Collection, r As vtkReference
'    refNames = Array("Scripting", "VBIDE", "Shell32", "MSXML2", "ADODB", "VBAToolKit_DEV", "VBA", "Excel", "stdole", "Office", "MSForms")
'
'   On Error Resume Next
'    Set c1 = newConfManager.references
'    ' Rearrange the collection by name
'    Set c2 = New Collection
'    For Each r In c1
'        c2.Add r, r.name
'    Next
'    mAssert.Equals c2.count, UBound(refNames) - LBound(refNames) + 1, "Count of all references of a new workbook"
'    ' il faut boucler sur le tableau et rechercher dans la collection (si pas trouvé = erreur)
'    For i = LBound(refNames) To UBound(refNames)
'        Set r = c2(refNames(i))
'        mAssert.Equals Err.Number, 0, "Error when getting " & refNames(i) & " reference"
'    Next i
'   On Error GoTo 0
'End Sub
'
'Public Sub TestGetAllReferencesFromExistingWorkbook()
''       Verify that all references are listed from existing workbook
'    Dim refNames(), i As Integer, c1 As Collection, c2 As Collection, r As vtkReference
'    refNames = Array("VBA", "Excel", "stdole", "Office", "MSForms", "Scripting", "VBIDE", "Shell32", "MSXML2", "VBAToolKit", "EventSystemLib")
'
'   On Error Resume Next
'    Set c1 = existingConfManager.references
'    ' Rearrange the collection by name
'    Set c2 = New Collection
'    For Each r In c1
'        c2.Add r, r.name
'    Next
'    mAssert.Equals c2.count, UBound(refNames) - LBound(refNames) + 1, "Count of all references of a new workbook"
'    ' il faut boucler sur le tableau et rechercher dans la collection (si pas trouvé = erreur)
'    For i = LBound(refNames) To UBound(refNames)
'        Set r = c2(refNames(i))
'        mAssert.Equals Err.Number, 0, "Error when getting " & refNames(i) & " reference"
'    Next i
'   On Error GoTo 0
'End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


