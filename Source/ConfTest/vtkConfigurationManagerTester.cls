VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkConfigurationManagerTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITest
Implements ITestCase

Private mManager As TestCaseManager
Private mAssert As IAssert

Private newWorkbook As Workbook         ' New Workbook created for each test
Private newWorkbookName As String
Private newConfManager As vtkConfigurationManager   ' Configuration Manager for the new workbook
Private Const newProjectName As String = "NewProject"

Private existingWorkbook As Workbook    ' Existing Workbook opened for each test
Private Const existingWorkbookName As String = "WorkBookForConfigurationsTests.xlsm"
Private existingConfManager As vtkConfigurationManager   ' Configuration Manager for the new workbook
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
    Set newWorkbook = vtkCreateExcelWorkbookWithProjectName(newProjectName)    ' create an Excel project for tests
    newWorkbookName = newWorkbook.name
    Set newConfManager = vtkConfigurationManagerForWorkbook(newWorkbookName)
    FileCopy Source:=vtkPathToTemplateFolder & "\" & existingWorkbookName, Destination:=vtkPathToTestFolder & "\" & existingWorkbookName
    Set existingWorkbook = Workbooks.Open(Filename:=vtkPathToTestFolder & "\" & existingWorkbookName)
    Set existingConfManager = vtkConfigurationManagerForWorkbook(existingWorkbookName)
End Sub

Private Sub ITestCase_TearDown()
    newWorkbook.Close SaveChanges:=False      ' close the new Excel project
    existingWorkbook.Close SaveChanges:=False ' close the existing Excel project
    vtkResetConfigurationManagers           ' reset all configuration managers
    On Error Resume Next
    Kill PathName:=vtkPathToTestFolder & "\" & "TestFile.xlsx"
    On Error GoTo 0
End Sub

'   Liste des tests à faire
'       Ajout d'un module
'       Définition du chemin pour un module
'       Définition du chemin pour une configuration
'       et les mêmes tests lors de la lecture d'une configuration dans un fichier existant

Public Sub TestConfigurationSheetCreationForNewProject()
'       Verify that a Configuration Sheet is created in a new project
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = newWorkbook.Sheets("vtkConfigurations")
    mAssert.Equals Err, 0, "A configuration manager must create a Configuration sheet"
    On Error GoTo 0
    mAssert.Should newWorkbook.Sheets("vtkConfigurations") Is newConfManager.configurationSheet, "The configurationSheet property of the conf manager must be equal to the configuration sheet of the workbook"
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
    mAssert.Equals ws.Range("B2"), "Delivery\" & newWorkbookName, "Expected related Path for new main workbook"
    mAssert.Equals ws.Range("C1"), newProjectName & "_DEV", "Expected Title for development project column"
    mAssert.Equals ws.Range("C2"), "Project\" & newWorkbookName & "_DEV", "Expected related Path for new development workbook"
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

Public Sub TestGetModulesFromNewProject()
'       Verify the list of the modules of a new project
    mAssert.Equals newConfManager.moduleCount, 0, "There must be no module in a new project"
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
    newWorkbook.Close SaveChanges:=True, Filename:=vtkPathToTestFolder & "\" & "TestFile.xlsx"
    vtkResetConfigurationManagers           ' reset all configuration managers
    Set newWorkbook = Workbooks.Open(Filename:=vtkPathToTestFolder & "\" & "TestFile.xlsx")
    Set newConfManager = vtkConfigurationManagerForWorkbook("TestFile.xlsx")
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

' pour tous les tests de path, faire avec numéro, puis avec nom de la configuration
' Test Get configuration by name
' Test Get path by name
' Test Set path by name

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationSheetCreationForNewProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationSheetRetrievalForExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationSheetFormatForNewProjet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationSheetFormatForExistingProjet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetConfigurationsFromNewProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetModulesFromNewProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetConfigurationsFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetConfigurationPathWithNumberFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestSetConfigurationPathWithNumberToNewProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestSetConfigurationPathWithNumberToSavedProject"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestConfigurationSheetCreationForNewProject": TestConfigurationSheetCreationForNewProject
        Case "TestConfigurationSheetRetrievalForExistingProject": TestConfigurationSheetRetrievalForExistingProject
        Case "TestConfigurationSheetFormatForNewProjet": TestConfigurationSheetFormatForNewProjet
        Case "TestConfigurationSheetFormatForExistingProjet": TestConfigurationSheetFormatForExistingProjet
        Case "TestGetConfigurationsFromNewProject": TestGetConfigurationsFromNewProject
        Case "TestGetModulesFromNewProject": TestGetModulesFromNewProject
        Case "TestGetConfigurationsFromExistingProject": TestGetConfigurationsFromExistingProject
        Case "TestGetConfigurationPathWithNumberFromExistingProject": TestGetConfigurationPathWithNumberFromExistingProject
        Case "TestSetConfigurationPathWithNumberToNewProject": TestSetConfigurationPathWithNumberToNewProject
        Case "TestSetConfigurationPathWithNumberToSavedProject": TestSetConfigurationPathWithNumberToSavedProject
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

