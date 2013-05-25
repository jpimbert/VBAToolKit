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
End Sub

'   Liste des tests à faire
'       Par défaut deux configurations sont créées dans un projet
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

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationSheetCreationForNewProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationSheetRetrievalForExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationSheetFormatForNewProjet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationSheetFormatForExistingProjet"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestConfigurationSheetCreationForNewProject": TestConfigurationSheetCreationForNewProject
        Case "TestConfigurationSheetRetrievalForExistingProject": TestConfigurationSheetRetrievalForExistingProject
        Case "TestConfigurationSheetFormatForNewProjet": TestConfigurationSheetFormatForNewProjet
        Case "TestConfigurationSheetFormatForExistingProjet": TestConfigurationSheetFormatForExistingProjet
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

