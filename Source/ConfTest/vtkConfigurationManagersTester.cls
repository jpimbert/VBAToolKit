VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkConfigurationManagersTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITest
Implements ITestCase

Private mManager As TestCaseManager
Private mAssert As IAssert

Private wbForTest As Workbook       ' Default test workbook
Private workbookName As String      ' Default name of the test workbook
Private Const projectName As String = "TestProject"     ' Default name of the test project

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
    ' Create the default Excel project for tests
    Set wbForTest = vtkCreateExcelWorkbookForTestWithProjectName(projectName)
    workbookName = wbForTest.name
End Sub

Private Sub ITestCase_TearDown()
    Dim name As String
    vtkCloseAndKillWorkbook wb:=wbForTest
    ' Reset all configuration managers
    vtkResetConfigurationManagers
End Sub

Public Sub TestUtilityFunctionCreateExcelProject()
    ' Verify that the utility function createExcelProjectNamed correctly creates a new Excel file
    ' with proper project name
    mAssert.Equals wbForTest.name, workbookName, "Expected name for Excel project"
End Sub

Public Sub TestConfigurationManagerIsNotNull()
    ' Verify that the configuration Manager is not null
    mAssert.Should Not vtkConfigurationManagerForProject(projectName) Is Nothing, "A configuartion manager must not be null"
End Sub

Public Sub TestConfigurationManagersAreDifferentForTwoProjects()
    ' Verify that the configuration managers for two projects are different
    Dim secondWb As Workbook
    ' Create a second Excel project for test
    Set secondWb = vtkCreateExcelWorkbookForTestWithProjectName("SecondTest")
    mAssert.Should Not vtkConfigurationManagerForProject(projectName) Is vtkConfigurationManagerForProject("SecondTest"), "Two configuration managers must not be equal"
    vtkCloseAndKillWorkbook wb:=secondWb
End Sub

Public Sub TestConfigurationManagerIsInvariant()
    ' Verify that the configuration manager for a project is always the same
    Dim cm As vtkConfigurationManager
    Set cm = vtkConfigurationManagerForProject(projectName)
    mAssert.Should vtkConfigurationManagerForProject(projectName) Is cm, "The configuration manager for a project must be always the same"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestUtilityFunctionCreateExcelProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationManagerIsNotNull"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationManagersAreDifferentForTwoProjects"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationManagerIsInvariant"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestUtilityFunctionCreateExcelProject": TestUtilityFunctionCreateExcelProject
        Case "TestConfigurationManagerIsNotNull": TestConfigurationManagerIsNotNull
        Case "TestConfigurationManagersAreDifferentForTwoProjects": TestConfigurationManagersAreDifferentForTwoProjects
        Case "TestConfigurationManagerIsInvariant": TestConfigurationManagerIsInvariant
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


