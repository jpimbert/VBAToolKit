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

Private wbForTest As Workbook       '   default WorkBook for test
Private Const projectName As String = "projectTest" ' default project Name for test

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
    Set wbForTest = createExcelProjectNamed(projectName)    ' create the default Excel project for tests
End Sub

Private Sub ITestCase_TearDown()
    wbForTest.Close SaveChanges:=False      ' close the default Excel project
End Sub

'---------------------------------------------------------------------------------------
' Procedure : createExcelProjectNamed
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Utility function for Excel project creation with a given project name
'---------------------------------------------------------------------------------------
'
Private Function createExcelProjectNamed(projectName As String) As Workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.VBProject.name = projectName
    Set createExcelProjectNamed = wb
End Function

Public Sub TestUtilityFunctionCreateExcelProject()
'       Verify that the utility function createExcelProjectNamed correctly create a new Excel file with proper project name
    mAssert.Equals wbForTest.VBProject.name, projectName, "Expected name for Excel project"
End Sub

Public Sub TestConfigurationManagerIsNotNull()
'       Verify that the configuration Manager is not null
    mAssert.Should Not configurationManagerForProject(projectName) Is Nothing, "A configuartion manager must not be null"
End Sub

Public Sub TestConfigurationManagersAreDifferentForTwoProjects()
'       Verify that the configuration Managers for two project are different
    Dim secondWb As Workbook
    Set secondWb = createExcelProjectNamed("SecondProject")    ' create the default Excel project for tests
    mAssert.Should Not configurationManagerForProject(projectName) Is configurationManagerForProject("SecondProject"), "Two configuration managers must not be equal"
    secondWb.Close SaveChanges:=False
End Sub

Public Sub TestConfigurationManagerIsInvariant()
'       Verify that the configuration Manager for a project is always the same
    Dim cm As vtkConfigurationManager
    Set cm = configurationManagerForProject(projectName)
    mAssert.Should configurationManagerForProject(projectName) Is cm, "The configuration manager for a project must be always the same"
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

