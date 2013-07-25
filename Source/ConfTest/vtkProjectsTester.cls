VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkProjectsTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITest
Implements ITestCase

Private mManager As TestCaseManager
Private mAssert As IAssert

Private projectName As String      ' default workbook Name for test

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
    projectName = "TestProject"
End Sub

Private Sub ITestCase_TearDown()
End Sub

Public Sub TestProjectIsNotNull()
    ' Verify that the configuration lanager is not null
    mAssert.Should Not vtkProjectForName(projectName) Is Nothing, "A project must not be null"
End Sub

Public Sub TestProjectsAreDifferentForTwoNames()
    ' Verify that the configuration managers for two projects are different
    mAssert.Should Not vtkProjectForName(projectName) Is vtkProjectForName("SecondProject"), "Two projects must not be equal"
End Sub

Public Sub TestProjectIsInvariant()
    ' Verify that the configuration manager for a project is always the same
    Dim cm As vtkProject
    Set cm = vtkProjectForName(projectName)
    mAssert.Should vtkProjectForName(projectName) Is cm, "The project for a project name must be always the same"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectIsNotNull"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectsAreDifferentForTwoNames"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectIsInvariant"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestProjectIsNotNull": TestProjectIsNotNull
        Case "TestProjectsAreDifferentForTwoNames": TestProjectsAreDifferentForTwoNames
        Case "TestProjectIsInvariant": TestProjectIsInvariant
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub



