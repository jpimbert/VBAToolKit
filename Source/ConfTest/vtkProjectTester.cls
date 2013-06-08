VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkProjectTester"
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
Private project As vtkProject      ' Default vtkProject for test

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
    Set project = vtkProjectForName(projectName)
End Sub

Private Sub ITestCase_TearDown()
End Sub

Public Sub TestProjectNameForDevelopment()
'       Verify the project name for development
    mAssert.Equals project.projectDEVName, "TestProject_DEV", "The project development is suffixed by _DEV"
End Sub

Public Sub TestProjectNameForDevelopmentWithNullProjectName()
'       Verify that the project name for development is null when project name is null
    Dim p As New vtkProject
    mAssert.Equals p.projectDEVName, "", "The project development name mus be null"
End Sub

Public Sub TestProjectDevelopmentRelativePath()
'       Verify the project path for development
    mAssert.Equals project.projectDEVStandardRelativePath, "Project\" & "TestProject_DEV.xlsm", "The project development standard path"
End Sub

Public Sub TestProjectRelativePath()
'       Verify the project path
    mAssert.Equals project.projectStandardRelativePath, "Delivery\" & "TestProject.xlsm", "The project development standard path"
End Sub

Public Sub TestProjectDevelopmentWorkbookName()
'       Verify the project path for development
    mAssert.Equals project.workbookDEVName, "TestProject_DEV.xlsm", "The project development standard path"
End Sub

Public Sub TestProjectWorkbookName()
'       Verify the project path
    mAssert.Equals project.workbookName, "TestProject.xlsm", "The project development standard path"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectNameForDevelopment"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectNameForDevelopmentWithNullProjectName"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectDevelopmentRelativePath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectRelativePath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectDevelopmentWorkbookName"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectWorkbookName"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestProjectNameForDevelopment": TestProjectNameForDevelopment
        Case "TestProjectNameForDevelopmentWithNullProjectName": TestProjectNameForDevelopmentWithNullProjectName
        Case "TestProjectDevelopmentRelativePath": TestProjectDevelopmentRelativePath
        Case "TestProjectRelativePath": TestProjectRelativePath
        Case "TestProjectDevelopmentWorkbookName": TestProjectDevelopmentWorkbookName
        Case "TestProjectWorkbookName": TestProjectWorkbookName
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub




