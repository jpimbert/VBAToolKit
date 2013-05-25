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
End Sub

Private Sub ITestCase_TearDown()

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
    Dim wb As Workbook
    Set wb = createExcelProjectNamed("Test1")
    mAssert.Equals wb.VBProject.name, "Test1", "Expected name for Excel project"
    wb.Close (False)
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestUtilityFunctionCreateExcelProject"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestUtilityFunctionCreateExcelProject": TestUtilityFunctionCreateExcelProject
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

