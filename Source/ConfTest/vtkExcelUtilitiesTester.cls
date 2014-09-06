VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkExcelUtilitiesTester"
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
    VBAToolKit.resetTestFolder
End Sub

Public Sub Test_vtkDefaultIsAddIn()
    mAssert.Should vtkDefaultIsAddIn("TestFile.xla"), "xla extension is Add-In"
    mAssert.Should vtkDefaultIsAddIn("TestFile.xlam"), "xlam extension is Add-In"
    mAssert.Should Not vtkDefaultIsAddIn("TestFile.xlsx"), "xlsx extension is not Add-In"
    mAssert.Should Not vtkDefaultIsAddIn("TestFile.xltx"), "xltx extension is not Add-In"
    mAssert.Should Not vtkDefaultIsAddIn("TestFile.xlsm"), "xlsm extension is not Add-In"
    mAssert.Should Not vtkDefaultIsAddIn("TestFile.xltm"), "xltm extension is not Add-In"
    mAssert.Should Not vtkDefaultIsAddIn("TestFile.xls"), "xls extension is not Add-In"
End Sub

Public Sub Test_vtkDefaultExcelExtension()
    mAssert.Equals ".xlsm", vtkDefaultExcelExtension, "Default Excel 2007 extension"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDefaultIsAddIn"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDefaultExcelExtension"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkDefaultIsAddIn": Test_vtkDefaultIsAddIn
        Case "Test_vtkDefaultExcelExtension": Test_vtkDefaultExcelExtension
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

