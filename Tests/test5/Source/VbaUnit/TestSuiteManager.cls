VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestSuiteManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITestManager
Private mClassName

Private Function ITestManager_CountTestCases(test As ITest) As Long
    Dim Count As Long
    Dim TestSuite As TestSuite, tests As Collection
    Dim SubTest As ITest
    Set TestSuite = test
    Set tests = TestSuite.tests()
    For Each SubTest In tests
        Count = Count + SubTest.Manager.CountTestCases(SubTest)
    Next
    ITestManager_CountTestCases = Count
End Function

Private Property Let ITestManager_ClassName(RHS As String)
    mClassName = RHS
End Property

Private Property Get ITestManager_ClassName() As String
    ITestManager_ClassName = mClassName
End Property

Private Function ITestManager_Run(test As ITest, Optional RunManager As IRunManager) As IRunManager
    Dim TestSuite As TestSuite, tests As Collection
    Dim SubTest As ITest
    Set TestSuite = test
    Set tests = TestSuite.tests()
    If RunManager Is Nothing Then Set RunManager = New RunManager
    For Each SubTest In TestSuite.tests()
        SubTest.Manager.Run SubTest, RunManager
    Next
    Set ITestManager_Run = RunManager
End Function
