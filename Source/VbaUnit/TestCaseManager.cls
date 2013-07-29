VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCaseManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestManager

Private mMethodName As String, mClassName As String

Public Sub Init(methodName As String)
    mMethodName = methodName
End Sub

Public Property Get methodName() As String
    methodName = mMethodName
End Property

Private Property Let ITestManager_ClassName(RHS As String)
    mClassName = RHS
End Property

Private Property Get ITestManager_ClassName() As String
    ITestManager_ClassName = mClassName
End Property

Private Function ITestManager_CountTestCases(test As ITest) As Long
    ITestManager_CountTestCases = 1
End Function

Private Function ITestManager_Run(test As ITest, Optional RunManager As IRunManager) As IRunManager
    Dim TestCase As ITestCase
    If RunManager Is Nothing Then Set RunManager = New RunManager
    Set TestCase = test
    RunManager.Result.StartTest TestCase
    TestCase.SetUp RunManager.Assert
    TestCase.RunTest
    'debug.print "Run Runtest"
    TestCase.TearDown
    'debug.print "Run teardown"
    RunManager.Result.EndTest TestCase
    'debug.print "Run endtest"
End Function
