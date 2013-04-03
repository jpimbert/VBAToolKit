VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private mAssert As Assert
Private mCurrentTestCase As ITestCase
Private mTestCasesRun As Long
Private mFailures As Collection

Public Property Get CurrentTestCase() As ITestCase
    Set CurrentTestCase = mCurrentTestCase
End Property

Public Sub StartTest(TestCase As ITestCase)
    Set mCurrentTestCase = TestCase
    mTestCasesRun = mTestCasesRun + 1
End Sub

Public Property Get TestCasesRun() As Long
    TestCasesRun = mTestCasesRun
End Property

Public Sub EndTest(TestCase As ITestCase)
    Set mCurrentTestCase = Nothing
End Sub

Public Sub AddFailure(TestCase As ITestCase, Comment As String)
    Dim Failure As TestFailure
    Set Failure = New TestFailure
    Failure.Init TestCase, Comment
    Failures.Add Failure
End Sub

Public Property Get Failures() As Collection
    Set Failures = mFailures
End Property

Private Sub Class_Initialize()
    Set mFailures = New Collection
End Sub

Public Property Get WasSuccessful() As Boolean
    WasSuccessful = Failures.Count = 0
End Property
