VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RunManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IRunManager

Private mAssert As Assert
Private mAssertAsResultUser As IResultUser

Private Sub Class_Initialize()
    Set mAssert = New Assert
    Set mAssertAsResultUser = mAssert
    Set mAssertAsResultUser.Result = New TestResult
End Sub

Private Property Get IRunManager_Assert() As IAssert
    Set IRunManager_Assert = mAssert
End Property

Private Sub IRunManager_Report()
    Dim Failure As TestFailure, RM As IRunManager, test As ITest, TestCase As ITestCase
    If Result.WasSuccessful Then
        Debug.Print "OK (" & Result.TestCasesRun & ")"
    Else
        Debug.Print "Tests run: " & Result.TestCasesRun & " Failures: " & Result.Failures.Count
        For Each Failure In Result.Failures
            Set TestCase = Failure.TestCase
            Set test = TestCase
            Debug.Print test.Manager.ClassName & "." & TestCase.Manager.methodName & ": " & Failure.Comment
        Next
    End If
End Sub

Private Property Get IRunManager_Result() As TestResult
    Set IRunManager_Result = Result
End Property

Public Property Get Result() As TestResult
    Set Result = mAssertAsResultUser.Result
End Property
