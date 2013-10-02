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
    Set mAssertAsResultUser.result = New TestResult
End Sub

Private Property Get IRunManager_Assert() As IAssert
    Set IRunManager_Assert = mAssert
End Property

Private Sub IRunManager_Report()
    Dim Failure As TestFailure, RM As IRunManager, test As ITest, TestCase As ITestCase
    If result.WasSuccessful Then
        Debug.Print "OK (" & result.TestCasesRun & ")"
    Else
        Debug.Print "Tests run: " & result.TestCasesRun & " Failures: " & result.Failures.Count
        For Each Failure In result.Failures
            Set TestCase = Failure.TestCase
            Set test = TestCase
            Debug.Print test.Manager.ClassName & "." & TestCase.Manager.methodName & ": " & Failure.Comment
        Next
    End If
End Sub

Private Property Get IRunManager_Result() As TestResult
    Set IRunManager_Result = result
End Property

Public Property Get result() As TestResult
    Set result = mAssertAsResultUser.result
End Property
