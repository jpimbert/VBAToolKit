VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IAssert
Implements IResultUser

Private mResult As TestResult

Private Sub AddFailure(TestCase As ITestCase, Comment As String)
    mResult.AddFailure TestCase, Comment
    'Debug.Print "Test failed: " & Comment
End Sub

Private Sub IAssert_Delta(Actual As Variant, expected As Variant, Delta As Variant, Optional Comment As String)

End Sub

Private Sub IAssert_Equals(Actual As Variant, expected As Variant, Optional Comment As String)
    If Actual <> expected Then AddFailure CurrentTestCase, NotEqualsComment(Comment, Actual, expected)
End Sub

Private Sub IAssert_Should(Condition As Boolean, Optional Comment As String)
    If Not Condition Then AddFailure CurrentTestCase, Comment
End Sub

Private Property Get CurrentTestCase() As ITestCase
    Set CurrentTestCase = mResult.CurrentTestCase
End Property

Private Property Set IResultUser_Result(RHS As TestResult)
    Set mResult = RHS
End Property

Private Property Get IResultUser_Result() As TestResult
    Set IResultUser_Result = mResult
End Property

Private Function NotEqualsComment(Comment As String, Actual As Variant, expected As Variant) As String
    NotEqualsComment = Comment & ":" & vbCrLf & "expected: " & vbCrLf & QW(CStr(expected)) & vbCrLf & "but was: " & vbCrLf & QW(CStr(Actual))
End Function
