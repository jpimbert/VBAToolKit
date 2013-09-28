VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestFailure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mTestCase As ITestCase, mComment As String

Public Sub init(TestCase As ITestCase, Comment As String)
    Set mTestCase = TestCase
    mComment = Comment
End Sub

Public Property Get TestCase() As ITestCase
    Set TestCase = mTestCase
End Property

Public Property Get Comment() As String
    Comment = mComment
End Property
