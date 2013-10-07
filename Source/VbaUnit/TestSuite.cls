VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITest

Private mTests As Collection
Private mManager As TestSuiteManager

Public Function TESTS() As Collection
    Set TESTS = mTests
End Function

Public Sub AddTest(TestClassName As String, Optional methodName As String)
    Dim test As ITest
    Dim TestCase As ITestCase
    Dim TL As TestClassLister
    Set TL = New TestClassLister
    Set test = TL.NewTestClass(TestClassName)
    If methodName <> "" Then
        Set TestCase = test
        TestCase.Manager.init methodName
        mTests.Add test
    Else
        mTests.Add test.Suite
    End If
End Sub

Private Sub Class_Initialize()
    Set mTests = New Collection
    Set mManager = New TestSuiteManager
End Sub

Private Property Get ITest_Manager() As ITestManager
    Set ITest_Manager = mManager
End Property

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = Me
End Function
