VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RectangleTesterDisabled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITest
Implements ITestCase

Private mManager As TestCaseManager
Private mAssert As IAssert

Private mRectangle As RectangleClass

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
    Set mRectangle = New RectangleClass
    mRectangle.Init 2, 3
End Sub

Private Sub ITestCase_TearDown()

End Sub

Public Sub TestArea()
    mAssert.Should mRectangle.Area() = 67, "Area"
End Sub

Public Sub TestPerimeter()
    mAssert.Should mRectangle.Perimeter() = 107, "Perimeter"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestArea"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestPerimeter"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestArea": TestArea
        Case "TestPerimeter": TestPerimeter
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub
