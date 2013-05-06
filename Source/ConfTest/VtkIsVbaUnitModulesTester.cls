VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VtkIsVbaUnitModulesTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

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

End Sub

Public Sub TestVbaUnitClassName()
'a very simple test
    mAssert.Equals a = VtkExistVbaUnit("assert"), False, "this module name must exist"
    mAssert.Equals a = VtkExistVbaUnit("ASSERT"), False, "this module name must exist"
    mAssert.Equals a = VtkExistVbaUnit("Test"), True, "this module name must not exist"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestVbaUnitClassName"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestVbaUnitClassName": TestVbaUnitClassName
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


