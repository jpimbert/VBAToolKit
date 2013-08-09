VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkToolBarsTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITest
Implements ITestCase

Private mManager As TestCaseManager
Private mAssert As IAssert

Private Const toolBarName = "VBAToolKit_DEV_Toolbar"

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
    vtkDeleteToolbars   ' The test toolbars have to be deleted after Unit Test ran
End Sub

Public Sub TestExcelToolBarCreation()
    '   Verify that the Excel toolbar exists after creation
    Dim bar As CommandBar
    
    vtkCreateToolbars
    
    On Error Resume Next
    Set bar = Application.CommandBars(toolBarName)
    On Error GoTo 0
    
    mAssert.Should Not bar Is Nothing, "The Excel Toolbar must exist"
End Sub

Public Sub TestVBEToolBarCreation()
    '   Verify that the VBE toolbar exists after creation
    Dim bar As CommandBar
    
    vtkCreateToolbars
    
    On Error Resume Next
    Set bar = Application.VBE.CommandBars(toolBarName)
    On Error GoTo 0
    
    mAssert.Should Not bar Is Nothing, "The VBE Toolbar must exist"
End Sub

Public Sub TestExcelToolBarDeletion()
    '   Verify that the Excel toolbar doesn't exist after deletion
    Dim bar As CommandBar
    
    vtkCreateToolbars
    vtkDeleteToolbars
    
    On Error Resume Next
    Set bar = Application.CommandBars(toolBarName)
    On Error GoTo 0
    
    mAssert.Should bar Is Nothing, "The Excel Toolbar must not exist"
End Sub

Public Sub TestVBEToolBarDeletion()
    '   Verify that the VBE toolbar doesn't exist after deletion
    Dim bar As CommandBar
    
    vtkCreateToolbars
    vtkDeleteToolbars
    
    On Error Resume Next
    Set bar = Application.VBE.CommandBars(toolBarName)
    On Error GoTo 0
    
    mAssert.Should bar Is Nothing, "The VBE Toolbar must not exist"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestExcelToolBarCreation"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestVBEToolBarCreation"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestExcelToolBarDeletion"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestVBEToolBarDeletion"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestExcelToolBarCreation": TestExcelToolBarCreation
        Case "TestVBEToolBarCreation": TestVBEToolBarCreation
        Case "TestExcelToolBarDeletion": TestExcelToolBarDeletion
        Case "TestVBEToolBarDeletion": TestVBEToolBarDeletion
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

