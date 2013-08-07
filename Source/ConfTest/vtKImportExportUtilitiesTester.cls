VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtKImportExportUtilitiesTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
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

Public Sub TestNumberOfVBAUnitModules()
    ' Verify the correct number of VBAUnit Modules in the list
    mAssert.Equals vtkVBAUnitModulesList.Count, 18, "There are 18 VBAUnit Modules"
End Sub

Public Sub TestVBAUnitMainIsTheFirstVBAUnitModule()
    ' Verify that VBAUnitMain is the first VBAUnit Module in the list
    '   The purpose is just to access the list by Item number
    mAssert.Equals vtkVBAUnitModulesList.Item(1), "VbaUnitMain", "VBAUnitMain must be the first VBAUnit Module"
End Sub

Public Sub TestIAssertIsInVBAUnitModulesList()
    ' Verify that IAssert is included in the VBAUnit Module list
    '   The purpose is just to access the list by Item name
    mAssert.Equals vtkVBAUnitModulesList.Item("IAssert"), "IAssert", "IAssert must be included in the list"
End Sub

Public Sub TestVBAUnitStandardCategory()
    ' Verify that the category of a VBAUnit Module is "VBAUnit"
    mAssert.Equals vtkStandardCategoryForModuleName("TesterTemplate"), "VBAUnit", "TesterTemplate is a VBAUnit Module"
End Sub

Public Sub TestTestStandardCategory()
    ' Verify that the category of a Module ending with Tester is "Test"
    mAssert.Equals vtkStandardCategoryForModuleName("TemplateTester"), "Test", "TemplateTester is a Test Module"
End Sub

Public Sub TestProdStandardCategory()
    ' Verify that the category of a standard Module is "Prod"
    mAssert.Equals vtkStandardCategoryForModuleName("Template"), "Prod", "Template is a standard Module"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestNumberOfVBAUnitModules"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestVBAUnitMainIsTheFirstVBAUnitModule"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestIAssertIsInVBAUnitModulesList"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestVBAUnitStandardCategory"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestTestStandardCategory"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProdStandardCategory"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestNumberOfVBAUnitModules": TestNumberOfVBAUnitModules
        Case "TestVBAUnitMainIsTheFirstVBAUnitModule": TestVBAUnitMainIsTheFirstVBAUnitModule
        Case "TestIAssertIsInVBAUnitModulesList": TestIAssertIsInVBAUnitModulesList
        Case "TestVBAUnitStandardCategory": TestVBAUnitStandardCategory
        Case "TestTestStandardCategory": TestTestStandardCategory
        Case "TestProdStandardCategory": TestProdStandardCategory
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

