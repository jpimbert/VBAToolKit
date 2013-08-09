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

Private Const existingWorkbookName As String = "WorkBookForConfigurationsTests.xlsm"     ' workbook Name for test
Private Const moduleToImport As String = "TestModule.bas"                             ' Module to import for test
Private Const classToImport As String = "TestClass.cls"                             ' Class to import for test
Private Const sheetToImport As String = "Feuil1.cls"                             ' Sheet to import for test
Private Const workbookToImport As String = "ThisWorkbook.cls"                             ' Workbook to import for test

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
    
    FileCopy source:=vtkPathToTemplateFolder & "\" & existingWorkbookName, destination:=vtkPathToTestFolder & "\" & existingWorkbookName
    FileCopy source:=vtkPathToTemplateFolder & "\" & moduleToImport, destination:=vtkPathToTestFolder & "\" & moduleToImport
    FileCopy source:=vtkPathToTemplateFolder & "\" & classToImport, destination:=vtkPathToTestFolder & "\" & classToImport
    FileCopy source:=vtkPathToTemplateFolder & "\" & sheetToImport, destination:=vtkPathToTestFolder & "\" & sheetToImport
    FileCopy source:=vtkPathToTemplateFolder & "\" & workbookToImport, destination:=vtkPathToTestFolder & "\" & workbookToImport
End Sub

Private Sub ITestCase_TearDown()
    On Error Resume Next
    Kill PathName:=vtkPathToTestFolder & "\" & existingWorkbookName
    Kill PathName:=vtkPathToTestFolder & "\" & moduleToImport
    Kill PathName:=vtkPathToTestFolder & "\" & classToImport
    Kill PathName:=vtkPathToTestFolder & "\" & sheetToImport
    Kill PathName:=vtkPathToTestFolder & "\" & workbookToImport
    On Error GoTo 0
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

'
'   Tests for importing one Module. The following configurations have to be considered
'   A/ About the presence of the module to import
'       A1/ the imported module is a new one in the project
'       A2/ the file to import doesn't exist
'       A3/ the imported module has to replace an existing module
'   B/ About the type of module to import
'       B1/ standard module
'       B2/ class module
'       B3/ sheet code
'       B4/ workbook code
'       B5/ user form
'   The configuration to tests are :
'       - B1 and B5 with A1, A2 and A3
'       - B2, B3 and B4 with A3
'

Public Sub TestImportInexistantFileDoesNothing()
    '   Verify that an existing Module isn't deleted by a bad import (A2-B1)
    Dim wbForTest As Workbook
    Set wbForTest = Workbooks.Open(Filename:=vtkPathToTestFolder & "\" & existingWorkbookName)
    
    vtkImportOneModule project:=wbForTest.VBProject, moduleName:="TestModule", filePath:="DummyFile.xxx"
    
    mAssert.Equals wbForTest.VBProject.VBComponents("TestModule").CodeModule.CountOfLines, 5, "The TestModule of the Test Workbook must not have been modified. It contains 5 code lines"
    vtkCloseAndKillWorkbook wb:=wbForTest
End Sub

Public Sub TestImportExistantFileInExistingModule()
    '   Verify that the existing Module replaced by the new one (A3-B1)
    Dim wbForTest As Workbook
    Set wbForTest = Workbooks.Open(Filename:=vtkPathToTestFolder & "\" & existingWorkbookName)
    
    vtkImportOneModule project:=wbForTest.VBProject, moduleName:="TestModule", filePath:=vtkPathToTestFolder & "\" & moduleToImport
    
    mAssert.Equals wbForTest.VBProject.VBComponents("TestModule").CodeModule.CountOfLines, 8, "The TestModule of the Test Workbook must not have been modified. It contains now 8 code lines"
    vtkCloseAndKillWorkbook wb:=wbForTest
End Sub

Public Sub TestImportExistantFileInExistingModuleWithNewName()
    '   Verify that the new Module doesn't replace the old one (A1-B1)
    Dim wbForTest As Workbook
    Set wbForTest = Workbooks.Open(Filename:=vtkPathToTestFolder & "\" & existingWorkbookName)
    
    vtkImportOneModule project:=wbForTest.VBProject, moduleName:="NewModule", filePath:=vtkPathToTestFolder & "\" & moduleToImport
    
    mAssert.Equals wbForTest.VBProject.VBComponents("TestModule").CodeModule.CountOfLines, 5, "The TestModule of the Test Workbook must not have been modified. It contains 5 code lines"
    mAssert.Equals wbForTest.VBProject.VBComponents("NewModule").CodeModule.CountOfLines, 8, "The NewModule had to be added to the test Workbook. It contains 8 code lines"
    vtkCloseAndKillWorkbook wb:=wbForTest
End Sub

Public Sub TestImportExistantFileInExistingClassModule()
    '   Verify that the existing Module replaced by the new one (A3-B2)
    Dim wbForTest As Workbook
    Set wbForTest = Workbooks.Open(Filename:=vtkPathToTestFolder & "\" & existingWorkbookName)
    
    vtkImportOneModule project:=wbForTest.VBProject, moduleName:="TestClass", filePath:=vtkPathToTestFolder & "\" & classToImport
    
    mAssert.Equals wbForTest.VBProject.VBComponents("TestClass").CodeModule.CountOfLines, 122, "The TestClass of the Test Workbook must have been modified. It must contain 122 code lines"
    vtkCloseAndKillWorkbook wb:=wbForTest
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestNumberOfVBAUnitModules"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestVBAUnitMainIsTheFirstVBAUnitModule"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestIAssertIsInVBAUnitModulesList"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestVBAUnitStandardCategory"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestTestStandardCategory"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProdStandardCategory"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestImportInexistantFileDoesNothing"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestImportExistantFileInExistingModule"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestImportExistantFileInExistingModuleWithNewName"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestImportExistantFileInExistingClassModule"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestNumberOfVBAUnitModules": TestNumberOfVBAUnitModules
        Case "TestVBAUnitMainIsTheFirstVBAUnitModule": TestVBAUnitMainIsTheFirstVBAUnitModule
        Case "TestIAssertIsInVBAUnitModulesList": TestIAssertIsInVBAUnitModulesList
        Case "TestVBAUnitStandardCategory": TestVBAUnitStandardCategory
        Case "TestTestStandardCategory": TestTestStandardCategory
        Case "TestProdStandardCategory": TestProdStandardCategory
        Case "TestImportInexistantFileDoesNothing": TestImportInexistantFileDoesNothing
        Case "TestImportExistantFileInExistingModule": TestImportExistantFileInExistingModule
        Case "TestImportExistantFileInExistingModuleWithNewName": TestImportExistantFileInExistingModuleWithNewName
        Case "TestImportExistantFileInExistingClassModule": TestImportExistantFileInExistingClassModule
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

