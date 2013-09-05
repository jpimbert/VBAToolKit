VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkExportOneModuleTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkExportOneModuleTester
' Author    : Jean-Pierre Imbert
' Date      : 22/08/2013
' Purpose   : Test the ExportOneModule function
'
' Copyright 2013 Skwal-Soft (http://skwalsoft.com)
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.
'---------------------------------------------------------------------------------------

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
Private Const formToImport As String = "TestForm.frm"                             ' UserForm to import for test

Private wbForTest As Workbook, fso As New FileSystemObject

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
    FileCopy source:=vtkPathToTemplateFolder & "\" & formToImport, destination:=vtkPathToTestFolder & "\" & formToImport
    
    Set wbForTest = Workbooks.Open(fileName:=vtkPathToTestFolder & "\" & existingWorkbookName)
    
End Sub

'---------------------------------------------------------------------------------------
'   The existing workbook template contains modules
'     - Feuil1, Feuil2, Feuil3 (Sheets)
'     - ThisWorkbook (workbook)
'     - TestForm (UserForm)
'     - TestModule (Standard)
'     - TestClass (Class)
'   It exists an exported template for each (except Feuil2 and Feuil3)
'---------------------------------------------------------------------------------------

Private Sub ITestCase_TearDown()
    On Error Resume Next
    
    vtkCloseAndKillWorkbook Wb:=wbForTest

    Kill PathName:=vtkPathToTestFolder & "\" & existingWorkbookName
    Kill PathName:=vtkPathToTestFolder & "\" & moduleToImport
    Kill PathName:=vtkPathToTestFolder & "\" & classToImport
    Kill PathName:=vtkPathToTestFolder & "\" & sheetToImport
    Kill PathName:=vtkPathToTestFolder & "\" & workbookToImport
    Kill PathName:=vtkPathToTestFolder & "\" & formToImport
    Kill PathName:=vtkPathToTestFolder & "\" & "TestModule2.bas"
    Kill PathName:=vtkPathToTestFolder & "\" & "TestForm.frx"
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
'   Tests for exporting one Module. The following configurations have to be considered
'   A/ About the presence of the module to export
'       A1/ the module doesn't exist in the project
'       A2/ the module exists in the project
'   B/ About the presence of the file in which to export
'       B1/ the file doesn't exist
'       B2/ the file already exists
'   C/ About the type of module to export
'       C1/ standard module
'       C2/ class module
'       C3/ sheet code
'       C4/ workbook code
'       C5/ user form
'   The configuration to tests are :
'       - A1-B1     : Test_ModuleAndFileDoesntExist
'       - A1-B2     : Test_ModuleDoesntExistAndFileExists
'       - A2-B1-C1  : Test_StandardModuleExistsAndFileDoesntExist
'       - A2-B1-C2  : Test_ClassModuleExistsAndFileDoesntExist
'       - A2-B2-C1  : Test_StandardModuleExistsAndFileExists
'       - A2-B2-C3  : Test_SheetModuleExistsAndFileExists
'       - A2-B2-C5  : Test_FormModuleExistsAndFRMFileExists
'       -             Test_FormModuleExistsAndFRXFileExists
'---------------------------------------------------------------------------------------

Public Sub Test_ModuleAndFileDoesntExist()
    Dim filePath As String, moduleName As String
    
    filePath = "DummyFile.xxx"
    moduleName = "InexistantModule"
   
   On Error GoTo Test_Error
    vtkExportOneModule project:=wbForTest.VBProject, moduleName:=moduleName, filePath:=filePath
   On Error GoTo 0
   
Test_Error:
    mAssert.Equals Err.number, VTK_UNKNOWN_MODULE, "Export doesn't work when module is inexistant"
    mAssert.Should Not fso.FileExists(filePath), "File must not be created when exported module is inexistant"
End Sub

Public Sub Test_ModuleDoesntExistAndFileExists()
    Dim sizeBefore As Long, sizeAfter As Long, filePath As String, moduleName As String
   
    filePath = vtkPathToTestFolder & "\" & moduleToImport
    moduleName = "InexistantModule"
   
   On Error GoTo Test_Error
    sizeBefore = fso.GetFile(filePath).Size
    vtkExportOneModule project:=wbForTest.VBProject, moduleName:=moduleName, filePath:=filePath
   On Error GoTo 0
   
Test_Error:
    sizeAfter = fso.GetFile(filePath).Size
    mAssert.Equals Err.number, VTK_UNKNOWN_MODULE, "Export doesn't work when module is inexistant"
    mAssert.Should fso.FileExists(filePath), "File is not deleted when exported module is inexistant"
    mAssert.Equals sizeAfter, sizeBefore, "The size of the file must not be changed when exported module is inexistant"
End Sub

Public Sub Test_StandardModuleExistsAndFileDoesntExist()
    Dim existingFilePath As String, newFilePath As String, moduleName As String

    existingFilePath = vtkPathToTestFolder & "\" & "TestModule.bas"
    newFilePath = vtkPathToTestFolder & "\" & "TestModule2.bas"
    moduleName = "TestModule"

   On Error GoTo Test_Error
    vtkExportOneModule project:=wbForTest.VBProject, moduleName:=moduleName, filePath:=newFilePath
    mAssert.Should fso.FileExists(newFilePath), "File must be created when exported existant module"
    mAssert.Should Not (fso.GetFile(existingFilePath).Size = fso.GetFile(newFilePath).Size), "Size of Exported module must be different than the existing file"
   On Error GoTo 0
    Exit Sub
    
Test_Error:
    mAssert.Equals Err.number, 0, "No error must be thrown when Module exists"
End Sub

Public Sub Test_ClassModuleExistsAndFileDoesntExist()
    Dim existingFilePath As String, newFilePath As String, moduleName As String

    existingFilePath = vtkPathToTestFolder & "\" & "TestClass.cls"
    newFilePath = vtkPathToTestFolder & "\" & "TestClass.cls"
    moduleName = "TestClass"

   On Error GoTo Test_Error
    vtkExportOneModule project:=wbForTest.VBProject, moduleName:=moduleName, filePath:=newFilePath
    mAssert.Should fso.FileExists(newFilePath), "File must be created when exported existant module"
   On Error GoTo 0
    Exit Sub
    
Test_Error:
    mAssert.Equals Err.number, 0, "No error must be thrown when Module exists"
End Sub

Public Sub Test_StandardModuleExistsAndFileExists()
    Dim sizeBefore As Long, sizeAfter As Long, filePath As String, moduleName As String

    filePath = vtkPathToTestFolder & "\" & "TestModule.bas"
    moduleName = "TestModule"

   On Error GoTo Test_Error
    sizeBefore = fso.GetFile(filePath).Size
    vtkExportOneModule project:=wbForTest.VBProject, moduleName:=moduleName, filePath:=filePath
    sizeAfter = fso.GetFile(filePath).Size
    mAssert.Should fso.FileExists(filePath), "File must exist when exported existant module"
    mAssert.Should sizeAfter <> sizeBefore, "Size of new file (" & sizeAfter & ") must be different than size of old file (" & sizeBefore & ")"
   On Error GoTo 0
    Exit Sub
    
Test_Error:
    mAssert.Equals Err.number, 0, "No error must be thrown when Module exists"
End Sub

Public Sub Test_SheetModuleExistsAndFileExists()
    Dim sizeBefore As Long, sizeAfter As Long, filePath As String, moduleName As String

    filePath = vtkPathToTestFolder & "\" & "Feuil1.cls"
    moduleName = "Feuil1"

   On Error GoTo Test_Error
    sizeBefore = fso.GetFile(filePath).Size
    vtkExportOneModule project:=wbForTest.VBProject, moduleName:=moduleName, filePath:=filePath
    sizeAfter = fso.GetFile(filePath).Size
    mAssert.Should fso.FileExists(filePath), "File must exist when exported existant module"
    mAssert.Should sizeAfter <> sizeBefore, "Size of new file (" & sizeAfter & ") must be different than size of old file (" & sizeBefore & ")"
   On Error GoTo 0
    Exit Sub
    
Test_Error:
    mAssert.Equals Err.number, 0, "No error must be thrown when Module exists"
End Sub

Public Sub Test_FormModuleExistsAndFRMFileExists()
    Dim sizeBefore As Long, sizeAfter As Long, filePath As String, moduleName As String, secondFilePath As String

    filePath = vtkPathToTestFolder & "\" & "TestForm.frm"
    secondFilePath = vtkPathToTestFolder & "\" & "TestForm.frx"
    moduleName = "TestForm"

   On Error GoTo Test_Error
    sizeBefore = fso.GetFile(filePath).Size
    vtkExportOneModule project:=wbForTest.VBProject, moduleName:=moduleName, filePath:=filePath
    sizeAfter = fso.GetFile(filePath).Size
    mAssert.Should fso.FileExists(filePath), "File must exist when exported existant module"
    mAssert.Should fso.FileExists(secondFilePath), "File must exist when exported existant module"
    mAssert.Should sizeAfter <> sizeBefore, "Size of new file (" & sizeAfter & ") must be different than size of old file (" & sizeBefore & ")"
   On Error GoTo 0
    Exit Sub
    
Test_Error:
    mAssert.Equals Err.number, 0, "No error must be thrown when Module exists"
End Sub

Public Sub Test_FormModuleExistsAndFRXFileExists()
    Dim sizeBefore As Long, sizeAfter As Long, filePath As String, moduleName As String, secondFilePath As String

    filePath = vtkPathToTestFolder & "\" & "TestForm.frm"
    secondFilePath = vtkPathToTestFolder & "\" & "TestForm.frx"
    moduleName = "TestForm"
    
    fso.CreateTextFile secondFilePath

   On Error GoTo Test_Error
    sizeBefore = fso.GetFile(secondFilePath).Size
    vtkExportOneModule project:=wbForTest.VBProject, moduleName:=moduleName, filePath:=filePath
    sizeAfter = fso.GetFile(secondFilePath).Size
    mAssert.Should fso.FileExists(filePath), "File must exist when exported existant module"
    mAssert.Should fso.FileExists(secondFilePath), "File must exist when exported existant module"
    mAssert.Should sizeAfter <> sizeBefore, "Size of new file (" & sizeAfter & ") must be different than size of old file (" & sizeBefore & ")"
   On Error GoTo 0
    Exit Sub
    
Test_Error:
    mAssert.Equals Err.number, 0, "No error must be thrown when Module exists"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ModuleAndFileDoesntExist"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ModuleDoesntExistAndFileExists"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_StandardModuleExistsAndFileDoesntExist"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ClassModuleExistsAndFileDoesntExist"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_StandardModuleExistsAndFileExists"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_SheetModuleExistsAndFileExists"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_FormModuleExistsAndFRMFileExists"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_FormModuleExistsAndFRXFileExists"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_ModuleAndFileDoesntExist": Test_ModuleAndFileDoesntExist
        Case "Test_ModuleDoesntExistAndFileExists": Test_ModuleDoesntExistAndFileExists
        Case "Test_StandardModuleExistsAndFileDoesntExist": Test_StandardModuleExistsAndFileDoesntExist
        Case "Test_ClassModuleExistsAndFileDoesntExist": Test_ClassModuleExistsAndFileDoesntExist
        Case "Test_StandardModuleExistsAndFileExists": Test_StandardModuleExistsAndFileExists
        Case "Test_SheetModuleExistsAndFileExists": Test_SheetModuleExistsAndFileExists
        Case "Test_FormModuleExistsAndFRMFileExists": Test_FormModuleExistsAndFRMFileExists
        Case "Test_FormModuleExistsAndFRXFileExists": Test_FormModuleExistsAndFRXFileExists
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

