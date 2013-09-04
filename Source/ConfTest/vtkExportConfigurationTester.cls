VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkExportConfigurationTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkExportConfigurationTester
' Author    : Lucas Vitorino
' Purpose   : Test the vtkExportConfiguration function.
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

Private Const testWorkbookName As String = "TestProject_DEV.xlsm"
Private Const testProjectName As String = "TestProject"
Private Const testModule1FileName As String = "TestModule1.bas"
Private Const testModule2FileName As String = "TestModule2.bas"
Private Const importedModuleFileName As String = "TestModule.bas"
Private Const testModule1Name As String = "TestModule1"
Private Const testModule2Name As String = "TestModule2"
Private Const importedModuleName As String = "TestModule"
Private Const totalNumberOfModules As Integer = 2

Private cm As vtkConfigurationManager
Private testConf As vtkConfiguration

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
    On Error Resume Next

    Set mAssert = Assert
    
    FileCopy source:=vtkPathToTemplateFolder & "\" & testWorkbookName, destination:=vtkPathToTestFolder & "\" & testWorkbookName

    Set wbForTest = Workbooks.Open(fileName:=vtkPathToTestFolder & "\" & testWorkbookName)
    
    vtkResetConfigurationManagers
    Set cm = vtkConfigurationManagerForProject(testProjectName)
    Set testConf = cm.configurations(wbForTest.VBProject.name)

    On Error GoTo 0
    Exit Sub
    
End Sub

Private Sub ITestCase_TearDown()
    
    vtkCloseAndKillWorkbook Wb:=wbForTest
    vtkCleanFolder vtkPathToTestFolder

End Sub

' Quick procedure to modify a module
Private Sub insertDummyProcedureInCodeModule(codemo As VBIDE.CodeModule, dummyInt As Integer)
    Dim dummyProcedure As String
    
    dummyProcedure = _
    "Public Sub dummyProcedure" & dummyInt & "()" & vbNewLine & _
    "End Sub" & vbNewLine
    
    With codemo
        .InsertLines .CountOfLines + 1, dummyProcedure
    End With

End Sub

' -------------------------------------------------------------------------
'                               TESTS
' -------------------------------------------------------------------------
'
' - Source files are not yet created, onlyModified = True
'    - expected : all modules are exported
'    - Test_vtkExportConfiguration_filesNotYetCreatedTrue
'
' - Source files are not yet created, onlyModified = False
'    - expected : all modules are exported
'    - Test_vtkExportConfiguration_filesNotYetCreatedFalse
'
' - Source files are not yet created, onlyModified = False
'    - expected : testModule1FileName has been created
'    - Test_vtkExportConfiguration_filesNotYetCreatedFalseFilesAreCreated1
'
' - Source files are not yet created, onlyModified = False
'    - expected : testModule2FileName has been created
'    - Test_vtkExportConfiguration_filesNotYetCreatedFalseFilesAreCreated2
'
' - Source files are created, onlyModified = True
'    - expected : 0 modules are exported
'    - Test_vtkExportConfiguration_filesJustExportedTrue
'
' - Source files are created, onlyModified = False
'    - expected : all modules are exported
'    - Test_vtkExportConfiguration_filesJustExportedFalse
'
' - Source files are exported, 1 module is modified, onlyModified = True
'    - expected : 1 module is exported
'    - Test_vtkExportConfiguration_oneFileModifiedTrue
'
' - Source files are exported, 1 module is modified, onlyModified = False
'    - expected : all modules are exported
'    - Test_vtkExportConfiguration_oneFileModifiedFalse
'
' - Source files are exported, 2 modules are modified, onlyModified = True
'    - expected : 2 modules are exported
'    - Test_vtkExportConfiguration_twoFilesModifiedTrue
'
' - Source files are exported, 2 modules are modified, onlyModified = False
'    - expected : all modules are exported
'    - Test_vtkExportConfiguration_twoFilesModifiedFalse

' - Source files are exported, 1 module is imported, onlyModified = True
'    - expected : 1 module is exported
'    - Test_vtkExportConfiguration_importedModuleMustBeExported
'
' -------------------------------------------------------------------------
Public Sub Test_vtkExportConfiguration_filesNotYetCreatedTrue()
    Dim retVal As Integer
    
    On Error GoTo Test_vtkExportConfiguration_filesNotYetCreatedTrue_Error

    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=True)
    mAssert.Equals retVal, totalNumberOfModules, totalNumberOfModules & " should be exported as files do not yet exist."

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_filesNotYetCreatedTrue_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_filesNotYetCreatedTrue of Class Module vtkExportConfigurationTester"
    Resume Next
    
End Sub

Public Sub Test_vtkExportConfiguration_filesNotYetCreatedFalse()
    Dim retVal As Integer

    On Error GoTo Test_vtkExportConfiguration_filesNotYetCreatedFalse_Error

    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=False)
    mAssert.Equals retVal, totalNumberOfModules, totalNumberOfModules & " should be exported as files do not yet exist."

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_filesNotYetCreatedFalse_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_filesNotYetCreatedFalse of Class Module vtkExportConfigurationTester"
    Resume Next

End Sub

Public Sub Test_vtkExportConfiguration_filesNotYetCreatedFalseFilesAreCreated1()
    Dim retVal As Integer

    On Error GoTo Test_vtkExportConfiguration_filesNotYetCreatedFalseFilesAreCreated1_Error

    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=False)
    mAssert.Should fso.FileExists(vtkPathToTestFolder & "\" & testModule1FileName), testModule1FileName & " should have been created."

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_filesNotYetCreatedFalseFilesAreCreated1_Error:
    Debug.Print "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_filesNotYetCreatedFalseFilesAreCreated of Class Module vtkExportConfigurationTester"
    Resume Next

End Sub


Public Sub Test_vtkExportConfiguration_filesNotYetCreatedFalseFilesAreCreated2()
    Dim retVal As Integer

    On Error GoTo Test_vtkExportConfiguration_filesNotYetCreatedFalseFilesAreCreated2_Error

    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=False)
    mAssert.Should fso.FileExists(vtkPathToTestFolder & "\" & testModule2FileName), testModule2FileName & " should have been created."

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_filesNotYetCreatedFalseFilesAreCreated2_Error:
    Debug.Print "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_filesNotYetCreatedFalseFilesAreCreated of Class Module vtkExportConfigurationTester"
    Resume Next

End Sub

Public Sub Test_vtkExportConfiguration_filesJustExportedTrue()
    Dim retVal As Integer
    
    On Error GoTo Test_vtkExportConfiguration_filesYetCreatedTrue_Error

    vtkExportOneModule wbForTest.VBProject, testModule1Name, vtkPathToTestFolder & "\" & testModule1FileName
    vtkExportOneModule wbForTest.VBProject, testModule2Name, vtkPathToTestFolder & "\" & testModule2FileName

    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=True)
    mAssert.Equals retVal, 0, "No module should be exported as files yet exist and onlyModified = True"

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_filesYetCreatedTrue_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_filesYetCreatedTrue of Class Module vtkExportConfigurationTester"
    Resume Next
End Sub

Public Sub Test_vtkExportConfiguration_filesJustExportedFalse()
    Dim retVal As Integer
    
    On Error GoTo Test_vtkExportConfiguration_filesJustExportedFalse_Error
    
    vtkExportOneModule wbForTest.VBProject, testModule1Name, vtkPathToTestFolder & "\" & testModule1FileName
    vtkExportOneModule wbForTest.VBProject, testModule2Name, vtkPathToTestFolder & "\" & testModule2FileName

    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=False)
    mAssert.Equals retVal, totalNumberOfModules, totalNumberOfModules & " modules should be exported as onlyModified = False."

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_filesJustExportedFalse_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_filesJustExportedFalse of Class Module vtkExportConfigurationTester"
    Resume Next
End Sub

Public Sub Test_vtkExportConfiguration_oneFileModifiedTrue()
    Dim retVal As Integer
    
    On Error GoTo Test_vtkExportConfiguration_oneFileModifiedTrue_Error

    vtkExportOneModule wbForTest.VBProject, testModule1Name, vtkPathToTestFolder & "\" & testModule1FileName
    vtkExportOneModule wbForTest.VBProject, testModule2Name, vtkPathToTestFolder & "\" & testModule2FileName

    insertDummyProcedureInCodeModule wbForTest.VBProject.VBComponents(testModule1Name).CodeModule, dummyInt:=1
    
    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=True)
    mAssert.Equals retVal, 1, "1 module should be exported as 1 module has been modified and onlyModified = True"

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_oneFileModifiedTrue_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_oneFileModifiedTrue of Class Module vtkExportConfigurationTester"
    Resume Next
    
End Sub


Public Sub Test_vtkExportConfiguration_oneFileModifiedFalse()
    Dim retVal As Integer
    
    On Error GoTo Test_vtkExportConfiguration_oneFileModifiedFalse_Error

    vtkExportOneModule wbForTest.VBProject, testModule1Name, vtkPathToTestFolder & "\" & testModule1FileName
    vtkExportOneModule wbForTest.VBProject, testModule2Name, vtkPathToTestFolder & "\" & testModule2FileName

    insertDummyProcedureInCodeModule wbForTest.VBProject.VBComponents(testModule1Name).CodeModule, dummyInt:=1
    
    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=False)
    mAssert.Equals retVal, totalNumberOfModules, totalNumberOfModules & " modules should be exported as onlyModified = False"

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_oneFileModifiedFalse_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_oneFileModifiedFalse of Class Module vtkExportConfigurationTester"
    Resume Next
    
End Sub


Public Sub Test_vtkExportConfiguration_twoFilesModifiedTrue()
    Dim retVal As Integer
    
    On Error GoTo Test_vtkExportConfiguration_twoFilesModifiedTrue_Error

    vtkExportOneModule wbForTest.VBProject, testModule1Name, vtkPathToTestFolder & "\" & testModule1FileName
    vtkExportOneModule wbForTest.VBProject, testModule2Name, vtkPathToTestFolder & "\" & testModule2FileName

    insertDummyProcedureInCodeModule wbForTest.VBProject.VBComponents(testModule1Name).CodeModule, dummyInt:=1
    insertDummyProcedureInCodeModule wbForTest.VBProject.VBComponents(testModule2Name).CodeModule, dummyInt:=2
    
    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=True)
    mAssert.Equals retVal, 2, "2 modules should be exported as 2 modules have been modified and onlyModified = True"

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_twoFilesModifiedTrue_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_twoFilesModifiedTrue of Class Module vtkExportConfigurationTester"
    Resume Next

End Sub

Public Sub Test_vtkExportConfiguration_twoFilesModifiedFalse()
    Dim retVal As Integer
    
    On Error GoTo Test_vtkExportConfiguration_twoFilesModifiedFalse_Error

    vtkExportOneModule wbForTest.VBProject, testModule1Name, vtkPathToTestFolder & "\" & testModule1FileName
    vtkExportOneModule wbForTest.VBProject, testModule2Name, vtkPathToTestFolder & "\" & testModule2FileName

    insertDummyProcedureInCodeModule wbForTest.VBProject.VBComponents(testModule1Name).CodeModule, dummyInt:=1
    insertDummyProcedureInCodeModule wbForTest.VBProject.VBComponents(testModule1Name).CodeModule, dummyInt:=2
    
    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=False)
    mAssert.Equals retVal, totalNumberOfModules, totalNumberOfModules & " modules should be exported as onlyModified = False"

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_twoFilesModifiedFalse_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_twoFilesModifiedFalse of Class Module vtkExportConfigurationTester"
    Resume Next

End Sub

Public Sub Test_vtkExportConfiguration_importedModuleMustBeExported()
    Dim retVal As Integer

    On Error GoTo Test_vtkExportConfiguration_importedModuleMustBeExported_Error

    ' We export the modules already present so that they don't get exported again
    vtkExportOneModule wbForTest.VBProject, testModule1Name, vtkPathToTestFolder & "\" & testModule1FileName
    vtkExportOneModule wbForTest.VBProject, testModule1Name, vtkPathToTestFolder & "\" & testModule1FileName

    ' We import the module
    vtkImportOneModule wbForTest.VBProject, importedModuleName, vtkPathToTestFolder & "\" & importedModuleFileName
    
    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=True)
    
    mAssert.Equals retVal, 1, "1 module should be exported as the others already exist and 1 module has been imported."

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_importedModuleMustBeExported_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_importedFileMustBeExported of Class Module vtkExportConfigurationTester"
    Resume Next
End Sub


Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_filesNotYetCreatedTrue"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_filesNotYetCreatedFalse"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_filesNotYetCreatedFalseFilesAreCreated1"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_filesNotYetCreatedFalseFilesAreCreated2"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_filesJustExportedTrue"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_filesJustExportedFalse"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_oneFileModifiedTrue"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_oneFileModifiedFalse"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_twoFilesModifiedTrue"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_twoFilesModifiedFalse"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_importedModuleMustBeExported"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkExportConfiguration_filesNotYetCreatedTrue": Test_vtkExportConfiguration_filesNotYetCreatedTrue
        Case "Test_vtkExportConfiguration_filesNotYetCreatedFalse": Test_vtkExportConfiguration_filesNotYetCreatedFalse
        Case "Test_vtkExportConfiguration_filesNotYetCreatedFalseFilesAreCreated1": Test_vtkExportConfiguration_filesNotYetCreatedFalseFilesAreCreated1
        Case "Test_vtkExportConfiguration_filesNotYetCreatedFalseFilesAreCreated2": Test_vtkExportConfiguration_filesNotYetCreatedFalseFilesAreCreated2
        Case "Test_vtkExportConfiguration_filesJustExportedTrue": Test_vtkExportConfiguration_filesJustExportedTrue
        Case "Test_vtkExportConfiguration_filesJustExportedFalse": Test_vtkExportConfiguration_filesJustExportedFalse
        Case "Test_vtkExportConfiguration_oneFileModifiedTrue": Test_vtkExportConfiguration_oneFileModifiedTrue
        Case "Test_vtkExportConfiguration_oneFileModifiedFalse": Test_vtkExportConfiguration_oneFileModifiedFalse
        Case "Test_vtkExportConfiguration_twoFilesModifiedTrue": Test_vtkExportConfiguration_twoFilesModifiedTrue
        Case "Test_vtkExportConfiguration_twoFilesModifiedFalse": Test_vtkExportConfiguration_twoFilesModifiedFalse
        Case "Test_vtkExportConfiguration_importedModuleMustBeExported": Test_vtkExportConfiguration_importedModuleMustBeExported
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

